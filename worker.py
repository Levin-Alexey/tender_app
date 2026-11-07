import os
import json
import asyncio
import tempfile
from pathlib import Path
import mimetypes
import urllib.parse
import re

import httpx
import aio_pika
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from supabase import create_client, Client

# Настройки
SUPABASE_URL = os.getenv("SUPABASE_URL", "")
SUPABASE_KEY = os.getenv("SUPABASE_KEY", "")
CLOUDAMQP_URL = os.getenv("CLOUDAMQP_URL", "")
OPENROUTER_API_KEY = os.getenv("OPENROUTER_API_KEY", "")
UNSTRUCTURED_API_KEY = os.getenv("UNSTRUCTURED_API_KEY", "")

UNSTRUCTURED_API_URL = os.getenv(
    "UNSTRUCTURED_API_URL", "https://api.unstructuredapp.io/general/v0/general"
).strip()

if UNSTRUCTURED_API_URL.startswith("https://api.unstructured.io"):
    UNSTRUCTURED_API_URL = UNSTRUCTURED_API_URL.replace(
        "https://api.unstructured.io", "https://api.unstructuredapp.io", 1
    )

supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)


def _parse_llm_json(raw_text: str) -> dict:
    try:
        data = json.loads(raw_text)
        if isinstance(data, dict):
            return data
    except json.JSONDecodeError:
        pass
    match = re.search(r"\{[\s\S]*\}", raw_text)
    if match:
        try:
            data = json.loads(match.group(0))
            if isinstance(data, dict):
                return data
        except json.JSONDecodeError:
            pass
    raise RuntimeError("LLM вернул невалидный JSON")


def _has_extracted_tender_facts(raw_data: dict) -> bool:
    sections = raw_data.get("sections")
    if not isinstance(sections, dict):
        return False

    for section_key, facts in sections.items():
        allowed_parameters = TENDER_DATA_PARAMETERS.get(section_key)
        if not allowed_parameters:
            continue
        if not isinstance(facts, list):
            continue
        for fact in facts:
            if (
                not isinstance(fact, dict)
                or fact.get("parameter") not in allowed_parameters
                or fact.get("status") == "missing"
            ):
                continue
            if any(
                str(fact.get(field) or "").strip()
                for field in ("value", "source_document", "raw_context")
            ):
                return True
    return False

# ИЗМЕНЕНИЕ 1: Добавили слово "ИНВЕНТАРЬ" в название раздела
TENDER_DATA_SECTIONS = (
    ("НМЦК И ДОГОВОР", "nmck_contract"),
    ("ОБЪЕКТЫ И ПЛОЩАДИ", "sites_areas"),
    ("ПЕРСОНАЛ", "personnel"),
    ("ОБОРУДОВАНИЕ И ТЕХНИКА", "equipment"),
    ("РАСХОДНИКИ, ИНВЕНТАРЬ И МАТЕРИАЛЫ", "consumables"), 
    ("ДОПОЛНИТЕЛЬНЫЕ РАБОТЫ", "additional_services"),
    ("РИСКИ И ТРЕБОВАНИЯ", "risks_requirements"),
    ("ПРОЧЕЕ", "other"),
)

FACT_STATUSES = {
    "found": "Найдено",
    "missing": "Не найдено",
    "insufficient": "Недостаточно данных",
    "conflict": "Конфликт документов",
    "manual_review": "Требуется ручная проверка",
}

# ИЗМЕНЕНИЕ 2: Разбили химию и инвентарь на отдельные конкретные параметры
TENDER_DATA_PARAMETERS = {
    "nmck_contract": (
        "Наличие НМЦК",
        "Разбивка начальной цены по видам работ (или общая сумма)",
        "Срок оказания услуг (период, в течение скольких месяцев/лет)",
        "Требуются ли лицензии на дезинфекцию, дезинсекцию, дератизацию",
        "Наличие у Исполнителя договоров с полигонами (вывоз мусора, ТБО, снега, снегоплавильная станция)",
        "Наличие допусков, разрешений, лицензий, сертификатов ФСБ",
        "Страхование (гражданская ответственность, жизнь и здоровье)",
        "Предоставляет ли Заказчик помещения/технику в аренду (взимание платы)",
    ),
    "sites_areas": (
        "Вид деятельности и назначение объекта (пищевое, склады, офис, медицина и т.д.)",
        "Расстояние между адресами объектов",
        "Оказываются ли услуги на территории ЗАТО",
        "Площади помещений для уборки",
        "Площади уборки внешней территории",
        "Удаленность объекта от остановок общественного транспорта",
    ),
    "personnel": (
        "Гражданство персонала (иностранцы или только РФ)",
        "Перечень должностей (уборщицы, дворники, слесари, электрики, менеджеры и т.д.)",
        "Количество персонала (штат, в смену, для видов уборки)",
        "Норма площади уборки на 1 человека",
        "Часы работы, график, смены, выходные, праздники, круглосуточная уборка",
        "Требования о постоянном присутствии менеджера на объекте",
        "Требования по охране труда и технике безопасности",
        "Санитарные и медицинские книжки, допуски для работы на высоте",
        "Доставка персонала (силами Исполнителя или Заказчика)",
    ),
    "equipment": (
        "Уборка с помощью СВЕП-систем, флаундеров (Vileda)",
        "Оборудование для уборки помещений (ПММ, роторы, пылесосы, водососы и т.д.) и его количество",
        "Оборудование для уборки территории (снегоуборщики, газонокосилки, АВД и т.д.)",
        "Механизированная уборка спецтехникой (тракторы, погрузчики, самосвалы)",
        "Оборудование и инвентарь для слесарных, плотницких, ремонтных работ",
        "Скреперы, снегорезы, скребки для уборки крыши",
    ),
    "consumables": (
        "Что из химии, инвентаря и оборудования предоставляет Заказчик",
        "Что из химии, инвентаря и оборудования предоставляет Исполнитель",
        "Перечень профессиональной химии и моющих средств",
        "Перечень уборочного инвентаря (швабры, МОПы, ведра, тележки, сгоны, щетки, метлы и т.д.)",
        "Перечень мелких расходных материалов для уборки (пакеты для мусора, салфетки, губки, тряпки)",
        "Противогололедные реагенты, техническая соль, песок, гранитная крошка",
        "Перечень санитарно-гигиенических материалов (бумага, мыло, освежители, средства для ПММ)",
        "Количество сотрудников Заказчика и посетителей (для расчета расходников)",
        "Спецодежда (огнеупорные костюмы, каски, жилеты, противогазы, обувь)",
        "Бахилы, одноразовые халаты, шапочки, перчатки",
        "Материалы для ремонта, отделки, разметки, побелки и сторона их предоставления",
        "Предоставление Исполнителем диспенсеров, дозаторов, электросушилок, корзин для мусора",
    ),
    "additional_services": (
        "Химчистка мягкой мебели, штор, тюля, ковров (площади, количество, периодичность)",
        "Мытье окон, витражей внутри и снаружи (площади, количество, периодичность)",
        "Мытье окон альпинистами, с автовышки, гидроподъемника",
        "Телескопическое оборудование для мытья окон",
        "Вывоз мусора, ТБО, строительного мусора (объем, периодичность, предоставление контейнеров)",
        "Вывоз и утилизация снега (объем, периодичность)",
        "Замена грязезащитных ковров (ковросервис, размеры, периодичность)",
        "Чистка кровли, сбивание сосулек, наледи (площади, оборудование)",
        "Благоустройство, озеленение (полив, покос, посадка, спил деревьев, отлов животных)",
        "Дезинсекция и дератизация (площади, периодичность)",
        "Мытье бассейна, фонтана",
    ),
    "risks_requirements": (
        "Штрафные санкции и удержания",
        "Особые условия приемки услуг",
    ),
    "other": ("Прочие значимые данные",),
}


# --- БЛОК 1: Парсеры ---
def _safe_filename_from_url_or_response(url: str, response: httpx.Response) -> str:
    parsed = urllib.parse.urlparse(url)
    query = urllib.parse.parse_qs(parsed.query)
    filename_values = query.get("filename")
    if filename_values and filename_values[0].strip():
        return Path(filename_values[0].strip()).name
    content_disposition = response.headers.get("content-disposition", "")
    match = re.search(r'filename\*?=(?:UTF-8\'\')?"?([^";]+)"?', content_disposition)
    if match:
        return Path(urllib.parse.unquote(match.group(1))).name
    suffix = Path(parsed.path).suffix
    if suffix:
        return "document" + suffix
    content_type = response.headers.get("content-type", "")
    ext = mimetypes.guess_extension(content_type.split(";")[0].strip()) or ""
    return "document" + ext


def extract_text_from_excel(file_path: str) -> str:
    sheets = pd.read_excel(file_path, sheet_name=None, header=None, dtype=str)
    parts = []
    for sheet_name, df in sheets.items():
        df = df.fillna("")
        parts.append(f"\n\n# Лист: {sheet_name}")
        for row_num, row in enumerate(df.itertuples(index=False), start=1):
            values = [str(cell).strip() for cell in row if str(cell).strip()]
            if values:
                parts.append(f"Строка {row_num}: " + " | ".join(values))
    text = "\n".join(parts).strip()
    if not text:
        raise RuntimeError("Excel-файл пустой или не удалось извлечь текст")
    return text


def _elements_to_text(elements) -> str:
    parts = []
    for element in elements:
        if not isinstance(element, dict):
            continue
        category = element.get("type") or element.get("category") or "Element"
        text = (element.get("text") or "").strip()
        metadata = element.get("metadata") or {}
        table_html = (metadata.get("text_as_html") or "").strip()
        if table_html:
            parts.append(f"\n\n## {category}\n{table_html}")
        elif text:
            parts.append(f"\n\n## {category}\n{text}")
    result = "".join(parts).strip()
    if not result:
        raise RuntimeError("Unstructured вернул пустой текст")
    return result


async def call_smart_parser(file_path: str) -> str:
    if not UNSTRUCTURED_API_KEY:
        raise RuntimeError("UNSTRUCTURED_API_KEY is not set")
    filename = os.path.basename(file_path) or "document"
    mime_type = mimetypes.guess_type(filename)[0] or "application/octet-stream"
    headers = {
        "accept": "application/json",
        "unstructured-api-key": UNSTRUCTURED_API_KEY,
    }
    data = {
        "strategy": "hi_res",
        "languages": "rus",
        "skip_infer_table_types": "[]",
        "include_page_breaks": "true",
    }
    timeout_config = httpx.Timeout(900.0, read=900.0, connect=60.0)

    async with httpx.AsyncClient(
        timeout=timeout_config, follow_redirects=True
    ) as http_client:
        with open(file_path, "rb") as f:
            response = await http_client.post(
                UNSTRUCTURED_API_URL,
                headers=headers,
                data=data,
                files={"files": (filename, f, mime_type)},
            )
    try:
        response.raise_for_status()
    except httpx.HTTPStatusError as e:
        raise RuntimeError(
            f"Ошибка Unstructured API {response.status_code}: {response.text[:2000]}"
        ) from e
    elements = response.json()
    if not isinstance(elements, list):
        raise RuntimeError(f"Unstructured вернул неожиданный ответ: {elements}")
    return _elements_to_text(elements)


# --- БЛОК 2: LLM (DeepSeek) ---
async def analyze_text_with_deepseek_analytics(text: str) -> dict:
    if not OPENROUTER_API_KEY:
        raise RuntimeError("OPENROUTER_API_KEY is not set")
    api_url = "https://openrouter.ai/api/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {OPENROUTER_API_KEY}",
        "Content-Type": "application/json",
    }
    system_prompt = """Ты — Главный аналитик тендерной документации и старший сметчик клининговой компании.
Твоя задача — подготовить КРАТКИЙ УПРАВЛЕНЧЕСКИЙ АУДИТ тендера: понять, можно ли переходить к ручному расчету, какие условия влияют на стоимость и какие риски нужно снять до участия.

НЕ рассчитывай цены, ФОТ, прибыль, налоги и итоговую стоимость. Не придумывай данные.

Верни только строгий JSON по указанной схеме:
{
    "tender_card": {
        "nmck": {"value": "", "status": "found|missing|insufficient|conflict|manual_review", "source_document": "", "raw_context": ""},
        "service_term": {"value": "", "status": "", "source_document": "", "raw_context": ""},
        "object_type": {"value": "", "status": "", "source_document": "", "raw_context": ""},
        "addresses_and_areas": {"value": "", "status": "", "source_document": "", "raw_context": ""},
        "personnel_summary": {"value": "", "status": "", "source_document": "", "raw_context": ""},
        "work_schedule": {"value": "", "status": "", "source_document": "", "raw_context": ""}
    },
    "cost_drivers": [
        {"factor": "", "impact": "", "status": "", "source_document": "", "raw_context": ""}
    ],
    "risks": [
        {"risk": "", "severity": "high|medium|low", "status": "", "source_document": "", "raw_context": "", "recommendation": ""}
    ],
    "conflicts": [
        {"topic": "", "details": "", "source_document": "", "raw_context": ""}
    ],
    "critical_questions": [
        {"question": "", "reason": "", "severity": "high|medium|low"}
    ],
    "overall_assessment": {
        "readiness": "ready|needs_clarification|high_risk",
        "summary": ""
    }
}

ПРАВИЛА:
1. Статусы: found — данные достаточны; missing — данных нет; insufficient — данные есть, но для решения недостаточны; conflict — документы противоречат; manual_review — нужна ручная трактовка.
2. У каждого найденного факта указывай source_document и дословную raw_context. Для missing источник и цитата могут быть пустыми.
3. При конфликте НЕ выбирай версию самостоятельно. Зафиксируй оба значения и источники в details/raw_context, а также добавь запись в conflicts.
4. К критичным вопросам относись строго: включай только то, что блокирует расчет, меняет цену или создает договорной риск.
5. В cost_drivers отрази персонал и графики, площади, работу по выходным, расходники, оборудование, спецтехнику, дополнительные работы, удаленность, лицензии и другие условия, влияющие на стоимость.
6. В risks отрази лицензии, допуски, медкнижки, гражданство, страхование, аренду, сторону предоставления оборудования/материалов, неясные объемы и рискованные договорные требования.
7. Если по карточке тендера данных нет — ставь status="missing", value="". НИКОГДА НЕ ДОДУМЫВАЙ.
8. В overall_assessment.readiness: ready — критичных пробелов нет; needs_clarification — есть вопросы до расчета; high_risk — есть существенные риски или конфликты.
"""
    payload = {
        "model": "deepseek/deepseek-v4-pro",
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": f"Проанализируй текст и верни JSON:\n\n{text}"},
        ],
        "reasoning": {"enabled": True},
        "temperature": 0.1,
        "response_format": {"type": "json_object"},
    }
    async with httpx.AsyncClient(timeout=240.0) as client:
        response = await client.post(api_url, headers=headers, json=payload)
        response.raise_for_status()
        return _parse_llm_json(
            response.json()["choices"][0]["message"]
            .get("content", "")
            .replace("```json", "")
            .replace("```", "")
            .strip()
        )


async def extract_tender_data_for_estimator(text: str) -> dict:
    if not OPENROUTER_API_KEY:
        raise RuntimeError("OPENROUTER_API_KEY is not set")
    api_url = "https://openrouter.ai/api/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {OPENROUTER_API_KEY}",
        "Content-Type": "application/json",
    }
    
    sections = ", ".join(section_key for _, section_key in TENDER_DATA_SECTIONS)
    
    # Формируем список параметров для промпта динамически
    parameter_lists = "\n".join([f"- Раздел {section}: " + ", ".join(params) for section, params in TENDER_DATA_PARAMETERS.items()])

    # ИЗМЕНЕНИЕ 3: Добавили жесткое ПРАВИЛО №7 в промпт ИИ
    system_prompt = f"""Ты — старший сметчик клининговой компании. Прочитай ОБЪЕДИНЕННЫЙ текст документов тендера и подготовь КАРТУ ИСХОДНЫХ ДАННЫХ для человека-сметчика.

Твоя задача — только извлечь факты. НЕ рассчитывай цены, ФОТ, прибыль, налоги и не сопоставляй с прайс-листом.

РАЗДЕЛЫ JSON: {sections}.

ОБЯЗАТЕЛЬНО ПРОВЕРЬ КАЖДЫЙ ИЗ ЭТИХ ПАРАМЕТРОВ (сохраняй их точные названия):
{parameter_lists}

Верни JSON с объектом "sections". Внутри него обязательно должны быть все разделы: {sections}.
В каждом разделе верни отдельную запись для КАЖДОГО перечисленного параметра с точным названием.
Если в документах сведений по параметру нет, всё равно верни запись с status="missing",
пустым value и вопросом заказчику в comment. Не пропускай разделы или параметры.

Для каждого найденного факта верни объект:
{{
    "parameter": "точное название проверяемого параметра из списка выше",
    "status": "found | missing | insufficient | conflict | manual_review",
    "value": "значение только из документов",
    "source_document": "имя документа после заголовка ДОКУМЕНТ",
    "source_location": "страница, лист, строка или раздел, если есть",
    "raw_context": "точная цитата из документа",
    "comment": "что должен проверить сметчик"
}}

ПРАВИЛА:
1. Не придумывай данные. Если значение отсутствует — status="missing", value="", а в comment сформулируй вопрос заказчику.
2. Если данные есть, но не позволяют считать (например, расходники без количества) — status="insufficient" и укажи чего не хватает.
3. Если документы противоречат друг другу — status="conflict"; сохрани оба значения и оба источника в value/raw_context. Не выбирай одно значение сам.
4. Если формулировка двусмысленна — status="manual_review" и объясни причину.
5. Для каждого статуса кроме missing обязательно заполняй source_document и raw_context дословной цитатой.
6. Все количественные значения сохраняй вместе с единицами измерения, периодичностью, площадью, графиком или стороной предоставления, когда они указаны.
7. ВАЖНО: Если параметр запрашивает ПЕРЕЧЕНЬ (например, список инвентаря, химии или расходников), выписывай ВСЕ найденные позиции полностью, сохраняя их количество и характеристики. Не сокращай и не обобщай списки!
8. Не создавай параметров вне списка выше. Возвращай только строгий JSON.

Формат ответа:
{{
    "sections": {{
        "personnel": [{{"parameter": "Количество персонала (штат, в смену, для видов уборки)", "status": "found", "value": "8 уборщиков", "source_document": "ТЗ.pdf", "source_location": "стр. 12", "raw_context": "...", "comment": ""}}]
    }},
    "critical_questions": ["вопрос заказчику, без которого нельзя рассчитывать"],
    "documents_summary": ["краткое замечание по документу"]
}}"""
    payload = {
        "model": "deepseek/deepseek-v4-pro",
        "messages": [
            {"role": "system", "content": system_prompt},
            {
                "role": "user",
                "content": f"Вот объединенный текст всех документов тендера:\n\n{text}",
            },
        ],
        "reasoning": {"enabled": True},
        "temperature": 0.1,
        "response_format": {"type": "json_object"},
    }
    async with httpx.AsyncClient(timeout=240.0) as client:
        for attempt in range(2):
            response = await client.post(api_url, headers=headers, json=payload)
            response.raise_for_status()
            raw_data = _parse_llm_json(
                response.json()["choices"][0]["message"]
                .get("content", "")
                .replace("```json", "")
                .replace("```", "")
                .strip()
            )
            if _has_extracted_tender_facts(raw_data):
                return raw_data
            if attempt == 0:
                payload["messages"].append(
                    {
                        "role": "user",
                        "content": (
                            "Предыдущий ответ не содержал извлеченных фактов. "
                            "Повтори анализ исходного текста выше. Верни объект "
                            f'"sections" со всеми разделами ({sections}) и записями '
                            "по каждому параметру. Если факта нет, укажи status "
                            '"missing"; если факт есть — укажи value, '
                            "source_document и дословный raw_context. "
                            "Не возвращай пустые разделы."
                        ),
                    }
                )

    raise RuntimeError(
        "LLM дважды вернул отчет без извлеченных фактов; пустой отчет не сформирован"
    )


# --- БЛОК 3: Обработка очередей ---
DOWNLOAD_TIMEOUT = httpx.Timeout(300.0, connect=30.0)
DOWNLOAD_ATTEMPTS = 4


def _is_yandex_disk_url(url: str) -> bool:
    host = (urllib.parse.urlparse(url).hostname or "").lower()
    return (
        host == "yadi.sk"
        or host.startswith("disk.yandex.")
        or host.startswith("disk.360.yandex.")
    )


async def _get_with_retries(client: httpx.AsyncClient, url: str, **kwargs):
    last_error = None
    for attempt in range(1, DOWNLOAD_ATTEMPTS + 1):
        try:
            resp = await client.get(url, **kwargs)
            if resp.status_code in (429, 500, 502, 503, 504):
                last_error = RuntimeError(f"HTTP {resp.status_code}")
            else:
                resp.raise_for_status()
                return resp
        except httpx.HTTPStatusError as e:
            raise RuntimeError(
                f"Не удалось скачать файл: HTTP {e.response.status_code} {e.response.text[:300]}"
            ) from e
        except httpx.RequestError as e:
            last_error = RuntimeError(f"{type(e).__name__}: {e}")
        if attempt < DOWNLOAD_ATTEMPTS:
            await asyncio.sleep(2**attempt)
    raise RuntimeError(f"Не удалось скачать файл после {DOWNLOAD_ATTEMPTS} попыток: {last_error}")


async def download_file(url: str, temp_dir: str) -> str:
    url = url.strip()
    async with httpx.AsyncClient(
        follow_redirects=True, timeout=DOWNLOAD_TIMEOUT
    ) as client:
        if _is_yandex_disk_url(url):
            resp = await _get_with_retries(
                client,
                "https://cloud-api.yandex.net/v1/disk/public/resources/download",
                params={"public_key": url},
            )
            href = resp.json().get("href")
            if not href:
                raise RuntimeError(f"Яндекс.Диск не вернул ссылку на скачивание: {url}")
            url = href

        resp = await _get_with_retries(client, url)
        filename = _safe_filename_from_url_or_response(url, resp)
        path = os.path.join(temp_dir, filename)
        with open(path, "wb") as f:
            f.write(resp.content)
        return path


# --- БЛОК 4: ГЕНЕРАТОР КАРТЫ ДАННЫХ В EXCEL ---
def upload_excel_to_supabase(file_path: str, filename: str) -> str:
    with open(file_path, "rb") as f:
        supabase.storage.from_("tenders").upload(
            filename,
            f.read(),
            file_options={
                "content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            },
        )
    public_url = supabase.storage.from_("tenders").get_public_url(filename)
    return (
        public_url.get("publicURL") or public_url.get("public_url") or ""
        if isinstance(public_url, dict)
        else str(public_url)
    )


def _normalize_tender_fact(fact, parameter: str) -> dict:
    fact = fact if isinstance(fact, dict) else {}
    status = fact.get("status", "missing")
    if status not in FACT_STATUSES:
        status = "manual_review"
    return {
        "parameter": fact.get("parameter") or parameter,
        "status": status,
        "value": fact.get("value", ""),
        "source_document": fact.get("source_document", ""),
        "source_location": fact.get("source_location", ""),
        "raw_context": fact.get("raw_context", ""),
        "comment": fact.get("comment", ""),
    }


def normalize_tender_data(raw_data: dict) -> dict:
    raw_sections = raw_data.get("sections")
    if not _has_extracted_tender_facts(raw_data):
        raise RuntimeError(
            "LLM не вернул извлеченные факты; нормализация пустого отчета запрещена"
        )
    if not isinstance(raw_sections, dict):
        raise RuntimeError("LLM вернул ответ без объекта sections")
    normalized_sections = {}
    clarification_questions = list(raw_data.get("critical_questions", []))

    for _, section_key in TENDER_DATA_SECTIONS:
        facts_by_parameter = {}
        for fact in raw_sections.get(section_key, []):
            normalized = _normalize_tender_fact(fact, "Прочие значимые данные")
            facts_by_parameter.setdefault(normalized["parameter"], []).append(
                normalized
            )

        section_facts = []
        for parameter in TENDER_DATA_PARAMETERS[section_key]:
            matched_facts = facts_by_parameter.pop(parameter, [])
            if matched_facts:
                section_facts.extend(matched_facts)
            else:
                section_facts.append(
                    _normalize_tender_fact(
                        {
                            "status": "missing",
                            "comment": f"Уточнить у заказчика: {parameter}.",
                        },
                        parameter,
                    )
                )

        for extra_facts in facts_by_parameter.values():
            section_facts.extend(extra_facts)

        normalized_sections[section_key] = section_facts

    for facts in normalized_sections.values():
        for fact in facts:
            if fact["status"] in {"missing", "insufficient", "conflict"}:
                question = fact["comment"] or fact["parameter"]
                if question not in clarification_questions:
                    clarification_questions.append(question)

    return {
        "sections": normalized_sections,
        "critical_questions": clarification_questions,
        "documents_summary": raw_data.get("documents_summary", []),
    }


def create_tender_data_report(data: dict, temp_dir: str, task_id: str) -> str:
    wb = openpyxl.Workbook()
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="4F81BD")
    status_fills = {
        "found": PatternFill("solid", fgColor="C6EFCE"),
        "missing": PatternFill("solid", fgColor="FFC7CE"),
        "insufficient": PatternFill("solid", fgColor="FFEB9C"),
        "conflict": PatternFill("solid", fgColor="F4B183"),
        "manual_review": PatternFill("solid", fgColor="D9EAD3"),
    }
    columns = [
        "Параметр",
        "Статус",
        "Найденные данные",
        "Документ-источник",
        "Страница / лист / строка",
        "Цитата из документа",
        "Комментарий для сметчика",
        "Решение сметчика",
    ]

    ws_summary = wb.active
    ws_summary.title = "СВОДКА"
    ws_summary.append(["Карта исходных данных тендера", ""])
    ws_summary.cell(row=1, column=1).font = Font(bold=True, size=14)
    ws_summary.append(["Раздел", "Найдено", "Пробелы/уточнения", "Конфликты"])
    for column in range(1, 5):
        ws_summary.cell(row=2, column=column).font = header_font
        ws_summary.cell(row=2, column=column).fill = header_fill

    row = 3
    total_found = total_gaps = total_conflicts = 0
    for section_name, section_key in TENDER_DATA_SECTIONS:
        facts = data["sections"][section_key]
        found = sum(fact["status"] == "found" for fact in facts)
        gaps = sum(fact["status"] in {"missing", "insufficient"} for fact in facts)
        conflicts = sum(fact["status"] == "conflict" for fact in facts)
        ws_summary.append([section_name, found, gaps, conflicts])
        total_found += found
        total_gaps += gaps
        total_conflicts += conflicts
        row += 1

    ws_summary.append(["ИТОГО", total_found, total_gaps, total_conflicts])
    for column in range(1, 5):
        ws_summary.cell(row=row, column=column).font = Font(bold=True)
    ws_summary.append([])
    ws_summary.append(["Документы", ""])
    for document_summary in data.get("documents_summary", []):
        ws_summary.append([str(document_summary), ""])
    for column, width in enumerate([42, 15, 20, 15], start=1):
        ws_summary.column_dimensions[openpyxl.utils.get_column_letter(column)].width = (
            width
        )

    ws_questions = wb.create_sheet(title="КРИТИЧНЫЕ УТОЧНЕНИЯ")
    ws_questions.append(["Вопрос или риск", "Действие сметчика"])
    for column in range(1, 3):
        ws_questions.cell(row=1, column=column).font = header_font
        ws_questions.cell(row=1, column=column).fill = PatternFill(
            "solid", fgColor="C0504D"
        )
    questions = data.get("critical_questions", [])
    if questions:
        for question in questions:
            ws_questions.append([str(question), "Запросить уточнение у заказчика"])
    else:
        ws_questions.append(["Критичных уточнений не выявлено", ""])
    ws_questions.column_dimensions["A"].width = 90
    ws_questions.column_dimensions["B"].width = 38

    ws_request = wb.create_sheet(title="ЗАПРОС ЗАКАЗЧИКУ")
    ws_request.append(["№", "Вопрос заказчику", "Причина"])
    for column in range(1, 4):
        ws_request.cell(row=1, column=column).font = header_font
        ws_request.cell(row=1, column=column).fill = header_fill
    for index, question in enumerate(questions, start=1):
        ws_request.append([index, str(question), "Данные нужны для расчета сметы"])
    if not questions:
        ws_request.append([1, "Уточнения не требуются", ""])
    ws_request.column_dimensions["A"].width = 8
    ws_request.column_dimensions["B"].width = 90
    ws_request.column_dimensions["C"].width = 38

    ws_sources = wb.create_sheet(title="ИСТОЧНИКИ")
    ws_sources.append(["Раздел", *columns[:-1]])
    for column in range(1, 8):
        ws_sources.cell(row=1, column=column).font = header_font
        ws_sources.cell(row=1, column=column).fill = header_fill

    for section_name, section_key in TENDER_DATA_SECTIONS:
        ws_section = wb.create_sheet(title=section_name[:31])
        ws_section.append(columns)
        for column in range(1, len(columns) + 1):
            ws_section.cell(row=1, column=column).font = header_font
            ws_section.cell(row=1, column=column).fill = header_fill

        for fact in data["sections"][section_key]:
            values = [
                fact["parameter"],
                FACT_STATUSES[fact["status"]],
                fact["value"],
                fact["source_document"],
                fact["source_location"],
                fact["raw_context"],
                fact["comment"],
                "",
            ]
            ws_section.append(values)
            current_row = ws_section.max_row
            ws_section.cell(row=current_row, column=2).fill = status_fills[
                fact["status"]
            ]
            ws_sources.append([section_name, *values[:-1]])

        for column, width in enumerate([34, 24, 48, 28, 24, 65, 50, 30], start=1):
            ws_section.column_dimensions[
                openpyxl.utils.get_column_letter(column)
            ].width = width
        ws_section.freeze_panes = "A2"
        ws_section.auto_filter.ref = ws_section.dimensions
        for worksheet_row in ws_section.iter_rows(min_row=2):
            for cell in worksheet_row:
                cell.alignment = Alignment(wrap_text=True, vertical="top")

    for column, width in enumerate([28, 34, 24, 48, 28, 24, 65], start=1):
        ws_sources.column_dimensions[openpyxl.utils.get_column_letter(column)].width = (
            width
        )
    ws_sources.freeze_panes = "A2"
    ws_sources.auto_filter.ref = ws_sources.dimensions

    for worksheet in (ws_summary, ws_questions, ws_request, ws_sources):
        for worksheet_row in worksheet.iter_rows():
            for cell in worksheet_row:
                cell.alignment = Alignment(wrap_text=True, vertical="top")

    file_path = os.path.join(temp_dir, f"Karta_Dannyh_Tendera_{task_id[:8]}.xlsx")
    wb.save(file_path)
    return file_path


def _normalize_extracted_data(raw):
    if isinstance(raw, dict):
        return raw
    if isinstance(raw, str):
        try:
            parsed = json.loads(raw)
            if isinstance(parsed, dict):
                return parsed
        except json.JSONDecodeError:
            return {}
    return {}


async def finalize_task_if_ready(task_id: str, mode: str):
    docs_resp = await asyncio.to_thread(
        lambda: supabase.table("documents")
        .select("id, status, document_name, extracted_data")
        .eq("task_id", task_id)
        .execute()
    )
    docs = docs_resp.data or []
    if not docs:
        raise RuntimeError(f"По task_id={task_id} не найдено документов")

    statuses = [doc.get("status", "pending") for doc in docs]
    if any(status in ("pending", "processing") for status in statuses):
        return

    if any(status == "error" for status in statuses):
        return

    extracted_by_doc = []
    for doc in docs:
        doc_name = doc.get("document_name", "Unknown_Doc")
        parsed_data = _normalize_extracted_data(doc.get("extracted_data"))
        text = parsed_data.get("parsed_text", "")
        if not text:
            raise RuntimeError(f"У документа '{doc_name}' отсутствует parsed_text")
        extracted_by_doc.append((doc_name, text))

    combined_text = ""
    for doc_name, text in extracted_by_doc:
        combined_text += f"\n\n{'='*30}\nДОКУМЕНТ: {doc_name}\n{'='*30}\n\n{text}"

    if mode == "analytics":
        final_result = await analyze_text_with_deepseek_analytics(combined_text)
    elif mode == "estimate":
        raw_tender_data = await extract_tender_data_for_estimator(combined_text)
        final_result = normalize_tender_data(raw_tender_data)
        final_result["report_type"] = "tender_data_map"
        final_result["report_note"] = (
            "Карта исходных данных для ручного расчета сметчиком. "
            "Финансовый расчет не выполнялся."
        )

        with tempfile.TemporaryDirectory() as temp_dir:
            excel_path = await asyncio.to_thread(
                create_tender_data_report, final_result, temp_dir, task_id
            )
            final_result["excel_download_url"] = await asyncio.to_thread(
                upload_excel_to_supabase,
                excel_path,
                f"Karta_Dannyh_Tendera_{task_id[:8]}.xlsx",
            )
    else:
        raise RuntimeError(f"Неизвестный режим: {mode}")

    payload = {
        "final_result": final_result,
        "stage": "finalized",
    }
    await asyncio.to_thread(
        lambda: supabase.table("documents")
        .update({"status": "completed", "extracted_data": payload})
        .eq("task_id", task_id)
        .execute()
    )


async def process_message(message: aio_pika.IncomingMessage):
    async with message.process(requeue=True, reject_on_redelivered=True):
        body = json.loads(message.body.decode())
        task_id = body.get("task_id")
        mode = body.get("mode", "analytics")

        if not task_id:
            raise RuntimeError("Некорректное сообщение: нет task_id")

        try:
            docs_resp = await asyncio.to_thread(
                lambda: supabase.table("documents")
                .select("id, document_url, document_name, status")
                .eq("task_id", task_id)
                .execute()
            )
            documents = docs_resp.data or []
            if not documents:
                raise RuntimeError(f"По task_id={task_id} не найдено документов")

            failed_docs = 0

            async def set_doc(doc_id, status, data):
                await asyncio.to_thread(
                    lambda: supabase.table("documents")
                    .update({"status": status, "extracted_data": data})
                    .eq("id", doc_id)
                    .execute()
                )

            async def fail_doc(doc_id, doc_name, err):
                error_msg = f"{doc_name}: {type(err).__name__}: {err}"
                print(f"[WORKER ERROR] {error_msg}")
                await set_doc(doc_id, "error", {"error": error_msg})

            with tempfile.TemporaryDirectory() as temp_dir:
                downloaded = []
                for doc in documents:
                    doc_id = doc.get("id")
                    url = doc.get("document_url")
                    doc_name = doc.get("document_name") or "Unknown_Doc"
                    if not doc_id or not url:
                        continue
                    await set_doc(doc_id, "processing", {"stage": "downloading"})
                    try:
                        sub_dir = os.path.join(temp_dir, str(doc_id))
                        os.makedirs(sub_dir, exist_ok=True)
                        file_path = await download_file(url, sub_dir)
                        downloaded.append((doc_id, doc_name, file_path))
                    except Exception as doc_error:
                        failed_docs += 1
                        await fail_doc(doc_id, doc_name, doc_error)

                semaphore = asyncio.Semaphore(3)

                async def parse_doc(doc_id, doc_name, file_path):
                    async with semaphore:
                        try:
                            await set_doc(doc_id, "processing", {"stage": "parsing"})
                            ext = Path(file_path).suffix.lower()
                            if ext in [".xlsx", ".xls"]:
                                text = await asyncio.to_thread(
                                    extract_text_from_excel, file_path
                                )
                            else:
                                text = await call_smart_parser(file_path)
                            await set_doc(
                                doc_id,
                                "completed",
                                {
                                    "stage": "parsed",
                                    "parsed_text": text,
                                    "source_file": os.path.basename(file_path),
                                    "document_name": doc_name,
                                },
                            )
                            return True
                        except Exception as doc_error:
                            await fail_doc(doc_id, doc_name, doc_error)
                            return False

                results = await asyncio.gather(
                    *(parse_doc(*item) for item in downloaded)
                )
                failed_docs += results.count(False)

            if failed_docs:
                return

            await finalize_task_if_ready(task_id, mode)

        except Exception as e:
            error_msg = f"{type(e).__name__}: {e}"
            await asyncio.to_thread(
                lambda: supabase.table("documents")
                .update({"status": "error", "extracted_data": {"error": error_msg}})
                .eq("task_id", task_id)
                .execute()
            )


async def main():
    conn = await aio_pika.connect_robust(CLOUDAMQP_URL)
    async with conn:
        channel = await conn.channel()
        await channel.set_qos(prefetch_count=1)
        queue = await channel.declare_queue("tender_tasks", durable=True)
        await queue.consume(process_message)
        await asyncio.Future()


if __name__ == "__main__":
    asyncio.run(main())