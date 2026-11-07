import os
import uuid
import json
import asyncio
import urllib.parse
from typing import List, Literal
from fastapi import FastAPI, HTTPException, Security, Depends
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials
from pydantic import BaseModel
import httpx
import aio_pika

app = FastAPI(title="Tender AI API")
security = HTTPBearer()

SUPABASE_URL = os.getenv("SUPABASE_URL", "")
SUPABASE_KEY = os.getenv("SUPABASE_KEY", "")
CLOUDAMQP_URL = os.getenv("CLOUDAMQP_URL", "")
API_TOKEN = os.getenv("API_TOKEN", "my_secret_token_for_gpt")


class DocumentInput(BaseModel):
    url: str
    name: str


class AnalyzeRequest(BaseModel):
    client_name: str = "Неизвестный клиент"
    mode: Literal["analytics", "estimate"] = "analytics"
    documents: List[DocumentInput]


class SupabaseAsync:
    def __init__(self, url: str, key: str):
        self.base_url = url.rstrip("/")
        self.headers = {
            "apikey": key,
            "Authorization": f"Bearer {key}",
            "Content-Type": "application/json",
            "Prefer": "return=representation",
        }

    async def _request(self, method: str, endpoint: str, **kwargs):
        async with httpx.AsyncClient(http2=False, timeout=30.0) as client:
            url = f"{self.base_url}/rest/v1/{endpoint}"
            last_err = None
            for _ in range(3):
                try:
                    resp = await client.request(
                        method, url, headers=self.headers, **kwargs
                    )
                    resp.raise_for_status()
                    if resp.content:
                        return resp.json()
                    return {}
                except httpx.HTTPStatusError as e:
                    raise Exception(
                        f"БД вернула ошибку {e.response.status_code}: {e.response.text}"
                    )
                except Exception as e:
                    last_err = e
                    await asyncio.sleep(1)
            raise last_err

    async def get_client_by_name(self, name: str):
        encoded = urllib.parse.quote(name)
        return await self._request("GET", f"clients?name=eq.{encoded}&select=id")

    async def create_client(self, name: str):
        return await self._request(
            "POST",
            "clients",
            json={"name": name},
        )

    async def create_calculation(self, data: dict):
        return await self._request(
            "POST",
            "calculations",
            json=data,
        )

    async def create_documents(self, data: list):
        return await self._request(
            "POST",
            "documents",
            json=data,
        )

    async def get_documents_by_calc(self, calc_id: str):
        return await self._request(
            "GET", f"documents?calculation_id=eq.{calc_id}&select=*"
        )

    async def get_documents_by_task(self, task_id: str):
        return await self._request("GET", f"documents?task_id=eq.{task_id}&select=*")

    async def update_calculation(self, calc_id: str, data: dict):
        return await self._request("PATCH", f"calculations?id=eq.{calc_id}", json=data)


def verify_token(credentials: HTTPAuthorizationCredentials = Security(security)):
    if credentials.credentials != API_TOKEN:
        raise HTTPException(status_code=401, detail="Invalid token")
    return credentials.credentials


# ИЗМЕНЕНИЕ: Отправляем в очередь task_id, а воркер сам читает документы из БД
async def send_to_queue(task_id: str, mode: str):
    connection = await aio_pika.connect_robust(CLOUDAMQP_URL)
    async with connection:
        channel = await connection.channel()
        queue = await channel.declare_queue("tender_tasks", durable=True)
        message_body = json.dumps(
            {
                "task_id": task_id,
                "mode": mode,
            }
        ).encode("utf-8")
        await channel.default_exchange.publish(
            aio_pika.Message(
                body=message_body, delivery_mode=aio_pika.DeliveryMode.PERSISTENT
            ),
            routing_key=queue.name,
        )


@app.post("/api/v1/tenders/analyze", dependencies=[Depends(verify_token)])
async def analyze_tenders(request: AnalyzeRequest):
    try:
        db = SupabaseAsync(SUPABASE_URL, SUPABASE_KEY)

        clients = await db.get_client_by_name(request.client_name)
        if clients:
            db_client_id = clients[0]["id"]
        else:
            new_client = await db.create_client(request.client_name)
            db_client_id = new_client[0]["id"]

        calculation_id = str(uuid.uuid4())
        await db.create_calculation(
            {"id": calculation_id, "client_id": db_client_id, "status": "pending"}
        )

        # Создаем ЕДИНУЮ задачу и отдельную запись в documents на каждый файл.
        # Важно: в одном batch не должно быть повторяющихся ссылок на один и тот же документ.
        task_id = str(uuid.uuid4())
        seen_urls = set()
        docs_payload = []

        for doc in request.documents:
            doc_url = (doc.url or "").strip()
            if not doc_url or doc_url in seen_urls:
                continue
            seen_urls.add(doc_url)
            docs_payload.append(
                {
                    "task_id": task_id,
                    "calculation_id": calculation_id,
                    "document_url": doc_url,
                    "document_name": doc.name,
                    "status": "pending",
                    "extracted_data": {"stage": "queued"},
                }
            )

        if not docs_payload:
            raise HTTPException(
                status_code=400,
                detail="В запросе нет валидных документов для обработки.",
            )

        await db.create_documents(docs_payload)

        # Отправляем задачу на пакетную обработку.
        await send_to_queue(task_id, request.mode)

        return {
            "batch_id": calculation_id,
            "task_id": task_id,
            "documents_total": len(docs_payload),
            "message": f"Принято {len(docs_payload)} документов в пакет. Режим: {request.mode}",
        }

    except Exception as e:
        print(f"[API ERROR] {str(e)}")
        raise HTTPException(status_code=500, detail=f"Ошибка БД или Очереди: {str(e)}")


@app.get("/api/v1/tenders/status/{batch_id}", dependencies=[Depends(verify_token)])
async def check_status(batch_id: str):
    try:
        db = SupabaseAsync(SUPABASE_URL, SUPABASE_KEY)
        documents = await db.get_documents_by_calc(batch_id)

        if not documents:
            raise HTTPException(status_code=404, detail="Расчет не найден")

        statuses = [doc.get("status", "pending") for doc in documents]
        total_docs = len(documents)
        processing_statuses = {"pending", "processing"}
        done_docs = sum(1 for s in statuses if s not in processing_statuses)
        failed_docs = sum(1 for s in statuses if s == "error")
        parsed_docs = sum(1 for s in statuses if s == "completed")

        docs_state = []
        final_result = None
        task_id = documents[0].get("task_id") if documents else None

        for doc in documents:
            extracted = doc.get("extracted_data")
            if isinstance(extracted, str):
                try:
                    extracted = json.loads(extracted)
                except json.JSONDecodeError:
                    pass

            if isinstance(extracted, dict) and "final_result" in extracted:
                final_result = extracted.get("final_result")

            docs_state.append(
                {
                    "document_name": doc.get("document_name", "Unknown_Doc"),
                    "status": doc.get("status", "pending"),
                    "error": (
                        extracted.get("error") if isinstance(extracted, dict) else None
                    ),
                }
            )

        if any(s in processing_statuses for s in statuses):
            return {
                "status": "processing",
                "task_id": task_id,
                "progress": {
                    "total_docs": total_docs,
                    "done_docs": done_docs,
                    "parsed_docs": parsed_docs,
                    "failed_docs": failed_docs,
                },
                "documents": docs_state,
                "message": "Документы еще обрабатываются, ждем завершения...",
            }

        if failed_docs > 0:
            return {
                "status": "error",
                "task_id": task_id,
                "progress": {
                    "total_docs": total_docs,
                    "done_docs": done_docs,
                    "parsed_docs": parsed_docs,
                    "failed_docs": failed_docs,
                },
                "documents": docs_state,
                "message": "Часть документов не обработана. Итоговый расчет не построен.",
            }

        if final_result is None:
            return {
                "status": "processing",
                "task_id": task_id,
                "progress": {
                    "total_docs": total_docs,
                    "done_docs": done_docs,
                    "parsed_docs": parsed_docs,
                    "failed_docs": failed_docs,
                },
                "documents": docs_state,
                "message": "Документы обработаны, ожидается финальный расчет.",
            }

        try:
            await db.update_calculation(batch_id, {"status": "completed"})
        except Exception:
            pass

        return {
            "status": "completed",
            "task_id": task_id,
            "progress": {
                "total_docs": total_docs,
                "done_docs": done_docs,
                "parsed_docs": parsed_docs,
                "failed_docs": failed_docs,
            },
            "documents": docs_state,
            "data": final_result,
        }

    except HTTPException:
        raise
    except Exception as e:
        print(f"[API ERROR] 500: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))
