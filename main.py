import json
import logging
import os
import uuid
from contextlib import asynccontextmanager
from typing import Literal

import aio_pika
import asyncpg
from fastapi import Depends, FastAPI, HTTPException, Security
from fastapi.security import HTTPAuthorizationCredentials, HTTPBearer
from pydantic import BaseModel
from dotenv import load_dotenv

from database import create_database_pool

load_dotenv()

logger = logging.getLogger(__name__)

CLOUDAMQP_URL = os.getenv("CLOUDAMQP_URL", "")
API_TOKEN = os.getenv("API_TOKEN", "")


@asynccontextmanager
async def lifespan(app: FastAPI):
    if not API_TOKEN:
        raise RuntimeError("API_TOKEN is not set")
    app.state.pool = await create_database_pool()
    try:
        yield
    finally:
        await app.state.pool.close()


app = FastAPI(title="Tender AI API", lifespan=lifespan)
security = HTTPBearer()


class DocumentInput(BaseModel):
    url: str
    name: str


class AnalyzeRequest(BaseModel):
    client_name: str = "Неизвестный клиент"
    mode: Literal["analytics", "estimate"] = "analytics"
    documents: list[DocumentInput]


def verify_token(
    credentials: HTTPAuthorizationCredentials = Security(security),
) -> str:
    if not API_TOKEN or credentials.credentials != API_TOKEN:
        raise HTTPException(status_code=401, detail="Invalid token")
    return credentials.credentials


async def send_to_queue(task_id: uuid.UUID, mode: str) -> None:
    if not CLOUDAMQP_URL:
        raise RuntimeError("CLOUDAMQP_URL is not set")
    connection = await aio_pika.connect_robust(CLOUDAMQP_URL)
    async with connection:
        channel = await connection.channel()
        queue = await channel.declare_queue("tender_tasks", durable=True)
        message_body = json.dumps(
            {"task_id": str(task_id), "mode": mode}
        ).encode("utf-8")
        await channel.default_exchange.publish(
            aio_pika.Message(
                body=message_body,
                delivery_mode=aio_pika.DeliveryMode.PERSISTENT,
            ),
            routing_key=queue.name,
        )


@app.post("/api/v1/tenders/analyze", dependencies=[Depends(verify_token)])
async def analyze_tenders(request: AnalyzeRequest):
    seen_urls: set[str] = set()
    document_inputs = []
    for document in request.documents:
        document_url = document.url.strip()
        if document_url and document_url not in seen_urls:
            seen_urls.add(document_url)
            document_inputs.append((document_url, document.name))

    if not document_inputs:
        raise HTTPException(
            status_code=400,
            detail="В запросе нет валидных документов для обработки.",
        )

    pool: asyncpg.Pool = app.state.pool
    calculation_id = uuid.uuid4()
    task_id = uuid.uuid4()

    try:
        async with pool.acquire() as connection:
            async with connection.transaction():
                client_id = await connection.fetchval(
                    """
                    INSERT INTO clients (name)
                    VALUES ($1)
                    ON CONFLICT (name) DO UPDATE SET name = EXCLUDED.name
                    RETURNING id
                    """,
                    request.client_name,
                )
                await connection.execute(
                    """
                    INSERT INTO calculations (id, client_id, status)
                    VALUES ($1, $2, 'pending')
                    """,
                    calculation_id,
                    client_id,
                )
                await connection.executemany(
                    """
                    INSERT INTO documents (
                        task_id, calculation_id, document_url, document_name,
                        status, extracted_data
                    )
                    VALUES ($1, $2, $3, $4, 'pending', '{"stage":"queued"}'::jsonb)
                    """,
                    [
                        (task_id, calculation_id, document_url, document_name)
                        for document_url, document_name in document_inputs
                    ],
                )
    except Exception as error:
        logger.exception("Could not create analysis records")
        raise HTTPException(
            status_code=500, detail="Не удалось создать задачу анализа."
        ) from error

    try:
        await send_to_queue(task_id, request.mode)
    except Exception as error:
        logger.exception("Could not enqueue analysis task")
        async with pool.acquire() as connection:
            async with connection.transaction():
                await connection.execute(
                    """
                    UPDATE documents
                    SET status = 'error',
                        extracted_data = '{"error":"Не удалось поставить задачу в очередь"}'::jsonb
                    WHERE task_id = $1
                    """,
                    task_id,
                )
                await connection.execute(
                    "UPDATE calculations SET status = 'error' WHERE id = $1",
                    calculation_id,
                )
        raise HTTPException(
            status_code=503, detail="Не удалось поставить задачу в очередь."
        ) from error

    return {
        "batch_id": str(calculation_id),
        "task_id": str(task_id),
        "documents_total": len(document_inputs),
        "message": (
            f"Принято {len(document_inputs)} документов в пакет. "
            f"Режим: {request.mode}"
        ),
    }


@app.get(
    "/api/v1/tenders/status/{batch_id}",
    dependencies=[Depends(verify_token)],
)
async def check_status(batch_id: uuid.UUID):
    pool: asyncpg.Pool = app.state.pool
    try:
        async with pool.acquire() as connection:
            documents = await connection.fetch(
                """
                SELECT task_id, document_name, status, extracted_data
                FROM documents
                WHERE calculation_id = $1
                ORDER BY created_at, id
                """,
                batch_id,
            )
    except Exception as error:
        logger.exception("Could not read calculation status")
        raise HTTPException(
            status_code=500, detail="Не удалось получить статус расчета."
        ) from error

    if not documents:
        raise HTTPException(status_code=404, detail="Расчет не найден")

    statuses = [document["status"] for document in documents]
    total_docs = len(documents)
    processing_statuses = {"pending", "processing"}
    done_docs = sum(status not in processing_statuses for status in statuses)
    failed_docs = statuses.count("error")
    parsed_docs = statuses.count("completed")

    docs_state = []
    final_result = None
    task_id = documents[0]["task_id"]
    for document in documents:
        extracted = document["extracted_data"] or {}
        if isinstance(extracted, str):
            try:
                extracted = json.loads(extracted)
            except json.JSONDecodeError:
                extracted = {}
        if not isinstance(extracted, dict):
            extracted = {}
        if "final_result" in extracted:
            final_result = extracted["final_result"]
        docs_state.append(
            {
                "document_name": document["document_name"],
                "status": document["status"],
                "error": extracted.get("error"),
            }
        )

    progress = {
        "total_docs": total_docs,
        "done_docs": done_docs,
        "parsed_docs": parsed_docs,
        "failed_docs": failed_docs,
    }
    if any(status in processing_statuses for status in statuses):
        return {
            "status": "processing",
            "task_id": str(task_id),
            "progress": progress,
            "documents": docs_state,
            "message": "Документы еще обрабатываются...",
        }
    if failed_docs:
        return {
            "status": "error",
            "task_id": str(task_id),
            "progress": progress,
            "documents": docs_state,
            "message": "Часть документов не обработана.",
        }
    if final_result is None:
        return {
            "status": "processing",
            "task_id": str(task_id),
            "progress": progress,
            "documents": docs_state,
            "message": "Документы обработаны, ожидается финальный расчет.",
        }

    return {
        "status": "completed",
        "task_id": str(task_id),
        "progress": progress,
        "documents": docs_state,
        "data": final_result,
    }
