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
import asyncpg
from contextlib import asynccontextmanager

# Настройки
DATABASE_URL = os.getenv("DATABASE_URL", "") # Сюда положишь ссылку на новый PostgreSQL
CLOUDAMQP_URL = os.getenv("CLOUDAMQP_URL", "")
API_TOKEN = os.getenv("API_TOKEN", "my_secret_token_for_gpt")

@asynccontextmanager
async def lifespan(app: FastAPI):
    # Подключаемся к базе при старте HF Space
    app.state.pool = await asyncpg.create_pool(DATABASE_URL, min_size=1, max_size=10)
    yield
    await app.state.pool.close()

app = FastAPI(title="Tender AI API", lifespan=lifespan)
security = HTTPBearer()

class DocumentInput(BaseModel):
    url: str
    name: str

class AnalyzeRequest(BaseModel):
    client_name: str = "Неизвестный клиент"
    mode: Literal["analytics", "estimate"] = "analytics"
    documents: List[DocumentInput]

def verify_token(credentials: HTTPAuthorizationCredentials = Security(security)):
    if credentials.credentials != API_TOKEN:
        raise HTTPException(status_code=401, detail="Invalid token")
    return credentials.credentials

async def send_to_queue(task_id: str, mode: str):
    connection = await aio_pika.connect_robust(CLOUDAMQP_URL)
    async with connection:
        channel = await connection.channel()
        queue = await channel.declare_queue("tender_tasks", durable=True)
        message_body = json.dumps({
            "task_id": task_id,
            "mode": mode,
        }).encode("utf-8")
        await channel.default_exchange.publish(
            aio_pika.Message(body=message_body, delivery_mode=aio_pika.DeliveryMode.PERSISTENT),
            routing_key=queue.name,
        )

@app.post("/api/v1/tenders/analyze", dependencies=[Depends(verify_token)])
async def analyze_tenders(request: AnalyzeRequest):
    try:
        pool = app.state.pool
        async with pool.acquire() as conn:
            # Ищем или создаем клиента
            client_record = await conn.fetchrow("SELECT id FROM clients WHERE name = $1", request.client_name)
            if client_record:
                db_client_id = client_record["id"]
            else:
                db_client_id = await conn.fetchval("INSERT INTO clients (name) VALUES ($1) RETURNING id", request.client_name)

            calculation_id = str(uuid.uuid4())
            await conn.execute("INSERT INTO calculations (id, client_id, status) VALUES ($1, $2, $3)", 
                               calculation_id, db_client_id, "pending")

            task_id = str(uuid.uuid4())
            seen_urls = set()
            docs_payload = []

            for doc in request.documents:
                doc_url = (doc.url or "").strip()
                if not doc_url or doc_url in seen_urls:
                    continue
                seen_urls.add(doc_url)
                # Готовим данные для вставки
                docs_payload.append((
                    task_id, calculation_id, doc_url, doc.name, "pending", json.dumps({"stage": "queued"})
                ))

            if not docs_payload:
                raise HTTPException(status_code=400, detail="В запросе нет валидных документов.")

            # Пакетная вставка документов с защитой от 409 (DO NOTHING)
            await conn.executemany("""
                INSERT INTO documents (task_id, calculation_id, document_url, document_name, status, extracted_data)
                VALUES ($1, $2, $3, $4, $5, $6::jsonb)
                ON CONFLICT (task_id) DO NOTHING
            """, docs_payload)

        await send_to_queue(task_id, request.mode)

        return {
            "batch_id": calculation_id,
            "task_id": task_id,
            "documents_total": len(docs_payload),
            "message": f"Принято {len(docs_payload)} документов в пакет. Режим: {request.mode}",
        }

    except Exception as e:
        print(f"[API ERROR] {str(e)}")
        raise HTTPException(status_code=500, detail=f"Ошибка БД: {str(e)}")


@app.get("/api/v1/tenders/status/{batch_id}", dependencies=[Depends(verify_token)])
async def check_status(batch_id: str):
    try:
        pool = app.state.pool
        async with pool.acquire() as conn:
            documents = await conn.fetch("SELECT task_id, document_name, status, extracted_data::text FROM documents WHERE calculation_id = $1", batch_id)

        if not documents:
            raise HTTPException(status_code=404, detail="Расчет не найден")

        statuses = [doc["status"] for doc in documents]
        total_docs = len(documents)
        processing_statuses = {"pending", "processing"}
        done_docs = sum(1 for s in statuses if s not in processing_statuses)
        failed_docs = sum(1 for s in statuses if s == "error")
        parsed_docs = sum(1 for s in statuses if s == "completed")

        docs_state = []
        final_result = None
        task_id = documents[0]["task_id"] if documents else None

        for doc in documents:
            extracted_raw = doc["extracted_data"]
            extracted = json.loads(extracted_raw) if extracted_raw else {}
            
            if "final_result" in extracted:
                final_result = extracted.get("final_result")

            docs_state.append({
                "document_name": doc["document_name"],
                "status": doc["status"],
                "error": extracted.get("error"),
            })

        progress = {
            "total_docs": total_docs, "done_docs": done_docs,
            "parsed_docs": parsed_docs, "failed_docs": failed_docs,
        }

        if any(s in processing_statuses for s in statuses):
            return {
                "status": "processing", "task_id": task_id, "progress": progress,
                "documents": docs_state, "message": "Документы еще обрабатываются..."
            }

        if failed_docs > 0:
            return {
                "status": "error", "task_id": task_id, "progress": progress,
                "documents": docs_state, "message": "Часть документов не обработана."
            }

        if final_result is None:
            return {
                "status": "processing", "task_id": task_id, "progress": progress,
                "documents": docs_state, "message": "Документы обработаны, ожидается финальный расчет."
            }

        try:
            async with pool.acquire() as conn:
                await conn.execute("UPDATE calculations SET status = 'completed' WHERE id = $1", batch_id)
        except Exception:
            pass

        return {
            "status": "completed", "task_id": task_id, "progress": progress,
            "documents": docs_state, "data": final_result,
        }

    except HTTPException:
        raise
    except Exception as e:
        print(f"[API ERROR] 500: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))