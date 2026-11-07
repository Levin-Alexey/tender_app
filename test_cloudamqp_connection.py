#!/usr/bin/env python3
"""Скрипт для проверки подключения к CloudAMQP"""

import asyncio
import os
import sys
from dotenv import load_dotenv

# Загрузить переменные окружения
load_dotenv()

CLOUDAMQP_URL = os.getenv('CLOUDAMQP_URL')

print("=" * 70)
print("🔍 ПРОВЕРКА ПОДКЛЮЧЕНИЯ К CLOUDAMQP (RabbitMQ)")
print("=" * 70)

# Проверить наличие URL
if not CLOUDAMQP_URL:
    print("\n❌ ОШИБКА: CLOUDAMQP_URL не установлен в .env")
    sys.exit(1)

# Скрыть чувствительные данные для вывода
url_display = CLOUDAMQP_URL.replace(CLOUDAMQP_URL.split('@')[0], '***')
print(f"\n📋 Конфигурация:")
print(f"  URL: {url_display}")

async def test_cloudamqp_connection():
    """Асинхронная проверка подключения к CloudAMQP"""
    try:
        import aio_pika
        
        print("\n⏳ Попытка подключения...")
        
        # Подключиться к RabbitMQ
        connection = await aio_pika.connect_robust(CLOUDAMQP_URL, timeout=10)
        print(f"✅ Подключение к серверу успешно!")
        
        # Получить канал
        print("\n🔗 Открытие канала...")
        channel = await connection.channel()
        print(f"✅ Канал открыт успешно!")
        
        # Объявить тестовую очередь
        queue_name = f"test-queue-{os.getpid()}"
        print(f"\n📦 Создание тестовой очереди '{queue_name}'...")
        queue = await channel.declare_queue(queue_name, auto_delete=True)
        print(f"✅ Очередь создана: {queue.name}")
        
        # Отправить тестовое сообщение
        test_message = "Test message from connection check"
        print(f"\n📤 Отправка тестового сообщения...")
        await channel.default_exchange.publish(
            aio_pika.Message(body=test_message.encode()),
            routing_key=queue_name
        )
        print(f"✅ Сообщение отправлено: '{test_message}'")
        
        # Получить сообщение
        print(f"\n📥 Получение сообщения...")
        message = await asyncio.wait_for(queue.get(), timeout=5)
        async with message.process():
            received_message = message.body.decode()
            print(f"✅ Сообщение получено: '{received_message}'")
            
            if received_message == test_message:
                print(f"   ✅ Содержимое совпадает!")
        
        # Получить статистику соединения
        print(f"\n📊 Информация о соединении:")
        print(f"  Connection: {connection}")
        print(f"  Channel: {channel}")
        
        # Закрыть соединение
        print(f"\n🔌 Закрытие соединения...")
        await connection.close()
        print(f"✅ Соединение закрыто")
        
        print("\n" + "=" * 70)
        print("✅ ВСЕ ПРОВЕРКИ ПРОЙДЕНЫ УСПЕШНО!")
        print("=" * 70)
        
    except asyncio.TimeoutError:
        print("\n❌ ОШИБКА: Timeout при подключении")
        print("   Проверьте:")
        print("   - URL CloudAMQP")
        print("   - Интернет соединение")
        print("   - Доступность сервера")
        sys.exit(1)
        
    except aio_pika.exceptions.AMQPConnectionError as e:
        print(f"\n❌ ОШИБКА ПОДКЛЮЧЕНИЯ (AMQP):")
        print(f"   {str(e)}")
        print("\n   Проверьте:")
        print("   - Правильность URL CloudAMQP")
        print("   - Учетные данные (username/password)")
        print("   - Доступность сервера RabbitMQ")
        sys.exit(1)
        
    except Exception as e:
        print(f"\n❌ ОШИБКА:")
        print(f"   {type(e).__name__}: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

# Запустить асинхронную функцию
if __name__ == "__main__":
    try:
        asyncio.run(test_cloudamqp_connection())
    except KeyboardInterrupt:
        print("\n\n⚠️  Прерывание пользователем")
        sys.exit(0)
