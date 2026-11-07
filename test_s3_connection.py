#!/usr/bin/env python3
"""Скрипт для проверки подключения к S3"""

import os
import sys
from dotenv import load_dotenv
import boto3
from botocore.exceptions import ClientError, ConnectionError

# Загрузить переменные окружения
load_dotenv()

# Получить конфигурацию S3
S3_ENDPOINT_URL = os.getenv('S3_ENDPOINT_URL')
S3_ACCESS_KEY = os.getenv('S3_ACCESS_KEY')
S3_SECRET_KEY = os.getenv('S3_SECRET_KEY')
S3_BUCKET = os.getenv('S3_BUCKET')

print("=" * 60)
print("🔍 ПРОВЕРКА ПОДКЛЮЧЕНИЯ К S3")
print("=" * 60)

# Вывести конфигурацию (скрыть чувствительные данные)
print(f"\n📋 Конфигурация:")
print(f"  Endpoint: {S3_ENDPOINT_URL}")
print(f"  Access Key: {S3_ACCESS_KEY[:10]}...***")
print(f"  Secret Key: ***")
print(f"  Bucket: {S3_BUCKET}")

# Проверить наличие всех переменных
if not all([S3_ENDPOINT_URL, S3_ACCESS_KEY, S3_SECRET_KEY, S3_BUCKET]):
    print("\n❌ ОШИБКА: Не все переменные S3 установлены в .env")
    sys.exit(1)

try:
    print("\n⏳ Попытка подключения...")
    
    # Создать клиент S3
    s3_client = boto3.client(
        's3',
        endpoint_url=S3_ENDPOINT_URL,
        aws_access_key_id=S3_ACCESS_KEY,
        aws_secret_access_key=S3_SECRET_KEY,
        region_name='us-east-1'
    )
    
    # Проверить доступ к bucket'у
    print(f"\n🔗 Проверка доступа к бакету '{S3_BUCKET}'...")
    response = s3_client.head_bucket(Bucket=S3_BUCKET)
    
    print(f"✅ Подключение успешно!")
    print(f"   HTTP Status: {response['ResponseMetadata']['HTTPStatusCode']}")
    
    # Попробовать список объектов
    print(f"\n📁 Получение списка объектов в бакете...")
    response = s3_client.list_objects_v2(Bucket=S3_BUCKET, MaxKeys=5)
    
    if 'Contents' in response:
        print(f"   Найдено объектов: {response['KeyCount']}")
        print(f"   Первые объекты:")
        for obj in response['Contents'][:5]:
            print(f"     - {obj['Key']} ({obj['Size']} bytes)")
    else:
        print(f"   Бакет пуст или не содержит объектов")
    
    # Проверить возможность загрузки (создать тестовый файл)
    print(f"\n📤 Проверка возможности загрузки файла...")
    test_key = f"test-connection-{os.getpid()}.txt"
    test_content = b"Test connection file"
    
    s3_client.put_object(
        Bucket=S3_BUCKET,
        Key=test_key,
        Body=test_content
    )
    print(f"   ✅ Файл успешно загружен: {test_key}")
    
    # Проверить возможность чтения
    print(f"\n📥 Проверка возможности чтения файла...")
    response = s3_client.get_object(Bucket=S3_BUCKET, Key=test_key)
    content = response['Body'].read()
    
    if content == test_content:
        print(f"   ✅ Файл успешно прочитан, содержимое совпадает")
    else:
        print(f"   ⚠️  Содержимое не совпадает")
    
    # Удалить тестовый файл
    print(f"\n🗑️  Удаление тестового файла...")
    s3_client.delete_object(Bucket=S3_BUCKET, Key=test_key)
    print(f"   ✅ Тестовый файл удален")
    
    print("\n" + "=" * 60)
    print("✅ ВСЕ ПРОВЕРКИ ПРОЙДЕНЫ УСПЕШНО!")
    print("=" * 60)
    
except ClientError as e:
    error_code = e.response['Error']['Code']
    error_message = e.response['Error']['Message']
    print(f"\n❌ ОШИБКА S3 (ClientError):")
    print(f"   Код: {error_code}")
    print(f"   Сообщение: {error_message}")
    sys.exit(1)
    
except ConnectionError as e:
    print(f"\n❌ ОШИБКА ПОДКЛЮЧЕНИЯ:")
    print(f"   {str(e)}")
    print(f"\n   Проверьте:")
    print(f"   - URL endpoint: {S3_ENDPOINT_URL}")
    print(f"   - Интернет соединение")
    print(f"   - Доступность сервера S3")
    sys.exit(1)
    
except Exception as e:
    print(f"\n❌ НЕИЗВЕСТНАЯ ОШИБКА:")
    print(f"   {type(e).__name__}: {str(e)}")
    sys.exit(1)
