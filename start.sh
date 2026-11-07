#!/bin/bash
source /root/tender_app/venv/bin/activate
uvicorn main:app --host 0.0.0.0 --port 80 &
python worker.py