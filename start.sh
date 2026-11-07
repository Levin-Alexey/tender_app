set -euo pipefail

cd "$(dirname "$0")"

if [[ -x "./venv/bin/python" ]]; then
    PYTHON="./venv/bin/python"
else
    PYTHON="${PYTHON:-python3}"
fi

"$PYTHON" -m uvicorn main:app --host "${HOST:-0.0.0.0}" --port "${PORT:-80}" &
API_PID=$!
"$PYTHON" worker.py &
WORKER_PID=$!

cleanup() {
    kill "$API_PID" "$WORKER_PID" 2>/dev/null || true
    wait "$API_PID" 2>/dev/null || true
    wait "$WORKER_PID" 2>/dev/null || true
}
trap cleanup EXIT INT TERM