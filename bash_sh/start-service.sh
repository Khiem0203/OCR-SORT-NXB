#!/bin/bash
# Khởi động OCR SORT BOOK

# Cấu hình

# Script nằm ở bash_sh/sort/, .env ở thư mục gốc OCR_SORT_NXB
ENV_FILE="$(cd "$(dirname "$0")/../.." && pwd)/.env"
if [ -f "$ENV_FILE" ]; then
    set -a
    source <(tr -d '\r' < "$ENV_FILE")
    set +a
    echo "Loaded .env from $ENV_FILE"
else
    echo "Warning: .env file not found at $ENV_FILE"
fi

# Số job chạy cùng lúc trong API (không cần RabbitMQ, PostgreSQL)
export SORT_WORKERS="${SORT_WORKERS:-4}"
echo "SORT_WORKERS: $SORT_WORKERS"

# Spellfix
SPELLFIX_GPU="${SPELLFIX_GPU:-0}"
SPELLFIX_DTYPE="${SPELLFIX_DTYPE:-float16}"
SPELLFIX_WAIT="${SPELLFIX_WAIT:-180}"
SPELLFIX_PYTHON="${SPELLFIX_PYTHON:-/home/your_user/anaconda3/envs/ocr/bin/python}"

# Môi trường

CONDA_PATH="/home/your_user/anaconda3/condabin/conda"
CONDA_ENV_NAME="ocr"
PROJECT_DIR="/home/your_user/OCR_SORT_NXB"

# Khởi động
LOG_DIR="$PROJECT_DIR/logs"
mkdir -p "$LOG_DIR/api"
mkdir -p "$PROJECT_DIR/temp_processing"
mkdir -p "$LOG_DIR/services"
mkdir -p "$PROJECT_DIR/results"

echo "Activating Conda environment: $CONDA_ENV_NAME"
if [ -f "$(dirname "$CONDA_PATH")/../etc/profile.d/conda.sh" ]; then
    source "$(dirname "$CONDA_PATH")/../etc/profile.d/conda.sh"
else
    echo "Error: conda.sh not found. Check your Conda installation."
    exit 1
fi

conda activate "$CONDA_ENV_NAME"
if [ $? -ne 0 ]; then echo "Error: Failed to activate Conda env '$CONDA_ENV_NAME'."; exit 1; fi

cd "$PROJECT_DIR" || { echo "Error: Could not change to project directory $PROJECT_DIR"; exit 1; }

# Hàm bật dịch vụ

start_fastapi_server() {
    local log_file="$LOG_DIR/api/api_server.log"
    echo "Starting FastAPI API Server on port 9091..."
    nohup /home/your_user/anaconda3/envs/ocr/bin/uvicorn app.main:app --host 0.0.0.0 --port 9091 > "$log_file" 2>&1 &
    echo "FastAPI server started. Log: $log_file. PID: $!"
}

# Bật dịch vụ model nếu port còn trống
start_model_service() {
    local name=$1 port=$2 gpu=$3 python_bin=$4 module=$5 log_file=$6
    if ss -ltn 2>/dev/null | grep -q ":${port} "; then
        echo "$name already running on port $port, skip."
        eval "${name}_PID="
        return
    fi
    echo "Starting $name on port $port (GPU $gpu, $python_bin)..."
    CUDA_VISIBLE_DEVICES="$gpu" nohup "$python_bin" -m uvicorn "$module" --host 0.0.0.0 --port "$port" > "$log_file" 2>&1 &
    eval "${name}_PID=$!"
    echo "$name started. Log: $log_file. PID: $!"
}

# Chờ dịch vụ nạp model xong
wait_service_ready() {
    local name=$1 port=$2 max_wait=$3 log_file=$4
    local pid_var="${name}_PID"; local pid="${!pid_var}"
    echo "Waiting for $name to load model (max ${max_wait}s)..."
    local waited=0
    while [ "$waited" -lt "$max_wait" ]; do
        if curl -sf --max-time 3 "http://localhost:${port}/health" > /dev/null; then
            echo "$name ready after ${waited}s: $(curl -s --max-time 3 "http://localhost:${port}/health")"
            return 0
        fi
        if [ -n "$pid" ] && ! kill -0 "$pid" 2>/dev/null; then
            echo "ERROR: $name đã dừng khi đang nạp model. Cuối log $log_file:"
            tail -n 15 "$log_file"
            return 0
        fi
        sleep 5
        waited=$((waited + 5))
    done
    echo "Warning: $name chưa sẵn sàng sau ${max_wait}s, vẫn tiếp tục khởi động (xem $log_file)."
    return 0
}

# Chạy
echo "========================================"
echo "  Starting OCR SORT BOOK Services"
echo "========================================"

# 1. Bật Spellfix
export SPELLFIX_DTYPE
start_model_service SPELLFIX 8006 "$SPELLFIX_GPU" "$SPELLFIX_PYTHON" \
    app.spellfix_api:app "$LOG_DIR/services/spellfix_service.log"
wait_service_ready SPELLFIX 8006 "$SPELLFIX_WAIT" "$LOG_DIR/services/spellfix_service.log"

# 2. Bật API (job chạy ngay trong API, SORT_WORKERS job cùng lúc)
start_fastapi_server

echo ""
echo "All application services have been started."
echo "--------------------------------"
echo "Resources Utilization"
echo "RAM Usage:"
free -h | awk '/^Mem:/ {print "  Used: "$3" / Total: "$2" (Free: "$4")"}'
echo "VRAM Usage (GPU):"
nvidia-smi --query-gpu=memory.total,memory.used,memory.free --format=csv,noheader,nounits | awk -F', ' '{print "  Total: "$1" MiB, Used: "$2" MiB, Free: "$3" MiB"}'
echo "--------------------------------"
echo "Job đang chờ / chạy: curl http://localhost:9091/api/jobs/status"
echo "Kết quả: $PROJECT_DIR/results/sort_YYYY-MM-DD.jsonl"
echo ""
echo "To stop the application services, run:"
echo "  bash_sh/sort/stop-service.sh"
echo "--------------------------------"
