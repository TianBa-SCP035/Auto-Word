#!/bin/bash
# AutoWord 一键启动脚本（记录 PID）
set -e

# ====== 可调参数 ======
FRONTEND_PORT=8080
BACKEND_PORT=5000
BACKEND_DIR=/home/AutoWord/Code
FRONTEND_DIR=/home/AutoWord/dist
LOG_DIR=/home/AutoWord
RUN_DIR=/home/AutoWord/run
PYTHON_BIN=/home/project/software/anaconda3/bin/python   # 固定用 conda 的 python

# ====== 目录与提示 ======
mkdir -p "$RUN_DIR"
mkdir -p "$LOG_DIR"

echo "==== 启动 AutoWord 服务 ===="

# ====== 后端 ======
echo "[1/2] 启动后端..."
cd "$BACKEND_DIR"

# 如已有老进程，优雅处理（可选）
if [ -f "$RUN_DIR/backend.pid" ] && ps -p "$(cat "$RUN_DIR/backend.pid")" > /dev/null 2>&1; then
  echo "发现旧后端进程，PID=$(cat "$RUN_DIR/backend.pid")，无需重复启动。"
else
  nohup "$PYTHON_BIN" api.py > "$LOG_DIR/backend.log" 2>&1 &
  echo $! > "$RUN_DIR/backend.pid"
  echo "后端 PID: $(cat "$RUN_DIR/backend.pid") | 日志: $LOG_DIR/backend.log"
fi

# ====== 前端 ======
echo "[2/2] 启动前端 serve..."

if command -v serve >/dev/null 2>&1; then
  # 如已有老进程，优雅处理（可选）
  if [ -f "$RUN_DIR/frontend.pid" ] && ps -p "$(cat "$RUN_DIR/frontend.pid")" > /dev/null 2>&1; then
    echo "发现旧前端进程，PID=$(cat "$RUN_DIR/frontend.pid")，无需重复启动。"
  else
    nohup serve -s "$FRONTEND_DIR" -l "$FRONTEND_PORT" > "$LOG_DIR/frontend.log" 2>&1 &
    echo $! > "$RUN_DIR/frontend.pid"
    echo "前端 PID: $(cat "$RUN_DIR/frontend.pid") | 日志: $LOG_DIR/frontend.log"
  fi
else
  echo "未找到 'serve' 命令，请先安装：npm i -g serve"
  echo "或使用任意静态服务器手动托管 $FRONTEND_DIR"
fi

echo "==== 全部启动完成 ===="
echo "前端:  http://<服务器IP>:$FRONTEND_PORT"
echo "后端:  http://<服务器IP>:$BACKEND_PORT"
