#!/bin/bash
RUN_DIR=/home/AutoWord/run

stop_one () {
  local name="$1"
  local pid_file="$RUN_DIR/$name.pid"
  if [ -f "$pid_file" ]; then
    PID=$(cat "$pid_file")
    if kill -0 "$PID" 2>/dev/null; then
      echo "停止 $name (PID $PID)..."
      kill "$PID"
      sleep 1
      kill -9 "$PID" 2>/dev/null || true
      echo "已停止 $name。"
    else
      echo "$name 的 PID 文件存在但进程不在运行。"
    fi
    rm -f "$pid_file"
  else
    echo "未找到 $name 的 PID 文件。"
  fi
}

echo "==== 按 PID 停止 AutoWord ===="
stop_one backend
stop_one frontend
echo "==== 完成 ===="
