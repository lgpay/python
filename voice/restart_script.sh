#!/bin/bash

# 定义要监测的进程名称
PROCESS_NAME="python3 v2.py"

# 检查进程是否在运行
if pgrep -f "$PROCESS_NAME" > /dev/null; then
    echo "$PROCESS_NAME is running."
else
    echo "$PROCESS_NAME is not running. Restarting..."
    cd ~/voice && nohup python3 v2.py >/dev/null 2>&1 &
    echo "Process restarted."
fi
