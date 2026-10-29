#!/bin/bash
# Script khởi chạy ứng dụng TranHR Tax Tool

DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$DIR"

# Kiểm tra virtualenv
if [ -d "venv" ]; then
    PYTHON_EXEC="./venv/bin/python3"
else
    PYTHON_EXEC="python3"
fi

echo "Đang khởi động Công Cụ Đối Soát Thuế..."
$PYTHON_EXEC main.py
