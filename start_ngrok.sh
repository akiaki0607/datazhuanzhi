#!/bin/bash
# ngrok启动脚本

echo "启动Flask应用..."
source venv/bin/activate
python app.py &
FLASK_PID=$!

echo "等待Flask应用启动..."
sleep 3

echo "启动ngrok隧道..."
# 下载ngrok（如果还没有）
if [ ! -f "ngrok" ]; then
    echo "下载ngrok..."
    if [[ "$OSTYPE" == "darwin"* ]]; then
        # macOS
        curl -L https://bin.equinox.io/c/bNyj1mQVY4c/ngrok-v3-stable-darwin-amd64.zip -o ngrok.zip
        unzip ngrok.zip
        rm ngrok.zip
    else
        echo "请手动下载ngrok: https://ngrok.com/download"
        exit 1
    fi
fi

# 启动ngrok（需要先注册ngrok账号并获取authtoken）
./ngrok http 8080

# 清理函数
cleanup() {
    echo "停止服务..."
    kill $FLASK_PID 2>/dev/null
    pkill -f ngrok 2>/dev/null
    exit 0
}

# 捕获中断信号
trap cleanup SIGINT SIGTERM

# 等待
wait
