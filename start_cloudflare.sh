#!/bin/bash
# Cloudflare Tunnel启动脚本

echo "启动Flask应用..."
source venv/bin/activate
python app.py &
FLASK_PID=$!

echo "等待Flask应用启动..."
sleep 3

echo "启动Cloudflare Tunnel..."
# 安装cloudflared（如果还没有）
if ! command -v cloudflared &> /dev/null; then
    echo "安装cloudflared..."
    if [[ "$OSTYPE" == "darwin"* ]]; then
        # macOS
        brew install cloudflared
    else
        echo "请手动安装cloudflared: https://developers.cloudflare.com/cloudflare-one/connections/connect-apps/install-and-setup/installation/"
        exit 1
    fi
fi

# 启动tunnel（需要先登录cloudflare）
cloudflared tunnel --url http://localhost:8080

# 清理函数
cleanup() {
    echo "停止服务..."
    kill $FLASK_PID 2>/dev/null
    pkill -f cloudflared 2>/dev/null
    exit 0
}

# 捕获中断信号
trap cleanup SIGINT SIGTERM

# 等待
wait
