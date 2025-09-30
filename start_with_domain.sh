#!/bin/bash
# 使用自定义域名启动Excel转置处理工具

echo "🚀 启动Excel转置处理工具 (域名: aki-excel.liuchenglu.com)"
echo "=================================================="

# 检查Flask应用是否已经在运行
if lsof -i :8080 > /dev/null 2>&1; then
    echo "⚠️  端口8080已被占用，正在停止现有进程..."
    kill -9 $(lsof -ti :8080) 2>/dev/null
    sleep 2
fi

# 启动Flask应用
echo "📱 启动Flask应用..."
source venv/bin/activate
python app.py &
FLASK_PID=$!

echo "⏳ 等待Flask应用启动..."
sleep 3

# 检查Flask应用是否启动成功
if ! lsof -i :8080 > /dev/null 2>&1; then
    echo "❌ Flask应用启动失败！"
    exit 1
fi

echo "✅ Flask应用启动成功 (PID: $FLASK_PID)"
echo "🌐 本地访问地址: http://localhost:8080"

# 启动frp客户端
echo ""
echo "🔗 启动frp客户端..."
cd frp_0.52.3_darwin_amd64
./frpc -c ../frpc.ini &
FRP_PID=$!

echo "⏳ 等待frp连接建立..."
sleep 5

echo ""
echo "🎉 服务启动完成！"
echo "=================================================="
echo "📱 本地访问: http://localhost:8080"
echo "🌍 公网访问: https://aki-excel.liuchenglu.com"
echo "=================================================="
echo ""
echo "💡 提示:"
echo "- 按 Ctrl+C 停止所有服务"
echo "- 确保域名 aki-excel.liuchenglu.com 已正确解析到frp服务器"
echo ""

# 清理函数
cleanup() {
    echo ""
    echo "🛑 正在停止服务..."
    kill $FLASK_PID 2>/dev/null
    kill $FRP_PID 2>/dev/null
    pkill -f frpc 2>/dev/null
    echo "✅ 服务已停止"
    exit 0
}

# 捕获中断信号
trap cleanup SIGINT SIGTERM

# 等待
wait
