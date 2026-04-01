#!/bin/bash

# 一键启动脚本
# 检查环境并启动服务

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

echo "=== NotebookLM to PPT 一键启动 ==="

# 检查 Docker
if ! command -v docker &> /dev/null; then
    echo "错误：未检测到 Docker，请先安装 Docker"
    exit 1
fi

if ! docker info &> /dev/null; then
    echo "错误：Docker 未运行，请启动 Docker"
    exit 1
fi

echo "✓ Docker 运行正常"

# 检查 Docker Compose
if ! command -v docker-compose &> /dev/null && ! docker compose version &> /dev/null; then
    echo "错误：未检测到 Docker Compose，请先安装"
    exit 1
fi

echo "✓ Docker Compose 就绪"

# 检查.env 文件
if [ ! -f ".env" ]; then
    echo "正在初始化..."
    ./docker-deploy.sh setup
fi

# 检查 API Keys
if grep -q "sk-your-" .env 2>/dev/null; then
    echo "⚠ 警告：检测到默认 API Key"
    echo "请编辑.env 文件配置您的 API Keys"
    echo ""
    read -p "是否继续启动？(y/N): " confirm
    if [ "$confirm" != "y" ] && [ "$confirm" != "Y" ]; then
        echo "已退出"
        exit 0
    fi
fi

# 启动服务
echo ""
echo "正在启动服务..."
./docker-deploy.sh start

echo ""
echo "=== 启动完成 ==="
echo "访问地址：http://localhost:8501"
echo "查看日志：./docker-deploy.sh logs"
echo "停止服务：./docker-deploy.sh stop"
