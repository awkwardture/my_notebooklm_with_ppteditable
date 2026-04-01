#!/bin/bash

set -e

# 颜色定义
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

# 脚本目录
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

# 打印帮助信息
print_help() {
    cat << EOF
${BLUE}NotebookLM to PPT Docker 部署脚本${NC}

用法：$0 [命令]

命令:
    build       构建 Docker 镜像
    start       启动服务
    stop        停止服务
    restart     重启服务
    status      查看服务状态
    logs        查看日志
    clean       清理容器和镜像
    setup       初始化设置（创建.env 和目录）
    help        显示此帮助信息

示例:
    $0 setup    # 首次使用，初始化设置
    $0 build    # 构建镜像
    $0 start    # 启动服务

EOF
}

# 初始化设置
do_setup() {
    echo -e "${BLUE}正在初始化设置...${NC}"

    # 创建.env 文件
    if [ ! -f ".env" ]; then
        echo -e "${YELLOW}创建.env 文件...${NC}"
        cp .env.example .env
        echo -e "${GREEN}✓ .env 文件已创建，请编辑.env 文件配置 API Keys${NC}"
    else
        echo -e "${YELLOW}✓ .env 文件已存在${NC}"
    fi

    # 创建项目目录
    echo -e "${YELLOW}创建项目目录...${NC}"
    mkdir -p projects/原文档
    mkdir -p projects/优化 PP 页文档
    mkdir -p projects/生成的图片
    mkdir -p projects/最终文档
    mkdir -p comfyui/models/checkpoints
    mkdir -p comfyui/models/clip
    mkdir -p comfyui/models/vae
    mkdir -p comfyui/output

    # 创建示例文档
    if [ ! -f "projects/原文档/示例.md" ]; then
        cat > projects/原文档/示例.md << 'EXAMPLE'
# 示例文档

这是一个示例 NotebookLM 文档。

## 主要内容

1. 第一点内容
2. 第二点内容
3. 第三点内容

## 总结

这里是总结内容。
EXAMPLE
        echo -e "${GREEN}✓ 示例文档已创建${NC}"
    fi

    echo -e "${GREEN}✓ 初始化完成！${NC}"
    echo -e "${YELLOW}请编辑.env 文件配置您的 API Keys${NC}"
}

# 构建 Docker 镜像
do_build() {
    echo -e "${BLUE}正在构建 Docker 镜像...${NC}"
    docker-compose build
    echo -e "${GREEN}✓ 构建完成！${NC}"
}

# 启动服务
do_start() {
    echo -e "${BLUE}正在启动服务...${NC}"

    # 检查.env 文件
    if [ ! -f ".env" ]; then
        echo -e "${RED}错误：.env 文件不存在！${NC}"
        echo -e "${YELLOW}请先运行：$0 setup${NC}"
        exit 1
    fi

    # 检查 API Keys 是否配置
    if grep -q "sk-your-" .env; then
        echo -e "${YELLOW}警告：检测到默认 API Key，请确保已配置正确的 API Keys${NC}"
    fi

    docker-compose up -d

    echo -e "${GREEN}✓ 服务已启动！${NC}"
    echo -e "${BLUE}访问地址：http://localhost:8501${NC}"
    echo -e "${YELLOW}查看日志：$0 logs${NC}"
}

# 停止服务
do_stop() {
    echo -e "${BLUE}正在停止服务...${NC}"
    docker-compose down
    echo -e "${GREEN}✓ 服务已停止${NC}"
}

# 重启服务
do_restart() {
    do_stop
    sleep 2
    do_start
}

# 查看状态
do_status() {
    echo -e "${BLUE}服务状态：${NC}"
    docker-compose ps
}

# 查看日志
do_logs() {
    docker-compose logs -f "${1:-}"
}

# 清理
do_clean() {
    echo -e "${YELLOW}警告：这将删除所有容器和镜像！${NC}"
    read -p "确认继续？(y/N): " confirm
    if [ "$confirm" = "y" ] || [ "$confirm" = "Y" ]; then
        echo -e "${BLUE}正在清理...${NC}"
        docker-compose down --rmi all
        docker-compose rm -f
        echo -e "${GREEN}✓ 清理完成${NC}"
    else
        echo -e "${YELLOW}已取消${NC}"
    fi
}

# 主程序
case "${1:-help}" in
    build)
        do_build
        ;;
    start)
        do_start
        ;;
    stop)
        do_stop
        ;;
    restart)
        do_restart
        ;;
    status)
        do_status
        ;;
    logs)
        do_logs "$2"
        ;;
    clean)
        do_clean
        ;;
    setup)
        do_setup
        ;;
    help|--help|-h)
        print_help
        ;;
    *)
        echo -e "${RED}未知命令：$1${NC}"
        print_help
        exit 1
        ;;
esac
