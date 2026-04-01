# Docker 部署指南

## 快速开始

### 1. 初始化设置

首次使用前，需要运行初始化脚本：

```bash
./docker-deploy.sh setup
```

这会创建：
- `.env` 配置文件
- 项目目录结构
- 示例文档

### 2. 配置 API Keys

编辑 `.env` 文件，填入您的 API Keys：

```bash
# MiniMax API Key
MINIMAX_API_KEY=sk-your-minimax-api-key-here

# Aliyun API Key
ALIYUN_API_KEY=sk-your-aliyun-api-key-here

# ComfyUI 服务地址
# 如果使用 Docker 运行 ComfyUI，使用：http://comfyui:8188
# 如果使用宿主机的 ComfyUI，使用：http://host.docker.internal:8188
COMFYUI_URL=http://host.docker.internal:8188
```

### 3. 构建镜像

```bash
./docker-deploy.sh build
```

### 4. 启动服务

```bash
./docker-deploy.sh start
```

访问 http://localhost:8501

## 命令参考

| 命令 | 说明 |
|------|------|
| `./docker-deploy.sh setup` | 初始化设置 |
| `./docker-deploy.sh build` | 构建 Docker 镜像 |
| `./docker-deploy.sh start` | 启动服务 |
| `./docker-deploy.sh stop` | 停止服务 |
| `./docker-deploy.sh restart` | 重启服务 |
| `./docker-deploy.sh status` | 查看服务状态 |
| `./docker-deploy.sh logs` | 查看日志 |
| `./docker-deploy.sh logs notebooklm-ppt` | 查看应用日志 |
| `./docker-deploy.sh clean` | 清理容器和镜像 |

## 使用本地 ComfyUI

如果您已经在本地运行了 ComfyUI，需要在 `.env` 中配置：

```bash
COMFYUI_URL=http://host.docker.internal:8188
```

**注意**：`host.docker.internal` 在以下平台可用：
- Docker Desktop for Mac
- Docker Desktop for Windows
- Docker Desktop for Linux (需要添加 `--add-host=host.docker.internal:host-gateway`)

## 使用 Docker Compose 运行 ComfyUI

如果需要同时运行 ComfyUI，编辑 `docker-compose.yml`：

1. 取消 `comfyui` 服务的注释
2. 如果有 NVIDIA GPU，取消 `deploy` 部分的注释
3. 修改 `COMFYUI_URL` 为 `http://comfyui:8188`

然后启动：

```bash
docker-compose up -d
```

## 数据持久化

以下数据通过 volume 持久化：

- `./projects/` - 项目文件（原文档、优化稿、图片、最终文档）
- `./.env` - 环境配置（只读）
- `./comfyui/` - ComfyUI 模型和输出（如果运行 ComfyUI）

## 故障排查

### 查看日志

```bash
# 查看所有服务日志
./docker-deploy.sh logs

# 查看应用日志
./docker-deploy.sh logs notebooklm-ppt

# 查看 ComfyUI 日志
./docker-deploy.sh logs comfyui
```

### 容器无法启动

```bash
# 检查容器状态
./docker-deploy.sh status

# 查看详细错误
docker-compose logs notebooklm-ppt
```

### API 调用失败

1. 检查 `.env` 中的 API Key 是否正确
2. 检查网络连接
3. 查看日志确认错误信息

### ComfyUI 连接失败

1. 确认 ComfyUI 服务正在运行
2. 检查 `COMFYUI_URL` 配置是否正确
3. 如果在 Docker 中运行，确保网络互通

## 更新

```bash
# 拉取最新代码
git pull

# 重新构建
./docker-deploy.sh build

# 重启
./docker-deploy.sh restart
```

## 卸载

```bash
# 停止并清理
./docker-deploy.sh clean

# 删除项目文件（可选）
rm -rf projects/
rm -rf comfyui/
```
