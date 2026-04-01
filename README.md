# Our's NotebookLM

基于 AI 的演示文档生成器 —— 粘贴文字，自动生成专业信息图 PDF / 可编辑 PPT。

## 功能

1. **智能优化** — 粘贴原始文本，AI 自动拆页、提炼要点
2. **风格生成** — 根据内容主题生成统一视觉风格描述
3. **信息图渲染** — 逐页调用 AI 生成信息图幻灯片
4. **导出 PDF** — 合并所有图片为 PDF，即刻下载分享
5. **生成 PPT** — AI 分析每页信息图，自动生成 python-pptx 代码并输出可编辑 PPTX 文件
   - 一键生成完整 PPT
   - 逐页生成 / 重新生成
   - 在线编辑 AI 生成的代码后直接运行
   - 合并已有页面为完整 PPT

## 技术栈

- **前端/应用框架**: Streamlit
- **文本生成模型**:
  - MiniMax-M2.7-highspeed — 文本优化（高速）
  - GLM-5 / Qwen3-Max — 阿里云模型，文本优化
- **PPT 代码生成**:
  - Qwen3.5-Plus / Qwen3-Max — 阿里云视觉模型，图片分析 & PPT 代码生成
- **图片生成**:
  - ComfyUI (本地) — Z-Image-Turbo / Qwen Image 2512
  - MiniMax API (云端) — 图片生成
- **PDF 合成**: img2pdf + Pillow
- **PPT 生成**: python-pptx (AI 生成代码 → 动态执行)

## 项目结构

```
├── app.py                 # 主应用入口
├── src/
│   ├── optimizer.py       # 文档优化 & 幻灯片拆分
│   ├── image_generator.py # 信息图生成
│   ├── pdf_builder.py     # PDF 合并
│   ├── ppt_generator.py   # PPT 代码生成 & 执行
│   ├── minimax_client.py  # MiniMax API 客户端
│   ├── aliyun_client.py   # 阿里云 API 客户端
│   ├── comfyui_client.py  # ComfyUI 本地图片生成
│   └── prompts.py         # Prompt 模板
├── projects/              # 用户项目数据 (自动生成)
├── requirements.txt
├── .env.example
├── Dockerfile
├── docker-compose.yml
├── docker-deploy.sh       # Docker 部署脚本
└── start.sh               # 一键启动脚本
```

## 快速启动

### 方式一：Docker 部署（推荐）

```bash
# 1. 初始化设置
./docker-deploy.sh setup

# 2. 编辑.env 文件配置 API Keys

# 3. 一键启动
./start.sh
```

访问 http://localhost:8501

详细 Docker 部署说明请参考 [DOCKER.md](DOCKER.md)

### 方式二：本地运行

#### 1. 克隆项目

```bash
git clone <repo-url>
cd my_notebookLM
```

### 2. 创建虚拟环境

```bash
python -m venv venv
source venv/bin/activate  # macOS/Linux
# Windows: venv\Scripts\activate
```

### 3. 安装依赖

```bash
pip install -r requirements.txt
```

### 4. 配置环境变量

```bash
cp .env.example .env
```

编辑 `.env` 文件，填入你的 API Keys：

```
# MiniMax API Key
MINIMAX_API_KEY=sk-your-minimax-api-key-here

# Aliyun API Key
ALIYUN_API_KEY=sk-your-aliyun-api-key-here

# ComfyUI 服务地址 (本地图片生成)
COMFYUI_URL=http://127.0.0.1:8188
```

### 5. 启动服务

```bash
streamlit run app.py
```

浏览器会自动打开 `http://localhost:8501`，即可开始使用。

## 使用流程

1. 在侧边栏创建一个新项目
2. **Step 1**: 粘贴或输入原始文档内容，点击「生成优化稿」
3. **Step 2**: 查看/编辑 AI 生成的优化稿和风格描述
4. **Step 3**: 点击「一键生成所有图片」，等待信息图生成；对不满意的页面可单独重新生成
5. **Step 4 - 导出**:
   - **PDF 标签页**: 点击「合并为 PDF」，下载最终文档
   - **PPT 标签页**: 点击「一键生成完整 PPT」自动生成可编辑 PPTX；也可逐页生成、编辑代码后重新运行，最后合并为完整 PPT

## 项目数据结构

每个项目目录包含：

```
projects/<项目名>/
├── 原文档/           # 原始文档及图片
├── 优化PP页文档/     # AI 生成的优化稿和风格描述
├── 生成的图片/       # AI 生成的信息图 (01.jpg, 02.jpg, ...)
└── 最终文档/         # 导出结果
    ├── <项目名>.pdf  # 合并后的 PDF
    ├── <项目名>.pptx # 完整 PPT
    └── ppt_slides/   # 逐页 PPT 代码和单页 PPTX
```



## 文章

本项目属于[《我写了一个"可编辑PPT版"的 NotebookLM》](https://mp.weixin.qq.com/s/wNE871I5w73aISuPQ0WxCA)的演示代码项目。

关注公众号获取更多内容:

**AI Native启示录**

<img src="images/qrcode.jpg" width="200" />
