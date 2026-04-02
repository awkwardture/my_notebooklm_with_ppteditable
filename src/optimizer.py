from src.minimax_client import generate_text as generate_text_minimax
from src.aliyun_client import generate_text as generate_text_aliyun
from src.prompts import OPTIMIZE_SYSTEM_PROMPT, STYLE_SYSTEM_PROMPT
from src.template_renderer import extract_variables_from_content

# 可用的文本生成模型（优化稿生成）
TEXT_MODELS = {
    "MiniMax-M2.7-highspeed": {
        "name": "MiniMax-M2.7-highspeed",
        "client": "minimax",
        "description": "MiniMax 高速推理模型",
    },
    "glm-5": {
        "name": "glm-5",
        "client": "aliyun",
        "description": "阿里云 GLM-5 模型",
    },
    "qwen3-max-2026-01-23": {
        "name": "qwen3-max-2026-01-23",
        "client": "aliyun",
        "description": "阿里云 Qwen3-Max 模型",
    },
}

# 新增：PPT 风格模板变量提取配置
TEMPLATE_VARIABLES_CONFIG = {
    "extract_variables": True,  # 是否提取模板变量
    "variable_types": [
        "title",        # 页面标题
        "subtitle",     # 副标题
        "content",      # 主要内容
        "key_data",     # 关键数据
        "conclusion",   # 结论
        "chart_type",   # 图表类型
        "chart_data",   # 图表数据
    ]
}


def generate_text(model: str, system_prompt: str, user_prompt: str) -> str:
    """Generate text using the specified model."""
    model_info = TEXT_MODELS.get(model)
    if model_info is None:
        # 默认使用 MiniMax
        model_info = {"client": "minimax"}

    if model_info["client"] == "aliyun":
        return generate_text_aliyun(model=model, system_prompt=system_prompt, user_prompt=user_prompt)
    else:
        return generate_text_minimax(model=model, system_prompt=system_prompt, user_prompt=user_prompt)


def optimize_document(raw_md: str, model: str = "MiniMax-M2.7") -> tuple[str, str]:
    """Convert raw markdown to optimized slide document and style description.

    Returns:
        (optimized_md, style_md)
    """
    optimized_md = generate_text(
        model=model,
        system_prompt=OPTIMIZE_SYSTEM_PROMPT,
        user_prompt=f"请将以下文档优化为演示文档结构：\n\n{raw_md}",
    )

    style_md = generate_text(
        model=model,
        system_prompt=STYLE_SYSTEM_PROMPT,
        user_prompt=f"请根据以下演示文档内容，生成 PPT 样式风格描述：\n\n{optimized_md}",
    )

    return optimized_md, style_md


def optimize_document_with_variables(raw_md: str, model: str = "MiniMax-M2.7") -> dict:
    """Convert raw markdown to optimized slide document with extracted variables.

    此函数不仅生成优化稿和风格描述，还会提取可用于模板渲染的变量。

    Returns:
        {
            "optimized_md": str,      # 优化后的文档
            "style_md": str,          # 风格描述
            "slides": list,           # 分页后的内容
            "variables": list,        # 每页的变量提取结果
        }
    """
    optimized_md, style_md = optimize_document(raw_md, model)
    slides = parse_slides(optimized_md)

    result = {
        "optimized_md": optimized_md,
        "style_md": style_md,
        "slides": slides,
        "variables": []
    }

    # 为每一页提取变量
    for slide_content in slides:
        variables = extract_variables_from_content(slide_content)
        result["variables"].append(variables)

    return result


def extract_page_variables(page_content: str) -> dict:
    """Extract variables from a single page content.

    用于在优化稿编辑时提取变量，以便后续选择模板渲染。

    Args:
        page_content: 单页内容

    Returns:
        变量字典，包含 title, subtitle, content_points 等
    """
    return extract_variables_from_content(page_content)


def parse_slides(optimized_md: str) -> list[str]:
    """Split optimized markdown into individual slide contents."""
    slides = []
    for part in optimized_md.split("---"):
        content = part.strip()
        if content:
            slides.append(content)
    return slides
