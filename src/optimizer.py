from src.minimax_client import generate_text as generate_text_minimax
from src.aliyun_client import generate_text as generate_text_aliyun
from src.prompts import OPTIMIZE_SYSTEM_PROMPT, STYLE_SYSTEM_PROMPT

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


def parse_slides(optimized_md: str) -> list[str]:
    """Split optimized markdown into individual slide contents."""
    slides = []
    for part in optimized_md.split("---"):
        content = part.strip()
        if content:
            slides.append(content)
    return slides
