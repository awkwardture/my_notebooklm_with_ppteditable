from src.minimax_client import generate_text
from src.prompts import OPTIMIZE_SYSTEM_PROMPT, STYLE_SYSTEM_PROMPT

DEFAULT_MODEL = "MiniMax-M2.7"


def optimize_document(raw_md: str, model: str = DEFAULT_MODEL) -> tuple[str, str]:
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
        user_prompt=f"请根据以下演示文档内容，生成PPT样式风格描述：\n\n{optimized_md}",
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
