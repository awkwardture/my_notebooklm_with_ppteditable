from src.comfyui_client import generate_image_comfyui
from src.prompts import SLIDE_IMAGE_PROMPT_TEMPLATE

DEFAULT_MODEL = "z_image_turbo"

# 支持的图片生成模型
IMAGE_MODELS = {
    "z_image_turbo": {
        "name": "ComfyUI Z-Image-Turbo (最快，推荐)",
        "description": "本地 ComfyUI，使用 Z-Image-Turbo 模型，速度快，效果好",
    },
    "qwen_image_2512": {
        "name": "ComfyUI Qwen Image 2512 (高质量中文)",
        "description": "本地 ComfyUI，使用 Qwen Image 2512 模型，中文支持最好，50 步",
    },
    "qwen_image_fast": {
        "name": "ComfyUI Qwen Image 2512 LoRA (快速)",
        "description": "本地 ComfyUI，使用 Qwen Image 2512 LoRA 版本，4 步快速生成",
    },
    "minimax": {
        "name": "MiniMax API (云端)",
        "description": "使用 MiniMax 云端 API 生成图片，需要网络",
    },
}

MAX_PROMPT_LENGTH = 1000


def generate_slide_image(
    slide_content: str,
    style_desc: str,
    page_num: int,
    total_pages: int,
    model: str = DEFAULT_MODEL,
) -> bytes | None:
    """Generate an infographic image for a single slide.

    Args:
        slide_content: The slide text content
        style_desc: Style description
        page_num: Current page number
        total_pages: Total number of pages
        model: Model to use - "z_image_turbo", "qwen_image_2512", "qwen_image_fast", or "minimax"

    Returns:
        Generated image as bytes, or None if failed
    """
    prompt = SLIDE_IMAGE_PROMPT_TEMPLATE.format(
        style_description=style_desc,
        page_num=page_num,
        slide_content=slide_content,
        total_pages=total_pages,
    )

    # 如果 prompt 超长，进行压缩
    if len(prompt) > MAX_PROMPT_LENGTH:
        template_overhead = len(SLIDE_IMAGE_PROMPT_TEMPLATE.format(
            style_description="", page_num=page_num, slide_content="", total_pages=total_pages
        ))
        available = MAX_PROMPT_LENGTH - template_overhead - 100

        style_len = min(len(style_desc), available // 2)
        content_len = available - style_len

        truncated_style = style_desc[:style_len] + "..." if len(style_desc) > style_len else style_desc
        truncated_content = slide_content[:content_len] + "..." if len(slide_content) > content_len else slide_content

        prompt = SLIDE_IMAGE_PROMPT_TEMPLATE.format(
            style_description=truncated_style,
            page_num=page_num,
            slide_content=truncated_content,
            total_pages=total_pages,
        )

    # 根据选择的模型生成图片
    if model == "z_image_turbo":
        # 使用 ComfyUI Z-Image-Turbo (最快)
        return generate_image_comfyui(
            prompt=prompt,
            width=1920,
            height=1080,
            steps=20,
            use_z_image_turbo=True,
            use_qwen_2512=False,
            use_qwen_fast=False,
            use_flux=False,
        )
    elif model == "qwen_image_2512":
        # 使用 ComfyUI Qwen Image 2512 (高质量中文)
        return generate_image_comfyui(
            prompt=prompt,
            width=1920,
            height=1080,
            steps=50,
            use_z_image_turbo=False,
            use_qwen_2512=True,
            use_qwen_fast=False,
            use_flux=False,
        )
    elif model == "qwen_image_fast":
        # 使用 ComfyUI Qwen Image 2512 LoRA (快速)
        return generate_image_comfyui(
            prompt=prompt,
            width=1920,
            height=1080,
            steps=4,
            use_z_image_turbo=False,
            use_qwen_2512=False,
            use_qwen_fast=True,
            use_flux=False,
        )
    elif model == "minimax":
        # 使用 MiniMax API (云端)
        from src.minimax_client import generate_image_minimax
        return generate_image_minimax(
            prompt=prompt,
            width=1920,
            height=1080,
        )
    else:
        # 默认使用 Z-Image-Turbo
        return generate_image_comfyui(
            prompt=prompt,
            width=1920,
            height=1080,
            steps=20,
            use_z_image_turbo=True,
            use_qwen_2512=False,
            use_qwen_fast=False,
            use_flux=False,
        )
