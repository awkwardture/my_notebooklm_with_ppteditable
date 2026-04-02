"""基于风格模板的图片生成模块

功能：
1. 使用模板渲染风格描述
2. 支持变量替换
3. 生成用于 ComfyUI 的 prompt
"""

from src.template_renderer import get_template_manager, extract_variables_from_content
from src.prompts import TEMPLATE_BASED_IMAGE_PROMPT, SLIDE_IMAGE_PROMPT_TEMPLATE
from src.image_generator import generate_slide_image as generate_base_image
from src.comfyui_client import generate_image_comfyui

DEFAULT_TEMPLATE = "default"  # 默认模板名称


def generate_prompt_from_template(
    template_name: str,
    page_num: int,
    total_pages: int,
    page_content: str,
    variables: dict = None
) -> str:
    """从模板生成图片生成 prompt

    Args:
        template_name: 风格模板名称
        page_num: 页码
        total_pages: 总页数
        page_content: 页面内容
        variables: 变量字典（如不传则自动从内容提取）

    Returns:
        渲染后的 prompt 字符串
    """
    manager = get_template_manager()
    template = manager.get_template(template_name)

    if not template:
        # 模板不存在，使用基础 prompt
        from src.prompts import SLIDE_IMAGE_PROMPT_TEMPLATE
        return SLIDE_IMAGE_PROMPT_TEMPLATE.format(
            style_description="专业信息图风格",
            slide_content=page_content,
            page_num=page_num,
            total_pages=total_pages
        )

    # 获取页面模板
    slide_template = template.get_slide_template(page_num)
    if not slide_template:
        slide_template = template.get_slide_template(1) or {}

    # 提取变量
    if variables is None:
        variables = extract_variables_from_content(page_content)

    # 获取布局类型
    layout_type = slide_template.get("layout", "content_page")

    # 构建风格模板描述
    style_template = slide_template.get("style_template", {})
    style_desc = style_template.get("description", f"{template_name}风格")

    # 渲染 prompt
    content_points = variables.get("content_points", [])
    if isinstance(content_points, list):
        content_points_str = ", ".join(content_points[:3])
    else:
        content_points_str = str(content_points)[:200]

    prompt = TEMPLATE_BASED_IMAGE_PROMPT.format(
        style_template=style_desc,
        title=variables.get("title", ""),
        subtitle=variables.get("subtitle", ""),
        content_points=content_points_str,
        key_data=", ".join(variables.get("key_data", [])[:2]),
        layout_type=layout_type,
        page_num=page_num,
        total_pages=total_pages
    )

    return prompt


def generate_slide_image_with_template(
    template_name: str,
    page_num: int,
    total_pages: int,
    page_content: str,
    variables: dict = None,
    model: str = "z_image_turbo"
) -> bytes | None:
    """使用风格模板生成幻灯片图片

    Args:
        template_name: 风格模板名称
        page_num: 页码
        total_pages: 总页数
        page_content: 页面内容
        variables: 变量字典
        model: 图片生成模型

    Returns:
        生成的图片 bytes
    """
    # 生成 prompt
    prompt = generate_prompt_from_template(
        template_name=template_name,
        page_num=page_num,
        total_pages=total_pages,
        page_content=page_content,
        variables=variables
    )

    # 使用 ComfyUI 生成图片
    return generate_image_comfyui(
        prompt=prompt,
        width=1920,
        height=1080,
        steps=20 if model == "z_image_turbo" else 50,
        use_z_image_turbo=(model == "z_image_turbo"),
        use_qwen_2512=(model == "qwen_image_2512"),
        use_qwen_fast=(model == "qwen_image_fast"),
    )


def get_available_templates() -> list:
    """获取所有可用的风格模板列表

    Returns:
        模板名称列表
    """
    manager = get_template_manager()
    return manager.get_template_names()


def get_template_layouts(template_name: str) -> list:
    """获取指定模板的所有布局类型

    Args:
        template_name: 模板名称

    Returns:
        布局类型列表
    """
    manager = get_template_manager()
    template = manager.get_template(template_name)
    if not template:
        return []
    return template.get_layout_types()


def preview_template_description(template_name: str, page_num: int, variables: dict) -> str:
    """预览模板渲染后的描述

    Args:
        template_name: 模板名称
        page_num: 页码
        variables: 变量字典

    Returns:
        渲染后的描述字符串
    """
    manager = get_template_manager()
    return manager.render_page_description(template_name, page_num, variables)
