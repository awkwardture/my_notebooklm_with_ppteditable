"""PPT 风格模板渲染引擎

功能：
1. 加载和管理 PPT 风格模板（以页面为维度）
2. 支持变量化描述
3. 模板选择和变量替换渲染
4. 生成详细的视觉风格描述
"""

import os
import json
import glob
from typing import Optional

TEMPLATE_DIR = os.path.join(
    os.path.dirname(os.path.dirname(__file__)),
    "page_template"
)

# 布局类型中英文映射
LAYOUT_CATEGORY_CN = {
    "title": "封面标题页",
    "content": "内容页",
    "table": "表格页",
    "chart": "图表页",
    "bullets": "列表页",
}


def get_layout_category_cn(category: str) -> str:
    """获取布局类型的中文名称"""
    return LAYOUT_CATEGORY_CN.get(category, category)


class SlideTemplate:
    """单个页面的风格模板"""

    def __init__(self, slide_data: dict):
        self.page_num = slide_data.get("page_num", 1)
        self.layout_category = slide_data.get("layout_category", "content")  # title, content, table, chart, bullets
        self.layout_category_cn = get_layout_category_cn(self.layout_category)
        self.layout_name = slide_data.get("layout_name", "自定义布局")
        self.detailed_description = slide_data.get("detailed_description", "")
        self.style_descriptor = slide_data.get("style_descriptor", {})
        self.variables = slide_data.get("variables", {})
        self.render_template = slide_data.get("render_template", {})

    def get_colors(self) -> dict:
        """获取配色方案"""
        return self.style_descriptor.get("colors", {})

    def get_elements(self) -> dict:
        """获取元素配置"""
        return self.style_descriptor.get("elements", {})

    def get_table_structure(self) -> Optional[dict]:
        """获取表格结构（如果有）"""
        return self.style_descriptor.get("table_structure")

    def get_chart_type(self) -> Optional[str]:
        """获取图表类型（如果有）"""
        return self.style_descriptor.get("chart_type")

    def render(self, variables: dict = None) -> str:
        """渲染页面描述，替换变量

        Args:
            variables: 变量字典，如 {"title": "标题", "content_points": [...]}

        Returns:
            渲染后的详细描述字符串
        """
        # 使用 detailed_description 作为基础
        # 因为当前版本 detailed_description 不包含变量占位符，直接返回即可
        # 后续可以扩展支持变量替换
        return self.detailed_description


class StyleTemplate:
    """一个完整的 PPT 风格模板（包含多页）"""

    def __init__(self, template_data: dict):
        self.name = template_data.get("name", "unknown")
        self.source_file = template_data.get("source_file", "")
        self.description = template_data.get("description", "")
        self.total_slides = template_data.get("total_slides", 0)
        self.slides = [SlideTemplate(s) for s in template_data.get("slides", [])]

    def get_slide_template(self, page_num: int) -> Optional[SlideTemplate]:
        """获取指定页面的模板"""
        for slide in self.slides:
            if slide.page_num == page_num:
                return slide
        return None

    def get_slides_by_layout(self, layout_category: str) -> list:
        """根据布局类型筛选页面"""
        return [s for s in self.slides if s.layout_category == layout_category]

    def get_layout_categories(self) -> list:
        """获取所有布局类型"""
        return list(set(s.layout_category for s in self.slides))


class TemplateManager:
    """模板管理器类"""

    def __init__(self, template_dir: str = TEMPLATE_DIR):
        self.template_dir = template_dir
        self.templates: dict[str, StyleTemplate] = {}
        self._load_all_templates()

    def _load_all_templates(self):
        """加载所有模板文件"""
        if not os.path.exists(self.template_dir):
            os.makedirs(self.template_dir, exist_ok=True)
            return

        # 加载所有独立的模板文件（排除 all_templates.json 等汇总文件）
        for filepath in glob.glob(os.path.join(self.template_dir, "*.json")):
            if any(x in filepath for x in ["all_templates", "all_pages", "_analysis"]):
                continue

            try:
                with open(filepath, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                template = StyleTemplate(data)
                self.templates[template.name] = template
            except Exception as e:
                print(f"Error loading template {filepath}: {e}")

    def get_template_names(self) -> list:
        """获取所有模板名称列表"""
        return list(self.templates.keys())

    def get_template(self, name: str) -> Optional[StyleTemplate]:
        """根据名称获取模板"""
        return self.templates.get(name)

    def get_all_layout_categories(self) -> list:
        """获取所有可用的布局类型（中文）"""
        layouts = set()
        for template in self.templates.values():
            for cat in template.get_layout_categories():
                layouts.add(get_layout_category_cn(cat))
        # 按固定顺序返回
        order = ["封面标题页", "内容页", "表格页", "图表页", "列表页"]
        result = []
        for o in order:
            if o in layouts:
                result.append(o)
        for cat in sorted(layouts):
            if cat not in result:
                result.append(cat)
        return result

    def get_slides_by_layout(self, layout_category: str) -> list:
        """根据布局类型筛选所有模板的页面"""
        result = []
        for template in self.templates.values():
            result.extend(template.get_slides_by_layout(layout_category))
        return result

    def render_page_description(
        self,
        template_name: str,
        page_num: int,
        variables: dict = None
    ) -> str:
        """渲染指定模板页面的风格描述

        Args:
            template_name: 模板名称
            page_num: 页码
            variables: 变量字典

        Returns:
            渲染后的风格描述
        """
        template = self.get_template(template_name)
        if not template:
            return f"模板 '{template_name}' 不存在"

        slide = template.get_slide_template(page_num)
        if not slide:
            return f"第 {page_num} 页的模板不存在"

        return slide.render(variables)

    def add_template(self, name: str, template_data: dict):
        """添加新模板"""
        template = StyleTemplate(template_data)
        self.templates[name] = template

        # 保存到文件
        filepath = os.path.join(self.template_dir, f"{name}.json")
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(template_data, f, ensure_ascii=False, indent=2)

    def refresh(self):
        """重新加载所有模板"""
        self.templates.clear()
        self._load_all_templates()


# 全局模板管理器实例
_manager: Optional[TemplateManager] = None


def get_template_manager() -> TemplateManager:
    """获取全局模板管理器实例"""
    global _manager
    if _manager is None:
        _manager = TemplateManager()
    return _manager


def render_template(
    template_name: str,
    page_num: int,
    variables: dict = None
) -> str:
    """便捷函数：渲染模板描述

    Usage:
        description = render_template("科技蓝", 1, {"title": "工作汇报", "content": "主要内容"})
    """
    manager = get_template_manager()
    return manager.render_page_description(template_name, page_num, variables)


def list_templates() -> list:
    """便捷函数：列出所有模板名称"""
    manager = get_template_manager()
    return manager.get_template_names()


def list_layout_categories() -> list:
    """便捷函数：列出所有布局类型"""
    manager = get_template_manager()
    return manager.get_all_layout_categories()


def extract_variables_from_content(content: str) -> dict:
    """从优化稿内容中提取变量

    分析文档内容，提取可用于模板渲染的变量：
    - title: 标题
    - subtitle: 副标题
    - content_points: 内容要点列表
    - key_data: 关键数据
    - conclusion: 结论

    Args:
        content: 优化后的文档内容

    Returns:
        变量字典
    """
    variables = {
        "title": "",
        "subtitle": "",
        "content_points": [],
        "key_data": [],
        "conclusion": "",
        "style_suggestions": "",  # 视觉建议（不渲染到图上，仅作指引）
        "color_scheme": "",       # 配色建议（不渲染到图上，仅作指引）
    }

    lines = content.strip().split('\n')
    skip_mode = False  # 用于跳过视觉建议/配色建议的内容

    for line in lines:
        line_stripped = line.strip()
        if not line_stripped:
            continue

        # 跳过页码行
        if line_stripped.startswith('**页码**') or line_stripped.startswith('- **页码**'):
            continue

        # 跳过视觉建议行
        if line_stripped.startswith('- **视觉建议**') or line_stripped.startswith('**视觉建议**') or line_stripped.startswith('视觉建议'):
            skip_mode = True
            # 提取建议内容作为风格指引，不作为页面内容
            suggestion_text = line_stripped.split('：', 1)[-1].strip() if '：' in line_stripped else ''
            variables["style_suggestions"] += suggestion_text + " "
            continue

        # 跳过配色建议行
        if line_stripped.startswith('- **配色建议**') or line_stripped.startswith('**配色建议**') or line_stripped.startswith('配色建议'):
            skip_mode = True
            # 提取配色内容作为风格指引，不作为页面内容
            color_text = line_stripped.split('：', 1)[-1].strip() if '：' in line_stripped else ''
            variables["color_scheme"] += color_text + " "
            continue

        # 如果是新的内容部分，退出跳过模式
        if line_stripped.startswith('# ') or line_stripped.startswith('## ') or line_stripped.startswith('- '):
            if skip_mode and not line_stripped.startswith('- '):
                skip_mode = False

        # 提取标题
        if line_stripped.startswith('# '):
            variables["title"] = line_stripped[2:].strip()
            skip_mode = False
        elif line_stripped.startswith('## '):
            variables["subtitle"] = line_stripped[3:].strip()
            skip_mode = False
        # 提取要点（跳过视觉建议/配色建议后的内容）
        elif line_stripped.startswith('- ') and not skip_mode:
            # 跳过已经单独处理的视觉建议/配色建议
            if '**视觉建议**' not in line_stripped and '**配色建议**' not in line_stripped:
                variables["content_points"].append(line_stripped[1:].strip())
        # 提取关键数据（包含数字的句子，但跳过配色建议中的色值）
        elif any(c.isdigit() for c in line_stripped) and len(line_stripped) < 100 and not skip_mode:
            # 跳过纯配色代码行
            if not line_stripped.startswith('#'):
                variables["key_data"].append(line_stripped)
        # 提取结论
        elif '结论' in line_stripped or '总结' in line_stripped or '总之' in line_stripped:
            variables["conclusion"] = line_stripped

    return variables


def extract_page_variables(page_content: str) -> dict:
    """从单页内容中提取变量

    用于在优化稿编辑时提取变量，以便后续选择模板渲染。

    Args:
        page_content: 单页内容

    Returns:
        变量字典
    """
    return extract_variables_from_content(page_content)