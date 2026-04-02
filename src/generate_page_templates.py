#!/usr/bin/env python3
"""
从 PPTX 文件生成带缩略图的页面级模板。
每页导出为图片，并使用 qwen3.5-plus 模型识别提取详细的风格描述。
使用 LibreOffice + pdftoppm 生成真实幻灯片截图。
"""

import os
import json
import glob
import subprocess
import tempfile
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE
import io

# 导入阿里云客户端
from aliyun_client import generate_text_with_images

# PPT 样例目录
ppt_dir = None
# 尝试多个可能的路径
for search_path in [
    os.path.join(os.path.dirname(__file__), "ppt*"),
    os.path.join(os.path.dirname(os.path.dirname(__file__)), "ppt*"),
]:
    for item in glob.glob(search_path):
        if os.path.isdir(item) and not item.endswith('.py'):
            ppt_dir = item.strip()
            break
    if ppt_dir:
        break

if not ppt_dir:
    raise Exception("PPT 样例 directory not found")

output_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), "page_template")
thumbnail_dir = os.path.join(output_dir, "thumbnails")
os.makedirs(thumbnail_dir, exist_ok=True)

# 布局类型映射
LAYOUT_CATEGORY_MAP = {
    'title': '封面标题页',
    'content': '内容页',
    'table': '表格页',
    'chart': '图表页',
    'bullets': '列表页',
}


def convert_pptx_to_pdf(pptx_path, output_dir):
    """使用 LibreOffice 将 PPTX 转换为 PDF"""
    pptx_path = os.path.abspath(pptx_path).strip()
    output_dir = output_dir.strip()

    # 根据系统选择 LibreOffice 路径
    import platform
    if platform.system() == 'Windows':
        soffice_path = 'C:\\Program Files\\LibreOffice\\program\\soffice.exe'
    else:
        soffice_path = '/Applications/LibreOffice.app/Contents/MacOS/soffice'

    cmd = [
        soffice_path,
        '--headless', '--convert-to', 'pdf',
        pptx_path, '--outdir', output_dir
    ]
    result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)

    pdf_path = os.path.join(output_dir, os.path.basename(pptx_path).replace('.pptx', '.pdf'))
    return pdf_path if os.path.exists(pdf_path) else None


def extract_pdf_pages_to_images(pdf_path, output_prefix, size=(800, 450)):
    """使用 pdftoppm 将 PDF 每页转为图片"""
    pdf_path = pdf_path.strip()
    output_prefix = output_prefix.strip()

    # 根据系统选择 pdftoppm 路径
    import platform
    if platform.system() == 'Windows':
        # Windows 需要指定 poppler 路径
        pdftoppm_path = r'D:\迅雷下载\Release-25.12.0-0\poppler-25.12.0\Library\bin\pdftoppm.exe'
    else:
        pdftoppm_path = 'pdftoppm'

    cmd = [pdftoppm_path, '-png', '-scale-to', str(size[0]), pdf_path, output_prefix]
    result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)

    # 返回生成的图片列表
    images = []
    idx = 1
    while True:
        img_path = f"{output_prefix}-{idx:02d}.png"
        if os.path.exists(img_path):
            images.append(img_path)
            idx += 1
        else:
            break
    return images


def analyze_slide(slide, slide_num, filename):
    """分析单页幻灯片"""
    elements = {
        'has_title': False,
        'has_subtitle': False,
        'has_text_boxes': 0,
        'has_bullets': False,
        'has_table': False,
        'table_structure': None,
        'has_chart': False,
        'chart_type': None,
        'has_image': False,
        'has_shape': False,
        'shape_count': 0,
    }

    # 分析标题
    title = slide.shapes.title if slide.shapes.title else None
    if title and title.text.strip():
        elements['has_title'] = True

    # 分析 shapes
    for shape in slide.shapes:
        try:
            if shape.has_table:
                elements['has_table'] = True
                elements['table_structure'] = {
                    'rows': len(shape.table.rows),
                    'cols': len(shape.table.columns),
                }
            elif shape.has_chart:
                elements['has_chart'] = True
                elements['chart_type'] = str(shape.chart.chart_type).split('.')[-1]
            elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                elements['has_image'] = True
            elif shape.has_text_frame:
                text = shape.text_frame.text.strip()
                if text:
                    elements['has_text_boxes'] += 1
                    for paragraph in shape.text_frame.paragraphs:
                        if paragraph.level > 0:
                            elements['has_bullets'] = True
                            break
            elif shape.shape_type in [1, 2, 3, 4, 5]:
                elements['has_shape'] = True
                elements['shape_count'] += 1
        except:
            continue

    # 确定布局类型
    if elements['has_title'] and not elements['has_table'] and not elements['has_chart'] and elements['has_text_boxes'] <= 2:
        layout_category = 'title'
    elif elements['has_table']:
        layout_category = 'table'
    elif elements['has_chart']:
        layout_category = 'chart'
    elif elements['has_bullets']:
        layout_category = 'bullets'
    else:
        layout_category = 'content'

    return elements, layout_category


def generate_style_description(filename, slide_num, elements, layout_category, thumbnail_path=None):
    """使用 qwen3.5-plus 模型分析缩略图，生成真正独特的风格描述"""

    # 如果没有缩略图，返回基础描述
    if not thumbnail_path or not os.path.exists(thumbnail_path):
        return generate_basic_description(filename, slide_num, elements, layout_category)

    # 布局类型中文映射
    layout_map = {
        'title': '封面标题页',
        'content': '内容页',
        'table': '表格页',
        'chart': '图表页',
        'bullets': '列表页',
    }
    layout_category_cn = layout_map.get(layout_category, layout_category)

    # 构建元素信息供模型参考
    elements_info = f"""
页面元素分析：
- 是否有标题: {elements['has_title']}
- 是否有副标题: {elements['has_subtitle']}
- 文本框数量: {elements['has_text_boxes']}
- 是否有列表: {elements['has_bullets']}
- 是否有表格: {elements['has_table']}
- 表格结构: {elements['table_structure'] if elements['table_structure'] else '无'}
- 是否有图表: {elements['has_chart']}
- 图表类型: {elements['chart_type'] if elements['chart_type'] else '无'}
- 是否有图片: {elements['has_image']}
- 是否有形状: {elements['has_shape']}
- 形状数量: {elements['shape_count']}
"""

    system_prompt = """你是专业的PPT视觉风格分析专家。请仔细观察PPT截图，按照以下结构提取该页的独特视觉风格要素：

# PPT 样式风格描述

## 整体风格
- **定位**：分析该页的整体设计定位（如企业汇报、商务提案、技术展示等）
- **基调**：描述整体基调（如专业、简洁、活泼、严肃等）
- **气质**：描述视觉气质和传递的品牌形象

## 主色调
请观察截图中的实际颜色，列出：
| 类型 | 色值（尽量准确估计） | 用途 |

## 字体风格
观察文字的字体、字号、字重特征

## 背景风格
描述背景的颜色、图案、渐变、装饰元素

## 图形元素
描述线条、形状、图标等视觉元素的风格特征

## 排版风格
描述页面布局方式、对齐方式、留白比例

## 本页视觉元素
**页面类型**：[指定类型]
**视觉构成**：描述该页具体的视觉元素组成

## 本页图表视觉建议
针对该页类型给出具体的图表/表格视觉建议

## 本页内容结构
描述该页的内容组织结构

## 总结
一句话总结该页的视觉呈现特点

请确保每项描述都基于截图的实际内容，不要使用模板化的通用描述。"""

    user_prompt = f"""请分析这张PPT截图，提取其独特的视觉风格要素。

来源文件：{filename}
页码：{slide_num}
布局类型：{layout_category_cn}

参考元素信息（已通过程序分析）：
{elements_info}

请仔细观察截图，按照指定结构输出风格描述，确保描述符合该页的实际视觉效果。"""

    try:
        # 调用阿里云 qwen3.5-plus 识别图片
        style_desc = generate_text_with_images(
            model="qwen3.5-plus",
            system_prompt=system_prompt,
            user_prompt=user_prompt,
            image_paths=[thumbnail_path]
        )
        return style_desc
    except Exception as e:
        print(f"    [WARN] 模型分析失败 ({filename} p{slide_num}): {e}")
        return generate_basic_description(filename, slide_num, elements, layout_category)


def generate_basic_description(filename, slide_num, elements, layout_category):
    """生成基础风格描述（备用方案）"""

    layout_map = {
        'title': '封面标题页',
        'content': '内容页',
        'table': '表格页',
        'chart': '图表页',
        'bullets': '列表页',
    }

    visual_elements = []
    if elements['has_title']:
        visual_elements.append("- 顶部标题区")
    if elements['has_table'] and elements['table_structure']:
        ts = elements['table_structure']
        visual_elements.append(f"- {ts['rows']}行×{ts['cols']}列表格")
    if elements['has_chart']:
        visual_elements.append(f"- {elements['chart_type']}图表")
    if elements['has_image']:
        visual_elements.append("- 装饰图片")
    if elements['has_bullets']:
        visual_elements.append("- 项目符号列表")
    if elements['has_text_boxes'] > 0:
        visual_elements.append(f"- {elements['has_text_boxes']} 个文本区域")

    visual_elements_str = "\n".join(visual_elements) if visual_elements else "- 简洁内容布局"

    return f"""# PPT 样式风格描述

## 整体风格
- **定位**：企业级汇报风格
- **基调**：专业、清晰

## 本页视觉元素

**页面类型**：{layout_map.get(layout_category, layout_category)}

**来源**：{filename.replace('.pptx', '')}

**视觉构成**：
{visual_elements_str}

## 总结
本页为{layout_map.get(layout_category, layout_category)}，内容层次清晰。"""


# 收集所有页面模板
all_page_templates = []

# 获取所有 PPTX 文件（排除临时文件）
ppt_files = [f.strip() for f in os.listdir(ppt_dir) if f.endswith('.pptx') and not f.startswith('~')]
print(f"找到 {len(ppt_files)} 个 PPTX 文件")

# 批量转换 PPTX 为 PDF（并行处理）
pdf_cache = {}
temp_base = tempfile.mkdtemp(prefix="pptx_batch_")
print(f"临时目录：{temp_base}")

import concurrent.futures

def convert_single_pptx(filename):
    """转换单个 PPTX 为 PDF"""
    filepath = os.path.join(ppt_dir, filename).strip()
    ppt_temp_dir = os.path.join(temp_base, filename.replace('.pptx', ''))
    os.makedirs(ppt_temp_dir, exist_ok=True)

    pdf_path = convert_pptx_to_pdf(filepath, ppt_temp_dir)
    return (filename, pdf_path)

# 并行转换所有 PPTX 文件
print(f"\n=== 并行转换 PPTX 为 PDF（{len(ppt_files)}个文件）===")
with concurrent.futures.ThreadPoolExecutor(max_workers=4) as executor:
    futures = {executor.submit(convert_single_pptx, fn): fn for fn in ppt_files}
    for future in concurrent.futures.as_completed(futures):
        filename, pdf_path = future.result()
        if pdf_path:
            pdf_cache[filename] = pdf_path
            print(f"  [OK] {filename}: PDF 已生成")
        else:
            print(f"  [FAIL] {filename}: PDF 生成失败")

# 将 PDF 页面转为图片（并行处理）
print(f"\n=== 并行生成缩略图 ===")

def process_pptx_thumbnails(filename, pdf_path):
    """处理单个 PPTX 的缩略图生成"""
    style_name = filename.replace('.pptx', '').replace('副本', '')[:50].strip()
    img_prefix = os.path.join(os.path.dirname(pdf_path), f"thumb_{style_name}")

    images = extract_pdf_pages_to_images(pdf_path, img_prefix)
    results = []

    for idx, img_path in enumerate(images):
        slide_num = idx + 1
        thumb_filename = f"{style_name}_p{slide_num}.jpg"
        thumb_path = os.path.join(thumbnail_dir, thumb_filename)

        # 转换 PNG 为 JPEG
        from PIL import Image
        img = Image.open(img_path)
        if img.mode in ('RGBA', 'P'):
            img = img.convert('RGB')
        img.save(thumb_path, quality=85)

        # 删除临时 PNG
        os.remove(img_path)
        results.append((slide_num, thumb_path))

    return (filename, len(results), results)

# 并行生成所有缩略图
with concurrent.futures.ThreadPoolExecutor(max_workers=4) as executor:
    futures = {executor.submit(process_pptx_thumbnails, fn, pdf): (fn, pdf) for fn, pdf in pdf_cache.items()}
    for future in concurrent.futures.as_completed(futures):
        filename, count, results = future.result()
        print(f"  [OK] {filename}: {count} 页缩略图")

# 分析 PPTX 并生成模板数据
print(f"\n=== 分析幻灯片基本元素 ===")

def analyze_pptx_slides_basic(filename):
    """分析单个 PPTX 的所有幻灯片基本元素（不生成风格描述）"""
    filepath = os.path.join(ppt_dir, filename).strip()
    style_name = filename.replace('.pptx', '').replace('副本', '')[:50].strip()

    templates = []
    prs = Presentation(filepath)

    for i, slide in enumerate(prs.slides):
        slide_num = i + 1

        # 分析幻灯片
        elements, layout_category = analyze_slide(slide, slide_num, filename)

        # 缩略图路径
        thumb_filename = f"{style_name}_p{slide_num}.jpg"
        thumb_path = os.path.join(thumbnail_dir, thumb_filename)
        thumb_relative = f"thumbnails/{thumb_filename}" if os.path.exists(thumb_path) else None

        # 构建页面模板（暂不含风格描述）
        page_template = {
            "id": f"{style_name}_page_{slide_num}",
            "source_file": filename,
            "source_name": style_name,
            "page_num": slide_num,
            "thumbnail": thumb_relative,
            "thumbnail_path": thumb_path,  # 保留绝对路径用于模型分析
            "layout_category": layout_category,
            "layout_category_cn": LAYOUT_CATEGORY_MAP.get(layout_category, layout_category),
            "style_description": None,  # 待模型填充
            "elements": elements,
            "colors": {
                "primary": "#0066CC",
                "secondary": "#66CCFF",
                "accent": "#FF9933",
            }
        }

        templates.append(page_template)

    return (filename, templates)

# 先分析基本元素
all_page_templates = []
for fn in ppt_files:
    filename, templates = analyze_pptx_slides_basic(fn)
    all_page_templates.extend(templates)
    print(f"  [OK] {filename}: {len(templates)} 页基本元素已分析")

# 使用 qwen3.5-plus 并行生成风格描述（并发10个）
print(f"\n=== 使用 qwen3.5-plus 并行分析风格（并发10个）===")

def generate_style_for_page(page_template):
    """为单个页面生成风格描述"""
    try:
        style_desc = generate_style_description(
            page_template['source_file'],
            page_template['page_num'],
            page_template['elements'],
            page_template['layout_category'],
            page_template.get('thumbnail_path')
        )
        page_template['style_description'] = style_desc
        return (page_template['id'], True, None)
    except Exception as e:
        # 使用备用描述
        page_template['style_description'] = generate_basic_description(
            page_template['source_file'],
            page_template['page_num'],
            page_template['elements'],
            page_template['layout_category']
        )
        return (page_template['id'], False, str(e))

# 并发10个调用模型
with concurrent.futures.ThreadPoolExecutor(max_workers=10) as executor:
    futures = {executor.submit(generate_style_for_page, pt): pt for pt in all_page_templates}
    completed = 0
    for future in concurrent.futures.as_completed(futures):
        page_id, success, error = future.result()
        completed += 1
        if success:
            print(f"  [{completed}/{len(all_page_templates)}] [OK] {page_id}: 模型分析完成")
        else:
            print(f"  [{completed}/{len(all_page_templates)}] [WARN] {page_id}: 使用备用描述 ({error})")

# 清理临时路径字段（不保存到JSON）
for pt in all_page_templates:
    pt.pop('thumbnail_path', None)

# 清理临时目录
import shutil
try:
    shutil.rmtree(temp_base)
except:
    pass

# 保存所有页面模板
output_file = os.path.join(output_dir, "page_templates.json")
with open(output_file, 'w', encoding='utf-8') as f:
    json.dump(all_page_templates, f, ensure_ascii=False, indent=2)

print(f"\n========== 完成 ==========")
print(f"总计：{len(all_page_templates)} 页模板")
print(f"输出文件：{output_file}")
print(f"缩略图目录：{thumbnail_dir}")
