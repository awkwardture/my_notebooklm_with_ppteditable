import os
import shutil
import concurrent.futures
import streamlit as st
from src.optimizer import optimize_document, parse_slides, optimize_document_with_variables, extract_page_variables
from src.image_generator import generate_slide_image, IMAGE_MODELS, DEFAULT_MODEL as DEFAULT_IMAGE_MODEL
from src.pdf_builder import build_pdf
from src.template_renderer import list_templates, list_layout_categories, render_template, get_template_manager, get_layout_category_cn
from src.template_image_generator import generate_prompt_from_template
import json
from src.ppt_generator import (
    generate_slide_code, build_single_slide_pptx, build_single_slide_pptx_with_retry,
    build_full_pptx, save_slide_code, load_slide_code, load_all_slide_codes,
    get_single_pptx_path, test_slide_code,
)

PROJECTS_DIR = os.path.join(os.path.dirname(__file__), "projects")
os.makedirs(PROJECTS_DIR, exist_ok=True)

st.set_page_config(page_title="Our's NotebookLM", layout="wide")
st.title("Our's NotebookLM")

# ── Sidebar: project management ──────────────────────────────────────
with st.sidebar:
    st.header("项目管理")

    existing = sorted(
        d for d in os.listdir(PROJECTS_DIR)
        if os.path.isdir(os.path.join(PROJECTS_DIR, d))
    )

    new_name = st.text_input("新建项目名称")
    if st.button("创建项目") and new_name:
        proj = os.path.join(PROJECTS_DIR, new_name)
        for sub in ["原文档/images", "优化PP页文档", "生成的图片", "最终文档"]:
            os.makedirs(os.path.join(proj, sub), exist_ok=True)
        st.session_state["selected_project"] = new_name
        st.rerun()

    if not existing:
        st.info("请先创建一个项目")
        st.stop()

    default_idx = 0
    if "selected_project" in st.session_state and st.session_state["selected_project"] in existing:
        default_idx = existing.index(st.session_state["selected_project"])

    project_name = st.selectbox("选择项目", existing, index=default_idx)
    proj_dir = os.path.join(PROJECTS_DIR, project_name)

    # 删除项目功能
    st.divider()
    st.subheader("删除项目")
    confirm_delete = st.checkbox("确认删除当前项目", key="confirm_delete")
    if st.button("🗑️ 删除项目", disabled=not confirm_delete, type="secondary"):
        if os.path.exists(proj_dir):
            shutil.rmtree(proj_dir)
            if "selected_project" in st.session_state:
                del st.session_state["selected_project"]
            st.success(f"项目 '{project_name}' 已删除")
            st.rerun()

    st.divider()
    st.markdown("""
**🚀 功能说明**

1. **智能优化** — 粘贴原始文本，AI 自动拆页、提炼要点
2. **风格生成** — 根据内容主题生成统一视觉风格
3. **信息图渲染** — 逐页生成专业信息图幻灯片
4. **导出** — 合并为 PDF / AI 生成可编辑 PPT
    """)

    # 导入优化稿模型配置
    from src.optimizer import TEXT_MODELS

    # 模型设置
    st.sidebar.subheader("模型设置")

    # 优化稿模型选择
    text_model_options = {v["name"]: k for k, v in TEXT_MODELS.items()}
    selected_text_model_name = st.sidebar.selectbox(
        "优化稿模型",
        options=list(text_model_options.keys()),
        index=0,  # 默认选择 MiniMax-M2.7
        help="选择生成优化稿和风格描述的模型",
    )
    text_model = text_model_options[selected_text_model_name]

    # 图片生成模型选择
    image_model_options = {v["name"]: k for k, v in IMAGE_MODELS.items()}
    selected_image_model_name = st.sidebar.selectbox(
        "图片生成模型",
        options=list(image_model_options.keys()),
        index=0,  # 默认选择第一个 (Z-Image-Turbo)
        help="选择生成信息图图片的模型",
    )
    image_model = image_model_options[selected_image_model_name]

    # PPT 代码生成模型
    from src.ppt_generator import PPT_MODELS
    ppt_model_options = {v["name"]: k for k, v in PPT_MODELS.items()}
    selected_ppt_model_name = st.sidebar.selectbox(
        "PPT 代码生成模型",
        options=list(ppt_model_options.keys()),
        index=0,
        help="选择生成 PPT 代码的 Vision 模型",
    )
    ppt_model = ppt_model_options[selected_ppt_model_name]

    # 执行模式选择：图片生成和 PPT 生成独立设置
    st.sidebar.subheader("执行模式")
    image_execution_mode = st.sidebar.radio(
        "图片生成",
        options=["并行", "串行"],
        index=0,
        key="image_mode",
        help="并行：同时生成所有图片（速度快）；串行：逐页生成（节省资源）",
    )
    ppt_execution_mode = st.sidebar.radio(
        "PPT 生成",
        options=["并行", "串行"],
        index=0,
        key="ppt_mode",
        help="并行：同时生成所有 PPT 页（速度快）；串行：逐页生成（节省资源）",
    )

# ── Helper paths ─────────────────────────────────────────────────────
raw_path = os.path.join(proj_dir, "原文档", "原稿.md")
opt_path = os.path.join(proj_dir, "优化PP页文档", "优化稿.md")
style_path = os.path.join(proj_dir, "优化PP页文档", "ppt样式风格描述.md")
img_dir = os.path.join(proj_dir, "生成的图片")
pdf_path = os.path.join(proj_dir, "最终文档", f"{project_name}.pdf")
ppt_path = os.path.join(proj_dir, "最终文档", f"{project_name}.pptx")

# 模板配置文件路径
template_config_path = os.path.join(proj_dir, "优化PP页文档", "template_config.json")
page_templates_path = os.path.join(os.path.dirname(__file__), "page_template", "page_templates.json")
page_templates_cache = None
page_templates_mtime = 0


def load_page_templates() -> list:
    """加载所有页面模板"""
    global page_templates_cache, page_templates_mtime
    # 检查文件是否变化
    current_mtime = os.path.getmtime(page_templates_path) if os.path.exists(page_templates_path) else 0
    if page_templates_cache is not None and page_templates_mtime == current_mtime:
        return page_templates_cache
    if os.path.exists(page_templates_path):
        with open(page_templates_path, "r", encoding="utf-8") as f:
            page_templates_cache = json.load(f)
        page_templates_mtime = current_mtime
        return page_templates_cache
    return []


def save_page_style(project_name: str, page_idx: int, style_desc: str, template_id: str = ""):
    """保存某页的风格描述"""
    config_path = os.path.join(PROJECTS_DIR, project_name, "优化PP页文档", "page_styles.json")
    if os.path.exists(config_path):
        with open(config_path, "r", encoding="utf-8") as f:
            data = json.load(f)
    else:
        data = {"pages": {}}
    data["pages"][str(page_idx)] = {"style_description": style_desc, "template_id": template_id}
    os.makedirs(os.path.dirname(config_path), exist_ok=True)
    with open(config_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def load_page_styles(project_name: str) -> dict:
    """加载项目的页面风格配置"""
    config_path = os.path.join(PROJECTS_DIR, project_name, "优化PP页文档", "page_styles.json")
    if os.path.exists(config_path):
        with open(config_path, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"pages": {}}


def read_template_config() -> dict:
    """读取模板配置"""
    import json
    if os.path.exists(template_config_path):
        with open(template_config_path, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"selected_template": "", "page_templates": {}}


def write_template_config(config: dict):
    """写入模板配置"""
    import json
    os.makedirs(os.path.dirname(template_config_path), exist_ok=True)
    with open(template_config_path, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)


def read_file(path: str) -> str:
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            return f.read()
    return ""


def write_file(path: str, content: str):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        f.write(content)


# ── Step 1: Raw document ─────────────────────────────────────────────
st.header("Step 1: 原稿编辑")

raw_text = st.text_area(
    "输入或粘贴原始文档 (Markdown)",
    value=read_file(raw_path),
    height=300,
    key="raw_text",
)

col1, col2 = st.columns(2)
with col1:
    if st.button("保存原稿"):
        write_file(raw_path, raw_text)
        st.success("已保存")

with col2:
    if st.button("生成优化稿", type="primary"):
        if not raw_text.strip():
            st.warning("请先输入原稿内容")
        else:
            write_file(raw_path, raw_text)
            with st.spinner("正在生成优化稿和风格描述..."):
                opt_md, sty_md = optimize_document(raw_text, model=text_model)
            write_file(opt_path, opt_md)
            write_file(style_path, sty_md)
            
            # 解析每页风格并自动保存
            from src.optimizer import parse_page_styles
            slides = parse_slides(opt_md)
            page_styles_list = parse_page_styles(sty_md)
            
            # 调试输出
            import streamlit as st
            with st.expander("查看原始风格描述", expanded=False):
                st.code(sty_md[:2000])
            
            if page_styles_list:
                for item in page_styles_list:
                    page_num = item.get("page_num", 0)
                    style_desc = item.get("style_description", "")
                    if page_num > 0 and style_desc:
                        save_page_style(project_name, page_num, style_desc)
                st.success(f"优化稿生成完成！同时生成了 {len(page_styles_list)} 页独立风格")
            else:
                st.warning("未能解析风格描述，请检查 AI 输出格式")
                st.code(sty_md[:500])
            st.rerun()

# ── Step 2: Optimized document & style ───────────────────────────────
st.header("Step 2: 优化稿 & 风格描述")

# 读取模板配置（保留用于兼容）
template_config = read_template_config()

# 加载所有页面模板
page_templates = load_page_templates()

# 加载项目的页面风格配置（每次都重新加载，确保获取最新值）
page_styles = load_page_styles(project_name)

# 解析优化稿为单页
current_opt = read_file(opt_path)
slides = parse_slides(current_opt) if current_opt else []

tab_opt = st.container()

with tab_opt:
    if not slides:
        st.info("请先生成优化稿")
    else:
        st.info(f"共 {len(slides)} 页，请在下方逐页编辑内容和风格")

        # 为每页显示独立的编辑器
        for i, slide in enumerate(slides):
            page_idx = i + 1
            page_key = str(page_idx)

            # 获取该页已保存的风格配置
            saved_style = page_styles.get("pages", {}).get(page_key, {})
            current_style_desc = saved_style.get("style_description", "")
            selected_template_id = saved_style.get("template_id", "")
            global_style = read_file(style_path)

            # 检查是否有 pending 的风格描述（刚选择的模板）
            pending_key = f"pending_page_style_{page_idx}"
            text_area_key = f"page_style_{page_idx}"
            # 使用版本号来强制刷新 text_area
            version_key = f"page_style_version_{page_idx}"
            if pending_key in st.session_state:
                current_style_desc = st.session_state[pending_key]
                del st.session_state[pending_key]
                # 增加版本号，强制 text_area 使用新的 key
                st.session_state[version_key] = st.session_state.get(version_key, 0) + 1
                st.toast(f"已更新第 {page_idx} 页风格描述", icon="✅")

            # 获取当前版本号
            style_version = st.session_state.get(version_key, 0)
            # 使用带版本号的 key，确保更新时创建新的 widget
            dynamic_text_key = f"page_style_{page_idx}_v{style_version}"

            # 使用 expander 折叠每页内容
            with st.expander(f"第 {page_idx} 页", expanded=(page_idx == 1)):
                # 页面内容编辑
                page_content_key = f"page_content_{page_idx}"
                page_content_version_key = f"page_content_version_{page_idx}"
                
                # 检查是否有 pending 的内容更新
                pending_content_key = f"pending_page_content_{page_idx}"
                if pending_content_key in st.session_state:
                    slide = st.session_state[pending_content_key]
                    del st.session_state[pending_content_key]
                    st.session_state[page_content_version_key] = st.session_state.get(page_content_version_key, 0) + 1
                
                # 使用带版本号的 key 强制刷新
                content_version = st.session_state.get(page_content_version_key, 0)
                dynamic_content_key = f"page_content_{page_idx}_v{content_version}"
                
                page_content = st.text_area(
                    "页面内容（可编辑）",
                    value=slide,
                    height=300,
                    key=dynamic_content_key,
                )
                
                # 风格描述编辑
                style_label = "风格描述（可编辑）"
                if selected_template_id:
                    style_label = f"风格描述 ✅ 已选模板"

                page_style = st.text_area(
                    style_label,
                    value=current_style_desc if current_style_desc else global_style,
                    height=200,
                    key=dynamic_text_key,
                )

                # 保存按钮行
                col_save_content, col_save_style, col_select = st.columns(3)
                with col_save_content:
                    if st.button("保存内容", key=f"save_content_{page_idx}"):
                        # 保存该页内容到优化稿
                        all_slides = parse_slides(current_opt)
                        all_slides[page_idx - 1] = page_content
                        new_opt_md = "\n\n---\n\n".join(all_slides)
                        write_file(opt_path, new_opt_md)
                        current_opt = new_opt_md  # 更新当前缓存
                        st.session_state[f"pending_page_content_{page_idx}"] = page_content
                        st.success(f"第 {page_idx} 页内容已保存")
                        st.rerun()
                
                with col_save_style:
                    if st.button("保存风格", key=f"save_style_{page_idx}_btn"):
                        save_page_style(project_name, page_idx, page_style, selected_template_id)
                        st.success(f"第 {page_idx} 页风格已保存")
                        st.rerun()

                with col_select:
                    # 模板选择按钮
                    if st.button("选择模板", key=f"select_template_{page_idx}", type="secondary"):
                        st.session_state[f"show_template_selector_{page_idx}"] = True

                # 显示模板选择器
                if st.session_state.get(f"show_template_selector_{page_idx}", False):
                    st.divider()
                    st.markdown(f"### 第 {page_idx} 页 - 选择模板")

                    # 获取该页的布局类型（用于筛选）
                    # 简单判断：如果有项目符号就是列表页，有表格就是表格页，等等
                    slide_lower = slide.lower()
                    layout_filter = "全部"
                    if "表格" in slide or "table" in slide_lower:
                        layout_filter = "表格页"
                    elif "图表" in slide or "chart" in slide_lower:
                        layout_filter = "图表页"
                    elif slide.startswith("# ") or "标题" in slide:
                        layout_filter = "封面标题页"

                    # 布局筛选
                    layout_categories = ["全部", "封面标题页", "内容页", "表格页", "图表页", "列表页"]
                    selected_layout = st.selectbox(
                        "筛选布局类型",
                        options=layout_categories,
                        index=layout_categories.index(layout_filter) if layout_filter in layout_categories else 0,
                        key=f"layout_select_{page_idx}"
                    )

                    # 筛选模板
                    filtered_templates = page_templates
                    if selected_layout != "全部":
                        layout_en_map = {"封面标题页": "title", "内容页": "content", "表格页": "table", "图表页": "chart", "列表页": "bullets"}
                        layout_en = layout_en_map.get(selected_layout, "")
                        filtered_templates = [t for t in page_templates if t.get("layout_category") == layout_en]

                    if not filtered_templates:
                        st.warning("没有符合条件的模板")
                    else:
                        st.caption(f"找到 {len(filtered_templates)} 个模板")

                        # 网格显示模板（每行 3 个）
                        cols_per_row = 3
                        for row_start in range(0, len(filtered_templates), cols_per_row):
                            cols = st.columns(cols_per_row)
                            for j, col in enumerate(cols):
                                idx = row_start + j
                                if idx >= len(filtered_templates):
                                    break
                                tpl = filtered_templates[idx]

                                with col:
                                    # 显示缩略图
                                    thumb_path = tpl.get("thumbnail")

                                    if thumb_path:
                                        full_thumb_path = os.path.join(os.path.dirname(__file__), "page_template", thumb_path)
                                        if os.path.exists(full_thumb_path):
                                            st.image(full_thumb_path, caption=f"{tpl['source_name']} - 第{tpl['page_num']}页", width='stretch')
                                        else:
                                            st.write(f"**{tpl['source_name']}**")
                                            st.caption(f"第 {tpl['page_num']} 页 - {tpl['layout_category_cn']}")
                                    else:
                                        st.write(f"**{tpl['source_name']}**")
                                        st.caption(f"第 {tpl['page_num']} 页 - {tpl['layout_category_cn']}")

                                    # 模板描述预览
                                    style_desc = tpl.get("style_description", "")
                                    if len(style_desc) > 100:
                                        style_desc = style_desc[:100] + "..."

                                    with st.expander("查看风格描述"):
                                        st.markdown(tpl.get("style_description", ""))

                                    # 选择按钮
                                    if st.button("选择此模板", key=f"select_tpl_{page_idx}_{tpl['id']}"):
                                        # 保存选择
                                        save_page_style(project_name, page_idx, tpl["style_description"], tpl["id"])
                                        # 使用 pending key 存储新值，下次渲染时应用
                                        st.session_state[f"pending_page_style_{page_idx}"] = tpl["style_description"]
                                        # 更新 session state
                                        st.session_state[f"show_template_selector_{page_idx}"] = False
                                        st.toast(f"已选择模板，风格描述已更新", icon="✅")
                                        st.rerun()

                    # 取消按钮
                    if st.button("取消选择", key=f"cancel_select_{page_idx}"):
                        st.session_state[f"show_template_selector_{page_idx}"] = False
                        st.rerun()

                # 显示当前选中的模板信息
                if selected_template_id:
                    matching_tpl = next((t for t in page_templates if t["id"] == selected_template_id), None)
                    if matching_tpl:
                        st.info(f"当前模板：{matching_tpl['source_name']} 第{matching_tpl['page_num']}页 ({matching_tpl['layout_category_cn']})")


# 删除不再使用的变量
# available_templates 和 template_config 已不再需要

# ── Step 3: Generate images ──────────────────────────────────────────
st.header("Step 3: 生成信息图")

current_opt = read_file(opt_path)
global_style = read_file(style_path)

if current_opt:
    slides = parse_slides(current_opt)
    st.info(f"共解析出 {len(slides)} 页幻灯片")

    def get_page_style(page_idx: int) -> str:
        """获取指定页面的风格描述（每次都重新加载最新数据）"""
        page_key = str(page_idx)
        # 重新加载最新的风格配置
        latest_styles = load_page_styles(project_name)
        saved = latest_styles.get("pages", {}).get(page_key, {})
        return saved.get("style_description", global_style)

    if st.button("一键生成所有图片", type="primary"):
        os.makedirs(img_dir, exist_ok=True)
        progress = st.progress(0)
        status = st.empty()

        # 显示每页使用的风格来源
        style_info = []
        for i in range(len(slides)):
            page_style = get_page_style(i + 1)
            if page_style != global_style:
                style_info.append(f"第{i+1}页: 使用选定的模板风格")
            else:
                style_info.append(f"第{i+1}页: 使用全局风格")
        with st.expander("风格来源", expanded=False):
            for info in style_info:
                st.write(info)

        def generate_single_image(args):
            """生成单张图片的辅助函数"""
            i, slide = args
            try:
                # 使用该页保存的风格描述
                page_style = get_page_style(i + 1)
                img_bytes = generate_slide_image(
                    slide, page_style, i + 1, len(slides), model=image_model
                )
                if img_bytes:
                    img_path = os.path.join(img_dir, f"{i+1:02d}.jpg")
                    with open(img_path, "wb") as f:
                        f.write(img_bytes)
                    return (i, True, None)
                return (i, False, "生成失败")
            except Exception as e:
                return (i, False, str(e))

        if image_execution_mode == "并行":
            # 并行生成图片（多 agent 同时执行）
            with concurrent.futures.ThreadPoolExecutor(max_workers=8) as executor:
                futures = list(executor.map(generate_single_image, enumerate(slides)))
                completed = 0
                for i, success, error in futures:
                    completed += 1
                    status.text(f"已完成 {completed}/{len(slides)} 页")
                    progress.progress(completed / len(slides))
        else:
            # 串行生成图片（逐页执行）
            for i, slide in enumerate(slides):
                result = generate_single_image((i, slide))
                completed = i + 1
                status.text(f"已完成 {completed}/{len(slides)} 页")
                progress.progress(completed / len(slides))

        status.empty()
        st.success("所有图片生成完成！")
        st.rerun()

    # Show existing images and allow per-page regeneration
    existing_imgs = sorted(
        f for f in os.listdir(img_dir)
        if f.lower().endswith((".jpg", ".jpeg", ".png"))
    ) if os.path.exists(img_dir) else []

    if existing_imgs:
        cols_per_row = 3
        for row_start in range(0, len(existing_imgs), cols_per_row):
            cols = st.columns(cols_per_row)
            for j, col in enumerate(cols):
                idx = row_start + j
                if idx >= len(existing_imgs):
                    break
                img_file = existing_imgs[idx]
                img_full = os.path.join(img_dir, img_file)
                with col:
                    st.image(img_full, caption=img_file, width='stretch')
                    page_idx = idx  # index into slides list
                    if page_idx < len(slides):
                        with st.expander(f"重新生成 {img_file}"):
                            custom_prompt = st.text_area(
                                "自定义该页内容（可选）",
                                value=slides[page_idx],
                                key=f"regen_{idx}",
                                height=150,
                            )
                            if st.button("重新生成", key=f"btn_regen_{idx}"):
                                with st.spinner("重新生成中..."):
                                    # 使用该页保存的风格描述
                                    page_style = get_page_style(page_idx + 1)
                                    # 显示使用的风格描述（前 200 字符）
                                    st.caption(f"风格描述: {page_style[:200]}...")
                                    new_bytes = generate_slide_image(
                                        custom_prompt,
                                        page_style,
                                        page_idx + 1,
                                        len(slides),
                                        model=image_model,
                                    )
                                    if new_bytes:
                                        with open(img_full, "wb") as f:
                                            f.write(new_bytes)
                                        st.success("已重新生成")
                                        st.rerun()
else:
    st.info("请先生成优化稿")

# ── Step 4: Export (PDF / PPT) ───────────────────────────────────────
st.header("Step 4: 导出")

has_images = (
    os.path.exists(img_dir)
    and any(f.endswith((".jpg", ".png")) for f in os.listdir(img_dir))
)

tab_pdf, tab_ppt = st.tabs(["合并为 PDF", "生成 PPT"])

# ── Tab 1: PDF ──
with tab_pdf:
    if has_images:
        if st.button("合并为 PDF", type="primary"):
            with st.spinner("正在合并..."):
                build_pdf(img_dir, pdf_path)
            st.success("PDF 生成完成！")
            st.rerun()

        if os.path.exists(pdf_path):
            with open(pdf_path, "rb") as f:
                st.download_button(
                    "下载 PDF",
                    data=f.read(),
                    file_name=f"{project_name}.pdf",
                    mime="application/pdf",
                )
    else:
        st.info("请先生成图片")

# ── Tab 2: PPT ──
with tab_ppt:
    if not has_images:
        st.info("请先生成图片")
    else:
        # Collect image files
        ppt_img_files = sorted(
            f for f in os.listdir(img_dir)
            if f.lower().endswith((".jpg", ".jpeg", ".png"))
        )
        total_pages = len(ppt_img_files)

        # Load existing slide codes
        saved_codes = load_all_slide_codes(proj_dir)

        st.caption(
            f"共 {total_pages} 页信息图。AI 将逐页分析图片并生成 python-pptx 代码，"
            "最终合并为可编辑的 PPTX 文件。"
        )

        # ── One-click full generation ──
        if st.button("一键生成完整 PPT", type="primary", key="btn_full_ppt"):
            progress = st.progress(0)
            status = st.empty()
            all_codes = {}
            failed_pages = []

            def generate_and_test_slide(args):
                """生成单页 PPT 代码并测试"""
                i, img_file = args
                page = i + 1
                img_full = os.path.join(img_dir, img_file)

                for attempt in range(3):  # 最多重试3次
                    try:
                        code = generate_slide_code(
                            image_path=img_full,
                            page_num=page,
                            total_pages=total_pages,
                            model=ppt_model,
                        )

                        # 测试代码是否有效
                        test_ok, test_error = test_slide_code(code)
                        if test_ok:
                            return (page, code, None)
                        else:
                            # 测试失败，如果还有重试机会则继续
                            if attempt < 2:
                                print(f"[Page {page}] Test failed (attempt {attempt+1}), retrying...")
                                continue
                            return (page, None, f"测试失败: {test_error[:500]}")
                    except Exception as e:
                        if attempt < 2:
                            continue
                        return (page, None, str(e))

                return (page, None, "超过最大重试次数")

            # 并行/串行生成 PPT 代码
            if ppt_execution_mode == "并行":
                # 并行生成 PPT 代码（多 agent 同时执行）
                with concurrent.futures.ThreadPoolExecutor(max_workers=8) as executor:
                    futures = list(executor.map(generate_and_test_slide, enumerate(ppt_img_files)))
                    completed = 0
                    for page, code, error in futures:
                        completed += 1
                        if code:
                            status.text(f"第 {page} 页：生成成功 ✅ ({completed}/{total_pages})")
                        else:
                            status.text(f"第 {page} 页：生成失败 ❌ ({completed}/{total_pages})")
                        progress.progress(completed / total_pages)
                        if code:
                            # 再次验证并生成 PPTX
                            single_path = get_single_pptx_path(proj_dir, page)
                            success, err, final_code = build_single_slide_pptx_with_retry(
                                code, single_path, max_retries=1
                            )
                            if success:
                                save_slide_code(proj_dir, page, final_code)
                                all_codes[page] = final_code
                                status.text(f"第 {page} 页：PPTX 生成成功 ✅ ({completed}/{total_pages})")
                            else:
                                failed_pages.append((page, err))
                                status.text(f"第 {page} 页：PPTX 生成失败 ❌ ({completed}/{total_pages})")
                        else:
                            failed_pages.append((page, error))
            else:
                # 串行生成 PPT 代码（逐页执行）
                for i, img_file in enumerate(ppt_img_files):
                    page = i + 1
                    status.text(f"正在生成第 {page} 页... ({page}/{total_pages})")
                    result = generate_and_test_slide((i, img_file))
                    page, code, error = result
                    progress.progress(page / total_pages)
                    if code:
                        single_path = get_single_pptx_path(proj_dir, page)
                        success, err, final_code = build_single_slide_pptx_with_retry(
                            code, single_path, max_retries=1
                        )
                        if success:
                            save_slide_code(proj_dir, page, final_code)
                            all_codes[page] = final_code
                            status.text(f"第 {page} 页：生成成功 ✅ ({page}/{total_pages})")
                        else:
                            failed_pages.append((page, err))
                            status.text(f"第 {page} 页：生成失败 ❌ ({page}/{total_pages})")
                    else:
                        failed_pages.append((page, error))
                        status.text(f"第 {page} 页：生成失败 ❌ ({page}/{total_pages})")

            if failed_pages:
                status.text(f"有 {len(failed_pages)} 页生成失败，正在重试...")
                # 可以选择重试失败的页面
                for page, err in failed_pages:
                    st.warning(f"第 {page} 页生成失败")

            status.text("正在合并为完整 PPTX...")
            success, error = build_full_pptx(all_codes, ppt_path)
            status.empty()
            if success:
                if failed_pages:
                    st.success(f"PPT 生成完成！({len(failed_pages)} 页失败)")
                else:
                    st.success("完整 PPT 生成完成！")
                st.rerun()
            else:
                st.error("PPT 生成失败")
                with st.expander("错误详情"):
                    st.code(error)

        st.divider()

        # ── Per-page grid: checkbox + status + regenerate ──
        st.subheader("逐页管理")

        cols_per_row = 3
        for row_start in range(0, total_pages, cols_per_row):
            cols = st.columns(cols_per_row)
            for j, col in enumerate(cols):
                idx = row_start + j
                if idx >= total_pages:
                    break
                page = idx + 1
                img_file = ppt_img_files[idx]
                img_full = os.path.join(img_dir, img_file)
                has_code = load_slide_code(proj_dir, page) is not None
                single_pptx = get_single_pptx_path(proj_dir, page)
                has_pptx = os.path.exists(single_pptx)

                with col:
                    st.image(img_full, caption=f"第 {page} 页", width='stretch')

                    if has_pptx:
                        st.caption("已生成")
                    else:
                        st.caption("未生成")

                    # AI generate / regenerate
                    if st.button(
                        "AI 重新生成代码" if has_code else "AI 生成代码",
                        key=f"btn_ppt_page_{page}",
                    ):
                        with st.spinner(f"正在生成第 {page} 页（带测试验证）..."):
                            # 生成代码
                            code = generate_slide_code(
                                image_path=img_full,
                                page_num=page,
                                total_pages=total_pages,
                                model=ppt_model,
                            )

                            # 测试并生成PPTX（带重试）
                            ok, err, final_code = build_single_slide_pptx_with_retry(
                                code,
                                single_pptx,
                                max_retries=3,
                                regenerate_func=generate_slide_code,
                                regenerate_args=(img_full, page, total_pages, ppt_model),
                            )
                            if ok:
                                save_slide_code(proj_dir, page, final_code)
                        if ok:
                            st.success(f"第 {page} 页生成完成")
                            st.rerun()
                        else:
                            st.error(f"第 {page} 页生成失败（已重试3次）")
                            with st.expander("错误详情"):
                                st.code(err)

                    # Editable code + run from code
                    if has_code:
                        with st.expander("编辑代码 / 生成 PPT 页"):
                            edited = st.text_area(
                                "代码",
                                value=load_slide_code(proj_dir, page),
                                height=300,
                                key=f"code_edit_{page}",
                            )
                            c1, c2 = st.columns(2)
                            with c1:
                                if st.button("保存代码", key=f"btn_save_code_{page}"):
                                    save_slide_code(proj_dir, page, edited)
                                    st.success("已保存")
                            with c2:
                                if st.button("生成 PPT 页", key=f"btn_run_code_{page}", type="primary"):
                                    # 先测试代码
                                    test_ok, test_err = test_slide_code(edited)
                                    if not test_ok:
                                        st.error("代码测试失败，请检查语法错误")
                                        with st.expander("测试错误详情"):
                                            st.code(test_err)
                                    else:
                                        save_slide_code(proj_dir, page, edited)
                                        with st.spinner("正在生成..."):
                                            ok, err = build_single_slide_pptx(
                                                edited, single_pptx
                                            )
                                        if ok:
                                            st.success(f"第 {page} 页 PPT 生成完成")
                                            st.rerun()
                                        else:
                                            st.error("生成失败")
                                            st.code(err)

        st.divider()

        # ── Merge existing single-page codes into full PPTX ──
        saved_codes = load_all_slide_codes(proj_dir)
        if saved_codes:
            st.subheader("合并为完整 PPT")
            st.info(f"已有 {len(saved_codes)}/{total_pages} 页代码")

            if st.button("合并已有页面为完整 PPT", key="btn_merge_ppt"):
                with st.spinner("正在合并..."):
                    success, error = build_full_pptx(saved_codes, ppt_path)
                if success:
                    st.success("合并完成！")
                    st.rerun()
                else:
                    st.error("合并失败")
                    with st.expander("错误详情"):
                        st.code(error)

        # ── Download ──
        if os.path.exists(ppt_path):
            with open(ppt_path, "rb") as f:
                st.download_button(
                    "下载完整 PPT",
                    data=f.read(),
                    file_name=f"{project_name}.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    key="btn_download_full_ppt",
                )
