def build_slide(slide):
    # Colors
    BLUE_DARK = RGBColor(0x1F, 0x4E, 0x79)
    BLUE_PRIMARY = RGBColor(0x00, 0x70, 0xC0)
    BLUE_LIGHT_BG = RGBColor(0xF4, 0xF8, 0xFC)
    BLUE_BORDER = RGBColor(0xBD, 0xD7, 0xEE)
    GRAY_TEXT = RGBColor(0x55, 0x55, 0x55)
    GRAY_BAR = RGBColor(0xA6, 0xA6, 0xA6)
    ORANGE_WARN = RGBColor(0xED, 0x7D, 0x31)
    BLACK_TEXT = RGBColor(0x33, 0x33, 0x33)

    # 1. Title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.4), Inches(9), Inches(0.8))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = "简单场景：AI 初显 50% 提效潜力 ⚡"
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.name = FONT_NAME

    # 2. Subtitle
    sub_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.1), Inches(8), Inches(0.5))
    tf = sub_box.text_frame
    p = tf.paragraphs[0]
    p.text = "案例A —— 标准 ALV 销售订单报表验证"
    p.font.size = Pt(16)
    p.font.color.rgb = GRAY_TEXT
    p.font.name = FONT_NAME

    # 3. Page Indicator (Top Right)
    page_shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(11.8), Inches(0.6), Inches(1.0), Inches(0.4))
    page_shape.fill.solid()
    page_shape.fill.fore_color.rgb = RGBColor(0xE6, 0xF0, 0xFA)
    page_shape.line.fill.background()
    tf = page_shape.text_frame
    tf.text = "P1 / 4"
    tf.paragraphs[0].font.size = Pt(12)
    tf.paragraphs[0].font.color.rgb = BLUE_PRIMARY
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER

    # 4. Left Column - Core Points Heading
    heading_box = slide.shapes.add_textbox(Inches(0.5), Inches(2.0), Inches(3), Inches(0.5))
    tf = heading_box.text_frame
    p = tf.paragraphs[0]
    p.text = "核心要点"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.name = FONT_NAME

    # 5. Left Column - Core Points Items
    items = [
        ("⚡", "1. 效率突破: ", "Claude Code 仅需 30 分钟完成开发，较手写提效 50%。"),
        ("🤖", "2. 工具差距: ", "GitHub Copilot 耗时与手写持平，且因 SQL 不兼容导致运行崩溃。"),
        ("🔀", "3. 交互对比: ", "Claude 仅需 2 轮提示即可运行，Copilot 需 5 轮以上人工干预。"),
        ("🔍", "4. 核心瓶颈: ", "[地址取数逻辑] 与 [过账状态字段] 是 AI 初始生成的共同盲区。")
    ]

    y_offset = 2.7
    for i, (icon, label, desc) in enumerate(items):
        # Icon
        icon_box = slide.shapes.add_textbox(Inches(0.5), y_offset, Inches(0.6), Inches(0.6))
        icon_box.text_frame.text = icon
        icon_box.text_frame.paragraphs[0].font.size = Pt(24)

        # Text
        text_box = slide.shapes.add_textbox(Inches(1.2), y_offset, Inches(4.8), Inches(0.8))
        text_box.text_frame.word_wrap = True
        p = text_box.text_frame.paragraphs[0]
        
        run1 = p.add_run()
        run1.text = label
        run1.font.bold = True
        run1.font.size = Pt(14)
        run1.font.color.rgb = BLUE_DARK
        run1.font.name = FONT_NAME
        
        run2 = p.add_run()
        run2.text = desc
        run2.font.size = Pt(14)
        run2.font.color.rgb = BLACK_TEXT
        run2.font.name = FONT_NAME

        # Separator line
        if i < len(items) - 1:
            line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(1.2), y_offset + 0.85, Inches(6.0), y_offset + 0.85)
            line.line.color.rgb = RGBColor(0xE0, 0xE0, 0xE0)
        
        y_offset += 1.0

    # 6. Right Column - Chart Box Background
    chart_bg = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.5), Inches(1.8), Inches(6.3), Inches(4.2))
    chart_bg.fill.solid()
    chart_bg.fill.fore_color.rgb = BLUE_LIGHT_BG
    chart_bg.line.color.rgb = BLUE_BORDER
    chart_bg.line.width = Pt(1.5)

    # Chart Title
    chart_title = slide.shapes.add_textbox(Inches(6.7), Inches(2.0), Inches(4), Inches(0.5))
    tf = chart_title.text_frame
    p = tf.paragraphs[0]
    p.text = "开发耗时对比 (分钟)"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.name = FONT_NAME

    # Chart
    chart_data = CategoryChartData()
    chart_data.categories = ["手写/Copilot\n(Manual/AI)", "Claude Code\n(AI)"]
    chart_data.add_series('耗时', [60, 30])

    x, y, cx, cy = Inches(6.6), Inches(2.8), Inches(5.8), Inches(2.5)
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED, x, y, cx, cy, chart_data
    ).chart

    chart.has_legend = False
    chart.plots[0].has_data_labels = False # We will add custom text boxes for labels
    chart.plots[0].gap_width = 100

    # Customize bar colors
    series = chart.plots[0].series[0]
    pt0 = series.points[0]
    pt0.format.fill.solid()
    pt0.format.fill.fore_color.rgb = GRAY_BAR
    pt1 = series.points[1]
    pt1.format.fill.solid()
    pt1.format.fill.fore_color.rgb = BLUE_PRIMARY

    # Custom Data Labels
    lbl1 = slide.shapes.add_textbox(Inches(11.9), Inches(4.15), Inches(1), Inches(0.4))
    lbl1.text_frame.text = "⚠️ 60"
    lbl1.text_frame.paragraphs[0].font.size = Pt(14)
    lbl1.text_frame.paragraphs[0].font.name = FONT_NAME

    lbl2 = slide.shapes.add_textbox(Inches(9.6), Inches(3.25), Inches(1), Inches(0.4))
    lbl2.text_frame.text = "⚡ 30"
    lbl2.text_frame.paragraphs[0].font.size = Pt(14)
    lbl2.text_frame.paragraphs[0].font.name = FONT_NAME

    # Chart Callout (50% 提效)
    callout = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGULAR_CALLOUT, Inches(9.1), Inches(2.4), Inches(1.6), Inches(0.6))
    callout.fill.solid()
    callout.fill.fore_color.rgb = RGBColor(0xE6, 0xF0, 0xFA)
    callout.line.color.rgb = BLUE_BORDER
    tf = callout.text_frame
    p = tf.paragraphs[0]
    p.text = "50% 提效"
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = BLUE_PRIMARY
    p.font.name = FONT_NAME
    p.alignment = PP_ALIGN.CENTER

    # Chart Warning Text
    warn_text = slide.shapes.add_textbox(Inches(10.3), Inches(4.6), Inches(2.5), Inches(0.4))
    tf = warn_text.text_frame
    p = tf.paragraphs[0]
    p.text = "SQL 不兼容/运行崩溃"
    p.font.size = Pt(10)
    p.font.color.rgb = ORANGE_WARN
    p.font.name = FONT_NAME

    # 7. Right Column - Conclusion Box
    conc_bg = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.5), Inches(6.2), Inches(6.3), Inches(0.8))
    conc_bg.fill.solid()
    conc_bg.fill.fore_color.rgb = BLUE_LIGHT_BG
    conc_bg.line.color.rgb = BLUE_BORDER
    conc_bg.line.width = Pt(1.5)

    conc_text = slide.shapes.add_textbox(Inches(6.5), Inches(6.35), Inches(6.3), Inches(0.5))
    tf = conc_text.text_frame
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    
    run1 = p.add_run()
    run1.text = "✔️ ⚡ AI 初显 "
    run1.font.size = Pt(22)
    run1.font.bold = True
    run1.font.color.rgb = BLACK_TEXT
    run1.font.name = FONT_NAME

    run2 = p.add_run()
    run2.text = "50% 提效潜力"
    run2.font.size = Pt(22)
    run2.font.bold = True
    run2.font.color.rgb = BLUE_PRIMARY
    run2.font.name = FONT_NAME