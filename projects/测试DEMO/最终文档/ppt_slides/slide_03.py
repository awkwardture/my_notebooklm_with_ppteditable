def build_slide(slide):
    # 背景设置 (深色主题)
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), SLIDE_WIDTH, Inches(7.5))
    bg.fill.solid()
    bg.fill.fore_color.rgb = RGBColor(0x14, 0x16, 0x1A)
    bg.line.fill.background()

    # 标题
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.4), Inches(10), Inches(0.8))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    
    run1 = p.add_run()
    run1.text = "高复杂度：效率反降 "
    run1.font.size = Pt(32)
    run1.font.bold = True
    run1.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    run1.font.name = FONT_NAME

    run2 = p.add_run()
    run2.text = "60%"
    run2.font.size = Pt(32)
    run2.font.bold = True
    run2.font.color.rgb = RGBColor(0xE5, 0x39, 0x35) # 红色
    run2.font.name = FONT_NAME

    run3 = p.add_run()
    run3.text = " 的修复陷阱"
    run3.font.size = Pt(32)
    run3.font.bold = True
    run3.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    run3.font.name = FONT_NAME

    # 副标题
    sub_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.1), Inches(10), Inches(0.5))
    tf_sub = sub_box.text_frame
    p_sub = tf_sub.paragraphs[0]
    p_sub.text = "案例C — 跨工厂 STO 报表（18+ 关联表）"
    p_sub.font.size = Pt(18)
    p_sub.font.color.rgb = RGBColor(0xAA, 0xAA, 0xAA)
    p_sub.font.name = FONT_NAME

    # 左上区域：-60% 效率损失 (使用文本和箭头代替复杂插图)
    loss_box = slide.shapes.add_textbox(Inches(4.5), Inches(2.2), Inches(2.5), Inches(1.5))
    tf_loss = loss_box.text_frame
    p_loss1 = tf_loss.paragraphs[0]
    p_loss1.text = "-60%"
    p_loss1.font.size = Pt(48)
    p_loss1.font.bold = True
    p_loss1.font.color.rgb = RGBColor(0xE5, 0x39, 0x35)
    p_loss1.alignment = PP_ALIGN.CENTER

    p_loss2 = tf_loss.add_paragraph()
    p_loss2.text = "效率损失"
    p_loss2.font.size = Pt(24)
    p_loss2.font.bold = True
    p_loss2.font.color.rgb = RGBColor(0xE5, 0x39, 0x35)
    p_loss2.alignment = PP_ALIGN.CENTER

    # 向下箭头
    arrow = slide.shapes.add_shape(MSO_SHAPE.DOWN_ARROW, Inches(5.8), Inches(3.4), Inches(0.8), Inches(1.0))
    arrow.fill.solid()
    arrow.fill.fore_color.rgb = RGBColor(0xE5, 0x39, 0x35)
    arrow.line.fill.background()
    arrow.rotation = -45

    # 右上区域：条形图对比
    box_tr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.5), Inches(1.8), Inches(5.3), Inches(2.2))
    box_tr.fill.solid()
    box_tr.fill.fore_color.rgb = RGBColor(0x2A, 0x2D, 0x35)
    box_tr.line.color.rgb = RGBColor(0x55, 0x55, 0x55)

    # AI 修复标签
    lbl_ai = slide.shapes.add_textbox(Inches(7.7), Inches(2.0), Inches(2.0), Inches(0.4))
    lbl_ai.text_frame.text = "AI 修复"
    lbl_ai.text_frame.paragraphs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    lbl_ai.text_frame.paragraphs[0].font.size = Pt(14)

    # AI 修复条 (红色)
    bar_ai = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.8), Inches(2.4), Inches(4.7), Inches(0.4))
    bar_ai.fill.solid()
    bar_ai.fill.fore_color.rgb = RGBColor(0x9E, 0x1B, 0x22)
    bar_ai.line.fill.background()
    
    txt_ai = slide.shapes.add_textbox(Inches(9.5), Inches(2.4), Inches(2.8), Inches(0.4))
    txt_ai.text_frame.text = "8 人天 (远超手写)"
    txt_ai.text_frame.paragraphs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    txt_ai.text_frame.paragraphs[0].font.size = Pt(12)
    txt_ai.text_frame.paragraphs[0].font.bold = True
    txt_ai.text_frame.paragraphs[0].alignment = PP_ALIGN.RIGHT

    # 手写开发标签
    lbl_man = slide.shapes.add_textbox(Inches(7.7), Inches(2.9), Inches(2.0), Inches(0.4))
    lbl_man.text_frame.text = "手写开发"
    lbl_man.text_frame.paragraphs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    lbl_man.text_frame.paragraphs[0].font.size = Pt(14)

    # 手写开发条 (蓝灰色)
    bar_man = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.8), Inches(3.3), Inches(3.0), Inches(0.4))
    bar_man.fill.solid()
    bar_man.fill.fore_color.rgb = RGBColor(0x4A, 0x62, 0x78)
    bar_man.line.fill.background()
    
    txt_man = slide.shapes.add_textbox(Inches(9.0), Inches(3.3), Inches(1.6), Inches(0.4))
    txt_man.text_frame.text = "5 人天"
    txt_man.text_frame.paragraphs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    txt_man.text_frame.paragraphs[0].font.size = Pt(12)
    txt_man.text_frame.paragraphs[0].font.bold = True
    txt_man.text_frame.paragraphs[0].alignment = PP_ALIGN.RIGHT

    # 左下区域：折线图与警告
    box_bl = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(4.3), Inches(6.5), Inches(2.8))
    box_bl.fill.solid()
    box_bl.fill.fore_color.rgb = RGBColor(0x1A, 0x1C, 0x22)
    box_bl.line.color.rgb = RGBColor(0x55, 0x55, 0x55)

    # 坐标轴
    slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(1.0), Inches(4.6), Inches(1.0), Inches(6.6)).line.color.rgb = RGBColor(0x88, 0x88, 0x88)
    slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(1.0), Inches(6.6), Inches(4.0), Inches(6.6)).line.color.rgb = RGBColor(0x88, 0x88, 0x88)

    # 坐标轴标签
    y_label = slide.shapes.add_textbox(Inches(0.5), Inches(4.8), Inches(0.4), Inches(1.5))
    y_label.text_frame.word_wrap = True
    y_label.text_frame.text = "错\n误\n数\n量"
    y_label.text_frame.paragraphs[0].font.size = Pt(10)
    y_label.text_frame.paragraphs[0].font.color.rgb = RGBColor(0xAA, 0xAA, 0xAA)
    y_label.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

    x_label = slide.shapes.add_textbox(Inches(3.2), Inches(6.7), Inches(1.0), Inches(0.4))
    x_label.text_frame.text = "修复迭代"
    x_label.text_frame.paragraphs[0].font.size = Pt(10)
    x_label.text_frame.paragraphs[0].font.color.rgb = RGBColor(0xAA, 0xAA, 0xAA)

    # 数据点与连线 (3 -> 2 -> 9)
    pt1_x, pt1_y = Inches(1.4), Inches(6.1)
    pt2_x, pt2_y = Inches(2.3), Inches(6.3)
    pt3_x, pt3_y = Inches(3.5), Inches(5.0)

    line1 = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, pt1_x, pt1_y, pt2_x, pt2_y)
    line1.line.color.rgb = RGBColor(0xE5, 0x39, 0x35)
    line1.line.width = Pt(2)
    
    line2 = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, pt2_x, pt2_y, pt3_x, pt3_y)
    line2.line.color.rgb = RGBColor(0xE5, 0x39, 0x35)
    line2.line.width = Pt(2)

    for px, py in [(pt1_x, pt1_y), (pt2_x, pt2_y), (pt3_x, pt3_y)]:
        dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, px - Inches(0.05), py - Inches(0.05), Inches(0.1), Inches(0.1))
        dot.fill.solid()
        dot.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        dot.line.color.rgb = RGBColor(0xE5, 0x39, 0x35)
        dot.line.width = Pt(1.5)

    # 数据点标签
    lbl_pt1 = slide.shapes.add_textbox(pt1_x - Inches(0.2), pt1_y - Inches(0.3), Inches(0.4), Inches(0.3))
    lbl_pt1.text_frame.text = "3"
    lbl_pt1.text_frame.paragraphs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    lbl_pt2 = slide.shapes.add_textbox(pt2_x - Inches(0.2), pt2_y - Inches(0.3), Inches(0.4), Inches(0.3))
    lbl_pt2.text_frame.text = "2"
    lbl_pt2.text_frame.paragraphs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    lbl_pt3 = slide.shapes.add_textbox(pt3_x - Inches(0.1), pt3_y - Inches(0.3), Inches(0.4), Inches(0.3))
    lbl_pt3.text_frame.text = "9"
    lbl_pt3.text_frame.paragraphs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    # 趋势标注
    anno = slide.shapes.add_textbox(Inches(2.4), Inches(5.4), Inches(1.0), Inches(0.3))
    anno.text_frame.text = "3→2→9"
    anno.text_frame.paragraphs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    anno.text_frame.paragraphs[0].font.size = Pt(10)
    anno.rotation = -35

    # 警告区域 (修复爆炸)
    warn_icon = slide.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, Inches(4.2), Inches(4.8), Inches(0.3), Inches(0.3))
    warn_icon.fill.solid()
    warn_icon.fill.fore_color.rgb = RGBColor(0xE5, 0x39, 0x35)
    warn_icon.line.fill.background()
    
    warn_ex = slide.shapes.add_textbox(Inches(4.15), Inches(4.8), Inches(0.4), Inches(0.3))
    warn_ex.text_frame.text = "!"
    warn_ex.text_frame.paragraphs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    warn_ex.text_frame.paragraphs[0].font.bold = True
    warn_ex.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

    warn_title = slide.shapes.add_textbox(Inches(4.6), Inches(4.75), Inches(2.0), Inches(0.4))
    warn_title.text_frame.text = "修复爆炸"
    warn_title.text_frame.paragraphs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    warn_title.text_frame.paragraphs[0].font.size = Pt(16)
    warn_title.text_frame.paragraphs[0].font.bold = True

    warn_desc = slide.shapes.add_textbox(Inches(4.2), Inches(5.3), Inches(2.6), Inches(1.5))
    tf_wd = warn_desc.text_frame
    tf_wd.word_wrap = True
    p_wd = tf_wd.paragraphs[0]
    p_wd.font.size = Pt(12)
    p_wd.font.color.rgb = RGBColor(0xCC, 0xCC, 0xCC)
    
    run_wd1 = p_wd.add_run()
    run_wd1.text = "错误呈 "
    run_wd2 = p_wd.add_run()
    run_wd2.text = "3→2→9"
    run_wd2.font.color.rgb = RGBColor(0xE5, 0x39, 0x35)
    run_wd3 = p_wd.add_run()
    run_wd3.text = " 非线性增长，前序修正引发后续逻辑大规模崩溃"

    # 右下区域：要点列表
    # 要点 1: 解析能力
    icon1 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.5), Inches(4.5), Inches(0.8), Inches(0.8))
    icon1.fill.solid()
    icon1.fill.fore_color.rgb = RGBColor(0x33, 0x33, 0x33)
    icon1.line.fill.background()
    
    title1 = slide.shapes.add_textbox(Inches(8.5), Inches(4.4), Inches(4.0), Inches(0.4))
    title1.text_frame.text = "解析能力"
    title1.text_frame.paragraphs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    title1.text_frame.paragraphs[0].font.size = Pt(16)
    title1.text_frame.paragraphs[0].font.bold = True

    desc1 = slide.shapes.add_textbox(Inches(8.5), Inches(4.8), Inches(4.5), Inches(0.8))
    desc1.text_frame.word_wrap = True
    p_d1 = desc1.text_frame.paragraphs[0]
    p_d1.font.size = Pt(12)
    p_d1.font.color.rgb = RGBColor(0xCC, 0xCC, 0xCC)
    p_d1.add_run().text = "Claude 虽支持文档解析，但在处理 "
    r_d1 = p_d1.add_run()
    r_d1.text = "10+"
    r_d1.font.color.rgb = RGBColor(0xE5, 0x39, 0x35)
    p_d1.add_run().text = " 项虚构数据字典时表现乏力"

    # 分隔线
    sep = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(7.5), Inches(5.7), Inches(12.8), Inches(5.7))
    sep.line.color.rgb = RGBColor(0x44, 0x44, 0x44)
    sep.line.dash_style = 2

    # 要点 2: 实测评价
    icon2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.5), Inches(6.0), Inches(0.8), Inches(0.8))
    icon2.fill.solid()
    icon2.fill.fore_color.rgb = RGBColor(0x33, 0x33, 0x33)
    icon2.line.fill.background()

    title2 = slide.shapes.add_textbox(Inches(8.5), Inches(5.9), Inches(4.0), Inches(0.4))
    title2.text_frame.text = "实测评价"
    title2.text_frame.paragraphs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    title2.text_frame.paragraphs[0].font.size = Pt(16)
    title2.text_frame.paragraphs[0].font.bold = True

    desc2 = slide.shapes.add_textbox(Inches(8.5), Inches(6.3), Inches(4.5), Inches(0.8))
    desc2.text_frame.word_wrap = True
    p_d2 = desc2.text_frame.paragraphs[0]
    p_d2.font.size = Pt(12)
    p_d2.font.color.rgb = RGBColor(0xCC, 0xCC, 0xCC)
    p_d2.text = "底层数据结构理解缺失，导致“修复成本 > 直接重写”"

    # 页脚
    footer = slide.shapes.add_textbox(Inches(12.0), Inches(7.0), Inches(1.0), Inches(0.4))
    footer.text_frame.text = "第 3 页"
    footer.text_frame.paragraphs[0].font.size = Pt(10)
    footer.text_frame.paragraphs[0].font.color.rgb = RGBColor(0x66, 0x66, 0x66)
    footer.text_frame.paragraphs[0].alignment = PP_ALIGN.RIGHT