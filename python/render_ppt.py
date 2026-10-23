import json
import os
import re

# ===== pptx =====
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR
from pptx.enum.dml import MSO_LINE
from pptx.oxml.ns import qn

# ===== pillow =====
from PIL import Image # 需要pip install pillow
from PIL import ImageDraw, ImageFilter # 引入图像滤镜库

# ===== lxml ===== 
from lxml import etree

# 配置
INPUT_JSON = "optimized.json"
TEMPLATE = "template.pptx"
OUTPUT = "final_presentation.pptx"

# 定义宏变量
current_part = "" # 当前部分的名称

# ===== 颜色设置部分 =====
## 渐变调色板
CARD_COLORS = [
    (0x1F, 0x4E, 0x79),   # 深蓝
    (0x4A, 0x90, 0xD9),   # 天蓝
    (0x8E, 0x24, 0x8E),   # 紫色
    (0xE6, 0x7E, 0x22),   # 橙色
    (0x27, 0xAE, 0x60),   # 绿色
    (0xC0, 0x39, 0x2B),   # 红色
]

## 映射 AI 输出的颜色字符串到 RGB
HIGHLIGHT_COLORS = {
    "red":    (0xE7, 0x4C, 0x3C),
    "blue":   (0x34, 0x98, 0xDB),
    "green":  (0x2E, 0xCC, 0x71),
    "orange": (0xE6, 0x7E, 0x22),
    "purple": (0x9B, 0x59, 0xB6),
    "gray":   (0x95, 0xA5, 0xA6),
    "black":  (0x33, 0x33, 0x33) # 默认
}

# ===========================================
# ===== 通用工具函数（画卡片、画文字等）=====
# ===========================================

# ===== 文字函数 ===== 

## 1.1 字体设置函数
# 全局默认字体
DEFAULT_FONT = "宋体"

def set_run_font(run, font_size, bold=False, color=(0, 0, 0), font_name=DEFAULT_FONT):
    """通用函数：设置 Run 的字体、大小、颜色，并完美支持中文字体"""
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = RGBColor(*color)
    
    # 核心：设置西文字体和中文字体
    run.font.name = font_name
    # 操作底层 XML 强制设置东亚字体（ea）
    r = run._r
    rPr = r.get_or_add_rPr()
    ea = rPr.find(qn('a:ea'))
    if ea is None:
        ea = etree.SubElement(rPr, qn('a:ea'))
    ea.set('typeface', font_name)

## 1.2 在指定位置绘制文本框
def draw_textbox(slide, left, top, width, height, text,
                 font_size=18, bold=False, color=(255, 255, 255), 
                 align=PP_ALIGN.LEFT,
                 vertical=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.text = text
    tf.vertical_anchor = vertical
    for para in tf.paragraphs:
        para.alignment = align
        for run in para.runs:
            set_run_font(run, font_size, bold, color)
    return tb

## 1.3 正则匹配高亮函数
def draw_highlighted_text(slide, left, top, width, height, text, highlights, 
                          layout_type,
                          default_font_size=20, default_color=(0x33, 0x33, 0x33),
                          align=PP_ALIGN.LEFT, vertical=MSO_ANCHOR.TOP):
    """
    绘制带有高亮片段的文本
    :param text: 完整的句子，如"技术突破: 加大研发投入"
    :param highlights: AI返回的列表，如[{"text": "技术突破", "color": "blue", "bold": True}]
    """
    tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = tb.text_frame
    tf.word_wrap = True # 自动换行
    tf.vertical_anchor = vertical

    p = tf.paragraphs[0]
    p.alignment = align

    # 如果需要添加项目符号，先单独塞一个Run进去
    if layout_type == "title_bullets":
        bullet_run = p.add_run()
        bullet_run.text = "• "
        bullet_run.font.size = Pt(default_font_size)
        bullet_run.font.bold = True
        bullet_run.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)

    # 如果没有任何高亮的需求，直接普通输出
    if not highlights: 
        run = p.add_run() # pptx 的三个层次: TextFrame → Paragraph → Run(文本块)
        run.text = text
        run.font.size = Pt(default_font_size)
        run.font.color.rgb = RGBColor(*default_color)
        return tb

    # 建立字典，方便O(1)查找
    kw_dict = {item["text"]: item for item in highlights if "text" in item}

    # 生成正则表达式，例如: "技术突破|研发投入"
    # 使用re.escape防止关键词里带有*（）等特殊符号导致报错
    pattern = "|".join(re.escape(kw) for kw in kw_dict.keys())

    # 使用re.split进行切割，重点：加上（）就是捕获组，切出来的数组会保留关键词本身
    parts = re.split(f"({pattern})", text)

    # 遍历切片，逐个生成Run
    for part in parts:

        if not part:
            continue

        run = p.add_run()
        run.text = part

        if part in kw_dict:
            # 如果命中关键词，应用高亮样式
            info = kw_dict[part]
            color_name = info.get("color", "black").lower()
            is_bold = info.get("bold", True)

            run.font.bold = is_bold
            # 找不到颜色就默认使用黑色兜底
            rgb_val = HIGHLIGHT_COLORS.get(color_name, (0, 0, 0))
            set_run_font(run, default_font_size, bold=is_bold, color=rgb_val)
        else:
            # 没命中就是普通文字
            set_run_font(run, default_font_size, bold=False, color=default_color)
        
    # 循环结束后返回对象    
    return tb

# ===== 绘制图形卡片 ===== 

## 2.1 画一张圆角矩形卡片
def draw_card(slide, left, top, width, height, bg_color=(0xF2, 0xF6, 0xFA), 
              border_color=(0x1F, 0x4E, 0x79)):
    card = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    card.fill.solid()
    card.fill.fore_color.rgb = RGBColor(*bg_color)
    card.line.color.rgb = RGBColor(*border_color)
    card.line.width = Pt(1.5)
    return card

## 2.2 绘制一张矩形卡片，无边沿线
def draw_rectangle(slide, left, top, width, height, bg_color=(0xF2, 0xF6, 0xFA),
                   border_color=(0x1F, 0x4E, 0x79)):
    card = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    card.fill.solid()
    card.fill.fore_color.rgb = RGBColor(*bg_color)
    card.line.fill.background() # 将线条填充设为“无”
    return card

## 2.3 绘制向右渐变蒙版
def draw_rectangle_mask(slide, left, top, width, height, bg_color=(0xF2, 0xF6, 0xFA),
                        border_color=(0x1F, 0x4E, 0x79)):
    card = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(left), Inches(top), Inches(width), Inches(height),
    )

    # 设置填充线条为无
    card.line.fill.background()
    
    # 启用渐变填充（创建两个渐变光圈）
    fill = card.fill
    fill.gradient()

    # 设置为左右渐变
    fill.gradient_angle = 0

    # 获取渐变光圈集合
    stops = fill.gradient_stops

    # ===== 设置第一个渐变光圈（左侧）=====
    stop1 = stops[0]
    stop1.position = 0.0 # 0.0代表最左侧
    stop1.color.rgb = RGBColor(*bg_color)
    
    # 修改光圈透明度（需要操作XML）
    color_element_1 = stop1.color._xFill.find(qn('a:srgbClr'))
    # 查看是否已有alpha节点（alpha是控制透明度的参数）
    alpha1 = color_element_1.find(qn('a:alpha'))
    if alpha1 is None:
        from lxml import etree
        alpha1 = etree.SubElement(color_element_1, qn('a:alpha'))
    alpha1.set('val', '0') # 0为全透明 

    # ===== 设置第二个渐变光圈（右侧）=====
    stop2 = stops[1]
    stop2.position = 0.9 # 90% 的位置
    stop2.color.rgb = RGBColor(*bg_color)

    color_element_2 = stop2.color._xFill.find(qn('a:srgbClr'))
    alpha2 = color_element_2.find(qn('a:alpha'))
    if alpha2 is None:
        from lxml import etree
        alpha2 = etree.SubElement(color_element_2, qn('a:alpha'))
    alpha2.set('val', '100000') # 默认值为100000，也就是完全不透明 

    return card

## 2.4 绘制三角形
### 绘制等腰三角形
def draw_isosceles_triangle(slide, left, top, width, height, bg_color=(0xF2, 0xF6, 0xFA),
                  border_color=(0x1F, 0x4E, 0x79)):
    triangle = slide.shapes.add_shape(
        MSO_SHAPE.ISOSCELES_TRIANGLE,
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    triangle.fill.solid()
    triangle.fill.fore_color.rgb = RGBColor(*bg_color)
    triangle.line.fill.background()
    return triangle

### 绘制直角三角形（默认直角在左下角）
def draw_right_triangle(slide, left, top, width, height, bg_color=(0xF2, 0xF6, 0xFA),
                  border_color=(0x1F, 0x4E, 0x79)):
    triangle = slide.shapes.add_shape(
        MSO_SHAPE.RIGHT_TRIANGLE,
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    triangle.fill.solid()
    triangle.fill.fore_color.rgb = RGBColor(*bg_color)
    triangle.line.fill.background()
    return triangle

## 2.5 将两个图形组合成一个
def draw_group(shape1, shape2):
    group = slide.shapes.add_group_shape()

    group.shapes._spTree.append(shape1._element)
    group.shapes._spTree.append(shape2._element)

    return group

## 2.6 绘制便签
def draw_folded_note(slide, left, top, width, height, fold_size=0.4):
    """
    绘制一个带折角的便签
    :param fold_size: 折角的大小（英寸），默认 0.4
    """
    from pptx.enum.shapes import MSO_SHAPE
    
    # 1. 绘制背景矩形（白色，带全边框）
    rect = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    rect.fill.solid()
    rect.fill.fore_color.rgb = RGBColor(255, 255, 255)
    rect.line.color.rgb = RGBColor(200, 200, 200) # 浅灰色边框
    rect.line.width = Pt(1)

    # 2. 绘制右下角的折角三角形
    tri_left = left + width - fold_size
    tri_top = top + height - fold_size
    
    # 默认 RIGHT_TRIANGLE 直角在左下，旋转 90 度让直角跑到右下角
    tri = slide.shapes.add_shape(
        MSO_SHAPE.RIGHT_TRIANGLE, 
        Inches(tri_left), Inches(tri_top), Inches(fold_size), Inches(fold_size)
    )
    tri.rotation = 90 
    
    # 3. 设置折角的外观
    tri.fill.solid()
    # 折角颜色可以比底色稍微深一点点，比如 250, 250, 250
    tri.fill.fore_color.rgb = RGBColor(245, 245, 245) 
    # 去掉三角形的边框（或者给右边和下边加边框）
    tri.line.color.rgb = RGBColor(200, 200, 200)
    tri.line.width = Pt(1)

    # 4. 给折角加向左上方的阴影，体现立体感
    add_shadow_to_shape(tri, direction=2700000) # 270度方向（左上）

# ===== 功能函数 =====
    
## 3.1 将图片加入幻灯片
def add_image_to_slide(slide, image_path, left, top, max_width, max_height):
    """按最大尺寸等比例缩放插入图片，避免形变"""
    if not os.path.exists(image_path):
        return 
    with Image.open(image_path) as img: 
        w, h = img.size
    ratio = min(max_width / w, max_height / h)
    new_w = w * ratio
    new_h = h * ratio
    slide.shapes.add_picture(image_path, Inches(left), Inches(top),
                             width=Inches(new_w), height=Inches(new_h))

## 3.2 边缘羽化函数
def create_faded_image(image_path, output_path, fade_width_ratio=0.15):
    """生成边缘渐变透明的图片"""
    try:
        # 1.打开图片并转换为 RGBA 模式（支持透明通道）
        img = Image.open(image_path).convert("RGBA")
        w, h = img.size

        # 2.创建一个全黑的蒙版（L模式）
        mask = Image.new("L", (w, h), 0)
        draw = ImageDraw.Draw(mask)

        # 3.计算边缘淡化的像素宽度
        fade_w = int(w * fade_width_ratio)
        fade_h = int(h * fade_width_ratio)

        # 4.在中心画一个实心白色矩形，四周留出fade宽度
        draw.rectangle(
            [fade_w, fade_h, w - fade_w, h - fade_h],
            fill=255
        )

        # 5.对蒙版进行高斯模糊，产生渐变过渡
        mask = mask.filter(ImageFilter.GaussianBlur(fade_w // 2))

        # 6.将处理好的蒙版应用到原图的Alpha通道
        img.putalpha(mask)

        # 7.保存为PNG（必须带透明通道）
        img.save(output_path, "PNG")
        return True
    except Exception as e:
        print(f"图片羽化失败: {e}")
        return False

# 3.3 阴影绘制函数
def add_shadow_to_picture(picture, blur_radius=320000, distance=80000, direction=5400000, alpha=80000):
    """
    给图片添加原生阴影
    :param picture: add_picture 返回的图片对象
    :param blur_radius: 模糊半径 (EMU单位，40000约等于3pt)
    :param distance: 阴影偏移距离 (EMU单位，20000约等于1.5pt)
    :param direction: 阴影方向 (60000分之一度，5400000表示正下方90度)
    :param alpha: 阴影透明度 (0-100000，40000表示40%不透明度)
    """
    # 获取图片的 XML 节点
    spPr = picture._element.spPr
    
    # 构造阴影的 XML 结构
    effectLst = etree.SubElement(spPr, qn('a:effectLst'))
    outerShdw = etree.SubElement(effectLst, qn('a:outerShdw'))
    
    # 设置阴影属性
    outerShdw.set('blurRad', str(blur_radius))
    outerShdw.set('dist', str(distance))
    outerShdw.set('dir', str(direction))
    outerShdw.set('algn', 'tl')
    outerShdw.set('rotWithShape', '0')
    
    # 设置阴影颜色为黑色，并指定透明度
    srgbClr = etree.SubElement(outerShdw, qn('a:srgbClr'))
    srgbClr.set('val', '000000')
    alpha_node = etree.SubElement(srgbClr, qn('a:alpha'))
    alpha_node.set('val', str(alpha))

# 3.4 设置图片位置
## 将图片置于最顶层
def bring_to_front(slide, shape):
    spTree = slide.shapes._spTree
    spTree.remove(shape._element) # 将其从当前位置删除
    spTree.append(shape._element) # 再将其添加到最后面（最顶层）

## 将图片置于最底层
def send_to_back(slide, shape):
    spTree = slide.shapes._spTree
    spTree.remove(shape._element)
    spTree.insert(2,shape._element)

# 3.5 （居中）等比例裁剪图片并填充到指定的位置和大小
def add_center_cropped_picture(slide, img_path, left, top, target_width, target_height):
    """
    :param left: 目标左上角 X 坐标（英寸）
    :param top: 目标左上角 Y 坐标（英寸）
    :param target_width: 目标宽度（英寸）
    :param target_height: 目标高度（英寸）
    """
    # 获取原图真实长宽比
    with Image.open(img_path) as im:
        img_w, img_h = im.size
    img_ratio = img_w / img_h
    target_ratio = target_width / target_height

    # 插入图片，先硬性占据目标大小
    pic = slide.shapes.add_picture(
        img_path, Inches(left), Inches(top),
        width=Inches(target_width), height=Inches(target_height)
    )

    # 根据长宽比差异，计算需要裁剪的百分比
    if img_ratio > target_ratio:
        # 原图太宽，裁两边
        crop_percent = (1 - (target_ratio / img_ratio)) / 2
        pic.crop_left = crop_percent
        pic.crop_right = crop_percent
    elif img_ratio < target_ratio:
        # 原图太高，裁上下
        crop_percent = (1 - (img_ratio / target_ratio)) / 2
        pic.crop_top = crop_percent
        pic.crop_bottom = crop_percent

    return pic

# 3.6 给形状添加阴影
def add_shadow_to_shape(shape, blur_radius=50000, distance=30000, direction=2700000, alpha=60000):
    """
    给形状（矩形、三角形等）添加原生阴影
    :param direction: 2700000 表示左上方（折角阴影通常打在左上或上方）
    """
    spPr = shape._element.spPr
    effectLst = etree.SubElement(spPr, qn('a:effectLst'))
    outerShdw = etree.SubElement(effectLst, qn('a:outerShdw'))
    
    outerShdw.set('blurRad', str(blur_radius))
    outerShdw.set('dist', str(distance))
    outerShdw.set('dir', str(direction))
    outerShdw.set('algn', 'tl')
    outerShdw.set('rotWithShape', '0')
    
    srgbClr = etree.SubElement(outerShdw, qn('a:srgbClr'))
    srgbClr.set('val', '000000') # 黑色阴影
    alpha_node = etree.SubElement(srgbClr, qn('a:alpha'))
    alpha_node.set('val', str(alpha))

# ===== 每种布局对应一个函数 =====
## 1.cover
def layout_cover(slide, data): 
    """封面布局：大标题居中 + 副标题""" 
    # 主标题居中
    draw_textbox(slide, 1.5, 2.5, 10.3, 1, 
                 data["title"], font_size=44, bold=True,
                 color=(0x1F, 0x4E, 0x79), align=PP_ALIGN.CENTER)
    # 副标题
    subtitle = data.get("subtitle", "") 
    if subtitle: 
        draw_textbox(slide, 1.5, 4.2, 10.3, 1, 
                     subtitle, font_size=24,
                     color=(0x66, 0x66, 0x66), align=PP_ALIGN.CENTER)

## 2.title_bullets
def layout_title_bullets(slide, data):
    # 标题
    draw_textbox(slide, 0.5, 1.1, 12.3, 1, 
                 data["title"], font_size=32, bold=True, 
                 color=(0x1F, 0x4E, 0x79))
    # 要点
    bullets = data.get("bullets", [])
    
    highlights = data.get("highlights", []) # 获取AI提取的高亮列表
    layout_type = data.get("layout") # 获取版式类型

    current_top = 1.8 # 正文部分的Y坐标

    for bullet in bullets:
        # 每条要点单独绘制一个文本框
        draw_highlighted_text(
            slide, 0.8, current_top, 11.7, 1.5,
            bullet, highlights,
            layout_type=layout_type,
            default_font_size=20,
            default_color=(0x33, 0x33, 0x33),
        )
        current_top += 0.5

    # 结论
    draw_textbox(slide, 0.5, 6.2, 12.3, 1, 
                 data["conclusion"], font_size=14, bold=True,
                 color=(0, 0, 0))

## 3.parallel_cards
    """多卡片布局：顶部标题 + 并列卡片，自适应卡片数量的布局，N张卡片均匀铺满一行"""
def layout_parallel_cards(slide, data, page_w=13.333, margin=0.3, gap=0.3):
    # 标题
    draw_textbox(slide, 0.5, 1.1, 12.3, 1, 
                 data["title"], font_size=32, bold=True,
                 color=(0x1F, 0x4E, 0x79))

    # 并行n张卡片
    items = data.get("items", data.get("bullets", []))
    n = min(len(items), 6)
    if n == 0: 
        return
    items = items[:n]

    # 用数量反推每张卡片的宽度
    available = page_w - 2 * margin
    card_width = (available - (n - 1) * gap) / n
    card_width = max(card_width, 1.5)
    card_height = 4.8
    top = 1.8
    
    # 图片参数
    newline = 13
    img_top = 2.2
    img_width = card_width - 0.2 * 2

    for i, item in enumerate(items):
        left = margin + i * (card_width + gap)

        # 正确提取文本和图片路径
        if isinstance(item, dict):
            text_content = item.get("text", "")
            img_path = item.get("image_path", "")
        else:
            text_content = item
            img_path = ""

		# 按索引取色，循环调用取色板
        card_color = CARD_COLORS[i % len(CARD_COLORS)]			

        card = draw_card(slide, left, top, card_width, card_height,
                         bg_color=card_color)
        
        # 在卡片上半部分绘制插图
        if img_path and os.path.exists(img_path):
            # 调用羽化函数，生成一张羽化后的图片
            faded_img_path = f"images/faded_{i}_{os.path.basename(img_path)}"
            # 如果原图为jpg，强制替换后后缀为png
            if not faded_img_path.lower().endswith(".png"):
                faded_img_path = faded_img_path.rsplit(".", 1)[0] + ".png"

            success = create_faded_image(img_path, faded_img_path, fade_width_ratio=0.15)

            if success:
                with Image.open(faded_img_path) as im:
                    w, h = im.size
            else:
                # 如果羽化失败，就用原图
                faded_img_path = img_path
                with Image.open(img_path) as im:
                    w, h = im.size
            
            # 计算缩放比例
            ratio = img_width / w
            new_h = h * ratio

            # 插入图片（位于卡片上部）
            slide.shapes.add_picture(faded_img_path, Inches(left + 0.2), Inches(img_top),
                                      width=Inches(img_width), height=Inches(new_h))
        
        card.text_frame.margin_top = Inches(3.2) # 将矩形框内的文字向下推3.2英寸

        tf = card.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.text = str(text_content)
        for para in tf.paragraphs:
            para.alignment = PP_ALIGN.CENTER
            for run in para.runs:
                set_run_font(run, 18, bold=True, color=(0xFF, 0xFF, 0xFF))
    
    # 结论
    draw_textbox(slide, 0.5, 7.0, 12.3, 1, 
                 data["conclusion"], font_size=14, bold=True,
                 color=(0, 0, 0))

## 4.left_right 
    """对比式布局：顶部主题 + 左右两张卡片进行对比"""
def layout_left_right(slide, data, page_w=13.333, margin=1, gap=2): 
    # 标题
    draw_textbox(slide, 0.5, 1.1, 12.3, 1,
                 data["title"], font_size=32, bold=True,
                 color=(0x1F, 0x4E, 0x79))

    items = data.get("bullets", [])

    card_width = (page_w - gap - 2 * margin) / 2
    card_height = 4.2
    top = 2.3

    # 创建中间分隔要点的虚线
    line = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT,
        Inches(page_w / 2), Inches(2), # 起点(left, top)
        Inches(page_w / 2), Inches(6.8) # 终点
    )

    ## 设置成虚线
    line.line.dash_style = MSO_LINE.DASH
    line.line.color.rgb = RGBColor(0x1F, 0x4E, 0x79)
    line.line.width = Pt(2)

    # 绘制包含对比要点的矩形框，并将要点填入
    for i, item in enumerate(items):
        left = margin + i * (card_width + gap)
        card = draw_card(slide, left, top, card_width, card_height,
                         bg_color=CARD_COLORS[i])
        
        tf = card.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.text = item
        for para in tf.paragraphs:
            para.alignment = PP_ALIGN.CENTER
            for run in para.runs:
                run.font.size = Pt(18)

    # 结论
    draw_textbox(slide, 0.5, 7.2, 12.3, 1, 
                 data["conclusion"], font_size=14, bold=True,
                 color=(0, 0, 0))

## 5.timeline
def layout_timeline(slide, data):
    # 标题
    draw_textbox(slide, 0.5, 0.8, 12.3, 1,
                 data["title"], font_size=32, bold=True,
                 color=(0x1F, 0x4E, 0x79))

    items = data.get("items", data.get("bullets", []))
    n = len(items)
    if n == 0: 
        return 

    # 水平线
    line_y = 4
    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(1), Inches(line_y),
        Inches(11.3), Pt(3)
    )
    line.fill.solid()
    line.fill.fore_color.rgb = RGBColor(0x1F, 0x4E, 0x79)
    line.line.fill.background()

    # 节点
    spacing = 11.3 / (n + 1)
    for i, item in enumerate(items):
        x = 1 + spacing * (i + 1)

        # 提取文本和图片路径
        if isinstance(item, dict):
            text_content = item.get("text", "")
            img_path = item.get("image_path", "")
        else:
            text_content = item
            img_path = ""

        # 圆点
        circle = slide.shapes.add_shape(
            MSO_SHAPE.OVAL,
            Inches(x - 0.15), Inches(line_y - 0.1),
            Inches(0.3), Inches(0.3)
        )
        circle.fill.solid()
        circle.fill.fore_color.rgb = RGBColor(0x1F, 0x4E, 0x79)
        circle.line.fill.background()

        # 文字（偶数时将文本分布在圆点下方，反之分布在上方）
        if i % 2 == 0: 
            draw_textbox(slide, x - 0.75, line_y + 0.4, 1.5, 1.5, 
                         str(text_content), font_size=18, color=(0, 0, 0), align=PP_ALIGN.CENTER)
        else: 
            draw_textbox(slide, x - 0.75, line_y - 1.9, 1.5, 1.5, 
                         str(text_content), font_size=18, color=(0, 0, 0), align=PP_ALIGN.CENTER)

        # 图片（分布方向与文字相反）
        if n <= 3: 
            img_width = 11.3 / 4
        else:
            img_width = 11.3 / n
        img_left = x - img_width / 2
        img_top = 1.8
        img_floor = 4.5

        if img_path and os.path.exists(img_path):
            with Image.open(img_path) as im:
                w, h = im.size
            ratio = img_width / w
            new_h = h * ratio

            if i % 2 == 0:
                pic = slide.shapes.add_picture(img_path, Inches(img_left), Inches(img_top),
                                         width=Inches(img_width), height=Inches(new_h))
            else:
                pic = slide.shapes.add_picture(img_path, Inches(img_left), Inches(img_floor),
                                         width=Inches(img_width), height=Inches(new_h))
            add_shadow_to_picture(pic, blur_radius=320000, distance=80000, alpha=80000)

    # 结论
    draw_textbox(slide, 0.5, 7.0, 12.3, 1, 
                 data["conclusion"], font_size=14, bold=True,
                 color=(0, 0, 0))

## 6.chart

## 7.quote
def layout_quote(slide, data):
    """金句布局：深色背景 + 居中大字 + 装饰引号"""
    text = data.get("conclusion", data.get("title", ""))

    # 1. 整页深色背景
    bg = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), Inches(0),
        Inches(13.333), Inches(7.5)
    )
    bg.fill.solid()
    bg.fill.fore_color.rgb = RGBColor(0x1F, 0x4E, 0x79)  # 深蓝
    bg.line.fill.background()

    # 2. 顶部装饰引号（用大号字符模拟）
    draw_textbox(slide, 1.5, 0.8, 3, 2,
                 "“", font_size=120, bold=True,
                 color=(0x4A, 0x90, 0xD9),   # 浅蓝，半透明感
                 align=PP_ALIGN.LEFT)

    # 3. 主文字居中
    draw_textbox(slide, 1.5, 2.5, 10.3, 3,
                 text, font_size=40, bold=True,
                 color=(0xFF, 0xFF, 0xFF),
                 align=PP_ALIGN.CENTER,
                 vertical=MSO_ANCHOR.MIDDLE)

    # 4. 底部装饰线
    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(5.5), Inches(5.8),
        Inches(2.3), Pt(3)
    )
    line.fill.solid()
    line.fill.fore_color.rgb = RGBColor(0x4A, 0x90, 0xD9)
    line.line.fill.background()

## 8.section
def layout_section(slide, data):
    """在左侧绘制渐变图片，右侧填入文本"""
    # 获取图片路径
    img_path = data.get("image_path", "")

    # 绘制剪切后的图片在左侧
    add_center_cropped_picture(slide, img_path, 0, 0, 6, 7.5)

    # 绘制白色蒙版并将其置于顶层
    card = draw_rectangle_mask(slide, 0, 0, 6, 7.5, bg_color=(255, 255, 255),
                               border_color=(255, 255, 255))
    bring_to_front(slide, card)

    # 绘制一个小的便签在第×部分左方
    draw_folded_note(slide, 5.7, 1.0, 2.7, 0.8, fold_size=0.3)
                                 
    # 将文本写入右侧
    draw_textbox(slide, 6.0, 1.0, 7, 2, 
                 data["title"], font_size=36, bold=True,
                 color=(0x1F, 0x4E, 0x79))

    # 将该部分内容写入
    subtitle = data.get("subtitle", "") 
    draw_textbox(slide, 7.5, 2.5, 6, 2, 
                 subtitle, font_size=40, bold=True,
                 color=(0x15, 0x73, 0xCB))

# 定义宏
## 定义layout参数
LAYOUT_DISPATCHER = {
    "cover":          layout_cover,
    "title_bullets":  layout_title_bullets,
    "parallel_cards": layout_parallel_cards,
    "left_right":     layout_left_right,
    "timeline":       layout_timeline,
    # "chart":          layout_chart,
    "quote":          layout_quote,
    "section":        layout_section,
}

# 主流程
## 读取JSON
with open(INPUT_JSON, "r", encoding="utf-8") as f:
    data = json.load(f)

## 设定幻灯片大小（16:9，采用Inches转换）
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

for i, slide_data in enumerate(data["slides"]):
    layout_type = slide_data.get("layout", "title_bullets")
    current_page = f"第{i}页"
    
    if layout_type == "section":
        current_part = slide_data.get("title")

    # 1.新建空白页
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # 2.读layout字段（默认title_bullets）
    layout_type = slide_data.get("layout", "title_bullets")

    # 3.查表找到对应函数
    layout_func = LAYOUT_DISPATCHER.get(layout_type, layout_title_bullets)

    # 4.调用函数渲染
    layout_func(slide, slide_data)

    # 5. 处理标签显示
    ## 左上角绘制当前部分
    if current_part and layout_type not in ("cover", "section"):
        draw_textbox(slide, 0.3, 0.15, 4, 0.5,
                     current_part, font_size=12, bold=True,
                     color=(0x66, 0x66, 0x66),
                     align=PP_ALIGN.LEFT)
        # 绘制分隔线
        line = slide.shapes.add_connector(
            MSO_CONNECTOR.STRAIGHT,
            Inches(0), Inches(0.65),
            Inches(13.333), Inches(0.65)
        )

        line.line.color.rgb = RGBColor(0x1F, 0x4E, 0x79)
        line.line.width = Pt(4)

    ## 绘制右下角页码
    draw_textbox(slide, 12.7, 7.2, 4, 0.5,
                 current_page, font_size=10, bold=True,
                 color=(0, 0, 0),
                 align=PP_ALIGN.LEFT)

    print(f"[Python] 第{i + 1} 页使用布局: {layout_type}")

# 保存
prs.save(OUTPUT)
print(f"[Python] 已生成最终PPT: {OUTPUT}")
