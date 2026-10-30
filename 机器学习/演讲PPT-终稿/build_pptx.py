# -*- coding: utf-8 -*-
"""
将演讲PPT（Beamer 版）原生重建为 WPS 可编辑的 16:9 PPTX。
真实文本框 / 表格 / 形状 / 图片，全部可编辑。
运行：python3 build_pptx.py
"""
import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

BASE = os.path.dirname(os.path.abspath(__file__))
FIG = os.path.join(BASE, "figures")
PIC = os.path.join(BASE, "pic")

# ---- 主题色（与 UESTC Beamer 模板一致）----
BLUE   = RGBColor(0x11, 0x5E, 0xA2)
LIGHT  = RGBColor(0xE9, 0xE9, 0xF3)
GREY   = RGBColor(0x59, 0x59, 0x59)
LBLUE  = RGBColor(0xBF, 0xD9, 0xEC)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
DARK   = RGBColor(0x1A, 0x1A, 0x1A)
GREEN  = RGBColor(0x00, 0x80, 0x00)
DBLUE  = RGBColor(0x00, 0x00, 0x80)

HEI = "微软雅黑"   # 黑体/标题
KAI = "楷体"       # 楷体（装饰用）
MONO = "Consolas"

TOTAL = 14

prs = Presentation()
prs.slide_width  = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]

# ---------------- 工具函数 ----------------

def set_font(run, name=HEI, size=18, bold=False, color=DARK, italic=False, ea=None):
    f = run.font
    f.name = name
    f.size = Pt(size)
    f.bold = bold
    f.italic = italic
    f.color.rgb = color
    rPr = run._r.get_or_add_rPr()
    ea_name = ea if ea is not None else name
    for tag in ("a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {})
            rPr.append(el)
        el.set("typeface", ea_name)

def tb(slide, x, y, w, h, anchor=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = Inches(0); tf.margin_right = Inches(0)
    tf.margin_top = Inches(0); tf.margin_bottom = Inches(0)
    return box, tf

def para(tf, runs, size=18, color=DARK, bullet=None, level=0, space_after=6,
         space_before=0, align=None, first=False, line_spacing=1.0):
    p = tf.paragraphs[0] if (first and not tf.paragraphs[0].runs) else tf.add_paragraph()
    p.level = level
    p.space_after = Pt(space_after)
    p.space_before = Pt(space_before)
    if line_spacing:
        p.line_spacing = line_spacing
    if align is not None:
        p.alignment = align
    if bullet:
        rb = p.add_run(); rb.text = bullet
        set_font(rb, HEI, size, False, BLUE)
    if isinstance(runs, str):
        runs = [(runs, False, None)]
    for it in runs:
        text, bold = it[0], it[1]
        col = it[2] if len(it) > 2 and it[2] is not None else color
        rb = p.add_run(); rb.text = text
        set_font(rb, HEI, size, bold, col)
    return p

def rect(slide, x, y, w, h, fill=None, line=None, shape=MSO_SHAPE.RECTANGLE, line_w=1.0, shadow=False):
    sp = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid(); sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line; sp.line.width = Pt(line_w)
    sp.shadow.inherit = False
    return sp

def set_text(sp, text, size=18, color=WHITE, bold=True, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE, name=HEI, ml=0.1, mr=0.1):
    tf = sp.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(ml); tf.margin_right = Inches(mr)
    tf.margin_top = Inches(0.02); tf.margin_bottom = Inches(0.02)
    tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    p.alignment = align
    if isinstance(text, str):
        text = [(text, bold, color)]
    for t, b, c in text:
        rb = p.add_run(); rb.text = t
        set_font(rb, name, size, b, c)
    return sp

def add_bar(slide, title, speaker):
    """顶部蓝色标题栏 + 主讲人标注"""
    rect(slide, 0, 0, 13.333, 0.85, fill=BLUE)
    _, tf = tb(slide, 0.45, 0.12, 10.4, 0.6, anchor=MSO_ANCHOR.MIDDLE)
    para(tf, [(title, True)], size=22, color=WHITE, first=True, space_after=0)
    _, tf2 = tb(slide, 9.3, 0.12, 3.55, 0.6, anchor=MSO_ANCHOR.MIDDLE)
    para(tf2, [("主讲 ", False, LBLUE), (speaker, False, LBLUE)], size=13,
         align=PP_ALIGN.RIGHT, first=True, space_after=0)

def add_footer(slide, n):
    rect(slide, 0, 7.06, 13.333, 0.016, fill=LBLUE)
    _, tf = tb(slide, 0.45, 7.12, 8.0, 0.3)
    para(tf, "基于机器学习的台风分类与路径预测 · 徐诗霖 蔡奕锴 李炜焜",
         size=9, color=GREY, first=True, space_after=0)
    _, tf2 = tb(slide, 12.3, 7.12, 0.6, 0.3)
    para(tf2, "%d / %d" % (n, TOTAL), size=9, color=GREY,
         align=PP_ALIGN.RIGHT, first=True, space_after=0)

def block(slide, x, y, w, h, title, body, title_size=14, body_size=13):
    """浅灰底 + 蓝色边框的块，title 为蓝色粗体"""
    rect(slide, x, y, w, h, fill=LIGHT, line=BLUE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, line_w=1.0)
    _, tf = tb(slide, x + 0.16, y + 0.10, w - 0.32, h - 0.18)
    if title:
        para(tf, [(title, True)], size=title_size, color=BLUE, first=True, space_after=4)
    for item in body:
        if isinstance(item, str):
            item = [(item, False)]
        para(tf, item, size=body_size, color=DARK, bullet="•  ", space_after=3, line_spacing=1.02)
    return tf

def image(slide, path, x, y, w=None, h=None):
    return slide.shapes.add_picture(path, Inches(x), Inches(y),
                                    width=Inches(w) if w else None,
                                    height=Inches(h) if h else None)

def caption(slide, x, y, w, text, size=10):
    _, tf = tb(slide, x, y, w, 0.28)
    para(tf, text, size=size, color=GREY, align=PP_ALIGN.CENTER, first=True, space_after=0)

# ==================== 幻灯片 1：封面 ====================
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, 13.333, 7.5, fill=WHITE)
# 顶部蓝色装饰条
rect(s, 0, 0, 13.333, 0.18, fill=BLUE)
# 标题框
tbox = rect(s, 1.4, 1.55, 10.53, 1.15, fill=BLUE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
set_text(tbox, "基于机器学习的台风分类与路径预测", size=32, color=WHITE, bold=True)
_, tf = tb(s, 1.5, 2.95, 10.33, 0.5)
para(tf, "机器学习课程分组演讲 · 基于 2024 MathorCup 大数据挑战赛赛道 A",
     size=16, color=GREY, align=PP_ALIGN.CENTER, first=True)
_, tf = tb(s, 1.5, 3.55, 10.33, 0.6)
para(tf, [("徐诗霖", True), ("　　蔡奕锴", True), ("　　李炜焜", True)],
     size=22, color=DARK, align=PP_ALIGN.CENTER, first=True)
_, tf = tb(s, 1.5, 4.2, 10.33, 0.45)
para(tf, "计算机科学与工程学院", size=15, color=GREY, align=PP_ALIGN.CENTER, first=True)
_, tf = tb(s, 1.5, 4.62, 10.33, 0.45)
para(tf, "2026 年 10 月", size=15, color=GREY, align=PP_ALIGN.CENTER, first=True)
logo = image(s, os.path.join(PIC, "logo.png"), 0, 5.2, h=1.15)
logo.left = int((prs.slide_width - logo.width) / 2)

# ==================== 幻灯片 2：报告提纲 ====================
s = prs.slides.add_slide(BLANK)
add_bar(s, "报告提纲", "全体")
add_footer(s, 2)
outline = [
    ("1. 问题定义", " —— 研究背景 · 三个核心子问题", "约 1 分钟 · 主讲：徐诗霖"),
    ("2. 技术方案", " —— 数据工程 · 分类 / 路径预测 / 登陆衰减建模", "约 2 分钟 · 主讲：徐诗霖、蔡奕锴、李炜焜"),
    ("3. 实验结果", " —— 台风“贝碧嘉”案例验证", "约 1.5 分钟 · 主讲：蔡奕锴、李炜焜"),
    ("4. 结论与展望", " —— 技术亮点 · 不足与改进", "约 0.5 分钟 · 主讲：徐诗霖"),
]
y = 1.45
for head, rest, time in outline:
    _, tf = tb(s, 1.2, y, 11.6, 0.75, anchor=MSO_ANCHOR.MIDDLE)
    para(tf, [(head, True), (rest, False), ("　　　" + time, False, GREY)],
         size=19, color=DARK, first=True, space_after=0)
    rect(s, 1.2, y + 0.72, 11.6, 0.012, fill=LBLUE)
    y += 0.86
_, tf = tb(s, 1.2, 5.9, 10.93, 0.6)
para(tf, "陈述严格控制在 5 分钟以内，随后进入提问与点评；组内三人分板块主讲，全员参与问答。",
     size=13, color=GREY, align=PP_ALIGN.CENTER, first=True)

# ==================== 幻灯片 3：研究背景 ====================
s = prs.slides.add_slide(BLANK)
add_bar(s, "研究背景：台风预测为何是机器学习问题", "徐诗霖")
add_footer(s, 3)
_, tf = tb(s, 0.6, 1.15, 6.0, 5.6)
para(tf, [("现实需求", True)], size=18, color=DARK, first=True, space_after=4)
para(tf, "台风突发性强、破坏力大，狂风、暴雨、风暴潮并发", size=15, bullet="•  ", space_after=5)
para(tf, "准确预测可延长预警提前量，支撑人员疏散与应急资源调配", size=15, bullet="•  ", space_after=10)
para(tf, [("预测难点", True)], size=18, color=DARK, space_after=4)
para(tf, "内力（地转偏向力）与外力（副热带高压）耦合作用", size=15, bullet="•  ", space_after=5)
para(tf, "路径与影响因素高度非线性，存在打转、停滞现象", size=15, bullet="•  ", space_after=5)
block(s, 7.0, 1.15, 5.73, 3.6, "机器学习视角",
      [[("台风路径预测本质上是一个", False), ("时空序列预测问题", True)],
       [("时序依赖：台风自身历史运动状态", False)],
       [("空间融合：气压、风场、海温等环境因素", False)],
       [("数据基础：1945–2023 年近 80 年观测", False)]],
      title_size=15, body_size=14)
_, tf = tb(s, 7.0, 5.0, 5.73, 1.4)
para(tf, "研究价值：提升精度 → 延长预警提前量 → 指导防风加固与人员疏散",
     size=13, color=BLUE, first=True, line_spacing=1.15)

# ==================== 幻灯片 4：三个核心子问题 ====================
s = prs.slides.add_slide(BLANK)
add_bar(s, "问题定义：三个核心子问题", "徐诗霖")
add_footer(s, 4)
rows = [
    ("子问题", "任务（输入 → 输出）", "验证案例", "评价指标"),
    ("问题一\n台风分类", "强度、风速、气压等特征 → 类别；区分夏季 / 秋季台风", "2024 年 7 月、9 月台风", "准确率、混淆矩阵"),
    ("问题二\n路径预测", "历史运动状态 + 环境因素 → 未来 6 / 12 / 24 h 中心位置", "“贝碧嘉” 9.13–17", "大圆距离、RMSE、DTW"),
    ("问题三\n登陆衰减", "风速 V → 降水 P；距中心距离 d → 降水 P", "“贝碧嘉” 9.16–18", "RMSE、R²、残差分析"),
]
tw = [1.7, 5.6, 2.5, 2.4]
tbl = s.shapes.add_table(4, 4, Inches(0.7), Inches(1.25), Inches(12.0), Inches(3.6)).table
for i, w in enumerate(tw):
    tbl.columns[i].width = Inches(w)
tbl.rows[0].height = Inches(0.7)
for r in range(1, 4):
    tbl.rows[r].height = Inches(0.95)
for r in range(4):
    for c in range(4):
        cell = tbl.cell(r, c)
        cell.margin_left = Inches(0.08); cell.margin_right = Inches(0.08)
        cell.margin_top = Inches(0.04); cell.margin_bottom = Inches(0.04)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        tfc = cell.text_frame; tfc.word_wrap = True
        txt = rows[r][c].split("\n")
        for j, line in enumerate(txt):
            p = tfc.paragraphs[0] if j == 0 else tfc.add_paragraph()
            p.alignment = PP_ALIGN.CENTER if c == 0 or r == 0 else PP_ALIGN.LEFT
            if c == 0:
                p.alignment = PP_ALIGN.CENTER
            rb = p.add_run(); rb.text = line
            if r == 0:
                set_font(rb, HEI, 14, True, WHITE)
            elif c == 0:
                set_font(rb, HEI, 13, True, BLUE)
            else:
                set_font(rb, HEI, 13, False, DARK)
        if r == 0:
            cell.fill.solid(); cell.fill.fore_color.rgb = BLUE
        elif c == 0:
            cell.fill.solid(); cell.fill.fore_color.rgb = LIGHT
        else:
            cell.fill.solid(); cell.fill.fore_color.rgb = WHITE
block(s, 0.7, 5.35, 12.0, 1.35, None,
      [[("三者共享同一套数据管线与评估思想：均为", False), ("从历史观测中学习非线性映射", True),
        ("，差异仅在映射的目标空间。", False)]],
      body_size=14)

# ==================== 幻灯片 5：技术路线总览 ====================
s = prs.slides.add_slide(BLANK)
add_bar(s, "技术路线总览", "徐诗霖")
add_footer(s, 5)
# 顶部：数据 → 数据工程
data = rect(s, 0.7, 1.25, 3.0, 1.0, fill=LIGHT, line=BLUE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
set_text(data, [("数据", True, BLUE)], size=13)
tf = data.text_frame
para(tf, "1945–2023 台风路径集合", size=11, color=DARK, align=PP_ALIGN.CENTER, line_spacing=1.0)
para(tf, "+ 2024 年个例观测", size=11, color=DARK, align=PP_ALIGN.CENTER, line_spacing=1.0)
ar = rect(s, 3.85, 1.62, 0.55, 0.26, fill=BLUE, shape=MSO_SHAPE.RIGHT_ARROW)
prep = rect(s, 4.55, 1.25, 3.0, 1.0, fill=LIGHT, line=BLUE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
set_text(prep, [("数据工程", True, BLUE)], size=13)
tf = prep.text_frame
para(tf, "插值 · 去异常 · 标准化 · PCA", size=11, color=DARK, align=PP_ALIGN.CENTER, line_spacing=1.0)
# 三个模型框
models = [
    ("问题一：分类", "K-Means + 决策树 / SVM / LSTM"),
    ("问题二：路径预测", "SVR → LSTM → CNN-LSTM+注意力"),
    ("问题三：登陆衰减", "幂律 / 指数 / 高斯 + RF / GPR"),
]
mx = [0.7, 4.7, 8.7]
for i, (t, sub) in enumerate(models):
    box = rect(s, mx[i], 2.75, 3.6, 1.0, fill=LIGHT, line=BLUE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    set_text(box, [(t, True, BLUE)], size=13)
    tf = box.text_frame
    para(tf, sub, size=10.5, color=DARK, align=PP_ALIGN.CENTER, line_spacing=1.0)
    # 数据工程 → 模型框 的向下箭头
    down = rect(s, mx[i] + 1.62, 2.38, 0.26, 0.34, fill=BLUE, shape=MSO_SHAPE.DOWN_ARROW)
# 评估框
ev = rect(s, 2.3, 4.3, 8.7, 0.9, fill=BLUE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
set_text(ev, [("评　估", True, WHITE)], size=13)
tf = ev.text_frame
para(tf, "混淆矩阵 · RMSE · DTW · 残差分析", size=12, color=WHITE, align=PP_ALIGN.CENTER)
for i in range(3):
    down = rect(s, mx[i] + 1.62, 3.82, 0.26, 0.45, fill=BLUE, shape=MSO_SHAPE.DOWN_ARROW)
_, tf = tb(s, 0.7, 5.6, 11.9, 1.2)
para(tf, [("技术主线：", True), ("历史序列建模、环境特征融合、直接预测位移、独立事件验证", False)],
     size=15, first=True, space_after=4)
para(tf, "三个子问题共享数据管线，评估口径按任务分别设定", size=13, color=GREY, bullet="•  ")

# ==================== 幻灯片 6：数据工程 ====================
s = prs.slides.add_slide(BLANK)
add_bar(s, "数据工程与特征分析", "徐诗霖")
add_footer(s, 6)
_, tf = tb(s, 0.6, 1.15, 5.9, 4.6)
para(tf, [("1. 缺失填补：", True), ("拉格朗日插值（移动方向、速度）", False)], size=14, first=True, space_after=8)
para(tf, [("2. 异常检测：", True), ("箱线法剔除野值（风速 >110 m/s 等）", False)], size=14, space_after=8)
para(tf, [("3. 标准化：", True), ("Z=(X−μ)/σ，参数仅由训练集估计", False)], size=14, space_after=8)
para(tf, [("4. 降维：", True), ("PCA，累计贡献率 >85%", False)], size=14, space_after=10)
para(tf, [("数据划分", True, BLUE)], size=14, space_after=3)
para(tf, "按年份与完整台风事件划分训练 / 验证 / 测试集，同一台风不跨集合，杜绝信息泄漏。",
     size=13, color=DARK, line_spacing=1.15)
img = image(s, os.path.join(FIG, "热力图.png"), 6.9, 1.2, w=3.3)
caption(s, 6.9, 1.2 + img.height / 914400 + 0.02, 3.3, "斯皮尔曼相关系数矩阵")
block(s, 0.6, 5.35, 12.13, 1.4, "特征分析的独立判断",
      [[("风速–等级 r=0.996、风速–气压 r=−0.943：强相关提示", False),
        ("特征冗余", True),
        ("，为特征筛选与模型输入设计提供量化依据，而非全量喂入模型。", False)]],
      body_size=13)

# ==================== 幻灯片 7：问题一分类 ====================
s = prs.slides.add_slide(BLANK)
add_bar(s, "问题一：台风分类模型", "徐诗霖")
add_footer(s, 7)
i1 = image(s, os.path.join(FIG, "肘部法.jpg"), 0.7, 1.2, w=2.75)
caption(s, 0.7, 1.2 + i1.height / 914400 + 0.02, 2.75, "肘部法则确定聚类数 K=3")
i2 = image(s, os.path.join(FIG, "台风聚类的结果.jpg"), 0.7, 3.65, w=2.75)
caption(s, 0.7, 3.65 + i2.height / 914400 + 0.02, 2.75, "PCA 降维后的聚类结构")
_, tf = tb(s, 3.9, 1.15, 8.8, 2.0)
para(tf, [("两阶段方法", True)], size=17, first=True, space_after=5)
para(tf, [("无监督探索：", True), ("K-Means 聚类观察特征空间的自然结构", False)], size=14, bullet="•  ", space_after=5)
para(tf, [("监督建模：", True), ("决策树 → SVM → LSTM，交叉验证比较精度", False)], size=14, bullet="•  ")
block(s, 3.9, 3.15, 8.8, 3.5, "结论：夏 / 秋台风差异",
      [[("夏台风（6–8 月）：数量多、风速高，多向西北移动", False)],
       [("秋台风（9–11 月）：生成纬度更高，平均强度不弱、路径更多变", False)],
       [("2024 年 7 月 / 9 月个例的判别结果与气象事实一致", False)]],
      body_size=14)

# ==================== 幻灯片 8：问题二模型 ====================
s = prs.slides.add_slide(BLANK)
add_bar(s, "问题二：递进式路径预测模型", "蔡奕锴")
add_footer(s, 8)
_, tf = tb(s, 0.6, 1.1, 12.13, 0.6)
para(tf, [("外推基线", True, BLUE), (" → ", False, GREY), ("SVR", True, BLUE),
          (" → ", False, GREY), ("LSTM", True, BLUE), (" → ", False, GREY),
          ("CNN-LSTM", True, BLUE), (" → ", False, GREY), ("+通道注意力", True, BLUE)],
     size=17, align=PP_ALIGN.CENTER, first=True)
_, tf = tb(s, 0.6, 1.62, 12.13, 0.5)
para(tf, "每一层针对上一层的明确缺陷：无法响应环境变化 → 忽略时序依赖 → 缺空间特征 → 特征无差异加权",
     size=11, color=GREY, align=PP_ALIGN.CENTER, first=True)
_, tf = tb(s, 0.6, 2.25, 12.13, 1.2)
para(tf, [("LSTM：", True), ("三门控机制（遗忘 / 输入 / 输出）建模过去 24 h 状态序列的长程依赖", False)],
     size=13, bullet="•  ", first=True, space_after=4)
para(tf, [("CNN–LSTM 双分支：", True), ("CNN 提取环境场空间特征，LSTM 提取时序特征，经通道注意力自适应加权融合", False)],
     size=13, bullet="•  ")
img = image(s, os.path.join(FIG, "神经网络通道图.jpg"), 0, 3.5, w=7.6)
img.left = int((prs.slide_width - img.width) / 2)
_, tf = tb(s, 0.6, 6.35, 12.13, 0.5)
para(tf, "通道注意力：压缩（Squeeze）全局信息 → 激励（Excitation）生成通道权重，突出关键特征通道",
     size=11, color=GREY, align=PP_ALIGN.CENTER, first=True)

# ==================== 幻灯片 9：问题二关键设计 ====================
s = prs.slides.add_slide(BLANK)
add_bar(s, "问题二：关键设计与实验规范", "蔡奕锴")
add_footer(s, 9)
codebox = rect(s, 0.6, 1.2, 6.4, 3.4, fill=RGBColor(0xF5, 0xF7, 0xFA),
               line=BLUE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
_, tf = tb(s, 0.85, 1.35, 5.9, 3.1)
code_lines = [
    ("# 目标重参数化：预测位移而非速度", True),
    ("delta = pos[t + h] - pos[t]   # h = 6/12/24 h", False),
    ("X = feats[t-24:t]             # 过去 24 h 状态序列", False),
    ("y_hat = CNN_LSTM_SE(X)        # 时空 + 通道注意力", False),
    ("loss = mse_loss(y_hat, delta)", False),
    ("", False),
    ("# 防信息泄漏：按完整台风事件划分", True),
    ("train, test = split_by_storm_id(...)", False),
]
for i, (line, is_comment) in enumerate(code_lines):
    para(tf, line, size=12, color=GREEN if is_comment else DARK,
         first=(i == 0), space_after=2, line_spacing=1.05)
_, tf = tb(s, 7.35, 1.25, 5.4, 3.3)
para(tf, [("防信息泄漏：", True), ("标准化参数仅由训练集估计；同一台风的相邻样本不跨集合", False)],
     size=13, bullet="•  ", first=True, space_after=8)
para(tf, [("消融实验：", True), ("逐一移除 CNN 分支 / 注意力模块，量化各模块贡献", False)],
     size=13, bullet="•  ", space_after=8)
para(tf, [("困难样本：", True), ("单独分析转向、登陆、快速移动样本", False)],
     size=13, bullet="•  ")
block(s, 0.6, 4.9, 12.13, 1.7, "为什么直接预测未来位移？",
      [[("以“移动速度”为输出时，位置需经积分还原，误差随步长累积且起点不定；", False),
        ("位移目标与“预测未来位置”任务直接对齐，逐段累加即可还原完整路径。", True)]],
      body_size=13)

# ==================== 幻灯片 10：问题三 ====================
s = prs.slides.add_slide(BLANK)
add_bar(s, "问题三：登陆后风速与降水衰减建模", "李炜焜")
add_footer(s, 10)
block(s, 0.6, 1.15, 6.0, 2.6, "物理先验：风速 — 降水关系",
      [[("动力学：强风增强对流、促进水汽上升；热力学：强风加剧海面蒸发", False)],
       [("幂律模型：P = aVᵇ", False)],
       [("指数模型：P = aeᵇᵛ", False)],
       [("多项式：P = c₀ + c₁V + ⋯ + cₙVⁿ", False)]],
      body_size=13)
block(s, 0.6, 3.95, 6.0, 2.0, "空间衰减：降水 — 距离关系",
      [[("指数衰减：P = P₀·e^(−λd)", False)],
       [("高斯分布：P = P₀·e^(−(d−μ)²/(2σ²))", False)]],
      body_size=13)
block(s, 6.85, 1.15, 5.88, 2.2, "多变量数据驱动扩展",
      [[("引入方位角、湿度、风切变、地形等协变量，采用随机森林、高斯过程回归刻画降水分布的空间不对称性", False)]],
      body_size=13)
_, tf = tb(s, 6.85, 3.6, 5.88, 2.4)
para(tf, "多模型竞争拟合 + 残差分析择优", size=14, bullet="•  ", first=True, space_after=6)
para(tf, "体现“物理先验 + 数据驱动”的互补建模思想", size=14, bullet="•  ", space_after=6)
para(tf, [("验证：", True), ("均方误差 MSE、RMSE、决定系数 R² 与残差分析", False)],
     size=13, color=GREY, space_after=0)

# ==================== 幻灯片 11：路径预测结果 ====================
s = prs.slides.add_slide(BLANK)
add_bar(s, "路径预测结果：台风“贝碧嘉”（9 月 13–17 日每日 14 时）", "蔡奕锴")
add_footer(s, 11)
img = image(s, os.path.join(FIG, "路径对比图.png"), 0.7, 1.3, w=4.9)
caption(s, 0.7, 1.3 + img.height / 914400 + 0.02, 4.9, "预测路径与实际路径对比")
block(s, 5.9, 1.3, 6.8, 1.0, "定量指标",
      [[("RMSE = 0.232　　　DTW 距离 = 94.25", True)]], body_size=15)
block(s, 5.9, 2.55, 6.8, 3.4, "定性分析",
      [[("大方向与整体趋势准确，平稳段高度重合", False)],
       [("急转弯处预测滞后：时序依赖未完全捕捉", False)],
       [("训练 / 验证 / 测试损失接近，无明显过拟合", False)]],
      body_size=14)

# ==================== 幻灯片 12：降水建模结果 ====================
s = prs.slides.add_slide(BLANK)
add_bar(s, "降水建模结果：台风“贝碧嘉”登陆后（9 月 16–18 日）", "李炜焜")
add_footer(s, 12)
img1 = image(s, os.path.join(FIG, "贝碧嘉台风预测降水量变化.png"), 0.7, 1.35, w=5.6)
caption(s, 0.7, 1.35 + img1.height / 914400 + 0.02, 5.6, "综合预测模型捕捉登陆后降水衰减趋势")
img2 = image(s, os.path.join(FIG, "模型预测降水量与实际降水量的比较.png"), 7.0, 1.35, w=5.6)
caption(s, 7.0, 1.35 + img2.height / 914400 + 0.02, 5.6, "残差围绕零线分布、无系统性偏差")
block(s, 0.7, 6.0, 11.93, 0.85, "结论",
      [[("综合预测模型显著优于单一拟合；风速与降水正相关，降水量在台风中心附近达到峰值、随距离呈指数衰减。", False)]],
      body_size=13)

# ==================== 幻灯片 13：结论 ====================
s = prs.slides.add_slide(BLANK)
add_bar(s, "结论：技术亮点与展望", "徐诗霖")
add_footer(s, 13)
block(s, 0.6, 1.1, 12.13, 3.1, "技术亮点",
      [[("严格数据治理：", True), ("标准化仅用训练集、按完整台风事件划分——从源头杜绝信息泄漏", False)],
       [("递进式基线 + 消融实验：", True), ("每一层模型的引入都有明确动机与可量化贡献", False)],
       [("目标重参数化：", True), ("直接预测未来位移，预测目标与任务对齐", False)],
       [("物理先验 + 数据驱动：", True), ("幂律 / 指数 / 高斯约束结合随机森林 / 高斯过程回归", False)]],
      body_size=13)
_, tf = tb(s, 0.6, 4.45, 5.9, 1.9)
para(tf, [("不足", True)], size=15, first=True, space_after=4)
para(tf, "环境场数据（ERA5 风场、位势高度场）尚未融合", size=13, bullet="•  ", space_after=4)
para(tf, "转向拐点预测存在滞后", size=13, bullet="•  ")
_, tf = tb(s, 6.85, 4.45, 5.9, 1.9)
para(tf, [("展望", True)], size=15, first=True, space_after=4)
para(tf, "集成学习降低预测方差", size=13, bullet="•  ", space_after=4)
para(tf, "物理约束神经网络（PINN）注入动力学先验", size=13, bullet="•  ")
_, tf = tb(s, 0.6, 6.35, 12.13, 0.5)
para(tf, "主要参考：Gao et al. (2018) LSTM 路径预报 · Giffard-Roisin et al. (2020) 环境场融合位移预测 · Liu et al. (2022) DBF-Net · He et al. (2023) 双注意力 ConvLSTM",
     size=9.5, color=GREY, align=PP_ALIGN.CENTER, first=True)

# ==================== 幻灯片 14：谢谢 ====================
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, 13.333, 7.5, fill=WHITE)
rect(s, 0, 0, 13.333, 0.18, fill=BLUE)
_, tf = tb(s, 1.5, 1.7, 10.33, 1.2)
para(tf, "谢谢！", size=54, color=BLUE, align=PP_ALIGN.CENTER, first=True)
_, tf = tb(s, 1.5, 3.15, 10.33, 0.6)
para(tf, "恳请各位老师、同学批评指正", size=20, color=DARK, align=PP_ALIGN.CENTER, first=True)
_, tf = tb(s, 1.5, 4.1, 10.33, 0.6)
para(tf, "主讲分工：徐诗霖（问题定义 · 分类 · 结论）　蔡奕锴（路径预测）　李炜焜（衰减建模）",
     size=14, color=GREY, align=PP_ALIGN.CENTER, first=True)
_, tf = tb(s, 1.5, 4.8, 10.33, 0.5)
para(tf, "计算机科学与工程学院　　·　　2026 年 10 月", size=13, color=GREY,
     align=PP_ALIGN.CENTER, first=True)

out = os.path.join(BASE, "演讲PPT.pptx")
prs.save(out)
print("saved:", out, "slides:", len(prs.slides._sldIdLst))
