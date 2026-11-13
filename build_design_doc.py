# -*- coding: utf-8 -*-
"""生成 MEET·DOCTOR 前端交互设计说明（docx）。"""

from pathlib import Path
import unicodedata

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

BASE = Path(__file__).resolve().parent
SHOT1 = BASE / "screenshots" / "case-01-queue.png"
SHOT2 = BASE / "screenshots" / "case-04-inconclusive.png"
OUT = BASE / "MEET-DOCTOR-前端交互设计说明.docx"

INK = "1A1A1A"
AMBER = "9A6413"
MUTED = "5C5A52"
RULE = "D8D2C4"
SURFACE = "F6F2E8"

CN_BODY = "宋体"
CN_HEAD = "微软雅黑"
MONO = "Consolas"


# ---------------------------------------------------------------- helpers
def set_run(run, size=12, bold=False, latin=CN_BODY, eastasia=CN_BODY,
            color=INK, italic=False):
    run.font.name = latin
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    run.font.color.rgb = RGBColor.from_string(color)
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:ascii"), latin)
    rfonts.set(qn("w:hAnsi"), latin)
    rfonts.set(qn("w:eastAsia"), eastasia)


def para(doc, text="", size=12, bold=False, color=INK, indent=True,
         space_before=0, space_after=6, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
         latin=CN_BODY, eastasia=CN_BODY, line_spacing=1.5, italic=False):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.line_spacing = line_spacing
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    pf.alignment = align
    if indent:
        pf.first_line_indent = Pt(size * 2)
    if text:
        run = p.add_run(text)
        set_run(run, size, bold, latin, eastasia, color, italic)
    return p


def rich(doc, parts, size=12, indent=True, space_before=0, space_after=6,
         line_spacing=1.5, align=WD_ALIGN_PARAGRAPH.JUSTIFY):
    """parts: list of (text, dict(bold=, color=, italic=))"""
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.line_spacing = line_spacing
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    pf.alignment = align
    if indent:
        pf.first_line_indent = Pt(size * 2)
    for text, opt in parts:
        run = p.add_run(text)
        set_run(run, size, opt.get("bold", False), CN_BODY, CN_BODY,
                opt.get("color", INK), opt.get("italic", False))
    return p


def bullet(doc, text, size=12, space_after=4, marker="·"):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.line_spacing = 1.5
    pf.space_before = Pt(0)
    pf.space_after = Pt(space_after)
    pf.left_indent = Cm(0.9)
    pf.first_line_indent = Cm(-0.45)
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run = p.add_run(f"{marker}  ")
    set_run(run, size, color=MUTED)
    run2 = p.add_run(text)
    set_run(run2, size)
    return p


def heading(doc, text, level=1):
    size = {1: 16, 2: 13.5, 3: 12}[level]
    p = doc.add_paragraph()
    p.style = doc.styles[f"Heading {level}"]
    pf = p.paragraph_format
    pf.space_before = Pt({1: 20, 2: 14, 3: 10}[level])
    pf.space_after = Pt({1: 8, 2: 6, 3: 4}[level])
    pf.line_spacing = 1.25
    pf.first_line_indent = Pt(0)
    pf.keep_with_next = True
    run = p.add_run(text)
    set_run(run, size, True, CN_HEAD, CN_HEAD, INK)
    return p


def caption(doc, text):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf.space_before = Pt(4)
    pf.space_after = Pt(12)
    pf.line_spacing = 1.15
    run = p.add_run(text)
    set_run(run, 10.5, False, CN_BODY, CN_BODY, MUTED)


def shade(cell, hexcolor):
    tcpr = cell._tc.get_or_add_tcPr()
    fill = OxmlElement("w:shd")
    fill.set(qn("w:val"), "clear")
    fill.set(qn("w:color"), "auto")
    fill.set(qn("w:fill"), hexcolor)
    tcpr.append(fill)


def table_borders(table, color=RULE, sz=4):
    tblpr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(sz))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        borders.append(el)
    tblpr.append(borders)


def spacer(doc, pts=6):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 1.0
    run = p.add_run("")
    run.font.size = Pt(pts)
    return p


def table(doc, headers, rows, widths, size=10, header_fill=SURFACE,
          align_center=True):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    table_borders(t)

    tblpr = t._tbl.tblPr
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblpr.append(layout)

    hdr = t.rows[0]
    trpr = hdr._tr.get_or_add_trPr()
    th = OxmlElement("w:tblHeader")
    trpr.append(th)

    for i, h in enumerate(headers):
        cell = hdr.cells[i]
        cell.width = Cm(widths[i])
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        shade(cell, header_fill)
        p = cell.paragraphs[0]
        p.paragraph_format.line_spacing = 1.1
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(h)
        set_run(run, size, True, CN_HEAD, CN_HEAD, INK)

    for row in rows:
        cells = t.add_row().cells
        for i, val in enumerate(row):
            cell = cells[i]
            cell.width = Cm(widths[i])
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(1)
            if align_center and i > 0 and len(str(val)) <= 14:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(str(val))
            set_run(run, size)

    # 让整张表尽量不被分页切开
    rows_all = t.rows
    for idx, row in enumerate(rows_all):
        trpr = row._tr.get_or_add_trPr()
        trpr.append(OxmlElement("w:cantSplit"))
        if idx < len(rows_all) - 1:
            for cell in row.cells:
                for p in cell.paragraphs:
                    p.paragraph_format.keep_with_next = True

    spacer(doc, 4)
    return t


def figure(doc, path, width_cm=15.4):
    p = doc.add_paragraph()
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True
    p.add_run().add_picture(str(path), width=Cm(width_cm))


def layout_diagram(doc):
    """三栏线框示意图：用合并单元格拼出来，避免等宽字符对不齐。"""
    widths = [4.2, 7.2, 4.0]
    t = doc.add_table(rows=3, cols=3)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    table_borders(t, color=RULE, sz=6)

    tblpr = t._tbl.tblPr
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblpr.append(layout)

    for row in t.rows:
        for i, cell in enumerate(row.cells):
            cell.width = Cm(widths[i])

    head = t.cell(0, 0).merge(t.cell(0, 2))
    foot = t.cell(2, 0).merge(t.cell(2, 2))

    def fill(cell, lines, size=9, bold=False, fill_color="F2EDE0", color=MUTED):
        shade(cell, fill_color)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        first = True
        for text in lines:
            p = cell.paragraphs[0] if first else cell.add_paragraph()
            first = False
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.15
            run = p.add_run(text)
            set_run(run, size, bold, CN_HEAD, CN_HEAD, color)

    fill(head, ["品牌　·　模型状态条　·　时钟"])
    fill(foot, ["工具调用计数　·　数据来源声明　·　快捷键"])
    fill(t.cell(1, 0), ["事件队列"], bold=True, color=INK)
    fill(t.cell(1, 1), ["标题 + 事故简报 + 元信息", "RTT 包络（基线带 / ONSET）",
                        "控制条", "证据时间轴（7 步，可展开 JSON）"],
         fill_color="FAF7EE", color=INK)
    fill(t.cell(1, 2), ["假设擂台"], bold=True, color=INK)

    for cell, sub in ((t.cell(1, 0), ["05 条卷宗", "状态点", "一句话结论"]),
                      (t.cell(1, 2), ["5 条发散条", "证据芯片", "盖章结论"])):
        for text in sub:
            p = cell.add_paragraph()
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.15
            run = p.add_run(text)
            set_run(run, 9, False, CN_HEAD, CN_HEAD, MUTED)

    spacer(doc, 4)


# ---------------------------------------------------------------- document
doc = Document()
sec = doc.sections[0]
sec.page_height = Cm(29.7)
sec.page_width = Cm(21.0)
sec.top_margin = Cm(2.5)
sec.bottom_margin = Cm(2.5)
sec.left_margin = Cm(2.5)
sec.right_margin = Cm(2.5)

normal = doc.styles["Normal"]
normal.font.name = CN_BODY
normal.font.size = Pt(12)
normal.element.rPr.rFonts.set(qn("w:eastAsia"), CN_BODY)

doc.core_properties.title = "MEET·DOCTOR 时延取证台 — 前端交互设计说明"
doc.core_properties.subject = "视频会议卡顿归因 Agent 界面设计"
doc.core_properties.author = ""

# 页脚页码（fldSimple + 占位结果，Word 打开时自动刷新）
footer_p = sec.footer.paragraphs[0]
footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
fld = OxmlElement("w:fldSimple")
fld.set(qn("w:instr"), " PAGE ")
fld_run = OxmlElement("w:r")
fld_text = OxmlElement("w:t")
fld_text.text = "1"
fld_run.append(fld_text)
fld.append(fld_run)
footer_p._p.append(fld)

# ---------------- 封面区
p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(36)
p.paragraph_format.space_after = Pt(2)
p.paragraph_format.line_spacing = 1.1
run = p.add_run("MEET·DOCTOR")
set_run(run, 26, True, MONO, MONO, AMBER)

p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(4)
p.paragraph_format.line_spacing = 1.1
run = p.add_run("时延取证台")
set_run(run, 30, True, CN_HEAD, CN_HEAD, INK)

p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(28)
p.paragraph_format.line_spacing = 1.2
run = p.add_run("视频会议卡顿归因 Agent · 前端交互设计说明")
set_run(run, 14, False, CN_HEAD, CN_HEAD, MUTED)

table(
    doc,
    ["项目", "内容"],
    [
        ["文档性质", "界面设计说明（Design Rationale）"],
        ["对应原型", "projects/meetdoctor-ui/index.html（单文件，零依赖）"],
        ["实现方式", "手写 HTML / CSS / JavaScript，无框架、无构建步骤"],
        ["截图环境", "Chrome Headless，1440 × 900，?static=1 静态模式"],
        ["版本 / 日期", "v1.0 / 2026-11-02"],
    ],
    widths=[3.4, 12.0],
    size=10,
    align_center=False,
)

para(doc, "", indent=False, space_after=6)

heading(doc, "内容提要", 2)
para(doc, "本文说明 MEET·DOCTOR 时延取证台的前端交互设计。该界面服务于视频会议卡顿的归因调查，"
          "把一次网络事故从“用户投诉”推进到“可复核的结论”。设计上放弃常见的对话式布局与 AI 仪表盘风格，"
          "改用事故取证台的视觉隐喻：三栏结构分别承载分诊、取证与裁决；用发散条而非进度条表达假设的支持与推翻；"
          "用印章而非标签承载最终结论；并专门设计“证据不足”的弃权态，使 Agent 的克制在界面上同样郑重。"
          "文中同时给出色彩、字体、动效与无障碍规格，以及界面与后端 Agent 的对接约定。")

doc.add_page_break()

# ---------------- 一
heading(doc, "一、设计目标", 1)
para(doc, "本界面服务于一个具体的调查任务：值班工程师接到“视频会议卡顿”的投诉，需要在"
          "排队拥塞、路由绕行、上行不足、无线抖动、终端侧这五类假设之间做排除。界面必须回答三个问题："
          "Agent 现在查到了哪一步？每条证据在支持或反对哪个假设？最终结论凭什么成立？")

rich(doc, [("因此它不是一个“聊天窗口加结果卡片”的界面，而是一张可以被复核的取证记录："
            "每一条结论都能追溯到一次具体的工具调用、一份原始返回和一个评分动作。", {})])

para(doc, "围绕这一点，界面在设计上有三条硬约束：", space_before=2)
bullet(doc, "可追溯：任何结论都能向下展开到工具入参与返回值，不允许出现“没有出处的判断”。")
bullet(doc, "可对照：五个假设并置在同一屏，支持与反对的证据同时可见，避免只看到被选中的那一个。")
bullet(doc, "可复现：同一输入产生同一画面，便于截图进报告、也便于答辩时重放。")

# ---------------- 二
heading(doc, "二、审美方向", 1)
heading(doc, "2.1 隐喻：事故取证台", 2)
para(doc, "产品做的事情是一次事故调查——提出假设、收集证据、推翻假设、盖章下结论。"
          "因此视觉语言取自航空事故调查与仪器取证：碳黑底、琥珀色批注、编号卷宗、示波器式的波形、"
          "以及最后那枚盖下去的印章。整个界面像一张摊在桌上的调查卷宗，而不是一块监控大屏。")

heading(doc, "2.2 为什么不做常见的“AI 仪表盘”", 2)
para(doc, "深色底加霓虹渐变、发光卡片、环形进度，这类做法的潜台词是“系统正在运行，请信任它”。"
          "但本产品要传达的恰恰相反：结论正在被检验，请审视它。这与课程评分原则三"
          "“专业正确性比语言流畅更重要”是同一件事——界面不应该替 Agent 争取信任，"
          "而应该把判断的过程摊开给工程师看。")

heading(doc, "2.3 三条视觉纪律", 2)
bullet(doc, "颜色只承担语义，不承担装饰。青色只表示“证据支持”，锈红只表示“证据反对”，"
            "琥珀只表示“待判定”或强调。没有第四种用法。")
bullet(doc, "数据与叙述分离。等宽字体承载数字与时间戳，衬线斜体承载事故经过，"
            "让工程师一眼分得清“这是测出来的”和“这是人说的”。")
bullet(doc, "动效只用于表达因果顺序。假设条不会自己飘，它只在“证据评分”那一步长出来。")

# ---------------- 三
heading(doc, "三、视觉系统", 1)

heading(doc, "3.1 色彩", 2)
table(
    doc,
    ["角色", "色值", "使用位置"],
    [
        ["底色", "#0A0B0A", "全局背景"],
        ["面板", "#101211", "左栏、右栏、顶栏"],
        ["主文字", "#EDE6D6", "标题、正文、关键数值"],
        ["次文字", "#A8A493", "说明、图例、次要字段"],
        ["弱文字", "#6F6E63", "标签、时间戳、单位"],
        ["琥珀（强调）", "#E9A14C", "结论、当前步骤、重点数字"],
        ["青色（支持）", "#74D3BE", "正向证据条、实测曲线"],
        ["锈红（反对）", "#D9563B", "负向证据条、异常起点、缺口"],
    ],
    widths=[3.2, 3.4, 8.8],
    size=10,
)

heading(doc, "3.2 字体", 2)
table(
    doc,
    ["角色", "字体", "中文回退", "用途"],
    [
        ["标题 / 假设名", "Antonio", "思源黑体、微软雅黑", "压缩字身，小面板里也能用大字号"],
        ["数据 / 标签", "DM Mono", "微软雅黑", "等宽对齐时间戳、置信度、指标"],
        ["事故简报", "Newsreader 斜体", "宋体", "把叙述与数据区分开"],
    ],
    widths=[3.4, 3.6, 4.2, 4.2],
    size=10,
    align_center=False,
)
para(doc, "中文回退必须显式指定。英文展示字体不含中文字形，若不写回退，标题会退化成浏览器默认宋体，"
          "中文与英文的视觉重量会完全脱节——这是中英混排界面最常见的一处失控。", space_before=2)

heading(doc, "3.3 空间与质感", 2)
bullet(doc, "8 px 基准网格；面板之间只用 1 px 分隔线，不用投影和圆角卡片。")
bullet(doc, "噪点以 3.5% 不透明度叠加在整屏之上，扫描线以 4 px 为周期缓慢漂移，营造仪器屏幕的质感。")
bullet(doc, "左上与右下各有一处极淡的径向辉光（琥珀 / 青色），避免纯黑背景的“贴纸感”。")

# ---------------- 四
heading(doc, "四、信息架构", 1)
para(doc, "采用非对称三栏：左侧是卷宗队列，中间是取证现场，右侧是假设擂台与结论文书。"
          "三栏分别对应工程师的三种动作——分诊、取证、裁决。")

layout_diagram(doc)
caption(doc, "图 1  三栏布局示意（左：分诊　中：取证　右：裁决）")

table(
    doc,
    ["区域", "内容", "回答的问题"],
    [
        ["左栏", "5 条事件、状态点、一句话结论、置信度", "先查哪一条？哪条还没结案？"],
        ["中栏", "事故简报、RTT 包络、证据时间轴", "Agent 查到了什么？凭什么？"],
        ["右栏", "五个假设的发散条、结论印章、行动建议", "谁被支持？谁被推翻？结论是什么？"],
    ],
    widths=[2.4, 8.0, 5.0],
    size=10,
    align_center=False,
)

# ---------------- 五
heading(doc, "五、核心界面元素", 1)

heading(doc, "5.1 事件队列：用一句话结论做分诊", 2)
para(doc, "每条卷宗三行：编号与来源、一句话结论、状态与置信度。队列表里刻意不放大标题，"
          "而放“出口 38%，却在排队”这种一句话判断——因为运维在队列里做的是分诊，"
          "他要看的不是事故全名，而是“这条到底是不是我要的那类问题”。"
          "状态点用双色：青色表示已定性，琥珀表示证据不足，扫一眼就知道队列里谁还悬着。")

heading(doc, "5.2 RTT 包络：让“正常”和“越界”同屏", 2)
para(doc, "波形不是普通折线图，而是带三条注释的包络：虚线框是基线安全带（正常 RTT 区间），"
          "红色竖虚线标出异常起点 ONSET，底部锈红点阵标出采样缺失的位置。"
          "之所以要画基线带，是因为“RTT 从 25 ms 涨到 150 ms”这句话，只有和基线放在一起看，"
          "才能让工程师判断这是量变还是质变。")

heading(doc, "5.3 证据时间轴：可追溯性的落点", 2)
para(doc, "中栏主体是 7 步时间轴。每一步只占一行，但工具调用行可以展开，"
          "看到完整的入参与返回 JSON——这正好对应课程要求的“必须展示至少一次真实的工具调用过程”。"
          "时间轴左侧的时间戳是真实的调用时刻，不是装饰，它让“先聚合、再复算、再比对路径”的顺序无法被含糊过去。")

heading(doc, "5.4 假设发散条：让“被推翻”看得见", 2)
para(doc, "右侧每条假设都有一根从中心线向两侧生长的条：向右（青）是支持，向左（锈红）是反对。"
          "普通进度条只能表达“多与少”，发散条能表达“方向与推翻”——当“路由绕行”显示为一根 −70 的红条时，"
          "工程师不需要读任何文字就知道这个假设已经被排除了。这是整个界面里信息密度最高的一处设计。")

heading(doc, "5.5 结论印章：不可被滚动带走的判断", 2)
para(doc, "结论区固定在右栏底部，是一枚倾斜 3.2 度、双线边框的印章，落下时从 1.6 倍缩放砸到 1 倍并轻微回弹。"
          "印章下方是一行斜体行动建议和置信度。当 Agent 证据不足时，印章变成灰色虚线边框，"
          "文字改为“证据不足 / INCONCLUSIVE”，右上角状态变为 AGENT ABSTAINED——"
          "弃权在视觉上必须和“得出结论”一样郑重，否则 Agent 会倾向于硬给答案。")

# ---------------- 六
heading(doc, "六、一次调查的交互时序", 1)
para(doc, "点击“开始调查”后，七步以 780 ms 的间隔依次落位。每一步只做一件事，界面同步给出一个反馈，"
          "工程师可以随时按“直接出结论”快进。")

table(
    doc,
    ["步", "Agent 动作", "界面反应"],
    [
        ["1", "接收 Incident", "时间轴出现第一行，顶部状态变为 investigating"],
        ["2", "提出 5 个假设", "右栏五条假设就位，但证据条仍为零"],
        ["3", "调用 get_window_metrics", "工具行出现，可展开原始 JSON；工具计数 +1"],
        ["4", "调用 mm1_queue_delay", "同上，同时显示 ρ 与理论排队时延"],
        ["5", "调用 diff_path_snapshots", "同上，显示跳数变化与是否度量最优"],
        ["6", "证据评分", "发散条从中心线生长，得分数字点亮"],
        ["7", "得出结论", "印章砸落，行动建议与置信度出现"],
    ],
    widths=[1.2, 6.4, 7.8],
    size=10,
)

# ---------------- 七
heading(doc, "七、动效规范", 1)
table(
    doc,
    ["元素", "时长", "缓动", "目的"],
    [
        ["时间轴单步入场", "500 ms", "cubic-bezier(.2,.7,.2,1)", "上移 8 px 淡入，形成“逐条落位”"],
        ["证据条生长", "900 ms", "cubic-bezier(.2,.8,.2,1)", "把“算出来”和“长出来”绑定"],
        ["结论印章", "500 ms", "cubic-bezier(.2,1.4,.4,1)", "过冲回弹，模拟盖章的物理感"],
        ["RTT 曲线描边", "1600 ms", "cubic-bezier(.4,0,.2,1)", "从左到右绘制，强调时间流逝"],
        ["扫描线漂移", "9 s", "linear 循环", "极低幅度的环境动效，不抢注意力"],
    ],
    widths=[3.4, 2.0, 6.0, 4.0],
    size=10,
)
para(doc, "动效原则：单次调查的总动画时长控制在 6 秒内；任何动效都不承载信息——"
          "关掉动画后，所有结论、证据与数值仍然完整可读。这一点在静态截图模式下有直接验证。")

# ---------------- 八
heading(doc, "八、无障碍与降级", 1)
bullet(doc, "尊重 prefers-reduced-motion：该媒体查询生效时，全部动画与扫描线关闭。")
bullet(doc, "键盘可达：左栏可用 Tab 聚焦、回车选中；数字键 1–5 直接切换事件。")
bullet(doc, "焦点可见：所有按钮有 2 px 青色焦点描边，不依赖浏览器默认样式。")
bullet(doc, "语义标注：事件队列为 listbox，时间轴为 aria-live=polite，波形提供替代文本。")
bullet(doc, "窄屏降级：宽度小于 1180 px 时隐藏右栏，小于 820 px 时收为单栏。")
bullet(doc, "对比度：主文字与底色对比度约 15:1，次要文字约 6:1，均高于 WCAG AA 要求。")

# ---------------- 九
heading(doc, "九、五个案例的设计意图", 1)
para(doc, "五个案例不是同一问题的数字变体，而是刻意覆盖不同的判断情形，直接对应课程第五部分的测试设计要求。")

table(
    doc,
    ["案例", "演示的 Agent 能力", "对应的考察点"],
    [
        ["C-001 出口 38% 却在排队", "用峰值而非均值推翻“带宽不足”的直觉", "专业误判边界"],
        ["C-002 RTT 翻 5 倍", "路径变了但仍是度量最优，据此排除错误归因", "反面证据"],
        ["C-003 只有上行在丢帧", "分方向拆指标，区分出口拥塞与上行溢出", "实质性计算"],
        ["C-004 六成采样缺失", "得分全部落在 ±0.2，主动弃权", "证据不足与不确定性"],
        ["C-005 全网正常只有一人卡", "承认超出网络域，给出可验证的下一步", "能力边界"],
    ],
    widths=[4.6, 7.4, 3.4],
    size=10,
    align_center=False,
)

# ---------------- 十
heading(doc, "十、工程实现与对接约定", 1)
bullet(doc, "单文件 45 KB，零依赖、零构建，双击即可运行；所有样式走 CSS 变量，改一处即可整体换肤。")
bullet(doc, "波形由 mulberry32 伪随机数在现场生成，固定种子 2026，因此刷新不会变形，也便于复现截图。")
bullet(doc, "URL 参数：case 指定启动案例，instant 跳过逐步动画，static 关闭全部动画用于稳定截图。")
bullet(doc, "数据模型与后端 Pydantic 输出一一对应：id、title、brief、meta、metrics、weights、"
            "evidence、hits、verdict（zh/en/tone/conf/action）、steps。")

para(doc, "接入真实 Agent 时不需要重写界面：把 runSequence() 中遍历本地 steps 的循环，"
          "换成消费后端事件流即可，每个事件带一个 kind 字段（recv / hypo / tool / score / verdict），"
          "右栏证据条直接读取后端返回的 weights 与 evidence。", space_before=2)

# ---------------- 十一
heading(doc, "十一、局限与后续", 1)
bullet(doc, "字体依赖外部字体服务，离线时回退到系统字体，标题的压缩感会有所损失。")
bullet(doc, "移动端只做了隐藏与单栏降级，未针对触屏重排；当前定位是桌面端答辩演示。")
bullet(doc, "案例数据内联在前端，尚未接入真实 Agent，评分动作也是按脚本重放的。")
bullet(doc, "后续可做：点击证据芯片联动高亮对应时间轴步骤；亮色主题；把案例抽成 cases.json 供后端驱动。")

# ---------------- 附录
heading(doc, "附录　设计决策清单", 1)
table(
    doc,
    ["决策", "被放弃的做法", "理由"],
    [
        ["三栏而非单栏对话", "聊天式界面", "对话流无法并置对照五个假设"],
        ["发散条而非进度条", "单向百分比条", "只有方向感才能表达“假设被推翻”"],
        ["印章而非徽章", "圆角标签", "结论需要不可滚走的重量感"],
        ["波形带基线带", "纯折线图", "没有基线就无法判断是否越界"],
        ["弃权态单独设计", "复用成功态", "弃权必须与结论同等郑重，否则 Agent 倾向硬答"],
        ["静态截图模式", "直接截动画", "报告截图必须可复现、可对比"],
    ],
    widths=[4.0, 3.8, 7.6],
    size=10,
    align_center=False,
)

# ---------------- 截图
heading(doc, "附图", 1)
figure(doc, SHOT1)
caption(doc, "图 2  案例 C-001 结论态：印章判定“排队拥塞”，证据条 +86 与 −70 同屏可见")

figure(doc, SHOT2)
caption(doc, "图 3  案例 C-004 弃权态：得分落在 ±0.2 内，印章转为灰色虚线“证据不足”")

doc.save(str(OUT))
print("saved:", OUT)
