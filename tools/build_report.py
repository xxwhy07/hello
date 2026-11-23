# -*- coding: utf-8 -*-
"""Build the Word lab report for 实训六 蜂鸣器案例."""
import os
import hashlib
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT

ROOT = r"E:\deepseek harness\学习\实训六_蜂鸣器"
ASSETS = os.path.join(ROOT, "assets")
REC = os.path.join(ROOT, "验证记录")
FW = r"E:\HiSparkStudioSDK\fbb_ws63\src\output\ws63\fwpkg\ws63-liteos-app\ws63-liteos-app_all.fwpkg"
OUT = os.path.join(ROOT, "实训六_蜂鸣器实训报告_向栩言.docx")

DARK = RGBColor(0x14, 0x2B, 0x47)
BLUE = RGBColor(0x1F, 0x4E, 0x79)
GRAY = RGBColor(0x5A, 0x6A, 0x7C)
LIGHT = RGBColor(0x7A, 0x8A, 0x9A)

def set_run_font(run, cn="宋体", en="Times New Roman", size=12, bold=False, color=None, italic=False):
    run.font.name = en
    run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), cn)
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    if color is not None:
        run.font.color.rgb = color

def fmt_para(p, first_indent=True, before=0, after=6, line=1.5, align=WD_ALIGN_PARAGRAPH.JUSTIFY, keep=False):
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line
    pf.alignment = align
    if first_indent:
        pf.first_line_indent = Pt(24)
    pf.keep_with_next = keep

def add_para(doc, text, size=12, bold=False, cn="宋体", en="Times New Roman", color=None,
             first_indent=True, before=0, after=6, line=1.5, align=WD_ALIGN_PARAGRAPH.JUSTIFY):
    p = doc.add_paragraph()
    fmt_para(p, first_indent, before, after, line, align)
    r = p.add_run(text)
    set_run_font(r, cn, en, size, bold, color)
    return p

def add_heading(doc, text, level=1):
    p = doc.add_paragraph()
    sizes = {1: 16, 2: 14, 3: 12}
    before = {1: 14, 2: 10, 3: 8}
    after = {1: 8, 2: 6, 3: 4}
    fmt_para(p, False, before[level], after[level], 1.25, WD_ALIGN_PARAGRAPH.LEFT, True)
    r = p.add_run(text)
    set_run_font(r, "黑体", "Times New Roman", sizes[level], True, DARK)
    return p

def shade_cell(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tcPr.append(shd)

def set_cell_margins(cell, top=60, left=90, bottom=60, right=90):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for tag, val in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        el = OxmlElement("w:" + tag)
        el.set(qn("w:w"), str(val))
        el.set(qn("w:type"), "dxa")
        tcMar.append(el)
    tcPr.append(tcMar)

def set_table_borders(table, val="single", sz=6, color="B7C4D4"):
    tbl = table._tbl
    tblPr = tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement("w:" + edge)
        el.set(qn("w:val"), val)
        el.set(qn("w:sz"), str(sz))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        borders.append(el)
    tblPr.append(borders)

def set_cell_borders(cell, val="single", sz=6, color="C9D3DF", dash=None):
    tcPr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement("w:" + edge)
        el.set(qn("w:val"), dash or val)
        el.set(qn("w:sz"), str(sz))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        borders.append(el)
    tcPr.append(borders)

def add_field(paragraph, instr):
    run = paragraph.add_run()
    fld1 = OxmlElement("w:fldChar"); fld1.set(qn("w:fldCharType"), "begin")
    it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = instr
    fld2 = OxmlElement("w:fldChar"); fld2.set(qn("w:fldCharType"), "separate")
    t = OxmlElement("w:t"); t.text = "1"
    fld3 = OxmlElement("w:fldChar"); fld3.set(qn("w:fldCharType"), "end")
    run._r.append(fld1); run._r.append(it); run._r.append(fld2); run._r.append(t); run._r.append(fld3)
    return run

def add_hyperlink(paragraph, url, text, size=10.5):
    part = paragraph.part
    r_id = part.relate_to(url, RT.HYPERLINK, is_external=True)
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)
    new_run = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    rFonts = OxmlElement("w:rFonts")
    rFonts.set(qn("w:ascii"), "Times New Roman")
    rFonts.set(qn("w:hAnsi"), "Times New Roman")
    rFonts.set(qn("w:eastAsia"), "宋体")
    rPr.append(rFonts)
    color = OxmlElement("w:color"); color.set(qn("w:val"), "1F4E79")
    u = OxmlElement("w:u"); u.set(qn("w:val"), "single")
    sz = OxmlElement("w:sz"); sz.set(qn("w:val"), str(int(size * 2)))
    rPr.append(color); rPr.append(u); rPr.append(sz)
    new_run.append(rPr)
    t = OxmlElement("w:t"); t.text = text; t.set(qn("xml:space"), "preserve")
    new_run.append(t)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)

def add_caption(doc, text):
    p = doc.add_paragraph()
    fmt_para(p, False, 2, 8, 1.0, WD_ALIGN_PARAGRAPH.CENTER)
    r = p.add_run(text)
    set_run_font(r, "宋体", "Times New Roman", 9, False, GRAY)
    return p

def add_image(doc, path, width_cm, caption):
    p = doc.add_paragraph()
    fmt_para(p, False, 4, 2, 1.0, WD_ALIGN_PARAGRAPH.CENTER)
    run = p.add_run()
    run.add_picture(path, width=Cm(width_cm))
    add_caption(doc, caption)

def add_code_block(doc, code, title=None, size=8.5, max_lines=30):
    lines = code.rstrip("\n").split("\n")
    chunks = [lines[i:i + max_lines] for i in range(0, len(lines), max_lines)] or [[""]]
    for idx, chunk in enumerate(chunks):
        cap = title or ""
        if len(chunks) > 1:
            cap = f"{cap}（{idx + 1}/{len(chunks)}）"
        if cap:
            add_caption(doc, cap)
        table = doc.add_table(rows=1, cols=1)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False
        table.allow_autofit = False
        cell = table.cell(0, 0)
        cell.width = Cm(15.8)
        shade_cell(cell, "F4F6F9")
        set_cell_borders(cell, "single", 4, "D5DEE8")
        set_cell_margins(cell, 100, 140, 100, 140)
        p = cell.paragraphs[0]
        pf = p.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.line_spacing = 1.0
        for i, line in enumerate(chunk):
            if i:
                p.add_run().add_break()
            r = p.add_run(line if line else " ")
            set_run_font(r, "宋体", "Consolas", size, False, RGBColor(0x22, 0x2A, 0x35))
        doc.add_paragraph().paragraph_format.space_after = Pt(0)

def add_table(doc, headers, rows, widths, font_size=10.5, header_fill="1F4E79",
              first_col_bold=False, align_center=True):
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    table.allow_autofit = False
    set_table_borders(table, "single", 6, "B7C4D4")
    hdr = table.rows[0]
    for i, h in enumerate(headers):
        cell = hdr.cells[i]
        cell.width = Cm(widths[i])
        shade_cell(cell, header_fill)
        set_cell_margins(cell, 70, 90, 70, 90)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(h)
        set_run_font(r, "黑体", "Times New Roman", font_size, True, RGBColor(0xFF, 0xFF, 0xFF))
    for row in rows:
        tr = table.add_row()
        for i, val in enumerate(row):
            cell = tr.cells[i]
            cell.width = Cm(widths[i])
            set_cell_margins(cell, 60, 90, 60, 90)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if (align_center and i > 0) else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.15
            r = p.add_run(str(val))
            set_run_font(r, "宋体", "Times New Roman", font_size, first_col_bold and i == 0)
    return table

def add_photo_placeholder(doc, text, height_cm=5.2):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    table.allow_autofit = False
    cell = table.cell(0, 0)
    cell.width = Cm(15.5)
    table.rows[0].height = Cm(height_cm)
    shade_cell(cell, "FAFBFC")
    set_cell_borders(cell, "single", 8, "AEBBCB", dash="dashed")
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run(text)
    set_run_font(r, "宋体", "Times New Roman", 10.5, False, LIGHT)
    doc.add_paragraph().paragraph_format.space_after = Pt(0)

def fmt_table_cell_text(cell, text, size=10.5, bold=False, color=None, align=WD_ALIGN_PARAGRAPH.LEFT):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.2
    r = p.add_run(text)
    set_run_font(r, "宋体", "Times New Roman", size, bold, color)

def page_border_bottom(p, color="1F4E79", sz=12):
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), str(sz))
    bottom.set(qn("w:space"), "4")
    bottom.set(qn("w:color"), color)
    pBdr.append(bottom)
    pPr.append(pBdr)

def read(path):
    with open(path, "r", encoding="utf-8") as f:
        return f.read()

def main():
    fw_size = os.path.getsize(FW)
    fw_hash = hashlib.sha256(open(FW, "rb").read()).hexdigest()
    code = read(os.path.join(ROOT, "code", "beep", "pwm_example.c"))

    doc = Document()
    # Normal style
    st = doc.styles["Normal"]
    st.font.name = "Times New Roman"
    st.font.size = Pt(12)
    st._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    st.paragraph_format.line_spacing = 1.5
    st.paragraph_format.space_after = Pt(6)

    sec = doc.sections[0]
    sec.page_width = Cm(21)
    sec.page_height = Cm(29.7)
    sec.top_margin = Cm(2.54)
    sec.bottom_margin = Cm(2.54)
    sec.left_margin = Cm(2.6)
    sec.right_margin = Cm(2.6)
    sec.header_distance = Cm(1.2)
    sec.footer_distance = Cm(1.2)

    # ---------------- Cover ----------------
    for _ in range(1):
        doc.add_paragraph()
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.line_spacing = 1.0
    r = p.add_run("重庆建筑科技职业学院"); set_run_font(r, "黑体", "Times New Roman", 22, True, DARK)
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.line_spacing = 1.0
    r = p.add_run("数字工程学院"); set_run_font(r, "黑体", "Times New Roman", 16, True, BLUE)
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.line_spacing = 1.0
    page_border_bottom(p, "1F4E79", 18)
    p.paragraph_format.space_after = Pt(16)

    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.line_spacing = 1.0
    r = p.add_run("实 训 报 告 书"); set_run_font(r, "黑体", "Times New Roman", 32, True, DARK)
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.line_spacing = 1.0
    r = p.add_run("实训六  蜂鸣器案例"); set_run_font(r, "黑体", "Times New Roman", 20, True, BLUE)
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.line_spacing = 1.0
    r = p.add_run("——基于 PWM 的蜂鸣器驱动实现"); set_run_font(r, "宋体", "Times New Roman", 14, False, GRAY)
    for _ in range(1):
        doc.add_paragraph()

    info = [
        ("课程名称", "星闪技术应用"),
        ("实训项目", "实训六  蜂鸣器案例"),
        ("姓    名", "向栩言"),
        ("学    号", "20250301072"),
        ("班    级", "物联网专贯本2班"),
        ("指导教师", "夏先玉"),
        ("实训日期", "2026 年 11 月 16 日"),
        ("成    绩", ""),
    ]
    table = doc.add_table(rows=0, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    table.allow_autofit = False
    for k, v in info:
        row = table.add_row()
        c0, c1 = row.cells
        c0.width = Cm(4.2); c1.width = Cm(8.8)
        for c in (c0, c1):
            set_cell_borders(c, "nil", 0, "FFFFFF")
            set_cell_margins(c, 24, 80, 24, 80)
        p0 = c0.paragraphs[0]; p0.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p0.paragraph_format.line_spacing = 1.0
        r0 = p0.add_run(k + "："); set_run_font(r0, "黑体", "Times New Roman", 11.5, True, DARK)
        p1 = c1.paragraphs[0]; p1.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p1.paragraph_format.line_spacing = 1.0
        r1 = p1.add_run(v); set_run_font(r1, "宋体", "Times New Roman", 11.5, False)
        if v:
            page_border_bottom(p1, "D5DEE8", 4)
    for _ in range(1):
        doc.add_paragraph()
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.line_spacing = 1.0
    r = p.add_run("2026 年 11 月"); set_run_font(r, "宋体", "Times New Roman", 13, False, GRAY)

    # ---------------- Body section ----------------
    body = doc.add_section(WD_SECTION.NEW_PAGE)
    body.page_width = Cm(21); body.page_height = Cm(29.7)
    body.top_margin = Cm(2.54); body.bottom_margin = Cm(2.54)
    body.left_margin = Cm(2.6); body.right_margin = Cm(2.6)
    body.header_distance = Cm(1.2); body.footer_distance = Cm(1.2)
    body.header.is_linked_to_previous = False
    body.footer.is_linked_to_previous = False
    # page number start at 1
    sectPr = body._sectPr
    pgNum = OxmlElement("w:pgNumType"); pgNum.set(qn("w:start"), "1"); sectPr.append(pgNum)
    # header
    hp = body.header.paragraphs[0]
    hp.paragraph_format.tab_stops.add_tab_stop(Cm(15.8), WD_TAB_ALIGNMENT.RIGHT)
    r = hp.add_run("实训六  蜂鸣器案例"); set_run_font(r, "宋体", "Times New Roman", 9, False, GRAY)
    r = hp.add_run("\t重庆建筑科技职业学院 · 数字工程学院"); set_run_font(r, "宋体", "Times New Roman", 9, False, LIGHT)
    page_border_bottom(hp, "D5DEE8", 4)
    # footer
    fp = body.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = fp.add_run("第 "); set_run_font(r, "宋体", "Times New Roman", 9, False, GRAY)
    r = add_field(fp, " PAGE "); set_run_font(r, "宋体", "Times New Roman", 9, False, GRAY)
    r = fp.add_run(" 页  共 "); set_run_font(r, "宋体", "Times New Roman", 9, False, GRAY)
    r = add_field(fp, " NUMPAGES "); set_run_font(r, "宋体", "Times New Roman", 9, False, GRAY)
    r = fp.add_run(" 页"); set_run_font(r, "宋体", "Times New Roman", 9, False, GRAY)

    # update fields on open
    settings = doc.settings.element
    uf = OxmlElement("w:updateFields"); uf.set(qn("w:val"), "true"); settings.append(uf)

    # ---------------- 目录 ----------------
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("目  录"); set_run_font(r, "黑体", "Times New Roman", 18, True, DARK)
    p.paragraph_format.space_after = Pt(12)
    toc = [
        "一、实训目的与要求",
        "二、实训设备与软件环境",
        "三、实训原理",
        "四、实训步骤（共 8 步）",
        "五、关键代码与解析",
        "六、编译与烧录",
        "七、实训结果",
        "八、问题与解决",
        "九、实训总结与体会",
        "十、教师评语与成绩",
        "附录 A  源码、命令与校验信息",
        "附录 B  交互式实训网页",
        "附录 C  操作检查表",
    ]
    for t in toc:
        p = doc.add_paragraph()
        fmt_para(p, False, 3, 3, 1.5, WD_ALIGN_PARAGRAPH.LEFT)
        r = p.add_run(t); set_run_font(r, "宋体", "Times New Roman", 12, False, DARK)
    doc.add_page_break()

    # ---------------- 1. 目的 ----------------
    add_heading(doc, "一、实训目的与要求", 1)
    add_heading(doc, "1.1 实训目的", 2)
    for t in [
        "1. 理解 PWM（脉冲宽度调制）的基本概念，掌握周期、频率、占空比之间的换算关系，能够根据 32 MHz 计数时钟计算目标频率。",
        "2. 熟悉 HiHope_NearLink_DK_WS63E_V03 开发板与交通灯板的硬件连接，明确交通灯板 BEEP → 底板 A9 → 核心板 GPIO9 的信号链路。",
        "3. 掌握 WS63 SDK 中 PWM 驱动的使用方法：uapi_pin_set_mode、uapi_pwm_init、uapi_pwm_open、uapi_pwm_set_group、uapi_pwm_start、uapi_pwm_close 等。",
        "4. 掌握在 HiSpark Studio / SDK 中新建样例目录、注册 CMake、增加 Kconfig、使用 menuconfig、编译、烧录的完整流程。",
        "5. 能够独立排查“路径不对、配置没生效、串口选错、烧录握手超时”等常见问题，并完成实物验证。",
        "6. 通过交互式网页复盘整个流程，理解“代码提交 → 编译通过 → 烧录完成 → 蜂鸣器鸣响”的先后关系。",
    ]:
        add_para(doc, t, first_indent=True)
    add_heading(doc, "1.2 实训要求", 2)
    for t in [
        "（1）按指导书的 8 个步骤独立完成，不跳过编译和烧录直接上板。",
        "（2）记录实际使用的 SDK 路径、串口、编译结果和烧录结果，保留关键日志。",
        "（3）实物验证时必须听到蜂鸣器鸣响；若未听到，按“问题与解决”逐项排查。",
        "（4）报告中的代码、参数和截图应与实际工程一致，不伪造实验现象。",
    ]:
        add_para(doc, t, first_indent=True)

    # ---------------- 2. 设备 ----------------
    add_heading(doc, "二、实训设备与软件环境", 1)
    add_para(doc, "本次实训使用星闪 WS63E 开发板套件，在 Windows 11 下通过 HiSpark Studio 插件和 SDK 命令行完成编译，使用 CH340K 串口烧录。主要设备与环境如下表所示。", first_indent=True)
    add_table(doc,
              ["类别", "名称 / 版本", "用途与说明"],
              [
                  ["硬件", "HiHope_NearLink_DK_WS63E_V03 核心板", "主控板，兼容文档中的 3863E_V03"],
                  ["硬件", "交通灯板 HiSpark_WiFi_IoT_SSL_VER.A", "板载 BEEP 蜂鸣器、红黄绿 LED、按键"],
                  ["硬件", "底板 HiSpark_WiFi_IoT_EXB_VER.A", "提供 A9 等排针，连接核心板与交通灯板"],
                  ["硬件", "USB Type-C 数据线", "供电 + CH340K 串口通信，必须是数据线"],
                  ["软件", "Windows 11 + VS Code", "开发主机与编辑器"],
                  ["软件", "HiSpark Studio 扩展 26.9.1", "新建工程、Kconfig、编译、程序加载"],
                  ["软件", "SDK fbb_ws63", r"E:\HiSparkStudioSDK\fbb_ws63"],
                  ["软件", "HiSparkStudioToolChain", "riscv32-linux-musl-gcc 7.3.0、CMake、Ninja"],
                  ["串口", "USB-SERIAL CH340K (COM5)", "开发板实际串口；COM3 为蓝牙虚拟串口，不能用"],
              ],
              widths=[1.5, 6.0, 8.3], font_size=10)
    add_caption(doc, "表 2-1  实训设备与软件环境")

    # ---------------- 3. 原理 ----------------
    add_heading(doc, "三、实训原理", 1)
    add_heading(doc, "3.1 PWM 与蜂鸣器", 2)
    add_para(doc, "PWM 是脉冲宽度调制（Pulse Width Modulation）的缩写。它通过快速切换高低电平输出方波，在一个周期内高电平所占的比例称为占空比。对无源蜂鸣器而言，PWM 频率决定音调，占空比影响音色与响度；频率太低或太高都可能听不到。", first_indent=True)
    add_para(doc, "WS63 的 PWM 以 32 MHz 工作时钟计数。设低电平计数为 low_time、高电平计数为 high_time，则：", first_indent=True)
    add_para(doc, "周期计数 = low_time + high_time；频率 f = 32 000 000 / (low_time + high_time)；占空比 = high_time / (low_time + high_time) × 100%。", first_indent=True, bold=True)
    add_para(doc, "样例取 low_time = 20000、high_time = 10000，因此周期为 30000 个计数，f ≈ 1066.7 Hz，占空比 ≈ 33.3%。这个频率落在人耳较敏感的范围，蜂鸣器会发出清晰的“滴——”声。", first_indent=True)
    add_heading(doc, "3.2 硬件连接", 2)
    add_para(doc, "交通灯板上的蜂鸣器丝印为 BEEP，它通过底板左侧的 A9 排针连接到核心板的 GPIO9。软件上需要先把 GPIO9 的引脚复用设置为 PWM 功能（pin mode = 1），再打开 PWM 通道 1、加入组 1，最后启动组输出。", first_indent=True)
    add_para(doc, "硬件链路：核心板 GPIO9 → 底板 A9 → 交通灯板 BEEP → 蜂鸣器。若板子插接不牢、Type-C 只供电不通信、或 GPIO9 没有切到 PWM 模式，都会导致“烧录成功但不响”。", first_indent=True)
    add_heading(doc, "3.3 软件流程", 2)
    add_para(doc, "样例 pwm_example.c 的执行流程如下：", first_indent=True)
    add_table(doc,
              ["顺序", "调用", "作用"],
              [
                  ["1", "uapi_pin_set_mode(9, 1)", "把 GPIO9 复用为 PWM 功能"],
                  ["2", "uapi_pwm_init()", "初始化 PWM 驱动"],
                  ["3", "uapi_pwm_open(1, &cfg)", "按 low/high 参数打开通道 1"],
                  ["4", "uapi_pwm_register_interrupt(1, cb)", "注册周期完成回调（样例用于打印）"],
                  ["5", "uapi_pwm_set_group(1, &ch, 1)", "把通道 1 加入组 1（v151 三参数）"],
                  ["6", "uapi_pwm_start(1)", "启动组 1，开始输出方波"],
                  ["7", "uapi_tcxo_delay_ms(1000)", "保持约 1 秒，蜂鸣器鸣响"],
                  ["8", "uapi_pwm_close(1) / uapi_pwm_deinit()", "停止输出并释放 PWM"],
              ],
              widths=[1.2, 6.8, 7.8], font_size=10)
    add_caption(doc, "表 3-1  PWM 样例调用顺序")

    # ---------------- 4. 步骤 ----------------
    add_heading(doc, "四、实训步骤", 1)
    add_para(doc, "本次实训严格按指导书的 8 个步骤执行。实际 SDK 中 vendor 样例位于 demo\\beep，比 PDF 写的路径多一层 demo；这一点在步骤二特别说明。", first_indent=True)

    add_heading(doc, "4.1 步骤一：新建 beep 样例文件夹", 2)
    add_para(doc, "在 SDK 的 src\\application\\samples\\peripheral 目录上右键，选择“新建文件夹”，命名为 beep。目录名必须全小写，后面 CMake 中的 add_subdirectory_if_exist(beep) 才能找到它。", first_indent=True)
    add_image(doc, os.path.join(ASSETS, "new_folder.png"), 12.5, "图 4-1  在 peripheral 下新建 beep 文件夹（SDK 参考界面）")

    add_heading(doc, "4.2 步骤二：复制 vendor 样例", 2)
    add_para(doc, "把 vendor\\HiHope_NearLink_DK_WS63E_V03\\demo\\beep 中的 pwm_example.c、CMakeLists.txt、README.md 三个文件复制到上一步创建的 beep 目录。注意：PDF 写的是 vendor\\...\\beep，实际 SDK 在 demo\\beep，照 PDF 路径会找不到文件。", first_indent=True)
    add_image(doc, os.path.join(ASSETS, "copy_sample.png"), 12.0, "图 4-2  复制 beep 样例文件（SDK 参考界面）")

    add_heading(doc, "4.3 步骤三：注册 CMake 编译案例", 2)
    add_para(doc, "打开 src\\application\\samples\\peripheral\\CMakeLists.txt，在 set(SOURCES \"${SOURCES}\" PARENT_SCOPE) 之前加入下面的 if 块。它的作用是：当 CONFIG_SAMPLE_SUPPORT_BEEP 被定义时，把 beep 子目录加入编译。", first_indent=True)
    add_code_block(doc, "if(DEFINED CONFIG_SAMPLE_SUPPORT_BEEP)\n    add_subdirectory_if_exist(beep)\nendif()", "代码 4-1  peripheral/CMakeLists.txt 新增片段", 9.5)
    add_image(doc, os.path.join(ASSETS, "cmake.png"), 12.0, "图 4-3  CMakeLists.txt 新增 BEEP 编译案例（SDK 参考界面）")

    add_heading(doc, "4.4 步骤四：新增 Kconfig 配置项", 2)
    add_para(doc, "打开 src\\application\\samples\\peripheral\\Kconfig，在文件末尾追加 BEEP 配置项。指导书截图的 help 文案误写成 “WATCHDOG”，不影响编译；本报告改为 “BEEP”，便于阅读。", first_indent=True)
    add_code_block(doc, "config SAMPLE_SUPPORT_BEEP\n    bool\n    prompt \"Support BEEP Sample.\"\n    default n\n    depends on ENABLE_PERIPHERAL_SAMPLE\n    help\n        This option means support BEEP Sample.", "代码 4-2  peripheral/Kconfig 新增配置项", 9.0)
    add_image(doc, os.path.join(ASSETS, "kconfig.png"), 13.0, "图 4-4  Kconfig 新增 BEEP 配置项（SDK 参考界面）")

    add_heading(doc, "4.5 步骤五：menuconfig 使能 BEEP", 2)
    add_para(doc, "打开 HiSpark Studio 的“系统配置 / Kconfig”，按路径 Application → Enable the Sample of peripheral → Support BEEP Sample. 勾选，然后点击 Save。保存后，ws63_liteos_app.config 中会出现 CONFIG_SAMPLE_SUPPORT_BEEP=y。", first_indent=True)
    add_image(doc, os.path.join(ASSETS, "menuconfig.png"), 8.6, "图 4-5  系统配置入口（SDK 参考界面）")
    add_image(doc, os.path.join(ASSETS, "support_beep.png"), 12.5, "图 4-6  勾选 Support BEEP Sample. 并保存（SDK 参考界面）")
    add_para(doc, "实际配置文件中还保留了 CONFIG_ENABLE_PERIPHERAL_SAMPLE=y。为避免 LED 样例和 BEEP 样例同时运行，本次实验暂时把 CONFIG_SAMPLE_SUPPORT_LED 设为 not set；原配置已备份，实验后可恢复。", first_indent=True)

    add_heading(doc, "4.6 步骤六：编译工程", 2)
    add_para(doc, "在 VS Code 的 HiSpark Studio 插件中点击 build 或 rebuild，或在 SDK 命令行执行 python build.py -c ws63-liteos-app。编译成功后，控制台会显示 SUCCESS Took ... seconds，并生成 ws63-liteos-app_all.fwpkg。", first_indent=True)
    add_code_block(doc, "-- The C compiler identification is GNU 7.3.0\n...\n[288/1435] Building C object application/samples/CMakeFiles/samples.dir/peripheral/beep/pwm_example.c.obj\n...\n######### Build target:ws63_liteos_app success\nws63-liteos-app takes 316.185354 s\npacket success!", "日志 4-1  真实编译关键输出", 8.5)
    add_image(doc, os.path.join(ASSETS, "build_success.jpg"), 14.0, "图 4-7  HiSpark Studio 编译成功（SDK 参考界面）")

    add_heading(doc, "4.7 步骤七：烧录到开发板", 2)
    add_para(doc, "打开“工程配置 → 程序加载”，传输方式选 serial，端口选 CH340K 对应的 COM5，波特率 115200，点击“程序加载”。出现 Connecting, please reset device... 时按一下开发板 RST（或重新上电），等待烧录完成。", first_indent=True)
    add_para(doc, "本次实际使用的烧录命令为：BurnToolCmd.exe --burn -n ws63 -m serial COM5 --baudRate 115200 -f ws63-liteos-app_all.fwpkg。", first_indent=True)
    add_image(doc, os.path.join(ASSETS, "project_config.jpg"), 14.0, "图 4-8  程序加载配置（SDK 参考界面）")
    add_image(doc, os.path.join(ASSETS, "flash_connecting.png"), 14.0, "图 4-9  Connecting 提示与复位时机（SDK 参考界面）")

    add_heading(doc, "4.8 步骤八：蜂鸣器鸣响", 2)
    add_para(doc, "烧录完成后设备自动复位。交通灯板上的蜂鸣器发出约 1 秒的鸣响，频率约 1066.7 Hz、占空比约 33.3%。本次实物验证已由本人确认听到声音。", first_indent=True)
    add_image(doc, os.path.join(ASSETS, "hardware.jpg"), 13.5, "图 4-10  核心板、底板与交通灯板连接（SDK 参考图）")

    # ---------------- 5. 代码 ----------------
    add_heading(doc, "五、关键代码与解析", 1)
    add_heading(doc, "5.1 pwm_example.c 完整代码", 2)
    add_para(doc, "下面是从 vendor\\...\\demo\\beep 复制并实际编译进固件的 pwm_example.c。代码保留了原样例的版权头、任务创建和回调打印，只在工程配置上做了使能。", first_indent=True)
    add_code_block(doc, code, "代码 5-1  pwm_example.c（共 75 行）", 8.0, 30)

    add_heading(doc, "5.2 关键参数与计算", 2)
    add_table(doc,
              ["参数", "取值", "说明"],
              [
                  ["PWM_CHANNEL", "1", "使用的 PWM 通道"],
                  ["PWM_GROUP_ID", "1", "通道分组号"],
                  ["CONFIG_PWM_PIN", "9", "GPIO9 / 底板 A9"],
                  ["CONFIG_PWM_PIN_MODE", "1", "GPIO9 复用为 PWM"],
                  ["low_time", "20000", "低电平 32 MHz 计数"],
                  ["high_time", "10000", "高电平 32 MHz 计数"],
                  ["周期计数", "30000", "20000 + 10000"],
                  ["频率", "1066.7 Hz", "32 000 000 / 30000"],
                  ["占空比", "33.3%", "10000 / 30000"],
                  ["鸣响时间", "约 1 s", "uapi_tcxo_delay_ms(1000)"],
              ],
              widths=[3.8, 3.8, 8.2], font_size=10, first_col_bold=True)
    add_caption(doc, "表 5-1  PWM 关键参数")

    add_heading(doc, "5.3 API 版本差异说明", 2)
    add_para(doc, "指导书 2.5 节给出的 uapi_pwm_set_group 定义是旧版两参数 (group, uint16_t channel_id)，而本机 SDK 的 pwm.h（v151）是三参数 (group, const uint8_t *channel_set, uint32_t channel_set_len)，样例代码也使用三参数：uapi_pwm_set_group(PWM_GROUP_ID, &channel_id, 1)。如果照 PDF 写两参数，编译会报参数个数不匹配。本报告以实际 SDK 和样例代码为准。", first_indent=True)

    # ---------------- 6. 编译烧录 ----------------
    add_heading(doc, "六、编译与烧录", 1)
    add_heading(doc, "6.1 编译过程与问题", 2)
    add_para(doc, "首次执行 python build.py -c ws63-liteos-app 时，Kconfig 解析报错：application/samples/peripheral/Kconfig:466: 'application/samples/peripheral/buttondemo/Kconfig' not found。原因是此前实验新增的 buttondemo 目录缺少 Kconfig 文件，而 peripheral/Kconfig 中又 source 了它。补一个只含注释的 buttondemo/Kconfig 后，重新编译通过。", first_indent=True)
    add_para(doc, f"第二次编译耗时 316.185354 秒，日志中出现 pwm_example.c.obj，说明 BEEP 样例确实参与编译。生成的固件为 ws63-liteos-app_all.fwpkg，大小 {fw_size:,} 字节（约 {fw_size/1024/1024:.2f} MB），SHA-256 为 {fw_hash}。", first_indent=True)
    add_para(doc, "编译日志中还有几处 'size' is not recognized 的提示，来自 codesize_statistic.py 调用的 size 工具不在 PATH；它只影响代码体积统计，不影响固件生成和烧录。", first_indent=True)

    add_heading(doc, "6.2 烧录过程与串口排查", 2)
    add_para(doc, "最初按口头信息使用 COM3 烧录，BurnToolCmd 连续报 Ack receiving timeout / Shakehand failed。查阅注册表 HKLM\\HARDWARE\\DEVICEMAP\\SERIALCOMM 后发现：COM3/4/6/7 都是 BthModem（蓝牙虚拟串口），只有 COM5 是 \\Device\\Serial2；进一步在 USB 设备中确认 USB-SERIAL CH340K (COM5)。改到 COM5 后，按提示复位，烧录一次成功。", first_indent=True)
    add_code_block(doc, "Please reset the device, the waiting time is about 30 seconds.\n...\ndownloaded: ws63-liteos-app-sign.bin successfully.\nBurn ...ws63-liteos-app_all.fwpkg successfully.\nBurned successfully.", "日志 6-1  COM5 烧录成功关键输出", 8.5)
    add_para(doc, "烧录完成时间约为 2026-11-16 15:12:24，烧录后设备自动复位。", first_indent=True)

    # ---------------- 7. 结果 ----------------
    add_heading(doc, "七、实训结果", 1)
    add_heading(doc, "7.1 实物现象", 2)
    add_para(doc, "烧录完成后，按开发板 RST 复位，交通灯板上的蜂鸣器发出约 1 秒的连续鸣响。本人已在现场确认听到声音，说明 GPIO9 已成功复用为 PWM、通道/组配置正确、固件已正常运行。", first_indent=True)
    add_table(doc,
              ["项目", "设计/预期值", "实测结果"],
              [
                  ["PWM 通道 / 组", "通道 1 / 组 1", "与代码一致"],
                  ["GPIO", "GPIO9（底板 A9）", "接线正确"],
                  ["频率", "约 1066.7 Hz", "听到连续蜂鸣，音调稳定"],
                  ["占空比", "约 33.3%", "声音清晰，无明显异常"],
                  ["鸣响时间", "约 1 s", "约 1 秒后停止"],
                  ["烧录串口", "COM5（CH340K）", "Burned successfully"],
              ],
              widths=[3.8, 5.0, 7.0], font_size=10, first_col_bold=True)
    add_caption(doc, "表 7-1  实物验证结果")

    add_heading(doc, "7.2 交互式网页仿真", 2)
    add_para(doc, "为复盘整个流程，本次还制作了单文件交互式网页。网页严格按 8 个步骤组织，只有代码提交、模拟编译、烧录模拟全部通过后，才解锁虚拟蜂鸣器；虚拟蜂鸣器用 Web Audio 合成与板载蜂鸣器同频的脉冲波，从电脑扬声器/耳机发声。", first_indent=True)
    add_image(doc, os.path.join(ASSETS, "web_preview8.jpg"), 15.5, "图 7-1  交互式网页第 8 步：虚拟蜂鸣器与 PWM 示波器")
    add_para(doc, "需要说明：浏览器没有权限直接驱动主板 PC 蜂鸣器，因此网页里的声音是“同频合成”的虚拟蜂鸣器；实物声音仍由开发板 PWM 驱动。网页中的编译是教学模拟器（规则校验 + 拟真日志），真实编译结果以第六章为准。", first_indent=True)

    add_heading(doc, "7.3 实物照片与记录", 2)
    add_para(doc, "以下位置请补拍/粘贴本人实物照片，作为实训过程与结果的佐证。", first_indent=True)
    add_photo_placeholder(doc, "【请粘贴照片 1】开发板 + 交通灯板连接、Type-C 数据线接 PC", 4.8)
    add_photo_placeholder(doc, "【请粘贴照片 2】HiSpark Studio 编译成功界面（SUCCESS）", 4.8)
    add_photo_placeholder(doc, "【请粘贴照片 3】烧录完成 / 蜂鸣器鸣响（可拍视频后截帧）", 4.8)
    add_para(doc, "实测记录：日期 2026 年 11 月 16 日；串口 COM5；编译结果 SUCCESS；烧录结果 Burned successfully；现象：蜂鸣器鸣响约 1 秒。学生签名：____________", first_indent=True)

    # ---------------- 8. 问题 ----------------
    add_heading(doc, "八、问题与解决", 1)
    add_para(doc, "实训过程中遇到并解决了以下问题，记录如下，供以后复盘。", first_indent=True)
    problems = [
        ("PDF 路径少了 demo", "指导书写 vendor\\...\\beep，实际 SDK 为 vendor\\...\\demo\\beep", "以实际目录为准，复制 demo\\beep 下的 3 个文件"),
        ("Kconfig 的 help 文案写成 WATCHDOG", "指导书截图复制粘贴笔误", "改为 BEEP；不影响编译，只影响菜单说明"),
        ("首次编译 Kconfig 报错，找不到 buttondemo/Kconfig", "之前实验新增的 buttondemo 目录缺少 Kconfig，而 peripheral/Kconfig source 了它", "补一个只含注释的 buttondemo/Kconfig，重新编译通过"),
        ("PDF 的 uapi_pwm_set_group 是两参数，样例是三参数", "指导书对应旧版 API，本机 SDK 为 v151", "以 pwm.h 和样例为准，使用 (group, &channel_id, 1)"),
        ("COM3 烧录握手超时", "COM3 是蓝牙虚拟串口，不是开发板", "查注册表/设备管理器，确认 CH340K 为 COM5"),
        ("BurnToolCmd 提示 Please reset device 后无响应", "没有在 30 秒窗口内复位，或串口选错", "改用 COM5，在提示后立即按 RST/重新上电，烧录成功"),
        ("LED 样例与 BEEP 样例同时运行", "之前实验使能了 CONFIG_SAMPLE_SUPPORT_LED", "临时关闭 LED，仅保留 BEEP；原配置已备份"),
        ("回调用 osal_printk 每周期打印", "PWM 周期约 1 kHz，串口打印量大", "本次未影响鸣响；若串口刷屏，可注释回调打印或降低频率"),
        ("cycles=1 与 repeat=true 看起来矛盾", "cycles 表示重复周期数，repeat 表示是否连续输出；样例按连续输出使用", "以实际鸣响为准，延时 1 秒后 close"),
    ]
    add_table(doc, ["问题", "原因", "解决方法"], problems, widths=[4.6, 5.6, 5.6], font_size=9.5)
    add_caption(doc, "表 8-1  问题与解决记录")

    # ---------------- 9. 总结 ----------------
    add_heading(doc, "九、实训总结与体会", 1)
    for t in [
        "本次实训从“新建一个 beep 文件夹”开始，完整走完了样例复制、CMake 注册、Kconfig 配置、menuconfig 使能、编译、烧录、实物鸣响的全过程。看似只是让蜂鸣器响一声，实际把 SDK 工程结构、构建系统、配置系统和硬件引脚复用串成了一条完整的链路。",
        "最大的收获是理解了“配置—编译—运行”的因果关系：Kconfig 里的 SAMPLE_SUPPORT_BEEP 决定 menuconfig 是否出现选项，menuconfig 保存的 .config 决定 CMake 的 if 是否成立，CMake 决定 pwm_example.c 是否被编译，编译出的固件再通过串口烧录到板子上。任何一环缺失，最后都表现为“没声音”。",
        "排查过程也让我印象很深。最初按口头信息用 COM3 烧录，连续握手超时；后来查注册表才发现 COM3 是蓝牙虚拟串口，真正的 CH340K 是 COM5。这个细节说明：做硬件实验不能只看“端口号”，要确认端口背后的设备。另一个问题是 PDF 的样例路径少了 demo、API 签名也是旧版；如果完全照抄文档，会在找文件和编译两处卡住。以实际 SDK 和样例代码为准，才是可靠的工程习惯。",
        "交互式网页让我把流程“讲清楚”了一遍：每一步都要有明确的通过条件，代码改了就要重新提交、重新编译、重新烧录，不能跳过。虚拟蜂鸣器用 Web Audio 合成，虽然不能代替实物，但能帮助理解频率、占空比和波形的关系。",
        "后续可以继续拓展：用按键控制蜂鸣器开关、播放简单旋律、用 PWM 调音做电子琴、把蜂鸣器接入交通灯控制逻辑。也可以把本次的 8 步流程做成可复用的实验模板，方便以后做 LED、按键、传感器等实训。",
    ]:
        add_para(doc, t, first_indent=True)

    # ---------------- 10. 教师 ----------------
    add_heading(doc, "十、教师评语与成绩", 1)
    table = doc.add_table(rows=3, cols=2)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    table.allow_autofit = False
    set_table_borders(table, "single", 6, "B7C4D4")
    table.rows[0].cells[0].width = Cm(3.0); table.rows[0].cells[1].width = Cm(12.8)
    table.rows[0].height = Cm(6.5)
    table.rows[1].height = Cm(1.6)
    table.rows[2].height = Cm(1.8)
    fmt_table_cell_text(table.rows[0].cells[0], "教师评语", 11, True, align=WD_ALIGN_PARAGRAPH.CENTER)
    fmt_table_cell_text(table.rows[0].cells[1], "", 11)
    fmt_table_cell_text(table.rows[1].cells[0], "成    绩", 11, True, align=WD_ALIGN_PARAGRAPH.CENTER)
    fmt_table_cell_text(table.rows[1].cells[1], "", 11)
    fmt_table_cell_text(table.rows[2].cells[0], "教师签名", 11, True, align=WD_ALIGN_PARAGRAPH.CENTER)
    fmt_table_cell_text(table.rows[2].cells[1], "                                  日期：      年    月    日", 11)
    add_caption(doc, "表 10-1  教师评语与成绩")

    # ---------------- Appendix A ----------------
    add_heading(doc, "附录 A  源码、命令与校验信息", 1)
    add_para(doc, "实际工程中的关键文件与路径如下：", first_indent=True)
    add_table(doc,
              ["文件 / 项目", "路径或说明"],
              [
                  ["样例源码", r"src\application\samples\peripheral\beep\pwm_example.c"],
                  ["样例 CMake", r"src\application\samples\peripheral\beep\CMakeLists.txt"],
                  ["构建注册", r"src\application\samples\peripheral\CMakeLists.txt"],
                  ["Kconfig", r"src\application\samples\peripheral\Kconfig"],
                  ["配置文件", r"src\build\config\target_config\ws63\menuconfig\acore\ws63_liteos_app.config"],
                  ["固件", r"src\output\ws63\fwpkg\ws63-liteos-app\ws63-liteos-app_all.fwpkg"],
                  ["固件大小", f"{fw_size:,} 字节（约 {fw_size/1024/1024:.2f} MB）"],
                  ["SHA-256", fw_hash],
              ],
              widths=[3.4, 12.4], font_size=9.5, first_col_bold=True)
    add_caption(doc, "表 A-1  工程文件与固件信息")
    add_para(doc, "编译命令（在 SDK 的 src 目录执行，工具链已在 PATH 中）：", first_indent=True)
    add_code_block(doc, "set PATH=<工具链>\\python;<工具链>\\python\\Scripts;<工具链>\\Windows\\ninja;<工具链>\\Windows\\cc_riscv32_win_env;%PATH%\ncd /d E:\\HiSparkStudioSDK\\fbb_ws63\\src\npython build.py -c ws63-liteos-app", "代码 A-1  编译命令", 9.0)
    add_para(doc, "烧录命令：", first_indent=True)
    add_code_block(doc, "BurnToolCmd.exe --burn -n ws63 -m serial COM5 --baudRate 115200 ^\n  -f E:\\HiSparkStudioSDK\\fbb_ws63\\src\\output\\ws63\\fwpkg\\ws63-liteos-app\\ws63-liteos-app_all.fwpkg", "代码 A-2  烧录命令", 9.0)

    # ---------------- Appendix B ----------------
    add_heading(doc, "附录 B  交互式实训网页", 1)
    add_para(doc, "网页为单文件 HTML，双击可离线打开，也可通过下面的公网地址访问。它按 8 个步骤还原实训流程，包含代码编辑/提交、模拟编译、烧录模拟和虚拟蜂鸣器。", first_indent=True)
    p = doc.add_paragraph(); fmt_para(p, False, 4, 4, 1.5, WD_ALIGN_PARAGRAPH.LEFT)
    r = p.add_run("公网地址："); set_run_font(r, "宋体", "Times New Roman", 11, True)
    add_hyperlink(p, "https://dxuy69.gicp.fun", "https://dxuy69.gicp.fun", 11)
    p = doc.add_paragraph(); fmt_para(p, False, 0, 4, 1.5, WD_ALIGN_PARAGRAPH.LEFT)
    r = p.add_run("资源 ID：1794813761412362123    发布时间：2026-11-16 15:24    状态：已认领、公开访问")
    set_run_font(r, "宋体", "Times New Roman", 10, False, GRAY)
    add_image(doc, os.path.join(ASSETS, "web_preview0.jpg"), 15.5, "图 B-1  交互式网页首页")
    add_image(doc, os.path.join(ASSETS, "web_preview6.jpg"), 15.5, "图 B-2  交互式网页代码工作区与模拟编译")
    add_para(doc, "使用说明：按左侧任务地图从 0 走到 8；步骤 3/4/6 需要写代码并点“提交代码”；步骤 6 编译通过后才能烧录；烧录完成后第 8 步的“鸣响 1 秒”才会解锁。网页中的编译为教学模拟器，真实编译以第六章为准；虚拟蜂鸣器用 Web Audio 从扬声器发声。", first_indent=True)

    # ---------------- Appendix C ----------------
    add_heading(doc, "附录 C  操作检查表", 1)
    checks = [
        "确认 SDK 路径为 E:\\HiSparkStudioSDK\\fbb_ws63，vendor 样例在 demo\\beep",
        "在 peripheral 下创建 beep 文件夹，并复制 3 个样例文件",
        "在 peripheral/CMakeLists.txt 的 set(SOURCES ...) 前加入 BEEP if 块",
        "在 peripheral/Kconfig 末尾加入 SAMPLE_SUPPORT_BEEP 配置项",
        "menuconfig 勾选 Support BEEP Sample. 并 Save",
        "编译成功，日志出现 pwm_example.c.obj 和 SUCCESS",
        "确认 CH340K 串口为 COM5（COM3 是蓝牙）",
        "烧录时在 Connecting 提示后按 RST，烧录 100% 成功",
        "复位后听到蜂鸣器约 1 秒鸣响",
        "补拍 3 张实物照片并填入报告",
    ]
    for c in checks:
        p = doc.add_paragraph()
        fmt_para(p, False, 2, 2, 1.5, WD_ALIGN_PARAGRAPH.LEFT)
        r = p.add_run("☐  "); set_run_font(r, "宋体", "Times New Roman", 12, False, BLUE)
        r = p.add_run(c); set_run_font(r, "宋体", "Times New Roman", 11.5)

    # save
    doc.save(OUT)
    print("saved", OUT, os.path.getsize(OUT))

    # verify
    d2 = Document(OUT)
    text = "\n".join(p.text for p in d2.paragraphs)
    print("paragraphs", len(d2.paragraphs), "tables", len(d2.tables), "images", len(d2.inline_shapes))
    for key in ["实训六  蜂鸣器案例", "向栩言", "20250301072", "夏先玉", "1066.7", "COM5", "https://dxuy69.gicp.fun", "Burned successfully"]:
        print(key, key in text or any(key in c.text for t in d2.tables for r in t.rows for c in r.cells))

if __name__ == "__main__":
    main()
