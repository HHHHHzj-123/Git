from __future__ import annotations

import base64
import hashlib
import json
import zipfile
from pathlib import Path

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

from finance_cases_data import CASES


ROOT = Path(r"C:\Users\HZJ\Desktop\Git\work")
BLUE = "2E74B5"
DARK_BLUE = "1F4D78"
INK = "203040"
MUTED = "64748B"
LIGHT_BLUE = "E8EEF5"
LIGHT_GRAY = "F4F6F9"
GOLD = "A66B14"
GREEN = "2F6B4F"
LATIN_FONT = "Calibri"
CJK_FONT = "Microsoft YaHei"


def set_run(run, size=11, bold=False, color=INK, italic=False):
    run.font.name = LATIN_FONT
    rpr = run._element.get_or_add_rPr()
    rpr.rFonts.set(qn("w:eastAsia"), CJK_FONT)
    rpr.rFonts.set(qn("w:ascii"), LATIN_FONT)
    rpr.rFonts.set(qn("w:hAnsi"), LATIN_FONT)
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = RGBColor.from_string(color)
    return run


def shade_cell(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_cell_margins(cell, top=80, start=120, bottom=80, end=120):
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for tag, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{tag}"))
        if node is None:
            node = OxmlElement(f"w:{tag}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def set_table_geometry(table, widths_inches):
    assert round(sum(widths_inches), 3) == 6.5
    table.autofit = False
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    tbl_pr = table._tbl.tblPr
    tbl_w = tbl_pr.find(qn("w:tblW"))
    if tbl_w is None:
        tbl_w = OxmlElement("w:tblW")
        tbl_pr.append(tbl_w)
    tbl_w.set(qn("w:w"), "9360")
    tbl_w.set(qn("w:type"), "dxa")
    tbl_ind = tbl_pr.find(qn("w:tblInd"))
    if tbl_ind is None:
        tbl_ind = OxmlElement("w:tblInd")
        tbl_pr.append(tbl_ind)
    tbl_ind.set(qn("w:w"), "120")
    tbl_ind.set(qn("w:type"), "dxa")
    dxa = [int(round(value * 1440)) for value in widths_inches]
    grid = table._tbl.tblGrid
    for child in list(grid):
        grid.remove(child)
    for width in dxa:
        col = OxmlElement("w:gridCol")
        col.set(qn("w:w"), str(width))
        grid.append(col)
    for row in table.rows:
        for index, cell in enumerate(row.cells):
            cell.width = Inches(widths_inches[index])
            tc_w = cell._tc.get_or_add_tcPr().find(qn("w:tcW"))
            if tc_w is None:
                tc_w = OxmlElement("w:tcW")
                cell._tc.get_or_add_tcPr().append(tc_w)
            tc_w.set(qn("w:w"), str(dxa[index]))
            tc_w.set(qn("w:type"), "dxa")
            set_cell_margins(cell)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def repeat_header(row):
    tr_pr = row._tr.get_or_add_trPr()
    node = OxmlElement("w:tblHeader")
    node.set(qn("w:val"), "true")
    tr_pr.append(node)


def add_page_number(paragraph):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    set_run(paragraph.add_run("第 "), 9, color=MUTED)
    run = paragraph.add_run()
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = "PAGE"
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run._r.extend([fld_begin, instr, fld_end])
    set_run(run, 9, color=MUTED)
    set_run(paragraph.add_run(" 页"), 9, color=MUTED)


def add_heading(doc, text, level):
    p = doc.add_paragraph(style=f"Heading {level}")
    p.paragraph_format.keep_with_next = True
    set_run(p.add_run(text), {1: 16, 2: 13, 3: 12}[level], True, {1: BLUE, 2: BLUE, 3: DARK_BLUE}[level])
    return p


def add_callout(doc, label, text, fill=LIGHT_BLUE, accent=BLUE):
    p = doc.add_paragraph()
    p_pr = p._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    p_pr.append(shd)
    borders = OxmlElement("w:pBdr")
    left = OxmlElement("w:left")
    left.set(qn("w:val"), "single")
    left.set(qn("w:sz"), "18")
    left.set(qn("w:space"), "8")
    left.set(qn("w:color"), accent)
    borders.append(left)
    p_pr.append(borders)
    p.paragraph_format.left_indent = Inches(0.12)
    p.paragraph_format.right_indent = Inches(0.06)
    p.paragraph_format.space_before = Pt(5)
    p.paragraph_format.space_after = Pt(7)
    p.paragraph_format.line_spacing = 1.2
    set_run(p.add_run(label + "  "), 10.5, True, accent)
    set_run(p.add_run(text), 10.5, color=INK)
    return p


def create_numbering(doc):
    numbering = doc.part.numbering_part.element

    def add_abstract(abstract_id, kind):
        abstract = OxmlElement("w:abstractNum")
        abstract.set(qn("w:abstractNumId"), str(abstract_id))
        multi = OxmlElement("w:multiLevelType")
        multi.set(qn("w:val"), "singleLevel")
        abstract.append(multi)
        lvl = OxmlElement("w:lvl")
        lvl.set(qn("w:ilvl"), "0")
        start = OxmlElement("w:start")
        start.set(qn("w:val"), "1")
        lvl.append(start)
        num_fmt = OxmlElement("w:numFmt")
        num_fmt.set(qn("w:val"), "bullet" if kind == "bullet" else "decimal")
        lvl.append(num_fmt)
        lvl_text = OxmlElement("w:lvlText")
        lvl_text.set(qn("w:val"), "•" if kind == "bullet" else "%1.")
        lvl.append(lvl_text)
        suff = OxmlElement("w:suff")
        suff.set(qn("w:val"), "tab")
        lvl.append(suff)
        ppr = OxmlElement("w:pPr")
        tabs = OxmlElement("w:tabs")
        tab = OxmlElement("w:tab")
        tab.set(qn("w:val"), "num")
        tab.set(qn("w:pos"), "540")
        tabs.append(tab)
        ppr.append(tabs)
        ind = OxmlElement("w:ind")
        ind.set(qn("w:left"), "540")
        ind.set(qn("w:hanging"), "270")
        ppr.append(ind)
        spacing = OxmlElement("w:spacing")
        spacing.set(qn("w:after"), "80")
        spacing.set(qn("w:line"), "300")
        spacing.set(qn("w:lineRule"), "auto")
        ppr.append(spacing)
        lvl.append(ppr)
        rpr = OxmlElement("w:rPr")
        fonts = OxmlElement("w:rFonts")
        fonts.set(qn("w:ascii"), LATIN_FONT)
        fonts.set(qn("w:hAnsi"), LATIN_FONT)
        fonts.set(qn("w:eastAsia"), CJK_FONT)
        rpr.append(fonts)
        lvl.append(rpr)
        abstract.append(lvl)
        numbering.append(abstract)

    def add_num(num_id, abstract_id):
        num = OxmlElement("w:num")
        num.set(qn("w:numId"), str(num_id))
        abstract_ref = OxmlElement("w:abstractNumId")
        abstract_ref.set(qn("w:val"), str(abstract_id))
        num.append(abstract_ref)
        numbering.append(num)

    add_abstract(90, "bullet")
    add_abstract(91, "decimal")
    add_num(90, 90)
    add_num(91, 91)
    return 90, 91


def add_list_item(doc, text, num_id=90, compact=False, dense=False, super_dense=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(0.8 if super_dense else (1.5 if dense else (3 if compact else 4)))
    p.paragraph_format.line_spacing = 1.08 if super_dense else (1.15 if dense else (1.2 if compact else 1.25))
    ppr = p._p.get_or_add_pPr()
    num_pr = OxmlElement("w:numPr")
    ilvl = OxmlElement("w:ilvl")
    ilvl.set(qn("w:val"), "0")
    num = OxmlElement("w:numId")
    num.set(qn("w:val"), str(num_id))
    num_pr.extend([ilvl, num])
    ppr.append(num_pr)
    set_run(p.add_run(text), 9.7 if super_dense else (10.2 if dense else (10.5 if compact else 11)), color=INK)
    return p


def setup_doc(case):
    doc = Document()
    section = doc.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    super_dense = case.get("super_dense_layout", False)
    section.top_margin = Inches(0.72 if super_dense else (0.88 if case.get("dense_layout", False) else 1))
    section.bottom_margin = Inches(0.72 if super_dense else (0.88 if case.get("dense_layout", False) else 1))
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)
    section.header_distance = Inches(0.492)
    section.footer_distance = Inches(0.492)

    normal = doc.styles["Normal"]
    normal.font.name = LATIN_FONT
    normal._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), CJK_FONT)
    normal._element.get_or_add_rPr().rFonts.set(qn("w:ascii"), LATIN_FONT)
    normal._element.get_or_add_rPr().rFonts.set(qn("w:hAnsi"), LATIN_FONT)
    normal.font.size = Pt(9.8 if super_dense else (10.5 if case.get("dense_layout", False) else 11))
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(3 if super_dense else (4 if case.get("dense_layout", False) else 6))
    normal.paragraph_format.line_spacing = 1.12 if super_dense else (1.2 if case.get("dense_layout", False) else 1.25)
    if super_dense:
        heading_specs = (
            ("Heading 1", 14.7, BLUE, 8, 4),
            ("Heading 2", 11.8, BLUE, 5, 2),
            ("Heading 3", 10.8, DARK_BLUE, 4, 1),
        )
    elif case.get("dense_layout", False):
        heading_specs = (
            ("Heading 1", 15.2, BLUE, 10, 5),
            ("Heading 2", 12.2, BLUE, 7, 3),
            ("Heading 3", 11.2, DARK_BLUE, 5, 2),
        )
    elif case.get("ultra_compact_layout", False):
        heading_specs = (
            ("Heading 1", 15.5, BLUE, 13, 7),
            ("Heading 2", 12.5, BLUE, 9, 4),
            ("Heading 3", 11.5, DARK_BLUE, 7, 3),
        )
    else:
        heading_specs = (
            ("Heading 1", 16, BLUE, 18, 10),
            ("Heading 2", 13, BLUE, 14, 7),
            ("Heading 3", 12, DARK_BLUE, 10, 5),
        )
    for name, size, color, before, after in heading_specs:
        style = doc.styles[name]
        style.font.name = LATIN_FONT
        rpr = style._element.get_or_add_rPr()
        rpr.rFonts.set(qn("w:eastAsia"), CJK_FONT)
        rpr.rFonts.set(qn("w:ascii"), LATIN_FONT)
        rpr.rFonts.set(qn("w:hAnsi"), LATIN_FONT)
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = RGBColor.from_string(color)
        style.paragraph_format.space_before = Pt(before)
        style.paragraph_format.space_after = Pt(after)
        style.paragraph_format.keep_with_next = True

    header_p = section.header.paragraphs[0]
    set_run(header_p.add_run(f"齐昊老师 · 学员实战案例 第 {case['episode']} 集"), 9, True, MUTED)
    add_page_number(section.footer.paragraphs[0])
    create_numbering(doc)
    return doc


def build_docx(case):
    doc = setup_doc(case)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(72)
    p.paragraph_format.space_after = Pt(18)
    set_run(p.add_run("财务管理实战案例"), 11, True, GOLD)

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.paragraph_format.space_after = Pt(10)
    set_run(
        title.add_run(f"第 {case['episode']} 集｜{case['title']}"),
        case.get("cover_title_size", 27),
        True,
        DARK_BLUE,
    )

    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle.paragraph_format.space_after = Pt(26)
    set_run(subtitle.add_run(case["subtitle"]), 14, False, BLUE)

    for label, value in (
        ("主讲账号", "齐昊老师"),
        ("发布时间", case["date"]),
        ("视频时长", case["duration"]),
        ("视频编号", case["video_id"]),
    ):
        row = doc.add_paragraph()
        row.alignment = WD_ALIGN_PARAGRAPH.CENTER
        row.paragraph_format.space_after = Pt(3)
        set_run(row.add_run(label + "："), 10.5, True, MUTED)
        set_run(row.add_run(value), 10.5, color=INK)

    doc.add_paragraph().paragraph_format.space_after = Pt(16)
    add_callout(doc, "核心结论", case["core"], "EAF3EE", GREEN)
    source_note = case.get(
        "source_note",
        "依据公开视频静默转写并结构化改写；不是逐字稿。视频中的金额、税率与政策口径保留其发表时语境，实际工作应按最新法规和具体事实复核。",
    )
    add_callout(doc, "整理口径", source_note, LIGHT_GRAY, MUTED)
    doc.add_page_break()

    add_heading(doc, "一页总览", 1)
    table = doc.add_table(rows=1, cols=2)
    table.style = "Table Grid"
    for index, text in enumerate(("节点", "核心内容")):
        shade_cell(table.rows[0].cells[index], LIGHT_BLUE)
        para = table.rows[0].cells[index].paragraphs[0]
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_run(para.add_run(text), 10.5, True, DARK_BLUE)
    for label, value in case["overview"]:
        cells = table.add_row().cells
        set_run(cells[0].paragraphs[0].add_run(label), 10.2, True, DARK_BLUE)
        set_run(cells[1].paragraphs[0].add_run(value), 10.2, color=INK)
    repeat_header(table.rows[0])
    set_table_geometry(table, [1.6, 4.9])

    add_callout(doc, "逻辑主线", case.get("overview_logic", "识别商业事实 → 找到管理根因 → 把财务控制嵌入业务流程 → 用数据支持经营决策"), "FFF7E8", GOLD)

    for section in case["sections"]:
        add_heading(doc, section["title"], 1)
        for sub_title, bullets in section["subs"]:
            add_heading(doc, sub_title, 2)
            for bullet in bullets:
                add_list_item(
                    doc,
                    bullet,
                    90,
                    case.get("compact_layout", False),
                    case.get("dense_layout", False),
                    case.get("super_dense_layout", False),
                )

    # The cover already identifies the account, episode and video ID. The full
    # source URL remains in the companion XMind/Markdown, avoiding a duplicate
    # source block that can create a nearly blank trailing Word page.
    doc.core_properties.title = f"齐昊老师学员实战案例第{case['episode']}集：{case['title']}"
    doc.core_properties.subject = "财务管理与财务总监实战案例整理"
    doc.core_properties.author = "Codex"
    doc.core_properties.keywords = "财务管理, 财务经理, 财务总监, 业财融合, 实战案例"
    out = ROOT / f"齐昊老师_学员实战案例_第{case['episode']}集_文字思维导图.docx"
    doc.save(out)
    return out


def stable_id(value):
    return hashlib.sha1(value.encode("utf-8")).hexdigest()[:26]


def topic(title, path, children=None, root=False):
    node = {"id": stable_id(path), "class": "topic", "title": title}
    if root:
        node["structureClass"] = "org.xmind.ui.map.unbalanced"
    if children:
        node["children"] = {"attached": children}
    return node


def walk(node):
    yield node
    for child in node.get("children", {}).get("attached", []):
        yield from walk(child)


def build_xmind(case):
    ep = case["episode"]
    overview_children = []
    for label, value in case["overview"]:
        leaves = [topic(part.strip(), f"ep{ep}/overview/{label}/{part.strip()}") for part in value.split("→")]
        overview_children.append(topic(label, f"ep{ep}/overview/{label}", leaves))
    overview_children.append(topic("核心结论", f"ep{ep}/overview/core", [topic(case["core"], f"ep{ep}/overview/core/text")]))

    detail_children = []
    for section in case["sections"]:
        sub_nodes = []
        for sub_title, bullets in section["subs"]:
            leaves = [topic(item, f"ep{ep}/{section['title']}/{sub_title}/{item}") for item in bullets]
            sub_nodes.append(topic(sub_title, f"ep{ep}/{section['title']}/{sub_title}", leaves))
        detail_children.append(topic(section["title"], f"ep{ep}/{section['title']}", sub_nodes))

    def sheet(title, root_title, children, key):
        return {
            "id": stable_id(f"ep{ep}/sheet/{key}"),
            "class": "sheet",
            "title": title,
            "rootTopic": topic(root_title, f"ep{ep}/root/{key}", children, root=True),
            "topicPositioning": "fixed",
        }

    content = [
        sheet("01 总览图", f"第{ep}集｜{case['title']}", overview_children, "overview"),
        sheet("02 详细知识树", f"第{ep}集｜{case['title']}", detail_children, "detail"),
    ]
    ids = [node["id"] for item in content for node in walk(item["rootTopic"])]
    assert len(ids) == len(set(ids))
    metadata = {"creator": {"name": "Xmind", "version": "24.0"}, "activeSheetId": content[0]["id"]}
    manifest = {"file-entries": {"content.json": {}, "metadata.json": {}, "Thumbnails/thumbnail.png": {}}}
    transparent_png = base64.b64decode(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4z8DwHwAFgAI/W9n7WQAAAABJRU5ErkJggg=="
    )
    out = ROOT / f"齐昊老师_学员实战案例_第{ep}集_思维导图.xmind"
    with zipfile.ZipFile(out, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("content.json", json.dumps(content, ensure_ascii=False, separators=(",", ":")))
        archive.writestr("metadata.json", json.dumps(metadata, ensure_ascii=False, separators=(",", ":")))
        archive.writestr("manifest.json", json.dumps(manifest, ensure_ascii=False, separators=(",", ":")))
        archive.writestr("Thumbnails/thumbnail.png", transparent_png)
    with zipfile.ZipFile(out) as archive:
        assert archive.testzip() is None
        loaded = json.loads(archive.read("content.json"))
        counts = [sum(1 for _ in walk(item["rootTopic"])) for item in loaded]
    return out, counts


def build_markdown(case):
    lines = [f"# 第{case['episode']}集｜{case['title']}", "", f"> {case['core']}", "", "## 一页总览"]
    for label, value in case["overview"]:
        lines.extend([f"### {label}", f"- {value}"])
    for section in case["sections"]:
        lines.extend(["", f"## {section['title']}"])
        for sub_title, bullets in section["subs"]:
            lines.append(f"### {sub_title}")
            lines.extend(f"- {item}" for item in bullets)
    lines.extend(["", "## 来源与准确性说明", f"- 视频编号：{case['video_id']}", f"- 原链接：{case['source_url']}"])
    out = ROOT / f"齐昊老师_学员实战案例_第{case['episode']}集_思维导图源.md"
    out.write_text("\n".join(lines), encoding="utf-8")
    return out


if __name__ == "__main__":
    for case in CASES:
        md = build_markdown(case)
        docx = build_docx(case)
        xmind, counts = build_xmind(case)
        print(f"EP{case['episode']}\tDOCX={docx.stat().st_size}\tXMIND={xmind.stat().st_size}\tTOPICS={counts}\tMD={md.stat().st_size}")
