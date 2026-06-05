import argparse
import re
import shutil
import sys
from pathlib import Path

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, Twips


if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")


TITLE_TOC = "目录"
TITLE_FIGURES = "插图清单"
TITLE_TABLES = "附表清单"
TITLE_INTRODUCTION = "引言"

HEADING_TEXT_RE = re.compile(r"^\d+(?:\.\d+){0,2}\s+\S+")
FIG_RE = re.compile(r"^图\s+\d+(?:\.\d+)?\s+")
TAB_RE = re.compile(r"^表\s+\d+(?:\.\d+)?\s+")
TOC_LINE_RE = re.compile(r"^(\d+(?:\.\d+){0,2})\s+(.+?)\s+(\d+)$")
HEADING_STYLE_IDS = {1: "Heading1", 2: "Heading2", 3: "Heading3"}
HEADING_STYLE_NAMES = {
    1: ("Heading 1", "标题 1"),
    2: ("Heading 2", "标题 2"),
    3: ("Heading 3", "标题 3"),
}
FRONT_TITLES = {TITLE_TOC, TITLE_FIGURES, TITLE_TABLES}


def clean(text: str) -> str:
    return " ".join(text.split())


def remove_paragraph(paragraph) -> None:
    element = paragraph._element
    element.getparent().remove(element)
    paragraph._p = paragraph._element = None


def ensure_run_fonts(run, font_name: str, size_pt: float | None = None, bold: bool | None = None):
    run.font.name = font_name
    if size_pt is not None:
        run.font.size = Pt(size_pt)
    if bold is not None:
        run.bold = bold
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.rFonts
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    for key in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rfonts.set(qn(key), font_name)


def ensure_style_fonts(style, font_name: str, size_pt: float | None = None, bold: bool | None = None):
    style.font.name = font_name
    if size_pt is not None:
        style.font.size = Pt(size_pt)
    if bold is not None:
        style.font.bold = bold
    rpr = style._element.get_or_add_rPr()
    rfonts = rpr.rFonts
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    for key in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rfonts.set(qn(key), font_name)


def set_paragraph_style_id(paragraph, style_id: str) -> None:
    ppr = paragraph._p.get_or_add_pPr()
    pstyle = ppr.pStyle
    if pstyle is None:
        pstyle = OxmlElement("w:pStyle")
        ppr.insert(0, pstyle)
    pstyle.set(qn("w:val"), style_id)


def clear_outline_level(paragraph) -> None:
    ppr = paragraph._p.get_or_add_pPr()
    for outline in list(ppr.xpath("./w:outlineLvl")):
        ppr.remove(outline)


def ensure_page_break_at_start(paragraph) -> None:
    for br in paragraph._p.xpath('.//w:br[@w:type="page"]'):
        return
    first_run = paragraph.runs[0] if paragraph.runs else paragraph.add_run()
    br = OxmlElement("w:br")
    br.set(qn("w:type"), "page")
    first_run._r.insert(0, br)


def remove_page_breaks(paragraph) -> None:
    for br in list(paragraph._p.xpath('.//w:br[@w:type="page"]')):
        br.getparent().remove(br)


def find_style(doc: Document, names=(), style_ids=()):
    name_set = {name.casefold() for name in names}
    id_set = set(style_ids)
    for style in doc.styles:
        if style.name and style.name.casefold() in name_set:
            return style
        if getattr(style, "style_id", None) in id_set:
            return style
    return None


def ensure_style(doc: Document, name: str, base: str = "Normal"):
    style = find_style(doc, names=(name,))
    if style is None:
        style = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
    base_style = find_style(doc, names=(base, "Normal", "正文"), style_ids=("Normal",))
    if base_style is not None:
        style.base_style = base_style
    ensure_style_fonts(style, "宋体", 12)
    style.paragraph_format.line_spacing = 1.5
    style.paragraph_format.space_before = Pt(0)
    style.paragraph_format.space_after = Pt(0)
    return style


def ensure_table_of_figures_style(doc: Document):
    style = find_style(doc, names=("table of figures",))
    if style is None:
        style = doc.styles.add_style("table of figures", WD_STYLE_TYPE.PARAGRAPH)
    base_style = find_style(doc, names=("Normal", "正文"), style_ids=("a", "Normal"))
    if base_style is not None:
        style.base_style = base_style
    ensure_style_fonts(style, "宋体", 12)
    style.paragraph_format.line_spacing = 1.5
    style.paragraph_format.space_before = Pt(0)
    style.paragraph_format.space_after = Pt(0)
    return style


def find_paragraph(doc: Document, text: str, start: int = 0):
    for i, paragraph in enumerate(doc.paragraphs[start:], start):
        if clean(paragraph.text) == text:
            return i
    return None


def remove_paragraphs_before(doc: Document, text: str) -> int:
    idx = find_paragraph(doc, text)
    if idx is None or idx == 0:
        return 0
    for paragraph in list(doc.paragraphs[:idx]):
        remove_paragraph(paragraph)
    return idx


def style_name(paragraph) -> str:
    style = paragraph.style
    if style is None:
        return ""
    return style.name or ""


def style_id(paragraph) -> str:
    style = paragraph.style
    if style is None:
        return ""
    return getattr(style, "style_id", "") or ""


def is_heading_style(paragraph, levels=(1, 2, 3)) -> bool:
    name = style_name(paragraph)
    sid = style_id(paragraph)
    for level in levels:
        if sid == HEADING_STYLE_IDS[level] or name in HEADING_STYLE_NAMES[level]:
            return True
    return False


def get_heading_style(doc: Document, level: int):
    style = find_style(
        doc,
        names=HEADING_STYLE_NAMES[level],
        style_ids=(HEADING_STYLE_IDS[level],),
    )
    if style is not None:
        return style

    style = doc.styles.add_style(HEADING_STYLE_NAMES[level][0], WD_STYLE_TYPE.PARAGRAPH)
    base_style = find_style(doc, names=("Normal", "正文"), style_ids=("Normal",))
    if base_style is not None:
        style.base_style = base_style
    ensure_style_fonts(style, "宋体", 14 if level == 1 else 12)
    return style


def first_body_index(doc: Document):
    front_end = find_paragraph(doc, TITLE_TABLES)
    search_start = 0 if front_end is None else front_end + 1
    for i, paragraph in enumerate(doc.paragraphs):
        if i < search_start:
            continue
        text = clean(paragraph.text)
        if text == TITLE_INTRODUCTION:
            return i
        if is_heading_style(paragraph, levels=(1,)) and HEADING_TEXT_RE.match(text):
            return i
    for i, paragraph in enumerate(doc.paragraphs):
        if i < search_start:
            continue
        text = clean(paragraph.text)
        if text == TITLE_INTRODUCTION:
            return i
        if HEADING_TEXT_RE.match(text) and text not in {TITLE_TOC, TITLE_FIGURES, TITLE_TABLES}:
            return i
    return None


def apply_heading_styles(doc: Document) -> int:
    body_start = first_body_index(doc)
    if body_start is None:
        return 0
    heading_styles = {
        1: get_heading_style(doc, 1),
        2: get_heading_style(doc, 2),
        3: get_heading_style(doc, 3),
    }
    changed = 0
    for paragraph in doc.paragraphs[body_start:]:
        text = clean(paragraph.text)
        if text in FRONT_TITLES:
            continue
        if text == TITLE_INTRODUCTION:
            paragraph.style = heading_styles[1]
            changed += 1
            continue
        if re.match(r"^\d+\s+\S+", text):
            paragraph.style = heading_styles[1]
            changed += 1
        elif re.match(r"^\d+\.\d+\s+\S+", text):
            paragraph.style = heading_styles[2]
            changed += 1
        elif re.match(r"^\d+\.\d+\.\d+\s+\S+", text):
            paragraph.style = heading_styles[3]
            changed += 1
    return changed


def make_toc_field(instr_text: str, placeholder: str):
    p = OxmlElement("w:p")
    ppr = OxmlElement("w:pPr")
    tabs = OxmlElement("w:tabs")
    tab = OxmlElement("w:tab")
    tab.set(qn("w:val"), "right")
    tab.set(qn("w:leader"), "dot")
    tab.set(qn("w:pos"), "8296")
    tabs.append(tab)
    spacing = OxmlElement("w:spacing")
    spacing.set(qn("w:line"), "360")
    spacing.set(qn("w:lineRule"), "auto")
    ppr.append(tabs)
    ppr.append(spacing)
    p.append(ppr)

    def add_child(child):
        r = OxmlElement("w:r")
        r.append(child)
        p.append(r)

    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    begin.set(qn("w:dirty"), "true")
    add_child(begin)

    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = instr_text
    add_child(instr)

    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    add_child(separate)

    text = OxmlElement("w:t")
    text.text = placeholder
    add_child(text)

    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    add_child(end)
    return p


def ensure_front_title(doc: Document, text: str, page_break_before: bool = False):
    idx = find_paragraph(doc, text)
    if idx is not None:
        paragraph = doc.paragraphs[idx]
        normal_style = find_style(doc, names=("Normal", "正文"), style_ids=("a", "Normal"))
        if normal_style is not None:
            paragraph.style = normal_style
            set_paragraph_style_id(paragraph, getattr(normal_style, "style_id", "a") or "a")
        else:
            set_paragraph_style_id(paragraph, "a")
        clear_outline_level(paragraph)
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraph.paragraph_format.line_spacing = None if page_break_before else 1.5
        paragraph.paragraph_format.space_before = None
        paragraph.paragraph_format.space_after = None
        if page_break_before:
            ensure_page_break_at_start(paragraph)
        else:
            remove_page_breaks(paragraph)
        for run in paragraph.runs:
            ensure_run_fonts(run, "仿宋", 16, True)
    return idx


def replace_between(doc: Document, start_text: str, end_text: str, replacement) -> bool:
    start = find_paragraph(doc, start_text)
    if start is None:
        return False
    end = find_paragraph(doc, end_text, start + 1)
    if end is None:
        return False
    for paragraph in list(doc.paragraphs[start + 1 : end]):
        remove_paragraph(paragraph)
    if replacement is not None:
        doc.paragraphs[start]._p.addnext(replacement)
    return True


def prepare_main_toc(doc: Document) -> bool:
    ensure_table_of_figures_style(doc)
    ensure_front_title(doc, TITLE_TOC)
    ensure_front_title(doc, TITLE_FIGURES, page_break_before=True)
    ensure_front_title(doc, TITLE_TABLES, page_break_before=True)
    return replace_between(
        doc,
        TITLE_TOC,
        TITLE_FIGURES,
        make_toc_field(r'TOC \o "1-3" \h \z \u', "目录将在 Word 中自动更新。"),
    )


def parse_toc_pages(doc: Document) -> dict[str, int]:
    pages = {}
    for paragraph in doc.paragraphs:
        if not style_name(paragraph).lower().startswith("toc "):
            continue
        text = clean(paragraph.text)
        match = TOC_LINE_RE.match(text)
        if not match:
            continue
        number, title, page = match.groups()
        pages[f"{number} {title}"] = int(page)
    return pages


def nearest_heading_page(heading_stack: list[str], pages: dict[str, int]):
    for heading in reversed(heading_stack):
        if heading in pages:
            return pages[heading]
    return None


def collect_caption_entries(doc: Document, pages: dict[str, int]):
    fig_style = ensure_style(doc, "图题注")
    tab_style = ensure_style(doc, "表题注")
    body_start = first_body_index(doc)
    if body_start is None:
        return [], []

    headings = []
    figures = []
    tables = []
    seen_figures = set()
    seen_tables = set()
    for paragraph in doc.paragraphs[body_start:]:
        text = clean(paragraph.text)
        if is_heading_style(paragraph):
            headings.append(text)
            continue
        if FIG_RE.match(text):
            paragraph.style = fig_style
            if text not in seen_figures:
                figures.append((text, nearest_heading_page(headings, pages)))
                seen_figures.add(text)
        elif TAB_RE.match(text) and "从技术环节" not in text:
            paragraph.style = tab_style
            if text not in seen_tables:
                tables.append((text, nearest_heading_page(headings, pages)))
                seen_tables.add(text)
    return figures, tables


def write_static_list(doc: Document, start_text: str, end_text: str, entries) -> bool:
    list_style = ensure_table_of_figures_style(doc)
    start = find_paragraph(doc, start_text)
    if start is None:
        return False
    end = find_paragraph(doc, end_text, start + 1)
    if end is None:
        return False

    for paragraph in list(doc.paragraphs[start + 1 : end]):
        remove_paragraph(paragraph)

    anchor = doc.paragraphs[start]._p
    for caption, page in reversed(entries):
        p = doc.add_paragraph()
        p.style = list_style
        p.paragraph_format.left_indent = Twips(483)
        p.paragraph_format.first_line_indent = Twips(-485)
        p.paragraph_format.tab_stops.add_tab_stop(
            Twips(8296), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS
        )
        p.paragraph_format.line_spacing = 1.5
        caption_run = p.add_run(caption)
        tab_run = p.add_run("\t")
        page_run = p.add_run("" if page is None else str(page))
        for run in (caption_run, tab_run, page_run):
            ensure_run_fonts(run, "宋体", 12)
        anchor.addnext(p._p)
    return True


def first_body_heading_text(doc: Document) -> str | None:
    idx = first_body_index(doc)
    if idx is None:
        return None
    return clean(doc.paragraphs[idx].text)


def ensure_body_starts_on_new_page(doc: Document) -> bool:
    idx = first_body_index(doc)
    if idx is None:
        return False
    ensure_page_break_at_start(doc.paragraphs[idx])
    return True


def prepare(path: Path, backup: bool) -> None:
    if backup:
        backup_path = path.with_name(f"{path.stem}.before-report-directories{path.suffix}")
        shutil.copy2(path, backup_path)
        print(f"backup={backup_path}")
    doc = Document(str(path))
    front_removed = remove_paragraphs_before(doc, TITLE_TOC)
    changed = apply_heading_styles(doc)
    toc_ok = prepare_main_toc(doc)
    doc.save(str(path))
    print(f"front_paragraphs_removed={front_removed}")
    print(f"heading_styles_applied={changed}")
    print(f"main_toc_prepared={toc_ok}")


def finalize(path: Path) -> None:
    doc = Document(str(path))
    pages = parse_toc_pages(doc)
    figures, tables = collect_caption_entries(doc, pages)
    body_heading = first_body_heading_text(doc)
    fig_ok = write_static_list(doc, TITLE_FIGURES, TITLE_TABLES, figures)
    table_ok = False if body_heading is None else write_static_list(doc, TITLE_TABLES, body_heading, tables)
    body_page_break_ok = ensure_body_starts_on_new_page(doc)
    doc.save(str(path))
    print(f"toc_pages_found={len(pages)}")
    print(f"figures={len(figures)} figure_list_updated={fig_ok} missing_pages={sum(p is None for _, p in figures)}")
    print(f"tables={len(tables)} table_list_updated={table_ok} missing_pages={sum(p is None for _, p in tables)}")
    print(f"body_page_break_updated={body_page_break_ok}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Prepare/update Word report directories.")
    parser.add_argument("path", type=Path)
    parser.add_argument("--phase", choices=["prepare", "finalize"], required=True)
    parser.add_argument("--no-backup", action="store_true")
    args = parser.parse_args()
    path = args.path.expanduser().resolve(strict=True)

    if args.phase == "prepare":
        prepare(path, backup=not args.no_backup)
    else:
        finalize(path)


if __name__ == "__main__":
    main()
