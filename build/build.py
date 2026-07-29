#!/usr/bin/env python3
"""
Build "Restore & Rebuild" from Markdown source into a typeset DOCX,
plus a single concatenated Markdown file.

Reading order is declared explicitly in manuscript/order.txt so chapters can be
reordered without renaming files.

Usage:  python3 build/build.py
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

ROOT = Path(__file__).resolve().parent.parent
MANUSCRIPT = ROOT / "manuscript"
DIST = ROOT / "dist"
ORDER_FILE = MANUSCRIPT / "order.txt"

BODY_FONT = "Georgia"
DISPLAY_FONT = "Georgia"
INK = RGBColor(0x1A, 0x1A, 0x1A)
MUTED = RGBColor(0x5A, 0x5A, 0x5A)

# Word-count band for daily chapters; anything outside gets flagged loudly.
# Set at 750-1100 while drafting, then widened once the chapters existed: a
# chapter carrying a verse, a mechanism, a meditation and an action does not
# fit in 900 words without going thin somewhere. 850-1250 is a four-to-six
# minute read, which is still one sitting before the market opens.
CHAPTER_MIN, CHAPTER_MAX = 850, 1250

# Sections whose body paragraphs get special treatment.
SEC_MEDITATION = "the meditation"
SEC_AFFIRMATION = "say it out loud"
SEC_CARRY = "carry this"


# --------------------------------------------------------------------------- #
# low-level docx helpers
# --------------------------------------------------------------------------- #

def _set_spacing(pf, before=0, after=8, line=1.3):
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line


def _letter_space(style, twentieths_of_point: int):
    """Apply character spacing (tracking) to a style."""
    rpr = style.element.get_or_add_rPr()
    spacing = OxmlElement("w:spacing")
    spacing.set(qn("w:val"), str(twentieths_of_point))
    rpr.append(spacing)


def _left_rule(style, color="C8B79A", size=18, space=14):
    """Draw a vertical rule down the left edge of a paragraph style."""
    ppr = style.element.get_or_add_pPr()
    borders = OxmlElement("w:pBdr")
    left = OxmlElement("w:left")
    left.set(qn("w:val"), "single")
    left.set(qn("w:sz"), str(size))
    left.set(qn("w:space"), str(space))
    left.set(qn("w:color"), color)
    borders.append(left)
    ppr.append(borders)


def _keep_with_next(style):
    ppr = style.element.get_or_add_pPr()
    for tag in ("w:keepNext", "w:keepLines"):
        el = OxmlElement(tag)
        el.set(qn("w:val"), "true")
        ppr.append(el)


def _no_widows(style):
    ppr = style.element.get_or_add_pPr()
    el = OxmlElement("w:widowControl")
    el.set(qn("w:val"), "true")
    ppr.append(el)


def _new_style(doc, name, *, size, bold=False, italic=False, align=None,
               color=INK, font=BODY_FONT, caps=False, base="Normal"):
    st = doc.styles.add_style(name, 1)  # 1 == WD_STYLE_TYPE.PARAGRAPH
    st.base_style = doc.styles[base]
    f = st.font
    f.name = font
    f.size = Pt(size)
    f.bold = bold
    f.italic = italic
    f.color.rgb = color
    f.all_caps = caps
    # East-Asian font mapping so LibreOffice doesn't substitute unpredictably
    rpr = st.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs"):
        rfonts.set(qn(attr), font)
    if align is not None:
        st.paragraph_format.alignment = align
    _no_widows(st)
    return st


def build_styles(doc):
    normal = doc.styles["Normal"]
    normal.font.name = BODY_FONT
    normal.font.size = Pt(10.5)
    normal.font.color.rgb = INK
    _set_spacing(normal.paragraph_format)

    s = _new_style
    C, L, J = WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.JUSTIFY

    st = s(doc, "RRBookTitle", size=26, bold=True, align=C, font=DISPLAY_FONT)
    _set_spacing(st.paragraph_format, before=0, after=10, line=1.05)

    st = s(doc, "RRBookSubtitle", size=12.5, italic=True, align=C, color=MUTED)
    _set_spacing(st.paragraph_format, before=0, after=24, line=1.3)

    st = s(doc, "RRBookByline", size=11, align=C, color=MUTED, caps=True)
    _letter_space(st, 30)
    _set_spacing(st.paragraph_format, before=36, after=0)

    st = s(doc, "RRPartKicker", size=9.5, align=C, color=MUTED, caps=True)
    _letter_space(st, 60)
    _set_spacing(st.paragraph_format, before=0, after=10)

    st = s(doc, "RRPartTitle", size=22, bold=True, align=C, font=DISPLAY_FONT)
    _set_spacing(st.paragraph_format, before=0, after=14, line=1.1)

    st = s(doc, "RRPartBlurb", size=11, italic=True, align=C, color=MUTED)
    _set_spacing(st.paragraph_format, before=0, after=0, line=1.45)
    st.paragraph_format.left_indent = Inches(0.6)
    st.paragraph_format.right_indent = Inches(0.6)

    st = s(doc, "RRDayKicker", size=9, color=MUTED, caps=True)
    _letter_space(st, 60)
    _set_spacing(st.paragraph_format, before=0, after=4)
    _keep_with_next(st)

    st = s(doc, "RRDayTitle", size=19, bold=True, font=DISPLAY_FONT)
    _set_spacing(st.paragraph_format, before=0, after=16, line=1.08)
    _keep_with_next(st)

    st = s(doc, "RRSectionHead", size=9.5, bold=True, color=MUTED, caps=True)
    _letter_space(st, 45)
    _set_spacing(st.paragraph_format, before=16, after=6)
    _keep_with_next(st)

    st = s(doc, "RRBodyText2", size=10.5, align=J)
    _set_spacing(st.paragraph_format, before=0, after=8, line=1.34)

    st = s(doc, "RRVerse", size=11, italic=True, align=L)
    _set_spacing(st.paragraph_format, before=4, after=2, line=1.4)
    st.paragraph_format.left_indent = Inches(0.28)
    st.paragraph_format.right_indent = Inches(0.28)
    _left_rule(st, color="B8A17E", size=14, space=12)

    st = s(doc, "RRVerseRef", size=9, color=MUTED, caps=True)
    _letter_space(st, 40)
    _set_spacing(st.paragraph_format, before=2, after=14)
    st.paragraph_format.left_indent = Inches(0.28)

    st = s(doc, "RRMeditation", size=10.5, align=L)
    _set_spacing(st.paragraph_format, before=0, after=8, line=1.38)
    st.paragraph_format.left_indent = Inches(0.22)
    _left_rule(st, color="D6D0C4", size=12, space=12)

    st = s(doc, "RRAffirmation", size=12, bold=True, italic=True, align=L,
           font=DISPLAY_FONT)
    _set_spacing(st.paragraph_format, before=6, after=6, line=1.3)
    st.paragraph_format.left_indent = Inches(0.3)
    st.paragraph_format.right_indent = Inches(0.2)

    st = s(doc, "RRCarryThis", size=12, italic=True, align=C, font=DISPLAY_FONT)
    _set_spacing(st.paragraph_format, before=10, after=10, line=1.35)
    st.paragraph_format.left_indent = Inches(0.4)
    st.paragraph_format.right_indent = Inches(0.4)

    st = s(doc, "RRBulletItem", size=10.5, align=L)
    _set_spacing(st.paragraph_format, before=0, after=5, line=1.32)
    st.paragraph_format.left_indent = Inches(0.32)
    st.paragraph_format.first_line_indent = Inches(-0.18)

    st = s(doc, "RROrnament", size=12, align=C, color=MUTED)
    _letter_space(st, 60)
    _set_spacing(st.paragraph_format, before=10, after=10)


# --------------------------------------------------------------------------- #
# inline markdown -> runs
# --------------------------------------------------------------------------- #

INLINE = re.compile(r"(\*\*.+?\*\*|(?<!\w)_[^_]+_(?!\w)|(?<!\*)\*[^*]+\*(?!\*))")


def typographic(text: str) -> str:
    """Straight quotes and three dots are fine in the source files; the printed
    page should have proper ones."""
    text = text.replace("...", "…")
    text = re.sub(r'(^|[\s(\[\u2014-])"', "\\1\u201c", text)
    text = text.replace('"', "\u201d")
    text = re.sub(r"(?<=[A-Za-z0-9])'(?=[A-Za-z])", "\u2019", text)
    text = re.sub(r"(^|[\s(\[\u2014])'", "\\1\u2018", text)
    return text.replace("'", "\u2019")


def add_runs(par, text):
    for tok in INLINE.split(text):
        if not tok:
            continue
        if tok.startswith("**") and tok.endswith("**"):
            par.add_run(typographic(tok[2:-2])).bold = True
        elif tok.startswith("_") and tok.endswith("_"):
            par.add_run(typographic(tok[1:-1])).italic = True
        elif tok.startswith("*") and tok.endswith("*"):
            par.add_run(typographic(tok[1:-1])).italic = True
        else:
            par.add_run(typographic(tok))
    return par


def para(doc, style, text=""):
    p = doc.add_paragraph(style=style)
    if text:
        add_runs(p, text)
    return p


def page_break(doc):
    """Break to a fresh page. The break lives in its own paragraph so that the
    space_before of whatever follows still applies on the new page."""
    p = doc.add_paragraph(style="RRBodyText2")
    p.paragraph_format.space_after = Pt(0)
    p.add_run().add_break(WD_BREAK.PAGE)


def spacers(doc, n):
    for _ in range(n):
        doc.add_paragraph(style="RRBodyText2")


# --------------------------------------------------------------------------- #
# markdown file -> docx
# --------------------------------------------------------------------------- #

def normalize(s: str) -> str:
    return re.sub(r"[^a-z ]", "", s.lower()).strip()


def emit_table(doc, rows):
    header, body = rows[0], rows[1:]
    table = doc.add_table(rows=1, cols=len(header))
    table.style = "Table Grid"
    table.autofit = True
    for cell, text in zip(table.rows[0].cells, header):
        cell.text = ""
        p = cell.paragraphs[0]
        p.style = doc.styles["RRBulletItem"]
        p.paragraph_format.left_indent = Inches(0.04)
        p.paragraph_format.first_line_indent = Inches(0)
        run = add_runs(p, text).runs
        for r in run:
            r.bold = True
    for line in body:
        cells = table.add_row().cells
        for cell, text in zip(cells, line):
            cell.text = ""
            p = cell.paragraphs[0]
            p.style = doc.styles["RRBulletItem"]
            p.paragraph_format.left_indent = Inches(0.04)
            p.paragraph_format.first_line_indent = Inches(0)
            add_runs(p, text)
    doc.add_paragraph(style="RRBodyText2")


def render_file(doc, path: Path, state: dict):
    raw = path.read_text(encoding="utf-8")
    is_title_page = "<!-- title-page -->" in raw
    is_part = "<!-- part -->" in raw

    lines = [ln.rstrip() for ln in raw.split("\n")]
    section = ""
    table_buf: list[list[str]] = []
    first_heading_seen = False

    def flush_table():
        nonlocal table_buf
        if table_buf:
            emit_table(doc, table_buf)
            table_buf = []

    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        if stripped.startswith("<!--"):
            i += 1
            continue

        # ---- tables ----
        if stripped.startswith("|") and stripped.endswith("|"):
            cells = [c.strip() for c in stripped.strip("|").split("|")]
            if not all(re.fullmatch(r":?-{2,}:?", c) for c in cells):
                table_buf.append(cells)
            i += 1
            continue
        flush_table()

        if not stripped:
            i += 1
            continue

        # ---- ornament / rule ----
        if stripped in ("---", "***", "* * *"):
            para(doc, "RROrnament", "§")
            i += 1
            continue

        # ---- headings ----
        if stripped.startswith("# "):
            title = stripped[2:].strip()
            needs_break = state["started"]
            state["started"] = True
            first_heading_seen = True

            if needs_break:
                page_break(doc)

            kicker = state.pop("pending_kicker", "")
            if is_title_page:
                spacers(doc, 5)
                para(doc, "RRBookTitle", title)
            elif is_part:
                spacers(doc, 6)
                para(doc, "RRPartTitle", title)
            else:
                if kicker:
                    para(doc, "RRDayKicker", kicker)
                para(doc, "RRDayTitle", title)
            i += 1
            continue

        if stripped.startswith("## "):
            head = stripped[3:].strip()
            section = normalize(head)
            if is_title_page:
                para(doc, "RRBookSubtitle", head)
            elif is_part:
                para(doc, "RRPartKicker", head)
            else:
                para(doc, "RRSectionHead", head)
            i += 1
            continue

        if stripped.startswith("### "):
            head = stripped[4:].strip()
            if is_title_page:
                para(doc, "RRBookByline", head)
            else:
                p = para(doc, "RRBodyText2", f"**{head}**")
                p.paragraph_format.space_before = Pt(10)
            i += 1
            continue

        # ---- verse block ----
        if stripped.startswith(">"):
            content = stripped.lstrip(">").strip()
            if content.startswith("—") or content.startswith("--"):
                para(doc, "RRVerseRef", content.lstrip("—- ").strip())
            elif content:
                para(doc, "RRVerse", content)
            i += 1
            continue

        # ---- lists ----
        m_num = re.match(r"^(\d+)\.\s+(.*)$", stripped)
        m_bul = re.match(r"^[-*]\s+(.*)$", stripped)
        if m_num or m_bul:
            text = (m_num.group(2) if m_num else m_bul.group(1)).strip()
            if section == SEC_AFFIRMATION:
                para(doc, "RRAffirmation", f"“{text}”" if not text.startswith("“") else text)
            elif m_num:
                para(doc, "RRBulletItem", f"{m_num.group(1)}.  {text}")
            else:
                para(doc, "RRBulletItem", f"•  {text}")
            i += 1
            continue

        # ---- plain paragraph ----
        if is_title_page and not first_heading_seen:
            para(doc, "RRBookSubtitle", stripped)
        elif is_title_page:
            para(doc, "RRBookSubtitle", stripped)
        elif is_part:
            para(doc, "RRPartBlurb", stripped)
        elif section == SEC_MEDITATION:
            para(doc, "RRMeditation", stripped)
        elif section == SEC_CARRY:
            para(doc, "RRCarryThis", stripped)
        else:
            para(doc, "RRBodyText2", stripped)
        i += 1

    flush_table()


# --------------------------------------------------------------------------- #
# page setup
# --------------------------------------------------------------------------- #

def enable_hyphenation(doc):
    """Justified text on a 4.1-inch measure looks gappy without hyphenation."""
    settings = doc.settings.element
    for tag, val in (("w:autoHyphenation", "true"),
                     ("w:doNotHyphenateCaps", "true"),
                     ("w:consecutiveHyphenLimit", "2")):
        el = OxmlElement(tag)
        el.set(qn("w:val"), val)
        settings.append(el)


def setup_page(doc):
    sec = doc.sections[0]
    sec.page_width = Inches(5.5)
    sec.page_height = Inches(8.5)
    sec.top_margin = Inches(0.72)
    sec.bottom_margin = Inches(0.72)
    sec.left_margin = Inches(0.68)
    sec.right_margin = Inches(0.68)
    sec.header_distance = Inches(0.4)
    sec.footer_distance = Inches(0.4)

    # page numbers, centered in the footer
    footer_p = sec.footer.paragraphs[0]
    footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = footer_p.add_run()
    run.font.name = BODY_FONT
    run.font.size = Pt(9)
    run.font.color.rgb = MUTED
    for instr, kind in (("begin", "fldChar"), ("PAGE", "instrText"), ("end", "fldChar")):
        el = OxmlElement(f"w:{kind}")
        if kind == "fldChar":
            el.set(qn("w:fldCharType"), instr)
        else:
            el.set(qn("xml:space"), "preserve")
            el.text = " PAGE "
        run._r.append(el)


# --------------------------------------------------------------------------- #
# orchestration
# --------------------------------------------------------------------------- #

def read_order() -> list[Path]:
    if not ORDER_FILE.exists():
        sys.exit(f"missing manifest: {ORDER_FILE}")
    paths = []
    for line in ORDER_FILE.read_text(encoding="utf-8").splitlines():
        line = line.split("#", 1)[0].strip()
        if not line:
            continue
        p = MANUSCRIPT / line
        if not p.exists():
            sys.exit(f"manifest lists a file that does not exist: {line}")
        paths.append(p)
    return paths


def kicker_for(path: Path) -> str:
    m = re.match(r"day-(\d+)-", path.name)
    return f"Day {int(m.group(1))}" if m else ""


def word_count(path: Path) -> int:
    """Words of actual prose: headings and section labels are fixed scaffolding
    and would flatter every chapter by the same ~45 words."""
    body = [ln for ln in path.read_text(encoding="utf-8").splitlines()
            if not ln.lstrip().startswith(("#", "<!--"))]
    txt = re.sub(r"[>*_|`]", " ", "\n".join(body))
    return len(txt.split())


def main():
    DIST.mkdir(exist_ok=True)
    files = read_order()

    doc = Document()
    setup_page(doc)
    enable_hyphenation(doc)
    build_styles(doc)

    state = {"started": False}
    md_parts = []
    for path in files:
        k = kicker_for(path)
        if k:
            state["pending_kicker"] = k
        render_file(doc, path, state)
        md_parts.append(path.read_text(encoding="utf-8").strip())

    docx_path = DIST / "restore-and-rebuild.docx"
    doc.save(docx_path)

    md_path = DIST / "restore-and-rebuild.md"
    md_path.write_text("\n\n\n".join(md_parts) + "\n", encoding="utf-8")

    # ---- report ----
    chapters = [p for p in files if p.name.startswith("day-")]
    total = sum(word_count(p) for p in files)
    print(f"\n  files: {len(files)}   chapters: {len(chapters)}   words: {total:,}\n")

    flagged = []
    print(f"  {'chapter':<46}{'words':>7}")
    print(f"  {'-' * 53}")
    for p in chapters:
        wc = word_count(p)
        mark = ""
        if wc < CHAPTER_MIN:
            mark, flagged = "  SHORT", flagged + [p.name]
        elif wc > CHAPTER_MAX:
            mark, flagged = "  LONG", flagged + [p.name]
        print(f"  {p.stem:<46}{wc:>7}{mark}")

    print(f"\n  wrote {docx_path.relative_to(ROOT)}")
    print(f"  wrote {md_path.relative_to(ROOT)}")
    if flagged:
        print(f"\n  {len(flagged)} chapter(s) outside the "
              f"{CHAPTER_MIN}-{CHAPTER_MAX} word band")
    print()


if __name__ == "__main__":
    main()
