#!/usr/bin/env python3
"""content.json -> 규칙(references/report-rules.md)을 따르는 DOCX 보고서 생성기.

사용법:
    python build_report.py content.json -o 보고서.docx

입력 형식은 references/content-schema.md 참고.
"""
import argparse
import json
import os
import re
import sys
import tempfile
from datetime import date

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Mm, RGBColor, Cm

# ---------------------------------------------------------------- design tokens (R5~R10)
FONT_LATIN = "Malgun Gothic"
FONT_EA = "맑은 고딕"
BODY_PT = 11
LINE_SPACING = 1.4
NAVY = RGBColor(0x1F, 0x38, 0x64)
BLUE = RGBColor(0x2E, 0x75, 0xB6)
GRAY = RGBColor(0x7F, 0x7F, 0x7F)
LINK = RGBColor(0x05, 0x63, 0xC1)
HEADER_FILL = "1F3864"
ZEBRA_FILL = "F2F5FA"
CALLOUT_FILL = "EAF1FB"
CALLOUT_LINE = "2E75B6"
MARGIN_MM = 25

CITE_RE = re.compile(r"\[@([^\]]+)\]")
XREF_RE = re.compile(r"\{(tbl|fig):([A-Za-z0-9_\-]+)\}")
BOLD_RE = re.compile(r"\*\*(.+?)\*\*")


class BuildError(Exception):
    pass


# ---------------------------------------------------------------- low-level XML helpers
def set_run_font(run, size=None, bold=None, color=None, italic=None):
    rpr = run._element.get_or_add_rPr()
    fonts = rpr.find(qn("w:rFonts"))
    if fonts is None:
        fonts = OxmlElement("w:rFonts")
        rpr.insert(0, fonts)
    fonts.set(qn("w:ascii"), FONT_LATIN)
    fonts.set(qn("w:hAnsi"), FONT_LATIN)
    fonts.set(qn("w:eastAsia"), FONT_EA)
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    if italic is not None:
        run.font.italic = italic
    if color is not None:
        run.font.color.rgb = color


def style_font(style, size, bold=False, color=None):
    style.font.name = FONT_LATIN
    style.font.size = Pt(size)
    style.font.bold = bold
    if color is not None:
        style.font.color.rgb = color
    rpr = style.element.get_or_add_rPr()
    fonts = rpr.find(qn("w:rFonts"))
    if fonts is None:
        fonts = OxmlElement("w:rFonts")
        rpr.insert(0, fonts)
    for a in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        fonts.set(qn(a), FONT_EA if a == "w:eastAsia" else FONT_LATIN)
    for a in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
        if fonts.get(qn(a)) is not None:
            del fonts.attrib[qn(a)]


def make_rpr(color=None, underline=False, bold=False, size=None):
    rpr = OxmlElement("w:rPr")
    f = OxmlElement("w:rFonts")
    f.set(qn("w:ascii"), FONT_LATIN)
    f.set(qn("w:hAnsi"), FONT_LATIN)
    f.set(qn("w:eastAsia"), FONT_EA)
    rpr.append(f)
    if bold:
        rpr.append(OxmlElement("w:b"))
    if color is not None:
        c = OxmlElement("w:color")
        c.set(qn("w:val"), str(color))
        rpr.append(c)
    if size is not None:
        s = OxmlElement("w:sz")
        s.set(qn("w:val"), str(int(size * 2)))
        rpr.append(s)
    if underline:
        u = OxmlElement("w:u")
        u.set(qn("w:val"), "single")
        rpr.append(u)
    return rpr


def text_run(text, rpr):
    r = OxmlElement("w:r")
    r.append(rpr)
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    r.append(t)
    return r


def add_external_link(par, text, url, size=None):
    """R12: 클릭 가능한 외부 하이퍼링크 (파란색 밑줄)."""
    r_id = par.part.relate_to(url, RT.HYPERLINK, is_external=True)
    h = OxmlElement("w:hyperlink")
    h.set(qn("r:id"), r_id)
    h.set(qn("w:history"), "1")
    h.append(text_run(text, make_rpr(color=LINK, underline=True, size=size)))
    par._p.append(h)


def add_internal_link(par, text, anchor, size=None):
    """R11: 본문 인용 [n] -> 참고문헌 항목 북마크로 이동하는 내부 링크."""
    h = OxmlElement("w:hyperlink")
    h.set(qn("w:anchor"), anchor)
    h.set(qn("w:history"), "1")
    h.append(text_run(text, make_rpr(color=BLUE, size=size)))
    par._p.append(h)


_bm_id = [100]


def add_bookmark(par, name):
    _bm_id[0] += 1
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(_bm_id[0]))
    start.set(qn("w:name"), name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(_bm_id[0]))
    ppr = par._p.find(qn("w:pPr"))
    idx = list(par._p).index(ppr) + 1 if ppr is not None else 0
    par._p.insert(idx, start)
    par._p.append(end)


def add_field(par, instr, placeholder="", size=None, color=None):
    def fld(kind):
        r = OxmlElement("w:r")
        r.append(make_rpr(color=color, size=size))
        fc = OxmlElement("w:fldChar")
        fc.set(qn("w:fldCharType"), kind)
        if kind == "begin":
            fc.set(qn("w:dirty"), "true")
        r.append(fc)
        return r

    par._p.append(fld("begin"))
    r = OxmlElement("w:r")
    r.append(make_rpr(color=color, size=size))
    it = OxmlElement("w:instrText")
    it.set(qn("xml:space"), "preserve")
    it.text = f" {instr} "
    r.append(it)
    par._p.append(r)
    par._p.append(fld("separate"))
    par._p.append(text_run(placeholder, make_rpr(color=color, size=size)))
    par._p.append(fld("end"))


def shade(element_pr, fill):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    element_pr.append(shd)


def para_border(par, sides=("bottom",), color="1F3864", sz=8, space=4):
    ppr = par._p.get_or_add_pPr()
    bdr = OxmlElement("w:pBdr")
    for side in sides:
        b = OxmlElement(f"w:{side}")
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), str(sz))
        b.set(qn("w:space"), str(space))
        b.set(qn("w:color"), color)
        bdr.append(b)
    ppr.append(bdr)


# ---------------------------------------------------------------- report builder
class ReportBuilder:
    def __init__(self, content, base_dir):
        self.c = content
        self.base_dir = base_dir
        self.meta = content["meta"]
        self.refs = {r["id"]: r for r in content.get("references", [])}
        self.cite_order = []  # ref ids in first-citation order (R12)
        self.tbl_no, self.fig_no = {}, {}
        self.tmp = tempfile.mkdtemp(prefix="report_")
        self.doc = Document()
        self._number_objects()
        self._collect_citations()

    # ---------- pre-passes
    def _iter_blocks(self):
        for sec in self.c["sections"]:
            yield from sec.get("blocks", [])
            for sub in sec.get("subsections", []):
                yield from sub.get("blocks", [])

    def _iter_texts(self):
        for b in self._iter_blocks():
            for k in ("text", "caption", "note"):
                if isinstance(b.get(k), str):
                    yield b[k]
            for item in b.get("items", []):
                yield item
            for row in b.get("rows", []):
                yield from (str(x) for x in row)
            for step in b.get("steps", []):
                yield step.get("desc", "")

    def _number_objects(self):
        t = f = 0
        for b in self._iter_blocks():
            if b["type"] == "table":
                t += 1
                if b.get("id"):
                    self.tbl_no[b["id"]] = t
            elif b["type"] in ("chart", "flow", "image"):
                f += 1
                if b.get("id"):
                    self.fig_no[b["id"]] = f

    def _collect_citations(self):
        for s in list(self.c.get("summary", [])) + list(self._iter_texts()):
            for m in CITE_RE.finditer(s):
                for key in self._split_keys(m.group(1)):
                    if key not in self.refs:
                        raise BuildError(f"참고문헌에 없는 인용 키: @{key}")
                    if key not in self.cite_order:
                        self.cite_order.append(key)
        unused = [k for k in self.refs if k not in self.cite_order]
        if unused:
            raise BuildError("본문에서 인용되지 않은 참고문헌(R12 위반): " + ", ".join(unused))
        for r in self.refs.values():
            if not str(r.get("url", "")).startswith(("http://", "https://")):
                raise BuildError(f"참고문헌 {r['id']}에 클릭 가능한 URL/DOI가 없습니다 (R12).")

    @staticmethod
    def _split_keys(group):
        return [k.strip().lstrip("@") for k in group.split(";") if k.strip()]

    # ---------- inline text: **bold**, [@cite], {tbl:id}/{fig:id}
    def inline(self, par, text, size=None, color=None, bold=False):
        def resolve_xref(m):
            kind, key = m.group(1), m.group(2)
            table = self.tbl_no if kind == "tbl" else self.fig_no
            if key not in table:
                raise BuildError(f"존재하지 않는 {kind} 참조: {key}")
            return ("표 " if kind == "tbl" else "그림 ") + str(table[key])

        text = XREF_RE.sub(resolve_xref, text)
        pos = 0
        tokens = []
        for m in re.finditer(r"\*\*(.+?)\*\*|\[@([^\]]+)\]", text):
            if m.start() > pos:
                tokens.append(("t", text[pos:m.start()]))
            if m.group(1) is not None:
                tokens.append(("b", m.group(1)))
            else:
                tokens.append(("c", m.group(2)))
            pos = m.end()
        if pos < len(text):
            tokens.append(("t", text[pos:]))

        for kind, val in tokens:
            if kind in ("t", "b"):
                r = par.add_run(val)
                set_run_font(r, size=size, bold=bold or kind == "b", color=color)
            else:
                nums = sorted(self.cite_order.index(k) + 1 for k in self._split_keys(val))
                self._raw(par, "[", size)
                for i, n in enumerate(nums):
                    if i:
                        self._raw(par, ", ", size)
                    add_internal_link(par, str(n), f"ref_{n}", size=size)
                self._raw(par, "]", size)

    def _raw(self, par, s, size):
        par._p.append(text_run(s, make_rpr(color=BLUE, size=size)))

    # ---------- document setup (R5~R8)
    def setup(self):
        d = self.doc
        st = d.styles
        normal = st["Normal"]
        style_font(normal, BODY_PT)
        pf = normal.paragraph_format
        pf.line_spacing = LINE_SPACING
        pf.space_after = Pt(6)
        h1, h2 = st["Heading 1"], st["Heading 2"]
        style_font(h1, 16, bold=True, color=NAVY)
        h1.paragraph_format.space_before = Pt(18)
        h1.paragraph_format.space_after = Pt(8)
        h1.paragraph_format.keep_with_next = True
        h1.paragraph_format.line_spacing = 1.2
        style_font(h2, 13, bold=True, color=BLUE)
        h2.paragraph_format.space_before = Pt(12)
        h2.paragraph_format.space_after = Pt(4)
        h2.paragraph_format.keep_with_next = True
        h2.paragraph_format.line_spacing = 1.2
        for name in ("List Bullet", "Caption", "Title", "Subtitle"):
            if name in [s.name for s in st]:
                style_font(st[name], BODY_PT)
        style_font(st["List Bullet"], BODY_PT)
        st["List Bullet"].paragraph_format.line_spacing = LINE_SPACING
        st["List Bullet"].paragraph_format.space_after = Pt(3)

        s = d.sections[0]
        s.page_width, s.page_height = Mm(210), Mm(297)
        for side in ("top_margin", "bottom_margin", "left_margin", "right_margin"):
            setattr(s, side, Mm(MARGIN_MM))
        s.header_distance = s.footer_distance = Mm(12)

        cp = d.core_properties
        cp.title = self.meta["title"]
        cp.subject = self.meta.get("subtitle", "")
        cp.author = self.meta.get("author", "")
        cp.keywords = ", ".join(self.meta.get("keywords", []))
        cp.language = "ko-KR"

        # Word가 열 때 목차/페이지 필드를 갱신하도록 (finalize_word.py가 처리하면 제거됨)
        settings = d.settings.element
        upd = OxmlElement("w:updateFields")
        upd.set(qn("w:val"), "true")
        settings.append(upd)

    # ---------- cover (R2)
    def cover(self):
        d, m = self.doc, self.meta
        p = d.add_paragraph()
        p.paragraph_format.space_before = Pt(150)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_run_font(p.add_run(m.get("doc_type", "보고서")), size=12, color=BLUE, bold=True)
        p = d.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.line_spacing = 1.2
        set_run_font(p.add_run(m["title"]), size=26, bold=True, color=NAVY)
        if m.get("subtitle"):
            p = d.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_run_font(p.add_run(m["subtitle"]), size=14, color=GRAY)
        line = d.add_paragraph()
        line.paragraph_format.space_before = Pt(18)
        para_border(line, ("bottom",), color=CALLOUT_LINE, sz=12)
        spacer = d.add_paragraph()
        spacer.paragraph_format.space_before = Pt(190)
        info = [("소속", m.get("department", "")), ("작성자", m.get("author", "")),
                ("작성일", m.get("date") or date.today().strftime("%Y. %m. %d."))]
        for k, v in info:
            if not v:
                continue
            p = d.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_after = Pt(2)
            set_run_font(p.add_run(f"{k}  "), size=12, color=GRAY)
            set_run_font(p.add_run(v), size=12, bold=True)

    # ---------- body section with header/footer (R8)
    def body_section(self):
        d = self.doc
        sec = d.add_section(WD_SECTION.NEW_PAGE)
        for part in (sec.header, sec.footer):
            part.is_linked_to_previous = False
        # cover section keeps an empty header/footer
        # page numbers restart at 1 for the body
        pg = OxmlElement("w:pgNumType")
        pg.set(qn("w:start"), "1")
        sec._sectPr.append(pg)

        hp = sec.header.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        set_run_font(hp.add_run(self.meta.get("short_title", self.meta["title"])), size=9, color=GRAY)
        para_border(hp, ("bottom",), color="BFBFBF", sz=4, space=2)

        fp = sec.footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        fp._p.append(text_run("- ", make_rpr(color="7F7F7F", size=9)))
        add_field(fp, "PAGE", "1", size=9, color="7F7F7F")
        fp._p.append(text_run(" / ", make_rpr(color="7F7F7F", size=9)))
        add_field(fp, "SECTIONPAGES", "1", size=9, color="7F7F7F")
        fp._p.append(text_run(" -", make_rpr(color="7F7F7F", size=9)))

    # ---------- TOC (R3)
    def toc(self):
        d = self.doc
        p = d.add_paragraph()
        set_run_font(p.add_run("목 차"), size=18, bold=True, color=NAVY)
        para_border(p, ("bottom",), color=HEADER_FILL, sz=8)
        p.paragraph_format.space_after = Pt(12)
        p = d.add_paragraph()
        add_field(p, 'TOC \\o "1-2" \\h \\z \\u', "목차를 표시하려면 Word에서 [필드 업데이트](F9)를 실행하세요.")
        p.add_run().add_break(WD_BREAK.PAGE)

    # ---------- blocks
    def heading(self, text, level):
        p = self.doc.add_heading(level=level)
        set_run_font(p.add_run(text), size=16 if level == 1 else 13, bold=True, color=NAVY if level == 1 else BLUE)
        if level == 1:
            para_border(p, ("bottom",), color="BFD0EA", sz=6, space=3)
        return p

    def paragraph(self, text):
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        self.inline(p, text)
        return p

    def bullets(self, items):
        for it in items:
            p = self.doc.add_paragraph(style="List Bullet")
            self.inline(p, it)

    def callout(self, title, items):
        """R13: 핵심 요약 강조 박스 (음영 + 왼쪽 굵은 선)."""
        paras = []
        p = self.doc.add_paragraph()
        set_run_font(p.add_run(title), size=12, bold=True, color=NAVY)
        paras.append(p)
        for it in items:
            p = self.doc.add_paragraph()
            set_run_font(p.add_run("■ "), size=9, color=BLUE)
            self.inline(p, it)
            p.paragraph_format.space_after = Pt(3)
            paras.append(p)
        for i, p in enumerate(paras):
            ppr = p._p.get_or_add_pPr()
            bdr = OxmlElement("w:pBdr")
            for side in ("top", "left", "bottom", "right"):
                b = OxmlElement(f"w:{side}")
                b.set(qn("w:val"), "single")
                b.set(qn("w:sz"), "24" if side == "left" else "4")
                b.set(qn("w:space"), "8" if side in ("left", "right") else "4")
                b.set(qn("w:color"), CALLOUT_LINE if side == "left" else CALLOUT_FILL)
                bdr.append(b)
            ppr.append(bdr)
            shade(ppr, CALLOUT_FILL)
            p.paragraph_format.left_indent = Cm(0.3)
            p.paragraph_format.right_indent = Cm(0.3)
        self.doc.add_paragraph().paragraph_format.space_after = Pt(0)

    def caption(self, label, n, text, before=False):
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(4 if before else 2)
        p.paragraph_format.space_after = Pt(4 if before else 12)
        if before:
            p.paragraph_format.keep_with_next = True
        set_run_font(p.add_run(f"{label} {n}. "), size=10, bold=True, color=NAVY)
        self.inline(p, text, size=10)
        return p

    def table(self, b):
        """R9: 표 위 캡션, 남색 머리행(반복), 줄무늬 행."""
        n = self.tbl_no.get(b.get("id")) or (len([1 for _ in self.doc.tables]) + 1)
        self.caption("표", n, b["caption"], before=True)
        cols, rows = b["columns"], b["rows"]
        total_cm = 16.0
        widths = b.get("widths") or [1] * len(cols)
        widths = [total_cm * w / sum(widths) for w in widths]
        t = self.doc.add_table(rows=1 + len(rows), cols=len(cols))
        t.style = "Table Grid"
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        t.autofit = False
        align = b.get("align") or ["left"] * len(cols)
        for ri, row in enumerate([cols] + rows):
            tr = t.rows[ri]
            if ri == 0:
                trpr = tr._tr.get_or_add_trPr()
                h = OxmlElement("w:tblHeader")
                h.set(qn("w:val"), "true")
                trpr.append(h)
            cant = OxmlElement("w:cantSplit")
            tr._tr.get_or_add_trPr().append(cant)
            for ci, val in enumerate(row):
                cell = tr.cells[ci]
                cell.width = Cm(widths[ci])
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                p = cell.paragraphs[0]
                p.paragraph_format.space_after = Pt(0)
                # 표가 짧으면(<=12행) 한 페이지에 함께 두고, 길면 머리행만 다음 행과 붙인다
                if ri < len(rows) and (len(rows) <= 12 or ri == 0):
                    p.paragraph_format.keep_with_next = True
                p.paragraph_format.line_spacing = 1.2
                if ri == 0:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    set_run_font(p.add_run(str(val)), size=9.5, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))
                    shade(cell._tc.get_or_add_tcPr(), HEADER_FILL)
                else:
                    p.alignment = {"center": WD_ALIGN_PARAGRAPH.CENTER, "right": WD_ALIGN_PARAGRAPH.RIGHT}.get(align[ci], WD_ALIGN_PARAGRAPH.LEFT)
                    self.inline(p, str(val), size=9.5)
                    if ri % 2 == 0:
                        shade(cell._tc.get_or_add_tcPr(), ZEBRA_FILL)
        if b.get("note"):
            p = self.doc.add_paragraph()
            p.paragraph_format.space_before = Pt(2)
            self.inline(p, "※ " + b["note"], size=9, color=GRAY)
        self.doc.add_paragraph().paragraph_format.space_after = Pt(0)

    def figure(self, b, img_path, width_cm=15.0):
        """R10: 가운데 정렬 그림 + 그림 아래 캡션."""
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.keep_with_next = True
        p.paragraph_format.space_before = Pt(6)
        p.add_run().add_picture(img_path, width=Cm(b.get("width_cm", width_cm)))
        cap = self.caption("그림", self.fig_no.get(b.get("id"), 0), b["caption"])
        if b.get("note"):
            cap.paragraph_format.space_after = Pt(2)
            cap.paragraph_format.keep_with_next = True  # 그림·캡션·주석을 같은 쪽에
            p = self.doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_after = Pt(12)
            self.inline(p, "※ " + b["note"], size=9, color=GRAY)

    # ---------- charts (matplotlib)
    def _mpl(self):
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        from matplotlib import font_manager
        names = {f.name for f in font_manager.fontManager.ttflist}
        for cand in ("Malgun Gothic", "NanumGothic", "AppleGothic", "Noto Sans CJK KR"):
            if cand in names:
                plt.rcParams["font.family"] = cand
                break
        plt.rcParams["axes.unicode_minus"] = False
        return plt

    def chart(self, b):
        plt = self._mpl()
        cats = b["categories"]
        series = b["series"]
        fig, ax = plt.subplots(figsize=(7.5, 3.3), dpi=200)
        colors = ["#1F3864", "#5B9BD5", "#A5A5A5", "#ED7D31"]
        n = len(series)
        w = 0.8 / n
        for i, s in enumerate(series):
            xs = [x + (i - (n - 1) / 2) * w for x in range(len(cats))]
            bars = ax.bar(xs, s["values"], w * 0.92, label=s["name"], color=colors[i % 4])
            labels = s.get("labels") or [f"{v:g}{b.get('unit', '')}" for v in s["values"]]
            for bar, lab in zip(bars, labels):
                ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + max(max(x["values"]) for x in series) * 0.015,
                        lab, ha="center", va="bottom", fontsize=9)
        ax.set_xticks(range(len(cats)))
        ax.set_xticklabels(cats, fontsize=10)
        ax.set_ylabel(b.get("ylabel", ""), fontsize=10)
        ax.set_ylim(0, b.get("ymax", max(max(s["values"]) for s in series) * 1.18))
        for sp in ("top", "right"):
            ax.spines[sp].set_visible(False)
        ax.grid(axis="y", alpha=0.3)
        ax.set_axisbelow(True)
        ax.legend(frameon=False, fontsize=9, loc="upper right")
        fig.tight_layout()
        path = os.path.join(self.tmp, f"chart_{len(os.listdir(self.tmp))}.png")
        fig.savefig(path)
        plt.close(fig)
        self.figure(b, path)

    def flow(self, b):
        plt = self._mpl()
        from matplotlib.patches import FancyBboxPatch
        steps = b["steps"]
        n = len(steps)
        fig, ax = plt.subplots(figsize=(8, 2.3), dpi=200)
        ax.set_xlim(0, n)
        ax.set_ylim(0, 1)
        ax.axis("off")
        palette = ["#1F3864", "#2E5597", "#2E75B6", "#5B9BD5", "#8FB4E3", "#B4CBEA"]
        for i, s in enumerate(steps):
            color = palette[min(i, len(palette) - 1)]
            box = FancyBboxPatch((i + 0.06, 0.18), 0.78, 0.64, boxstyle="round,pad=0.02,rounding_size=0.06",
                                 fc=color, ec="none")
            ax.add_patch(box)
            dark = i < 4
            ax.text(i + 0.45, 0.62, s["title"], ha="center", va="center", fontsize=10.5, weight="bold",
                    color="white" if dark else "#1F3864")
            ax.text(i + 0.45, 0.38, s.get("desc", ""), ha="center", va="center", fontsize=7.5,
                    color="white" if dark else "#1F3864", linespacing=1.3)
            if i < n - 1:
                ax.annotate("", xy=(i + 1.05, 0.5), xytext=(i + 0.86, 0.5),
                            arrowprops=dict(arrowstyle="-|>", color="#7F7F7F", lw=1.4))
        fig.tight_layout(pad=0.2)
        path = os.path.join(self.tmp, f"flow_{len(os.listdir(self.tmp))}.png")
        fig.savefig(path)
        plt.close(fig)
        self.figure(b, path, width_cm=16.0)

    def block(self, b):
        t = b["type"]
        if t == "paragraph":
            self.paragraph(b["text"])
        elif t == "bullets":
            self.bullets(b["items"])
        elif t == "table":
            self.table(b)
        elif t == "chart":
            self.chart(b)
        elif t == "flow":
            self.flow(b)
        elif t == "image":
            self.figure(b, os.path.join(self.base_dir, b["path"]))
        elif t == "callout":
            self.callout(b.get("title", "핵심 정리"), b["items"])
        else:
            raise BuildError(f"알 수 없는 블록 type: {t}")

    # ---------- references (R12)
    def references(self):
        self.heading("참고문헌", 1)
        for n, key in enumerate(self.cite_order, 1):
            r = self.refs[key]
            p = self.doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(0.9)
            p.paragraph_format.first_line_indent = Cm(-0.9)
            p.paragraph_format.space_after = Pt(4)
            p.paragraph_format.line_spacing = 1.25
            set_run_font(p.add_run(f"[{n}] "), size=9.5, bold=True, color=NAVY)
            body = f"{r['authors']}, “{r['title']},” {r['source']}, {r['year']}. "
            set_run_font(p.add_run(body), size=9.5)
            add_external_link(p, r.get("link_text") or r["url"], r["url"], size=9.5)
            add_bookmark(p, f"ref_{n}")

    # ---------- main
    def build(self, out):
        self.setup()
        self.cover()
        self.body_section()
        self.toc()
        if self.c.get("summary"):
            self.heading("요약", 1)
            self.callout("핵심 요약", self.c["summary"])
        for i, sec in enumerate(self.c["sections"], 1):
            self.heading(f"{i}. {sec['heading']}", 1)
            for b in sec.get("blocks", []):
                self.block(b)
            for j, sub in enumerate(sec.get("subsections", []), 1):
                self.heading(f"{i}.{j} {sub['heading']}", 2)
                for b in sub.get("blocks", []):
                    self.block(b)
        self.references()
        self.doc.save(out)
        return out


def default_filename(meta):
    topic = meta.get("file_topic") or re.sub(r"[^\w가-힣]+", "", meta["title"])[:20]
    d = re.sub(r"\D", "", meta.get("date", "")) or date.today().strftime("%Y%m%d")
    return f"{topic}_보고서_{d}.docx"


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("content", help="content.json")
    ap.add_argument("-o", "--out", help="출력 .docx (기본: {주제}_보고서_{YYYYMMDD}.docx, R15)")
    a = ap.parse_args()
    with open(a.content, encoding="utf-8") as f:
        content = json.load(f)
    out = a.out or os.path.join(os.path.dirname(os.path.abspath(a.content)), default_filename(content["meta"]))
    try:
        ReportBuilder(content, os.path.dirname(os.path.abspath(a.content))).build(out)
    except BuildError as e:
        print(f"[BUILD ERROR] {e}", file=sys.stderr)
        sys.exit(2)
    print(out)


if __name__ == "__main__":
    main()
