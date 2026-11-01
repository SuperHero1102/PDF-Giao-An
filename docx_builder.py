# -*- coding: utf-8 -*-
"""
Dựng file Word (.docx) từ danh sách "khối" (blocks) đã nhận dạng.

- Chữ: font Times New Roman, cỡ 13 (Unicode, không lỗi dấu tiếng Việt).
- Công thức: viết bằng LaTeX trong $...$, được Pandoc đổi sang công thức Word
  thật (OMML) -> bấm vào trong Word là sửa được.
- Hình: cắt từ ảnh trang gốc theo vùng (bbox, tính theo % trang).

Các loại khối (type):
  paragraph   : {text, align, indent, first_line}
  columns     : {cells:[...], aligns:[...], widths:[...]}   (bảng không viền)
  boxed       : {text}                                       (chữ trong khung)
  figure      : {bbox:[x1,y1,x2,y2]}
  text_figure : {text, bbox, side:"right"|"left"}            (chữ cạnh hình)
  formula     : {latex}                                      (công thức riêng 1 dòng)
Trong text: **đậm**, *nghiêng*, $công thức$.
"""
import copy
import os
import re
import subprocess
import tempfile
import zipfile
from concurrent.futures import ThreadPoolExecutor

from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from docx.oxml import parse_xml
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from lxml import etree
from PIL import Image

M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
FONT = "Times New Roman"
FONT_PT = 13
PAGE_W_CM = 21.0
CONTENT_W_CM = 17.0

TOKEN = re.compile(r"(\$[^$]+\$|\*\*.+?\*\*|(?<!\*)\*[^*\s][^*]*?\*(?!\*))")
MATH_ONLY = re.compile(r"\$([^$]+)\$")

ALIGN = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}


# --------------------------------------------------------------------------
# Công thức: LaTeX -> OMML bằng Pandoc
# --------------------------------------------------------------------------
def _latex_to_omml_xml(latex):
    """Trả về phần tử lxml m:oMath, hoặc None nếu Pandoc không đổi được."""
    with tempfile.TemporaryDirectory() as d:
        md, out = os.path.join(d, "a.md"), os.path.join(d, "a.docx")
        with open(md, "w", encoding="utf-8") as f:
            f.write("$" + latex + "$\n")
        r = subprocess.run(["pandoc", md, "-o", out], capture_output=True, timeout=60)
        if r.returncode != 0 or not os.path.exists(out):
            return None
        xml = zipfile.ZipFile(out).read("word/document.xml")
    root = etree.fromstring(xml)
    return root.find(".//{%s}oMath" % M_NS)


def _batch_to_omml(latex_list):
    """Đổi NHIỀU công thức bằng một lần gọi Pandoc (nhanh hơn nhiều trên máy chủ yếu).
    Trả về list m:oMath cùng thứ tự, hoặc None nếu số lượng không khớp (khi đó đổi lẻ từng cái)."""
    with tempfile.TemporaryDirectory() as d:
        md, out = os.path.join(d, "a.md"), os.path.join(d, "a.docx")
        with open(md, "w", encoding="utf-8") as f:
            f.write("\n\n".join("$" + l + "$" for l in latex_list) + "\n")
        r = subprocess.run(["pandoc", md, "-o", out], capture_output=True, timeout=300)
        if r.returncode != 0 or not os.path.exists(out):
            return None
        xml = zipfile.ZipFile(out).read("word/document.xml")
    root = etree.fromstring(xml)
    paras = root.findall(".//{%s}p" % W_NS)
    found = [p.find(".//{%s}oMath" % M_NS) for p in paras]
    found = [x for x in found if x is not None]
    return found if len(found) == len(latex_list) else None


class Formulas:
    """Đổi trước tất cả công thức và nhớ kết quả."""

    def __init__(self):
        self.cache = {}

    def prepare(self, latex_list):
        todo = sorted(set(l.strip() for l in latex_list if l.strip()) - set(self.cache))
        for i in range(0, len(todo), 40):
            chunk = todo[i:i + 40]
            res = None
            try:
                res = _batch_to_omml(chunk)
            except Exception:  # noqa
                res = None
            if res is not None:
                self.cache.update(zip(chunk, res))
                continue
            # lô có công thức lỗi -> đổi lẻ từng công thức để chỉ công thức lỗi bị tô vàng
            with ThreadPoolExecutor(max_workers=4) as ex:
                for latex, el in zip(chunk, ex.map(_latex_to_omml_xml, chunk)):
                    self.cache[latex] = el

    def get(self, latex):
        el = self.cache.get(latex.strip())
        return copy.deepcopy(el) if el is not None else None


def _collect_latex(blocks):
    out = []
    for b in blocks:
        texts = []
        t = b.get("type")
        if t in ("paragraph", "boxed", "text_figure"):
            texts.append(b.get("text", ""))
        elif t == "columns":
            texts.extend(b.get("cells", []))
        elif t == "formula":
            out.append(b.get("latex", ""))
        for s in texts:
            out.extend(MATH_ONLY.findall(s or ""))
    return out


# --------------------------------------------------------------------------
# Hàm dựng từng phần
# --------------------------------------------------------------------------
class Builder:
    def __init__(self, page_image_fn, sizes=None):
        """page_image_fn(page_no) -> đường dẫn ảnh PNG của trang (để cắt hình).
        sizes: {số_trang: (rộng_pt, cao_pt)} khổ giấy từng trang của PDF gốc (đã tính xoay)."""
        self.sizes = sizes or {}
        self.page_w_cm, self.content_w_cm = PAGE_W_CM, CONTENT_W_CM
        self.page_image_fn = page_image_fn
        self.doc = Document()
        self.warnings = []
        self.formulas = Formulas()
        self._setup()

    def _setup(self):
        s = self.doc.sections[0]
        s.page_width, s.page_height = Cm(21), Cm(29.7)
        s.left_margin = s.right_margin = Cm(2)
        s.top_margin, s.bottom_margin = Cm(1.8), Cm(1.5)
        st = self.doc.styles["Normal"]
        st.font.name, st.font.size = FONT, Pt(FONT_PT)
        st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
        st.paragraph_format.space_after = Pt(2)
        st.paragraph_format.space_before = Pt(0)
        st.paragraph_format.line_spacing = 1.05

    # ---- chữ + công thức trong một đoạn --------------------------------
    def _run(self, p, text, bold=False, italic=False, highlight=False):
        r = p.add_run(text)
        r.bold, r.italic = bold, italic
        r.font.name, r.font.size = FONT, Pt(FONT_PT)
        r._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
        if highlight:
            r.font.highlight_color = WD_COLOR_INDEX.YELLOW
        return r

    def _math(self, p, latex):
        el = self.formulas.get(latex)
        if el is None:
            # Không đổi được: giữ LaTeX thô, tô vàng để người dùng thấy và sửa
            self.warnings.append("Không đổi được công thức: " + latex)
            self._run(p, latex, highlight=True)
        else:
            p._p.append(el)

    def fill(self, p, text):
        for part in TOKEN.split(text or ""):
            if not part:
                continue
            if part.startswith("$") and part.endswith("$") and len(part) > 2:
                self._math(p, part[1:-1])
            elif part.startswith("**") and part.endswith("**") and len(part) > 4:
                self._run(p, part[2:-2], bold=True)
            elif part.startswith("*") and part.endswith("*") and len(part) > 2:
                self._run(p, part[1:-1], italic=True)
            else:
                self._run(p, part)

    def para(self, container, text, align="left", indent=0.0, first_line=False,
             before=0, after=None):
        p = container.add_paragraph()
        p.alignment = ALIGN.get(align, WD_ALIGN_PARAGRAPH.LEFT)
        pf = p.paragraph_format
        if indent:
            pf.left_indent = Cm(float(indent))
        if first_line:
            pf.first_line_indent = Cm(1.0)
        if before:
            pf.space_before = Pt(before)
        if after is not None:
            pf.space_after = Pt(after)
        self.fill(p, text)
        return p

    # ---- bảng không viền / khung -----------------------------------------
    @staticmethod
    def _no_borders(tbl):
        borders = "".join('<w:%s w:val="nil"/>' % s for s in
                          ("top", "left", "bottom", "right", "insideH", "insideV"))
        tbl._tbl.tblPr.append(parse_xml(
            '<w:tblBorders xmlns:w="%s">%s</w:tblBorders>' % (W_NS, borders)))

    @staticmethod
    def _widths(tbl, widths):
        tbl.autofit = False
        for col, w in zip(tbl.columns, widths):  # cả khung lưới cột (để mọi phần mềm đọc đúng bề rộng)
            col.width = Cm(w)
        for row in tbl.rows:
            for c, w in zip(row.cells, widths):
                c.width = Cm(w)

    def _box(self, container, text):
        """Chữ trong khung viền (vd: ĐỀ CHÍNH THỨC)."""
        w = min(max(len(text) * 0.33 + 1.0, 3.0), 12.0)
        box = container.add_table(rows=1, cols=1)
        box.style = "Table Grid"
        box.alignment = WD_TABLE_ALIGNMENT.CENTER
        self._widths(box, [w])
        cell = box.rows[0].cells[0]
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        self.fill(p, text if "**" in text else "**" + text + "**")
        container.add_paragraph().paragraph_format.space_after = Pt(2)

    def _cell_lines(self, cell, text, align):
        lines = (text or "").split("\n")
        first = True
        for ln in lines:
            if ln.startswith("[box]"):
                self._box(cell, ln[5:].strip())
                first = False
                continue
            if first:
                p = cell.paragraphs[0]
                p.alignment = ALIGN.get(align, WD_ALIGN_PARAGRAPH.LEFT)
                self.fill(p, ln)
                first = False
            else:
                self.para(cell, ln, align=align)

    # ---- hình cắt từ trang gốc -------------------------------------------
    def _crop(self, page_no, bbox):
        path = self.page_image_fn(page_no)
        im = Image.open(path).convert("RGB")
        W, H = im.size
        x1, y1, x2, y2 = [float(v) for v in bbox]
        mx, my = 0.01 * W, 0.006 * H  # lề nhỏ để không lẹm mép hình
        box = (max(0, x1 / 100 * W - mx), max(0, y1 / 100 * H - my),
               min(W, x2 / 100 * W + mx), min(H, y2 / 100 * H + my))
        crop = im.crop(tuple(int(v) for v in box))
        tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
        crop.save(tmp.name)
        return tmp.name, (box[2] - box[0]) / W  # tỉ lệ rộng so với trang

    def _picture(self, p, page_no, bbox, max_cm):
        try:
            path, frac = self._crop(page_no, bbox)
            width = min(frac * self.page_w_cm, max_cm)
            p.add_run().add_picture(path, width=Cm(width))
            return frac * self.page_w_cm
        except Exception as e:  # noqa
            self.warnings.append("Không cắt được hình (trang %s): %s" % (page_no, e))
            self._run(p, "[Không cắt được hình]", highlight=True)
            return 6.0

    # ---- từng loại khối ----------------------------------------------------
    def block(self, b, page_no):
        t = b.get("type", "paragraph")
        if t == "paragraph":
            self.para(self.doc, b.get("text", ""), b.get("align", "left"),
                      b.get("indent", 0), bool(b.get("first_line")),
                      before=b.get("space_before", 0))
        elif t == "boxed":
            self._box(self.doc, b.get("text", ""))
        elif t == "formula":
            p = self.doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            self._math(p, b.get("latex", ""))
        elif t == "columns":
            cells = b.get("cells", [])
            n = max(len(cells), 1)
            widths = b.get("widths") or [100 / n] * n
            tot = float(sum(widths)) or 1.0
            widths_cm = [self.content_w_cm * float(w) / tot for w in widths]
            aligns = b.get("aligns") or ["left"] * n
            tbl = self.doc.add_table(rows=1, cols=n)
            self._no_borders(tbl)
            self._widths(tbl, widths_cm)
            for i, c in enumerate(cells):
                self._cell_lines(tbl.rows[0].cells[i], c, aligns[i] if i < len(aligns) else "left")
        elif t == "figure":
            p = self.doc.add_paragraph()
            p.alignment = ALIGN.get(b.get("align", "center"))
            p.paragraph_format.space_before = Pt(6)
            self._picture(p, page_no, b.get("bbox", [0, 0, 100, 100]), self.content_w_cm)
        elif t == "text_figure":
            side = b.get("side", "right")
            bbox = b.get("bbox", [50, 0, 100, 20])
            frac_cm = min(max((float(bbox[2]) - float(bbox[0])) / 100 * self.page_w_cm, 4.0), 9.5)
            text_cm = self.content_w_cm - frac_cm - 0.2
            tbl = self.doc.add_table(rows=1, cols=2)
            self._no_borders(tbl)
            order = [text_cm, frac_cm] if side == "right" else [frac_cm, text_cm]
            self._widths(tbl, order)
            tcell = tbl.rows[0].cells[0 if side == "right" else 1]
            fcell = tbl.rows[0].cells[1 if side == "right" else 0]
            self._cell_lines(tcell, b.get("text", ""), b.get("align", "left"))
            fp = fcell.paragraphs[0]
            fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            self._picture(fp, page_no, bbox, frac_cm - 0.2)
        else:
            self.warnings.append("Loại khối lạ, bỏ qua: " + str(t))

    def _start_page(self, n, first):
        """Mỗi trang PDF = một "section" Word có đúng khổ giấy và hướng (dọc/ngang) của trang gốc."""
        sec = self.doc.sections[0] if first else self.doc.add_section(WD_SECTION.NEW_PAGE)
        w_pt, h_pt = self.sizes.get(n) or (595.0, 842.0)
        w_cm, h_cm = w_pt / 72 * 2.54, h_pt / 72 * 2.54
        for ref_w, ref_h in ((21.0, 29.7), (29.7, 21.0)):  # gần A4 thì lấy đúng A4
            if abs(w_cm - ref_w) < 0.6 and abs(h_cm - ref_h) < 0.6:
                w_cm, h_cm = ref_w, ref_h
        sec.page_width, sec.page_height = Cm(w_cm), Cm(h_cm)
        sec.orientation = WD_ORIENT.LANDSCAPE if w_cm > h_cm else WD_ORIENT.PORTRAIT
        margin = 2.0 if w_cm >= 18 else max(1.0, w_cm * 0.08)
        sec.left_margin = sec.right_margin = Cm(margin)
        sec.top_margin, sec.bottom_margin = Cm(1.8), Cm(1.5)
        self.page_w_cm, self.content_w_cm = w_cm, w_cm - 2 * margin

    # ---- toàn bộ tài liệu ---------------------------------------------------
    def build(self, pages, out_path):
        """pages: {số_trang(int): [khối,...]} theo thứ tự trang."""
        all_blocks = [b for n in sorted(pages) for b in pages[n]]
        self.formulas.prepare(_collect_latex(all_blocks))
        first = True
        for n in sorted(pages):
            self._start_page(n, first)
            first = False
            for b in pages[n]:
                self.block(b, n)
        self.doc.save(out_path)
        return self.warnings
