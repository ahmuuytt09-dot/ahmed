"""Render the KAZ Management & Staff EHS Induction to PPTX (presentation) and PDF (distribution).

Both outputs are produced from deck_content.py through a small canvas abstraction, so the two
formats stay visually consistent. Canvas units are points on a 960 x 540 (16:9) page —
the same geometry the original file used.
"""
import os, re, html as _html
from PIL import Image

import deck_content as C

CANVAS_W, CANVAS_H = 960, 540
ASSETS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets")
IMG_DIR = os.path.join(ASSETS, "img")
LOGO_DIR = os.path.join(ASSETS, "logos")
CROP_DIR = os.path.join(ASSETS, "crop")

# ------------------------------------------------------------------ brand
TEAL = (0x00, 0x99, 0x99)
TEAL_HEX = "009999"
PURPLE = (0x52, 0x17, 0x7A)
NAVY = (0x12, 0x26, 0x3A)
INK = (0x33, 0x3F, 0x4C)
GREY = (0x6B, 0x7A, 0x88)
LIGHT = (0xF4, 0xF7, 0xFA)
WHITE = (0xFF, 0xFF, 0xFF)
RED = (0xC0, 0x00, 0x00)
AMBER = (0xE8, 0xA3, 0x3D)


# ------------------------------------------------------------------ image helpers
def crop_to(img_path, target_w, target_h):
    """Center-crop (no distortion) an image to the target aspect ratio, cached on disk.

    Images that contain an alpha channel are kept transparent (the PDF backend paints the
    background itself); opaque photographs are flattened to RGB.
    """
    os.makedirs(CROP_DIR, exist_ok=True)
    base = os.path.splitext(os.path.basename(img_path))[0]
    has_alpha = Image.open(img_path).mode in ("RGBA", "LA", "P")
    tag = "rgba" if has_alpha else "rgb"
    out = os.path.join(CROP_DIR, f"{base}_{int(target_w)}x{int(target_h)}_{tag}.png")
    if os.path.exists(out):
        return out
    im = Image.open(img_path).convert("RGBA" if has_alpha else "RGB")
    target_ratio = target_w / target_h
    w, h = im.size
    ratio = w / h
    if ratio > target_ratio:                     # too wide -> crop sides
        new_w = int(h * target_ratio)
        off = (w - new_w) // 2
        im = im.crop((off, 0, off + new_w, h))
    elif ratio < target_ratio:                   # too tall -> crop top/bottom
        new_h = int(w / target_ratio)
        off = (h - new_h) // 2
        im = im.crop((0, off, w, off + new_h))
    if has_alpha:
        if im.width > 900:
            im = im.resize((900, int(im.height * 900 / im.width)), Image.LANCZOS)
        im.save(out, optimize=True)
    else:
        if im.width > 1700:
            im = im.resize((1700, int(im.height * 1700 / im.width)), Image.LANCZOS)
        im.convert("RGB").save(out.replace(".png", ".jpg"), "JPEG", quality=84, optimize=True, progressive=True)
        out = out.replace(".png", ".jpg")
    return out


def img(name):
    return os.path.join(IMG_DIR, name)


def logo(name):
    return os.path.join(LOGO_DIR, name)


# ------------------------------------------------------------------ canvas base
class Canvas:
    """Primitive drawing surface shared by the PPTX and PDF backends."""

    def rect(self, x, y, w, h, fill=None, line=None, line_w=0.75):
        raise NotImplementedError

    def text(self, x, y, w, h, runs, size=14, color=INK, bold=False, align="left",
             valign="top", spacing=1.0, italic=False):
        """runs: str or list of (text, {bold/color/size/italic}) tuples."""
        raise NotImplementedError

    def image(self, path, x, y, w, h):
        raise NotImplementedError

    def line(self, x1, y1, x2, y2, color=TEAL, width=1.0):
        raise NotImplementedError


# ------------------------------------------------------------------ PPTX backend
class PptxCanvas(Canvas):
    def __init__(self, slide, prs):
        self.slide, self.prs = slide, prs

    def _rgb(self, c):
        from pptx.dml.color import RGBColor
        return RGBColor(*c)

    def rect(self, x, y, w, h, fill=None, line=None, line_w=0.75):
        from pptx.util import Pt
        from pptx.enum.shapes import MSO_SHAPE
        sh = self.slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Pt(x), Pt(y), Pt(w), Pt(h))
        sh.shadow.inherit = False
        if fill is None:
            sh.fill.background()
        else:
            sh.fill.solid(); sh.fill.fore_color.rgb = self._rgb(fill)
        if line is None:
            sh.line.fill.background()
        else:
            sh.line.color.rgb = self._rgb(line); sh.line.width = Pt(line_w)
        sh.text_frame.text = ""
        return sh

    def text(self, x, y, w, h, runs, size=14, color=INK, bold=False, align="left",
             valign="top", spacing=1.0, italic=False):
        from pptx.util import Pt
        from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
        tb = self.slide.shapes.add_textbox(Pt(x), Pt(y), Pt(w), Pt(h))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = {"top": MSO_ANCHOR.TOP, "middle": MSO_ANCHOR.MIDDLE,
                              "bottom": MSO_ANCHOR.BOTTOM}[valign]
        items = runs if isinstance(runs, list) else [(runs, {})]
        first = True
        for text, style in items:
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            p.alignment = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER,
                           "right": PP_ALIGN.RIGHT}[style.get("align", align)]
            p.line_spacing = style.get("spacing", spacing)
            if style.get("space_before"):
                p.space_before = Pt(style["space_before"])
            r = p.add_run(); r.text = text
            f = r.font
            f.size = Pt(style.get("size", size))
            f.bold = style.get("bold", bold)
            f.italic = style.get("italic", italic)
            f.name = "Calibri"
            f.color.rgb = self._rgb(style.get("color", color))
        return tb

    def image(self, path, x, y, w, h):
        from pptx.util import Pt
        return self.slide.shapes.add_picture(crop_to(path, w, h), Pt(x), Pt(y), Pt(w), Pt(h))

    def line(self, x1, y1, x2, y2, color=TEAL, width=1.0):
        from pptx.util import Pt
        from pptx.enum.shapes import MSO_CONNECTOR
        cn = self.slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Pt(x1), Pt(y1), Pt(x2), Pt(y2))
        cn.line.color.rgb = self._rgb(color); cn.line.width = Pt(width)
        return cn


# ------------------------------------------------------------------ PDF backend
FONT_CSS = """
@font-face {font-family: dm; src: url(DejaVuSans.ttf); font-weight: normal;}
@font-face {font-family: dm; src: url(DejaVuSans-Bold.ttf); font-weight: bold;}
@font-face {font-family: dms; src: url(DejaVuSerif.ttf); font-weight: normal;}
@font-face {font-family: dms; src: url(DejaVuSerif-Bold.ttf); font-weight: bold;}
"""


def _font_archive():
    """Zip archive of the TTFs used by the PDF backend (numbers/bold render correctly)."""
    global _ARCHIVE
    try:
        return _ARCHIVE
    except NameError:
        pass
    import zipfile
    src = "/usr/share/fonts/truetype/dejavu"
    path = os.path.join(CROP_DIR, "fonts.zip")
    os.makedirs(CROP_DIR, exist_ok=True)
    if not os.path.exists(path):
        with zipfile.ZipFile(path, "w") as z:
            for fn in ("DejaVuSans.ttf", "DejaVuSans-Bold.ttf",
                       "DejaVuSerif.ttf", "DejaVuSerif-Bold.ttf"):
                z.write(os.path.join(src, fn), fn)
    _ARCHIVE = zipfile.ZipFile(path)
    return _ARCHIVE


class PdfCanvas(Canvas):
    """Draws on a PyMuPDF page using the DejaVu font family through insert_htmlbox."""

    def __init__(self, page):
        self.page = page
        self.css = FONT_CSS
        self.archive = _font_archive()

    def _col(self, c):
        return tuple(v / 255 for v in c)

    def rect(self, x, y, w, h, fill=None, line=None, line_w=0.75):
        r = __import__("pymupdf").Rect(x, y, x + w, y + h)
        self.page.draw_rect(r, color=self._col(line) if line else None,
                            fill=self._col(fill) if fill else None, width=line_w)

    def _runs_html(self, runs, size, color, bold, italic, align, spacing):
        items = runs if isinstance(runs, list) else [(runs, {})]
        parts = []
        for text, style in items:
            t = _html.escape(text).replace("\n", "<br/>")
            t = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", t)
            s = f'font-size:{style.get("size", size)}pt;'
            if style.get("bold", bold): s += "font-weight:bold;"
            if style.get("italic", italic): s += "font-style:italic;"
            col = style.get("color", color)
            s += f'color:#{"".join(f"{v:02X}" for v in col)};'
            parts.append(f'<span style="{s}">{t}</span>')
        return (f'<div style="font-family:dm;line-height:{spacing};text-align:{align}">'
                + "".join(parts) + "</div>")

    def text(self, x, y, w, h, runs, size=14, color=INK, bold=False, align="left",
             valign="top", spacing=1.0, italic=False):
        import pymupdf
        html = self._runs_html(runs, size, color, bold, italic, align, spacing)
        box = pymupdf.Rect(x, y, x + w, y + h)
        for scale in (1.0, 0.96, 0.92, 0.88, 0.84, 0.8, 0.76, 0.72):
            res = self.page.insert_htmlbox(box, html, css=self.css, archive=self.archive,
                                           scale_low=scale)
            if res and res[0] >= 0:
                return
        self.page.insert_htmlbox(box, html, css=self.css, archive=self.archive)

    def image(self, path, x, y, w, h):
        self.page.insert_image(__import__("pymupdf").Rect(x, y, x + w, y + h), filename=crop_to(path, w, h))

    def line(self, x1, y1, x2, y2, color=TEAL, width=1.0):
        self.page.draw_line(__import__("pymupdf").Point(x1, y1), __import__("pymupdf").Point(x2, y2),
                            color=self._col(color), width=width)


# ------------------------------------------------------------------ slide renderer
FOOTER = f"{C.DOC_NO} · {C.REV} · Management & Staff EHS Induction · Restricted"


def _runs_for_bullet(text, level, size):
    """Return (runs, indent, effective size) for one content line."""
    if level == 0:
        return ([("▪  ", {"color": TEAL, "bold": True, "size": size}),
                 (text, {"color": INK, "size": size})], 0, size)
    if level == 1:
        return ([("•  ", {"color": GREY, "bold": True, "size": size - 1}),
                 (text, {"color": INK, "size": size - 1})], 16, size - 1)
    return ([("", {}), (text, {"color": PURPLE, "bold": True, "size": size - 1})], 0, size - 1)


def _layout_bullets(bullets, size, width, x=72, y0=132, leading=1.30, gap=7.0, gap_sub=4.0):
    """Estimate the laid-out lines: returns (entries, bottom_y).

    char width is approximated at 0.545 * font size for DejaVu Sans / Calibri text.
    """
    entries, y = [], y0
    for level, text in bullets:
        runs, indent, eff = _runs_for_bullet(text, level, size)
        usable = width - indent - 2
        cpl = max(12, int(usable / (0.545 * eff)))
        lines = max(1, -(-len(text) // cpl))
        h = lines * eff * leading
        entries.append((runs, indent, eff, y, h))
        y += h + (gap_sub if level else gap)
    return entries, y


def render_bullets(cv, s, size=15.5, width=860, bottom=492):
    """Draw bullets, shrinking the type size until everything fits above the footer."""
    chosen = None
    for sz in [size, size - 0.5, size - 1, size - 1.5, size - 2, size - 2.5, size - 3]:
        entries, bottom_y = _layout_bullets(s["bullets"], sz, width)
        if bottom_y <= bottom:
            chosen = entries
            break
    if chosen is None:
        chosen = entries                                  # last attempt
    for runs, indent, eff, y, h in chosen:
        cv.text(72 + indent, y, width - indent, h + 2, runs, size=eff, spacing=1.2)
    return chosen[-1][3] + chosen[-1][4]


def _table_height(t, widths, scale):
    total = 34
    for row in t["rows"]:
        max_lines = 1
        for j, cell in enumerate(row):
            cpl = max(12, int((widths[j] - 20) / (0.545 * 11.5 * scale)))
            max_lines = max(max_lines, max(1, -(-len(cell) // cpl)))
        total += max(24, max_lines * (11.5 * scale * 1.25) + 10)
    return total


def render_slide(cv, s, n, total):
    layout = s["layout"]

    if layout in ("bullets", "image_right", "table"):
        cv.rect(0, 0, CANVAS_W, 6, fill=TEAL)                     # top accent
        cv.text(72, 40, 700, 34, s["title"], size=27, bold=True, color=NAVY)
        if s.get("sub"):
            cv.text(72, 76, 700, 22, s["sub"], size=13.5, color=GREY)
        cv.line(72, 108, 168, 108, color=TEAL, width=2.2)

        if layout == "table":
            t = s["table"]
            x0, y0, tw = 72, 148, 816
            widths = [tw * w for w in t["widths"]]
            row_h_head = 34
            cv.rect(x0, y0, tw, row_h_head, fill=NAVY)
            cx = x0
            for j, head in enumerate(t["head"]):
                cv.text(cx + 10, y0 + 8, widths[j] - 20, row_h_head - 10, head, size=12.5,
                        bold=True, color=WHITE, valign="middle")
                cx += widths[j]
            y = y0 + row_h_head
            scale = 1.0 if len(t["rows"]) <= 8 else 0.9
            while scale > 0.72 and _table_height(t, widths, scale) > (492 - y0):
                scale -= 0.04
            for i, row in enumerate(t["rows"]):
                max_lines = 1
                for j, cell in enumerate(row):
                    max_lines = max(max_lines, int(len(cell) / max(18, widths[j] / (6.2 * scale))) + 1)
                rh = max(24, max_lines * (11.5 * scale + 4) + 8)
                if i % 2 == 1:
                    cv.rect(x0, y, tw, rh, fill=LIGHT)
                cx = x0
                for j, cell in enumerate(row):
                    cv.text(cx + 10, y + 5, widths[j] - 20, rh - 10, cell, size=11.5 * scale, color=INK)
                    cx += widths[j]
                cv.line(x0, y + rh, x0 + tw, y + rh, color=(0xD5, 0xDD, 0xE4), width=0.6)
                y += rh
            cv.rect(x0, y0, tw, y - y0, fill=None, line=(0xC3, 0xCE, 0xD8), line_w=0.8)

        elif layout == "image_right":
            render_bullets(cv, s, size=14.5, width=530, bottom=492)
            ix, iy, iw = 632, 132, 256
            ih = 268
            cv.image(img(s["image"]), ix, iy, iw, ih)
            cv.rect(ix, iy, iw, ih, fill=None, line=(0xD5, 0xDD, 0xE4), line_w=0.8)
            if s.get("caption"):
                cv.rect(ix, iy + ih, iw, 34, fill=LIGHT)
                cv.text(ix + 8, iy + ih + 8, iw - 16, 20, s["caption"], size=10, color=NAVY,
                        align="center", italic=False)
        else:
            render_bullets(cv, s, size=15.5, width=816, bottom=492)

    elif layout == "cover":
        cv.rect(0, 0, CANVAS_W, CANVAS_H, fill=NAVY)
        cv.image(img(s["image"]), 0, 0, CANVAS_W, CANVAS_H * 0.52)
        cv.rect(0, CANVAS_H * 0.52 - 2, CANVAS_W, 3, fill=TEAL)
        # logo lock-up with clear space: brand owner top-right, partners on the light band
        cv.rect(668, 16, 264, 86, fill=WHITE)
        cv.image(logo("siemens_energy.png"), 676, 22, 248, 74)
        cv.rect(72, 300, 700, 4, fill=TEAL)
        cv.text(72, 320, 800, 46, C.TITLE, size=40, bold=True, color=WHITE)
        cv.text(72, 372, 800, 26, C.PROJECT, size=19, color=(0x9F, 0xD8, 0xD8))
        cv.text(72, 404, 800, 22, C.SITE, size=13, color=(0xC9, 0xD4, 0xDE))
        cv.text(72, 430, 800, 20, C.CLIENT + "   ·   " + C.CONTRACTOR, size=12, color=(0xC9, 0xD4, 0xDE))
        cv.rect(72, 462, 816, 34, fill=(0x1B, 0x35, 0x4C))
        cv.text(84, 470, 792, 20, C.PREPARED_BY, size=13, bold=True, color=WHITE)
        cv.text(72, 505, 500, 18, f"{C.DOC_NO} · {C.REV} · For review and approval", size=9.5,
                color=(0x8F, 0xA3, 0xB5))
        cv.text(500, 505, 388, 18, "Restricted · © Siemens Energy · Siemens Energy is a trademark licensed by Siemens AG",
                size=7.5, color=(0x8F, 0xA3, 0xB5), align="right")
        cv.image(logo("almial.png"), 76, 208, 168, 38)
        cv.image(logo("bgc.png"), 268, 196, 66, 62)
        return

    elif layout == "closing":
        cv.rect(0, 0, CANVAS_W, CANVAS_H, fill=NAVY)
        cv.rect(0, 0, CANVAS_W, 6, fill=TEAL)
        cv.rect(684, 24, 252, 80, fill=WHITE)
        cv.image(logo("siemens_energy.png"), 692, 30, 236, 68)
        cv.text(72, 190, 820, 44, s["title"], size=36, bold=True, color=WHITE)
        cv.text(72, 244, 820, 26, s["sub"], size=17, color=(0x9F, 0xD8, 0xD8))
        y = 300
        for level, text in s["bullets"]:
            cv.text(72, y, 820, 30, text, size=14, color=(0xDD, 0xE5, 0xEC))
            y += 34
        cv.rect(72, 400, 816, 40, fill=(0x1B, 0x35, 0x4C))
        cv.text(84, 410, 792, 22, C.PREPARED_BY, size=13, bold=True, color=WHITE)
        cv.text(72, 470, 816, 30, "Zero Harm — every person, every task, every day.",
                size=13, color=(0x9F, 0xD8, 0xD8))
        cv.image(logo("almial.png"), 72, 508, 120, 27)
        return

    # footer for all content slides
    cv.line(72, 500, 888, 500, color=(0xD5, 0xDD, 0xE4), width=0.8)
    cv.text(72, 508, 620, 18, FOOTER, size=8.5, color=GREY)
    cv.text(700, 508, 188, 18, f"Slide {n} of {total}", size=8.5, color=GREY, align="right")


def render_section_header(cv, s):
    pass


# ------------------------------------------------------------------ builders
def build_pptx(path):
    from pptx import Presentation
    from pptx.util import Pt
    prs = Presentation()
    prs.slide_width, prs.slide_height = Pt(CANVAS_W), Pt(CANVAS_H)
    blank = prs.slide_layouts[6]
    total = len(C.SLIDES)
    for i, s in enumerate(C.SLIDES, 1):
        slide = prs.slides.add_slide(blank)
        render_slide(PptxCanvas(slide, prs), s, i, total)
        if s.get("notes"):
            slide.notes_slide.notes_text_frame.text = s["notes"]
    prs.save(path)
    print(f"{path}: {total} slides")


def build_pdf(path):
    import pymupdf
    doc = pymupdf.open()
    total = len(C.SLIDES)
    for i, s in enumerate(C.SLIDES, 1):
        page = doc.new_page(width=CANVAS_W, height=CANVAS_H)
        render_slide(PdfCanvas(page), s, i, total)
    doc.save(path, garbage=3, deflate=True)
    print(f"{path}: {doc.page_count} pages")
    doc.close()


def compress(path):
    import pymupdf
    tmp = path + ".tmp"
    d = pymupdf.open(path)
    d.subset_fonts(verbose=False)
    d.save(tmp, garbage=4, deflate=True, deflate_images=True, deflate_fonts=True, clean=True)
    d.close()
    os.replace(tmp, path)
    print(f"{path}: {os.path.getsize(path)/1e6:.2f} MB")


if __name__ == "__main__":
    build_pptx("KAZ_Management_Staff_EHS_Induction_Rev02.pptx")
    build_pdf("KAZ_Management_Staff_EHS_Induction_Rev02.pdf")
    compress("KAZ_Management_Staff_EHS_Induction_Rev02.pdf")
