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

    def measure(self, runs, size, width, align="left", spacing=1.0, color=INK,
                bold=False, italic=False, ceiling=600.0):
        return _story_height(runs, size, width, align, spacing, color, bold, italic, ceiling)

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


def _runs_html_static(runs, size, color, bold, italic, align, spacing):
    """Same markup the PDF backend produces, usable without a canvas instance."""
    items = runs if isinstance(runs, list) else [(runs, {})]
    parts = []
    for text, style in items:
        t = _html.escape(text).replace("\n", "<br/>")
        t = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", t)
        st = f'font-size:{style.get("size", size)}pt;'
        if style.get("bold", bold): st += "font-weight:bold;"
        if style.get("italic", italic): st += "font-style:italic;"
        col = style.get("color", color)
        st += f'color:#{"".join(f"{v:02X}" for v in col)};'
        parts.append(f'<span style="{st}">{t}</span>')
    return (f'<div style="font-family:dm;line-height:{spacing};text-align:{align};'
            f'font-variant-ligatures:none">' + "".join(parts) + "</div>")


def _story_height(runs, size, width, align, spacing, color, bold, italic, ceiling=600.0):
    """Lay text out with MuPDF and return the height it needs (DejaVu metrics).

    DejaVu is wider than the Calibri used by the PPTX backend, so using this for
    both backends errs on the generous side - boxes grow, never overlap.
    """
    key = (repr(runs), size, round(width, 1), align, round(spacing, 2), color, bold, italic)
    if key in _MEASURE_CACHE:
        return _MEASURE_CACHE[key]
    import pymupdf
    html = _runs_html_static(runs, size, color, bold, italic, align, spacing)
    story = pymupdf.Story(html=html, user_css=FONT_CSS, archive=_font_archive())
    filled = story.place(pymupdf.Rect(0, 0, width, ceiling))[1]
    out = max(0.0, (filled[3] - filled[1]) - 24.0)
    _MEASURE_CACHE[key] = out
    return out


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
        return (f'<div style="font-family:dm;line-height:{spacing};text-align:{align};'
                f'font-variant-ligatures:none">'
                + "".join(parts) + "</div>")

    def measure(self, runs, size, width, align="left", spacing=1.0, color=INK,
                bold=False, italic=False, ceiling=600.0):
        """Height (pt) the run list needs at this width - see _story_height()."""
        return _story_height(runs, size, width, align, spacing, color, bold, italic, ceiling)

    def text(self, x, y, w, h, runs, size=14, color=INK, bold=False, align="left",
             valign="top", spacing=1.0, italic=False, autofit=True):
        import pymupdf
        if autofit:
            need = self.measure(runs, size, w, align, spacing, color, bold, italic)
            h = max(h, need + 4.0)
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
    """Return (marker, marker styled, text runs, indent, effective size) for one line.

    The marker is drawn in its own narrow column so that wrapped text aligns
    under the first word (hanging indent) instead of under the marker.
    """
    if level == 0:
        return ("▪", {"color": TEAL, "bold": True}, [(text, {"color": INK})], 0, size)
    if level == 1:
        return ("•", {"color": GREY, "bold": True}, [(text, {"color": INK})], 18, size - 1)
    return ("",  {}, [(text, {"color": PURPLE, "bold": True})], 0, size - 1)


MARKER_W = 16.0
_MEASURE_CACHE = {}      # width of the marker column
def _layout_bullets(bullets, size, width, x=72, y0=132, leading=1.30, gap=7.0, gap_sub=4.0):
    """Estimate the laid-out lines: returns (entries, bottom_y).

    char width is approximated at 0.545 * font size for DejaVu Sans / Calibri text.
    """
    entries, y = [], y0
    for level, text in bullets:
        marker, mstyle, runs, indent, eff = _runs_for_bullet(text, level, size)
        lead = MARKER_W if marker else 0.0
        usable = width - indent - lead - 2
        cpl = max(12, int(usable / (0.545 * eff)))
        lines = max(1, -(-len(text) // cpl))
        h = lines * eff * leading
        entries.append((marker, mstyle, runs, indent, eff, y, h))
        y += h + (gap_sub if level else gap)
    return entries, y


def render_bullets(cv, s, size=15.5, width=860, bottom=492, x=72, y0=132):
    """Draw bullets, shrinking the type size until everything fits above the footer."""
    entries = None
    for sz in [size - i * 0.5 for i in range(7)]:
        entries, bottom_y = _layout_bullets(s["bullets"], sz, width, x=x, y0=y0)
        if bottom_y <= bottom:
            break
    for marker, mstyle, runs, indent, eff, y, h in entries:
        if marker:
            cv.text(x + indent, y + 0.5, MARKER_W, h, marker, size=eff, spacing=1.2, **mstyle)
        left = x + indent + (MARKER_W if marker else 0.0)
        cv.text(left, y, width - indent - (MARKER_W if marker else 0.0), h + 2, runs,
                size=eff, spacing=1.2)
    return entries[-1][5] + entries[-1][6]


def _table_height(t, widths, scale):
    total = 34
    for row in t["rows"]:
        max_lines = 1
        for j, cell in enumerate(row):
            cpl = max(12, int((widths[j] - 20) / (0.545 * 11.5 * scale)))
            max_lines = max(max_lines, max(1, -(-len(cell) // cpl)))
        total += max(24, max_lines * (11.5 * scale * 1.25) + 10)
    return total



# ------------------------------------------------------------------ structure
def build_flow():
    """Expand C.SLIDES into the final running order, inserting chapter dividers.

    Every content slide is tagged with its chapter so the header and footer can
    label it without duplicating the structure in two places.
    """
    starts = {c["starts_at"]: c for c in C.CHAPTERS}
    out, current = [], None
    for s in C.SLIDES:
        if s["title"] in starts:
            current = starts[s["title"]]
            out.append(dict(layout="divider", chapter=current, title=current["name"],
                            notes=f"Chapter {current['no']} — {current['name']}. "
                                  f"{current['blurb']} This chapter covers: "
                                  f"{', '.join(current['topics'])}."))
        item = dict(s)
        item["_chapter"] = current
        out.append(item)

    # cross references such as "see slide {{slide:Emergency Contacts}}" are resolved
    # against the final running order, so inserting a slide can never leave a stale number
    number = {}
    for i, it in enumerate(out, 1):
        number.setdefault(it["title"], i)

    def fix(t):
        return re.sub(r"\{\{slide:(.+?)\}\}", lambda m: str(number.get(m.group(1), "?")), t)

    for it in out:
        for key in ("title", "sub", "caption", "notes"):
            if it.get(key):
                it[key] = fix(it[key])
        if it.get("bullets"):
            it["bullets"] = [(lvl, fix(t)) for lvl, t in it["bullets"]]
        t = it.get("table")
        if t:
            t["head"] = [fix(c) for c in t["head"]]
            t["rows"] = [[fix(c) for c in row] for row in t["rows"]]
    return out


def render_footer(cv, s, n, total):
    """Document control, chapter and page number - identical on every slide."""
    cv.line(72, 500, 98, 500, color=TEAL, width=2.4)
    cv.line(98, 500, 888, 500, color=(0xD5, 0xDD, 0xE4), width=0.8)
    cv.text(72, 508, 430, 18, FOOTER, size=8.5, color=GREY)
    ch = s.get("_chapter")
    if ch:
        cv.text(420, 508, 330, 18, f"{ch['no']} · {ch['name']}", size=8.5,
                color=GREY, align="center")
    cv.text(700, 508, 188, 18, f"Slide {n} of {total}", size=8.5, color=GREY, align="right")


def render_slide(cv, s, n, total):
    layout = s["layout"]

    if layout in ("bullets", "image_right", "table"):
        cv.rect(0, 0, CANVAS_W, 6, fill=TEAL)                     # top accent
        cv.image(logo("siemens_energy.png"), 788, 30, 100, 31)     # brand mark, top right
        ch = s.get("_chapter")
        if ch:
            cv.text(72, 28, 690, 15, f"CHAPTER {ch['no']}  ·  {ch['name'].upper()}",
                    size=8.5, bold=True, color=TEAL)
        tts = 27.0
        th = cv.measure(s["title"], tts, 700, bold=True, spacing=1.15)
        while th > 34 and tts > 21.0:                             # keep every title to one line
            tts -= 1.0
            th = cv.measure(s["title"], tts, 700, bold=True, spacing=1.15)
        title_y = 46.0 if ch else 40.0
        cv.text(72, title_y, 700, th + 6, s["title"], size=tts, bold=True, color=NAVY, spacing=1.15)
        sub_y = title_y + th + 9
        if s.get("sub"):
            cv.text(72, sub_y, 700, 22, s["sub"], size=13.5, color=GREY)
        rule_y = min(112.0, sub_y + 34)
        cv.line(72, rule_y, 168, rule_y, color=TEAL, width=2.4)
        cv.line(168, rule_y + 0.9, 888, rule_y + 0.9, color=(0xD5, 0xDD, 0xE4), width=0.9)
        body_top = max(134.0, rule_y + 26)

        if layout == "table":
            t = s["table"]
            x0, y0, tw = 72, max(148.0, body_top + 16), 816
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
            render_bullets(cv, s, size=14.5, width=530, bottom=492, y0=body_top)
            ix, iy, iw = 632, body_top, 256
            ih = 268
            if s.get("image"):
                cv.image(img(s["image"]), ix, iy, iw, ih)
                cv.rect(ix, iy, iw, ih, fill=None, line=(0xD5, 0xDD, 0xE4), line_w=0.8)
            else:
                # reserved slot: the site photograph is dropped in from PowerPoint
                cv.rect(ix, iy, iw, ih, fill=LIGHT, line=(0xB9, 0xC6, 0xD1), line_w=1.0)
                cv.rect(ix + 88, iy + 104, 80, 58, fill=None, line=(0x9F, 0xB2, 0xC2), line_w=1.4)
                cv.rect(ix + 96, iy + 112, 64, 42, fill=(0xE7, 0xEE, 0xF3), line=None)
                cv.text(ix + 20, iy + 176, iw - 40, 20, "INSERT SITE PHOTOGRAPH",
                        size=9.5, bold=True, color=NAVY, align="center")
                cv.text(ix + 20, iy + 194, iw - 40, 32,
                        "Right-click the frame in PowerPoint\nthen choose Change Picture",
                        size=8.5, color=GREY, align="center", spacing=1.3)
            if s.get("caption"):
                cap_h = 30 if len(s["caption"]) <= 42 else 44
                cv.rect(ix, iy + ih, iw, cap_h, fill=LIGHT)
                cv.text(ix + 8, iy + ih + 5, iw - 16, cap_h - 10, s["caption"], size=10,
                        color=NAVY, align="center", spacing=1.15)
        else:
            render_bullets(cv, s, size=15.5, width=816, bottom=492, y0=body_top)

    elif layout == "divider":
        ch = s["chapter"]
        cv.rect(0, 0, CANVAS_W, CANVAS_H, fill=NAVY)
        cv.rect(0, 0, CANVAS_W, 6, fill=TEAL)
        cv.rect(0, 6, 6, CANVAS_H - 6, fill=TEAL)                 # spine
        cv.rect(676, 24, 260, 82, fill=WHITE)
        cv.image(logo("siemens_energy.png"), 684, 30, 244, 74)
        cv.text(96, 128, 300, 150, ch["no"], size=120, bold=True, color=(0x18, 0x38, 0x51),
                spacing=1.0)
        nsize = 32.0
        nh = cv.measure(ch["name"], nsize, 760, bold=True, spacing=1.1)
        while nh > 40 and nsize > 24:                 # keep the chapter title on one line
            nsize -= 1.0
            nh = cv.measure(ch["name"], nsize, 760, bold=True, spacing=1.1)
        cv.text(96, 286, 760, nh + 6, ch["name"], size=nsize, bold=True, color=WHITE, spacing=1.1)
        cv.rect(96, 336, 96, 3, fill=TEAL)
        cv.text(96, 352, 720, 26, ch["blurb"], size=14.5, color=(0x9F, 0xD8, 0xD8), spacing=1.2)
        cv.text(96, 396, 780, 46, "  ·  ".join(ch["topics"]), size=11.5,
                color=(0x8F, 0xA8, 0xBC), spacing=1.35)
        cv.text(96, 468, 500, 20, f"{C.DOC_NO} · {C.REV} · Restricted", size=9,
                color=(0x6C, 0x83, 0x99))
        cv.text(600, 468, 276, 20, f"Slide {n} of {total}", size=9,
                color=(0x6C, 0x83, 0x99), align="right")
        return

    elif layout == "contents":
        cv.rect(0, 0, CANVAS_W, 6, fill=TEAL)
        cv.image(logo("siemens_energy.png"), 788, 30, 100, 31)
        cv.text(72, 28, 690, 15, "COURSE STRUCTURE", size=8.5, bold=True, color=TEAL)
        cv.text(72, 46, 700, 34, s["title"], size=27, bold=True, color=NAVY)
        cv.text(72, 84, 700, 22, s["sub"], size=13.5, color=GREY)
        cv.line(72, 112, 168, 112, color=TEAL, width=2.4)
        cv.line(168, 112.9, 888, 112.9, color=(0xD5, 0xDD, 0xE4), width=0.9)
        gap_x, gap_y = 18, 12
        cw = (816 - gap_x) / 2
        chh = (338 - 2 * gap_y) / 3
        for k, ch in enumerate(C.CHAPTERS):
            col, row = k % 2, k // 2
            x = 72 + col * (cw + gap_x)
            y = 138 + row * (chh + gap_y)
            cv.rect(x, y, cw, chh, fill=LIGHT, line=(0xD5, 0xDD, 0xE4), line_w=0.9)
            cv.rect(x, y, 3.5, chh, fill=TEAL)
            cv.text(x + 16, y + 13, 44, 26, ch["no"], size=21, bold=True, color=TEAL)
            nh = cv.measure(ch["name"], 12.5, cw - 78, bold=True, spacing=1.2)
            cv.text(x + 62, y + 13, cw - 78, nh + 4, ch["name"], size=12.5, bold=True,
                    color=NAVY, spacing=1.2)
            by = y + 15 + max(nh, 20) + 4
            cv.text(x + 62, by, cw - 78, chh - (by - y) - 8, ch["blurb"], size=9.5, color=GREY,
                    spacing=1.25)
        render_footer(cv, s, n, total)
        return

    elif layout == "cover":
        cv.image(img(s["image"]), 0, 0, CANVAS_W, CANVAS_H)
        cv.rect(0, 0, CANVAS_W, 6, fill=TEAL)
        cv.rect(668, 22, 268, 88, fill=WHITE)                     # brand mark, clear space
        cv.image(logo("siemens_energy.png"), 676, 28, 252, 76)

        cv.rect(72, 282, 96, 4, fill=TEAL)
        ts = 40.0
        th = cv.measure(C.TITLE, ts, 816, bold=True, spacing=1.1)
        while th > 52 and ts > 30:
            ts -= 2
            th = cv.measure(C.TITLE, ts, 816, bold=True, spacing=1.1)
        cv.text(72, 300, 816, th + 6, C.TITLE, size=ts, bold=True, color=WHITE, spacing=1.1)
        y = 300 + th + 12
        cv.text(72, y, 816, 26, C.PROJECT, size=19, color=(0x9F, 0xD8, 0xD8))
        y += 30
        cv.text(72, y, 816, 20, C.SITE, size=12.5, color=(0xC9, 0xD4, 0xDE))
        y += 22
        cv.text(72, y, 816, 20, C.CLIENT + "   ·   " + C.CONTRACTOR, size=11.5,
                color=(0xC9, 0xD4, 0xDE))
        y += 21
        cv.text(72, y, 816, 20, C.SUBCONTRACTOR, size=11.5, color=(0xC9, 0xD4, 0xDE))

        cv.rect(72, 456, 816, 34, fill=(0x12, 0x2A, 0x40))
        cv.rect(72, 456, 4, 34, fill=TEAL)
        cv.text(86, 464, 790, 20, C.PREPARED_BY, size=13, bold=True, color=WHITE)
        cv.text(72, 502, 500, 18, f"{C.DOC_NO} · {C.REV} · For review and approval", size=9.5,
                color=(0x9A, 0xAD, 0xBE))
        cv.text(500, 502, 388, 18,
                "Restricted · © Siemens Energy · Siemens Energy is a trademark licensed by Siemens AG",
                size=7.5, color=(0x9A, 0xAD, 0xBE), align="right")
        return

    elif layout == "closing":
        cv.rect(0, 0, CANVAS_W, CANVAS_H, fill=NAVY)
        cv.rect(0, 0, CANVAS_W, 6, fill=TEAL)
        cv.rect(684, 24, 252, 80, fill=WHITE)
        cv.image(logo("siemens_energy.png"), 692, 30, 236, 68)
        ts = 36
        th = cv.measure(s["title"], ts, 820, spacing=1.15, bold=True)
        while th > 95 and ts > 24:                   # keep the closing headline to two lines max
            ts -= 2
            th = cv.measure(s["title"], ts, 820, spacing=1.15, bold=True)
        cv.text(72, 170, 820, th + 5, s["title"], size=ts, bold=True, color=WHITE, spacing=1.15)
        cv.text(72, 170 + th + 12, 820, 30, s["sub"], size=17, color=(0x9F, 0xD8, 0xD8), spacing=1.15)
        y, last = 170 + th + 60, 170 + th + 60
        for level, text in s["bullets"]:
            bh = cv.measure(text, 14, 820, spacing=1.2)
            cv.text(72, y, 820, bh + 4, text, size=14, color=(0xDD, 0xE5, 0xEC), spacing=1.2)
            y += bh + 16
            last = y
        band = min(max(404.0, last + 14), 428.0)
        cv.rect(72, band, 816, 40, fill=(0x1B, 0x35, 0x4C))
        cv.text(84, band + 10, 792, 24, C.PREPARED_BY, size=13, bold=True, color=WHITE)
        cv.text(72, band + 56, 816, 30, "No task is so urgent that it cannot be done safely.",
                size=13, color=(0x9F, 0xD8, 0xD8))
        cv.text(72, 508, 816, 18, f"{C.DOC_NO} · {C.REV} · Restricted   ·   "
                f"{C.SUBCONTRACTOR}", size=8.5, color=(0x7C, 0x93, 0xA8))
        return

    # footer for all content slides
    render_footer(cv, s, n, total)


def render_section_header(cv, s):
    pass


# ------------------------------------------------------------------ builders
def build_pptx(path):
    from pptx import Presentation
    from pptx.util import Pt
    prs = Presentation()
    prs.slide_width, prs.slide_height = Pt(CANVAS_W), Pt(CANVAS_H)
    blank = prs.slide_layouts[6]
    items = build_flow()
    total = len(items)
    for i, s in enumerate(items, 1):
        slide = prs.slides.add_slide(blank)
        render_slide(PptxCanvas(slide, prs), s, i, total)
        if s.get("notes"):
            slide.notes_slide.notes_text_frame.text = s["notes"]
    prs.save(path)
    print(f"{path}: {total} slides")


def build_pdf(path):
    import pymupdf
    doc = pymupdf.open()
    items = build_flow()
    total = len(items)
    for i, s in enumerate(items, 1):
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
