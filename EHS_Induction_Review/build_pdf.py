"""Build the PDF deliverables for the KAZ HSSE Induction review.

  1) KAZ_Management_Staff_EHS_Induction_Rev01.pdf -- landscape slide deck, one page per slide
  2) KAZ_HSSE_Induction_Review_Report_AR.pdf      -- portrait RTL report with fully controlled tables

The report uses a small hand-rolled flow renderer so that every table row height is measured
before drawing (PyMuPDF's built-in table layout mis-sizes rows containing long wrapped text).
"""
import re
import html as _html
import pymupdf

FN = "figo"            # bundled font family with Arabic + Latin coverage
PAD, VPAD = 4.0, 2.6  # table cell padding
GRID = (0.72, 0.79, 0.84)
DARK = (0.059, 0.239, 0.361)
ALT = (0.965, 0.976, 0.984)


# --------------------------------------------------------------------------- helpers
def esc(t):
    return _html.escape(t, quote=False)


def _latin(x):
    return bool(re.search(r"[A-Za-z]", x))


def _arabic(x):
    return bool(re.search(r"[\u0600-\u06FF]", x))


def inline(t, rtl=False):
    """Markdown inline formatting + RTL-safe handling of Latin runs."""
    t = esc(t)
    t = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", t)
    t = re.sub(r"`(.+?)`", lambda m: '<span dir="ltr" style="font-family:monospace;background:#F2F4F6">'
                                     + m.group(1) + "</span>", t)
    if rtl:
        t = re.sub(r"<b>(.+?)</b>",
                   lambda m: ("<b><bdi>" + m.group(1) + "</bdi></b>") if _latin(m.group(1)) else m.group(0), t)
    return t


def table_widths(rows, total_w):
    n = len(rows[0])
    lens = [0.0] * n
    for r in rows:
        for j, c in enumerate(r[:n]):
            lens[j] += len(c) ** 0.62
    tot = sum(lens) or 1
    w = [100 * x / tot for x in lens]
    w = [min(26.0, max(5.0, x)) for x in w]
    tot = sum(w)
    return [total_w * x / tot for x in w]


# --------------------------------------------------------------------------- deck
def build_deck(src, dst):
    md = open(src, encoding="utf-8").read()
    slides, cur = [], None
    for ln in md.split("\n"):
        m = re.match(r"^### SLIDE (\d+)\s*[—-]\s*(.+)$", ln.strip())
        if m:
            cur = {"n": int(m.group(1)), "title": m.group(2).strip(), "lines": []}
            slides.append(cur)
        elif cur is not None:
            cur["lines"].append(ln)

    doc = pymupdf.open()
    W, H = pymupdf.paper_size("a4-l")
    M = 40

    p = doc.new_page(width=W, height=H)
    cover = f'''
    <div style="font-family:{FN};text-align:center">
      <div style="font-size:13px;color:#0F3D5C"><b>SIEMENS ENERGY · KAZ POWER PLANT UPGRADE PROJECT · KHOR AL-ZUBAIR, BASRAH — IRAQ</b></div>
      <div style="height:30px"></div>
      <div style="font-size:34px;color:#0F3D5C"><b>MANAGEMENT &amp; STAFF<br/>EHS INDUCTION</b></div>
      <div style="height:18px"></div>
      <div style="font-size:16px;color:#134B6E">Rev. 01 — Proposed Full Rewrite (Corrected &amp; Aligned with Siemens Energy EHS Requirements)</div>
      <div style="height:34px"></div>
      <div style="font-size:11px">
        Document No. <b>KAZ-EHS-IND-001</b> · Rev. <b>01</b> · Status: <b>For Review &amp; Approval</b><br/>
        Client: Ministry of Electricity (MoE) &amp; Basrah Gas Company (BGC) · Principal Contractor: Siemens Energy<br/>
        Prepared by: EHS Manager (Al-Mial / AMC) · Reviewed by: SE EHSMIP · Approved by: Project Director / EHS Director<br/>
        Classification: <b>Restricted</b> — Siemens Energy is a trademark licensed by Siemens AG
      </div>
      <div style="height:26px"></div>
      <div style="font-size:12px;color:#C00000"><b>Zero Harm · Zero incidents is achievable · Health and safety – no compromises · We take care of each other</b></div>
    </div>'''
    p.insert_htmlbox(pymupdf.Rect(M, 90, W - M, H - 60), cover)

    total_pages = len(slides) + 1
    for s in slides:
        txt = re.sub(r"\n{2,}", "\n", "\n".join(s["lines"])).strip()
        head = (f'<div style="font-family:{FN}">'
                f'<div style="font-size:9px;color:#7A8B99">KAZ PP Upgrade Project — Management &amp; Staff EHS Induction · '
                f'Rev. 01 · KAZ-EHS-IND-001 · Restricted</div>'
                f'<div style="font-size:17px;color:#0F3D5C"><b>SLIDE {s["n"]} — {esc(s["title"])}</b></div>'
                f'<div style="height:6px"></div></div>')
        placed = False
        for size in (11, 10.5, 10, 9.5, 9, 8.5, 8):
            page = doc.new_page(width=W, height=H)
            page.insert_htmlbox(pymupdf.Rect(M, 24, W - M, 40), head)
            content = _md_fragment(txt, size)
            res = page.insert_htmlbox(pymupdf.Rect(M + 6, 56, W - M - 6, H - 34), content)
            if res and res[0] >= 0:
                placed = True
                break
            doc.delete_page(doc.page_count - 1)
        if not placed:
            page = doc.new_page(width=W, height=H)
            page.insert_htmlbox(pymupdf.Rect(M, 24, W - M, 40), head)
            page.insert_htmlbox(pymupdf.Rect(M + 6, 60, W - M - 6, H - 34), _md_fragment(txt, 8))
        pg = doc[doc.page_count - 1]
        pg.insert_text(pymupdf.Point(M, H - 16), "Zero Harm — Everyone goes home safely, every day.",
                       fontname="helv", fontsize=7.5, color=(0.35, 0.45, 0.55))
        pg.insert_text(pymupdf.Point(W - M - 78, H - 16), f"Page {doc.page_count} of {total_pages}",
                       fontname="helv", fontsize=7.5, color=(0.35, 0.45, 0.55))
    doc.save(dst, garbage=3, deflate=True)
    print(f"{dst}: {doc.page_count} pages")
    doc.close()


def _md_fragment(md, size, rtl=False):
    """Small markdown -> html converter used for one slide / one text block."""
    out, i, lines = [], 0, md.split("\n")
    while i < len(lines):
        ln = lines[i].rstrip()
        if ln.strip().startswith("|") and i + 1 < len(lines) and set(lines[i + 1].replace("|", "").strip()) <= set("-: "):
            rows = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                if not set(lines[i].replace("|", "").strip()) <= set("-: "):
                    rows.append([c.strip() for c in lines[i].strip().strip("|").split("|")])
                i += 1
            widths = table_widths(rows, 100)
            out.append('<table style="width:100%;table-layout:fixed;border-collapse:collapse;font-size:8.5px" '
                       'border="1" bordercolor="#B8C6D1">')
            out.append("<colgroup>" + "".join(f'<col style="width:{w:.2f}%"/>' for w in widths) + "</colgroup>")
            out.append("<tr>" + "".join(
                f'<td bgcolor="#0F3D5C" style="padding:3px"><b><font color="#FFFFFF">{inline(c, rtl)}</font></b></td>'
                for c in rows[0]) + "</tr>")
            for r in rows[1:]:
                out.append("<tr>" + "".join(f'<td style="padding:3px">{inline(c, rtl)}</td>' for c in r) + "</tr>")
            out.append("</table><br/>")
            continue
        if ln.startswith("### "):
            out.append(f'<h3 style="font-size:{size + 1.5}px;color:#134B6E">{inline(ln[4:], rtl)}</h3>')
        elif ln.startswith("## "):
            out.append(f'<h2 style="font-size:{size + 3}px;color:#0F3D5C">{inline(ln[3:], rtl)}</h2>')
        elif ln.startswith("# "):
            out.append(f'<h1 style="font-size:{size + 5}px;color:#0F3D5C">{inline(ln[2:], rtl)}</h1>')
        elif ln.startswith("- ") or ln.startswith("* "):
            out.append(f'<li style="margin-left:8px;font-size:{size}px">{inline(ln[2:], rtl)}</li>')
        elif re.match(r"^\d+\.\s", ln):
            out.append('<li style="margin-left:8px;font-size:%spx">%s</li>'
                       % (size, inline(re.sub(r"^\d+\.\s", "", ln), rtl)))
        elif ln.startswith("> "):
            out.append(f'<div style="background:#EAF1F6;padding:5px;border-left:3px solid #0F3D5C;font-size:{size}px">'
                       + inline(ln[2:], rtl) + "</div>")
        elif ln.strip() == "---":
            out.append("<hr/>")
        elif ln.strip():
            out.append(f'<p style="margin:3px 0;font-size:{size}px">{inline(ln, rtl)}</p>')
        i += 1
    return f'<div style="font-family:{FN};direction:{"rtl" if rtl else "ltr"}">' + "\n".join(out) + "</div>"


# --------------------------------------------------------------------------- report
class FlowReport:
    """Portrait RTL report with measured blocks and fully controlled tables."""

    def __init__(self, path, rtl=True, footer="KAZ-EHS-REV-2026-001  |  Rev. 01  |  Restricted"):
        self.path = path
        self.rtl = rtl
        self.footer = footer
        self.doc = pymupdf.open()
        self.W, self.H = pymupdf.paper_size("a4")
        self.ml, self.mt, self.mr, self.mb = 34, 40, 34, 46
        self.cw = self.W - self.ml - self.mr
        self._scratch_doc = pymupdf.open()
        self._scratch = self._scratch_doc.new_page(width=self.W, height=self.H)
        self._measures = 0
        self.page = None
        self.y = 0
        self._new_page()

    # -- page management
    def _new_page(self):
        self.page = self.doc.new_page(width=self.W, height=self.H)
        self.y = self.mt

    @property
    def space(self):
        return (self.H - self.mb) - self.y

    def measure(self, html, width, height=4000):
        # recycle the measurement page: an ever-growing content stream becomes very slow
        self._measures += 1
        if self._measures % 200 == 0:
            self._scratch_doc.close()
            self._scratch_doc = pymupdf.open()
            self._scratch = self._scratch_doc.new_page(width=self.W, height=self.H)
        rect = pymupdf.Rect(0, 0, width, height)
        res = self._scratch.insert_htmlbox(rect, html)
        return height - res[0]

    def block(self, html, width=None, keep=True, gap_after=0.0):
        width = width or self.cw
        h = self.measure(html, width)
        if keep and h > self.space and h <= (self.H - self.mt - self.mb):
            self._new_page()
        x0 = self.ml if width == self.cw else self.ml
        self.page.insert_htmlbox(pymupdf.Rect(x0, self.y, x0 + width, self.y + h), html)
        self.y += h + gap_after

    def rule(self, gap=3.0):
        if self.space < 10:
            self._new_page()
        self.y += gap
        self.page.draw_line(pymupdf.Point(self.ml, self.y), pymupdf.Point(self.W - self.mr, self.y),
                            color=(0.80, 0.85, 0.89), width=0.6)
        self.y += gap + 2

    # -- tables
    def _cell_html(self, text, header=False):
        inner = inline(text, self.rtl)
        if self.rtl and _latin(text) and not _arabic(text):
            inner = f'<div dir="ltr" style="text-align:left">{inner}</div>'
        else:
            inner = f'<div dir="{"rtl" if self.rtl else "ltr"}" style="text-align:{"right" if self.rtl else "left"}">{inner}</div>'
        if header:
            inner = f'<b><font color="#FFFFFF">{inner}</font></b>'
        return f'<div style="font-family:{FN};font-size:7.6px;line-height:1.32">{inner}</div>'

    def table(self, rows, font=7.6):
        if not rows:
            return
        widths = table_widths(rows, self.cw)
        header, body = rows[0], rows[1:]

        def row_h(row):
            hs = [self.measure(self._cell_html(c), max(20.0, widths[j] - 2 * PAD))
                  for j, c in enumerate(row[:len(widths)])]
            return max(hs) + 2 * VPAD + 1.0

        def draw_row(row, header=False, fill=None, h=None):
            h = h if h is not None else row_h(row)
            if fill:
                self.page.draw_rect(pymupdf.Rect(self.ml, self.y, self.W - self.mr, self.y + h),
                                    color=None, fill=fill)
            x = self.ml
            for j, c in enumerate(row[:len(widths)]):
                self.page.insert_htmlbox(
                    pymupdf.Rect(x + PAD, self.y + VPAD, x + widths[j] - PAD, self.y + h - VPAD),
                    self._cell_html(c, header))
                self.page.draw_rect(pymupdf.Rect(x, self.y, x + widths[j], self.y + h), color=GRID, width=0.4)
                x += widths[j]
            self.y += h

        def draw_header():
            draw_row(header, header=True, fill=DARK)

        hh_header = row_h(header)
        if self.space < hh_header + 20:
            self._new_page()
        draw_header()
        for k, r in enumerate(body):
            h = row_h(r)
            if h > self.space:
                self._new_page()
                draw_header()
            if h > (self.H - self.mt - self.mb):          # oversized single row -> stacked fallback
                for j, c in enumerate(r[:len(widths)]):
                    hh = (f'<div style="font-family:{FN};font-size:7.8px;direction:{ "rtl" if self.rtl else "ltr"}">'
                          f'<b>{inline(header[j], self.rtl)}:</b> {inline(c, self.rtl)}</div>')
                    self.block(hh, gap_after=1.5)
                self.rule(2)
                continue
            draw_row(r, fill=ALT if k % 2 == 0 else None, h=h)
        self.y += 4

    # -- markdown
    def render_markdown(self, md):
        lines, i = md.split("\n"), 0
        while i < len(lines):
            ln = lines[i].rstrip()
            if ln.strip().startswith("|") and i + 1 < len(lines) and set(lines[i + 1].replace("|", "").strip()) <= set("-: "):
                rows = []
                while i < len(lines) and lines[i].strip().startswith("|"):
                    if not set(lines[i].replace("|", "").strip()) <= set("-: "):
                        rows.append([c.strip() for c in lines[i].strip().strip("|").split("|")])
                    i += 1
                self.table(rows)
                continue
            if ln.startswith("# "):
                self.block(f'<div style="font-family:{FN};direction:rtl;font-size:16px;color:#0F3D5C"><b>{inline(ln[2:], self.rtl)}</b></div>', gap_after=4)
            elif ln.startswith("## "):
                self.block(f'<div style="font-family:{FN};direction:rtl;font-size:13.5px;color:#0F3D5C"><b>{inline(ln[3:], self.rtl)}</b></div>', gap_after=3)
            elif ln.startswith("### "):
                self.block(f'<div style="font-family:{FN};direction:rtl;font-size:11.5px;color:#134B6E"><b>{inline(ln[4:], self.rtl)}</b></div>', gap_after=2)
            elif ln.strip() == "---":
                self.rule(5)
            elif not ln.strip():
                self.y += 2
            elif ln.startswith("- ") or ln.startswith("* ") or re.match(r"^\d+\.\s", ln):
                item = re.sub(r"^(-\s|\*\s|\d+\.\s)", "", ln)
                self.block(f'<div style="font-family:{FN};direction:rtl;font-size:9.5px;line-height:1.4">'
                           f'&#8226;&nbsp;&nbsp;{inline(item, self.rtl)}</div>', gap_after=0.5)
            elif ln.startswith("> "):
                self.block(f'<div style="font-family:{FN};direction:rtl;font-size:9px;background:#EAF1F6;'
                           f'border-left:3px solid #0F3D5C;padding:4px">{inline(ln[2:], self.rtl)}</div>', gap_after=2)
            elif "النص الحرفي" in ln and ":" in ln:
                label, quote = ln.split(":", 1)
                self.block(f'<div style="font-family:{FN};direction:rtl;font-size:9.5px">'
                           f'<b>{inline(label + ":", self.rtl)}</b></div>', gap_after=1)
                self.block(f'<div dir="ltr" style="font-family:{FN};text-align:left;font-size:8.4px;'
                           f'background:#F4F7FA;border-left:3px solid #0F3D5C;padding:4px">'
                           f'{inline(quote.strip(), False)}</div>', gap_after=2)
            else:
                self.block(f'<div style="font-family:{FN};direction:rtl;font-size:9.5px;line-height:1.45;'
                           f'text-align:justify">{inline(ln, self.rtl)}</div>', gap_after=1.5)
            i += 1

    def close(self):
        total = self.doc.page_count
        for i, pg in enumerate(self.doc):
            pg.insert_text(pymupdf.Point(self.ml, self.H - 22),
                           self.footer, fontname="helv",
                           fontsize=7, color=(0.35, 0.45, 0.55))
            pg.insert_text(pymupdf.Point(self.W - self.mr - 88, self.H - 22),
                           f"Page {i + 1} of {total}", fontname="helv", fontsize=7, color=(0.35, 0.45, 0.55))
        self.doc.save(self.path, garbage=3, deflate=True)
        print(f"{self.path}: {total} pages")
        self.doc.close()


def build_report(src, dst, footer="KAZ-EHS-REV-2026-001  |  Rev. 01  |  Restricted"):
    md = open(src, encoding="utf-8").read()
    md = re.sub(r"^---$", "", md, flags=re.M)
    rep = FlowReport(dst, rtl=True, footer=footer)
    rep.render_markdown(md)
    rep.close()


def compress(path):
    """Rebuild the file with subset fonts -- keeps PDFs at a shareable size."""
    import os
    tmp = path + ".tmp"
    d = pymupdf.open(path)
    d.subset_fonts(verbose=False)
    d.save(tmp, garbage=4, deflate=True, deflate_images=True, deflate_fonts=True, clean=True)
    d.close()
    os.replace(tmp, path)
    print(f"{path}: {os.path.getsize(path) / 1e6:.2f} MB")


if __name__ == "__main__":
    build_deck("KAZ_Management_Staff_EHS_Induction_Rev01_EN.md",
               "KAZ_Management_Staff_EHS_Induction_Rev01.pdf")
    compress("KAZ_Management_Staff_EHS_Induction_Rev01.pdf")
    build_report("KAZ_HSSE_Induction_Review_Report_AR.md",
                 "KAZ_HSSE_Induction_Review_Report_AR.pdf")
    compress("KAZ_HSSE_Induction_Review_Report_AR.pdf")
