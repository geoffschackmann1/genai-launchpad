"""Shared professional-formatting finisher for package documents.

finalize(doc, title): fonts (Aptos body / Aptos Display headings), 1-inch
margins, footer (title | Confidential | Page X of Y with live PAGE/NUMPAGES
fields), keep-with-next on headings, repeating table header rows across
page breaks.
"""
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt


XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

FONT_BODY = "Aptos"
FONT_DISPLAY = "Aptos Display"


def apply_fonts(doc, body=FONT_BODY, display=FONT_DISPLAY):
    """Modern Office default pairing: Aptos Display for headings/titles, Aptos for body/lists/tables."""
    styles = doc.styles
    for name in ("Normal", "List Bullet", "List Number", "List Paragraph", "Body Text", "No Spacing"):
        if name in styles:
            styles[name].font.name = body
    for name in ("Title", "Subtitle", "Heading 1", "Heading 2", "Heading 3", "Heading 4"):
        if name in styles:
            styles[name].font.name = display


def _field(paragraph, instr):
    r = paragraph.add_run()
    for el, attrs, text in (("w:fldChar", {"w:fldCharType": "begin"}, None),
                            ("w:instrText", {XML_SPACE: "preserve"}, instr),
                            ("w:fldChar", {"w:fldCharType": "end"}, None)):
        e = OxmlElement(el)
        for k, v in attrs.items():
            e.set(qn(k) if k.startswith("w:") else k, v)
        if text is not None:
            e.text = text
        r._r.append(e)
    return r


def _tbl_header_repeat(table):
    row = table.rows[0]
    trPr = row._tr.get_or_add_trPr()
    if trPr.find(qn("w:tblHeader")) is None:
        th = OxmlElement("w:tblHeader")
        th.set(qn("w:val"), "true")
        trPr.append(th)


def toc(doc, levels="1-2"):
    """Insert a Word TOC field (populates on open / F9 in Word)."""
    p = doc.add_paragraph()
    _field(p, f'TOC \\o "{levels}" \\h \\z \\u')
    note = doc.add_paragraph()
    r = note.add_run("(Table of contents - in Word: right-click the field and choose Update Field, or press F9.)")
    r.italic = True
    r.font.size = Pt(8)
    return p


def finalize(doc, title):
    apply_fonts(doc)
    for sec in doc.sections:
        sec.left_margin = sec.right_margin = Inches(1)
        sec.top_margin = Inches(0.9)
        sec.bottom_margin = Inches(0.9)
        footer = sec.footer
        footer.is_linked_to_previous = False
        fp = footer.paragraphs[0]
        for r in list(fp.runs):
            r._r.getparent().remove(r._r)
        left = fp.add_run(f"{title} | Confidential | ")
        left.font.size = Pt(8)
        pg = fp.add_run("Page ")
        pg.font.size = Pt(8)
        _field(fp, "PAGE").font.size = Pt(8)
        of = fp.add_run(" of ")
        of.font.size = Pt(8)
        _field(fp, "NUMPAGES").font.size = Pt(8)
    for p in doc.paragraphs:
        if p.style.name.startswith("Heading") or p.style.name == "Title":
            p.paragraph_format.keep_with_next = True
    for t in doc.tables:
        _tbl_header_repeat(t)
    return doc
