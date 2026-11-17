"""Salida con el formato de los informes mensuales de Asistencia Técnica FAMESA (PDF y Word).

El Word se genera sobre tu propia plantilla (plantilla_informe.docx): portada, índice,
encabezado, pie de página y títulos de sección son los de tu archivo.
"""
import copy
import io
import os
import re
import shutil
import tempfile
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

from lxml import etree

import assets
from informe import _cell, _mk, png_de

BASE = Path(__file__).parent
PLANTILLA = BASE / "plantilla_informe.docx"
FIRMA = BASE / "firma.png"   # opcional: si no existe, no se muestra
ROJO, NEGRO, PLOMO, BANDA, BARRA = "#FF0000", "#111111", "#6F7378", "#DEDEDE", "#262626"
ROJO_TXT = "#E10613"


def _media(nombre):
    with zipfile.ZipFile(PLANTILLA) as z:
        return z.read("word/media/" + nombre)


def _cierre_txt(meta):
    quien = ", ".join(x for x in (meta.get("cliente"), meta.get("unidad")) if x) or "la unidad minera"
    return (f"Se extiende el agradecimiento a {quien} por las facilidades brindadas durante la ejecución "
            "de las actividades programadas, cuyo soporte fue fundamental para el cumplimiento de los "
            "objetivos establecidos.")


def _lineas(txt, n=None):
    ls = [x.strip() for x in str(txt or "").split("\n") if x.strip()]
    return ls[:n] if n else ls


# =============================================================================== PDF
def a_pdf(data, meta):
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT, TA_RIGHT
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.lib.utils import ImageReader
    from reportlab.platypus import (CondPageBreak, Image, KeepTogether, PageBreak, Paragraph,
                                    SimpleDocTemplate, Spacer, Table, TableStyle)
    from reportlab.platypus.tableofcontents import TableOfContents

    W, H = A4
    HR = colors.HexColor
    tmp = tempfile.mkdtemp()
    pie = meta.get("pie") or ""

    def S(name, **k):
        base = dict(fontName="Helvetica", fontSize=10, leading=14, textColor=HR(NEGRO))
        base.update(k)
        return ParagraphStyle(name, **base)

    st = {
        "n": S("n", alignment=TA_JUSTIFY, leading=15, spaceAfter=7),
        "h1": S("h1", fontName="Helvetica-Bold", fontSize=13, leading=16),
        "h2": S("h2", fontName="Helvetica-Bold", fontSize=11.5, leading=15, spaceBefore=8, spaceAfter=3, keepWithNext=1),
        "h3": S("h3", fontName="Helvetica-Bold", fontSize=10, spaceBefore=8, spaceAfter=3, keepWithNext=1),
        "th": S("th", fontName="Helvetica-Bold", fontSize=8.5, leading=10.5, textColor=colors.white, alignment=TA_CENTER),
        "tdl": S("tdl", fontSize=8.5, leading=10.5, alignment=TA_LEFT),
        "tdc": S("tdc", fontSize=8.5, leading=10.5, alignment=TA_CENTER),
        "tdr": S("tdr", fontSize=8.5, leading=10.5, alignment=TA_RIGHT),
        "note": S("note", fontName="Helvetica-Oblique", fontSize=8.5, leading=11, textColor=HR(PLOMO), spaceAfter=8),
        "cap": S("cap", fontName="Helvetica-BoldOblique", fontSize=9, alignment=TA_CENTER, spaceAfter=8),
        "eq": S("eq", fontName="Courier-Bold", fontSize=11, alignment=TA_CENTER, spaceBefore=4, spaceAfter=8),
        "bul": S("bul", leftIndent=16, bulletIndent=3, spaceAfter=4, alignment=TA_JUSTIFY, leading=14.5,
                 bulletFontName="ZapfDingbats", bulletFontSize=8, bulletColor=HR(ROJO)),
        "kvl": S("kvl", fontName="Helvetica-Bold", fontSize=9.5, leading=12),
        "kvv": S("kvv", fontSize=9.5, leading=12),
        "kvh": S("kvh", fontName="Helvetica-Bold", fontSize=9.5, textColor=colors.white, alignment=TA_CENTER),
        "ok": S("ok", fontName="Helvetica-Bold", fontSize=8.5, leading=10.5, textColor=HR("#1E7B34"), alignment=TA_CENTER),
        "bad": S("bad", fontName="Helvetica-Bold", fontSize=8.5, leading=10.5, textColor=HR(ROJO_TXT), alignment=TA_CENTER),
        "right": S("right", alignment=TA_RIGHT, spaceBefore=10, spaceAfter=4),
    }

    def banda(texto, toc_nivel=0):
        t = Table([["", Paragraph(escape(texto), st["h1"])]], colWidths=[.45 * cm, 16.55 * cm])
        t.setStyle(TableStyle([("BACKGROUND", (0, 0), (0, 0), HR(ROJO)), ("BACKGROUND", (1, 0), (1, 0), HR(BANDA)),
                               ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("TOPPADDING", (0, 0), (-1, -1), 7),
                               ("BOTTOMPADDING", (0, 0), (-1, -1), 7), ("LEFTPADDING", (1, 0), (1, 0), 10)]))
        t.keepWithNext = True
        t._toc = (toc_nivel, texto)
        return t

    def tabla(b):
        head, rows = b["head"], b["rows"]
        w = b.get("widths") or [17 / len(head)] * len(head)
        al = b.get("align") or "l" + "c" * (len(head) - 1)
        d = [[Paragraph(_mk(h), st["th"]) for h in head]]
        for r in rows:
            fila = []
            for j, c in enumerate(r):
                txt, s_ = _cell(c)
                fila.append(Paragraph(_mk(txt), st[s_] if s_ else st["td" + al[j]]))
            d.append(fila)
        t = Table(d, colWidths=[x * cm for x in w], repeatRows=1)
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), HR(ROJO_TXT)), ("GRID", (0, 0), (-1, -1), .4, HR("#8C8C8C")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, HR("#F4F4F4")]),
            ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ("LEFTPADDING", (0, 0), (-1, -1), 3), ("RIGHTPADDING", (0, 0), (-1, -1), 3)]))
        return t

    def kv(titulo, rows):
        d = [[Paragraph(_mk(titulo), st["kvh"]), ""]]
        d += [[Paragraph(_mk(k), st["kvl"]), Paragraph(_mk(v), st["kvv"])] for k, v in rows]
        t = Table(d, colWidths=[6.2 * cm, 10.8 * cm])
        t.setStyle(TableStyle([
            ("SPAN", (0, 0), (1, 0)), ("BACKGROUND", (0, 0), (-1, 0), HR(ROJO_TXT)),
            ("BACKGROUND", (0, 1), (0, -1), HR("#EDEDED")), ("GRID", (0, 0), (-1, -1), .5, HR("#333333")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4)]))
        return t

    def imagen(png, ancho):
        ruta = os.path.join(tmp, f"i{len(os.listdir(tmp))}.png")
        with open(ruta, "wb") as f:
            f.write(png)
        w, h = ImageReader(ruta).getSize()
        return Image(ruta, width=ancho * cm, height=ancho * cm * h / w)

    # ---- historia (la primera página la dibuja la función portada)
    story = [PageBreak(), banda("ÍNDICE", -1), Spacer(1, 14)]
    toc = TableOfContents()
    toc.levelStyles = [S("t0", fontName="Helvetica-Bold", fontSize=10.5, leading=16, spaceBefore=3),
                       S("t1", fontSize=9.5, leading=13, leftIndent=16)]
    story += [toc, PageBreak()]
    nfig = 0
    for b in data["blocks"]:
        t = b["t"]
        if t == "h1":
            story += [CondPageBreak(5 * cm), Spacer(1, 6), banda(b["text"].upper(), 0), Spacer(1, 8)]
        elif t == "h2":
            p = Paragraph(_mk(b["text"]), st["h2"])
            p._toc = (1, b["text"])
            story.append(p)
        elif t == "h3":
            story += [CondPageBreak(7 * cm), Paragraph("<u>" + _mk(b["text"]).upper() + "</u>", st["h3"])]
        elif t == "p":
            story.append(Paragraph(_mk(b["text"]), st["n"]))
        elif t == "bullets":
            story += [Paragraph(_mk(i), st["bul"], bulletText="n") for i in b["items"]]
            story.append(Spacer(1, 4))
        elif t == "eq":
            story.append(Paragraph(escape(b["text"]), st["eq"]))
        elif t == "kv":
            story += [kv(b["title"], b["rows"]), Spacer(1, 8)]
        elif t == "table":
            story.append(tabla(b))
            story.append(Paragraph(_mk(b["note"]), st["note"]) if b.get("note") else Spacer(1, 8))
        elif t == "img":
            png = png_de(b)
            if png:
                nfig += 1
                story.append(KeepTogether([imagen(png, b.get("width", 16)),
                                           Paragraph(f"IMAGEN {nfig}: {_mk(b['caption'])}", st["cap"])]))
        elif t == "pb":
            story.append(PageBreak())
        elif t == "cierre":
            bloque = [Spacer(1, 10), Paragraph(escape(_cierre_txt(meta)), st["n"]),
                      Paragraph("Atentamente.", st["right"])]
            if meta.get("firma") and FIRMA.exists():
                bloque.append(imagen(FIRMA.read_bytes(), 6.8))
            story.append(KeepTogether(bloque))

    class Doc(SimpleDocTemplate):
        def afterFlowable(self, fl):
            x = getattr(fl, "_toc", None)
            if x and x[0] >= 0:
                self.notify("TOCEntry", (x[0], x[1], self.page))

    def portada(c, doc):
        c.saveState()
        foto = ImageReader(io.BytesIO(_media("image2.jpeg")))
        iw, ih = foto.getSize()
        s = max(W / iw, H / ih)
        p = c.beginPath()
        p.rect(0, 0, W, H)
        c.clipPath(p, stroke=0, fill=0)
        c.drawImage(foto, (W - iw * s) / 2, (H - ih * s) / 2, width=iw * s, height=ih * s)
        c.setFillColor(colors.Color(0, 0, 0, alpha=.28))
        c.rect(0, 0, W, H, stroke=0, fill=1)
        c.drawImage(ImageReader(io.BytesIO(assets.vineta_bytes())), .35 * cm, H - 1.45 * cm, width=1.1 * cm,
                    height=1.1 * cm)
        bx, by, bw, bh = .45 * cm, H - 5.2 * cm - 7.6 * cm, 7.9 * cm, 7.6 * cm
        c.setFillColor(HR(ROJO))
        c.rect(bx, by, bw, bh, stroke=0, fill=1)
        lineas = _lineas(meta.get("titulo"), 4) or ["INFORME"]
        fs = min(24, int(7.0 * cm / (0.64 * max(len(x) for x in lineas))))
        c.setFillColor(colors.white)
        c.setFont("Helvetica-Bold", fs)
        total = len(lineas) * fs * 1.55
        y = by + bh / 2 + total / 2 - fs * 1.1
        for k, ln in enumerate(lineas):
            c.drawCentredString(bx + bw / 2, y, ln.upper())
            if k == 0:
                tw = c.stringWidth(ln.upper(), "Helvetica-Bold", fs)
                c.setStrokeColor(colors.white)
                c.setLineWidth(1.3)
                c.line(bx + bw / 2 - tw / 2, y - 3, bx + bw / 2 + tw / 2, y - 3)
            y -= fs * 1.55
        c.setFillColor(HR("#000000"))
        c.rect(0, 3.0 * cm, W, 6.9 * cm, stroke=0, fill=1)
        c.setFillColor(colors.white)
        ytop = 3.0 * cm + 6.9 * cm - 1.2 * cm
        c.setFont("Helvetica-Bold", 14)
        c.drawString(.9 * cm, ytop, (meta.get("cliente") or "").upper())
        c.setFont("Helvetica", 12)
        c.drawString(.9 * cm, ytop - .85 * cm, (meta.get("unidad") or "").upper())
        c.drawRightString(W - 1 * cm, ytop - 1.75 * cm, (meta.get("periodo") or "").upper())
        y = ytop - 2.9 * cm
        pt = ParagraphStyle("pc", fontName="Helvetica", fontSize=8.5, leading=11.5, textColor=colors.white)
        for etq, ls in (("Dirigido", _lineas(meta.get("dirigido"), 2)), ("Copia", _lineas(meta.get("copia"), 3))):
            for k, ln in enumerate(ls or [""]):
                p1 = Paragraph(f"<b>{etq if k == 0 else ''}</b>", pt)
                p1.wrapOn(c, 2.2 * cm, 1 * cm)
                p1.drawOn(c, .9 * cm, y - 8)
                p2 = Paragraph(((": " if k == 0 else "") + escape(ln)), pt)
                _, h2 = p2.wrap(W - 4.2 * cm, 3 * cm)
                p2.drawOn(c, 2.9 * cm, y - h2 + 3)
                y -= max(h2, 11.5) + 3
            y -= 6
        c.setFillColor(HR(BARRA))
        c.rect(0, 0, W, 1.5 * cm, stroke=0, fill=1)
        c.setFillColor(colors.white)
        c.setFont("Helvetica", 8.5)
        c.drawString(1.2 * cm, .6 * cm, "ESTRICTAMENTE CONFIDENCIAL")
        c.restoreState()

    def interior(c, doc):
        c.saveState()
        c.drawImage(ImageReader(io.BytesIO(_media("image18.png"))), 2 * cm, H - 2.45 * cm, width=3.4 * cm,
                    height=3.4 * cm * 406 / 1339)
        c.setFont("Helvetica-Bold", 10)
        c.setFillColor(HR(NEGRO))
        c.drawRightString(W - 2 * cm, H - 1.55 * cm, "Abrimos paso")
        c.drawRightString(W - 2 * cm, H - 1.95 * cm, "al progreso")
        c.setStrokeColor(HR("#9A9A9A"))
        c.setLineWidth(.8)
        c.line(2 * cm, H - 2.7 * cm, W - 2 * cm, H - 2.7 * cm)
        c.setFillColor(HR(BARRA))
        c.rect(0, 0, W, 1.5 * cm, stroke=0, fill=1)
        c.setFillColor(colors.white)
        c.setFont("Helvetica", 8)
        c.drawString(1.2 * cm, .85 * cm, f"INFORME TÉCNICO {pie}")
        c.drawString(1.2 * cm, .4 * cm, "Documento Estrictamente Confidencial")
        c.drawRightString(W - 1.5 * cm, .85 * cm, f"Pág. {doc.page}")
        c.restoreState()

    buf = io.BytesIO()
    try:
        Doc(buf, pagesize=A4, leftMargin=2 * cm, rightMargin=2 * cm, topMargin=3.1 * cm, bottomMargin=2.2 * cm,
            title=" ".join(_lineas(meta.get("titulo"))) or "Informe", author=meta.get("elaborado_por") or "FAMESA"
            ).multiBuild(story, onFirstPage=portada, onLaterPages=interior)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    return buf.getvalue()


# =============================================================================== WORD
def a_docx(data, meta):
    from docx import Document
    from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt, RGBColor

    XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
    d = Document(str(PLANTILLA))
    body = d.element.body
    sect = body.find(qn("w:sectPr"))
    kids = list(body)
    P_H, P_T, P_C, P_B, P_S = kids[-6:-1]      # modelos: título, texto, pie de imagen, viñeta, subtítulo
    for p in (P_H, P_T, P_C, P_B, P_S):
        body.remove(p)
    portada = kids[:2]
    sdt = kids[3]

    def jt(p):
        """Texto del párrafo solo con sus propias corridas (los cuadros de texto anidados duplican itertext)."""
        out = []
        for r in p.findall(qn("w:r")):
            for e in r:
                if e.tag == qn("w:t"):
                    out.append(e.text or "")
                elif e.tag == qn("w:tab"):
                    out.append("\t")
        return "".join(out).strip()

    def set_text(p, text):
        runs = [r for r in p.findall(qn("w:r")) if r.find(qn("w:t")) is not None]
        if not runs:
            return
        t = runs[0].find(qn("w:t"))
        t.text = text
        t.set(XML_SPACE, "preserve")
        for r in runs[1:]:
            r.getparent().remove(r)

    def set_label(p, valor):
        runs = p.findall(qn("w:r"))
        idx = next((i for i, r in enumerate(runs) if ":" in "".join(r.itertext())), None)
        if idx is None:
            return
        t = runs[idx].find(qn("w:t"))
        t.text = ": " + valor
        t.set(XML_SPACE, "preserve")
        for r in runs[idx + 1:]:
            r.getparent().remove(r)

    # ---- portada (textos de tu plantilla reemplazados por los de la app)
    tit = _lineas(meta.get("titulo"), 3) + ["", "", ""]
    dirigido = _lineas(meta.get("dirigido"), 1) or [""]
    copia = _lineas(meta.get("copia"), 2) + ["", ""]
    for raiz in portada:
        for tb in raiz.iter(qn("w:txbxContent")):
            ps = list(tb.findall(qn("w:p")))
            textos = [jt(p) for p in ps]
            icop = next((i for i, t in enumerate(textos) if t.startswith("Copia")), None)
            for i, p in enumerate(ps):
                t = textos[i]
                if t.startswith("INFORME") and "MENSUAL" in t:
                    set_text(p, tit[0].upper())
                elif t.startswith("ASISTENCIA"):
                    set_text(p, tit[1].upper())
                elif t.startswith("UNDERGROUND"):
                    set_text(p, tit[2].upper())
                elif t.startswith("KEVIN SERRANO"):
                    set_text(p, (meta.get("elaborado_por") or "").upper())
                elif t.startswith("TRAFIGURA") or t == "CLIENTE":
                    set_text(p, (meta.get("cliente") or "").upper())
                elif t.startswith("UNIDAD MINERA"):
                    set_text(p, (meta.get("unidad") or "").upper())
                elif t.startswith("SETIEMBRE"):
                    set_text(p, (meta.get("periodo") or "").upper())
                elif t.startswith("Dirigido"):
                    set_label(p, dirigido[0])
                elif t.startswith("Copia"):
                    set_label(p, copia[0])
                elif icop is not None and i == icop + 1:
                    vivo = [r for r in p.findall(qn("w:r"))
                            if r.find(qn("w:t")) is not None and (r.find(qn("w:t")).text or "").strip()]
                    if vivo:
                        t0 = vivo[0].find(qn("w:t"))
                        t0.text = ("  " + copia[1]) if copia[1] else ""
                        t0.set(XML_SPACE, "preserve")
                        for r in vivo[1:]:
                            r.getparent().remove(r)

    # ---- pie de página
    pie = f"INFORME TÉCNICO {meta.get('pie') or ''}".strip()
    for rel in d.part.rels.values():
        if rel.reltype.endswith("/footer"):
            for p in rel.target_part.element.iter(qn("w:p")):
                if jt(p).startswith("INFORME"):
                    set_text(p, pie)

    # ---- índice (se refresca al abrir en Word)
    entradas = [(0 if b["t"] == "h1" else 1, b["text"]) for b in data["blocks"] if b["t"] in ("h1", "h2")]
    cont = sdt.find(qn("w:sdtContent"))
    for ch in list(cont):
        cont.remove(ch)

    def p_toc(texto, nivel, ini=False, fin=False):
        p = OxmlElement("w:p")
        ppr = OxmlElement("w:pPr")
        ps_ = OxmlElement("w:pStyle")
        ps_.set(qn("w:val"), "TDC1")
        ppr.append(ps_)
        if nivel:
            ind = OxmlElement("w:ind")
            ind.set(qn("w:left"), "340")
            ppr.append(ind)
        p.append(ppr)

        def run(**k):
            r = OxmlElement("w:r")
            if "fld" in k:
                e = OxmlElement("w:fldChar")
                e.set(qn("w:fldCharType"), k["fld"])
                r.append(e)
            if "instr" in k:
                e = OxmlElement("w:instrText")
                e.set(XML_SPACE, "preserve")
                e.text = k["instr"]
                r.append(e)
            if "text" in k:
                e = OxmlElement("w:t")
                e.set(XML_SPACE, "preserve")
                e.text = k["text"]
                r.append(e)
            p.append(r)
        if ini:
            run(fld="begin")
            run(instr=' TOC \\o "1-3" \\h \\z \\u ')
            run(fld="separate")
        run(text=texto)
        if fin:
            run(fld="end")
        return p

    for i, (nv, tx) in enumerate(entradas):
        cont.append(p_toc(tx, nv, ini=(i == 0), fin=(i == len(entradas) - 1)))
    upd = OxmlElement("w:updateFields")
    upd.set(qn("w:val"), "true")
    d.settings.element.append(upd)

    # ---- utilidades de cuerpo
    ids = [9000]

    def emit(el):
        sect.addprevious(el)
        return el

    def limpia(el):
        for p in el.iter(qn("w:p")):
            for a in list(p.attrib):
                if "paraId" in a or "textId" in a:
                    del p.attrib[a]
        for e in el.iter():
            if etree.QName(e).localname in ("docPr", "cNvPr"):
                ids[0] += 1
                e.set("id", str(ids[0]))
        return el

    def base_rpr(proto):
        r = proto.find(qn("w:r"))
        rp = r.find(qn("w:rPr")) if r is not None else None
        return copy.deepcopy(rp) if rp is not None else OxmlElement("w:rPr")

    def rellena(p, texto, proto, size=None, bold=False, italic=False):
        for ch in list(p):
            if ch.tag != qn("w:pPr"):
                p.remove(ch)
        rp0 = base_rpr(proto)
        for tok in re.split(r"(\*\*.+?\*\*|\[\[.+?\]\])", str(texto)):
            if not tok:
                continue
            rojo, neg = tok.startswith("[["), tok.startswith("**")
            r = OxmlElement("w:r")
            rp = copy.deepcopy(rp0)
            if (bold or rojo or neg) and rp.find(qn("w:b")) is None:
                rp.insert(1 if rp.find(qn("w:rFonts")) is not None else 0, OxmlElement("w:b"))
            if italic and rp.find(qn("w:i")) is None:
                rp.append(OxmlElement("w:i"))
            if rojo:
                c = OxmlElement("w:color")
                c.set(qn("w:val"), "E10613")
                rp.append(c)
            if size:
                for tag in ("w:sz", "w:szCs"):
                    e = rp.find(qn(tag))
                    if e is None:
                        e = OxmlElement(tag)
                        rp.append(e)
                    e.set(qn("w:val"), str(int(size * 2)))
            r.append(rp)
            t = OxmlElement("w:t")
            t.set(XML_SPACE, "preserve")
            t.text = tok[2:-2] if (rojo or neg) else tok
            r.append(t)
            p.append(r)
        return p

    def espaciado(p, line=360, before=None, after=120):
        ppr = p.find(qn("w:pPr"))
        sp = ppr.find(qn("w:spacing"))
        if sp is None:
            sp = OxmlElement("w:spacing")
            ppr.insert(0, sp)
        sp.set(qn("w:line"), str(line))
        sp.set(qn("w:lineRule"), "auto")
        sp.set(qn("w:after"), str(after))
        if before is not None:
            sp.set(qn("w:before"), str(before))

    def parrafo(texto, size=11):
        p = copy.deepcopy(P_T)
        rellena(p, texto, P_T, size=size)
        espaciado(p)
        return emit(limpia(p))

    def titulo1(texto):
        p = copy.deepcopy(P_H)
        for q in p.iter(qn("w:p")):
            set_text(q, texto.upper())
        emit(limpia(p))

    def titulo2(texto):
        p = copy.deepcopy(P_T)
        rellena(p, texto, P_T, size=12.5, bold=True)
        ppr = p.find(qn("w:pPr"))
        ppr.insert(0, OxmlElement("w:keepNext"))
        jc = ppr.find(qn("w:jc"))
        if jc is not None:
            jc.set(qn("w:val"), "left")
        ol = OxmlElement("w:outlineLvl")
        ol.set(qn("w:val"), "1")
        rp = ppr.find(qn("w:rPr"))
        (rp.addprevious(ol) if rp is not None else ppr.append(ol))
        espaciado(p, before=200, after=80)
        emit(limpia(p))

    def titulo3(texto):
        p = copy.deepcopy(P_S)
        rellena(p, texto.upper(), P_S, size=10.5)
        p.find(qn("w:pPr")).insert(0, OxmlElement("w:keepNext"))
        espaciado(p, line=300, before=160, after=60)
        emit(limpia(p))

    def vineta(texto):
        p = copy.deepcopy(P_B)
        rellena(p, texto, P_B, size=11)
        espaciado(p, line=340, before=80, after=60)
        emit(limpia(p))

    def nota(texto):
        p = d.add_paragraph()
        p.paragraph_format.space_after = Pt(6)
        r = p.add_run(texto)
        r.italic = True
        r.font.size = Pt(8.5)
        r.font.name = "Arial"
        r.font.color.rgb = RGBColor(0x6F, 0x73, 0x78)

    def sombra(cell, color):
        tcPr = cell._tc.get_or_add_tcPr()
        s = OxmlElement("w:shd")
        s.set(qn("w:val"), "clear")
        s.set(qn("w:color"), "auto")
        s.set(qn("w:fill"), color.lstrip("#"))
        tcPr.append(s)

    def bordes(t):
        b = OxmlElement("w:tblBorders")
        for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
            e = OxmlElement("w:" + edge)
            for k, v in (("val", "single"), ("sz", "4"), ("space", "0"), ("color", "8C8C8C")):
                e.set(qn("w:" + k), v)
            b.append(e)
        tblPr = t._tbl.tblPr
        look = tblPr.find(qn("w:tblLook"))
        (look.addprevious(b) if look is not None else tblPr.append(b))

    def escribe(cell, texto, size=8.5, bold=False, color=None, align=None):
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.0
        if align is not None:
            p.alignment = align
        for tok in re.split(r"(\*\*.+?\*\*|\[\[.+?\]\])", str(texto)):
            if not tok:
                continue
            rojo, neg = tok.startswith("[["), tok.startswith("**")
            r = p.add_run(tok[2:-2] if (rojo or neg) else tok)
            r.bold = bold or rojo or neg
            r.font.size = Pt(size)
            r.font.name = "Arial"
            c = "E10613" if rojo else color
            if c:
                r.font.color.rgb = RGBColor.from_string(c.lstrip("#"))

    def anchos(t, widths):
        t.autofit = False
        k = 16.5 / sum(widths)
        lay = OxmlElement("w:tblLayout")
        lay.set(qn("w:type"), "fixed")
        t._tbl.tblPr.append(lay)
        for col, w in zip(t.columns, widths):
            col.width = Cm(w * k)
        for row in t.rows:
            for c, w in zip(row.cells, widths):
                c.width = Cm(w * k)

    def tabla(b):
        head, rows = b["head"], b["rows"]
        w = b.get("widths") or [17 / len(head)] * len(head)
        al = b.get("align") or "l" + "c" * (len(head) - 1)
        t = d.add_table(rows=1, cols=len(head))
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        bordes(t)
        trPr = t.rows[0]._tr.get_or_add_trPr()
        th = OxmlElement("w:tblHeader")
        th.set(qn("w:val"), "true")
        trPr.append(th)
        for c, h in zip(t.rows[0].cells, head):
            sombra(c, "#E10613")
            escribe(c, h, bold=True, color="FFFFFF", align=WD_ALIGN_PARAGRAPH.CENTER)
        for k, r in enumerate(rows):
            cells = t.add_row().cells
            for j, (c, v) in enumerate(zip(cells, r)):
                txt, s_ = _cell(v)
                if k % 2:
                    sombra(c, "#F4F4F4")
                a = {"l": WD_ALIGN_PARAGRAPH.LEFT, "r": WD_ALIGN_PARAGRAPH.RIGHT}.get(al[j], WD_ALIGN_PARAGRAPH.CENTER)
                if s_:
                    escribe(c, txt, bold=True, color="1E7B34" if s_ == "ok" else "E10613",
                            align=WD_ALIGN_PARAGRAPH.CENTER)
                else:
                    escribe(c, txt, align=a)
        anchos(t, w)
        if b.get("note"):
            nota(b["note"])
        else:
            d.add_paragraph().paragraph_format.space_after = Pt(2)

    def kv(titulo, rows):
        t = d.add_table(rows=1, cols=2)
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        bordes(t)
        a = t.cell(0, 0).merge(t.cell(0, 1))
        sombra(a, "#E10613")
        escribe(a, titulo, 9.5, bold=True, color="FFFFFF", align=WD_ALIGN_PARAGRAPH.CENTER)
        for k, v in rows:
            c = t.add_row().cells
            sombra(c[0], "#EDEDED")
            escribe(c[0], k, 9.5, bold=True)
            escribe(c[1], v, 9.5)
        anchos(t, (6.2, 10.8))
        d.add_paragraph().paragraph_format.space_after = Pt(2)

    # ---- cuerpo
    nfig = 0
    for b in data["blocks"]:
        t = b["t"]
        if t == "h1":
            titulo1(b["text"])
        elif t == "h2":
            titulo2(b["text"])
        elif t == "h3":
            titulo3(b["text"])
        elif t == "p":
            parrafo(b["text"])
        elif t == "bullets":
            for it in b["items"]:
                vineta(it)
        elif t == "eq":
            p = d.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(b["text"])
            r.bold = True
            r.font.name = "Courier New"
            r.font.size = Pt(11)
        elif t == "kv":
            kv(b["title"], b["rows"])
        elif t == "table":
            tabla(b)
        elif t == "img":
            png = png_de(b)
            if png:
                nfig += 1
                d.add_picture(io.BytesIO(png), width=Cm(b.get("width", 16)))
                d.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
                d.paragraphs[-1].paragraph_format.keep_with_next = True
                cp = copy.deepcopy(P_C)
                rellena(cp, f"IMAGEN {nfig}: {b['caption']}", P_C)
                emit(limpia(cp))
        elif t == "pb":
            d.add_page_break()
        elif t == "cierre":
            parrafo(_cierre_txt(meta))
            ap = copy.deepcopy(P_T)
            rellena(ap, "Atentamente.", P_T, size=11)
            ap.find(qn("w:pPr")).find(qn("w:jc")).set(qn("w:val"), "right")
            emit(limpia(ap))
            if meta.get("firma") and FIRMA.exists():
                d.add_picture(io.BytesIO(FIRMA.read_bytes()), width=Cm(7))
                d.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    d.core_properties.title = " ".join(_lineas(meta.get("titulo"))) or "Informe"
    d.core_properties.author = meta.get("elaborado_por") or "FAMESA"
    out = io.BytesIO()
    d.save(out)
    return out.getvalue()
