"""Lee un .docx "normal" (Word o Google Docs) y lo convierte en una lista de bloques neutros.

El lector ignora a propósito casi todo el formato directo (fuentes, colores, tamaños,
sangrías): la plantilla oficial decide cómo se ve cada elemento. Solo se conserva el
formato *semántico*: negrita, cursiva, subrayado, superíndice/subíndice, enlaces y notas al pie.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any

import docx
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

from .xml import local

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
V_NS = "urn:schemas-microsoft-com:vml"

# --- convenciones de redacción (ver docs/GUIA_REDACCION.md) -----------------------------
RE_TITULO_OBJ = re.compile(r"^\s*(tabla|gr[aá]fico|figura|foto|mapa|cuadro)\s*(\d+)?\s*[:.\-–—]\s*", re.I)
RE_SUBTITULO = re.compile(r"^\s*(subt[ií]tulo|unidades?)\s*[:]\s*", re.I)
RE_PIE = re.compile(r"^\s*(fuentes?|elaboraci[oó]n propia)\s*[:.]", re.I)
RE_DESTACADO = re.compile(r"^\s*destacado\s*:\s*", re.I)
RE_NUM_MANUAL = re.compile(r"^\s*(?:\d{1,2}(?:\.\d{1,2})+\.?|\d{1,2}\.)\s*(?=[A-Za-zÁÉÍÓÚÑáéíóúñ¿¡(])")
# Título "escrito a mano" en un párrafo normal: "1. RESUMEN", "2.1. Objetivos", "3.OFERTA"
RE_TITULO_NUM = re.compile(r"^\s*(\d{1,2}(?:\s*\.\s*\d{1,2}){0,2})\s*\.\s*(?=[A-Za-zÁÉÍÓÚÑáéíóúñ¿¡(])(.{2,85})$")

NOMBRES_TITULO = {"title", "título", "titulo", "subtitle", "subtítulo", "subtitulo"}
NOMBRES_CAPTION = {"caption", "leyenda", "epígrafe", "epigrafe"}


@dataclass
class Documento:
    bloques: list[dict] = field(default_factory=list)
    avisos: list[str] = field(default_factory=list)


class Lector:
    def __init__(self, ruta: str, titulos_numerados: bool = False):
        self.titulos_numerados = titulos_numerados
        self.doc = docx.Document(ruta)
        self.out = Documento()
        self._num_cache: dict[tuple[str, int], bool] = {}
        self._notas = self._cargar_notas()

    # ------------------------------------------------------------------ utilidades
    def aviso(self, msg: str) -> None:
        if msg not in self.out.avisos:
            self.out.avisos.append(msg)

    def _cargar_notas(self) -> dict[str, Any]:
        try:
            part = self.doc.part.part_related_by(RT.FOOTNOTES)
        except KeyError:
            return {}
        from lxml import etree
        root = etree.fromstring(part.blob)
        return {n.get(qn("w:id")): n for n in root.findall(qn("w:footnote"))}

    def _estilo(self, p) -> str:
        try:
            return (Paragraph(p, self.doc._body).style.name or "").strip().lower()
        except Exception:
            return ""

    # ------------------------------------------------------------------ listas
    def _numpr(self, p):
        """(numId, ilvl) del párrafo, mirando formato directo y luego el estilo."""
        ppr = p.find(qn("w:pPr"))
        numpr = ppr.find(qn("w:numPr")) if ppr is not None else None
        ilvl = 0
        num_id = None
        if numpr is not None:
            n = numpr.find(qn("w:numId"))
            l = numpr.find(qn("w:ilvl"))
            num_id = n.get(qn("w:val")) if n is not None else None
            ilvl = int(l.get(qn("w:val"))) if l is not None else 0
        if num_id in (None, "0"):
            try:
                st = Paragraph(p, self.doc._body).style
                while st is not None and num_id in (None, "0"):
                    sp = st.element.find(qn("w:pPr"))
                    sn = sp.find(qn("w:numPr")) if sp is not None else None
                    if sn is not None:
                        n = sn.find(qn("w:numId"))
                        l = sn.find(qn("w:ilvl"))
                        num_id = n.get(qn("w:val")) if n is not None else None
                        if l is not None and numpr is None:
                            ilvl = int(l.get(qn("w:val")))
                    st = st.base_style
            except Exception:
                pass
        if num_id in (None, "0"):
            return None
        return num_id, ilvl

    def _es_ordenada(self, num_id: str, ilvl: int) -> bool:
        key = (num_id, ilvl)
        if key in self._num_cache:
            return self._num_cache[key]
        ordenada = False
        try:
            root = self.doc.part.numbering_part.element
            num = next(n for n in root.findall(qn("w:num")) if n.get(qn("w:numId")) == num_id)
            abs_id = num.find(qn("w:abstractNumId")).get(qn("w:val"))
            absn = next(a for a in root.findall(qn("w:abstractNum")) if a.get(qn("w:abstractNumId")) == abs_id)
            for lvl in absn.findall(qn("w:lvl")):
                if int(lvl.get(qn("w:ilvl"))) == ilvl:
                    fmt = lvl.find(qn("w:numFmt"))
                    ordenada = fmt is not None and fmt.get(qn("w:val")) not in ("bullet", "none")
                    break
        except Exception:
            ordenada = False
        self._num_cache[key] = ordenada
        return ordenada

    # ------------------------------------------------------------------ runs
    def _fmt(self, r) -> dict:
        rpr = r.find(qn("w:rPr"))
        f = {"b": False, "i": False, "u": False, "sup": False, "sub": False}
        if rpr is None:
            return f

        def on(tag):
            e = rpr.find(qn(tag))
            return e is not None and e.get(qn("w:val"), "true") not in ("0", "false", "off")

        f["b"], f["i"] = on("w:b"), on("w:i")
        u = rpr.find(qn("w:u"))
        f["u"] = u is not None and u.get(qn("w:val")) not in (None, "none")
        va = rpr.find(qn("w:vertAlign"))
        if va is not None:
            f["sup"] = va.get(qn("w:val")) == "superscript"
            f["sub"] = va.get(qn("w:val")) == "subscript"
        rs = rpr.find(qn("w:rStyle"))
        if rs is not None:
            name = rs.get(qn("w:val"), "").lower()
            f["b"] |= name in ("strong", "textoennegrita")
            f["i"] |= name in ("emphasis", "nfasis", "énfasis")
        return f

    def _runs_de(self, nodo, href: str | None = None) -> list[dict]:
        """Recorre hijos de un párrafo/hipervínculo y devuelve runs normalizados."""
        runs: list[dict] = []
        for child in nodo:
            t = child.tag
            if t == qn("w:r"):
                runs.extend(self._run(child, href))
            elif t == qn("w:hyperlink"):
                rid = child.get(f"{{{R_NS}}}id")
                url = None
                if rid and rid in self.doc.part.rels:
                    url = self.doc.part.rels[rid].target_ref
                elif child.get(qn("w:anchor")):
                    url = None  # enlace interno: se conserva el texto sin enlace
                runs.extend(self._runs_de(child, url or href))
            elif t in (qn("w:ins"), qn("w:smartTag"), qn("w:fldSimple"), qn("w:sdt")):
                inner = child.find(qn("w:sdtContent")) if t == qn("w:sdt") else child
                if inner is not None:
                    runs.extend(self._runs_de(inner, href))
            # w:del, w:proofErr, w:bookmark*, etc. se ignoran
        return runs

    def _run(self, r, href) -> list[dict]:
        out: list[dict] = []
        fmt = self._fmt(r)
        buf = ""

        def flush():
            nonlocal buf
            if buf:
                out.append({"text": buf, **fmt, "href": href})
                buf = ""

        for c in r:
            t = c.tag
            if t == qn("w:t"):
                buf += c.text or ""
            elif t == qn("w:tab"):
                buf += " "
            elif t in (qn("w:br"), qn("w:cr")):
                if c.get(qn("w:type")) in ("page", "column"):
                    continue
                buf += "\n"
            elif t == qn("w:noBreakHyphen"):
                buf += "‑"
            elif t == qn("w:footnoteReference"):
                flush()
                nota = self._nota(c.get(qn("w:id")))
                if nota:
                    out.append({"fn": nota})
        flush()
        return out

    def _nota(self, fid: str) -> list[dict] | None:
        n = self._notas.get(fid)
        if n is None:
            return None
        runs: list[dict] = []
        for p in n.findall(qn("w:p")):
            rs = [x for x in self._runs_de(p) if "text" in x]
            if runs and rs:
                runs.append({"text": " ", "b": False, "i": False, "u": False,
                             "sup": False, "sub": False, "href": None})
            runs.extend(rs)
        return runs

    # ------------------------------------------------------------------ imágenes
    def _imagenes(self, p) -> list[dict]:
        imgs: list[dict] = []
        for dr in p.iter(qn("w:drawing")):
            if dr.find(f".//{{{C_NS}}}chart") is not None:
                self.aviso("Hay un gráfico nativo de Word/Google en el documento: no se puede "
                           "trasladar. Exporta el gráfico como imagen (PNG) y pégalo como imagen.")
                continue
            blip = dr.find(f".//{{{A_NS}}}blip")
            if blip is None:
                continue
            rid = blip.get(f"{{{R_NS}}}embed")
            ext_el = dr.find(f".//{{{WP_NS}}}extent")
            cx = int(ext_el.get("cx")) if ext_el is not None else None
            cy = int(ext_el.get("cy")) if ext_el is not None else None
            doc_pr = dr.find(f".//{{{WP_NS}}}docPr")
            alt = (doc_pr.get("descr") or doc_pr.get("name")) if doc_pr is not None else ""
            self._add_img(imgs, rid, cx, cy, alt)
        for im in p.iter(f"{{{V_NS}}}imagedata"):  # imágenes antiguas (VML)
            self._add_img(imgs, im.get(f"{{{R_NS}}}id"), None, None, "")
        return imgs

    def _add_img(self, imgs, rid, cx, cy, alt):
        if not rid or rid not in self.doc.part.related_parts:
            return
        part = self.doc.part.related_parts[rid]
        imgs.append({"k": "img", "blob": part.blob, "ctype": part.content_type,
                     "cx": cx, "cy": cy, "alt": alt})

    # ------------------------------------------------------------------ tablas
    def _tabla(self, tbl) -> dict:
        filas = []
        for tr in tbl.findall(qn("w:tr")):
            fila = []
            for tc in tr.findall(qn("w:tc")):
                tcpr = tc.find(qn("w:tcPr"))
                span, vm = 1, None
                if tcpr is not None:
                    gs = tcpr.find(qn("w:gridSpan"))
                    if gs is not None:
                        span = int(gs.get(qn("w:val")))
                    v = tcpr.find(qn("w:vMerge"))
                    if v is not None:
                        vm = "restart" if v.get(qn("w:val")) == "restart" else "cont"
                paras = []
                for p in tc.findall(qn("w:p")):
                    runs = self._runs_de(p)
                    paras.append(runs)
                if tc.find(f".//{qn('w:drawing')}") is not None:
                    self.aviso("Una imagen dentro de una celda de tabla se ha omitido.")
                fila.append({"paras": paras, "span": span, "vmerge": vm})
            filas.append(fila)
        if len(list(tbl.iter(qn("w:tbl")))) > 1:
            self.aviso("Hay tablas anidadas: solo se traslada el nivel exterior.")
        return {"k": "tbl", "rows": filas}

    # ------------------------------------------------------------------ bucle principal
    def leer(self) -> Documento:
        for el in self.doc.element.body:
            t = local(el.tag)
            if t == "p":
                self._parrafo(el)
            elif t == "tbl":
                self.out.bloques.append(self._tabla(el))
            elif t == "sdt":  # tabla de contenidos, etc.
                continue
        self._resolver_captions()
        return self.out

    def _parrafo(self, p) -> None:
        estilo = self._estilo(p)
        # Índice del documento original: la plantilla genera el suyo.
        if estilo.startswith(("toc", "tdc", "table of contents", "contents", "índice", "indice")) \
                and estilo not in ("título índice",):
            return
        imgs = self._imagenes(p)
        runs = [r for r in self._runs_de(p)]
        texto = "".join(r.get("text", "") for r in runs).strip()

        if imgs and not texto:
            self.out.bloques.extend(imgs)
            return

        if not texto and not any("fn" in r for r in runs):
            return

        if estilo in NOMBRES_TITULO:
            self.aviso("Se ha ignorado un párrafo con estilo 'Título/Subtítulo': el título del "
                       "estudio se define en portada.toml.")
            return

        # --- encabezados
        m = re.match(r"^(?:heading|título|titulo|encabezado)\s*(\d)$", estilo)
        nivel = int(m.group(1)) if m else None
        if nivel is None:
            ppr = p.find(qn("w:pPr"))
            ol = ppr.find(qn("w:outlineLvl")) if ppr is not None else None
            if ol is not None and int(ol.get(qn("w:val"))) < 9:
                nivel = int(ol.get(qn("w:val"))) + 1
        if nivel:
            if nivel > 4:
                self.aviso("Hay títulos de nivel 5 o más: se han tratado como nivel 4.")
                nivel = 4
            runs = _quitar_prefijo(runs, RE_NUM_MANUAL)
            self.out.bloques.append({"k": "h", "lvl": nivel, "runs": runs})
            self.out.bloques.extend(imgs)
            return

        # --- títulos escritos a mano ("1. RESUMEN") en párrafos normales (opcional)
        if self.titulos_numerados and self._numpr(p) is None:
            m = RE_TITULO_NUM.match(texto)
            if m and not texto.rstrip().endswith((".", ":", ";", ",")):
                nivel = len(re.findall(r"\d+", m.group(1)))
                if nivel == 1 and m.group(1).strip() == "":
                    nivel = 1
                runs = _quitar_prefijo(runs, RE_NUM_MANUAL)
                if texto.upper() == texto:  # TODO EN MAYÚSCULAS → capitalización normal
                    runs = [dict(r, text=r["text"].lower()) if "text" in r else r for r in runs]
                    primero = next((r for r in runs if "text" in r and r["text"].strip()), None)
                    if primero is not None:
                        t = primero["text"]
                        i = len(t) - len(t.lstrip())
                        primero["text"] = t[:i] + t[i:i + 1].upper() + t[i + 1:]
                self.out.bloques.append({"k": "h", "lvl": nivel, "runs": runs})
                self.out.bloques.extend(imgs)
                return

        # --- listas
        np = self._numpr(p)
        if np is not None or estilo.startswith(("list bullet", "list number", "lista con", "lista numerada")):
            num_id, ilvl = np if np else ("0", 0)
            ordenada = (self._es_ordenada(num_id, ilvl) if np else estilo.startswith(("list number", "lista numerada")))
            if ilvl > 2:
                ilvl = 2
            self.out.bloques.append({"k": "li", "lvl": ilvl, "ordered": ordenada,
                                     "num_id": num_id, "runs": runs})
            self.out.bloques.extend(imgs)
            return

        # --- convenciones por prefijo
        if estilo in NOMBRES_CAPTION:
            self.out.bloques.append({"k": "caption", "runs": runs})
        elif RE_DESTACADO.match(texto):
            self.out.bloques.append({"k": "dest", "runs": _quitar_prefijo(runs, RE_DESTACADO)})
        elif RE_SUBTITULO.match(texto):
            self.out.bloques.append({"k": "sub", "runs": _quitar_prefijo(runs, RE_SUBTITULO)})
        elif RE_TITULO_OBJ.match(texto):
            self.out.bloques.append({"k": "title", "runs": _quitar_prefijo(runs, RE_TITULO_OBJ)})
        elif RE_PIE.match(texto):
            self.out.bloques.append({"k": "pie", "runs": runs})
        else:
            self.out.bloques.append({"k": "p", "runs": runs})
        self.out.bloques.extend(imgs)

    def _resolver_captions(self) -> None:
        """Un 'Caption' de Word/Google va antes (título) o después (pie) del objeto."""
        b = self.out.bloques
        for i, bl in enumerate(b):
            if bl["k"] != "caption":
                continue
            antes = b[i - 1]["k"] if i > 0 else None
            despues = b[i + 1]["k"] if i + 1 < len(b) else None
            if antes in ("tbl", "img") and despues not in ("tbl", "img"):
                bl["k"] = "pie"
            else:
                bl["k"] = "title"
                bl["runs"] = _quitar_prefijo(bl["runs"], RE_TITULO_OBJ)


def _quitar_prefijo(runs: list[dict], regex: re.Pattern) -> list[dict]:
    """Quita del inicio del texto lo que case con ``regex``, repartido entre runs."""
    texto = "".join(r.get("text", "") for r in runs)
    m = regex.match(texto)
    if not m:
        return runs
    quitar = m.end()
    out = []
    for r in runs:
        if "text" not in r or quitar <= 0:
            out.append(r)
            continue
        n = len(r["text"])
        if quitar >= n:
            quitar -= n
            continue
        r = dict(r, text=r["text"][quitar:])
        quitar = 0
        out.append(r)
    return out
