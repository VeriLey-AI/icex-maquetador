"""Vuelca los bloques del lector dentro de la plantilla oficial de ICEX.

La plantilla se usa como "molde": se conservan portada, contraportada, cabeceras, pies,
secciones, estilos y numeración. Solo se sustituye lo que hay entre el título «Índice» y el
final de la sección principal (índice de ejemplo + capítulos de ejemplo).
"""
from __future__ import annotations

import io
import re

import docx
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import parse_xml
from docx.oxml.ns import qn, nsdecls
from docx.shared import Emu
from lxml import etree

from . import portada
from .lector import Documento
from .xml import el, has_sectpr, paragraph_style, sub, text_run

# Ancho útil de la página (A4, márgenes 2 cm) y de las tablas de la plantilla, en twips (dxa)
ANCHO_UTIL = 9632
ANCHO_TABLA_STD = 7710
EMU_POR_CM = 360000

ESTILO_H = {1: "Ttulo1", 2: "Ttulo2", 3: "Ttulo3", 4: "Ttulo4"}
ESTILO_LI = {0: "Prrafodelista", 1: "Prrafodelista2", 2: "Prrafodelista3"}


class Escritor:
    def __init__(self, ruta_plantilla: str, max_img_cm: float = 15.5):
        self.tpl = docx.Document(ruta_plantilla)
        self.body = self.tpl.element.body
        self.max_img = int(max_img_cm * EMU_POR_CM)
        self.avisos: list[str] = []
        self._fn_nuevas: list[tuple[int, list[dict]]] = []
        self._fn_next = 1
        self.stats = {"capitulos": 0, "tablas": 0, "imagenes": 0, "palabras": 0, "notas": 0}
        self._rn: dict[int, int] = {}  # contadores de listas numeradas por nivel
        self._rn_id: str | None = None

    # ------------------------------------------------------------------ localizar zonas
    def _localizar(self):
        hijos = list(self.body)
        titulo_idx = next((i for i, e in enumerate(hijos)
                           if e.tag == qn("w:p") and paragraph_style(e) == "Ttulondice"), None)
        if titulo_idx is None:
            raise SystemExit("La plantilla no tiene el título 'Índice' (estilo Ttulondice): "
                             "¿es la plantilla oficial de ICEX 2024?")
        sect_idx = next((i for i in range(titulo_idx + 1, len(hijos))
                         if hijos[i].tag == qn("w:p") and has_sectpr(hijos[i])), None)
        if sect_idx is None:
            raise SystemExit("No se ha encontrado el salto de sección tras el cuerpo de la plantilla.")
        return hijos, titulo_idx, sect_idx

    # ------------------------------------------------------------------ construcción de runs
    def _rpr_de(self, r: dict) -> list:
        props = []
        if r.get("href"):
            props.append(el("w:rStyle", val="Hipervnculo"))
        if r.get("b"):
            props.append(el("w:b"))
        if r.get("i"):
            props.append(el("w:i"))
        if r.get("u") and not r.get("href"):
            props.append(el("w:u", val="single"))
        if r.get("sup"):
            props.append(el("w:vertAlign", val="superscript"))
        if r.get("sub"):
            props.append(el("w:vertAlign", val="subscript"))
        return props

    def _runs_xml(self, runs: list[dict]) -> list:
        # recorta espacios sobrantes en los extremos del párrafo
        runs = [dict(r) for r in runs]
        textos = [r for r in runs if "text" in r]
        if textos:
            textos[0]["text"] = textos[0]["text"].lstrip()
            textos[-1]["text"] = textos[-1]["text"].rstrip()
        out = []
        for r in runs:
            if "fn" in r:
                out.append(self._ref_nota(r["fn"]))
                continue
            if not r["text"]:
                continue
            self.stats["palabras"] += len(r["text"].split())
            run = text_run(r["text"], self._rpr_de(r))
            if r.get("href"):
                rid = self.tpl.part.relate_to(r["href"], RT.HYPERLINK, is_external=True)
                h = el("w:hyperlink", history="1")
                h.set(qn("r:id"), rid)
                h.append(run)
                out.append(h)
            else:
                out.append(run)
        return out

    def _p(self, runs, estilo: str | None = None, ppr_extra: list | None = None):
        p = el("w:p")
        if estilo or ppr_extra:
            ppr = sub(p, "w:pPr")
            if estilo:
                sub(ppr, "w:pStyle", val=estilo)
            for e in ppr_extra or []:
                ppr.append(e)
        for x in self._runs_xml(runs):
            p.append(x)
        return p

    # ------------------------------------------------------------------ notas al pie
    def _ref_nota(self, runs: list[dict]):
        fid = self._fn_next
        self._fn_next += 1
        self._fn_nuevas.append((fid, runs))
        self.stats["notas"] += 1
        r = el("w:r")
        rpr = sub(r, "w:rPr")
        sub(rpr, "w:rStyle", val="Refdenotaalpie")
        ref = sub(r, "w:footnoteReference")
        ref.set(qn("w:id"), str(fid))
        return r

    def _volcar_notas(self):
        try:
            part = self.tpl.part.part_related_by(RT.FOOTNOTES)
        except KeyError:
            if self._fn_nuevas:
                self.avisos.append("La plantilla no tiene notas al pie: se han omitido.")
            return
        root = etree.fromstring(part.blob)
        for fn in list(root.findall(qn("w:footnote"))):  # fuera las notas de ejemplo
            if fn.get(qn("w:type")) is None:
                root.remove(fn)
        for fid, runs in self._fn_nuevas:
            f = etree.SubElement(root, qn("w:footnote"))
            f.set(qn("w:id"), str(fid))
            p = etree.SubElement(f, qn("w:p"))
            ppr = etree.SubElement(p, qn("w:pPr"))
            etree.SubElement(ppr, qn("w:pStyle")).set(qn("w:val"), "Textonotapie")
            r = etree.SubElement(p, qn("w:r"))
            rpr = etree.SubElement(r, qn("w:rPr"))
            etree.SubElement(rpr, qn("w:rStyle")).set(qn("w:val"), "Refdenotaalpie")
            etree.SubElement(r, qn("w:footnoteRef"))
            p.append(text_run(" "))
            for x in self._runs_xml([dict(r2, text=r2["text"]) for r2 in runs if "text" in r2]):
                p.append(x)
        part._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

    # ------------------------------------------------------------------ bloques
    def _reset_num(self):
        self._rn, self._rn_id = {}, None

    def _bloque(self, b: dict, i: int, bloques: list[dict]):
        k = b["k"]
        sig = bloques[i + 1]["k"] if i + 1 < len(bloques) else None
        if k != "li":
            self._reset_num()
        if k == "h":
            if b["lvl"] == 1:
                self.stats["capitulos"] += 1
            return [self._p(b["runs"], ESTILO_H[b["lvl"]])]
        if k == "p":
            return [self._p(b["runs"])]
        if k == "dest":
            return [self._p(b["runs"], "Destacado")]
        if k == "title":
            return [self._p(b["runs"], "TtuloTablas", [el("w:keepNext")])]
        if k == "sub":
            return [self._p(b["runs"], "Subttuloparatablasygrficos", [el("w:keepNext")])]
        if k == "pie":
            return [self._p(b["runs"], "Piedetablaygrfico")]
        if k == "li":
            return [self._lista(b, sig)]
        if k == "img":
            return self._imagen(b, sig)
        if k == "tbl":
            t = self._tabla(b, sig)
            return t
        return []

    def _lista(self, b: dict, sig: str | None):
        if not b["ordered"]:
            self._reset_num()
            return self._p(b["runs"], ESTILO_LI[b["lvl"]])
        lvl = b["lvl"]
        if self._rn_id != b["num_id"] and lvl == 0:
            self._rn = {}
        self._rn_id = b["num_id"]
        self._rn[lvl] = self._rn.get(lvl, 0) + 1
        for deeper in [x for x in self._rn if x > lvl]:
            del self._rn[deeper]
        n = self._rn[lvl]
        marca = [f"{n}.", f"{chr(96 + ((n - 1) % 26) + 1)})", f"{n})"][lvl]
        izq = 454 + 340 * lvl
        ppr = [el("w:tabs"), el("w:spacing", after="240" if sig != "li" else "80"),
               el("w:ind", left=str(izq), hanging="454")]
        ppr[0].append(el("w:tab", val="left", pos=str(izq)))
        p = self._p([], None, ppr)
        p.append(text_run(f"{marca}\t"))
        for x in self._runs_xml(b["runs"]):
            p.append(x)
        return p

    def _imagen(self, b: dict, sig: str | None):
        blob = b["blob"]
        try:
            ancho = b["cx"] or None
            if ancho and ancho > self.max_img:
                ancho = self.max_img
            if not ancho:
                ancho = min(self.max_img, 4_200_000)
            inline = self.tpl.part.new_pic_inline(io.BytesIO(blob), Emu(ancho), None)
        except Exception as e:  # EMF/WMF, formatos raros…
            self.avisos.append(f"Imagen no trasladada ({b['ctype']}): {e}. Reinsértala como PNG/JPG.")
            return [self._p([{"text": "[IMAGEN NO TRASLADADA — reinsertar como PNG/JPG]",
                              "b": True, "i": False, "u": False, "sup": False, "sub": False,
                              "href": None}])]
        self.stats["imagenes"] += 1
        if b.get("alt"):
            for dp in inline.iter(qn("wp:docPr")):
                dp.set("descr", b["alt"][:250])
        p = el("w:p")
        ppr = sub(p, "w:pPr")
        if sig == "pie":
            sub(ppr, "w:keepNext")
        sub(ppr, "w:jc", val="left")
        r = sub(p, "w:r")
        drawing = sub(r, "w:drawing")
        drawing.append(inline)
        return [p]

    # ------------------------------------------------------------------ tablas
    def _tabla(self, b: dict, sig: str | None):
        filas = b["rows"]
        if not filas or not any(filas):
            return []
        self.stats["tablas"] += 1
        ncols = max(sum(c["span"] for c in f) for f in filas)
        total = ANCHO_TABLA_STD if ncols <= 6 else ANCHO_UTIL
        # anchos proporcionales al texto más largo de cada columna (con topes)
        pesos = [6] * ncols
        for f in filas:
            col = 0
            for c in f:
                if c["span"] == 1 and col < ncols:
                    ln = max((len("".join(r.get("text", "") for r in p)) for p in c["paras"]), default=0)
                    pesos[col] = max(pesos[col], min(ln, 36))
                col += c["span"]
        sp = sum(pesos)
        anchos = [max(int(total * w / sp), 600) for w in pesos]
        anchos[-1] += total - sum(anchos)

        tbl = el("w:tbl")
        tblpr = sub(tbl, "w:tblPr")
        sub(tblpr, "w:tblStyle", val="Tabladelista2-nfasis1")
        sub(tblpr, "w:tblW", w=str(total), type="dxa")
        sub(tblpr, "w:tblLook", val="0620", firstRow="1", lastRow="0", firstColumn="0",
            lastColumn="0", noHBand="1", noVBand="1")
        grid = sub(tbl, "w:tblGrid")
        for w in anchos:
            sub(grid, "w:gridCol", w=str(w))

        n = len(filas)
        for ri, fila in enumerate(filas):
            tr = sub(tbl, "w:tr")
            trpr = sub(tr, "w:trPr")
            if ri == 0:
                sub(trpr, "w:cnfStyle", val="100000000000", firstRow="1", lastRow="0",
                    firstColumn="0", lastColumn="0", oddVBand="0", evenVBand="0", oddHBand="0",
                    evenHBand="0", firstRowFirstColumn="0", firstRowLastColumn="0",
                    lastRowFirstColumn="0", lastRowLastColumn="0")
            sub(trpr, "w:cantSplit")
            sub(trpr, "w:trHeight", val="379")
            if ri == 0:
                sub(trpr, "w:tblHeader")
            col = 0
            for c in fila:
                w = sum(anchos[col:col + c["span"]])
                col += c["span"]
                tc = sub(tr, "w:tc")
                tcpr = sub(tc, "w:tcPr")
                sub(tcpr, "w:tcW", w=str(w), type="dxa")
                if c["span"] > 1:
                    sub(tcpr, "w:gridSpan", val=str(c["span"]))
                if c["vmerge"]:
                    vm = sub(tcpr, "w:vMerge")
                    if c["vmerge"] == "restart":
                        vm.set(qn("w:val"), "restart")
                if ri == 0:
                    bd = sub(tcpr, "w:tcBorders")
                    sub(bd, "w:top", val="nil")
                    sub(bd, "w:bottom", val="single", sz="12", space="0", color="E2231A", themeColor="accent1")
                elif ri == 1:
                    bd = sub(tcpr, "w:tcBorders")
                    sub(bd, "w:top", val="single", sz="12", space="0", color="E2231A", themeColor="accent1")
                sub(tcpr, "w:vAlign", val="center")
                estilo = "Rango1Tablas" if ri == 0 else "DatosTabla"
                paras = c["paras"] or [[]]
                for pi, runs in enumerate(paras):
                    extra = [el("w:keepNext")] if (sig == "pie" and ri == n - 1) or ri == 0 else None
                    # el contenido de celda no debe heredar negrita: la plantilla lo decide
                    tc.append(self._p(runs, estilo, extra))
        out = [tbl]
        if sig == "tbl":  # dos tablas seguidas se fundirían en una
            out.append(self._p([], None, [el("w:spacing", after="0")]))
        return out

    # ------------------------------------------------------------------ índice
    def _indice(self):
        p = el("w:p")
        r = sub(p, "w:r")
        fc = sub(r, "w:fldChar", fldCharType="begin")
        fc.set(qn("w:dirty"), "true")
        r = sub(p, "w:r")
        it = sub(r, "w:instrText")
        it.text = ' TOC \\o "1-3" '
        it.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        sub(sub(p, "w:r"), "w:fldChar", fldCharType="separate")
        p.append(text_run("Pulsa F9 (o botón derecho → Actualizar campos) para generar el índice."))
        sub(sub(p, "w:r"), "w:fldChar", fldCharType="end")
        return p

    # ------------------------------------------------------------------ limpieza final
    def _limpiar_relaciones_huerfanas(self):
        xml = etree.tostring(self.body, encoding="unicode")
        usados = set(re.findall(r'(?:r:id|r:embed|r:link|r:pict)="(rId\d+)"', xml))
        for rid, rel in list(self.tpl.part.rels.items()):
            if rel.reltype in (RT.IMAGE, RT.CHART) and rid not in usados:
                self.tpl.part.rels.pop(rid)

    def _forzar_actualizar_campos(self):
        st = self.tpl.settings.element
        if st.find(qn("w:updateFields")) is not None:
            return
        nodo = el("w:updateFields", val="true")
        for tag in ("w:hdrShapeDefaults", "w:footnotePr", "w:endnotePr", "w:compat", "w:docVars",
                    "w:rsids", "m:mathPr", "w:themeFontLang", "w:clrSchemeMapping"):
            ref = st.find(qn(tag))
            if ref is not None:
                ref.addprevious(nodo)
                return
        st.append(nodo)

    # ------------------------------------------------------------------ API
    def construir(self, origen: Documento, cfg: dict) -> None:
        hijos, t_idx, s_idx = self._localizar()
        mapa = portada.construir_mapa(cfg)

        # 1) portada y contraportada (antes del índice) + cabeceras/pies
        for e in hijos[:t_idx]:
            portada.aplicar(e, mapa)
        for rel in self.tpl.part.rels.values():
            if rel.reltype in (RT.HEADER, RT.FOOTER):
                portada.aplicar(rel.target_part.element, mapa)

        # 2) vaciar índice de ejemplo y capítulos de ejemplo (conserva título «Índice»)
        for e in hijos[t_idx + 1:s_idx]:
            self.body.remove(e)
        destino = hijos[s_idx]
        destino.addprevious(self._indice())

        # 3) cuerpo
        self._reset_num()
        for i, b in enumerate(origen.bloques):
            for nodo in self._bloque(b, i, origen.bloques):
                destino.addprevious(nodo)

        # 4) notas al pie, relaciones, campos, metadatos
        self._volcar_notas()
        self._limpiar_relaciones_huerfanas()
        self._forzar_actualizar_campos()
        cp = self.tpl.core_properties
        if cfg.get("sector") and cfg.get("pais"):
            cp.title = f"El mercado de {cfg['sector']} en {cfg['pais']}"
        if cfg.get("autor"):
            cp.author = cfg["autor"]

        pend = []
        for e in hijos[:t_idx]:
            pend += portada.marcadores_pendientes(e)
        for rel in self.tpl.part.rels.values():
            if rel.reltype in (RT.HEADER, RT.FOOTER):
                pend += portada.marcadores_pendientes(rel.target_part.element)
        pend = sorted(set(pend))
        if pend:
            self.avisos.append("Marcadores sin rellenar en portada/cabeceras: " + ", ".join(pend)
                               + " → complétalos en portada.toml.")

    def guardar(self, ruta: str) -> None:
        self.tpl.save(ruta)
