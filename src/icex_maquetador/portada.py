"""Rellena los marcadores [entre corchetes] de portada, contraportada, cabeceras y pies."""
from __future__ import annotations

import datetime as _dt
import re
import tomllib
from pathlib import Path

from docx.oxml.ns import qn

from .xml import XML_SPACE

MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
         "agosto", "septiembre", "octubre", "noviembre", "diciembre"]

CAMPOS = ("sector", "pais", "ciudad", "autor")


def cargar_config(ruta: str | Path | None) -> dict:
    if ruta is None:
        return {}
    with open(ruta, "rb") as f:
        return tomllib.load(f)


def construir_mapa(cfg: dict) -> dict[str, str]:
    """Traduce la configuración (TOML) a pares ``marcador -> texto``."""
    m: dict[str, str] = {}
    sector, pais = cfg.get("sector"), cfg.get("pais")
    ciudad, autor = cfg.get("ciudad"), cfg.get("autor")

    fecha = cfg.get("fecha")
    if fecha:
        d = fecha if isinstance(fecha, _dt.date) else _dt.date.fromisoformat(str(fecha))
        m["[día]"] = str(d.day)
        m["[mes]"] = MESES[d.month - 1]
        m["[año]"] = str(d.year)
    if cfg.get("anio") or cfg.get("año"):
        m["[año]"] = str(cfg.get("anio") or cfg.get("año"))

    if sector:
        m["[sector]"] = sector
        m["[SECTOR]"] = sector.upper()
    if pais:
        m["[País]"] = pais
        m["[PAÍS]"] = pais.upper()
    if ciudad:
        m["[Ciudad]"] = ciudad
    if autor:
        m["[Nombre y Apellidos]"] = autor
    if "enlace" in cfg:
        m["[INSERTAR ENLACE]"] = cfg["enlace"]
    if cfg.get("web_pais"):  # http://<web_pais>.oficinascomerciales.es
        m["país."] = f"{cfg['web_pais']}."
    if cfg.get("nipo"):
        m["224240080"] = str(cfg["nipo"])
    return m


def _runs_con_texto(p):
    """``w:t`` que cuelgan directamente de este párrafo (sin entrar en cuadros de texto)."""
    ts = []
    for child in p:
        if child.tag == qn("w:r"):
            ts.extend(child.findall(qn("w:t")))
        elif child.tag == qn("w:hyperlink"):
            for r in child.findall(qn("w:r")):
                ts.extend(r.findall(qn("w:t")))
    return ts


def reemplazar_en_parrafo(p, mapa: dict[str, str]) -> int:
    """Sustituye marcadores aunque Word los haya partido en varios *runs*."""
    n = 0
    for viejo, nuevo in mapa.items():
        while True:
            ts = _runs_con_texto(p)
            textos = [t.text or "" for t in ts]
            full = "".join(textos)
            idx = full.find(viejo)
            if idx < 0:
                break
            fin = idx + len(viejo)
            pos = 0
            puesto = False
            for t, txt in zip(ts, textos):
                a, b = pos, pos + len(txt)
                pos = b
                if b <= idx or a >= fin:
                    continue
                izq = txt[: max(idx - a, 0)]
                der = txt[max(fin - a, 0):] if fin < b else ""
                t.text = izq + (nuevo if not puesto else "") + der
                t.set(XML_SPACE, "preserve")
                puesto = True
            n += 1
    return n


def aplicar(raiz, mapa: dict[str, str]) -> int:
    """Aplica el mapa a todos los párrafos bajo ``raiz`` (incluye cuadros de texto)."""
    total = 0
    for p in raiz.iter(qn("w:p")):
        total += reemplazar_en_parrafo(p, mapa)
    return total


def marcadores_pendientes(raiz) -> list[str]:
    """Marcadores ``[...]`` que siguen sin rellenar (para avisar al usuario)."""
    vistos: list[str] = []
    for p in raiz.iter(qn("w:p")):
        txt = "".join(t.text or "" for t in _runs_con_texto(p))
        for m in re.findall(r"\[[^\]]{1,40}\]", txt):
            if m not in vistos:
                vistos.append(m)
    return vistos
