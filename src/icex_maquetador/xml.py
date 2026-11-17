"""Utilidades pequeñas para construir y manipular XML de WordprocessingML."""
from __future__ import annotations

import re
from typing import Iterable

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def el(tag: str, **attrs: str):
    """Crea un elemento ``w:*`` con atributos ``w:*`` (p. ej. ``el('w:pStyle', val='Ttulo1')``)."""
    e = OxmlElement(tag)
    for k, v in attrs.items():
        e.set(qn(f"w:{k}"), str(v))
    return e


def sub(parent, tag: str, **attrs: str):
    child = el(tag, **attrs)
    parent.append(child)
    return child


def text_run(text: str, rpr_children: Iterable = ()):
    """Run con texto; los ``\\t`` y ``\\n`` se convierten en ``w:tab`` / ``w:br``."""
    r = el("w:r")
    children = list(rpr_children)
    if children:
        rpr = sub(r, "w:rPr")
        for c in children:
            rpr.append(c)
    parts = re.split(r"(\t|\n)", text)
    for part in parts:
        if part == "\t":
            sub(r, "w:tab")
        elif part == "\n":
            sub(r, "w:br")
        elif part:
            t = sub(r, "w:t")
            t.text = part
            t.set(XML_SPACE, "preserve")
    return r


def local(tag: str) -> str:
    return tag.split("}")[-1]


def paragraph_style(p) -> str | None:
    ppr = p.find(qn("w:pPr"))
    if ppr is None:
        return None
    s = ppr.find(qn("w:pStyle"))
    return None if s is None else s.get(qn("w:val"))


def has_sectpr(p) -> bool:
    ppr = p.find(qn("w:pPr"))
    return ppr is not None and ppr.find(qn("w:sectPr")) is not None
