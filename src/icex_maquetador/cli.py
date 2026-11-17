from __future__ import annotations

import argparse
import sys
from pathlib import Path

from . import __version__
from .escritor import Escritor
from .lector import Lector
from .portada import cargar_config


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(
        prog="icex-maquetador",
        description="Traslada un .docx normal a la plantilla oficial de estudios de mercado de ICEX.")
    ap.add_argument("entrada", help="documento .docx redactado en Word o exportado de Google Docs")
    ap.add_argument("-p", "--plantilla", default="plantilla/PlantillaWord_ICEX_EstudioMercado_2024.docx",
                    help="plantilla oficial .docx (por defecto, plantilla/…2024.docx)")
    ap.add_argument("-c", "--portada", help="fichero TOML con los datos de portada (ver ejemplos/portada.toml)")
    ap.add_argument("-o", "--salida", help="documento de salida (por defecto <entrada>_ICEX.docx)")
    ap.add_argument("--max-img-cm", type=float, default=15.5, help="ancho máximo de imágenes en cm")
    ap.add_argument("--titulos-numerados", action="store_true",
                    help="detecta como títulos los párrafos normales que empiezan por '1.', '2.1.'… "
                         "(para borradores sin estilos de título)")
    ap.add_argument("--actualizar-indice", action="store_true",
                    help="(macOS + Word) abre el resultado en Word, actualiza el índice y guarda")
    ap.add_argument("--version", action="version", version=__version__)
    a = ap.parse_args(argv)

    entrada = Path(a.entrada)
    plantilla = Path(a.plantilla)
    if not entrada.exists():
        ap.error(f"no existe {entrada}")
    if not plantilla.exists():
        ap.error(f"no encuentro la plantilla en {plantilla}. Descárgala de ICEX y colócala ahí "
                 f"(ver plantilla/LEEME.md) o indica otra ruta con --plantilla.")
    salida = Path(a.salida) if a.salida else entrada.with_name(entrada.stem + "_ICEX.docx")

    lector = Lector(str(entrada), titulos_numerados=a.titulos_numerados)
    doc = lector.leer()
    cfg = cargar_config(a.portada)
    esc = Escritor(str(plantilla), max_img_cm=a.max_img_cm)
    esc.construir(doc, cfg)
    esc.guardar(str(salida))

    s = esc.stats
    print(f"✔ {salida}")
    print(f"  {s['capitulos']} capítulos · {s['tablas']} tablas · {s['imagenes']} imágenes · "
          f"{s['notas']} notas al pie · ~{s['palabras']} palabras")
    avisos = doc.avisos + esc.avisos
    if avisos:
        print("\nAvisos:", file=sys.stderr)
        for av in avisos:
            print(f"  ⚠ {av}", file=sys.stderr)

    if a.actualizar_indice:
        from .word_mac import actualizar_con_word
        actualizar_con_word(salida)
        print("✔ Índice actualizado con Word")
    else:
        print("\nAl abrir el .docx en Word, acepta «actualizar campos» para generar el índice.")
    return 0
