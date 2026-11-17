"""Prueba de humo: convierte el ejemplo y comprueba la estructura del resultado.
Necesita la plantilla oficial en plantilla/ (no se distribuye con el repositorio)."""
from pathlib import Path

import docx
import pytest
from docx.oxml.ns import qn

RAIZ = Path(__file__).resolve().parents[1]
PLANTILLA = RAIZ / "plantilla" / "PlantillaWord_ICEX_EstudioMercado_2024.docx"

pytestmark = pytest.mark.skipif(not PLANTILLA.exists(), reason="falta la plantilla oficial en plantilla/")


def _convertir(tmp_path):
    from icex_maquetador.cli import main
    salida = tmp_path / "salida.docx"
    rc = main([str(RAIZ / "ejemplos" / "estudio_ejemplo.docx"), "-p", str(PLANTILLA),
               "-c", str(RAIZ / "ejemplos" / "portada.toml"), "-o", str(salida)])
    assert rc == 0
    return docx.Document(str(salida))


def test_estilos_y_estructura(tmp_path):
    d = _convertir(tmp_path)
    estilos = [p.style.style_id for p in d.paragraphs]
    assert estilos.count("Ttulo1") == 4          # cuatro capítulos
    assert "TtuloTablas" in estilos and "Piedetablaygrfico" in estilos
    assert "Destacado" in estilos and "Prrafodelista" in estilos
    assert len(d.tables) == 1 + 0                # solo la tabla del ejemplo (la plantilla se vació)


def test_portada_sin_marcadores(tmp_path):
    d = _convertir(tmp_path)
    texto = "".join(t.text or "" for t in d.element.body.iter(qn("w:t")))
    assert "[sector]" not in texto and "[País]" not in texto and "RECORDAMOS" not in texto
    assert "17 de noviembre de 2026" in texto


def test_sin_graficos_huerfanos(tmp_path):
    import zipfile
    _convertir(tmp_path)
    with zipfile.ZipFile(tmp_path / "salida.docx") as z:
        assert not [n for n in z.namelist() if n.startswith("word/charts/")]
