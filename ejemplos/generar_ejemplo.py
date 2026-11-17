"""Genera ejemplos/estudio_ejemplo.docx: un documento "normal" redactado con las convenciones de
docs/GUIA_REDACCION.md. Sirve para probar el conversor sin depender de documentos reales."""
import struct
import zlib
from pathlib import Path

import docx
from docx.shared import Cm

AQUI = Path(__file__).parent


def png_degradado(ruta: Path, w=800, h=480):
    filas = b"".join(b"\x00" + bytes(v for x in range(w) for v in (30 + x * 150 // w, 90 + y * 100 // h, 160))
                     for y in range(h))
    def chunk(t, d):
        c = struct.pack(">I", len(d)) + t + d
        return c + struct.pack(">I", zlib.crc32(t + d) & 0xFFFFFFFF)
    ruta.write_bytes(b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0))
                     + chunk(b"IDAT", zlib.compress(filas)) + chunk(b"IEND", b""))


def _enlace(p, texto, url):
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.opc.constants import RELATIONSHIP_TYPE as RT
    rid = p.part.relate_to(url, RT.HYPERLINK, is_external=True)
    h = OxmlElement("w:hyperlink")
    h.set(qn("r:id"), rid)
    r = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = texto
    r.append(t)
    h.append(r)
    p._p.append(h)


def _notas_al_pie(ruta: Path, notas: dict[str, str]):
    """python-docx no crea notas al pie; las inyectamos a mano para que el ejemplo las incluya."""
    import shutil
    import zipfile
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    tmp = ruta.with_suffix(".tmp")
    with zipfile.ZipFile(ruta) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml":
                x = data.decode("utf8")
                for i, key in enumerate(notas, 1):
                    x = x.replace(f"<w:t>[[{key}]]</w:t>",
                                  f'<w:footnoteReference w:id="{i}"/>')
                data = x.encode("utf8")
            elif item.filename == "[Content_Types].xml":
                data = data.decode("utf8").replace("</Types>",
                    '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>').encode("utf8")
            elif item.filename == "word/_rels/document.xml.rels":
                data = data.decode("utf8").replace("</Relationships>",
                    '<Relationship Id="rIdFn1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>').encode("utf8")
            zout.writestr(item, data)
        fn = "".join(f'<w:footnote w:id="{i}"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> {t}</w:t></w:r></w:p></w:footnote>'
                     for i, t in enumerate(notas.values(), 1))
        zout.writestr("word/footnotes.xml", f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                      f'<w:footnotes xmlns:w="{W}"><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>'
                      f'<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>{fn}</w:footnotes>')
    shutil.move(tmp, ruta)


def main():
    png = AQUI / "_grafico.png"
    png_degradado(png)
    d = docx.Document()
    lorem = ("Lorem ipsum dolor sit amet, consectetur adipiscing elit. Integer pulvinar congue orci. "
             "Nulla facilisi. Ut ullamcorper luctus mauris in posuere. Donec posuere mattis tincidunt. ") * 3

    d.add_heading("Resumen ejecutivo", 1)
    d.add_heading("Visión general", 2)
    d.add_paragraph("Destacado: El mercado crecerá un 6 % anual hasta 2030.")
    p = d.add_paragraph(lorem)
    p.add_run(" Texto en negrita").bold = True
    p.add_run(" y en cursiva.").italic = True
    for t in ("Primer punto clave", "Segundo punto clave", "Tercer punto clave"):
        d.add_paragraph(t, style="List Bullet")
    d.add_paragraph("Subpunto del tercero", style="List Bullet 2")

    d.add_heading("2. Definición del sector", 1)   # la numeración manual se elimina sola
    d.add_paragraph(lorem)
    d.add_heading("Clasificación arancelaria", 2)
    d.add_paragraph(lorem)
    d.add_paragraph("Tabla 1: Importaciones por país de origen")
    d.add_paragraph("Subtítulo: Millones de euros")
    t = d.add_table(rows=4, cols=4)
    t.style = "Table Grid"
    datos = [("País", "2021", "2022", "2023"), ("Italia", "120", "135", "150"),
             ("Alemania", "98", "101", "110"), ("España", "75", "88", "97")]
    for r, fila in enumerate(datos):
        for c, v in enumerate(fila):
            t.cell(r, c).text = v
    d.add_paragraph("Fuente: Eurostat, 2024.")

    d.add_heading("3. Oferta – Análisis de competidores", 1)
    d.add_heading("Principales actores", 2)
    d.add_paragraph(lorem)
    d.add_heading("Cuota de mercado", 3)
    d.add_paragraph(lorem)
    d.add_paragraph("Pasos recomendados para el exportador:")
    for t_ in ("Registrar la marca", "Elegir distribuidor local", "Adaptar el etiquetado"):
        d.add_paragraph(t_, style="List Number")
    d.add_paragraph("Gráfico 1: Evolución de la cuota de mercado")
    d.add_picture(str(png), width=Cm(12))
    d.add_paragraph("Fuente: elaboración propia.")

    d.add_heading("4. Demanda", 1)
    d.add_paragraph(lorem)
    p = d.add_paragraph("Más información en la web de ")
    _enlace(p, "ICEX", "https://www.icex.es")
    p.add_run(". Dato con nota al pie.")
    p.add_run("[[FN1]]")          # marcador: se convierte en nota al pie real (ver _notas_al_pie)
    d.add_paragraph(lorem)

    out = AQUI / "estudio_ejemplo.docx"
    d.save(out)
    _notas_al_pie(out, {"FN1": "Fuente: Ministerio de Economía de Alemania, 2024."})
    png.unlink()
    print(out)


if __name__ == "__main__":
    main()
