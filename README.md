# icex-maquetador

Traslada un estudio de mercado redactado en **Word o Google Docs normal** a la **plantilla oficial de
estudios de mercado de ICEX (2024)**, sin romper el formato: portada, contraportada, cabeceras, pies,
estilos, numeración de capítulos e índice son los de la plantilla.

La idea: el equipo escribe (≈ 40-60 páginas) sin pensar en el diseño, y **el maquetado oficial se hace al
final con un solo comando**.

```
estudio.docx  ──┐
portada.toml  ──┼──►  icex-maquetador  ──►  estudio_ICEX.docx  (plantilla oficial rellena)
plantilla/*.docx┘
```

## Instalación

Requiere Python 3.11 o superior.

```bash
git clone https://github.com/<usuario>/icex-maquetador.git
cd icex-maquetador
python3 -m venv .venv && source .venv/bin/activate
pip install -e .
```

Coloca la plantilla oficial en `plantilla/` (ver [plantilla/LEEME.md](plantilla/LEEME.md)): **no se incluye en
el repositorio** porque es material de ICEX.

## Uso

1. Escribe el estudio siguiendo la [guía de redacción](docs/GUIA_REDACCION.md) (usar estilos de título,
   listas y tablas nativas, y los prefijos `Tabla:` / `Fuente:` para títulos y fuentes).
2. Copia `ejemplos/portada.toml`, rellénalo con los datos del estudio.
3. Convierte:

```bash
icex-maquetador mi_estudio.docx --portada portada.toml -o mi_estudio_ICEX.docx
```

4. Abre el resultado en Word y acepta **«actualizar campos»** para generar el índice. (En macOS con Word
   instalado puedes añadir `--actualizar-indice` y lo hace el propio programa.)

El programa imprime un resumen (capítulos, tablas, imágenes, notas) y **avisos** de todo lo que no ha podido
trasladar. Léelos antes de dar el documento por bueno.

### Opciones

| Opción | Descripción |
|---|---|
| `-p, --plantilla` | Ruta de la plantilla (por defecto `plantilla/PlantillaWord_ICEX_EstudioMercado_2024.docx`) |
| `-c, --portada` | Fichero TOML con datos de portada |
| `-o, --salida` | Fichero de salida (por defecto `<entrada>_ICEX.docx`) |
| `--max-img-cm` | Ancho máximo de las imágenes (15,5 cm por defecto) |
| `--actualizar-indice` | (macOS + Word) actualiza el índice y guarda |

Prueba rápida con el ejemplo incluido:

```bash
python ejemplos/generar_ejemplo.py
icex-maquetador ejemplos/estudio_ejemplo.docx -c ejemplos/portada.toml -o /tmp/ejemplo_ICEX.docx
```

## Cómo funciona

No genera el Word desde cero: **abre la plantilla oficial y sustituye solo su contenido de ejemplo**.

1. **Lee** tu `.docx` y lo reduce a bloques semánticos (títulos, párrafos, listas, tablas, imágenes, notas,
   títulos/pies de tabla). El formato directo (fuentes, colores, tamaños) se descarta a propósito.
2. **Rellena** los marcadores `[sector]`, `[País]`, `[Ciudad]`, `[año]`… de portada, contraportada, cabeceras
   y pies (aunque Word los haya partido en varios fragmentos internos).
3. **Vacía** el índice y los capítulos de ejemplo de la plantilla, conservando secciones, cabeceras y pies.
4. **Escribe** cada bloque con el estilo de la plantilla (`Título 1-4`, `Título Tablas`, `Datos Tabla`,
   `Pie de tabla y gráfico`, `Párrafo de lista 1-3`, `Destacado`…), y reconstruye las tablas con el diseño ICEX.
5. **Inserta** un campo de índice (`TOC`) que Word completa al abrir, y traslada las notas al pie.

Código en `src/icex_maquetador/` (`lector.py`, `escritor.py`, `portada.py`).

## Limitaciones conocidas

* Los **gráficos nativos** de Word/Excel/Docs no se trasladan: hay que pegarlos como imagen.
* Las imágenes EMF/WMF no se trasladan (aviso en pantalla); usa PNG o JPG.
* No se crean ni se reordenan capítulos: el orden es el de tu documento.
* Las **tablas** se reconstruyen con el diseño de la plantilla: anchos de columna proporcionales al contenido,
  no los del original. Se admiten celdas combinadas básicas.
* El número de página del índice solo es definitivo cuando Word lo actualiza.
* Pensado para la plantilla **2024**. Si ICEX publica otra versión con estilos distintos, habrá que ajustar
  los nombres de estilo en `escritor.py`.
* El diseño final se ha comprobado abriendo y actualizando el resultado en Word para Mac; conviene revisar
  visualmente cada estudio (saltos de página, tablas largas) antes de entregarlo.

## Desarrollo

```bash
pip install -e ".[test]"
pytest        # necesita la plantilla en plantilla/
```

Contribuciones bienvenidas: abre un *issue* con el `.docx` problemático (sin datos sensibles) y el aviso que
imprimió el programa.

## Licencia y marca

Código bajo licencia MIT (ver `LICENSE`). La plantilla, el logotipo y la marca ICEX pertenecen a ICEX España
Exportación e Inversiones y **no forman parte de este repositorio**. Este proyecto es una herramienta
independiente, no oficial.
