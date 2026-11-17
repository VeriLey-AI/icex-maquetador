# Guía de redacción (para quien escribe el estudio)

Escribe el estudio en **Word o Google Docs normal**, sin preocuparte del diseño. Al final, el
conversor lo traslada a la plantilla oficial. Para que lo entienda bien, sigue estas reglas.

## 1. Estructura: usa los estilos de título, no formato manual

| En tu documento | Resultado en la plantilla ICEX |
|---|---|
| Título 1 (Heading 1) | Capítulo (empieza página nueva, numerado) |
| Título 2 | Epígrafe `1.1` |
| Título 3 | Subepígrafe `1.1.1` |
| Título 4 | Título de nivel 4 |

* **No numeres a mano** ("1. Resumen…"): la numeración la pone la plantilla. Si lo haces, se elimina sola.
* No uses negrita + tamaño grande en lugar de un título: no se reconocerá como título.
* Los capítulos de la plantilla son: Resumen ejecutivo, Definición del sector, Oferta – Análisis de
  competidores, Demanda, Precios, Percepción del producto español, Canales de distribución, Acceso al
  mercado – Barreras, Perspectivas del sector, Oportunidades, Información práctica y Otros
  capítulos/anexos. El conversor **no** los crea ni los reordena: es tu documento quien manda.
* No escribas portada ni índice: se generan solos (portada con `portada.toml`).

## 2. Texto y listas

* Párrafos normales: texto normal. Se conserva **negrita**, *cursiva*, subrayado, superíndices,
  subíndices y enlaces. Se descartan colores, fuentes, tamaños y sangrías (los decide la plantilla).
* Listas: usa las listas nativas de Word/Docs (viñetas o numeradas, hasta 3 niveles).
* Notas al pie: las nativas de Word/Docs se trasladan a la plantilla.
* Un párrafo que empiece por `Destacado:` se convierte en la cita destacada de la plantilla.

## 3. Tablas, gráficos y fotos

Cada objeto puede llevar **título**, **subtítulo** y **pie con la fuente**, escritos como párrafos
normales con estos prefijos:

```
Tabla 1: Importaciones por país de origen      ← ANTES del objeto  → "Título Tablas"
Subtítulo: Millones de euros                   ← opcional          → "Subtítulo para tablas y gráficos"
[ la tabla o la imagen ]
Fuente: Eurostat, 2024.                        ← DESPUÉS del objeto → "Pie de tabla y gráfico"
```

* Prefijos válidos para el título: `Tabla`, `Gráfico`, `Figura`, `Foto`, `Mapa`, `Cuadro` (seguidos de `:`,
  `.` o `-`, con número opcional: `Gráfico 3. Evolución…`). El prefijo y el número se eliminan.
* Prefijos válidos para el pie: `Fuente:`, `Fuentes:`, `Elaboración propia`.
* Si prefieres, también vale el estilo **Leyenda/Caption** de Word: se coloca como título (antes del
  objeto) o como pie (después), según la posición.
* **Tablas:** usa tablas nativas; la primera fila es la cabecera. Se admiten celdas combinadas.
  El diseño (líneas rojas, tipografías, anchos) lo pone la plantilla.
* **Imágenes:** insértalas en línea, en su propio párrafo, mejor en PNG o JPG. Se ajustan al ancho de página.
  Los formatos EMF/WMF no se trasladan (el conversor te avisa).
* **Gráficos:** los gráficos nativos de Word/Excel/Docs **no se pueden trasladar**. Pégalos como imagen
  (PNG) o expórtalos desde Excel/Sheets como imagen.

## 4. Qué no se traslada (y el conversor te avisa)

Gráficos nativos, imágenes en celdas de tabla, tablas anidadas, cuadros de texto y formas, comentarios y
cambios con control de cambios sin aceptar (se aceptan las inserciones y se descartan las eliminaciones).

## 5. Flujo recomendado

1. Redacta el estudio completo (≈ 40-60 páginas) en Word o Google Docs. Si usas Google Docs:
   *Archivo → Descargar → Microsoft Word (.docx)*.
2. Rellena `portada.toml` (sector, país, ciudad, autor, fecha…).
3. Ejecuta el conversor (ver `README.md`).
4. Abre el resultado en Word, acepta «actualizar campos» y revisa el índice y los saltos de página.
5. Revisa los avisos que imprimió el conversor.
