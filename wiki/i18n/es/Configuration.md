<!-- Translated from wiki/Configuration.md (git blob f4889ce51830df74593cc0803cf9ec6d5ccc4402). See wiki/TRANSLATION_GUIDE.md. -->
# Configurar docxi: márgenes, configuración de página y unidades

docxi no tiene archivo de configuración, variables de entorno ni ajustes globales. Todo se define por documento o por elemento. Esta página cubre los ajustes a nivel de documento y las unidades que usan las opciones. Las opciones de cada elemento están en [Funciones](Features).

## <a id="page-margins"></a>Márgenes de página

Los márgenes se pasan al crear el documento:

```ruby
require "docxi"

# Todos los valores están en veinteavos de punto (1440 = 1 pulgada).
word = Docxi::Document.new(top: 1440, right: 1080, bottom: 1440, left: 1080, header: 720, footer: 720)
word.document.p { |p| p.text "One-inch top and bottom margins" }

File.binwrite("margins.docx", word.render.string)
```

| Opción | Valor predeterminado | Significado |
|---|---|---|
| `top` | `720` (0.5 pulgadas) | Margen superior |
| `right` | `1440` (1 pulgada) | Margen derecho |
| `bottom` | `720` (0.5 pulgadas) | Margen inferior |
| `left` | `1440` (1 pulgada) | Margen izquierdo |
| `header` | `190` | Distancia del borde superior de la página al encabezado |
| `footer` | `200` | Distancia del borde inferior de la página al pie de página |

El hash que pasas sustituye a los valores predeterminados, así que define `top`, `right`, `bottom` y `left` juntos. Cualquiera de esos cuatro que omitas se escribe como un valor vacío. `header` y `footer` vuelven a 190 y 200 cuando faltan o están vacíos.

## <a id="page-size-and-orientation"></a>Tamaño y orientación de la página

Las páginas son US Letter en vertical: 12240 × 15840 veinteavos de punto, es decir, 8.5 × 11 pulgadas. El tamaño y la orientación no se pueden cambiar en la 0.1.0.

## <a id="first-page"></a>Primera página

docxi siempre activa la opción «Primera página diferente» de Word. Los encabezados y pies de página añadidos con `add_header` y `add_footer` se usan a partir de la página 2. La página 1 usa `add_first_header` y `add_first_footer`, así que defínelos también si la primera página necesita encabezado o pie de página. Consulta [Funciones](Features#headers-and-footers).

## <a id="what-every-document-contains"></a>Qué contiene cada documento

- La fuente predeterminada es Arial de 12 pt.
- Estilos definidos: `Normal`, `DefaultParagraphFont`, `TableNormal`, `TableCustom` y `NoList`.
- Los elementos de lista con viñetas (`ul: true`) usan un guion como viñeta.
- Las propiedades del documento indican «Docxi Ruby Gem» como nombre de la aplicación y registran la hora en que se generó el documento.

## <a id="units"></a>Unidades

| Opción | Unidad |
|---|---|
| `size:` del texto | Puntos (docxi lo duplica para obtener los medios puntos de Word) |
| Márgenes, `left:` y `right:` de párrafo, `iwidth:` de tabla, `height:` de fila | Veinteavos de punto (1440 por pulgada) |
| Tamaños de `borders:` de celda | Octavos de punto |
| `columns_width:` y `width:` de celda | Se multiplican por 14.92 para obtener veinteavos de punto, cerca de píxeles a 96 ppp |
| `style:` de imagen en línea | Longitudes CSS como `"120pt"` |
| `pH:` y `pV:` de imagen posicionada | EMU (914400 por pulgada) |
| Colores de texto y de relleno | Hexadecimal sin `#`, por ejemplo `"8B0000"` |
| `fill_color:` de imagen | Color VML con `#`, por ejemplo `"#00FF00"` |

## <a id="related-pages"></a>Páginas relacionadas

- [Funciones](Features)
- [Referencia de la API](API-Reference)
- [Buenas prácticas](Best-Practices)

---

[← Inicio de la wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
