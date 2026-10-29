<!-- Translated from wiki/Features.md (git blob 3eb1a7d53ab3801f4d8e200cedf6c2a9e8fef2be). See wiki/TRANSLATION_GUIDE.md. -->
# Funciones de docxi: párrafos, texto, tablas, imágenes, encabezados y pies de página

Esta página enumera todos los elementos que docxi puede añadir a un documento de Word y las opciones que acepta cada uno. Los ejemplos parten de esta configuración:

```ruby
require "docxi"

word = Docxi::Document.new
document = word.document
```

Las opciones que no necesites se pueden omitir. Los colores son valores hexadecimales sin `#`, por ejemplo `"8B0000"`. Para las unidades y la configuración de página, consulta [Configuración](Configuration).

## <a id="paragraphs"></a>Párrafos

`document.p(options) { |p| ... }` añade un párrafo.

| Opción | Valores | Efecto |
|---|---|---|
| `align` | `"left"`, `"center"`, `"right"`, `"both"` | Alineación. `"both"` justifica el texto |
| `fill` | Color hexadecimal | Sombreado detrás del párrafo y su texto |
| `left`, `right` | Veinteavos de punto | Sangría. Define ambos, porque uno solo se ignora |
| `ul` | `true` | Convierte el párrafo en un elemento de lista con viñetas |
| `bottom` | `true` | Dibuja un borde bajo el párrafo |

Las opciones de texto como `size:`, `bold:` o `color:` también se pueden pasar a `p`. Se convierten en los valores predeterminados de cada fragmento de texto del párrafo.

Dentro del bloque puedes llamar a `text`, `br`, `tab`, `image`, `hyperlink` y `frame`. `document.text(string, options)` es un atajo para un párrafo con un único fragmento.

```ruby
document.p(align: "both", left: 720, right: 720) do |p|
  p.text "An indented, justified paragraph."
end

document.p(ul: true) { |p| p.text "A bullet item" }
document.text "A one-line paragraph", size: 14, bold: true
```

## <a id="text"></a>Texto

`p.text(string, options)` añade un fragmento de texto (run).

| Opción | Valores | Efecto |
|---|---|---|
| `bold` | `true` | Negrita |
| `italic` | `true` | Cursiva |
| `underline` | `true` | Subrayado simple |
| `color` | Color hexadecimal | Color del texto |
| `size` | Puntos | Tamaño de fuente |
| `font` | Nombre de fuente | Fuente, por ejemplo `"Georgia"` |
| `fill` | Color hexadecimal | Sombreado detrás del texto |
| `space` | `true` | Conserva los espacios al principio y al final del texto |
| `br` | `true` | Añade un salto de línea después del texto |

Una cadena `nil` se omite.

```ruby
document.p do |p|
  p.text "Total: ", bold: true, space: true
  p.text "$1,200", color: "007A33", font: "Georgia", size: 14
end
```

## <a id="line-breaks-page-breaks-and-tabs"></a>Saltos de línea, saltos de página y tabulaciones

- `p.br`, o `p.br(times: 3)`, añade saltos de línea dentro de un párrafo.
- `document.br(page: true)` empieza una página nueva.
- `p.tab`, o `p.tab(times: 2)`, añade caracteres de tabulación dentro de un párrafo.
- `document.tab(page: true)` añade un párrafo que contiene una sola tabulación.

No llames a `document.br` ni a `document.tab` sin `page: true`. Colocan un fragmento de texto directamente en el cuerpo del documento, lo que no es WordprocessingML válido.

```ruby
document.p do |p|
  p.text "Name:", br: true
  p.tab
  p.text "Signature line"
  p.br(times: 2)
end

document.br(page: true)
```

## <a id="tables"></a>Tablas

`document.table(options)` recibe un bloque. Dentro de él, `table.tr(options)` añade filas y `tr.tc(options)` añade celdas.

**Opciones de tabla**

| Opción | Valores | Efecto |
|---|---|---|
| `columns_width` | Array de números | Anchos de columna, multiplicados por 14.92 para obtener veinteavos de punto. También son los anchos predeterminados de las celdas |
| `height` | Veinteavos de punto | Altura mínima de cada fila |
| `iwidth` | Veinteavos de punto | Sangría de la tabla desde el margen izquierdo |
| `style` | Identificador de estilo | Estilo de tabla. docxi define `TableNormal` y `TableCustom` |
| `width` | Número | Ancho preferido de la tabla, escrito con el tipo `auto`. Sin `columns_width`, cada celda recibe también este ancho, multiplicado por 14.92 |
| `align` | `"left"`, `"center"`, `"right"` | Alineación de la tabla. El texto añadido con `tc.text` se alinea igual |
| `padding` | `true` | Añade un relleno fijo dentro de las celdas |
| `borders` | Hash | Problema conocido: se aplica a cada celda en lugar de al contorno de la tabla |

**Opciones de fila**: `height:` define la altura mínima de la fila. Cualquier opción de celda que se pase a `tr` se aplica a todas las celdas de la fila. Las opciones de la tabla se transmiten de la misma forma a cada fila y celda.

**Opciones de celda**

| Opción | Valores | Efecto |
|---|---|---|
| `borders` | `{ top:, right:, bottom:, left: }` | Grosor del borde en octavos de punto. `nil` significa sin borde en ese lado |
| `fill` | Color hexadecimal | Sombreado de la celda. Las celdas sin relleno son blancas (`"FFFFFF"`) |
| `valign` | `"top"`, `"center"`, `"bottom"` | Alineación vertical |
| `colspan` | Entero | Número de columnas que abarca la celda |
| `rowspan` | `true` | Inicia una combinación vertical |
| `merged` | `true` | Continúa la combinación vertical de la celda de arriba. Su contenido se ignora |
| `width` | Número | Ancho de la celda, multiplicado por 14.92 |

**Contenido de la celda**

- `tc.text(string, options)`: las opciones de celda y de fila, como `bold:` o `align:`, también se aplican al texto.
- `tc.p(options) { |p| ... }`: un párrafo solo con sus propias opciones.
- `tc.image(media, width:, height:)`: una imagen. Consulta el problema de tamaño en [Imágenes](#images).
- `tc.page_numbers`: un campo de número de página. Problema conocido: lleva delante el texto «GlobalOptions».

```ruby
document.table(columns_width: [150, 150]) do |table|
  table.tr do |tr|
    tr.tc(rowspan: true, valign: "center", borders: { top: 4, right: 4, bottom: 4, left: 4 }) do |tc|
      tc.text "Spans two rows"
    end
    tr.tc { |tc| tc.text "Row 1" }
  end
  table.tr do |tr|
    tr.tc(merged: true)
    tr.tc { |tc| tc.text "Row 2" }
  end
end
```

## <a id="images"></a>Imágenes

`media = document.add_media(source)` registra una imagen. `source` puede ser:

- un objeto IO, como un `File` o un `StringIO`
- una ruta de archivo
- una URL `http`, `https` o `ftp`

La imagen se lee cuando llamas a `render` y siempre se guarda como una parte PNG, así que los archivos PNG son la opción segura.

**Imágenes en línea**: `p.image(media, options)` coloca la imagen en la línea de texto.

| Opción | Valores | Efecto |
|---|---|---|
| `style` | Hash de propiedades CSS | Tamaño, por ejemplo `{ width: "120pt", height: "60pt" }` |
| `title` | Cadena | Título de la imagen |
| `fill_color` | Color VML como `"#00FF00"` | Relleno detrás de la imagen |
| `opacity` | Cadena como `"50%"` | Opacidad del relleno |
| `wrap` | `true` | Desactiva el ajuste del texto alrededor de la imagen |

```ruby
logo = document.add_media(File.open("logo.png", "rb"))
document.p(align: "center") { |p| p.image logo, style: { width: "120pt", height: "60pt" } }
```

**Imágenes posicionadas**: `p.image(media, pH:, pV:, wrap: true, width:, height:)` ancla la imagen con un desplazamiento respecto a la columna (`pH`) y al párrafo (`pV`), en EMU. Hay que indicar `pH`, `pV` y `wrap`; si falta alguno, la imagen se coloca en línea.

**Imágenes de encabezado**: `h.image(media, align:, width:, height:, pH:, pV:, behind:, wrap:)`. `wrap: "none"` desactiva el ajuste. Cualquier otro valor ajusta el texto alrededor de la imagen en forma cuadrada.

**Problema conocido en la 0.1.0**: las imágenes posicionadas, las de encabezado y las de celdas de tabla calculan tanto su ancho como su alto como `width * height * 14.92` EMU. Siempre salen cuadradas y normalmente diminutas. Las imágenes en línea con tamaño definido mediante `style:` no se ven afectadas. Consulta [Buenas prácticas](Best-Practices#images) para una solución alternativa.

## <a id="headers-and-footers"></a>Encabezados y pies de página

```ruby
header = Docxi::Word::Headers::Header.new do |h|
  h.text "Quarterly report", size: 9, align: "right"
end

footer = Docxi::Word::Footers::Footer.new do |f|
  f.page_numbers name: "Page", align: "center"
end

document.add_header(header)
document.add_footer(footer)
document.add_first_header(Docxi::Word::Headers::Header.new { |h| h.text "Cover page" })
```

- Un encabezado puede contener `text`, `br`, `table`, `image` y `page_numbers`. Un pie de página puede contener `text`, `br`, `table` y `page_numbers`.
- Las opciones que se pasan a `Header.new` o `Footer.new` se aplican a su texto.
- `add_header` y `add_footer` cubren de la página 2 en adelante. La página 1 siempre usa `add_first_header` y `add_first_footer`. Consulta [Configuración](Configuration#first-page).

## <a id="page-numbers"></a>Números de página

`page_numbers(name: "Page", align: "center")` escribe el texto de `name` seguido de un campo PAGE. El número se muestra en Arial de 10 pt, gris oscuro. `align` es `"right"` de forma predeterminada.

Problema conocido: al código de campo le falta la barra invertida de `\* MERGEFORMAT`.

## <a id="table-of-contents"></a>Tabla de contenido

```ruby
document.table_of_content do |toc|
  toc.add_row "Introduction", 1
  toc.add_row "Appendix", 7, leader: "hyphen", pos: 8000
  toc.add_row "Notes"
end
```

- `add_row(text, page = nil, leader: "dot", pos: 9360)`.
- Con un número de página, la entrada recibe una tabulación alineada a la derecha en `pos` (veinteavos de punto; 9360 son 6.5 pulgadas) rellena con el carácter de `leader`. Los valores de relleno proceden de WordprocessingML, como `"dot"`, `"hyphen"`, `"underscore"`, `"middleDot"` y `"none"`.
- Sin número de página, la entrada es el texto tras una tabulación.
- Los números de página son los que pasas. docxi no crea un campo TOC de Word ni calcula páginas.

## <a id="hyperlinks"></a>Hipervínculos

```ruby
url = "https://github.com/hammad417/docxi"
link = document.add_hyperlink(url)
document.p { |p| p.hyperlink "docxi on GitHub", url, link }
```

`add_hyperlink(url)` devuelve un identificador de relación, y `hyperlink(text, url, id)` coloca el vínculo. El destino sale de la relación, así que pasa el identificador que se devolvió para la URL que quieras. Solo se admiten vínculos externos.

## <a id="frames-text-boxes"></a>Marcos (cuadros de texto)

```ruby
document.p do |p|
  p.frame(style: "width:200pt;height:40pt") do |f|
    f.text "Text inside a box"
  end
end

document.frame "A box with one line", style: "width:200pt;height:30pt"
```

- Los marcos son cuadros de texto VML con un relleno fijo gris claro semitransparente y un contorno negro. El texto se coloca por encima y por debajo de ellos.
- Dentro de un marco puedes llamar a `text`, `br`, `tab`, `image` y `frame`.
- `spacing: true` añade espacio antes del texto.

## <a id="output"></a>Salida

`word.render` devuelve un `StringIO` con el paquete `.docx` completo. El archivo ZIP se escribe sin registros Zip64, así que Google Docs y las versiones antiguas de LibreOffice pueden abrirlo. [Arquitectura](Architecture) enumera cada parte del paquete.

## <a id="related-pages"></a>Páginas relacionadas

- [Ejemplos de uso](Usage-Examples)
- [Referencia de la API](API-Reference)
- [Hoja de ruta](Roadmap), con los problemas conocidos previstos para la 0.2.0

---

[← Inicio de la wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
