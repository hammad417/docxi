<!-- Translated from wiki/API-Reference.md (git blob a682a57d84aa1c2b92c26f109cace3bc47079342). See wiki/TRANSLATION_GUIDE.md. -->
# Referencia de la API de docxi: clases y métodos

docxi es una biblioteca de Ruby, así que esta página enumera su API de Ruby. No hay API HTTP. Las opciones de cada método se describen en [Funciones](Features), y las clases se explican en [Arquitectura](Architecture).

## <a id="docxidocument"></a>Docxi::Document

| Método | Devuelve | Notas |
|---|---|---|
| `Docxi::Document.new(options = {})` | Document | `options` son los márgenes de página. Consulta [Configuración](Configuration) |
| `#document` | `Docxi::Word::Document` | El cuerpo al que añades contenido |
| `#render` | `StringIO` | El paquete `.docx` completo. `.string` devuelve los bytes |

## <a id="docxiworddocument-worddocument"></a>Docxi::Word::Document (`word.document`)

| Método | Devuelve | Añade |
|---|---|---|
| `p(options = {}) { \|p\| }` | Paragraph | Un párrafo |
| `text(string, options = {})` | Paragraph | Un párrafo con un solo fragmento de texto |
| `br(page: true)` | Break | Un salto de página |
| `tab(page: true)` | Tab | Un párrafo con una sola tabulación |
| `table(options = {}) { \|table\| }` | Table | Una tabla |
| `table_of_content(options = {}) { \|toc\| }` | TableOfContent | Entradas de la tabla de contenido |
| `frame(text, options = {})` | Paragraph | Un párrafo con un cuadro de texto |
| `add_media(io_path_or_url, options = {})` | Media | Una imagen para colocar con `image` |
| `add_hyperlink(url, options = {})` | String | Un identificador de relación para `hyperlink` |
| `add_header(header)`, `add_first_header(header)` | Header | Encabezado de la página 2 en adelante, y de la página 1 |
| `add_footer(footer)`, `add_first_footer(footer)` | Footer | Pie de página de la página 2 en adelante, y de la página 1 |

## <a id="paragraph-docxiwordcontentsparagraph"></a>Párrafo (`Docxi::Word::Contents::Paragraph`)

| Método | Añade |
|---|---|
| `text(string, options = {})` | Un fragmento de texto. Las opciones del párrafo son sus valores predeterminados |
| `br(times: 1)` | Saltos de línea |
| `tab(times: 1)` | Caracteres de tabulación |
| `image(media, options = {})` | Una imagen en línea o posicionada |
| `hyperlink(text, url, id)` | Un vínculo que usa el `id` de relación devuelto por `add_hyperlink` |
| `frame(options = {}) { \|frame\| }` | Un cuadro de texto |

## <a id="tables"></a>Tablas

| Método | Añade |
|---|---|
| `Table#tr(options = {}) { \|tr\| }` | Una fila. Las opciones de la tabla son sus valores predeterminados |
| `TableRow#tc(options = {}) { \|tc\| }` | Una celda. Las opciones de la fila son sus valores predeterminados |
| `TableCell#text(string, options = {})` | Un párrafo. Las opciones de la celda son sus valores predeterminados |
| `TableCell#p(options = {}) { \|p\| }` | Un párrafo con sus propias opciones |
| `TableCell#image(media, width:, height:)` | Una imagen |
| `TableCell#page_numbers` | Un campo de número de página |

## <a id="table-of-contents"></a>Tabla de contenido

| Método | Añade |
|---|---|
| `TableOfContent#add_row(text, page = nil, leader: "dot", pos: 9360)` | Una entrada de la tabla de contenido |

## <a id="headers-and-footers"></a>Encabezados y pies de página

| Clase | Métodos |
|---|---|
| `Docxi::Word::Headers::Header.new(options = {}) { \|h\| }` | `text`, `br`, `table`, `image`, `page_numbers(options = {})` |
| `Docxi::Word::Footers::Footer.new(options = {}) { \|f\| }` | `text`, `br`, `table`, `page_numbers(options = {})` |

## <a id="frames"></a>Marcos

`Docxi::Word::Contents::Frame` acepta `text`, `br`, `tab`, `image` y `frame` dentro de su bloque.

## <a id="constants-and-helpers"></a>Constantes y métodos auxiliares

| Nombre | Significado |
|---|---|
| `Docxi::VERSION` | La versión instalada, por ejemplo `"0.1.0"` |
| `Docxi.write_entry(zip, name, data)` | Escribe una parte en el ZIP con su tamaño fijado de antemano. Todas las partes lo usan internamente |
| `Docxi::Util.blank?`, `Docxi::Util.present?` | Métodos auxiliares internos con las reglas de `blank?` de ActiveSupport |

## <a id="related-pages"></a>Páginas relacionadas

- [Funciones](Features)
- [Ejemplos de uso](Usage-Examples)
- [Arquitectura](Architecture)

---

[← Inicio de la wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
