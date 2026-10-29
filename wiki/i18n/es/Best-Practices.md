<!-- Translated from wiki/Best-Practices.md (git blob eb5af21efac1eb144bd0ebde76bd10c6b87d00ab). See wiki/TRANSLATION_GUIDE.md. -->
# Buenas prácticas con docxi: evita las sorpresas habituales

Estos hábitos salen directamente de cómo funciona docxi, y la mayoría evita alguno de los problemas que aparecen en [Solución de problemas](Troubleshooting). Están agrupados por tema. Los consejos de configuración de página sirven para todo el mundo, y los de seguridad y rendimiento importan sobre todo en aplicaciones en producción.

## <a id="page-setup"></a>Configuración de página

- **Pasa los cuatro márgenes juntos.** El hash que pasas a `Docxi::Document.new` sustituye a los valores predeterminados, y cualquiera de `top`, `right`, `bottom` y `left` que omitas se escribe vacío.
- **Define el encabezado y el pie de página de la primera página.** docxi siempre activa una primera página distinta. Si la página 1 necesita encabezado o pie de página, llama también a `add_first_header` y `add_first_footer`.

## <a id="text"></a>Texto

- **Usa `space: true` para los espacios iniciales o finales.** De lo contrario, Word los elimina y las palabras de dos fragmentos se juntan.
- **Pon el formato común en el párrafo.** Las opciones que se pasan a `p`, como `size:` o `font:`, se aplican a todos sus fragmentos, así que hay menos que repetir.
- **Usa `p.br` para los saltos de línea y `document.br(page: true)` para los saltos de página.** `document.br` sin `page: true` genera XML que no es válido.

## <a id="tables"></a>Tablas

- **Define los bordes en las filas o en las celdas, no en la tabla.** En la 0.1.0, los `borders:` a nivel de tabla se aplican a cada celda en lugar de al contorno de la tabla.
- **Indica `columns_width` para todas las columnas.** Las celdas toman su ancho de ahí, así que las tablas mantienen el diseño que esperas.

## <a id="images"></a>Imágenes

- **Usa archivos PNG.** Todas las imágenes se guardan como una parte PNG.
- **Define el tamaño de las imágenes en línea con `style:`**, por ejemplo `style: { width: "120pt", height: "60pt" }`. Es el único tamaño de imagen que funciona como cabría esperar en la 0.1.0.
- **Pasa una ruta en lugar de un `File` abierto si el documento puede generarse más de una vez.** Las imágenes IO se leen durante `render`, y una segunda generación obtiene una imagen vacía.
- **Imágenes de encabezado, de celdas de tabla y posicionadas hasta la 0.2.0.** Su ancho y su alto se calculan como `width * height * 14.92` EMU, así que salen cuadradas. Para obtener una imagen cuadrada de `S` pulgadas, pasa `width` y `height` iguales a `Math.sqrt(S * 914400 / 14.92).round`, que es aproximadamente 248 para una pulgada.

## <a id="security"></a>Seguridad

- **No pases cadenas no confiables a `add_media`.** Una ruta hace que docxi lea ese archivo de tu servidor, y una URL hace que tu servidor la descargue. Pasa un IO que hayas abierto tú, o comprueba antes el valor. Consulta [Seguridad](Security).
- **El texto se escapa automáticamente.** Las cadenas que se pasan a `text` se escriben como texto XML, así que la entrada de los usuarios no puede cambiar la estructura del documento.

## <a id="performance-and-reliability"></a>Rendimiento y fiabilidad

- **Genera los documentos grandes en una tarea en segundo plano.** Todo el documento se construye en memoria cuando llamas a `render`. Consulta [Integraciones](Integrations).
- **Genera una vez y reutiliza los bytes.** Guarda `word.render.string` si necesitas enviar o almacenar el documento más de una vez.
- **Fija una versión compatible** con `gem "docxi", "~> 0.1"`.
- **Prueba lo que generas.** Un `.docx` es un ZIP de archivos XML, así que las pruebas pueden descomprimirlo y comprobar el XML. Las propias pruebas de docxi funcionan así. Consulta [Arquitectura](Architecture).

## <a id="related-pages"></a>Páginas relacionadas

- [Funciones](Features)
- [Configuración](Configuration)
- [Solución de problemas](Troubleshooting)

---

[← Inicio de la wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
