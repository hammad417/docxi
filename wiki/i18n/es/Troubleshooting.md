<!-- Translated from wiki/Troubleshooting.md (git blob 76c38be3731e1d9e4696f2aab1eba96257a98970). See wiki/TRANSLATION_GUIDE.md. -->
# Solución de problemas de docxi: errores de instalación, errores en tiempo de ejecución y documentos dañados

Busca tu síntoma a continuación. La mayoría de los problemas de las aplicaciones antiguas vienen de docxi 0.0.x, y actualizar a la 0.1.0 los resuelve sin cambiar el código.

## <a id="installation"></a>Instalación

**Bundler no encuentra versiones compatibles de `nokogiri` en una aplicación con Rails 6+.**
docxi 0.0.6 fija `nokogiri < 1.7`, y todas las versiones de Rails desde la 6.0 necesitan nokogiri 1.8.5 o posterior. Usa `gem "docxi", "~> 0.1"`.

**Bundler no encuentra versiones compatibles de `rubyzip`.**
docxi 0.0.6 fija `rubyzip ~> 0.9.9`, lo que entra en conflicto con cualquier gema que necesite un rubyzip más reciente. docxi 0.1.0 acepta rubyzip 1.3 a 3.x.

**Bundler sigue instalando la 0.0.6.**
O bien el Gemfile fija docxi en una 0.0.x, por ejemplo con `"~> 0.0.6"`, o bien Ruby es anterior a la 2.5, que la 0.1.0 necesita. Cambia la línea del Gemfile a `gem "docxi", "~> 0.1"` y comprueba `ruby -v`.

**nokogiri no se compila.**
Ninguna gema precompilada de nokogiri coincide con tu plataforma, así que se está compilando desde el código fuente. Instala un compilador de C y sigue la [guía de instalación de nokogiri](https://nokogiri.org/tutorials/installing_nokogiri.html).

## <a id="errors-while-building-a-document"></a>Errores al construir un documento

**`NoMethodError: undefined method 'present?'` fuera de Rails.**
docxi 0.0.6 daba por hecho que ActiveSupport estaba cargado. La 0.1.0 no lo necesita.

**`NameError` al llamar a `tab`, `frame` o al `page_numbers` de un encabezado.**
Estaban rotos en la 0.0.6 y funcionan en la 0.1.0, incluido `document.frame("text")`.

**`Errno::ENOENT` al añadir una imagen.**
La ruta no existe respecto al directorio en el que se ejecuta tu programa. En Ruby 3 con docxi 0.0.6, las URL de imagen también fallan así, porque la 0.0.6 usaba `Kernel#open`. La 0.1.0 descarga las URL con `URI#open`.

## <a id="documents-that-dont-look-right"></a>Documentos que no se ven bien

**Google Docs o LibreOffice no abren el archivo.**
docxi 0.0.6 con rubyzip 3 escribe registros Zip64 en cada entrada, y estos programas los rechazan. La 0.1.0 escribe los archivos ZIP sin ellos.

**La primera página no tiene encabezado ni pie de página.**
docxi da a la primera página su propio encabezado y pie de página. Defínelos con `add_first_header` y `add_first_footer`. Consulta [Configuración](Configuration#first-page).

**Falta un margen o el diseño de la página se ve mal.**
Cuando definas los márgenes, pasa `top`, `right`, `bottom` y `left` juntos. Los que omitas se escriben vacíos.

**Desaparecen los espacios entre fragmentos de texto.**
Word elimina los espacios al principio o al final de un fragmento. Añade `space: true` a esa llamada a `text`.

**Las imágenes de los encabezados, de las celdas de tabla o con `pH`/`pV` salen diminutas o cuadradas.**
Problema conocido en la 0.1.0: su tamaño se calcula como `width * height * 14.92` EMU en ambas dimensiones. Usa imágenes en línea con tamaño definido mediante `style:`, o consulta la solución alternativa en [Buenas prácticas](Best-Practices#images).

**Falta una imagen la segunda vez que se genera el mismo documento.**
Las imágenes añadidas desde un `File` abierto u otro IO se leen durante `render`, y una segunda generación lee un flujo vacío. Pasa una ruta en su lugar, o genera el documento una sola vez.

**Los bordes definidos en `table` aparecen alrededor de cada celda.**
Problema conocido: los `borders:` a nivel de tabla se aplican a cada celda en lugar de al contorno de la tabla. Para dibujar solo el contorno, define `borders:` en las celdas del perímetro, solo para sus lados exteriores.

**Los números de página en las celdas de tabla empiezan por «GlobalOptions».**
Problema conocido en `TableCell#page_numbers`. Pon los números de página en un pie de página o en un encabezado.

**Una imagen JPEG o GIF se ve mal o no aparece.**
Todas las imágenes se guardan como una parte PNG con un tipo de contenido PNG, sea cual sea su formato real, así que algunos programas pueden no mostrar otros formatos. Convierte las imágenes a PNG antes de añadirlas.

## <a id="inspecting-a-generated-file"></a>Inspeccionar un archivo generado

Un `.docx` es un archivo ZIP, así que puedes ver lo que hay dentro:

```sh
unzip -l report.docx
unzip -p report.docx word/document.xml
```

Si el problema persiste, [abre un issue](https://github.com/hammad417/docxi/issues) con un script breve que lo reproduzca y las versiones de Ruby, docxi, nokogiri y rubyzip.

## <a id="related-pages"></a>Páginas relacionadas

- [Preguntas frecuentes](FAQ)
- [Buenas prácticas](Best-Practices)
- [Instalación](Installation)

---

[← Inicio de la wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
