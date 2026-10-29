<!-- Translated from wiki/Getting-Started.md (git blob fa1a6a4130a2bf1dbae73e806e7c6d261123f8d7). See wiki/TRANSLATION_GUIDE.md. -->
# Primeros pasos con docxi: tu primer documento de Word en Ruby

Esta guía te lleva de una carpeta vacía a un archivo `.docx` con formato. Se da por hecho que conoces lo básico de Ruby. No necesitas Rails ni Microsoft Word.

## <a id="1-install-the-gem"></a>1. Instala la gema

```sh
gem install docxi
```

En un proyecto con Gemfile, añade `gem "docxi", "~> 0.1"` y ejecuta `bundle install` en su lugar. Encontrarás los detalles en [Instalación](Installation).

## <a id="2-create-a-document"></a>2. Crea un documento

Guarda esto como `hello.rb`:

```ruby
require "docxi"

word = Docxi::Document.new
document = word.document

document.p(align: "center") do |p|
  p.text "Hello, Word", size: 24, bold: true
end

File.binwrite("hello.docx", word.render.string)
```

Ejecuta `ruby hello.rb` (o `bundle exec ruby hello.rb` en un proyecto con Bundler) y abre `hello.docx`.

## <a id="3-what-each-line-does"></a>3. Qué hace cada línea

- `Docxi::Document.new` crea un documento vacío con los márgenes de página predeterminados.
- `word.document` es el cuerpo del documento. Todo lo que añades pasa por él.
- `document.p` añade un párrafo. Las opciones que se pasan a `p`, como `align:`, se aplican al párrafo y al texto que contiene.
- `p.text` añade un fragmento de texto (un run). `size:` se expresa en puntos.
- `word.render` construye el paquete `.docx` y devuelve un `StringIO`. Su `.string` contiene los bytes del archivo.

## <a id="4-add-a-list-and-a-table"></a>4. Añade una lista y una tabla

```ruby
require "docxi"

word = Docxi::Document.new
document = word.document

document.p { |p| p.text "Quarterly summary", size: 16, bold: true }

document.p do |p|
  p.text "Revenue grew ", space: true
  p.text "12%", bold: true, color: "007A33"
  p.text " compared with last quarter.", space: true
end

document.p(ul: true) { |p| p.text "New customers: 48" }
document.p(ul: true) { |p| p.text "Customers lost: 3" }

document.table(columns_width: [200, 100]) do |table|
  table.tr(fill: "D9D9D9", bold: true) do |tr|
    tr.tc { |tc| tc.text "Region" }
    tr.tc { |tc| tc.text "Sales" }
  end
  table.tr do |tr|
    tr.tc { |tc| tc.text "North" }
    tr.tc { |tc| tc.text "$41,000", align: "right" }
  end
end

File.binwrite("summary.docx", word.render.string)
```

Dos cosas a tener en cuenta:

- `space: true` conserva el espacio al principio o al final de un fragmento. Sin esa opción, Word elimina el espacio y las palabras quedan pegadas.
- Las opciones de `tr` se aplican a todas las celdas de esa fila, por eso toda la fila de encabezado queda sombreada y en negrita.

## <a id="5-where-to-go-next"></a>5. Siguientes pasos

- [Ejemplos de uso](Usage-Examples): recetas completas para informes, facturas y cartas
- [Funciones](Features): todas las opciones de párrafos, texto, tablas, imágenes, encabezados y pies de página
- [Integraciones](Integrations): enviar el documento desde un controlador de Rails
- [Buenas prácticas](Best-Practices): evita las sorpresas habituales antes de que ocurran

---

[← Inicio de la wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
