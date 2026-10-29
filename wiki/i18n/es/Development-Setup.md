<!-- Translated from wiki/Development-Setup.md (git blob 6fd2ffba3077016b05d1d9ab9467d15fc88c8e7d). See wiki/TRANSLATION_GUIDE.md. -->
# Entorno de desarrollo: ejecutar las pruebas y herramientas de docxi

Esta página te ayuda a poner en marcha una copia de trabajo de docxi y explica las pruebas, los documentos de ejemplo y la configuración de CI.

## <a id="get-the-code"></a>Obtener el código

```sh
git clone https://github.com/hammad417/docxi.git
cd docxi
bundle install
bundle exec rake test
```

Para desarrollar sirve cualquier Ruby a partir de la 2.5. Las dependencias de desarrollo son minitest y rake.

## <a id="rake-tasks"></a>Tareas de Rake

| Comando | Qué hace |
|---|---|
| `bundle exec rake test` | Ejecuta el conjunto de pruebas (la tarea predeterminada) |
| `bundle exec rake samples` | Escribe los documentos de ejemplo en `tmp/samples/*.docx`, para abrirlos en Word, LibreOffice o Google Docs |
| `bundle exec rake golden` | Reescribe `test/golden` a partir del código actual. Úsala solo después de un cambio intencionado en la salida |
| `bundle exec rake build` | Construye la gema en `pkg/` |

## <a id="running-the-tests-with-rails-loaded"></a>Ejecutar las pruebas con Rails cargado

La carpeta `gemfiles/` tiene un Gemfile por cada versión de Rails, además de algunos para versiones antiguas de rubyzip y nokogiri:

```sh
BUNDLE_GEMFILE=gemfiles/rails_8.1.gemfile bundle install
BUNDLE_GEMFILE=gemfiles/rails_8.1.gemfile DOCXI_WITH_RAILS=1 bundle exec rake test
```

`DOCXI_WITH_RAILS=1` carga el framework completo antes de las pruebas. También compara las reglas de `blank?` de docxi con las de ActiveSupport.

## <a id="what-the-tests-cover"></a>Qué cubren las pruebas

| Archivo | Comprueba |
|---|---|
| `test/test_golden_output.rb` | Cada parte de dos documentos de ejemplo coincide con lo que generó la 0.0.6 |
| `test/test_package.rb` | El XML está bien formado, las relaciones apuntan a partes existentes, cada parte tiene un tipo de contenido y no hay registros Zip64 |
| `test/test_media.rb` | Imágenes desde rutas, objetos IO y URL, y que una cadena con `\|` nunca se ejecuta como comando |
| `test/test_fixed_features.rb` | `tab`, `frame` y los números de página en los encabezados |
| `test/test_util.rb` | `Docxi::Util.blank?` sigue las reglas de ActiveSupport |
| `test/test_rails_independence.rb` | docxi funciona sin ActiveSupport |
| `test/test_warnings.rb` | La generación no emite advertencias de obsolescencia desde `lib/` |

Los documentos de ejemplo están en `test/support/sample_documents.rb`, y `test/support/package.rb` lee un `.docx` directamente de sus estructuras ZIP.

## <a id="continuous-integration"></a>Integración continua

`.github/workflows/ci.yml` se ejecuta en cada push a `master` y en cada pull request. Cubre:

- Ruby 2.5, 2.6 y 4.0 sin Rails
- Rails 5.2 (Ruby 2.5), 6.0 (2.7), 6.1 (3.0), 7.0 (3.1), 7.1 (3.2), 7.2 (3.3), 8.0 (3.4) y 8.1 (4.0)
- Rails `main`, al que se le permite fallar
- rubyzip 1.3 y 2.4, y nokogiri 1.11
- un trabajo que convierte los documentos de ejemplo a PDF con LibreOffice

Las versiones se tratan en [Despliegue](Deployment).

## <a id="related-pages"></a>Páginas relacionadas

- [Contribuir](Contributing)
- [Arquitectura](Architecture)
- [Despliegue](Deployment)

---

[← Inicio de la wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
