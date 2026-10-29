<!-- Translated from wiki/Installation.md (git blob 0bd43a0434cbf17fa4dcd95d837e40ca50c52119). See wiki/TRANSLATION_GUIDE.md. -->
# Instalar docxi: requisitos, configuración de Bundler y actualización

docxi es una gema escrita solo en Ruby. Al instalarla se añaden dos dependencias, nokogiri y rubyzip, y nada más. No hace falta ningún software de ofimática.

## <a id="requirements"></a>Requisitos

| Requisito | Versión |
|---|---|
| Ruby | 2.5 o posterior |
| nokogiri | `>= 1.8, < 2`, se instala automáticamente |
| rubyzip | `>= 1.3, < 4`, se instala automáticamente |
| Rails | No es necesario |

## <a id="install-with-bundler"></a>Instalar con Bundler

Añade la gema a tu Gemfile:

```ruby
gem "docxi", "~> 0.1"
```

Después instálala:

```sh
bundle install
```

## <a id="install-without-bundler"></a>Instalar sin Bundler

```sh
gem install docxi
```

## <a id="check-the-installation"></a>Comprobar la instalación

```sh
ruby -e 'require "docxi"; puts Docxi::VERSION'
```

En un proyecto con Bundler, ejecútalo como `bundle exec ruby -e 'require "docxi"; puts Docxi::VERSION'`. Muestra la versión instalada, por ejemplo `0.1.0`.

## <a id="upgrading-from-00x"></a>Actualizar desde 0.0.x

Si tu Gemfile fija docxi en una 0.0.x, por ejemplo con `"~> 0.0.6"`, cambia la línea a `gem "docxi", "~> 0.1"`. Después ejecuta `bundle update docxi`. No hace falta cambiar el código, y la 0.1.0 genera los mismos documentos que la 0.0.6: el conjunto de pruebas compara la salida con archivos generados por la 0.0.6.

Qué ha cambiado internamente:

- nokogiri y rubyzip pueden tener ahora cualquier versión dentro de los rangos anteriores, así que docxi se instala junto a las versiones actuales de Rails.
- docxi ya no necesita que ActiveSupport esté cargado.
- Los archivos ZIP se escriben sin registros Zip64.
- Las imágenes se cargan sin `Kernel#open`, así que las URL funcionan en Ruby 3.
- `tab`, `frame` y los números de página en los encabezados funcionan.

Las aplicaciones con una versión de Ruby anterior a la 2.5 siguen recibiendo la 0.0.6, porque la 0.1.0 declara Ruby 2.5 como mínimo. El [registro de cambios](Changelog) enumera todos los cambios.

## <a id="versions-tested-in-ci"></a>Versiones probadas en CI

Cada versión de Rails se ejecuta con el framework completo cargado junto a docxi.

| Ruby | Rails |
|---|---|
| 2.5 | 5.2 |
| 2.7 | 6.0 |
| 3.0 | 6.1 |
| 3.1 | 7.0 |
| 3.2 | 7.1 |
| 3.3 | 7.2 |
| 3.4 | 8.0 |
| 4.0 | 8.1 |

CI también ejecuta:

- las pruebas sin Rails en Ruby 2.5, 2.6 y 4.0
- las pruebas con rubyzip 1.3 (Ruby 2.5), rubyzip 2.4 (Ruby 3.3) y nokogiri 1.11 (Ruby 2.5)
- una comprobación de que LibreOffice puede convertir los documentos de ejemplo

## <a id="platform-notes"></a>Notas sobre plataformas

- nokogiri publica gemas precompiladas para las plataformas Linux, macOS y Windows más comunes. Cuando ninguna coincide, se compila desde el código fuente y necesita un compilador de C. Consulta la [guía de instalación de nokogiri](https://nokogiri.org/tutorials/installing_nokogiri.html).
- CI se ejecuta en Ubuntu 24.04, y las pruebas también se han ejecutado en macOS. Windows no se ha probado.

## <a id="if-installation-fails"></a>Si la instalación falla

Los conflictos de Bundler con nokogiri o rubyzip casi siempre significan que se está resolviendo una versión antigua, docxi 0.0.6. Consulta [Solución de problemas](Troubleshooting#installation).

## <a id="related-pages"></a>Páginas relacionadas

- [Primeros pasos](Getting-Started)
- [Configuración](Configuration)
- [Despliegue](Deployment)

---

[← Inicio de la wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
