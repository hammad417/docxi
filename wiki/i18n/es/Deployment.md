<!-- Translated from wiki/Deployment.md (git blob 9058153b973e13980c113b2d7be11debb7bfff40). See wiki/TRANSLATION_GUIDE.md. -->
# docxi en producción y publicación de versiones

docxi es una biblioteca, no un servicio, así que no hay nada que desplegar por sí solo. Esta página cubre lo que importa cuando publicas una aplicación que usa docxi y cómo se publican las nuevas versiones de docxi.

## <a id="shipping-an-app-that-uses-docxi"></a>Publicar una aplicación que usa docxi

- **Dependencias.** `bundle install` instala nokogiri y rubyzip. nokogiri publica gemas precompiladas para las plataformas Linux, macOS y Windows más comunes. En otras plataformas, como algunas imágenes de contenedor mínimas, se compila desde el código fuente y necesita un compilador de C. Consulta la [guía de instalación de nokogiri](https://nokogiri.org/tutorials/installing_nokogiri.html).
- **Sin software de ofimática.** Los servidores y contenedores no necesitan Microsoft Word ni LibreOffice.
- **Memoria.** `render` construye todo el documento en memoria y lo devuelve como un `StringIO`. El uso de memoria crece con el número de páginas y el tamaño de las imágenes.
- **Red y archivos.** docxi solo lee el sistema de archivos para las rutas de imagen que pasas a `add_media`, y solo usa la red para las URL de imagen que pasas. Consulta [Seguridad](Security).
- **Documentos largos.** Genéralos en una tarea en segundo plano en lugar de durante una petición web. Consulta [Integraciones](Integrations).

## <a id="how-docxi-releases-are-published"></a>Cómo se publican las versiones de docxi

Estos pasos son para quienes mantienen el proyecto. Las versiones se publican en RubyGems mediante GitHub Actions con publicación de confianza (trusted publishing), así que no se guarda ninguna clave de API en el repositorio.

1. Actualiza `lib/docxi/version.rb` y `CHANGELOG.md`, y fusiona los cambios en `master`.
2. Etiqueta el commit con la versión precedida de `v`, por ejemplo `v0.1.1`, y sube la etiqueta.
3. `.github/workflows/release.yml` construye la gema y la publica en RubyGems.

Las versiones preliminares siguen los mismos pasos con una versión como `0.2.0.rc1`. Bundler solo instala versiones preliminares cuando se piden explícitamente. CI (`.github/workflows/ci.yml`) se ejecuta en cada push a `master` y en cada pull request.

## <a id="related-pages"></a>Páginas relacionadas

- [Instalación](Installation)
- [Entorno de desarrollo](Development-Setup)
- [Registro de cambios](Changelog)

---

[← Inicio de la wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
