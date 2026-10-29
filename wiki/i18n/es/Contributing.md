<!-- Translated from wiki/Contributing.md (git blob a37042b915e5170b5cf706a4c253276b50db8aa0). See wiki/TRANSLATION_GUIDE.md. -->
# Contribuir a docxi

Los informes de errores, las correcciones, la documentación y las traducciones son bienvenidos. Esta página explica cómo informar de un problema y qué necesita un pull request. Para preparar una copia de trabajo, consulta [Entorno de desarrollo](Development-Setup).

## <a id="reporting-a-bug"></a>Informar de un error

Abre un issue en [github.com/hammad417/docxi/issues](https://github.com/hammad417/docxi/issues) con:

- un script breve de Ruby que reproduzca el problema
- las versiones de Ruby, docxi, nokogiri y rubyzip (y de Rails, si lo usas)
- lo que esperabas y lo que obtuviste. Si el documento sale dañado, indica con qué programa lo abriste

Los problemas de seguridad se informan en privado. Consulta [Seguridad](Security).

## <a id="pull-requests"></a>Pull requests

1. Haz un fork del repositorio y crea una rama.
2. Haz el cambio respetando el estilo del código que lo rodea.
3. Añade o actualiza las pruebas y ejecuta `bundle exec rake test`.
4. Si el XML generado cambia a propósito, ejecuta `bundle exec rake golden` y explica el cambio en el pull request. Los archivos de referencia protegen la salida de la que dependen los usuarios actuales.
5. Actualiza `CHANGELOG.md` y la documentación si cambia el comportamiento.
6. Abre el pull request. CI ejecuta las pruebas en Ruby 2.5 a 4.0 y Rails 5.2 a 8.1, y comprueba que LibreOffice abre los documentos de ejemplo.

Los cambios que modifican documentos existentes, como las correcciones previstas para la 0.2.0, van en su propio pull request para que la diferencia en los archivos de referencia sea fácil de revisar.

## <a id="documentation-and-translations"></a>Documentación y traducciones

- El README en inglés es la fuente de 15 traducciones en `docs/i18n/`. Sigue `docs/i18n/TRANSLATION_GUIDE.md` del repositorio cuando lo modifiques.
- Esta wiki se escribe en la carpeta `wiki/` del repositorio y se publica desde ahí. `wiki/README.md` explica cómo.

## <a id="related-pages"></a>Páginas relacionadas

- [Entorno de desarrollo](Development-Setup)
- [Arquitectura](Architecture)
- [Hoja de ruta](Roadmap)

---

[← Inicio de la wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
