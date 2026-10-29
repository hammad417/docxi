<!-- Translated from wiki/Changelog.md (git blob 4dabeabaf2ac6e0562ba7ff98a24098001aa1e67). See wiki/TRANSLATION_GUIDE.md. -->
# Registro de cambios e historial de versiones de docxi

Esta página resume cada versión. Las notas completas están en [`CHANGELOG.md`](https://github.com/hammad417/docxi/blob/master/CHANGELOG.md), en el repositorio, que es la referencia si alguna vez difieren.

## <a id="releases"></a>Versiones

| Versión | Fecha | Notas |
|---|---|---|
| 0.1.1 | 2026-10-29 | Resumen y descripción de la gema más precisos, y enlace a la documentación. Sin cambios en el código |
| 0.1.0 | 2026-10-28 | Funciona con las versiones actuales de Ruby, Rails y rubyzip. Genera los mismos documentos que la 0.0.6 |
| 0.1.0.rc1 | 2026-10-28 | Versión preliminar de la 0.1.0 |
| 0.0.6 | 2016-10-18 | Última versión antes de la actualización 0.1.0 |
| 0.0.2 | 2016-10-14 | No se registraron notas de la versión |
| 0.0.1 | 2016-10-14 | Primera versión. No se registraron notas de la versión |

Las versiones 0.0.3 a 0.0.5 nunca se publicaron en RubyGems.

## <a id="011"></a>0.1.1

Sin cambios en el código, y los documentos generados son los mismos que con la 0.1.0.

- El resumen y la descripción de la gema ahora explican qué hace docxi. Antes lo presentaban como un generador de plantillas.
- RubyGems enlaza a esta wiki como documentación de la gema.

## <a id="010"></a>0.1.0

**Compatibilidad**

- Requiere Ruby 2.5 o posterior. Las aplicaciones con versiones de Ruby más antiguas se quedan en la 0.0.6.
- Funciona con o sin Rails, y se prueba junto a Rails 5.2 a 8.1.
- nokogiri `>= 1.8, < 2` en lugar de `>= 1.6.3, < 1.7`, que entraba en conflicto con todas las versiones de Rails desde la 6.0.
- rubyzip `>= 1.3, < 4` en lugar de `~> 0.9.9`.
- Ya no llama a `blank?` y `present?` de ActiveSupport, que solo funcionaban cuando Rails los había cargado.

**Correcciones**

- Los documentos escritos con rubyzip 3 ya no contienen registros Zip64, que Google Docs y las versiones antiguas de LibreOffice rechazan.
- `add_media` con una URL de imagen vuelve a funcionar en Ruby 3.
- `add_media` ya no ejecuta como comando de shell una cadena que empieza por `|`.
- `tab` y `frame` ya no lanzan `NameError`.
- `document.frame("text")` crea un cuadro de texto que contiene el texto.
- `page_numbers` en un encabezado ya no lanza `NameError`.

**Gema**

- Cambio en el mantenimiento del proyecto y nuevos enlaces a la página principal y al registro de cambios.
- Las publicaciones requieren MFA.

## <a id="related-pages"></a>Páginas relacionadas

- [Hoja de ruta](Roadmap)
- [Instalación](Installation#upgrading-from-00x)

---

[← Inicio de la wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
