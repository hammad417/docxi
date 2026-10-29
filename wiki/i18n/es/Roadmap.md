<!-- Translated from wiki/Roadmap.md (git blob c9fde5b9818383a900ca60bf167fac5895ec82c5). See wiki/TRANSLATION_GUIDE.md. -->
# Hoja de ruta y problemas conocidos de docxi

docxi no tiene un calendario de versiones ni una hoja de ruta comprometida. El único trabajo previsto que consta es corregir los problemas conocidos de abajo. Están previstos para la 0.2.0, sin fecha. Se dejaron fuera de la 0.1.0 porque corregirlos cambia los documentos que los usuarios actuales ya generan.

## <a id="known-issues-planned-for-020"></a>Problemas conocidos previstos para la 0.2.0

| Problema | Solución alternativa en la 0.1.0 |
|---|---|
| Las imágenes posicionadas, de encabezado y de celdas de tabla calculan su tamaño como `width * height * 14.92` EMU en ambas dimensiones | Usa imágenes en línea con tamaño definido mediante `style:`, o consulta [Buenas prácticas](Best-Practices#images) |
| Los `borders:` a nivel de tabla se aplican a cada celda en lugar de al contorno de la tabla | Para obtener solo el contorno, define `borders:` en los lados exteriores de las celdas del perímetro |
| Al código del campo de número de página le falta la barra invertida de `\* MERGEFORMAT` | Ninguna |
| Los números de página en las celdas de tabla llevan delante «GlobalOptions» | Pon los números de página en un encabezado o pie de página |
| Las imágenes siempre se guardan como partes PNG, y `[Content_Types].xml` solo declara los dos primeros encabezados y pies de página | Usa imágenes PNG y como máximo dos encabezados y dos pies de página |
| Pasar solo algunos márgenes escribe vacíos los demás | Pasa `top`, `right`, `bottom` y `left` juntos |
| `docProps/core.xml` registra la hora local con un indicador UTC y nunca escribe el autor | Ninguna |
| `Docxi.test` lanza `NoMethodError` | No lo llames. Es un método auxiliar que quedó de versiones anteriores |
| `document.br` y `document.tab` sin `page: true` generan XML no válido | Usa `p.br` y `p.tab` dentro de un párrafo |

## <a id="not-planned"></a>No previsto

Estos puntos no están en la lista anterior, y no se ha decidido nada sobre ellos:

- editar documentos existentes o rellenar plantillas
- listas numeradas, estilos de título o una tabla de contenido automática
- otros tamaños de página o páginas en horizontal

## <a id="related-pages"></a>Páginas relacionadas

- [Registro de cambios](Changelog)
- [Contribuir](Contributing)
- [Preguntas frecuentes](FAQ)

---

[← Inicio de la wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
