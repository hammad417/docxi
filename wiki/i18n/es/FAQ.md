<!-- Translated from wiki/FAQ.md (git blob e799a79ef71c9461e89ce6c6912798eb8c7b99c4). See wiki/TRANSLATION_GUIDE.md. -->
# Preguntas frecuentes sobre docxi: documentos de Word en Ruby y Rails

Respuestas breves a las preguntas más habituales. Cada respuesta enlaza a una explicación más larga.

**¿Necesito tener instalado Microsoft Word o LibreOffice?**
No. docxi escribe el XML y el archivo ZIP por sí mismo.

**¿docxi necesita Rails?**
No. Funciona en cualquier programa Ruby. Las pruebas también se ejecutan con Rails 5.2 a 8.1 cargado. Consulta [Integraciones](Integrations).

**¿Cómo genero un documento de Word en Ruby on Rails?**
Constrúyelo en un controlador o en una tarea en segundo plano y envía `word.render.string` con `send_data`. Consulta [Integraciones](Integrations).

**¿Qué versiones de Ruby y Rails se admiten?**
Ruby 2.5 y posteriores. Rails es opcional, y CI prueba Rails 5.2 a 8.1. Consulta [Instalación](Installation).

**¿Puede docxi editar un .docx existente o rellenar una plantilla de Word?**
No, solo crea documentos nuevos. Para plantillas con campos de combinación, usa una gema basada en plantillas, como sablon.

**¿La tabla de contenido se genera a partir de los títulos?**
No. Tú añades cada entrada y su número de página. Consulta [Funciones](Features#table-of-contents).

**¿Puedo usar listas numeradas o estilos de título?**
No en la 0.1.0. Los párrafos pueden ser elementos con viñetas (`ul: true`), y puedes dar formato a los títulos con opciones de texto como `size:` y `bold:`.

**¿Puedo cambiar el tamaño o la orientación de la página?**
Todavía no. Las páginas son US Letter en vertical, y solo puedes cambiar los márgenes. Consulta [Configuración](Configuration).

**¿Qué formatos de imagen funcionan?**
PNG. Todas las imágenes se guardan como una parte PNG.

**¿Puedo añadir una imagen desde una URL?**
Sí. `add_media` acepta URL `http`, `https` y `ftp`, y las descarga cuando llamas a `render`. No pases URL que vengan de los usuarios sin comprobarlas. Consulta [Seguridad](Security).

**¿Puedo escribir el documento directamente en un archivo?**
`render` devuelve un `StringIO`. Escribe su `.string` con `File.binwrite("out.docx", word.render.string)`.

**¿Puede docxi generar documentos muy grandes en streaming?**
No. Todo el documento se construye en memoria cuando llamas a `render`. Consulta [Despliegue](Deployment).

**¿Puedo definir párrafos de derecha a izquierda?**
La 0.1.0 no tiene opciones de derecha a izquierda. El texto se escribe tal como lo pasas.

**¿Dónde están los problemas conocidos?**
En la [hoja de ruta](Roadmap) y en la sección «Limitations and known issues» (Limitaciones y problemas conocidos) del README.

## <a id="related-pages"></a>Páginas relacionadas

- [Solución de problemas](Troubleshooting)
- [Funciones](Features)
- [Primeros pasos](Getting-Started)

---

[← Inicio de la wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
