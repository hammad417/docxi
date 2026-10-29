<!-- Translated from wiki/Home.md (git blob 4724ddd12e6c213e92a003c3fc882558d8e2edfc). See wiki/TRANSLATION_GUIDE.md. -->
# Wiki de docxi: genera documentos de Word (.docx) con Ruby

docxi es una gema de Ruby que genera archivos `.docx` de Microsoft Word a partir de código. Ofrece párrafos, texto con formato, tablas, imágenes, encabezados y pies de página, números de página, hipervínculos y una tabla de contenido sencilla. No necesita Microsoft Word, LibreOffice ni una plantilla, y funciona tanto en Ruby puro como en aplicaciones Ruby on Rails.

El [README](https://github.com/hammad417/docxi#readme) es la introducción breve. Esta wiki cubre todas las opciones, incluye recetas listas para copiar, explica cómo funciona docxi por dentro y te ayuda cuando algo sale mal.

Versión actual: **0.1.1**, para Ruby 2.5 o posterior. Consulta el [registro de cambios](Changelog).

## <a id="start-here"></a>Empieza aquí

**Si es tu primera vez con docxi**

- [Primeros pasos](Getting-Started): instala la gema y crea tu primer documento
- [Instalación](Installation): requisitos, Bundler y actualización desde 0.0.x
- [Ejemplos de uso](Usage-Examples): informes, facturas, cartas y tablas a partir de datos, listos para copiar

**Crear documentos**

- [Funciones](Features): cada elemento y las opciones que acepta
- [Configuración](Configuration): márgenes, configuración de página y unidades
- [Referencia de la API](API-Reference): clases y métodos de un vistazo
- [Buenas prácticas](Best-Practices): hábitos que evitan las sorpresas habituales

**Cuando algo sale mal**

- [Solución de problemas](Troubleshooting): mensajes de error y documentos que no se ven bien
- [Preguntas frecuentes](FAQ): respuestas breves a preguntas comunes
- [Seguridad](Security): cómo informar de problemas y manejar entradas no confiables

**Para profundizar**

- [Arquitectura](Architecture): cómo las llamadas de Ruby se convierten en un paquete `.docx`
- [Integraciones](Integrations): controladores de Rails, aplicaciones Rack y tareas en segundo plano
- [Despliegue](Deployment): docxi en producción y cómo se publican las versiones
- [Contribuir](Contributing) y [Entorno de desarrollo](Development-Setup)
- [Hoja de ruta](Roadmap), [Registro de cambios](Changelog) y [Créditos](Credits)

## <a id="what-docxi-does-and-what-it-doesnt"></a>Qué hace docxi y qué no

docxi crea documentos nuevos. No puede abrir ni editar un `.docx` existente, rellenar una plantilla de Word ni hacer una combinación de correspondencia. No tiene listas numeradas, estilos de título ni tabla de contenido automática, y las páginas siempre son US Letter en vertical. Los problemas conocidos previstos para la 0.2.0 están en la [hoja de ruta](Roadmap).

## <a id="credits"></a>Créditos

docxi lo mantiene [Hammad Maqbool](https://github.com/hammad417) en [Phaedra Solutions](https://www.phaedrasolutions.com/). Nació en 2016 a partir de [docxer](https://github.com/gtimti/docxer), de Timur Gabdrakipov. Consulta [Créditos](Credits) para más detalles.

## <a id="languages"></a>Idiomas

La versión en inglés es la versión de referencia de esta wiki. Las traducciones están enlazadas en la barra lateral y al principio de cada página.

---

[README](https://github.com/hammad417/docxi#readme) · [RubyGems](https://rubygems.org/gems/docxi) · [Informar de un problema](https://github.com/hammad417/docxi/issues)
