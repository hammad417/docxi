<!-- Translated from wiki/Security.md (git blob 2358bb39bd6772aaea144aa8ca2434b43c086728). See wiki/TRANSLATION_GUIDE.md. -->
# Seguridad en docxi: cómo informar de problemas y manejar entradas no confiables

Esta página explica cómo informar de un problema de seguridad y qué aspectos de seguridad conviene conocer cuando tu aplicación pasa entradas de usuarios a docxi.

## <a id="reporting-a-vulnerability"></a>Informar de una vulnerabilidad

No informes de problemas de seguridad en issues públicos. Escríbelos por correo electrónico a las direcciones de mantenimiento que figuran en [`docxi.gemspec`](https://github.com/hammad417/docxi/blob/master/docxi.gemspec). Incluye la versión de docxi y una forma breve de reproducir el problema.

## <a id="upgrade-from-00x"></a>Actualiza desde 0.0.x

docxi 0.0.x pasa las cadenas de ruta de imagen a `Kernel#open`. `Kernel#open` ejecuta como comando de shell una cadena que empieza por `|`, así que quien controle una ruta de imagen podría ejecutar comandos en tu servidor. La 0.1.0 abre los archivos con `File.open` y descarga las URL con `URI#open`. Ninguno de los dos ejecuta comandos.

## <a id="untrusted-input"></a>Entradas no confiables

**El texto se puede pasar sin riesgo.** El texto, el contenido de las tablas y otras cadenas se escriben como texto XML o valores de atributo, que nokogiri escapa. La entrada de los usuarios no puede añadir elementos al documento.

**Los orígenes de imagen requieren cuidado.** `add_media` acepta:

| Origen | Qué hace docxi |
|---|---|
| Un IO que has abierto tú | Lo lee |
| Una ruta de archivo | Lee ese archivo del disco de tu servidor |
| Una URL `http`, `https` o `ftp` | La descarga desde tu servidor |

Si un usuario puede elegir la ruta, podría conseguir que cualquier archivo legible de tu servidor, como la configuración o las claves, acabe incrustado en el documento. Si un usuario puede elegir la URL, tu servidor la solicitará, y eso puede alcanzar direcciones internas (falsificación de solicitudes del lado del servidor, SSRF). Pasa un IO que hayas abierto tú, o permite solo rutas y hosts conocidos.

**Los hipervínculos** apuntan a donde diga la URL. Comprueba las URL que proporcionan los usuarios antes de pasarlas a `add_hyperlink`.

## <a id="how-releases-are-protected"></a>Cómo se protegen las versiones

- Las versiones se publican desde GitHub Actions con la publicación de confianza (trusted publishing) de RubyGems, así que no se guarda ninguna clave de API en el repositorio.
- Los metadatos de la gema definen `rubygems_mfa_required`, así que las publicaciones hechas con una clave de API personal requieren autenticación multifactor.

## <a id="related-pages"></a>Páginas relacionadas

- [Buenas prácticas](Best-Practices)
- [Despliegue](Deployment)
- [Contribuir](Contributing)

---

[← Inicio de la wiki](Home) · [README](https://github.com/hammad417/docxi#readme)
