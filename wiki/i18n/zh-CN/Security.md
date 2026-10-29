<!-- Translated from wiki/Security.md (git blob 2358bb39bd6772aaea144aa8ca2434b43c086728). See wiki/TRANSLATION_GUIDE.md. -->
# docxi 安全：报告问题与处理不可信输入

本页说明如何报告安全问题，以及当你的应用把用户输入传给 docxi 时需要了解的安全要点。

## <a id="reporting-a-vulnerability"></a>报告漏洞

请不要在公开的 issue 中报告安全问题，而是通过电子邮件联系维护者。邮箱地址在 [`docxi.gemspec`](https://github.com/hammad417/docxi/blob/master/docxi.gemspec) 中。请注明 docxi 的版本，并提供重现问题的简短方法。

## <a id="upgrade-from-00x"></a>从 0.0.x 升级

docxi 0.0.x 会把图片路径字符串传给 `Kernel#open`。`Kernel#open` 会把以 `|` 开头的字符串当作 shell 命令执行，因此能够控制图片路径的攻击者可以在你的服务器上执行命令。0.1.0 用 `File.open` 打开文件，用 `URI#open` 下载 URL，两者都不会执行命令。

## <a id="untrusted-input"></a>不可信输入

**传入文本是安全的**。文本、表格内容和其他字符串会作为 XML 文本或属性值写入，并由 nokogiri 转义。用户输入无法向文档中添加元素。

**图片来源需要谨慎处理**。`add_media` 接受：

| 来源 | docxi 的处理方式 |
|---|---|
| 你打开的 IO | 读取它 |
| 文件路径 | 从服务器磁盘读取该文件 |
| `http`、`https` 或 `ftp` URL | 从你的服务器下载它 |

如果用户可以指定路径，他们就能让服务器上任何可读的文件（例如配置或密钥）被嵌入到文档中。如果用户可以指定 URL，你的服务器就会去请求它，这可能会访问到内部地址（服务器端请求伪造，SSRF）。请传入你自己打开的 IO，或者只允许已知的路径和主机。

**超链接**会指向 URL 所指的任何地方。把用户提供的 URL 传给 `add_hyperlink` 之前，请先检查。

## <a id="how-releases-are-protected"></a>版本发布如何受到保护

- 版本通过 GitHub Actions 以 RubyGems 可信发布（trusted publishing）的方式发布，因此仓库中不保存任何 API 密钥。
- gem 元数据设置了 `rubygems_mfa_required`，因此使用个人 API 密钥推送时需要多因素认证。

## <a id="related-pages"></a>相关页面

- [最佳实践](Best-Practices)
- [部署](Deployment)
- [参与贡献](Contributing)

---

[← Wiki 首页](Home) · [README](https://github.com/hammad417/docxi#readme)
