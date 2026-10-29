<!-- Translated from wiki/Deployment.md (git blob 9058153b973e13980c113b2d7be11debb7bfff40). See wiki/TRANSLATION_GUIDE.md. -->
# 在生产环境中运行 docxi 以及发布新版本

docxi 是一个库，而不是服务，本身没有什么需要部署的。本页介绍发布使用 docxi 的应用时需要注意的事项，以及 docxi 新版本的发布方式。

## <a id="shipping-an-app-that-uses-docxi"></a>发布使用 docxi 的应用

- **依赖**：`bundle install` 会安装 nokogiri 和 rubyzip。nokogiri 为常见的 Linux、macOS 和 Windows 平台提供预编译 gem。在其他平台上（例如某些精简的容器镜像），它会从源码编译，此时需要 C 编译器。参见 [nokogiri 安装指南](https://nokogiri.org/tutorials/installing_nokogiri.html)。
- **不需要办公软件**：服务器和容器不需要 Microsoft Word 或 LibreOffice。
- **内存**：`render` 在内存中构建整个文档，并以 `StringIO` 的形式返回。内存占用随页数和图片大小而增长。
- **网络和文件**：docxi 只会为你传给 `add_media` 的图片路径读取文件系统，也只会为你传入的图片 URL 访问网络。参见[安全](Security)。
- **长文档**：请在后台任务中构建，而不是在 Web 请求过程中构建。参见[集成](Integrations)。

## <a id="how-docxi-releases-are-published"></a>docxi 版本的发布方式

以下步骤面向维护者。版本通过 GitHub Actions 以可信发布（trusted publishing）的方式发布到 RubyGems，因此仓库中不保存任何 API 密钥。

1. 更新 `lib/docxi/version.rb` 和 `CHANGELOG.md`，并合并到 `master`。
2. 给该提交打上以 `v` 为前缀的版本标签，例如 `v0.1.1`，然后推送该标签。
3. `.github/workflows/release.yml` 构建 gem 并推送到 RubyGems。

预发布版本遵循相同的步骤，版本号形如 `0.2.0.rc1`。只有在明确指定时，Bundler 才会安装预发布版本。CI（`.github/workflows/ci.yml`）会在每次推送到 `master` 以及每个 pull request 上运行。

## <a id="related-pages"></a>相关页面

- [安装](Installation)
- [开发环境搭建](Development-Setup)
- [更新日志](Changelog)

---

[← Wiki 首页](Home) · [README](https://github.com/hammad417/docxi#readme)
