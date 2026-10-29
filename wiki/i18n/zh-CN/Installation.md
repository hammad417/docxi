<!-- Translated from wiki/Installation.md (git blob 0bd43a0434cbf17fa4dcd95d837e40ca50c52119). See wiki/TRANSLATION_GUIDE.md. -->
# 安装 docxi：环境要求、Bundler 配置与升级

docxi 是一个纯 Ruby gem。安装时只会引入两个依赖：nokogiri 和 rubyzip，不需要任何办公软件。

## <a id="requirements"></a>环境要求

| 要求 | 版本 |
|---|---|
| Ruby | 2.5 或更高版本 |
| nokogiri | `>= 1.8, < 2`，自动安装 |
| rubyzip | `>= 1.3, < 4`，自动安装 |
| Rails | 不需要 |

## <a id="install-with-bundler"></a>使用 Bundler 安装

在 Gemfile 中添加该 gem：

```ruby
gem "docxi", "~> 0.1"
```

然后安装：

```sh
bundle install
```

## <a id="install-without-bundler"></a>不使用 Bundler 安装

```sh
gem install docxi
```

## <a id="check-the-installation"></a>检查安装

```sh
ruby -e 'require "docxi"; puts Docxi::VERSION'
```

在 Bundler 项目中，请以 `bundle exec ruby -e 'require "docxi"; puts Docxi::VERSION'` 的形式运行。它会输出已安装的版本，例如 `0.1.0`。

## <a id="upgrading-from-00x"></a>从 0.0.x 升级

如果 Gemfile 把 docxi 固定在 0.0.x（例如 `"~> 0.0.6"`），请先把这一行改为 `gem "docxi", "~> 0.1"`，然后运行 `bundle update docxi`。不需要修改代码，0.1.0 生成的文档与 0.0.6 相同：测试套件会把输出与 0.0.6 生成的文件进行比较。

底层的变化：

- nokogiri 和 rubyzip 现在可以是上述范围内的任意版本，因此 docxi 可以与当前的 Rails 版本一起安装。
- docxi 不再需要加载 ActiveSupport。
- 写出的压缩包不含 Zip64 记录。
- 加载图片不再使用 `Kernel#open`，因此 URL 在 Ruby 3 上可以正常使用。
- `tab`、`frame` 和页眉中的页码可以正常工作。

运行 Ruby 2.5 以下版本的应用会继续安装 0.0.6，因为 0.1.0 声明的最低 Ruby 版本是 2.5。[更新日志](Changelog)列出了所有变更。

## <a id="versions-tested-in-ci"></a>CI 中测试的版本

每个 Rails 版本都会在完整加载框架的情况下与 docxi 一起运行。

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

CI 还会运行：

- 在 Ruby 2.5、2.6 和 4.0 上不加载 Rails 的测试
- 使用 rubyzip 1.3（Ruby 2.5）、rubyzip 2.4（Ruby 3.3）和 nokogiri 1.11（Ruby 2.5）的测试
- 检查 LibreOffice 能否转换示例文档

## <a id="platform-notes"></a>平台说明

- nokogiri 为常见的 Linux、macOS 和 Windows 平台提供预编译 gem。如果没有匹配的版本，它会从源码编译，此时需要 C 编译器。参见 [nokogiri 安装指南](https://nokogiri.org/tutorials/installing_nokogiri.html)。
- CI 运行在 Ubuntu 24.04 上，测试也在 macOS 上运行过。Windows 未经测试。

## <a id="if-installation-fails"></a>安装失败时

Bundler 在 nokogiri 或 rubyzip 上报告冲突，几乎总是因为解析到了旧的 docxi 0.0.6。参见[故障排除](Troubleshooting#installation)。

## <a id="related-pages"></a>相关页面

- [快速入门](Getting-Started)
- [配置](Configuration)
- [部署](Deployment)

---

[← Wiki 首页](Home) · [README](https://github.com/hammad417/docxi#readme)
