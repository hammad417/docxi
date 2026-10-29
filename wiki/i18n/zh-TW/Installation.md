<!-- Translated from wiki/Installation.md (git blob 0bd43a0434cbf17fa4dcd95d837e40ca50c52119). See wiki/TRANSLATION_GUIDE.md. -->
# 安裝 docxi：系統需求、Bundler 設定與升級

docxi 是純 Ruby 的 gem。安裝時只會引入兩個相依套件：nokogiri 和 rubyzip，不需要任何辦公室軟體。

## <a id="requirements"></a>系統需求

| 需求 | 版本 |
|---|---|
| Ruby | 2.5 以上版本 |
| nokogiri | `>= 1.8, < 2`，自動安裝 |
| rubyzip | `>= 1.3, < 4`，自動安裝 |
| Rails | 不需要 |

## <a id="install-with-bundler"></a>使用 Bundler 安裝

在 Gemfile 中加入這個 gem：

```ruby
gem "docxi", "~> 0.1"
```

然後安裝：

```sh
bundle install
```

## <a id="install-without-bundler"></a>不使用 Bundler 安裝

```sh
gem install docxi
```

## <a id="check-the-installation"></a>確認安裝

```sh
ruby -e 'require "docxi"; puts Docxi::VERSION'
```

在 Bundler 專案中，請以 `bundle exec ruby -e 'require "docxi"; puts Docxi::VERSION'` 的形式執行。它會印出已安裝的版本，例如 `0.1.0`。

## <a id="upgrading-from-00x"></a>從 0.0.x 升級

如果 Gemfile 把 docxi 固定在 0.0.x（例如 `"~> 0.0.6"`），請先把這一行改為 `gem "docxi", "~> 0.1"`，然後執行 `bundle update docxi`。不需要修改程式碼，0.1.0 產生的文件與 0.0.6 相同：測試套件會將輸出與 0.0.6 產生的檔案進行比對。

底層的變更：

- nokogiri 和 rubyzip 現在可以是上述範圍內的任何版本，因此 docxi 可以與目前的 Rails 版本一起安裝。
- docxi 不再需要載入 ActiveSupport。
- 寫出的壓縮檔不含 Zip64 記錄。
- 載入圖片時不再使用 `Kernel#open`，因此 URL 在 Ruby 3 上可以正常使用。
- `tab`、`frame` 和頁首中的頁碼都能正常運作。

使用 Ruby 2.5 以下版本的應用程式會繼續安裝 0.0.6，因為 0.1.0 宣告的最低 Ruby 版本是 2.5。[變更紀錄](Changelog)列出了所有變更。

## <a id="versions-tested-in-ci"></a>CI 測試的版本

每個 Rails 版本都會在完整載入框架的情況下與 docxi 一起執行。

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

CI 還會執行：

- 在 Ruby 2.5、2.6 和 4.0 上不載入 Rails 的測試
- 使用 rubyzip 1.3（Ruby 2.5）、rubyzip 2.4（Ruby 3.3）和 nokogiri 1.11（Ruby 2.5）的測試
- 確認 LibreOffice 能轉換範例文件的檢查

## <a id="platform-notes"></a>平台說明

- nokogiri 為常見的 Linux、macOS 和 Windows 平台提供預先編譯的 gem。如果沒有相符的版本，它會從原始碼編譯，這時需要 C 編譯器。請參閱 [nokogiri 安裝指南](https://nokogiri.org/tutorials/installing_nokogiri.html)。
- CI 在 Ubuntu 24.04 上執行，測試也曾在 macOS 上執行過。Windows 未經測試。

## <a id="if-installation-fails"></a>安裝失敗時

Bundler 在 nokogiri 或 rubyzip 上回報衝突，幾乎都是因為解析到了舊版 docxi 0.0.6。請參閱[疑難排解](Troubleshooting#installation)。

## <a id="related-pages"></a>相關頁面

- [快速入門](Getting-Started)
- [設定](Configuration)
- [部署](Deployment)

---

[← Wiki 首頁](Home) · [README](https://github.com/hammad417/docxi#readme)
