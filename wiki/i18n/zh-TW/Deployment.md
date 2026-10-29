<!-- Translated from wiki/Deployment.md (git blob 9058153b973e13980c113b2d7be11debb7bfff40). See wiki/TRANSLATION_GUIDE.md. -->
# 在正式環境中執行 docxi 以及發布新版本

docxi 是程式庫，而不是服務，本身沒有什麼需要部署的。本頁介紹交付使用 docxi 的應用程式時需要注意的事項，以及 docxi 新版本的發布方式。

## <a id="shipping-an-app-that-uses-docxi"></a>交付使用 docxi 的應用程式

- **相依套件**：`bundle install` 會安裝 nokogiri 和 rubyzip。nokogiri 為常見的 Linux、macOS 和 Windows 平台提供預先編譯的 gem。在其他平台上（例如某些精簡的容器映像檔），它會從原始碼編譯，這時需要 C 編譯器。請參閱 [nokogiri 安裝指南](https://nokogiri.org/tutorials/installing_nokogiri.html)。
- **不需要辦公室軟體**：伺服器和容器不需要 Microsoft Word 或 LibreOffice。
- **記憶體**：`render` 在記憶體中建立整份文件，並以 `StringIO` 的形式傳回。記憶體用量會隨頁數和圖片大小增加。
- **網路和檔案**：docxi 只會為了你傳給 `add_media` 的圖片路徑讀取檔案系統，也只會為了你傳入的圖片 URL 使用網路。請參閱[安全性](Security)。
- **長篇文件**：請在背景工作中建立，而不是在 Web 請求期間建立。請參閱[整合](Integrations)。

## <a id="how-docxi-releases-are-published"></a>docxi 版本的發布方式

以下步驟供維護者參考。版本透過 GitHub Actions 以可信任發布（trusted publishing）的方式發布到 RubyGems，因此儲存庫中不會存放任何 API 金鑰。

1. 更新 `lib/docxi/version.rb` 和 `CHANGELOG.md`，並合併到 `master`。
2. 為該提交加上以 `v` 為前綴的版本標籤，例如 `v0.1.1`，然後推送該標籤。
3. `.github/workflows/release.yml` 會建置 gem 並推送到 RubyGems。

預先發行版本遵循相同的步驟，版本號碼形如 `0.2.0.rc1`。只有在明確指定時，Bundler 才會安裝預先發行版本。CI（`.github/workflows/ci.yml`）會在每次推送到 `master` 以及每個 pull request 上執行。

## <a id="related-pages"></a>相關頁面

- [安裝](Installation)
- [開發環境設定](Development-Setup)
- [變更紀錄](Changelog)

---

[← Wiki 首頁](Home) · [README](https://github.com/hammad417/docxi#readme)
