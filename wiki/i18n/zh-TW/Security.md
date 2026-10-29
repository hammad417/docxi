<!-- Translated from wiki/Security.md (git blob 2358bb39bd6772aaea144aa8ca2434b43c086728). See wiki/TRANSLATION_GUIDE.md. -->
# docxi 安全性：回報問題與處理不受信任的輸入

本頁說明如何回報安全性問題，以及當你的應用程式將使用者輸入傳給 docxi 時需要了解的安全性重點。

## <a id="reporting-a-vulnerability"></a>回報漏洞

請不要在公開的 issue 中回報安全性問題，而是以電子郵件聯絡維護者。電子郵件地址在 [`docxi.gemspec`](https://github.com/hammad417/docxi/blob/master/docxi.gemspec) 中。請註明 docxi 的版本，並提供重現問題的簡短方法。

## <a id="upgrade-from-00x"></a>從 0.0.x 升級

docxi 0.0.x 會將圖片路徑字串傳給 `Kernel#open`。`Kernel#open` 會把以 `|` 開頭的字串當作 shell 命令執行，因此能控制圖片路徑的攻擊者可以在你的伺服器上執行命令。0.1.0 以 `File.open` 開啟檔案，以 `URI#open` 下載 URL，兩者都不會執行命令。

## <a id="untrusted-input"></a>不受信任的輸入

**傳入文字是安全的**。文字、表格內容和其他字串會以 XML 文字或屬性值寫入，並由 nokogiri 跳脫。使用者輸入無法在文件中加入元素。

**圖片來源需要謹慎處理**。`add_media` 接受：

| 來源 | docxi 的處理方式 |
|---|---|
| 你開啟的 IO | 讀取它 |
| 檔案路徑 | 從伺服器磁碟讀取該檔案 |
| `http`、`https` 或 `ftp` URL | 從你的伺服器下載它 |

如果使用者可以指定路徑，他們就能讓伺服器上任何可讀取的檔案（例如設定或金鑰）被嵌入文件中。如果使用者可以指定 URL，你的伺服器就會去請求它，而這可能觸及內部位址（伺服器端請求偽造，SSRF）。請傳入你自己開啟的 IO，或只允許已知的路徑和主機。

**超連結**會指向 URL 所指的任何地方。將使用者提供的 URL 傳給 `add_hyperlink` 之前，請先檢查。

## <a id="how-releases-are-protected"></a>版本發布如何受到保護

- 版本透過 GitHub Actions 以 RubyGems 可信任發布（trusted publishing）的方式發布，因此儲存庫中不會存放任何 API 金鑰。
- gem 中繼資料設定了 `rubygems_mfa_required`，因此使用個人 API 金鑰推送時需要多重要素驗證。

## <a id="related-pages"></a>相關頁面

- [最佳做法](Best-Practices)
- [部署](Deployment)
- [參與貢獻](Contributing)

---

[← Wiki 首頁](Home) · [README](https://github.com/hammad417/docxi#readme)
