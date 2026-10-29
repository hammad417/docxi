<!-- Translated from wiki/Changelog.md (git blob 4dabeabaf2ac6e0562ba7ff98a24098001aa1e67). See wiki/TRANSLATION_GUIDE.md. -->
# docxi 變更紀錄與版本歷史

本頁概述每個版本。完整的說明位於儲存庫中的 [`CHANGELOG.md`](https://github.com/hammad417/docxi/blob/master/CHANGELOG.md)，兩者如有出入，以它為準。

## <a id="releases"></a>版本

| 版本 | 日期 | 說明 |
|---|---|---|
| 0.1.1 | 2026-10-29 | gem 的摘要與描述更準確，並新增文件連結。程式碼沒有變更 |
| 0.1.0 | 2026-10-28 | 相容於目前的 Ruby、Rails 和 rubyzip。產生的文件與 0.0.6 相同 |
| 0.1.0.rc1 | 2026-10-28 | 0.1.0 的預先發行版本 |
| 0.0.6 | 2016-10-18 | 0.1.0 更新之前的最後一個版本 |
| 0.0.2 | 2016-10-14 | 沒有留下發行說明 |
| 0.0.1 | 2016-10-14 | 第一個版本。沒有留下發行說明 |

0.0.3 到 0.0.5 從未發布到 RubyGems。

## <a id="011"></a>0.1.1

程式碼沒有變更，產生的文件與 0.1.0 相同。

- gem 的摘要與描述現在說明了 docxi 的實際用途。之前它們把 docxi 稱為範本產生器。
- RubyGems 把本 wiki 連結為 gem 的文件。

## <a id="010"></a>0.1.0

**相容性**

- 需要 Ruby 2.5 以上版本。使用較舊 Ruby 的應用程式會停留在 0.0.6。
- 有沒有 Rails 都能使用，並與 Rails 5.2 到 8.1 一起測試過。
- nokogiri 的需求從 `>= 1.6.3, < 1.7` 改為 `>= 1.8, < 2`，前者與 Rails 6.0 以來的每個版本都衝突。
- rubyzip 的需求從 `~> 0.9.9` 改為 `>= 1.3, < 4`。
- 不再呼叫 ActiveSupport 的 `blank?` 和 `present?`，這兩個方法只有在 Rails 載入它們之後才能使用。

**修正**

- 使用 rubyzip 3 寫出的文件不再包含 Zip64 記錄，Google Docs 和較舊的 LibreOffice 版本會拒絕這種記錄。
- 在 Ruby 3 上，以圖片 URL 呼叫 `add_media` 恢復正常。
- `add_media` 不再將以 `|` 開頭的字串當作 shell 命令執行。
- `tab` 和 `frame` 不再引發 `NameError`。
- `document.frame("text")` 會建立一個包含該文字的文字方塊。
- 頁首中的 `page_numbers` 不再引發 `NameError`。

**gem 套件**

- 新的維護者，以及新的專案首頁和變更紀錄連結。
- 發布版本需要 MFA。

## <a id="related-pages"></a>相關頁面

- [路線圖](Roadmap)
- [安裝](Installation#upgrading-from-00x)

---

[← Wiki 首頁](Home) · [README](https://github.com/hammad417/docxi#readme)
