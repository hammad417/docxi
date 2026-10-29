<!-- Translated from wiki/Troubleshooting.md (git blob 76c38be3731e1d9e4696f2aab1eba96257a98970). See wiki/TRANSLATION_GUIDE.md. -->
# docxi 疑難排解：安裝錯誤、執行時期錯誤和損壞的輸出

請在下方找到你遇到的症狀。舊版應用程式中的大多數問題都來自 docxi 0.0.x，升級到 0.1.0 即可解決，不需要修改程式碼。

## <a id="installation"></a>安裝

**在 Rails 6+ 應用程式中，Bundler 找不到相容的 `nokogiri` 版本**
docxi 0.0.6 鎖定了 `nokogiri < 1.7`，而 Rails 6.0 之後的每個版本都需要 nokogiri 1.8.5 以上版本。請使用 `gem "docxi", "~> 0.1"`。

**Bundler 找不到相容的 `rubyzip` 版本**
docxi 0.0.6 鎖定了 `rubyzip ~> 0.9.9`，與任何需要較新版 rubyzip 的 gem 衝突。docxi 0.1.0 接受 rubyzip 1.3 到 3.x。

**Bundler 一直安裝 0.0.6**
原因可能是 Gemfile 把 docxi 固定在 0.0.x（例如 `"~> 0.0.6"`），也可能是 Ruby 低於 0.1.0 所需的 2.5。請把 Gemfile 中的這一行改為 `gem "docxi", "~> 0.1"`，並檢查 `ruby -v`。

**nokogiri 編譯失敗**
沒有與你的平台相符的預先編譯 nokogiri gem，所以它正在從原始碼建置。請安裝 C 編譯器，並依照 [nokogiri 安裝指南](https://nokogiri.org/tutorials/installing_nokogiri.html)操作。

## <a id="errors-while-building-a-document"></a>建立文件時發生錯誤

**在 Rails 之外出現 `NoMethodError: undefined method 'present?'`**
docxi 0.0.6 假設 ActiveSupport 已經載入。0.1.0 不需要它。

**呼叫 `tab`、`frame` 或頁首的 `page_numbers` 時出現 `NameError`**
這些功能在 0.0.6 中無法使用，在 0.1.0 中可以正常運作，包括 `document.frame("text")`。

**加入圖片時出現 `Errno::ENOENT`**
該路徑不存在（相對於程式執行時所在的目錄）。在 Ruby 3 上使用 docxi 0.0.6 時，圖片 URL 也會以同樣的方式失敗，因為 0.0.6 使用的是 `Kernel#open`。0.1.0 以 `URI#open` 下載 URL。

## <a id="documents-that-dont-look-right"></a>文件顯示不正常

**Google Docs 或 LibreOffice 無法開啟檔案**
docxi 0.0.6 搭配 rubyzip 3 時，會在每個項目中寫入 Zip64 記錄，而這些程式會拒絕這種記錄。0.1.0 寫出的壓縮檔不含這些記錄。

**第一頁沒有頁首或頁尾**
docxi 會為第一頁另外設定頁首和頁尾。請用 `add_first_header` 和 `add_first_footer` 設定它們。請參閱[設定](Configuration#first-page)。

**缺少某個邊界，或版面配置看起來不對**
設定邊界時，請同時傳入 `top`、`right`、`bottom` 和 `left`。漏掉的那些會寫成空值。

**文字片段之間的空格消失了**
Word 會捨棄文字片段開頭或結尾的空格。請在該 `text` 呼叫中加入 `space: true`。

**頁首中、表格儲存格中或使用 `pH`/`pV` 的圖片很小或是正方形**
0.1.0 的已知問題：它們的兩個尺寸都以 `width * height * 14.92` EMU 計算。請使用以 `style:` 設定尺寸的內嵌圖片，或參閱[最佳做法](Best-Practices#images)中的因應措施。

**同一份文件第二次渲染時缺少圖片**
從已開啟的 `File` 或其他 IO 加入的圖片會在 `render` 期間讀取，第二次渲染讀到的是空的串流。請改為傳入路徑，或只渲染一次。

**在 `table` 上設定的框線出現在每個儲存格周圍**
已知問題：表格層級的 `borders:` 會套用到每個儲存格，而不是表格的外框。如果只想繪製外框，請只在邊緣儲存格上為其外側設定 `borders:`。

**表格儲存格中的頁碼以「GlobalOptions」開頭**
這是 `TableCell#page_numbers` 的已知問題。請改為將頁碼放在頁尾或頁首中。

**JPEG 或 GIF 圖片顯示不正確或沒有出現**
無論實際格式為何，每張圖片都以 PNG 組件和 PNG 內容類型儲存，因此某些程式可能無法顯示其他格式。請在加入前將圖片轉換為 PNG。

## <a id="inspecting-a-generated-file"></a>檢查產生的檔案

`.docx` 是 zip 壓縮檔，所以可以查看其內部：

```sh
unzip -l report.docx
unzip -p report.docx word/document.xml
```

如果問題仍然存在，請[提交 issue](https://github.com/hammad417/docxi/issues)，附上能重現問題的簡短指令碼，以及 Ruby、docxi、nokogiri 和 rubyzip 的版本。

## <a id="related-pages"></a>相關頁面

- [常見問題](FAQ)
- [最佳做法](Best-Practices)
- [安裝](Installation)

---

[← Wiki 首頁](Home) · [README](https://github.com/hammad417/docxi#readme)
