<!-- Translated from wiki/Configuration.md (git blob f4889ce51830df74593cc0803cf9ec6d5ccc4402). See wiki/TRANSLATION_GUIDE.md. -->
# 設定 docxi：邊界、版面設定和單位

docxi 沒有設定檔、環境變數或全域設定，所有設定都依文件或依元素指定。本頁介紹文件層級的設定，以及各選項使用的單位。各元素的選項列在[功能](Features)頁面。

## <a id="page-margins"></a>邊界

邊界在建立文件時傳入：

```ruby
require "docxi"

# 所有數值的單位都是二十分之一點（1440 = 1 英吋）。
word = Docxi::Document.new(top: 1440, right: 1080, bottom: 1440, left: 1080, header: 720, footer: 720)
word.document.p { |p| p.text "One-inch top and bottom margins" }

File.binwrite("margins.docx", word.render.string)
```

| 選項 | 預設值 | 意義 |
|---|---|---|
| `top` | `720`（0.5 英吋） | 上邊界 |
| `right` | `1440`（1 英吋） | 右邊界 |
| `bottom` | `720`（0.5 英吋） | 下邊界 |
| `left` | `1440`（1 英吋） | 左邊界 |
| `header` | `190` | 頁面上緣到頁首的距離 |
| `footer` | `200` | 頁面下緣到頁尾的距離 |

你傳入的雜湊會整個取代預設值，所以請同時設定 `top`、`right`、`bottom` 和 `left`。這四項中漏掉的任何一項都會寫成空值。`header` 和 `footer` 缺少或為空白時，會改用 190 和 200。

## <a id="page-size-and-orientation"></a>頁面大小和方向

頁面為直向的 US Letter：12240 × 15840 個二十分之一點，也就是 8.5 × 11 英吋。0.1.0 無法變更頁面大小和方向。

## <a id="first-page"></a>第一頁

docxi 一律開啟 Word 的「首頁不同」設定。以 `add_header` 和 `add_footer` 加入的頁首和頁尾從第 2 頁開始使用。第 1 頁使用 `add_first_header` 和 `add_first_footer`，因此如果第一頁需要頁首或頁尾，也要設定它們。請參閱[功能](Features#headers-and-footers)。

## <a id="what-every-document-contains"></a>每份文件都包含的內容

- 預設字型為 Arial 12 點。
- 定義的樣式：`Normal`、`DefaultParagraphFont`、`TableNormal`、`TableCustom` 和 `NoList`。
- 項目符號清單項目（`ul: true`）使用短橫線作為項目符號。
- 文件屬性將應用程式名稱記錄為「Docxi Ruby Gem」，並記錄文件渲染的時間。

## <a id="units"></a>單位

| 選項 | 單位 |
|---|---|
| 文字 `size:` | 點（docxi 會將其乘以 2，換算成 Word 的半點） |
| 邊界、段落 `left:` 和 `right:`、表格 `iwidth:`、列 `height:` | 二十分之一點（每英吋 1440） |
| 儲存格 `borders:` 的粗細 | 八分之一點 |
| `columns_width:` 和儲存格 `width:` | 乘以 14.92 換算成二十分之一點，在 96 dpi 下接近像素值 |
| 內嵌圖片 `style:` | CSS 長度，例如 `"120pt"` |
| 定位圖片 `pH:` 和 `pV:` | EMU（每英吋 914400） |
| 文字和填滿色彩 | 不含 `#` 的十六進位值，例如 `"8B0000"` |
| 圖片 `fill_color:` | 含 `#` 的 VML 色彩，例如 `"#00FF00"` |

## <a id="related-pages"></a>相關頁面

- [功能](Features)
- [API 參考](API-Reference)
- [最佳做法](Best-Practices)

---

[← Wiki 首頁](Home) · [README](https://github.com/hammad417/docxi#readme)
