<!-- Translated from wiki/Features.md (git blob 3eb1a7d53ab3801f4d8e200cedf6c2a9e8fef2be). See wiki/TRANSLATION_GUIDE.md. -->
# docxi の機能：段落、テキスト、表、画像、ヘッダーとフッター

このページでは、docxi が Word ドキュメントに追加できるすべての要素と、それぞれが受け付けるオプションを紹介します。例はすべて次のセットアップを前提にしています。

```ruby
require "docxi"

word = Docxi::Document.new
document = word.document
```

不要なオプションは省略できます。色は `#` なしの 16 進数で指定します（例：`"8B0000"`）。単位とページ設定については[設定](Configuration)を参照してください。

## <a id="paragraphs"></a>段落

`document.p(options) { |p| ... }` は段落を追加します。

| オプション | 値 | 効果 |
|---|---|---|
| `align` | `"left"`、`"center"`、`"right"`、`"both"` | 配置。`"both"` は両端揃えです |
| `fill` | 16 進数の色 | 段落とそのテキストの背景の網かけ |
| `left`、`right` | 1/20 ポイント | インデント。片方だけでは無視されるため、両方を指定します |
| `ul` | `true` | 段落を箇条書きの項目にします |
| `bottom` | `true` | 段落の下に罫線を引きます |

`size:`、`bold:`、`color:` などのテキストのオプションも `p` に渡せます。渡した値は、その段落内のすべてのテキストランのデフォルトになります。

ブロックの中では `text`、`br`、`tab`、`image`、`hyperlink`、`frame` を呼び出せます。`document.text(string, options)` は、ランが 1 つだけの段落を作るショートカットです。

```ruby
document.p(align: "both", left: 720, right: 720) do |p|
  p.text "An indented, justified paragraph."
end

document.p(ul: true) { |p| p.text "A bullet item" }
document.text "A one-line paragraph", size: 14, bold: true
```

## <a id="text"></a>テキスト

`p.text(string, options)` はテキストのラン（run）を追加します。

| オプション | 値 | 効果 |
|---|---|---|
| `bold` | `true` | 太字 |
| `italic` | `true` | 斜体 |
| `underline` | `true` | 一重下線 |
| `color` | 16 進数の色 | 文字の色 |
| `size` | ポイント | フォントサイズ |
| `font` | フォント名 | フォント（例：`"Georgia"`） |
| `fill` | 16 進数の色 | テキストの背景の網かけ |
| `space` | `true` | テキストの先頭と末尾の空白を保持します |
| `br` | `true` | テキストの後に改行を追加します |

`nil` の文字列はスキップされます。

```ruby
document.p do |p|
  p.text "Total: ", bold: true, space: true
  p.text "$1,200", color: "007A33", font: "Georgia", size: 14
end
```

## <a id="line-breaks-page-breaks-and-tabs"></a>改行、改ページ、タブ

- `p.br` または `p.br(times: 3)` は、段落内に改行を追加します。
- `document.br(page: true)` は新しいページを始めます。
- `p.tab` または `p.tab(times: 2)` は、段落内にタブ文字を追加します。
- `document.tab(page: true)` は、タブを 1 つだけ含む段落を追加します。

`document.br` と `document.tab` を `page: true` なしで呼び出さないでください。テキストランが本文に直接置かれ、有効な WordprocessingML になりません。

```ruby
document.p do |p|
  p.text "Name:", br: true
  p.tab
  p.text "Signature line"
  p.br(times: 2)
end

document.br(page: true)
```

## <a id="tables"></a>表

`document.table(options)` はブロックを受け取り、その中で `table.tr(options)` が行を、`tr.tc(options)` がセルを追加します。

**表のオプション**

| オプション | 値 | 効果 |
|---|---|---|
| `columns_width` | 数値の配列 | 列の幅。14.92 を掛けて 1/20 ポイントに変換されます。セルのデフォルトの幅にもなります |
| `height` | 1/20 ポイント | 各行の最小の高さ |
| `iwidth` | 1/20 ポイント | 左余白からの表のインデント |
| `style` | スタイル ID | 表のスタイル。docxi は `TableNormal` と `TableCustom` を定義しています |
| `width` | 数値 | 表の優先幅として、種類 `auto` で書き出されます。`columns_width` がない場合は、各セルにもこの幅に 14.92 を掛けた値が使われます |
| `align` | `"left"`、`"center"`、`"right"` | 表の配置。`tc.text` で追加したテキストも同じように配置されます |
| `padding` | `true` | セル内に固定のパディングを追加します |
| `borders` | ハッシュ | 既知の問題：表の外枠ではなく各セルに適用されます |

**行のオプション**：`height:` は行の最小の高さを設定します。`tr` に渡したセルのオプションは、その行のすべてのセルに適用されます。表のオプションも同じように、すべての行とセルに引き継がれます。

**セルのオプション**

| オプション | 値 | 効果 |
|---|---|---|
| `borders` | `{ top:, right:, bottom:, left: }` | 罫線の太さ（1/8 ポイント単位）。`nil` はその辺に罫線がないことを意味します |
| `fill` | 16 進数の色 | セルの網かけ。塗りつぶしのないセルは白（`"FFFFFF"`）になります |
| `valign` | `"top"`、`"center"`、`"bottom"` | 縦方向の配置 |
| `colspan` | 整数 | セルがまたぐ列の数 |
| `rowspan` | `true` | 縦方向の結合を開始します |
| `merged` | `true` | 上のセルからの縦方向の結合を続けます。このセルの内容は無視されます |
| `width` | 数値 | セルの幅。14.92 を掛けた値になります |

**セルの内容**

- `tc.text(string, options)`：`bold:` や `align:` など、セルと行のオプションがテキストにも適用されます。
- `tc.p(options) { |p| ... }`：自分のオプションだけを使う段落。
- `tc.image(media, width:, height:)`：画像。サイズの問題については[画像](#images)を参照してください。
- `tc.page_numbers`：ページ番号フィールド。既知の問題：先頭に「GlobalOptions」というテキストが付きます。

```ruby
document.table(columns_width: [150, 150]) do |table|
  table.tr do |tr|
    tr.tc(rowspan: true, valign: "center", borders: { top: 4, right: 4, bottom: 4, left: 4 }) do |tc|
      tc.text "Spans two rows"
    end
    tr.tc { |tc| tc.text "Row 1" }
  end
  table.tr do |tr|
    tr.tc(merged: true)
    tr.tc { |tc| tc.text "Row 2" }
  end
end
```

## <a id="images"></a>画像

`media = document.add_media(source)` は画像を登録します。`source` には次のものを指定できます。

- `File` や `StringIO` などの IO オブジェクト
- ファイルパス
- `http`、`https`、`ftp` の URL

画像は `render` を呼び出したときに読み込まれ、常に PNG パーツとして保存されます。そのため、PNG ファイルを使うのが安全です。

**インライン画像**：`p.image(media, options)` は、画像をテキストの行内に配置します。

| オプション | 値 | 効果 |
|---|---|---|
| `style` | CSS プロパティのハッシュ | サイズ（例：`{ width: "120pt", height: "60pt" }`） |
| `title` | 文字列 | 画像のタイトル |
| `fill_color` | `"#00FF00"` などの VML の色 | 画像の背景の塗りつぶし |
| `opacity` | `"50%"` などの文字列 | 塗りつぶしの不透明度 |
| `wrap` | `true` | 画像の周りの文字列の折り返しをオフにします |

```ruby
logo = document.add_media(File.open("logo.png", "rb"))
document.p(align: "center") { |p| p.image logo, style: { width: "120pt", height: "60pt" } }
```

**配置を指定した画像**：`p.image(media, pH:, pV:, wrap: true, width:, height:)` は、列（`pH`）と段落（`pV`）からのオフセットを EMU で指定して画像を固定します。`pH`、`pV`、`wrap` はすべて指定する必要があり、そうでない場合はインラインで配置されます。

**ヘッダーの画像**：`h.image(media, align:, width:, height:, pH:, pV:, behind:, wrap:)`。`wrap: "none"` は折り返しをオフにします。それ以外の値では、テキストが画像の周りを四角形に回り込みます。

**0.1.0 の既知の問題**：配置を指定した画像、ヘッダーの画像、表のセル内の画像は、幅と高さをどちらも `width * height * 14.92` EMU として計算します。そのため常に正方形になり、たいていはとても小さくなります。`style:` でサイズを指定したインライン画像は影響を受けません。回避策は[ベストプラクティス](Best-Practices#images)を参照してください。

## <a id="headers-and-footers"></a>ヘッダーとフッター

```ruby
header = Docxi::Word::Headers::Header.new do |h|
  h.text "Quarterly report", size: 9, align: "right"
end

footer = Docxi::Word::Footers::Footer.new do |f|
  f.page_numbers name: "Page", align: "center"
end

document.add_header(header)
document.add_footer(footer)
document.add_first_header(Docxi::Word::Headers::Header.new { |h| h.text "Cover page" })
```

- ヘッダーには `text`、`br`、`table`、`image`、`page_numbers` を入れられます。フッターには `text`、`br`、`table`、`page_numbers` を入れられます。
- `Header.new` や `Footer.new` に渡したオプションは、その中のテキストに適用されます。
- `add_header` と `add_footer` は 2 ページ目以降に適用されます。1 ページ目には常に `add_first_header` と `add_first_footer` が使われます。[設定](Configuration#first-page)を参照してください。

## <a id="page-numbers"></a>ページ番号

`page_numbers(name: "Page", align: "center")` は、`name` のテキストに続けて PAGE フィールドを書き出します。番号は Arial 10 pt、濃いグレーで表示されます。`align` のデフォルトは `"right"` です。

既知の問題：フィールドコードで `\* MERGEFORMAT` のバックスラッシュが欠けています。

## <a id="table-of-contents"></a>目次

```ruby
document.table_of_content do |toc|
  toc.add_row "Introduction", 1
  toc.add_row "Appendix", 7, leader: "hyphen", pos: 8000
  toc.add_row "Notes"
end
```

- `add_row(text, page = nil, leader: "dot", pos: 9360)`。
- ページ番号を指定すると、その項目には `pos`（1/20 ポイント単位。9360 は 6.5 インチ）に右揃えのタブ位置が設定され、`leader` の文字で埋められます。リーダーの値は WordprocessingML のもので、`"dot"`、`"hyphen"`、`"underscore"`、`"middleDot"`、`"none"` などがあります。
- ページ番号を指定しない場合、その項目はタブの後にテキストが続く形になります。
- ページ番号は渡した値がそのまま使われます。docxi は Word の TOC フィールドを作らず、ページの計算もしません。

## <a id="hyperlinks"></a>ハイパーリンク

```ruby
url = "https://github.com/hammad417/docxi"
link = document.add_hyperlink(url)
document.p { |p| p.hyperlink "docxi on GitHub", url, link }
```

`add_hyperlink(url)` はリレーションシップ ID を返し、`hyperlink(text, url, id)` がリンクを配置します。リンク先はリレーションシップから決まるため、目的の URL に対して返された ID を渡してください。外部リンクのみ対応しています。

## <a id="frames-text-boxes"></a>フレーム（テキストボックス）

```ruby
document.p do |p|
  p.frame(style: "width:200pt;height:40pt") do |f|
    f.text "Text inside a box"
  end
end

document.frame "A box with one line", style: "width:200pt;height:30pt"
```

- フレームは VML のテキストボックスで、薄いグレーの半透明の塗りつぶしと黒い枠線が固定で付きます。テキストはフレームの上下に配置されます。
- フレームの中では `text`、`br`、`tab`、`image`、`frame` を呼び出せます。
- `spacing: true` はテキストの前に間隔を追加します。

## <a id="output"></a>出力

`word.render` は、完全な `.docx` パッケージを含む `StringIO` を返します。アーカイブは Zip64 レコードなしで書き出されるため、Google Docs や古いバージョンの LibreOffice でも開けます。パッケージ内のパーツの一覧は[アーキテクチャ](Architecture)にあります。

## <a id="related-pages"></a>関連ページ

- [使用例](Usage-Examples)
- [API リファレンス](API-Reference)
- [ロードマップ](Roadmap)：0.2.0 で対応予定の既知の問題

---

[← Wiki ホーム](Home) · [README](https://github.com/hammad417/docxi#readme)
