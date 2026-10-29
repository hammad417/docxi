<!-- Translated from wiki/Configuration.md (git blob f4889ce51830df74593cc0803cf9ec6d5ccc4402). See wiki/TRANSLATION_GUIDE.md. -->
# docxi の設定：余白、ページ設定、単位

docxi には設定ファイルも環境変数もグローバル設定もありません。すべてドキュメントごと、または要素ごとに指定します。このページでは、ドキュメント単位の設定と、オプションで使う単位を説明します。各要素のオプションは[機能](Features)にあります。

## <a id="page-margins"></a>余白

余白はドキュメントの作成時に渡します。

```ruby
require "docxi"

# 値はすべて 1/20 ポイント単位です（1440 = 1 インチ）。
word = Docxi::Document.new(top: 1440, right: 1080, bottom: 1440, left: 1080, header: 720, footer: 720)
word.document.p { |p| p.text "One-inch top and bottom margins" }

File.binwrite("margins.docx", word.render.string)
```

| オプション | デフォルト | 意味 |
|---|---|---|
| `top` | `720`（0.5 インチ） | 上余白 |
| `right` | `1440`（1 インチ） | 右余白 |
| `bottom` | `720`（0.5 インチ） | 下余白 |
| `left` | `1440`（1 インチ） | 左余白 |
| `header` | `190` | ページの上端からヘッダーまでの距離 |
| `footer` | `200` | ページの下端からフッターまでの距離 |

渡したハッシュはデフォルト値を丸ごと置き換えるため、`top`、`right`、`bottom`、`left` はまとめて指定してください。この 4 つのうち省略したものは空の値として書き出されます。`header` と `footer` は、指定がないか空の場合に 190 と 200 になります。

## <a id="page-size-and-orientation"></a>ページサイズと向き

ページは縦向きの US Letter で、1/20 ポイント単位で 12240 × 15840、つまり 8.5 × 11 インチです。0.1.0 ではサイズと向きを変更できません。

## <a id="first-page"></a>1 ページ目

docxi は Word の「先頭ページのみ別指定」を常に有効にします。`add_header` と `add_footer` で追加したヘッダーとフッターは 2 ページ目以降に使われます。1 ページ目には `add_first_header` と `add_first_footer` が使われるので、1 ページ目にヘッダーやフッターが必要な場合はこちらも設定してください。[機能](Features#headers-and-footers)を参照してください。

## <a id="what-every-document-contains"></a>すべてのドキュメントに含まれるもの

- デフォルトのフォントは Arial 12 pt です。
- 定義されているスタイル：`Normal`、`DefaultParagraphFont`、`TableNormal`、`TableCustom`、`NoList`。
- 箇条書きの項目（`ul: true`）は行頭文字にダッシュを使います。
- ドキュメントのプロパティには、アプリケーション名として「Docxi Ruby Gem」と、ドキュメントがレンダリングされた時刻が記録されます。

## <a id="units"></a>単位

| オプション | 単位 |
|---|---|
| テキストの `size:` | ポイント（docxi が 2 倍にして Word の半ポイント単位に変換します） |
| 余白、段落の `left:` と `right:`、表の `iwidth:`、行の `height:` | 1/20 ポイント（1 インチあたり 1440） |
| セルの `borders:` の太さ | 1/8 ポイント |
| `columns_width:` とセルの `width:` | 14.92 を掛けて 1/20 ポイントに変換されます。96 dpi のピクセルに近い値です |
| インライン画像の `style:` | `"120pt"` などの CSS の長さ |
| 配置指定した画像の `pH:` と `pV:` | EMU（1 インチあたり 914400） |
| テキストと塗りつぶしの色 | `#` なしの 16 進数（例：`"8B0000"`） |
| 画像の `fill_color:` | `#` 付きの VML の色（例：`"#00FF00"`） |

## <a id="related-pages"></a>関連ページ

- [機能](Features)
- [API リファレンス](API-Reference)
- [ベストプラクティス](Best-Practices)

---

[← Wiki ホーム](Home) · [README](https://github.com/hammad417/docxi#readme)
