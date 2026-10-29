<!-- Translated from wiki/Getting-Started.md (git blob fa1a6a4130a2bf1dbae73e806e7c6d261123f8d7). See wiki/TRANSLATION_GUIDE.md. -->
# はじめての docxi：Ruby で最初の Word ドキュメントを作る

このガイドでは、空のフォルダーから書式付きの `.docx` ファイルを作るまでを説明します。Ruby の基本を知っていることを前提としています。Rails や Microsoft Word は必要ありません。

## <a id="1-install-the-gem"></a>1. gem をインストールする

```sh
gem install docxi
```

Gemfile のあるプロジェクトでは、代わりに `gem "docxi", "~> 0.1"` を追加して `bundle install` を実行します。詳しくは[インストール](Installation)を参照してください。

## <a id="2-create-a-document"></a>2. ドキュメントを作成する

次の内容を `hello.rb` として保存します。

```ruby
require "docxi"

word = Docxi::Document.new
document = word.document

document.p(align: "center") do |p|
  p.text "Hello, Word", size: 24, bold: true
end

File.binwrite("hello.docx", word.render.string)
```

`ruby hello.rb`（Bundler を使うプロジェクトでは `bundle exec ruby hello.rb`）を実行し、`hello.docx` を開きます。

## <a id="3-what-each-line-does"></a>3. 各行の役割

- `Docxi::Document.new` は、デフォルトの余白を持つ空のドキュメントを作成します。
- `word.document` はドキュメントの本文です。追加するものはすべてここを通ります。
- `document.p` は段落を追加します。`align:` など `p` に渡したオプションは、段落とその中のテキストに適用されます。
- `p.text` はテキストのラン（run）を追加します。`size:` の単位はポイントです。
- `word.render` は `.docx` パッケージを組み立てて `StringIO` を返します。その `.string` にファイルのバイト列が入っています。

## <a id="4-add-a-list-and-a-table"></a>4. リストと表を追加する

```ruby
require "docxi"

word = Docxi::Document.new
document = word.document

document.p { |p| p.text "Quarterly summary", size: 16, bold: true }

document.p do |p|
  p.text "Revenue grew ", space: true
  p.text "12%", bold: true, color: "007A33"
  p.text " compared with last quarter.", space: true
end

document.p(ul: true) { |p| p.text "New customers: 48" }
document.p(ul: true) { |p| p.text "Customers lost: 3" }

document.table(columns_width: [200, 100]) do |table|
  table.tr(fill: "D9D9D9", bold: true) do |tr|
    tr.tc { |tc| tc.text "Region" }
    tr.tc { |tc| tc.text "Sales" }
  end
  table.tr do |tr|
    tr.tc { |tc| tc.text "North" }
    tr.tc { |tc| tc.text "$41,000", align: "right" }
  end
end

File.binwrite("summary.docx", word.render.string)
```

注意点が 2 つあります。

- `space: true` は、ランの先頭や末尾の空白を保持します。指定しないと Word が空白を削除し、単語がつながってしまいます。
- `tr` のオプションはその行のすべてのセルに適用されるため、見出し行全体が網かけと太字になります。

## <a id="5-where-to-go-next"></a>5. 次に読むページ

- [使用例](Usage-Examples)：レポート、請求書、手紙の完全なレシピ
- [機能](Features)：段落、テキスト、表、画像、ヘッダーとフッターのすべてのオプション
- [連携](Integrations)：Rails のコントローラーからドキュメントを送信する
- [ベストプラクティス](Best-Practices)：よくある落とし穴を事前に避ける

---

[← Wiki ホーム](Home) · [README](https://github.com/hammad417/docxi#readme)
