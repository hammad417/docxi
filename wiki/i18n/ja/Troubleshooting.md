<!-- Translated from wiki/Troubleshooting.md (git blob 76c38be3731e1d9e4696f2aab1eba96257a98970). See wiki/TRANSLATION_GUIDE.md. -->
# docxi のトラブルシューティング：インストールエラー、実行時エラー、壊れた出力

以下から症状を探してください。古いアプリケーションの問題のほとんどは docxi 0.0.x が原因で、0.1.0 にアップグレードすればコードを変えずに解決します。

## <a id="installation"></a>インストール

**Rails 6 以降のアプリで、Bundler が `nokogiri` の互換バージョンを見つけられない**
docxi 0.0.6 は `nokogiri < 1.7` に固定されていますが、Rails 6.0 以降はすべて nokogiri 1.8.5 以降を必要とします。`gem "docxi", "~> 0.1"` を使ってください。

**Bundler が `rubyzip` の互換バージョンを見つけられない**
docxi 0.0.6 は `rubyzip ~> 0.9.9` に固定されており、より新しい rubyzip を必要とする gem と競合します。docxi 0.1.0 は rubyzip 1.3 から 3.x まで受け付けます。

**Bundler が 0.0.6 をインストールし続ける**
Gemfile で docxi が 0.0.x に固定されている（例：`"~> 0.0.6"`）か、Ruby が 0.1.0 に必要な 2.5 より古いかのどちらかです。Gemfile の行を `gem "docxi", "~> 0.1"` に変更し、`ruby -v` を確認してください。

**nokogiri のコンパイルに失敗する**
お使いのプラットフォームに合うプリコンパイル済みの nokogiri gem がないため、ソースからビルドされています。C コンパイラをインストールし、[nokogiri のインストールガイド](https://nokogiri.org/tutorials/installing_nokogiri.html)に従ってください。

## <a id="errors-while-building-a-document"></a>ドキュメント作成中のエラー

**Rails の外で `NoMethodError: undefined method 'present?'` が出る**
docxi 0.0.6 は ActiveSupport が読み込まれていることを前提にしていました。0.1.0 では不要です。

**`tab`、`frame`、ヘッダーの `page_numbers` を呼び出すと `NameError` が出る**
これらは 0.0.6 では壊れていましたが、0.1.0 では動作します。`document.frame("text")` も同様です。

**画像を追加すると `Errno::ENOENT` が出る**
プログラムを実行しているディレクトリから見て、そのパスが存在しません。Ruby 3 で docxi 0.0.6 を使うと、画像の URL も同じように失敗します。0.0.6 が `Kernel#open` を使っていたためです。0.1.0 は URL を `URI#open` でダウンロードします。

## <a id="documents-that-dont-look-right"></a>正しく表示されないドキュメント

**Google Docs や LibreOffice でファイルが開けない**
rubyzip 3 と組み合わせた docxi 0.0.6 はすべてのエントリーに Zip64 レコードを書き込み、これらのアプリはそれを受け付けません。0.1.0 は Zip64 レコードなしでアーカイブを書き出します。

**1 ページ目にヘッダーやフッターがない**
docxi は 1 ページ目に専用のヘッダーとフッターを設定します。`add_first_header` と `add_first_footer` で設定してください。[設定](Configuration#first-page)を参照してください。

**余白が欠けている、またはページのレイアウトがおかしい**
余白を設定するときは、`top`、`right`、`bottom`、`left` をまとめて渡してください。省略したものは空で書き出されます。

**テキストラン同士の間の空白が消える**
Word はランの先頭や末尾の空白を削除します。その `text` の呼び出しに `space: true` を追加してください。

**ヘッダーや表のセル内、または `pH`/`pV` を使った画像が小さすぎる、または正方形になる**
0.1.0 の既知の問題です。幅と高さの両方が `width * height * 14.92` EMU として計算されます。`style:` でサイズを指定したインライン画像を使うか、[ベストプラクティス](Best-Practices#images)の回避策を参照してください。

**同じドキュメントを 2 回目にレンダリングすると画像が消える**
開いた `File` などの IO から追加した画像は `render` の中で読み込まれるため、2 回目のレンダリングでは空のストリームを読むことになります。代わりにパスを渡すか、レンダリングを 1 回だけにしてください。

**`table` に設定した罫線がすべてのセルの周りに表示される**
既知の問題です。表単位の `borders:` は、表の外枠ではなく各セルに適用されます。外枠だけを描くには、端のセルに外側の辺だけの `borders:` を設定してください。

**表のセル内のページ番号が「GlobalOptions」で始まる**
`TableCell#page_numbers` の既知の問題です。ページ番号はフッターかヘッダーに入れてください。

**JPEG や GIF の画像が正しく表示されない、または表示されない**
すべての画像は、実際の形式にかかわらず、PNG のコンテンツタイプを持つ PNG パーツとして保存されます。そのため、アプリによっては他の形式を表示できない場合があります。追加する前に画像を PNG に変換してください。

## <a id="inspecting-a-generated-file"></a>生成されたファイルを調べる

`.docx` は zip アーカイブなので、中身を見ることができます。

```sh
unzip -l report.docx
unzip -p report.docx word/document.xml
```

それでも問題が解決しない場合は、問題を再現できる短いスクリプトと、Ruby、docxi、nokogiri、rubyzip のバージョンを添えて [Issue を作成](https://github.com/hammad417/docxi/issues)してください。

## <a id="related-pages"></a>関連ページ

- [よくある質問](FAQ)
- [ベストプラクティス](Best-Practices)
- [インストール](Installation)

---

[← Wiki ホーム](Home) · [README](https://github.com/hammad417/docxi#readme)
