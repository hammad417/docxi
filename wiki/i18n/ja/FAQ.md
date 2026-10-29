<!-- Translated from wiki/FAQ.md (git blob e799a79ef71c9461e89ce6c6912798eb8c7b99c4). See wiki/TRANSLATION_GUIDE.md. -->
# docxi のよくある質問：Ruby と Rails での Word ドキュメント

よく寄せられる質問への短い回答です。各回答から詳しい説明へのリンクをたどれます。

**Microsoft Word や LibreOffice のインストールは必要ですか？**
いいえ。docxi は XML と zip アーカイブを自分で書き出します。

**docxi には Rails が必要ですか？**
いいえ。どの Ruby プログラムでも動きます。テストは Rails 5.2 から 8.1 を読み込んだ状態でも実行しています。[連携](Integrations)を参照してください。

**Ruby on Rails で Word ドキュメントを生成するには？**
コントローラーかバックグラウンドジョブでドキュメントを組み立て、`word.render.string` を `send_data` で送信します。[連携](Integrations)を参照してください。

**どの Ruby と Rails のバージョンに対応していますか？**
Ruby 2.5 以降です。Rails は任意で、CI では Rails 5.2 から 8.1 をテストしています。[インストール](Installation)を参照してください。

**docxi で既存の .docx を編集したり、Word テンプレートに値を埋め込んだりできますか？**
いいえ、新しいドキュメントを作るだけです。差し込みフィールドのあるテンプレートには、sablon などのテンプレートベースの gem を使ってください。

**目次は見出しから自動で作られますか？**
いいえ。各項目とそのページ番号は自分で追加します。[機能](Features#table-of-contents)を参照してください。

**番号付きリストや見出しスタイルは使えますか？**
0.1.0 では使えません。段落を箇条書きの項目（`ul: true`）にすることはでき、見出しは `size:` や `bold:` などのテキストのオプションで書式を整えられます。

**ページのサイズや向きを変えられますか？**
まだできません。ページは縦向きの US Letter で、変更できるのは余白だけです。[設定](Configuration)を参照してください。

**どの画像形式が使えますか？**
PNG です。すべての画像は PNG パーツとして保存されます。

**URL から画像を追加できますか？**
はい。`add_media` は `http`、`https`、`ftp` の URL を受け付け、`render` を呼び出したときにダウンロードします。ユーザーから受け取った URL は、確認せずに渡さないでください。[セキュリティ](Security)を参照してください。

**ドキュメントを直接ファイルに書き出せますか？**
`render` は `StringIO` を返します。その `.string` を `File.binwrite("out.docx", word.render.string)` で書き出してください。

**docxi で非常に大きなドキュメントをストリーミング出力できますか？**
いいえ。`render` を呼び出すと、ドキュメント全体がメモリ上で組み立てられます。[デプロイ](Deployment)を参照してください。

**右から左に書く段落を設定できますか？**
0.1.0 には右から左に書くためのオプションはありません。テキストは渡したとおりに書き出されます。

**既知の問題はどこにまとまっていますか？**
[ロードマップ](Roadmap)と、README の「Limitations and known issues」（制限事項と既知の問題）セクションにあります。

## <a id="related-pages"></a>関連ページ

- [トラブルシューティング](Troubleshooting)
- [機能](Features)
- [はじめての docxi](Getting-Started)

---

[← Wiki ホーム](Home) · [README](https://github.com/hammad417/docxi#readme)
