<!-- Translated from wiki/Home.md (git blob 4724ddd12e6c213e92a003c3fc882558d8e2edfc). See wiki/TRANSLATION_GUIDE.md. -->
# docxi Wiki：Ruby で Word（.docx）ドキュメントを生成する

docxi は、コードから Microsoft Word の `.docx` ファイルを生成する Ruby の gem です。段落、書式付きテキスト、表、画像、ヘッダーとフッター、ページ番号、ハイパーリンク、シンプルな目次を作成できます。Microsoft Word や LibreOffice、テンプレートは不要で、素の Ruby でも Ruby on Rails アプリケーションでも動作します。

[README](https://github.com/hammad417/docxi#readme) は短い紹介です。この Wiki では、すべてのオプションの説明、そのままコピーできるレシピ、docxi の内部構造の解説、問題が起きたときの対処法をまとめています。

現在のリリース：**0.1.1**（Ruby 2.5 以降に対応）。[変更履歴](Changelog)を参照してください。

## <a id="start-here"></a>はじめに

**docxi を初めて使う方へ**

- [はじめての docxi](Getting-Started)：gem をインストールして最初のドキュメントを作成する
- [インストール](Installation)：必要な環境、Bundler、0.0.x からのアップグレード
- [使用例](Usage-Examples)：データから作るレポート、請求書、手紙、表。そのままコピーして使えます

**ドキュメントを作る**

- [機能](Features)：各要素と、指定できるオプション
- [設定](Configuration)：余白、ページ設定、単位
- [API リファレンス](API-Reference)：クラスとメソッドの一覧
- [ベストプラクティス](Best-Practices)：よくある落とし穴を避けるための習慣

**問題が起きたとき**

- [トラブルシューティング](Troubleshooting)：エラーメッセージと、正しく表示されないドキュメント
- [よくある質問](FAQ)：よくある質問への短い回答
- [セキュリティ](Security)：問題の報告方法と、信頼できない入力の扱い方

**さらに詳しく**

- [アーキテクチャ](Architecture)：Ruby の呼び出しが `.docx` パッケージになるまで
- [連携](Integrations)：Rails のコントローラー、Rack アプリ、バックグラウンドジョブ
- [デプロイ](Deployment)：本番環境での docxi の運用と、リリースの公開方法
- [コントリビュート](Contributing)と[開発環境のセットアップ](Development-Setup)
- [ロードマップ](Roadmap)、[変更履歴](Changelog)、[クレジット](Credits)

## <a id="what-docxi-does-and-what-it-doesnt"></a>docxi にできること、できないこと

docxi は新しいドキュメントを作成します。既存の `.docx` を開いたり編集したりすること、Word テンプレートに値を埋め込むこと、差し込み印刷を行うことはできません。番号付きリスト、見出しスタイル、自動生成の目次はなく、ページは常に縦向きの US Letter です。0.2.0 で対応予定の既知の問題は[ロードマップ](Roadmap)にまとめています。

## <a id="credits"></a>クレジット

docxi は [Phaedra Solutions](https://www.phaedrasolutions.com/) の [Hammad Maqbool](https://github.com/hammad417) がメンテナンスしています。2016 年に Timur Gabdrakipov の [docxer](https://github.com/gtimti/docxer) をもとに始まりました。詳しくは[クレジット](Credits)を参照してください。

## <a id="languages"></a>言語

この Wiki は英語版が基準です。各言語の翻訳へのリンクは、サイドバーと各ページの先頭にあります。

---

[README](https://github.com/hammad417/docxi#readme) · [RubyGems](https://rubygems.org/gems/docxi) · [問題を報告する](https://github.com/hammad417/docxi/issues)
