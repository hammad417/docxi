<!-- Translated from wiki/Deployment.md (git blob 9058153b973e13980c113b2d7be11debb7bfff40). See wiki/TRANSLATION_GUIDE.md. -->
# 本番環境での docxi の運用とリリースの公開

docxi はサービスではなくライブラリなので、それ単体でデプロイするものはありません。このページでは、docxi を使うアプリケーションをリリースするときに気をつけることと、docxi の新しいバージョンの公開方法を説明します。

## <a id="shipping-an-app-that-uses-docxi"></a>docxi を使うアプリケーションのリリース

- **依存関係**：`bundle install` で nokogiri と rubyzip がインストールされます。nokogiri は一般的な Linux、macOS、Windows 向けにプリコンパイル済みの gem を提供しています。それ以外のプラットフォーム（一部の最小構成のコンテナイメージなど）ではソースからコンパイルされ、C コンパイラが必要です。[nokogiri のインストールガイド](https://nokogiri.org/tutorials/installing_nokogiri.html)を参照してください。
- **オフィスソフトは不要**：サーバーやコンテナに Microsoft Word や LibreOffice は必要ありません。
- **メモリ**：`render` はドキュメント全体をメモリ上で組み立て、`StringIO` として返します。メモリ使用量はページ数と画像のサイズに応じて増えます。
- **ネットワークとファイル**：docxi がファイルシステムを読むのは `add_media` に渡した画像パスのためだけで、ネットワークを使うのは渡した画像 URL のためだけです。[セキュリティ](Security)を参照してください。
- **長いドキュメント**：Web リクエストの処理中ではなく、バックグラウンドジョブで生成してください。[連携](Integrations)を参照してください。

## <a id="how-docxi-releases-are-published"></a>docxi のリリースの公開方法

以下はメンテナー向けの手順です。リリースは GitHub Actions から Trusted Publishing で RubyGems に公開されるため、リポジトリには API キーを保存していません。

1. `lib/docxi/version.rb` と `CHANGELOG.md` を更新し、`master` にマージします。
2. コミットに `v` を付けたバージョン（例：`v0.1.1`）のタグを付け、タグをプッシュします。
3. `.github/workflows/release.yml` が gem をビルドして RubyGems にプッシュします。

プレリリースも同じ手順で、`0.2.0.rc1` のようなバージョンを使います。Bundler は、明示的に指定された場合にしかプレリリースをインストールしません。CI（`.github/workflows/ci.yml`）は、`master` へのプッシュとすべてのプルリクエストで実行されます。

## <a id="related-pages"></a>関連ページ

- [インストール](Installation)
- [開発環境のセットアップ](Development-Setup)
- [変更履歴](Changelog)

---

[← Wiki ホーム](Home) · [README](https://github.com/hammad417/docxi#readme)
