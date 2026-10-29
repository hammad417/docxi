<!-- Translated from wiki/Contributing.md (git blob a37042b915e5170b5cf706a4c253276b50db8aa0). See wiki/TRANSLATION_GUIDE.md. -->
# docxi へのコントリビュート

バグ報告、修正、ドキュメント、翻訳など、どれも歓迎します。このページでは、問題の報告方法と、プルリクエストに必要なものを説明します。作業用のコピーの準備については[開発環境のセットアップ](Development-Setup)を参照してください。

## <a id="reporting-a-bug"></a>バグの報告

[github.com/hammad417/docxi/issues](https://github.com/hammad417/docxi/issues) で Issue を作成し、次の情報を添えてください。

- 問題を再現する短い Ruby スクリプト
- Ruby、docxi、nokogiri、rubyzip のバージョン（Rails を使っている場合はそのバージョンも）
- 期待した結果と、実際の結果。ドキュメントが壊れている場合は、開いたアプリケーション名も

セキュリティ上の問題は非公開で報告してください。[セキュリティ](Security)を参照してください。

## <a id="pull-requests"></a>プルリクエスト

1. リポジトリをフォークし、ブランチを作成します。
2. 変更を加えます。周囲のコードのスタイルに合わせてください。
3. テストを追加または更新し、`bundle exec rake test` を実行します。
4. 生成される XML を意図的に変えた場合は、`bundle exec rake golden` を実行し、プルリクエストで変更内容を説明してください。ゴールデンファイルは、既存のユーザーが頼りにしている出力を守るためのものです。
5. 動作が変わる場合は、`CHANGELOG.md` とドキュメントを更新します。
6. プルリクエストを作成します。CI が Ruby 2.5 から 4.0、Rails 5.2 から 8.1 でテストを実行し、LibreOffice でサンプルドキュメントを開けるかを確認します。

0.2.0 で予定している修正のように、既存のドキュメントを変える変更は、ゴールデンファイルの差分をレビューしやすいように、独立したプルリクエストにしてください。

## <a id="documentation-and-translations"></a>ドキュメントと翻訳

- 英語の README は、`docs/i18n/` にある 15 の翻訳の元になっています。README を変更するときは、リポジトリの `docs/i18n/TRANSLATION_GUIDE.md` に従ってください。
- この Wiki はリポジトリの `wiki/` フォルダーで書かれ、そこから公開されます。手順は `wiki/README.md` にあります。

## <a id="related-pages"></a>関連ページ

- [開発環境のセットアップ](Development-Setup)
- [アーキテクチャ](Architecture)
- [ロードマップ](Roadmap)

---

[← Wiki ホーム](Home) · [README](https://github.com/hammad417/docxi#readme)
