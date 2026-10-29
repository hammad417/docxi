<!-- Translated from wiki/Development-Setup.md (git blob 6fd2ffba3077016b05d1d9ab9467d15fc88c8e7d). See wiki/TRANSLATION_GUIDE.md. -->
# 開発環境のセットアップ：docxi のテストとツールを実行する

このページでは、docxi の作業用コピーを動かす方法と、テスト、サンプルドキュメント、CI の構成を説明します。

## <a id="get-the-code"></a>コードを取得する

```sh
git clone https://github.com/hammad417/docxi.git
cd docxi
bundle install
bundle exec rake test
```

開発には Ruby 2.5 以降ならどのバージョンでも使えます。開発用の依存関係は minitest と rake です。

## <a id="rake-tasks"></a>Rake タスク

| コマンド | 内容 |
|---|---|
| `bundle exec rake test` | テストスイートを実行します（デフォルトのタスク） |
| `bundle exec rake samples` | サンプルドキュメントを `tmp/samples/*.docx` に書き出します。Word、LibreOffice、Google Docs で開いて確認できます |
| `bundle exec rake golden` | 現在のコードから `test/golden` を書き直します。出力を意図的に変えた後にだけ使ってください |
| `bundle exec rake build` | gem を `pkg/` にビルドします |

## <a id="running-the-tests-with-rails-loaded"></a>Rails を読み込んだ状態でテストを実行する

`gemfiles/` フォルダーには、Rails のバージョンごとの Gemfile と、古い rubyzip や nokogiri のバージョン用の Gemfile がいくつかあります。

```sh
BUNDLE_GEMFILE=gemfiles/rails_8.1.gemfile bundle install
BUNDLE_GEMFILE=gemfiles/rails_8.1.gemfile DOCXI_WITH_RAILS=1 bundle exec rake test
```

`DOCXI_WITH_RAILS=1` は、テストの前にフレームワーク全体を読み込みます。また、docxi の `blank?` の規則を ActiveSupport の規則と比較します。

## <a id="what-the-tests-cover"></a>テストの対象

| ファイル | 確認する内容 |
|---|---|
| `test/test_golden_output.rb` | 2 つのサンプルドキュメントのすべてのパーツが、0.0.6 の生成結果と一致すること |
| `test/test_package.rb` | XML が整形式であること、リレーションシップが存在するパーツを指していること、すべてのパーツにコンテンツタイプがあること、Zip64 レコードがないこと |
| `test/test_media.rb` | パス、IO オブジェクト、URL からの画像と、`\|` の文字列がコマンドとして実行されないこと |
| `test/test_fixed_features.rb` | `tab`、`frame`、ヘッダーのページ番号 |
| `test/test_util.rb` | `Docxi::Util.blank?` が ActiveSupport の規則に従うこと |
| `test/test_rails_independence.rb` | docxi が ActiveSupport なしで動くこと |
| `test/test_warnings.rb` | レンダリング時に `lib/` から非推奨警告が出ないこと |

サンプルドキュメントは `test/support/sample_documents.rb` にあり、`test/support/package.rb` は `.docx` を zip の構造から直接読み込みます。

## <a id="continuous-integration"></a>継続的インテグレーション

`.github/workflows/ci.yml` は、`master` へのプッシュとすべてのプルリクエストで実行されます。対象は次のとおりです。

- Rails なしの Ruby 2.5、2.6、4.0
- Rails 5.2（Ruby 2.5）、6.0（2.7）、6.1（3.0）、7.0（3.1）、7.1（3.2）、7.2（3.3）、8.0（3.4）、8.1（4.0）
- Rails `main`（失敗しても許容）
- rubyzip 1.3 と 2.4、nokogiri 1.11
- LibreOffice でサンプルドキュメントを PDF に変換するジョブ

リリースについては[デプロイ](Deployment)で説明しています。

## <a id="related-pages"></a>関連ページ

- [コントリビュート](Contributing)
- [アーキテクチャ](Architecture)
- [デプロイ](Deployment)

---

[← Wiki ホーム](Home) · [README](https://github.com/hammad417/docxi#readme)
