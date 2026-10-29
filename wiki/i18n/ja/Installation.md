<!-- Translated from wiki/Installation.md (git blob 0bd43a0434cbf17fa4dcd95d837e40ca50c52119). See wiki/TRANSLATION_GUIDE.md. -->
# docxi のインストール：必要な環境、Bundler の設定、アップグレード

docxi は Ruby だけで書かれた gem です。インストール時に追加される依存関係は nokogiri と rubyzip の 2 つだけです。オフィスソフトは必要ありません。

## <a id="requirements"></a>必要な環境

| 項目 | バージョン |
|---|---|
| Ruby | 2.5 以降 |
| nokogiri | `>= 1.8, < 2`（自動でインストールされます） |
| rubyzip | `>= 1.3, < 4`（自動でインストールされます） |
| Rails | 不要 |

## <a id="install-with-bundler"></a>Bundler でインストールする

Gemfile に gem を追加します。

```ruby
gem "docxi", "~> 0.1"
```

続いてインストールします。

```sh
bundle install
```

## <a id="install-without-bundler"></a>Bundler を使わずにインストールする

```sh
gem install docxi
```

## <a id="check-the-installation"></a>インストールを確認する

```sh
ruby -e 'require "docxi"; puts Docxi::VERSION'
```

Bundler を使うプロジェクトでは、`bundle exec ruby -e 'require "docxi"; puts Docxi::VERSION'` として実行します。インストールされたバージョン（例：`0.1.0`）が表示されます。

## <a id="upgrading-from-00x"></a>0.0.x からのアップグレード

Gemfile で docxi が 0.0.x に固定されている場合（例：`"~> 0.0.6"`）は、その行を `gem "docxi", "~> 0.1"` に変更します。そのうえで `bundle update docxi` を実行します。コードの変更は不要で、0.1.0 は 0.0.6 と同じドキュメントを生成します。テストスイートが出力を 0.0.6 で生成したファイルと比較しています。

内部で変わった点：

- nokogiri と rubyzip は上記の範囲内のどのバージョンでもよくなったため、docxi を現行の Rails と一緒にインストールできます。
- docxi は ActiveSupport が読み込まれていることを前提にしなくなりました。
- アーカイブは Zip64 レコードなしで書き出されます。
- 画像の読み込みに `Kernel#open` を使わなくなったため、Ruby 3 でも URL が使えます。
- `tab`、`frame`、ヘッダーのページ番号が動作するようになりました。

Ruby 2.5 より古い Ruby で動くアプリケーションには、引き続き 0.0.6 がインストールされます。0.1.0 は Ruby 2.5 を最低バージョンとして宣言しているためです。すべての変更は[変更履歴](Changelog)にあります。

## <a id="versions-tested-in-ci"></a>CI でテストしているバージョン

各 Rails バージョンは、フレームワーク全体を docxi と一緒に読み込んだ状態でテストしています。

| Ruby | Rails |
|---|---|
| 2.5 | 5.2 |
| 2.7 | 6.0 |
| 3.0 | 6.1 |
| 3.1 | 7.0 |
| 3.2 | 7.1 |
| 3.3 | 7.2 |
| 3.4 | 8.0 |
| 4.0 | 8.1 |

CI では次も実行しています。

- Ruby 2.5、2.6、4.0 での Rails なしのテスト
- rubyzip 1.3（Ruby 2.5）、rubyzip 2.4（Ruby 3.3）、nokogiri 1.11（Ruby 2.5）でのテスト
- LibreOffice でサンプルドキュメントを変換できるかの確認

## <a id="platform-notes"></a>プラットフォームに関する注意

- nokogiri は、一般的な Linux、macOS、Windows 向けにプリコンパイル済みの gem を提供しています。該当するものがない場合はソースからコンパイルされ、C コンパイラが必要です。[nokogiri のインストールガイド](https://nokogiri.org/tutorials/installing_nokogiri.html)を参照してください。
- CI は Ubuntu 24.04 で動いており、テストは macOS でも実行済みです。Windows ではテストしていません。

## <a id="if-installation-fails"></a>インストールに失敗する場合

nokogiri や rubyzip で Bundler の競合が起きる場合、ほぼ確実に古い docxi 0.0.6 が解決されています。[トラブルシューティング](Troubleshooting#installation)を参照してください。

## <a id="related-pages"></a>関連ページ

- [はじめての docxi](Getting-Started)
- [設定](Configuration)
- [デプロイ](Deployment)

---

[← Wiki ホーム](Home) · [README](https://github.com/hammad417/docxi#readme)
