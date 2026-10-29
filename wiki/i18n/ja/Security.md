<!-- Translated from wiki/Security.md (git blob 2358bb39bd6772aaea144aa8ca2434b43c086728). See wiki/TRANSLATION_GUIDE.md. -->
# docxi のセキュリティ：問題の報告と、信頼できない入力の扱い

このページでは、セキュリティ上の問題の報告方法と、アプリケーションがユーザーの入力を docxi に渡すときに知っておくべきセキュリティ上の注意点を説明します。

## <a id="reporting-a-vulnerability"></a>脆弱性の報告

セキュリティ上の問題は、公開の Issue では報告しないでください。代わりにメンテナーにメールで連絡してください。メールアドレスは [`docxi.gemspec`](https://github.com/hammad417/docxi/blob/master/docxi.gemspec) に記載しています。docxi のバージョンと、問題を再現する簡単な手順を添えてください。

## <a id="upgrade-from-00x"></a>0.0.x からアップグレードする

docxi 0.0.x は画像のパスの文字列を `Kernel#open` に渡します。`Kernel#open` は `|` で始まる文字列をシェルコマンドとして実行するため、画像のパスを操作できる攻撃者は、サーバー上でコマンドを実行できてしまいます。0.1.0 はファイルを `File.open` で開き、URL を `URI#open` でダウンロードします。どちらもコマンドを実行しません。

## <a id="untrusted-input"></a>信頼できない入力

**テキストは安全に渡せます**。テキスト、表の内容、その他の文字列は、XML のテキストまたは属性値として書き出され、nokogiri がエスケープします。ユーザーの入力でドキュメントに要素を追加することはできません。

**画像のソースには注意が必要です**。`add_media` は次のものを受け付けます。

| ソース | docxi の動作 |
|---|---|
| 自分で開いた IO | それを読み込みます |
| ファイルパス | サーバーのディスクからそのファイルを読み込みます |
| `http`、`https`、`ftp` の URL | サーバーからダウンロードします |

ユーザーがパスを選べる場合、設定ファイルや鍵など、サーバー上の読み取り可能な任意のファイルをドキュメントに埋め込ませることができてしまいます。ユーザーが URL を選べる場合、サーバーがその URL にリクエストを送るため、内部のアドレスにも到達しえます（サーバーサイドリクエストフォージェリ、SSRF）。自分で開いた IO を渡すか、既知のパスとホストだけを許可してください。

**ハイパーリンク**は、URL が示す場所をそのまま指します。ユーザーが指定した URL は、`add_hyperlink` に渡す前に確認してください。

## <a id="how-releases-are-protected"></a>リリースの保護

- リリースは GitHub Actions から RubyGems の Trusted Publishing で公開されるため、リポジトリには API キーを保存していません。
- gem のメタデータで `rubygems_mfa_required` を設定しているため、個人の API キーでのプッシュには多要素認証が必要です。

## <a id="related-pages"></a>関連ページ

- [ベストプラクティス](Best-Practices)
- [デプロイ](Deployment)
- [コントリビュート](Contributing)

---

[← Wiki ホーム](Home) · [README](https://github.com/hammad417/docxi#readme)
