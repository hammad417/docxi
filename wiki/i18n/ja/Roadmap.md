<!-- Translated from wiki/Roadmap.md (git blob c9fde5b9818383a900ca60bf167fac5895ec82c5). See wiki/TRANSLATION_GUIDE.md. -->
# docxi のロードマップと既知の問題

docxi にはリリーススケジュールも、確約されたロードマップもありません。予定として記録されている作業は、以下の既知の問題の修正だけです。修正は 0.2.0 で予定していますが、時期は決まっていません。既存のユーザーがすでに生成しているドキュメントが変わってしまうため、0.1.0 には含めていません。

## <a id="known-issues-planned-for-020"></a>0.2.0 で対応予定の既知の問題

| 問題 | 0.1.0 での回避策 |
|---|---|
| 配置を指定した画像、ヘッダーの画像、表のセル内の画像は、幅と高さの両方を `width * height * 14.92` EMU として計算する | `style:` でサイズを指定したインライン画像を使うか、[ベストプラクティス](Best-Practices#images)を参照 |
| 表単位の `borders:` が表の外枠ではなく各セルに適用される | 外枠だけが必要なら、端のセルの外側の辺に `borders:` を設定する |
| ページ番号のフィールドコードで `\* MERGEFORMAT` のバックスラッシュが欠けている | なし |
| 表のセル内のページ番号の先頭に「GlobalOptions」が付く | ページ番号はヘッダーかフッターに入れる |
| 画像が常に PNG パーツとして保存され、`[Content_Types].xml` が最初の 2 つのヘッダーとフッターしか宣言しない | PNG 画像を使い、ヘッダーとフッターはそれぞれ 2 つまでにする |
| 余白の一部だけを渡すと、残りが空で書き出される | `top`、`right`、`bottom`、`left` をまとめて渡す |
| `docProps/core.xml` がローカル時刻に UTC の印を付けて記録し、作成者を書き込まない | なし |
| `Docxi.test` が `NoMethodError` を発生させる | 呼び出さない。以前のバージョンから残っているヘルパーです |
| `page: true` なしの `document.br` と `document.tab` が無効な XML を生成する | 段落の中で `p.br` と `p.tab` を使う |

## <a id="not-planned"></a>予定していないこと

以下は上の一覧に含まれておらず、何も決まっていません。

- 既存のドキュメントの編集や、テンプレートへの値の埋め込み
- 番号付きリスト、見出しスタイル、自動生成の目次
- 他のページサイズや横向きのページ

## <a id="related-pages"></a>関連ページ

- [変更履歴](Changelog)
- [コントリビュート](Contributing)
- [よくある質問](FAQ)

---

[← Wiki ホーム](Home) · [README](https://github.com/hammad417/docxi#readme)
