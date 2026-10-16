# マイ出納帳

iPhone 向けの手入力式の家計簿（出納帳）アプリです。

- 毎日の支出と毎月の収入を手入力で記帳します
- 項目の追加・削除・名前の変更、メモの入力候補、毎月の定期支払いの自動記帳に対応しています
- 月別・年間の集計を Excel（.xlsx）に書き出せます
- データは端末（ブラウザ）の中にだけ保存され、このリポジトリやサーバーには送られません
- 設定画面からバックアップ（.json）の保存と復元ができます
- ご自身の Google スプレッドシートへ自動保存できます（準備は [SETUP_GOOGLE.md](SETUP_GOOGLE.md)）

GitHub Pages で公開し、iPhone の Safari で開いて「ホーム画面に追加」して使います。

`xlsx.full.min.js` は SheetJS Community Edition 0.18.5（Apache License 2.0、`xlsx.LICENSE` を参照）です。
