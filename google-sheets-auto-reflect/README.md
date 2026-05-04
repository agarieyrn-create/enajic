# 3施設別在庫・報告数字 自動反映

既存の牛肉棚卸管理表から、アロマ加工場・今帰仁冷凍施設のブロック肉入出庫状況を読み取り、日付別の報告書フォームへ自動反映するGoogle Apps Scriptです。

## 連携先

- 報告先Spreadsheet: https://docs.google.com/spreadsheets/d/1p4np21i3CzGXYu0piVnI_WjMOYc8Eo9CS0D47G0-QEk/edit
- 元データSpreadsheet: https://docs.google.com/spreadsheets/d/1K-cR1Cw4ycldHk6wl4MLwjnL0UejdTYJXIpycNaE8ck/edit
- Apps Script: https://script.google.com/d/1n-T9jg3noaqFUqGey9ouB9fLdYB4iIXPECtClDgZvKD7eMp2pSFDpuaQ/edit

## 読み取り元

- `2026.3～今帰仁冷凍施設(ブロック肉)`
  - 日付行: `90`
  - ヘッダー行: `91`
  - 明細開始行: `92`
- `2026.3～アロマ加工場（ブロック肉）`
  - 日付行: `89`
  - ヘッダー行: `90`
  - 明細開始行: `91`

`設定` シートの `報告対象日` に一致する日付列から `入庫 / 出庫 / 残数` を読み取り、1日前の日付列の `残数` を `前日` として使います。

## 作成・更新されるシート

- `設定`: 報告対象日、元データSpreadsheet ID、元シート名、PDF保存先を管理。
- `部位マスタ`: 報告書フォームの部位リスト。
- `3施設別在庫`: 読み取った日付・施設・部位別の集計結果。
- `3施設別報告表`: プレビュー用の報告書フォーム。
- `YYYY.M.D(3)`: `日付別報告作成` で作成/更新される日次報告書。

## 使い方

1. 報告先Spreadsheetを開く。
2. メニュー `3施設 自動反映 > 初期セットアップ` を初回実行する。
3. `設定` の `報告対象日` を報告したい日にする。
4. `3施設 自動反映 > 日付別報告作成` を実行する。
5. `2026.5.4(3)` のような日付シートが作成/更新される。

`カンゲンファームブロック肉在庫` 欄は、今回の元データに対象シートがないためv1では0で出力します。

## ボタン化

スプレッドシート上で図形を挿入し、右上メニューから「スクリプトを割り当て」を選ぶと、以下の関数をボタンにできます。

- `createDailyReport`
- `autoReflect`
- `exportReportPdf`

## 開発

```bash
clasp push -f
```

このリポジトリの `google-sheets-auto-reflect/.clasp.json` は、報告先Spreadsheetに紐づくApps Scriptを指しています。
