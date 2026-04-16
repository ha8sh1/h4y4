# 部品情報CSV統合ツール

要望に合わせて **HTMLページで操作するツール** を追加しました。

- 画面: `parts_web_tool.html`
- 既存CLI: `part_inventory_csv_tool.py`（バッチ用途として継続利用可）

## HTML版でできること
- ユーザー名を不完全一致で検索し、入力欄から選択
- ユーザーを追加するたびに部品リストへ追記・再集約
- 集約結果を「CSV出力」ボタンでダウンロード
- JSON / HTMLテーブル形式のレスポンスに対応（`auto`判定）

## 使い方（HTML版）
1. `parts_web_tool.html` をブラウザで開く
2. `取得URL` と `ユーザーパラメータ` を設定
3. `候補ユーザー一覧` にユーザー名を改行またはカンマ区切りで入力
4. `ユーザー選択` に一部の文字を入力して候補から選び、`ユーザー追加`
5. 必要ユーザーを追加後、`CSV出力`

## 注意
- URLの `#no-back` などフラグメントはHTTP送信されません。
- 認証が必要なサイトでは、同一オリジンでページを配信し Cookie を使う構成にしてください（`fetch(..., { credentials: 'include' })`）。
- 別オリジンAPIにアクセスする場合は CORS 設定が必要です。

## CLI版（参考）
```bash
python3 part_inventory_csv_tool.py \
  --base-url 'http://egl620043.jpn.mds.honda.com/11G/3V5_WHS/#no-back' \
  --users u001,u002,u003 \
  --user-param user \
  --format auto \
  --header 'Cookie: session=YOUR_SESSION' \
  --output parts_merged.csv
```
