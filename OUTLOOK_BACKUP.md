# Outlookメール定期バックアップツール

## 改善ポイント
- 既存コードの方針（送信者フィルタ / MSG保存 / 添付抽出 / HTML保存）を維持しています。
- 初回設定で「対象Outlookアカウント」「送信者メールアドレス」「保存先」「間隔（分）」を入力し、`backup_config.json` に保存します。
- 次回起動以降は設定を自動読込して、そのまま定期実行します。
- 重複保存は `backup_state.json`（`アカウント名:EntryID`）で防止します。

## アカウント指定
- `*` を指定: Outlookにある**全アカウント**を対象
- 文字列を指定: アカウント名に部分一致するものだけ対象（例: `contoso.com`）


## バックアップ間隔
GUIで次を選択できます。
- 15分 / 30分 / 1時間 / 6時間 / 12時間
- 1日 / 1週間 / 1か月(30日)
- カスタム(分)

## 生成されるバックアップ
アカウントごとにサブフォルダを作成して保存します。
- `<保存先>/<アカウント名>/msg/` : 元メール（.msg）
- `<保存先>/<アカウント名>/attachments/` : 添付ファイル
- `<保存先>/<アカウント名>/html/` : 本文HTML（整形済み）

## 実行（Python）
```bash
python outlook_backup.py
```

## exe化（PythonなしPC向け）
開発PCで以下を実行します。

```bash
python -m pip install -r outlook_backup_requirements.txt
pyinstaller --onefile --noconsole --name outlook_backup outlook_backup.py
```

配布物:
- `dist/outlook_backup.exe`

## 配布先PCの前提
- Windows
- Outlookデスクトップアプリがインストール済み
- 対象アカウントがOutlookに設定済み

## 自動起動（任意）
ログオン時に自動起動したい場合は、`outlook_backup.exe` のショートカットを以下に置いてください。
- `shell:startup`

## ログ
- `logs/outlook_backup.log`
