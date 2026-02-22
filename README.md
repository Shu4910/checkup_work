# Personal Productivity Tracker

自分自身のPC使用状況を記録して生産性を可視化するWindows用ツールです。

> **データはあなたのPCのみに保存されます。外部への送信は一切ありません。**

## 概要

どのアプリを何時間使っているか、いつ離席しているかを自動記録し、日々の作業を振り返るためのツールです。収集したデータはローカルのCSVファイルにのみ保存され、ネットワーク通信は行いません。

## 主な機能

- **アクティビティ記録**: 30秒ごとにアクティブなアプリケーションとウィンドウタイトルを記録
- **ブラウザURL取得**: Chrome、Edge、Firefoxの現在表示中のURLとページタイトルを取得
- **本日のレポート**: トレイメニューから今日の使用時間をアプリ別に確認
- **スリープ検知**: PCスリープ時は自動的にポーリングを停止
- **ロック検知**: PCロック時は状態を`locked`として記録（離席時間の把握）
- **CSVログ出力**: 1日1ファイルでCSV形式で保存
- **システムトレイ常駐**: システムトレイにアイコンを表示し、バックグラウンドで動作
- **自動ログローテーション**: 90日以上古いログファイルは自動削除

## システム要件

- **OS**: Windows 10 (64bit) / Windows 11
- **Python**: 3.12以降
- **メモリ**: 50MB以上の空きメモリ

## インストール

### 1. 依存パッケージのインストール

```bash
pip install -r requirements.txt
```

### 2. pywin32のインストール後処理

```bash
python -m pywin32_postinstall -install
```

## 設定

`config.json`で以下の設定が可能です。

```json
{
    "log_directory": "%USERPROFILE%\\Documents\\ProductivityTracker\\logs",
    "polling_interval": 30,
    "log_retention_days": 90
}
```

| 設定項目 | 説明 | デフォルト値 |
|---------|------|-------------|
| `log_directory` | ログ保存先ディレクトリ（`%USERPROFILE%` などの環境変数が使えます） | `%USERPROFILE%\Documents\ProductivityTracker\logs` |
| `polling_interval` | ポーリング間隔（秒） | 30 |
| `log_retention_days` | ログ保存期間（日数） | 90 |

## 使い方

### 起動

```bash
python src/activity_monitor.py
```

起動するとシステムトレイにアイコンが表示されます。

### システムトレイメニュー

- **現在の状態を表示**: 実行状態・ログ保存先を確認
- **本日のレポートを表示**: 今日のアクティブ時間・離席時間・アプリ別使用時間（上位5件）を表示
- **終了**: アプリケーションを終了（確認ダイアログが表示されます）

### レポート表示例

```
本日のアクティビティ (2026-02-22)
────────────────────────────────
アクティブ:    5時間23分
離席/ロック:   1時間12分
────────────────────────────────
アプリ別使用時間 (上位5件):
  chrome.exe          2時間15分
  excel.exe           1時間30分
  code.exe              45分
```

### ログファイル

- 保存先: `%USERPROFILE%\Documents\ProductivityTracker\logs\`
- ファイル名: `activity_log_YYYY-MM-DD.csv`（1日1ファイル）

#### CSVフォーマット

```csv
timestamp,machine_name,user_name,process_name,window_title,url,page_title,status
2026-02-22 09:01:00,MY-PC,username,chrome.exe,Google Chrome,https://mail.google.com/,受信トレイ - Gmail,active
2026-02-22 09:02:00,MY-PC,username,excel.exe,作業記録.xlsx - Microsoft Excel,,,active
2026-02-22 09:03:00,MY-PC,username,,,,,locked
```

## 自動起動設定（オプション）

Windowsタスクスケジューラを使用して、PC起動時に自動起動させることができます。

1. タスクスケジューラを開く
2. 「基本タスクの作成」を選択
3. トリガーを「コンピューターの起動時」に設定
4. 操作を「プログラムの起動」に設定し、`pythonw.exe` を指定
5. 引数に `src\activity_monitor.py` の絶対パスを設定
6. 開始ディレクトリにプロジェクトのルートディレクトリを設定

## トラブルシューティング

### ブラウザURLが取得できない

- UIAutomationがブラウザのアドレスバーにアクセスできない場合があります
- ブラウザの言語設定が日本語でない場合、`browser_tracker.py`の`Name`パラメータを調整してください

### システムトレイアイコンが表示されない

- Pillowが正しくインストールされているか確認してください
- Windowsの通知領域設定で、アイコンが非表示になっていないか確認してください

## プライバシー

- 収集したデータはすべてローカルPCの指定フォルダにのみ保存されます
- ネットワーク通信は一切行いません
- ログファイルは`.gitignore`で除外されており、Gitリポジトリには含まれません

## ライセンス

MIT License

## バージョン

- **バージョン**: 2.0.0
- **更新日**: 2026-02-22
