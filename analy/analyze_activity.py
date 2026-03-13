"""
activity_log_*.csv を読み込み、業務カテゴリ別の時間分析を行うスクリプト。

出力:
  - コンソール: 日別カテゴリ集計
  - output/activity_categorized_<日時>.csv : カテゴリ付き全データ
"""

import sys
from datetime import datetime
from pathlib import Path

# Windows コンソールの文字コードを UTF-8 に強制
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

try:
    import pandas as pd
except ImportError:
    print("=" * 60)
    print("  [エラー] pandas がインストールされていません")
    print("=" * 60)
    print()
    print("以下のコマンドを実行してください:")
    print("    pip install pandas")
    print()
    input("Enterキーを押して終了...")
    sys.exit(1)

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox
except ImportError:
    print("[エラー] tkinter が使用できません。Python を公式サイトから再インストールしてください。")
    input("Enterキーを押して終了...")
    sys.exit(1)

# ─────────────────────────────────────────────
# 設定
# ─────────────────────────────────────────────
OUTPUT_DIR = Path(__file__).parent / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

MAX_GAP_SEC = 300  # 5分

# ─────────────────────────────────────────────
# カテゴリ定義
# ─────────────────────────────────────────────
CATEGORY_COLORS = {
    "ミーティング":           "#4C72B0",
    "ウェビナー・展示会対応": "#DD8452",
    "メールマーケティング":   "#55A868",
    "リード管理":             "#C44E52",
    "メール対応":             "#8172B2",
    "データ分析・レポート":   "#937860",
    "プロジェクト管理":       "#DA8BC3",
    "コンテンツ制作":         "#8C8C8C",
    "広告運用":               "#CCB974",
    "コミュニケーション":     "#64B5CD",
    "AI活用":                 "#39CCCC",
    "開発・ツール作業":       "#2ECC40",
    "HubSpot":                "#FF7A59",
    "immedio":                "#1E88E5",
    "展示会":                 "#FB8C00",
    "ウェビナー":             "#FDD835",
    "ビデオ会議":             "#43A047",
    "クロードコード":         "#9C27B0",
    "経理":                   "#FF6B6B",
    "メモ":                   "#90A4AE",
    "雑務・その他":           "#DDDDDD",
    "離席":                   "#F0F0F0",
}

CHROME_WINDOW_RULES = [
    ("HubSpot",    ["HubSpot"]),
    ("経理",       ["バクラク"]),
    ("展示会",     ["expo", "展示会"]),
    ("ウェビナー", ["ウェビナー"]),
    ("ビデオ会議", ["MEET", "bizibl"]),
]

CHROME_RULES = [
    ("ミーティング", [
        "Meet -", "Google Meet", "Zoom Workplace", "Zoom から参加",
        "ウェビナーを編集", "ウェビナー情報 - Zoom",
        "本日のセミナーの「評価」",
        "immedio - カレンダー",
    ]),
    ("ウェビナー・展示会対応", [
        "ウェビナー経由推移", "動画ウェビナー", "展示会経由推移",
        "展示会・カンファレンス", "MarkeZine", "EightEXPO", "Eight Expo",
        "セミナー・イベント", "登壇", "カンファレンス",
        "インポートシートセミナー",
    ]),
    ("メールマーケティング", [
        "Eメールの編集", "マーケティングEメール", "メルマガ配信カレンダー",
        "Test send", "メール確認 - メール本文",
    ]),
    ("リード管理", [
        "セグメント", "コンタクト | 全て", "インポート詳細", "インポート",
        "immedio Forms", "Forms | HubSpot", "immedio\n",
        "データ連携", "プロパティー設定",
    ]),
    ("メール対応", [
        "受信トレイ", "immedio メール", "- Gmail", "Outlook",
        "検索結果 - komaki", "送信", "Test send - ",
        "【ご提出のお願い】", "【MarkeZine Day】", "【immedio】",
        "マーケ全体推移_FY25 - komaki",
        "動画ウェビナー経由推移_202510 - komaki",
    ]),
    ("データ分析・レポート", [
        "スプレッドシート", "Salesforce", "レポートビルダー",
        "アナリティクス", "OKR", "工数管理", "目標設計",
        "マーケ全体推移", "ウェビナー経由推移", "ホワイトペーパー",
        "Bill One", "AI活用申請フォーム",
    ]),
    ("プロジェクト管理", [
        "Notion", "Google ドキュメント", "ドキュメント",
    ]),
    ("コンテンツ制作", [
        "Google スライド", "Figma", "Canva",
        "WordPress", "投稿を編集", "マーケ全般",
        "immedio（イメディオ） — WordPress",
        "プレスリリース",
    ]),
    ("広告運用", [
        "広告マネージャ",
    ]),
    ("経理", [
        "バクラク",
    ]),
    ("コミュニケーション", [
        "Slack", "Messenger", "Chat", "LINE",
        "Facebook",
    ]),
    ("AI活用", [
        "ChatGPT", "Claude",
    ]),
    ("雑務・その他", [
        "新しいタブ", "最近のダウンロード", "Google 検索", "- Google 検索",
        "SBI証券", "ポートフォリオ", "口座管理",
        "ホームズ", "スーツ", "シーラベル",
        "Keeper SSO", "HubSpot Login",
        "開く", "無題", "Google ドライブ",
    ]),
]

PROCESS_CATEGORY = {
    "LockApp.exe":              "離席",
    "Slack.exe":                "コミュニケーション",
    "ChatGPT.exe":              "AI活用",
    "claude.exe":               "クロードコード",
    "Code.exe":                 "開発・ツール作業",
    "python.exe":               "開発・ツール作業",
    "EXCEL.EXE":                "データ分析・レポート",
    "POWERPNT.EXE":             "コンテンツ制作",
    "Figma.exe":                "コンテンツ制作",
    "Acrobat.exe":              "雑務・その他",
    "Notepad.exe":              "メモ",
    "SnippingTool.exe":         "雑務・その他",
    "Photos.exe":               "雑務・その他",
    "explorer.exe":             "雑務・その他",
    "CredentialUIBroker.exe":   "離席",
    "ShellHost.exe":            "離席",
    "Clibor.exe":               "雑務・その他",
    "LINE.exe":                 "コミュニケーション",
    "Zoom.exe":                 "ミーティング",
}


def classify_chrome(page_title: str) -> str:
    if not isinstance(page_title, str):
        return "雑務・その他"
    for category, keywords in CHROME_RULES:
        for kw in keywords:
            if kw in page_title:
                return category
    if page_title.strip().lower() == "immedio":
        return "リード管理"
    return "雑務・その他"


def classify_row(row) -> str:
    process = row["process_name"]
    if process in PROCESS_CATEGORY:
        return PROCESS_CATEGORY[process]
    if process == "chrome.exe":
        window_title = str(row["window_title"]) if pd.notna(row["window_title"]) else ""
        page_title   = str(row["page_title"])   if pd.notna(row["page_title"])   else ""
        url          = str(row["url"])           if pd.notna(row.get("url"))      else ""

        if "chatgpt.com" in url:
            return "AI活用"

        for category, keywords in CHROME_WINDOW_RULES:
            for kw in keywords:
                if kw in window_title:
                    return category

        category = classify_chrome(page_title)
        if category != "雑務・その他":
            return category

        if "immedio" in window_title or "immedio" in page_title:
            return "immedio"

        return "雑務・その他"
    return "雑務・その他"


def select_log_files() -> list:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    files = filedialog.askopenfilenames(
        title="分析するログCSVファイルを選択してください（複数可）",
        filetypes=[("CSVファイル", "*.csv"), ("すべてのファイル", "*.*")],
    )
    root.destroy()
    return sorted(files)


def load_logs(files: list) -> pd.DataFrame:
    if not files:
        raise FileNotFoundError("ファイルが選択されませんでした。")

    dfs = []
    for f in files:
        try:
            df = pd.read_csv(f, encoding="utf-8-sig")
            dfs.append(df)
        except Exception as e:
            print(f"  [警告] {f} の読み込みに失敗しました: {e}")

    if not dfs:
        raise ValueError("読み込めたファイルがありませんでした。")

    df = pd.concat(dfs, ignore_index=True)

    df["timestamp"] = pd.to_datetime(df["timestamp"], errors="coerce")
    df = df.dropna(subset=["timestamp"])
    df = df[df["process_name"].notna() & (df["process_name"] != "process_name")]
    df = df.sort_values("timestamp").reset_index(drop=True)

    df["page_title"] = df["page_title"].fillna(df["window_title"]).fillna("")
    df["category"]   = df.apply(classify_row, axis=1)
    df["date"]       = df["timestamp"].dt.date

    df["next_ts"]      = df["timestamp"].shift(-1)
    df["duration_sec"] = (df["next_ts"] - df["timestamp"]).dt.total_seconds()
    df["duration_sec"] = df["duration_sec"].clip(upper=MAX_GAP_SEC).fillna(30)
    df["duration_min"] = df["duration_sec"] / 60

    return df


def summarize_by_day(df: pd.DataFrame) -> pd.DataFrame:
    return (
        df[df["category"] != "離席"]
        .groupby(["date", "category"])["duration_min"]
        .sum()
        .unstack(fill_value=0)
    )


def print_summary(summary: pd.DataFrame):
    print("\n" + "=" * 60)
    print("  業務カテゴリ別 時間サマリー（分）")
    print("=" * 60)
    for date, row in summary.iterrows():
        total = row.sum()
        print(f"\n【{date}】  合計: {total:.0f}分 ({total/60:.1f}h)")
        for cat, mins in row.sort_values(ascending=False).items():
            if mins > 0:
                bar = "#" * int(mins / 10)
                print(f"  {cat:<18} {mins:5.0f}分  {bar}")
    print("\n" + "=" * 60)


def main():
    files = select_log_files()
    if not files:
        print("ファイルが選択されませんでした。終了します。")
        input("\nEnterキーを押して終了...")
        return

    print(f"{len(files)} 件のファイルを読み込み中...")
    for f in files:
        print(f"  {f}")

    try:
        df = load_logs(files)
    except Exception as e:
        print(f"\n[エラー] ファイルの読み込みに失敗しました:\n  {e}")
        input("\nEnterキーを押して終了...")
        return

    timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_csv = OUTPUT_DIR / f"activity_categorized_{timestamp_str}.csv"

    try:
        df.to_csv(out_csv, index=False, encoding="utf-8-sig")
        print(f"\nカテゴリ付きCSV出力: {out_csv}")
    except PermissionError:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "書き込みエラー",
            f"CSVファイルへの書き込みに失敗しました。\n{out_csv}\n\n"
            "Excelなどで開いている場合は閉じてから再実行してください。",
        )
        root.destroy()
        return
    except Exception as e:
        print(f"\n[エラー] CSV出力に失敗しました:\n  {e}")
        input("\nEnterキーを押して終了...")
        return

    summary = summarize_by_day(df)
    print_summary(summary)

    print("\n完了。output/ フォルダを確認してください。")
    input("\nEnterキーを押して終了...")


if __name__ == "__main__":
    main()
