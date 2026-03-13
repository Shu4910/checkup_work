"""
neinei.png から neinei.ico を生成し、
デスクトップに Activity Analyzer のショートカットを作成するスクリプト。
"""

import subprocess
import sys
from pathlib import Path

ANALY_DIR = Path(__file__).parent
PNG_PATH  = ANALY_DIR / "neinei.png"      # analy/ 内の neinei.png を参照
ICO_PATH  = ANALY_DIR / "neinei.ico"
BAT_PATH  = ANALY_DIR / "launch_analyzer.bat"
DESKTOP   = Path.home() / "Desktop"
LNK_PATH  = DESKTOP / "Activity Analyzer.lnk"

print("=" * 50)
print("  Activity Analyzer セットアップ")
print("=" * 50)
print()

# ── ICO 生成 ──────────────────────────────────────────────────
if not ICO_PATH.exists():
    if not PNG_PATH.exists():
        print(f"[警告] アイコン画像が見つかりません: {PNG_PATH}")
        print("       neinei.png を analy フォルダ内に置いてください。")
        print("       アイコンなしでショートカットを作成します。")
        ICO_PATH = None
    else:
        try:
            from PIL import Image
            img = Image.open(PNG_PATH).convert("RGBA")
            img.save(
                ICO_PATH,
                format="ICO",
                sizes=[(256, 256), (128, 128), (64, 64), (32, 32), (16, 16)],
            )
            print(f"アイコン生成完了: {ICO_PATH}")
        except ImportError:
            print("[警告] Pillow がインストールされていないためアイコンを生成できません。")
            ICO_PATH = None
        except Exception as e:
            print(f"[警告] アイコン生成に失敗しました: {e}")
            ICO_PATH = None
else:
    print(f"アイコンファイルを使用: {ICO_PATH}")

# ── デスクトップショートカット作成 ────────────────────────────
icon_line = f'$Shortcut.IconLocation = "{ICO_PATH}"' if ICO_PATH else ""

ps_script = f"""
$WshShell = New-Object -comObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut("{LNK_PATH}")
$Shortcut.TargetPath = "{BAT_PATH}"
$Shortcut.WorkingDirectory = "{ANALY_DIR}"
$Shortcut.Description = "Activity Analyzer - 業務カテゴリ別時間分析"
{icon_line}
$Shortcut.Save()
"""

result = subprocess.run(
    ["powershell", "-ExecutionPolicy", "Bypass", "-Command", ps_script],
    capture_output=True,
    text=True,
)

if result.returncode == 0:
    print(f"デスクトップショートカット作成完了: {LNK_PATH}")
    print()
    print("デスクトップの「Activity Analyzer」アイコンをダブルクリックして起動できます。")
else:
    print(f"[エラー] ショートカットの作成に失敗しました:")
    print(result.stderr)
    sys.exit(1)
