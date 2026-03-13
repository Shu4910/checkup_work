"""
neinei.png から neinei.ico を生成し、
デスクトップに install_and_run.bat へのショートカットを作成するスクリプト。
"""

import subprocess
import sys
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent

# ── neinei.png を探す ──────────────────────────────────────────
candidates = [
    SCRIPT_DIR / "neinei.png",
    SCRIPT_DIR.parent / "neinei.png",
]
png_path = next((p for p in candidates if p.exists()), None)

# ── ICO 生成 ──────────────────────────────────────────────────
ico_path = SCRIPT_DIR / "neinei.ico"

if png_path and not ico_path.exists():
    try:
        from PIL import Image
        img = Image.open(png_path).convert("RGBA")
        img.save(
            ico_path,
            format="ICO",
            sizes=[(256, 256), (128, 128), (64, 64), (32, 32), (16, 16)],
        )
        print(f"アイコン生成完了: {ico_path}")
    except Exception as e:
        print(f"[警告] アイコン生成に失敗しました: {e}")
        ico_path = None
elif ico_path.exists():
    print(f"アイコンファイルを使用: {ico_path}")
else:
    print("[警告] neinei.png が見つかりません。アイコンなしでショートカットを作成します。")
    ico_path = None

# ── デスクトップショートカット作成 ────────────────────────────
bat_path = SCRIPT_DIR / "install_and_run.bat"
desktop = Path.home() / "Desktop"
lnk_path = desktop / "Activity Analyzer.lnk"

icon_line = f'$Shortcut.IconLocation = "{ico_path}"' if ico_path else ""

ps_script = f"""
$WshShell = New-Object -comObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut("{lnk_path}")
$Shortcut.TargetPath = "{bat_path}"
$Shortcut.WorkingDirectory = "{SCRIPT_DIR}"
{icon_line}
$Shortcut.Save()
"""

result = subprocess.run(
    ["powershell", "-ExecutionPolicy", "Bypass", "-Command", ps_script],
    capture_output=True,
    text=True,
)

if result.returncode == 0:
    print(f"デスクトップショートカット作成完了: {lnk_path}")
else:
    print(f"[エラー] ショートカットの作成に失敗しました:")
    print(result.stderr)
    sys.exit(1)
