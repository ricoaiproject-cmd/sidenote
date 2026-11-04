"""Sidenote のショートカットを作る（Windows）。

  python install_shortcuts.py               スタートメニューとスタートアップ（起動時に裏で待機）に登録
  python install_shortcuts.py --no-startup  スタートメニューだけ
  python install_shortcuts.py --desktop     デスクトップにも作る
  python install_shortcuts.py --remove      作ったショートカットを消す
"""
import os
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
APP = HERE / "meeting_app.py"
ICON = HERE / "icon.ico"
NAME = "Sidenote.lnk"


def places(startup, desktop):
    appdata = Path(os.environ["APPDATA"])
    out = [(appdata / "Microsoft" / "Windows" / "Start Menu" / "Programs" / NAME, "")]
    if startup:  # Windows の起動時は窓を出さずに常駐だけする
        out.append((appdata / "Microsoft" / "Windows" / "Start Menu" / "Programs" / "Startup" / NAME, "--background"))
    if desktop:
        out.append((Path.home() / "Desktop" / NAME, ""))
    return out


def main():
    args = sys.argv[1:]
    targets = places(startup="--no-startup" not in args, desktop="--desktop" in args)
    if "--remove" in args:
        for path, _ in places(True, True):
            if path.exists():
                path.unlink()
                print(f"消しました: {path}")
        return
    # Microsoft Store 版は %APPDATA% への書き込みが別の場所へ振り替えられ、登録も設定の共有もできない
    if "WindowsApps" in sys.executable:
        sys.exit("Microsoft Store 版の Python では登録できません。python.org 版か winget の Python で実行してください（setup.bat なら自動で選びます）")
    # 窓（黒い画面）を出さずに動かすため、python.exe ではなく pythonw.exe で起動する
    pythonw = Path(sys.executable).with_name("pythonw.exe")
    if not pythonw.exists():
        sys.exit(f"pythonw.exe が見つかりません: {pythonw}")
    from win32com.client import Dispatch
    shell = Dispatch("WScript.Shell")
    for path, extra in targets:
        path.parent.mkdir(parents=True, exist_ok=True)
        lnk = shell.CreateShortCut(str(path))
        lnk.TargetPath = str(pythonw)
        lnk.Arguments = f'"{APP}" {extra}'.strip()
        lnk.WorkingDirectory = str(HERE)
        if ICON.exists():
            lnk.IconLocation = str(ICON)
        lnk.Description = "Sidenote — 会議を文字にして Claude に聞ける道具"
        lnk.save()
        print(f"作りました: {path}")


if __name__ == "__main__":
    main()
