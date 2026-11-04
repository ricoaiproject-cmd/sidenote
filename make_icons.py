"""通知領域のアイコンを作る。icon.ico（いつもの青灰）と icon_rec.ico（記録中・右下に橙の点）"""
import sys
from pathlib import Path
from PIL import Image, ImageDraw
out = Path(sys.argv[1]); S = 1024
BLUE, ORANGE, WHITE = (74, 98, 120, 255), (196, 119, 47, 255), (255, 255, 255, 255)
def draw(rec):
    im = Image.new("RGBA", (S, S), (0, 0, 0, 0)); d = ImageDraw.Draw(im)
    d.ellipse((0, 0, S - 1, S - 1), fill=BLUE)
    cx = S // 2  # マイク（画面左上のマークと同じ形）
    d.rounded_rectangle((cx - 120, 190, cx + 120, 600), radius=120, fill=WHITE)
    d.arc((cx - 230, 300, cx + 230, 720), start=0, end=180, fill=WHITE, width=70)
    d.rectangle((cx - 35, 700, cx + 35, 830), fill=WHITE)
    d.rounded_rectangle((cx - 150, 800, cx + 150, 860), radius=30, fill=WHITE)
    if rec:  # 記録中の印。小さく表示しても見えるよう大きめにし、白い縁で背景から浮かせる
        r = 250; x0, y0 = S - 2 * r - 10, S - 2 * r - 10
        d.ellipse((x0 - 40, y0 - 40, x0 + 2 * r + 40, y0 + 2 * r + 40), fill=WHITE)
        d.ellipse((x0, y0, x0 + 2 * r, y0 + 2 * r), fill=ORANGE)
    return im
sizes = [(16, 16), (20, 20), (24, 24), (32, 32), (40, 40), (48, 48), (64, 64), (256, 256)]
for name, rec in (("icon.ico", False), ("icon_rec.ico", True)):
    im = draw(rec); im.save(out / name, sizes=sizes)
    im.resize((128, 128), Image.LANCZOS).save(out / name.replace(".ico", "_preview.png"))
    for s in (16, 32):  # 実際の表示の大きさでも確かめる
        im.resize((s, s), Image.LANCZOS).resize((128, 128), Image.NEAREST).save(out / name.replace(".ico", f"_{s}px.png"))
