# -*- coding: utf-8 -*-
"""Монтажная сетка слайдов презентации одной PNG (быстрый просмотр глазами).

Запуск:
  python scripts/pptx_thumb.py <file.pptx> [-o grid.png] [-W 480] [-c 4]
      — рендерит все слайды через PowerPoint COM (pptx_render.ps1) во
      временную папку и собирает сетку с номерами слайдов;
  python scripts/pptx_thumb.py <папка_с_slideN.png> [-o grid.png]
      — собирает сетку из уже отрендеренных PNG (_build\\render).

PNG ≤1500 px по ширине читаются зрением модели целиком — один файл вместо
N штук экономит контекст на просмотре проекта/шаблона.
"""
import re
import subprocess
import sys
import tempfile
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

HERE = Path(__file__).resolve().parent
RENDER_PS1 = HERE / "pptx_render.ps1"
PAD, LABEL_H, BG = 12, 28, (255, 255, 255)


def find_slides(pptx):
    """Свежий рендер pptx через COM во временную папку; список PNG по номерам."""
    with tempfile.TemporaryDirectory(prefix="deck_thumb_") as td:
        r = subprocess.run(
            ["powershell", "-ExecutionPolicy", "Bypass", "-File", str(RENDER_PS1),
             "-Src", str(pptx), "-Dst", td, "-Width", "1000"],
            capture_output=True, text=True)
        if r.returncode != 0:
            sys.exit("рендер не удалось:\n" + (r.stderr or r.stdout))
        imgs = sort_pngs(Path(td).glob("slide*.png"))
        out_dir = Path(tempfile.mkdtemp(prefix="deck_grid_"))
        for p in imgs:
            p.replace(out_dir / p.name)
    return sort_pngs(out_dir.glob("slide*.png"))


def sort_pngs(paths):
    def num(p):
        m = re.search(r"(\d+)", p.stem)
        return int(m.group(1)) if m else 10**9
    return sorted(paths, key=num)


def grid(imgs, out, tile_w=480, cols=4):
    tiles = []
    for p in imgs:
        im = Image.open(p).convert("RGB")
        h = round(im.height * tile_w / im.width)
        tiles.append((p, im.resize((tile_w, h), Image.LANCZOS)))
    rows = (len(tiles) + cols - 1) // cols
    tile_h = max(t.height for _, t in tiles)
    W = cols * tile_w + (cols + 1) * PAD
    H = rows * (tile_h + LABEL_H) + (rows + 1) * PAD
    canvas = Image.new("RGB", (W, H), BG)
    d = ImageDraw.Draw(canvas)
    try:
        font = ImageFont.truetype("arial.ttf", 18)
    except OSError:
        font = ImageFont.load_default()
    for n, (p, t) in enumerate(tiles, 1):
        r, c = divmod(n - 1, cols)
        x = PAD + c * (tile_w + PAD)
        y = PAD + r * (tile_h + LABEL_H + PAD)
        d.text((x + 2, y), f"{n}  {p.name}", font=font, fill=(41, 41, 41))
        canvas.paste(t, (x, y + LABEL_H))
    canvas.save(out)
    return out


def main(src, out=None, tile_w=480, cols=4):
    src = Path(src)
    if src.is_dir():
        imgs = sort_pngs(src.glob("slide*.png"))
        if not imgs:
            sys.exit(f"в {src} нет файлов slideN.png")
    elif src.suffix.lower() == ".pptx":
        imgs = find_slides(src)
    else:
        sys.exit("нужен .pptx или папка с slideN.png")
    if not imgs:
        sys.exit("слайдов не найдено")
    out = Path(out) if out else (src.parent / f"{src.stem}_grid.png")
    print(grid(imgs, out, tile_w, cols))
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(2)
    args, opts = [], {}
    it = iter(sys.argv[1:])
    for a in it:
        if a == "-o":
            opts["out"] = next(it)
        elif a == "-W":
            opts["tile_w"] = int(next(it))
        elif a == "-c":
            opts["cols"] = int(next(it))
        else:
            args.append(a)
    sys.exit(main(args[0], **opts))
