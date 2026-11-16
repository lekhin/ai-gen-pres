# -*- coding: utf-8 -*-
"""Единый CLI фабрики презентаций (точка входа вместо длинных путей).

Запуск из любой папки (путь к deck.py — канонический для фабрики, но скилл может
быть установлен куда угодно; служебная папка ищется сама, см. deck_paths.py;
указать её вручную: переменная окружения DECK_SERVICE_DIR):
  python _service\\skills\\universal-pptx\\scripts\\deck.py inspect <file.pptx> [--full]
  python _service\\skills\\universal-pptx\\scripts\\deck.py themes
  python _service\\skills\\universal-pptx\\scripts\\deck.py build   plan.json [-o out.pptx] [--theme dark]
  python _service\\skills\\universal-pptx\\scripts\\deck.py mmd     diag.mmd [-o diag.json] [--emit py] [--preview out.pptx]
  python _service\\skills\\universal-pptx\\scripts\\deck.py qa      <file.pptx> [-t тема]
  python _service\\skills\\universal-pptx\\scripts\\deck.py check   <file.pptx>
  python _service\\skills\\universal-pptx\\scripts\\deck.py render  <file.pptx> [-o dir] [-W 1500]
                                                            # Win: COM; Linux: soffice+pdftoppm
  python _service\\skills\\universal-pptx\\scripts\\deck.py thumb   <file.pptx | папка\\render> [-o grid.png] [-c 4]
  python _service\\skills\\universal-pptx\\scripts\\deck.py make-template [тема …]  # пересобрать template.pptx из theme.json

Тема оформления: --theme / поле "theme" в plan.json / env DECK_THEME
(приоритет по убыванию); список тем — `themes`. Тема фиксируется при импорте
helper'ов, поэтому CLI выставляет DECK_THEME ДО импорта.
Команды — тонкая обёртка над pptx_inspect / pptx_qa / pptx_check_ph /
pptx_render.ps1 (или soffice-рендер в Linux) / pptx_thumb; сами скрипты
остаются рабочими по отдельности.
qa = плейсхолдеры + текст + структура пакета ДО рендера (код возврата 1 = ошибки).
"""
import argparse
import json
import os
import shutil
import subprocess
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")
HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

import pptx_check_ph  # noqa: E402
import pptx_inspect  # noqa: E402
import pptx_thumb  # noqa: E402


def cmd_render(a):
    dst = Path(a.out) if a.out else Path(a.file).parent / "_build" / "render"
    dst.mkdir(parents=True, exist_ok=True)
    if a.fast and _render_deckcraft(a, dst):
        pass                                    # внутренний цикл: без Office
    elif sys.platform == "win32":
        cmd = ["powershell", "-ExecutionPolicy", "Bypass",
               "-File", str(HERE / "pptx_render.ps1"),
               "-Src", a.file, "-Dst", str(dst),
               "-Width", str(a.width)]
        if a.svg:
            cmd.append("-Svg")
        rc = subprocess.run(cmd).returncode
        if rc:
            return rc
    else:
        if a.svg:
            print("render: --svg — только Windows (PowerPoint COM); делаю PNG")
        rc = _render_linux(a, dst)
        if rc:
            return rc
    if a.trim:
        return _trim(dst)
    return 0


def _render_deckcraft(a, dst):
    """Быстрый рендер ArtCraft DeckCraft (если установлен): все слайды в PNG
    headless, ~100× быстрее COM; для ФИНАЛЬНОЙ приёмки всё равно COM/soffice
    (DeckCraft — чистая комната, мелкие расхождений с PowerPoint возможны)."""
    from deck_paths import artcraft_cli
    dc = artcraft_cli("deckcraft")
    if not dc:
        print("render: --fast — deckcraft-cli не найден, рендерю обычно")
        return False
    r = subprocess.run([dc, "render", a.file, "--all", str(dst)]
                       + (["--transparent"] if a.transparent else []),
                       capture_output=True, text=True)
    if r.returncode:
        print((r.stderr or r.stdout).strip()[:300] or "render: deckcraft-cli отказал")
        return False
    for png in dst.glob("slide-*.png"):         # slide-01.png → канон. slide1.png
        png.rename(png.with_name(f"slide{int(png.stem.split('-')[-1])}.png"))
    _downscale(dst, a.width)
    print(f"rendered (deckcraft, быстро): {dst} — финальную приёмку смотри COM-рендером")
    return True


def _trim(dst):
    """Обрезка полей рендеров до bounding box контента (иллюстрация как
    самостоятельный файл). Фон = цвет угла страницы, допуск — шум сглаживания."""
    try:
        from PIL import Image, ImageChops
    except ImportError:
        print("trim: нужен Pillow")
        return 2
    n, pad = 0, 0
    for png in sorted(dst.glob("slide*.png")):
        with Image.open(png) as im:
            rgb = im.convert("RGB")
            bg = Image.new("RGB", rgb.size, rgb.getpixel((2, 2)))
            diff = ImageChops.difference(rgb, bg).convert("L")
            bbox = diff.point(lambda v: 255 if v > 12 else 0).getbbox()
            if bbox:
                pad = max(4, rgb.width // 150)
                box = (max(0, bbox[0] - pad), max(0, bbox[1] - pad),
                       min(rgb.width, bbox[2] + pad), min(rgb.height, bbox[3] + pad))
                im.crop(box).save(png)          # режим (RGBA при --transparent) сохраняется
                n += 1
    print(f"trim: обрезано {n} PNG (фон по углу страницы, поле {pad} px)")
    return 0


def _render_linux(a, dst):
    """Рендер вне Windows: soffice → pdf → pdftoppm (или PyMuPDF) с каноническими
    именами slideN.png (N=1..), шириной <= --width. Скрипт-обход ручных цепочек
    и рассинхона имён PNG в Linux-контейнерах."""
    src = Path(a.file)
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        print("render: нет ни PowerPoint (Windows), ни soffice/libreoffice (Linux)")
        return 2
    import tempfile
    with tempfile.TemporaryDirectory() as td:
        r = subprocess.run([soffice, "--headless", "--convert-to", "pdf",
                            "--outdir", td, str(src)],
                           capture_output=True, text=True)
        pdfs = list(Path(td).glob("*.pdf"))
        if r.returncode or not pdfs:
            print((r.stderr or r.stdout or "").strip() or "render: soffice не выдал pdf")
            return 1
        pdf = pdfs[0]
        # 13.333" ширина слайда: dpi => пикселей по ширине
        dpi = max(50, round(a.width / (33.87 / 2.54)))
        made = 0
        if shutil.which("pdftoppm"):
            r = subprocess.run([shutil.which("pdftoppm"), "-png", "-r", str(dpi),
                                str(pdf), str(dst / "raw")],
                               capture_output=True, text=True)
            if r.returncode:
                print((r.stderr or r.stdout).strip())
                return 1
            for i, png in enumerate(sorted(dst.glob("raw-*.png")), 1):
                png.rename(dst / f"slide{i}.png")   # raw-01.png → slide1.png
                made += 1
        else:
            try:
                import fitz  # PyMuPDF
            except ImportError:
                print("render: нет ни pdftoppm, ни PyMuPDF (pip install pymupdf)")
                return 2
            doc = fitz.open(pdf)
            for i, page in enumerate(doc, 1):
                pm = page.get_pixmap(matrix=fitz.Matrix(dpi / 72, dpi / 72))
                pm.save(dst / f"slide{i}.png")
                made += 1
            doc.close()
        _downscale(dst, a.width)
        print(f"rendered {made} slides -> {dst} (slideN.png)")
    return 0


def _downscale(dst, max_w):
    """Урезать PNG шире max_w (зрение модели рассчитана на <=1500 px)."""
    try:
        from PIL import Image
    except ImportError:
        return
    for png in dst.glob("slide*.png"):
        with Image.open(png) as im:
            if im.width > max_w:
                im = im.resize((max_w, round(im.height * max_w / im.width)),
                               Image.LANCZOS)
                im.save(png)


def load_plan_checked(path):
    """Поймать частые ошибки ручных plan.json (кавычки/запятые/true) с указанием строки."""
    try:
        txt = Path(path).read_text(encoding="utf-8-sig")  # BOM (notepad/PS) допустим
    except FileNotFoundError:
        print(f"план не найден: {path}")
        return None
    try:
        json.loads(txt)
    except json.JSONDecodeError as e:
        lines = txt.splitlines()
        ctx = "\n".join(f"  {n}: {lines[n - 1]}"
                        for n in range(max(1, e.lineno - 2),
                                       min(len(lines), e.lineno + 2) + 1))
        print(f"НЕ JSON, строка {e.lineno}, колонка {e.colno}: {e.msg}\n{ctx}\n"
              f"Частые причины: кавычки-«ёлочки» вместо \", лишняя/недостающая "
              f"запятая, true/None вместо true, конкатенация двух объектов.")
        return None
    return txt


def peek_plan_theme(path):
    """Тема из plan.json (поле "theme") — чтобы выставить DECK_THEME ДО импорта
    helper'ов: их палитра вычисляется при импорте."""
    try:
        return json.loads(Path(path).read_text(encoding="utf-8-sig")).get("theme")
    except Exception:
        return None


def main():
    p = argparse.ArgumentParser(prog="deck.py", description=__doc__,
                                formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--theme", "-T", help="тема оформления (см. команду themes)")
    sub = p.add_subparsers(dest="cmd", required=True)

    sub.add_parser("themes", help="список тем оформления")

    s = sub.add_parser("inspect", help="дамп структуры pptx")
    s.add_argument("file")
    s.add_argument("--full", action="store_true", help="с полными XML слайдов")

    s = sub.add_parser("qa", help="аудит пакета/плейсхолдеров/текста до рендера")
    s.add_argument("file")
    s.add_argument("-t", "--theme", dest="file_theme",
                   help="сверять цвета/шрифты с этой темой (без неё — со всеми)")

    s = sub.add_parser("check", help="только аудит плейсхолдеров")
    s.add_argument("file")

    s = sub.add_parser("render", help="рендер слайдов в PNG slideN.png "
                                      "(Windows: PowerPoint COM; Linux: soffice+pdftoppm)")
    s.add_argument("file")
    s.add_argument("-o", "--out", help="папка PNG (по умолчанию <pptx>_parent\\_build\\render)")
    s.add_argument("-W", "--width", type=int, default=1500)
    s.add_argument("--trim", action="store_true",
                   help="обрезать поля до контента (иллюстрация-файл; с -W 2400 — печатное качество)")
    s.add_argument("--svg", action="store_true", help="ещё и SVG по слайдам (Windows COM)")
    s.add_argument("--fast", action="store_true",
                   help="внутренний цикл: headless-рендер ArtCraft DeckCraft без Office "
                        "(если установлен); финальная приёмка — обычный COM/soffice-рендер")
    s.add_argument("--transparent", action="store_true",
                   help="фон слайда прозрачный (--fast; иллюстрация на подложку заказчика)")

    s = sub.add_parser("dump", help="читатель материалов: docx/pdf/pptx/md → "
                                    "нумерованный текст (+картинки из docx)")
    s.add_argument("file")
    s.add_argument("-o", "--out", help="файл дампа (иначе stdout)")
    s.add_argument("--media", nargs="?", const="", metavar="ПАПКА",
                   help="извлечь word/media из docx (по умолчанию <проект>\\_build\\media)")

    s = sub.add_parser("facts", help="сверка чисел на слайдах с facts.md + приёмочная строка")
    s.add_argument("file")
    s.add_argument("--facts", required=True, help="facts.md (таблица значение|статус|источник)")
    s.add_argument("--strict", action="store_true",
                   help="rc=1, если на слайдах есть числа не из facts.md")

    s = sub.add_parser("thumb", help="монтировка всех слайдов одной PNG")
    s.add_argument("src", help=".pptx или папка со slideN.png")
    s.add_argument("-o", "--out", help="файл сетки (по умолчанию <имя>_grid.png рядом)")
    s.add_argument("-c", "--cols", type=int, default=4)
    s.add_argument("-W", "--tile", type=int, default=480, help="ширина плитки, px")

    s = sub.add_parser("build", help="собрать pptx из декларативного плана plan.json")
    s.add_argument("plan")
    s.add_argument("-o", "--out", help="куда сохранить (по умолчанию plan['output'])")

    s = sub.add_parser("make-template", help="сгенерировать template.pptx из theme.json")
    s.add_argument("themes", nargs="*", help="имена тем (по умолчанию все найденные)")

    s = sub.add_parser("mmd", help="mermaid flowchart → схема нативными фигурами (json/py/preview)")
    s.add_argument("file", help="mermaid-файл (или - для stdin)")
    s.add_argument("-o", "--out", help="файл результата (json/py; иначе stdout)")
    s.add_argument("--emit", choices=("json", "py"), default="json")
    s.add_argument("--preview", metavar="PPTX", help="собрать однослайдовый pptx")
    s.add_argument("--title", help="заголовок слайда превью")
    s.add_argument("--legend", help='легенда: "solid=запросы,dash=OIDC"')
    s.add_argument("--dir", dest="direction", choices=("LR", "RL", "TB", "TD", "BT"),
                   help="переопределить направление (TB-схемы в LR читабельнее)")
    s.add_argument("--block", type=int, metavar="N",
                   help="N-й flowchart-блок mermaid в .md (несколько схем в файле)")

    a = p.parse_args()
    # тема ДО импорта helper'ов: CLI-флаг сильнее плана, план сильнее env
    theme = a.theme if a.cmd != "qa" else None
    if a.cmd == "build":
        theme = theme or peek_plan_theme(a.plan)
    if theme:
        os.environ["DECK_THEME"] = theme

    if a.cmd == "themes":
        import deck_theme
        for n in deck_theme.list_themes():
            t = deck_theme.load(n)
            mark = "*" if n == (os.environ.get("DECK_THEME") or deck_theme.DEFAULT_THEME) else " "
            print(f"{mark} {n:12s} {t['title']:12s} {t['description']}")
        return 0
    if a.cmd == "inspect":
        pptx_inspect.main(a.file, a.full)
        return 0
    if a.cmd == "qa":
        import pptx_qa
        return pptx_qa.main(a.file, theme=a.file_theme or a.theme)
    if a.cmd == "check":
        return pptx_check_ph.main(a.file)
    if a.cmd == "render":
        return cmd_render(a)
    if a.cmd == "dump":
        import deck_dump
        return deck_dump.main(a.file, a.out, a.media)
    if a.cmd == "facts":
        import deck_facts
        return deck_facts.main(a.file, a.facts, strict=a.strict)
    if a.cmd == "thumb":
        return pptx_thumb.main(a.src, a.out, a.tile, a.cols)
    if a.cmd == "build":
        if load_plan_checked(a.plan) is None:
            return 1
        import plan as plan_mod
        out = plan_mod.build(a.plan, a.out)
        print(f"built: {out}")
        return 0
    if a.cmd == "make-template":
        import deck_make_template
        return deck_make_template.main(a.themes)
    if a.cmd == "mmd":
        import mmd2diagram
        argv = [a.file, "--emit", a.emit]
        if a.out:
            argv += ["-o", a.out]
        if a.preview:
            argv += ["--preview", a.preview]
        if a.title:
            argv += ["--title", a.title]
        if a.legend:
            argv += ["--legend", a.legend]
        if a.direction:
            argv += ["--dir", a.direction]
        if a.block:
            argv += ["--block", str(a.block)]
        return mmd2diagram.main(argv)


if __name__ == "__main__":
    sys.exit(main())
