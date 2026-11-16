# -*- coding: utf-8 -*-
"""Штатный читатель материалов задачи: docx/pdf/pptx/md/txt → нумерованный текст
(+ извлечение картинок из docx). Входной тракт режима «из материалов»: ссылки в
facts.md держатся за НОМЕРА СТРОК этого дампа, поэтому нумерация печатается и в
файл, и в stdout, и пустые строки не нумеруются (ссылки стабильны).

deck.py dump <файл> [-o дамп.txt] [--media [папка]]
  docx — абзацы и строки таблиц (" ячейка | ячейка "), word/media/* → в папку;
  pdf  — постраничный текст (нужен pypdf: pip install pypdf);
  pptx — текст по слайдам ("=== слайд N ===");
  md/txt — как есть (utf-8, фолбэк cp1251).

Скрипт рабочий и без deck.py: python deck_dump.py <файл> …
"""
import re
import sys
import zipfile
from pathlib import Path

NS_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _w_text(p):
    """Текст w:p:runs + табы как пробелы."""
    parts = []
    for node in p.iter():
        tag = node.tag
        if tag == NS_W + "t":
            parts.append(node.text or "")
        elif tag in (NS_W + "tab", NS_W + "noBreakHyphen"):
            parts.append(" ")
        elif tag in (NS_W + "br", NS_W + "cr"):
            parts.append(" ")
    return "".join(parts).strip()


def dump_docx(path):
    """docx → [(строки…)]; таблицы — по одной строкой на строку таблицы."""
    lines = []
    with zipfile.ZipFile(path) as z:
        from lxml import etree
        root = etree.fromstring(z.read("word/document.xml"))
        body = root.find(NS_W + "body")
        for el in body:
            if el.tag == NS_W + "p":
                t = _w_text(el)
                if t:
                    lines.append(t)
            elif el.tag == NS_W + "tbl":
                for tr in el.findall(NS_W + "tr"):
                    cells = []
                    for tc in tr.findall(NS_W + "tc"):
                        cell = " ".join(filter(None, (_w_text(p)
                                                      for p in tc.iter(NS_W + "p"))))
                        cells.append(cell)
                    if any(cells):
                        lines.append(" | ".join(cells))
    return lines


def dump_pdf(path):
    from deck_paths import artcraft_cli
    pdf = artcraft_cli("pdfcraft")
    if pdf:                                       # ArtCraft PdfCraft: извлечение крепче pypdf
        import subprocess
        r = subprocess.run([pdf, "text", str(path)], capture_output=True, text=True)
        if not r.returncode and r.stdout.strip():
            return [l for l in r.stdout.replace("\f", "\n").splitlines() if l.strip()]
    try:
        from pypdf import PdfReader
    except ImportError:
        sys.exit("dump(pdf): нужен pypdf (pip install pypdf) или ArtCraft PdfCraft "
                 "(pdfcraft-cli в PATH / DECK_ARTCRAFT_DIR)")
    lines = []
    for i, page in enumerate(PdfReader(str(path)).pages, 1):
        lines.append(f"[страница {i}]")
        lines.extend(t for t in (page.extract_text() or "").splitlines() if t.strip())
    return lines


def dump_pptx(path):
    from pptx import Presentation
    lines = []
    for i, s in enumerate(Presentation(str(path)).slides, 1):
        lines.append(f"=== слайд {i} ===")

        def walk(shapes):
            for sh in shapes:
                if sh.shape_type == 6:                       # группа
                    walk(sh.shapes)
                elif sh.has_text_frame and sh.text_frame.text.strip():
                    lines.append(sh.text_frame.text.strip().replace("\n", " ⏎ "))
        walk(s.shapes)
    return lines


def dump_text(path):
    raw = Path(path).read_bytes()
    for enc in ("utf-8-sig", "cp1251"):
        try:
            txt = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    return [t for t in (l.strip() for l in txt.splitlines()) if t]


def extract_media(path, out_dir):
    """word/media/* из docx в папку; возвращает список файлов."""
    out_dir.mkdir(parents=True, exist_ok=True)
    made = []
    with zipfile.ZipFile(path) as z:
        for n in z.namelist():
            if n.startswith("word/media/"):
                dst = out_dir / Path(n).name
                dst.write_bytes(z.read(n))
                made.append(dst)
    return made


def main(src, out=None, media=None):
    src = Path(src)
    if not src.exists():
        sys.exit(f"dump: файл не найден: {src}")
    kind = src.suffix.lower()
    lines = {"docx": dump_docx, "pdf": dump_pdf, "pptx": dump_pptx}.get(
        kind.lstrip("."), dump_text)(src)
    text = "\n".join(f"{i:5d} | {t}" for i, t in enumerate(lines, 1))
    if out:
        Path(out).write_text(text + "\n", encoding="utf-8")
        print(f"dump: {len(lines)} строк -> {out}")
    else:
        print(text)
    if media is not None and kind == ".docx":
        mdir = Path(media) if media else src.parent / "_build" / "media"
        made = extract_media(src, mdir)
        print(f"media: {len(made)} файлов -> {mdir}")
        for f in made:
            print(f"  {f.name}  {f.stat().st_size // 1024} КБ")
    elif media is not None:
        print("media: извлечение картинок пока только для docx")
    return 0


if __name__ == "__main__":
    args = sys.argv[1:]
    if not args:
        sys.exit(__doc__)
    out = media = None
    if "-o" in args:
        i = args.index("-o")
        out = args[i + 1]
        del args[i:i + 2]
    if "--media" in args:
        i = args.index("--media")
        media = ""
        if i + 1 < len(args) and not args[i + 1].startswith("-"):
            media = args[i + 1]
            del args[i:i + 2]
        else:
            del args[i]
    sys.exit(main(args[0], out, media))
