# -*- coding: utf-8 -*-
"""Дамп структуры .pptx: размеры, мастер, layout'ы, слайды, плейсхолдеры, шрифты.

Запуск:  python pptx_inspect.py <file.pptx> [--full]
--full — дополнительно печатать полные XML слайдов (для разбора полей/наследований).
"""
import sys
import io
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

from pptx import Presentation
from pptx.util import Emu


def fmt_len(v):
    return f"{Emu(v).cm:.2f}cm" if v is not None else "None"


def dump_shape(sh, ind="  "):
    pos = ""
    if sh.left is not None:
        pos = (f"pos=({fmt_len(sh.left)},{fmt_len(sh.top)}) "
               f"size=({fmt_len(sh.width)}x{fmt_len(sh.height)})")
    print(f"{ind}name={sh.name!r} type={sh.shape_type} {pos}")
    if sh.is_placeholder:
        pf = sh.placeholder_format
        print(f"{ind}  ph idx={pf.idx} type={pf.type}")
    if sh.has_text_frame:
        for p in sh.text_frame.paragraphs:
            if not p.runs:
                continue
            for r in p.runs:
                f = r.font
                col = None
                try:
                    if f.color is not None and f.color.type is not None:
                        col = str(f.color.rgb) if f.color.type == 1 else f"theme:{f.color.theme_color}"
                except Exception:
                    pass
                rpr = r.font._rPr
                spc = rpr.get("spc") if rpr is not None else None
                cap = rpr.get("cap") if rpr is not None else None
                print(f"{ind}  run={r.text[:60]!r} font={f.name} size={f.size and f.size.pt} "
                      f"bold={f.bold} italic={f.italic} color={col} spc={spc} cap={cap}")


def main(path, full=False):
    prs = Presentation(path)
    try:                                    # в дамп — относительный путь (переносимость)
        shown = Path(path).relative_to(Path.cwd())
    except (ValueError, NameError):
        shown = path
    print(f"FILE: {shown}")
    print(f"slide size: {prs.slide_width} x {prs.slide_height} EMU = "
          f"{Emu(prs.slide_width).cm:.2f} x {Emu(prs.slide_height).cm:.2f} cm")
    print(f"slides: {len(prs.slides)}   masters: {len(prs.slide_masters)}")
    for mi, m in enumerate(prs.slide_masters):
        print(f"=== MASTER {mi} ===")
        for sh in m.shapes:
            dump_shape(sh)
        for li, l in enumerate(m.slide_layouts):
            print(f"=== LAYOUT {li}: {l.name!r} ===")
            for sh in l.shapes:
                dump_shape(sh)
    for si, s in enumerate(prs.slides):
        print(f"=== SLIDE {si + 1} (layout={s.slide_layout.name!r}) ===")
        for sh in s.shapes:
            dump_shape(sh)
        if full:
            print(s.element.xml)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    main(sys.argv[1], "--full" in sys.argv)
