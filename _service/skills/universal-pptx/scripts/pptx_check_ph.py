# -*- coding: utf-8 -*-
"""Аудит плейсхолдеров собранного pptx: пустые заглушки и поля-колонтитулы.

Запуск:  python scripts/pptx_check_ph.py <file.pptx>
Выход:   построчно «№слайда: idx:тип (текст|ПУСТО!) …», в конце — вердикт OK
         или список проблем (пустые плейсхолдеры = PowerPoint покажет
         заглушку «Текст слайда» в режиме правки).
Код возврата: 0 — чисто, 1 — есть пустые плейсхолдеры.
"""
import sys

from pptx import Presentation


def main(path):
    prs = Presentation(path)
    problems = []
    for i, s in enumerate(prs.slides, 1):
        items = []
        for p in s.placeholders:
            has = bool(p.text_frame.text.strip()) if p.has_text_frame else False
            tag = f"idx{p.placeholder_format.idx}:{p.placeholder_format.type}"
            items.append(f"{tag}{'(текст)' if has else '(ПУСТО!)'}")
            if p.has_text_frame and not has:
                problems.append((i, p.placeholder_format.idx))
        print(i, " | ".join(items))
    if problems:
        print("ПРОБЛЕМЫ: пустые плейсхолдеры на слайдах:",
              ", ".join(f"слайд {i} idx {idx}" for i, idx in problems))
        return 1
    print("OK: пустых плейсхолдеров нет")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(2)
    sys.exit(main(sys.argv[1]))
