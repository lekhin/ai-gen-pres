# -*- coding: utf-8 -*-
"""Сверка данных: числа на слайдах ↔ значения из facts.md (дисциплина AGENTS.md).

deck.py facts <file.pptx> --facts <facts.md> [--strict]

facts.md — таблица `значение | статус | источник` (первая колонка — значение,
вторая содержит ключевое слово sourced/derived/common_knowledge/estimate/missing,
третья — источник). Токенизация одинаковая с обеих сторон: числа, коды, даты
(«2.6.4115-25», «323-ФЗ», «12 %») сводятся к одним токенам.

Что печатает:
  * счётчики статусов и ГОТОВУЮ приёмочную строку AGENTS.md (часть про
    наличие источников в материалах подтверждает агент — инструмент её помечает);
  * токены со слайдов, которых нет в facts.md (реальные расхождения; шум —
    сквозная нумерация списков: она отбрасывается, если это маркер в начале
    абзаца вида «01 »);
  * значения из facts.md, не найденные на слайдах (могли съехать).
Текст плейсхолдеров (заголовок, колонтитулы, номер слайда, дата) не сверяется —
служебная область. --strict: код возврата 1 при непривязанных токенах.
"""
import re
import sys
from pathlib import Path

from pptx import Presentation

TOKEN = re.compile(r"\d[\d\u00a0 .,/:\-–—]*(?=\D|$)")
NORM = re.compile(r"[\s\u00a0]+")
STATUSES = ("sourced", "derived", "common_knowledge", "estimate", "missing")


def tokens(text):
    """Значимые числовые токены текста, нормализованные (пробелы внутри вынуты,
    разделители списков/дат ',' '/' сводятся к '.', десятичная запятая → точка:
    «105, 121, 135» и «105/121/135» → один токен, «2.6.4115-25» — свой)."""
    out = set()
    for m in TOKEN.finditer(text):
        t = NORM.sub("", m.group(0)).strip(".,:–-/")
        if not t or not any(c.isdigit() for c in t):
            continue
        t = re.sub(r"[,/]", ".", t)
        out.add(t)
    return out


MARK = re.compile(r"(?iu)^(уровень|шаг|этап|ступень|пункт|раздел|часть)\s*\.?\s*\d{1,2}$")


def slide_tokens(path):
    """(strict, visible): токены авторского набора (без плейсхолдеров, без
    номерных маркеров «01 », бейджей и подписей-меток «Уровень 1») + дополненный
    набор со всеми цифровыми метками — для обратной сверки «есть ли значение из
    facts.md вообще на слайдах»."""
    found, visible_add = {}, set()
    prs = Presentation(str(path))

    def walk(shapes, i):
        for sh in shapes:
            if sh.shape_type == 6:
                walk(sh.shapes, i)
                continue
            if sh.is_placeholder or not sh.has_text_frame:
                continue
            for p in sh.text_frame.paragraphs:
                txt = "".join(r.text for r in p.runs)
                if re.fullmatch(r"\s*\d{1,2}\s*", txt) or MARK.match(txt.strip()):
                    visible_add.update(tokens(txt))       # метка: бейдж / «Уровень N»
                    continue
                txt = re.sub(r"^\s*\d{1,2}\s", "", txt)   # маркер списка «01 …»
                for t in tokens(txt):
                    found.setdefault(t, (i, txt.strip()[:60]))
    for i, s in enumerate(prs.slides, 1):
        walk(s.shapes, i)
    return found, set(found) | visible_add


COUNT = re.compile(r"\d+[\s\-]*(пункт|вопрос|шаг|уровен|групп|колонк|строк|значен|элем)\w*", re.I)


def value_tokens(cell):
    """Токены ячейки «значения»: мета-счётчики («10 пунктов», «9 шагов») не в счёт —
    это про объём списка, а не про значения на слайде."""
    return tokens(COUNT.sub(" ", cell))


def parse_facts(path):
    """rows = [(значение, статус|None, токены)]; из markdown-таблицы. Колонки
    гибко: статус ищем keyword'ом, значение — колонки левее статуса (формы
    «значение|статус|источник» и «полоса|значения|статус|источник»), колонку-
    метку полосы («1 · Прил. 1.1») и токены по отдельным ячейкам (без склейки
    «14 | 5 уровней» в «145»)."""
    rows = []
    for line in Path(path).read_text(encoding="utf-8-sig").splitlines():
        if not line.strip().startswith("|"):
            continue
        cells = [c.strip() for c in line.strip().strip("|").split("|")]
        if len(cells) < 2 or all(re.fullmatch(r":?-{2,}:?", c or "-") for c in cells):
            continue
        if cells[0].lower() in ("значение", "полоса", "value", "число"):
            continue                              # строка заголовков таблицы
        s_idx = next((i for i, c in enumerate(cells)
                      if any(s in c.lower() for s in STATUSES)), None)
        val_cells = cells[:s_idx] if s_idx else cells[:-1]
        val_cells = [c for c in val_cells if not re.match(r"^\d+\s*[·•]", c)]
        val = " | ".join(val_cells)
        toks = set()
        for c in val_cells:
            toks |= value_tokens(c)
        status = None
        if s_idx is not None:
            raw = cells[s_idx].lower()
            status = next((s for s in STATUSES if s in raw), None)
        rows.append((val, status, toks))
    return rows


def main(pptx, facts, strict=False):
    facts_rows = parse_facts(facts)
    if not facts_rows:
        print(f"facts: в {facts} не найдено строк таблицы «значение | статус | …»")
        return 1
    counts = {s: 0 for s in STATUSES}
    counts[None] = 0
    facts_tokens = set()
    for val, status, toks in facts_rows:
        counts[status] += 1
        facts_tokens |= toks
    slide, visible = slide_tokens(pptx)

    unlisted = {t: v for t, v in slide.items() if t not in facts_tokens}
    absent = sorted(facts_tokens - visible)

    print(f"facts.md: значений {len(facts_rows)}  "
          f"(sourced {counts['sourced']}, derived {counts['derived']}, "
          f"common_knowledge {counts['common_knowledge']}, "
          f"estimate {counts['estimate']}, missing {counts['missing']}"
          + (f", БЕЗ СТАТУСА {counts[None]}" if counts[None] else "") + ")")
    if unlisted:
        print(f"\nНА СЛАЙДАХ, но НЕ в facts.md — {len(unlisted)}:")
        for t, (i, frag) in list(unlisted.items())[:40]:
            print(f"  {t:>12s}  стр.{i} «{frag}»")
        if len(unlisted) > 40:
            print(f"  … и ещё {len(unlisted) - 40}")
    else:
        print("\nна слайдах всё заявлено в facts.md")
    if absent:
        print(f"\nв facts.md, но НЕ найдено на слайдах — {len(absent)}: "
              + ", ".join(absent[:40]) + (" …" if len(absent) > 40 else ""))

    em = [v for v, s, _ in facts_rows if s in ("estimate", "missing")]
    print("\n— приёмочная строка (часть про источники подтверждает агент) —")
    print(f"ПРОВЕРКА ДАННЫХ: {len(facts_rows)} значений — sourced/derived "
          f"{counts['sourced'] + counts['derived']}, common_knowledge "
          f"{counts['common_knowledge']}, estimate {counts['estimate']}, "
          f"missing {counts['missing']}; список estimate/missing: "
          + ("; ".join(em) if em else "пуст")
          + "; все указанные на слайдах источники присутствуют среди "
            "материалов: <да/нет>")
    return 1 if (strict and unlisted) else 0


if __name__ == "__main__":
    argv = sys.argv[1:]
    if "--facts" not in argv:
        sys.exit(__doc__)
    i = argv.index("--facts")
    fx = argv[i + 1]
    del argv[i:i + 2]
    sys.exit(main(argv[0], fx, strict="--strict" in argv))
