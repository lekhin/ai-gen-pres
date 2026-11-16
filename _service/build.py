# -*- coding: utf-8 -*-
r"""Заготовка скрипта сборки презентации (python-pptx) — общая, без контента.

Как использовать: скопировать в служебную подпапку <Имя_Проекта>\_build\, задать
PROJECT, ТЕМУ и OUT, написать слайды между метками НАЧАЛО/КОНЕЦ, запустить
отовсюду (папка скилла и служебная папка с темами/шрифтами ищутся сами;
переопределение — env DECK_SKILL_SCRIPTS и DECK_SERVICE_DIR, см. deck_paths.py).
Типовые слайды дешевле собирать декларативным plan.json (`deck.py build`),
этот шаблон — для экзотики и гибрида.

Хелперы: _service\skills\universal-pptx\scripts\pptx_helpers.py (палитра темы,
сетка, new_slide с живыми колонтитулами, set_text, rect/hline/vline/textbox,
send_to_back/recolor_ph/drop_ph, chart/chart_image). Грабли и рецепты: SKILL.md
того же скилла. Книжную/«полосную» вёрстку (тема book, A4 портрет) собирать не от
этой заготовки, а сразу через pptx_book + pptx_metrics (SKILL.md «Книжная
вёрстка»; рабочий пример — `scripts\smoke_book.py` (4 полосы: чек-листы, список,
лента, пирамида), регрессия — тот же smoke + qa).
"""
import os
import sys

# build.py лежит в <папка проекта>\_build\; PROJECT_DIR — папка проекта (на уровень выше)
HERE = os.path.dirname(os.path.abspath(__file__))                    # <проект>\_build
PROJECT_DIR = os.path.dirname(HERE)                                   # <проект>

# ----- ТЕМА: ОБЯЗАТЕЛЬНО задать ДО импорта pptx_helpers (палитра вычисляется
# при импорте). minimal | dark | academic | book | своя из _service\themes\ ----
os.environ.setdefault("DECK_THEME", "minimal")


def _find_scripts_dir():
    """Папка скриптов скилла (работает при установке скилла в любое место):
    env DECK_SKILL_SCRIPTS → поиск вверх отсюда (канонический путь
    _service\\skills\\universal-pptx\\scripts и просто skills\\universal-pptx\\scripts)
    → фолбэк каноническая раскладка от PROJECT_DIR."""
    env = os.environ.get("DECK_SKILL_SCRIPTS")
    if env and os.path.isdir(env):
        return env
    d = HERE
    while True:
        for rel in (os.path.join("_service", "skills", "universal-pptx", "scripts"),
                    os.path.join("skills", "universal-pptx", "scripts")):
            c = os.path.join(d, rel)
            if os.path.isdir(c):
                return c
        parent = os.path.dirname(d)
        if parent == d:
            c = os.path.join(os.path.dirname(PROJECT_DIR),
                             "_service", "skills", "universal-pptx", "scripts")
            if os.path.isdir(c):
                return c
            raise RuntimeError(
                "папка скриптов скилла universal-pptx не найдена при поиске вверх "
                "от %s; укажи её переменной окружения DECK_SKILL_SCRIPTS" % HERE)
        d = parent


sys.path.insert(0, _find_scripts_dir())

from pptx.util import Cm, Pt
from pptx.enum.text import PP_ALIGN

import deck_theme                                                    # noqa: E402
from pptx_helpers import (  # noqa: E402
    ACCENT, INK, PAPER, LINE, MUTED, HEAD, BODY,
    BLACK, WHITE, RED, GRAY, CG, INTER,          # исторические алиасы
    SLIDE_W, SLIDE_H, MARGIN, GRID_W, COL_W, col_x, CONTENT_TOP,
    load_template, layouts, new_slide, ph, place,
    set_text, textbox, rect, hline, vline, R,
    send_to_back, recolor_ph, drop_ph,
)

# ----- конфигурация проекта ------------------------------------------------
PROJECT = "Имя_Проекта"                    # <- переименовать
DARK_SLIDES = False   # тёмные подложки в светлой теме: только по запросу заказа
SRC = deck_theme.template_path()           # template.pptx активной темы
OUT = os.path.join(PROJECT_DIR, PROJECT + ".pptx")   # в корень папки проекта

prs = load_template(SRC)
L_TITL, L_CONT, L_IMG = layouts(prs)


def dark_slide():
    """Тёмный титульный/финальный слайд (цвета dark_slide из theme.json):
    плашка под плейсхолдеры + перекраска колонтитулов + белизна заголовка —
    по run'ам (у plain-текста плейсхолдера явного rPr нет)."""
    from pptx.dml.color import RGBColor
    ds = deck_theme.active().get("dark_slide") or {}
    bg = RGBColor.from_string(ds.get("background", "111827").upper())
    fg = RGBColor.from_string(ds.get("text", "F9FAFB").upper())
    s = new_slide(prs, L_TITL, header=None, with_date=True)
    send_to_back(rect(s, 0, 0, SLIDE_W, SLIDE_H, fill=bg))
    for idx in (10, 12):
        recolor_ph(ph(s, idx), fg)
    return s, fg


# ==================== НАЧАЛО: СЛАЙДЫ ПРОЕКТА ====================

# --- светлый титул ---
s = new_slide(prs, L_TITL, header=None)
ph(s, 0).text_frame.text = "Заголовок презентации"
ph(s, 1).text_frame.text = "Подзаголовок / контекст"

# --- контентный слайд: заголовок + текст ---
s = new_slide(prs, L_CONT, header="Короткий заголовок")   # шапка-дубль
ph(s, 0).text_frame.text = "Заголовок слайда"             # стиль — из layout темы
b = ph(s, 1)
place(b, MARGIN, CONTENT_TOP, GRID_W, 11.4)
set_text(b.text_frame, [
    {"runs": [R("Тезис. ", 15, bold=True), R("Основной текст слайда.", 15)],
     "bullet": False, "line": 1.2, "space_after": 12},
    {"runs": [R("Второй абзац. Ключевые слова — жирным, акцент — цветом ACCENT.", 15)],
     "bullet": False, "line": 1.2},
])

# --- тёмный слайд: по умолчанию ВЫКЛЮЧЕН (DARK_SLIDES=False) --------------
# Включать, если тёмный тон запрошен заказом или выбрана тема dark.
if DARK_SLIDES:
    s, fg = dark_slide()
    t = ph(s, 0)
    place(t, 4.2, 5.6, 25.47, 3.1)
    set_text(t.text_frame, [
        {"runs": [R("Тёмный финальный слайд", 32, HEAD, True, color=fg)],
         "align": PP_ALIGN.CENTER, "bullet": False},
    ])

# ==================== КОНЕЦ: СЛАЙДЫ ПРОЕКТА ======================

# --- зачистка пустых плейсхолдеров-заглушек («Текст слайда» в PowerPoint) ---
for s in prs.slides:
    for idx in (1, 2):
        p = ph(s, idx)
        if p is not None and p.has_text_frame and not p.text_frame.text.strip():
            drop_ph(s, idx)

os.makedirs(os.path.dirname(OUT), exist_ok=True)
prs.save(OUT)
print("saved:", OUT, "slides:", len(prs.slides._sldIdLst))
