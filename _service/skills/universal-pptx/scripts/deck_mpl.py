# -*- coding: utf-8 -*-
"""Стиль matplotlib в дизайне активной темы: палитра из theme.json, шрифты из
_service\\fonts, размеры ровно под сетку слайда (figsize в см → PNG вставляется
на слайд 1:1).

Использование (скрипт-генератор графики лежит в _build\\ проекта):
    from deck_mpl import figure, save
    fig, ax = figure(15.4, 10.5)          # размер места на слайде, см
    ax.barh(...); ax.set_yticks(...)      # обычная ось matplotlib
    save(fig, "gantt.png")                # 300 dpi, прозрачный фон
Вставка на слайд — pptx_helpers.chart_image(slide, "gantt.png", x, y, w).
PNG сохраняем БЕЗ bbox_inches="tight": figsize = финальный размер, любой
tight меняет пропорции и текст на слайде растянется.
"""
import matplotlib

matplotlib.use("Agg")
from matplotlib import font_manager  # noqa: E402
import matplotlib.pyplot as plt  # noqa: E402

from deck_paths import service_file  # noqa: E402
from deck_theme import active as _active_theme  # noqa: E402

FONTS_DIR = service_file("fonts")  # папка fonts в служебной папке, поиск — deck_paths.py

_T = _active_theme()
_P = _T["palette"]
# hex-строки (matplotlib не принимает RGBColor из pptx_helpers!)
ACCENT = "#" + _P["accent"].upper()
INK = "#" + _P["ink"].upper()
PAPER = "#" + _P["paper"].upper()
LINE = "#" + _P["line"].upper()
NAV = ["#" + c.upper() for c in _P.get("nav", [])]
# алиасы старой кодовой базы
RED, BLACK, WHITE, GRAY = ACCENT, INK, PAPER, LINE
ROYAL, AMETHYST, APPLE = (NAV + [INK, LINE, LINE])[:3]
# цикл цветов (акцент первым, далее навигационная палитра темы)
CYCLE = [ACCENT] + NAV + [INK, LINE]

_fonts_done = False


def setup_fonts():
    """Зарегистрировать шрифты темы из папки _service\\fonts (идемпотентно).
    Variable-Inter matplotlib видит как статичный Regular — для выделений
    используйте weight='bold' только если убедились, что вес реально есть,
    иначе выделяйте размером/цветом."""
    global _fonts_done
    if _fonts_done:
        return
    for f in sorted(FONTS_DIR.glob("*.ttf")):
        try:
            font_manager.fontManager.addfont(str(f))
        except Exception:
            pass  # кривой/переменный файл — оси, просто не зарегистрируется
    _fonts_done = True


_RC = {
    "font.family": _T["fonts"]["body"],
    "font.size": 10,                      # на слайде = 10 pt (figsize 1:1)
    "text.color": BLACK,
    "axes.edgecolor": BLACK,
    "axes.linewidth": 0.75,
    "axes.labelcolor": BLACK,
    "axes.labelsize": 10,
    "axes.titlecolor": BLACK,
    "axes.grid": True,
    "axes.axisbelow": True,
    "grid.color": GRAY,
    "grid.linewidth": 0.5,
    "axes.spines.top": False,
    "axes.spines.right": False,
    "xtick.color": BLACK,
    "xtick.labelsize": 9,
    "ytick.color": BLACK,
    "ytick.labelsize": 9,
    "axes.prop_cycle": plt.cycler(color=CYCLE),
    "figure.facecolor": "none",
    "savefig.facecolor": "none",
    "savefig.transparent": True,
    "savefig.dpi": 300,
    "legend.frameon": False,
    "legend.fontsize": 9,
}


def style():
    """Применить rcParams темы к plt (для кода, создающего fig самому)."""
    setup_fonts()
    plt.rcParams.update(_RC)


def figure(w_cm, h_cm, **kw):
    """Фигура размером w×h САНТИМЕТРОВ (ровно ячейка слайда), стиль темы."""
    setup_fonts()
    plt.rcParams.update(_RC)
    fig, ax = plt.subplots(figsize=(w_cm / 2.54, h_cm / 2.54), **kw)
    fig.set_layout_engine("constrained")
    return fig, ax


def save(fig, path):
    """Сохранить PNG 300 dpi с прозрачным фоном; размер = figsize 1:1."""
    fig.savefig(path, dpi=300, transparent=True)
    plt.close(fig)
    return path
