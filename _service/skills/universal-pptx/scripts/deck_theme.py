# -*- coding: utf-8 -*-
"""Темы оформления презентаций — дизайн-контур фабрики (палитра/шрифты/сетка).

Тема = папка `_service\\themes\\<имя>\\` с двумя артефактами:
  * theme.json   — палитра, шрифты, имена layout'ов, сетка, настройки QA;
  * template.pptx — носитель стилей (мастер + layout'ы), генерируется
    deck_make_template.py из theme.json (пересобирать после правки theme.json).

Выбор активной темы (по приоритету):
  1. переменная окружения DECK_THEME;
  2. поле "theme" в plan.json (deck.py build подхватывает его сам);
  3. тема по умолчанию — minimal.

ВНИМАНИЕ: константы pptx_helpers/deck_mpl/pptx_diagram вычисляются при ИМПОРТЕ
из активной темы. Значит тему надо задать ДО первого импорта helper'ов:
  * через CLI — `deck.py build/mmd/qa --theme dark …` (ставит DECK_THEME);
  * в кодовом build.py проекта — строка `os.environ["DECK_THEME"] = "…"`
    до импорта pptx_helpers (в заготовке _service\\build.py есть метка).
Перезагрузки модулей не предусмотрено — это осознанно, импорт один на процесс.
"""
import json
import os

from deck_paths import service_file

DEFAULT_THEME = "minimal"
_cache = {}


def themes_dir():
    return service_file("themes")


def list_themes():
    d = themes_dir()
    return sorted(p.parent.name for p in d.glob("*/theme.json"))


def load(name=None):
    """Загрузить (и закэшировать) theme.json темы по имени."""
    name = name or os.environ.get("DECK_THEME") or DEFAULT_THEME
    if name not in _cache:
        path = themes_dir() / name / "theme.json"
        if not path.exists():
            raise FileNotFoundError(
                f"тема {name!r} не найдена: ждём {path}; доступны: "
                f"{', '.join(list_themes()) or '(нет)'}")
        t = json.loads(path.read_text(encoding="utf-8-sig"))
        t["name"] = name
        t["dir"] = str(path.parent)
        _cache[name] = t
    return _cache[name]


def active():
    return load()


def template_path(name=None):
    t = load(name)
    return os.path.join(t["dir"], t.get("template", "template.pptx"))


def inspect_dump_path(name=None):
    t = load(name)
    return os.path.join(t["dir"], t.get("inspect", "inspect.md"))


def qa_allowlists(name=None):
    """(разрешённые цвета hex, разрешённые шрифты) для предупреждений qa.

    name=None — объединение ВСЕХ тем: `qa` без знания темы не должен
    заваливать предупреждениями корректные сборки других тем."""
    if name:
        themes = [load(name)]
    else:
        themes = [load(n) for n in list_themes()] or [load(DEFAULT_THEME)]
    colors, fonts = set(), set()
    for t in themes:
        p = t["palette"]
        for v in p.values():
            if isinstance(v, str):
                colors.add(v.upper())
            elif isinstance(v, list):
                colors.update(c.upper() for c in v)
        for d in (t.get("dark_slide") or {},):
            for v in d.values():
                if isinstance(v, str):
                    colors.add(v.upper())
        fonts.update(t.get("qa_fonts") or t["fonts"].values())
    return colors, fonts
