# -*- coding: utf-8 -*-
"""Поиск ресурсов фабрики презентаций (папки themes/, fonts/, скриптов скилла)
без привязки к месту установки скилла. Работает, если скилл лежит в каноническом
месте (<корень>\\_service\\skills\\universal-pptx\\) или установлен в любую
другую папку.

Порядок поиска служебной папки `_service` (или её аналога):
  1. переменная окружения DECK_SERVICE_DIR (явный указатель — основной способ
     для установки скилла вне репозитория);
  2. вверх от расположения самих скриптов (каноническая раскладка: скилл живёт
      внутри _service\\skills\\…);
  3. вверх от текущей папки (запуск из репозитория);
  4. фолбэк — канонический parents[3].

Переменные окружения:
  DECK_SERVICE_DIR — папка с themes/ / fonts/ / build.py;
  DECK_THEME       — активная тема (см. deck_theme.py).
"""
import os
from pathlib import Path

ENV_SERVICE_DIR = "DECK_SERVICE_DIR"


def _walk_up(start):
    """start и все его предки (ближайшие первыми)."""
    start = Path(start).resolve()
    return [start, *start.parents]


def _is_service_dir(d):
    return (d / "themes").is_dir() or (d / "fonts").is_dir()


def service_dir():
    """Служебная папка с темами/шрифтами (см. докстринг модуля)."""
    env = os.environ.get(ENV_SERVICE_DIR)
    if env:
        return Path(env).expanduser().resolve()
    for start in (Path(__file__).resolve().parent, Path.cwd()):
        for d in _walk_up(start):
            if _is_service_dir(d):
                return d
    child = service_dir_child()            # раскладка «…/_service/…» от предков
    if child:
        return child
    return Path(__file__).resolve().parents[3]   # каноническая раскладка


def service_file(name):
    return service_dir() / name


def artcraft_cli(app):
    """Путь к *-cli.exe из комплекта ArtCraft (github.com/storytold) или None.
    Приложения комплекта имеют headless-режимы для агентов: deckcraft-cli рендерит
    .pptx в PNG без Office за доли секунды, pdfcraft-cli/wordcraft-cli читают
    pdf/docx. Ищем: PATH → env DECK_ARTCRAFT_DIR → «C:\\Program Files\\ArtCraft»."""
    import shutil as _sh
    found = _sh.which(f"{app}-cli")
    if found:
        return found
    roots = [Path(os.environ["DECK_ARTCRAFT_DIR"])] if os.environ.get("DECK_ARTCRAFT_DIR") \
        else [Path(r"C:\Program Files\ArtCraft"), Path.home() / "AppData" / "Local" / "ArtCraft"]
    for root in roots:
        for cand in (root / app.capitalize() / f"{app}-cli.exe",
                     root / app / f"{app}-cli.exe",
                     root / app.capitalize() / app, root / app / app):
            if cand.exists():
                return str(cand)
    return None


def skill_scripts():
    """Папка скриптов скилла (для sys.path в копиях build.py в _build\\ проектов;
    обычно не нужна — deck_paths сам регистрируется при импорте)."""
    return Path(__file__).resolve().parent


scripts_dir = skill_scripts()


def service_dir_child():
    """Служебная папка как Дочерняя `_service` ближайшего предка cwd/скриптов —
    раскладки «work/_service/…» без прямой связи по родительским цепочкам."""
    for start in (Path(__file__).resolve().parent, Path.cwd()):
        for d in _walk_up(start):
            c = d / "_service"
            if _is_service_dir(c):
                return c
    return None


def repo_root(start=None):
    """Ближайший предок, содержащий папку `_service` (корень фабрики), или None —
    вне такой фабрики."""
    for d in _walk_up(start or Path.cwd()):
        if (d / "_service").is_dir():
            return d
    return None


def resolve_asset(rel, *bases):
    """Относительный путь из плана → на диск. Пробуется по очереди: каждый
    base как есть и его предки (существование), затем предки с отсечённым
    префиксом `_service/` (план ссылается от корня фабрики), затем существующая
    папка-первый-компонент (файл выхода ещё не создан), затем фолбэк bases[0]."""
    p = Path(str(rel))
    if p.is_absolute():
        return p
    dirs = []
    for b in bases:
        if b:
            dirs.extend(_walk_up(b))
    for d in dirs:
        if (d / p).exists():
            return d / p
    parts = p.parts
    if parts and parts[0].lower() == "_service":
        parts = parts[1:] or parts       # "_service/…" без содержимого не бывает
    svc = service_dir()
    for d in [svc, *dirs]:
        c = d.joinpath(*parts)
        if c.exists() or (parts and (d / parts[0]).is_dir()):
            return c
    return (dirs[0] if dirs else svc) / p
