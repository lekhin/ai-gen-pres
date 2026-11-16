# -*- coding: utf-8 -*-
r"""Эталон 03: алгоритм (флоучарт) с ветвлением и возвратом.

Процесс генерации презентаций из ai-chat: очередь заданий → сборка → QA-гейт
(возврат «нет») → выдача. Показывает: flow() с ромбом решения и петлей no_to,
формы start/end/io, колонку-пояснение. Пересборка: python build.py.
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))


def _find_service():
    d = HERE
    while True:
        if os.path.basename(d) == "_service":
            return d
        parent = os.path.dirname(d)
        if parent == d:
            raise RuntimeError("_service не найден при поиске вверх")
        d = parent


SERVICE = os.environ.get("DECK_SERVICE_DIR") or _find_service()
os.environ.setdefault("DECK_SERVICE_DIR", SERVICE)
os.environ.setdefault("DECK_SKILL_SCRIPTS",
                      os.path.join(SERVICE, "skills", "universal-pptx", "scripts"))
sys.path.insert(0, os.environ["DECK_SKILL_SCRIPTS"])

import deck_theme
from deck_paths import service_file        # noqa: E402
from pptx_helpers import load_template, layouts, new_slide, ph, drop_ph  # noqa: E402
from pptx_diagram import flow, caption                     # noqa: E402

SRC = str(deck_theme.template_path())
OUT = os.path.join(HERE, "example.pptx")

prs = load_template(SRC)
L_TITL, L_CONT, L_IMG = layouts(prs)

s = new_slide(prs, L_CONT, header="Генерация презентаций")
ph(s, 0).text_frame.text = "Алгоритм генерации презентаций"
drop_ph(s, 1)

steps = [
    {"kind": "start", "text": "Заявка пользователя на pptx", "sub": "контекст чата + корпоративный шаблон"},
    {"text": "Постановка в очередь заданий (Postgres)"},
    {"text": "Сборка: фабрика корпоративных шаблонов + LibreOffice", "sub": "запросы к модели через шлюз"},
    {"kind": "dec", "text": "QA-гейт: pptx открывается, аудит зелёный?", "no_to": 2},
    {"text": "Файл в хранилище (TTL), ссылка в чате", "sub": "автоочистка медиа по сроку хранения"},
    {"kind": "end", "text": "Готово"},
]

flow(s, 10.0, 5.8, steps, w=8.0)

caption(s, 20.5, 5.9, 11.8, 9.0,
        [("Очередь в БД, не в памяти. ", True),
         ("Задания на генерацию ставит бэкенд в таблицу очереди — воркер "
          "переживает перезапуски и держит одну тяжёлую сборку за раз.\n",
          False),
         ("QA-гейт обязательный. ", True),
         ("Готовый pptx проверяется программно (пакет, плейсхолдеры, текст); "
          "красный аудит — возврат на пересборку, файл пользователю не отдаётся.\n",
          False),
         ("Модель — только через шлюз. ", True),
         ("У воркера персональный виртуальный ключ с лимитом и списком "
          "разрешённых моделей (квоты и учёт расходов — в шлюзе).", False)])

prs.save(OUT)
print("saved:", OUT)
