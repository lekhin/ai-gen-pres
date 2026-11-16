# -*- coding: utf-8 -*-
r"""Эталон 04: схема подключений сервисов — mermaid → pptx (слайд 16:9).

Показывает сквозной путь: mermaid-исходник (diagram.mmd) разбирается
mmd2diagram.py (ранговая раскладка, subgraph→периметры, формы/роли узлов,
пунктир, метки связей в чистых зонах) и рендерится движком pptx_diagram.
Заказчик правит результат в PowerPoint как нативные фигуры.

Пересборка:  python build.py  → example.pptx рядом
(см. _service\diagrams\README.md). Тонкая правка меток после пересборки —
обычное дело, см. SKILL.md «Схемы».
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
            raise RuntimeError("_service не найден при поиске вверх от %s" % HERE)
        d = parent


SERVICE = os.environ.get("DECK_SERVICE_DIR") or _find_service()
os.environ.setdefault("DECK_SERVICE_DIR", SERVICE)
os.environ.setdefault("DECK_SKILL_SCRIPTS",
                      os.path.join(SERVICE, "skills", "universal-pptx", "scripts"))
sys.path.insert(0, os.environ["DECK_SKILL_SCRIPTS"])

import deck_theme
from deck_paths import service_file        # noqa: E402
from pptx_helpers import (load_template, layouts, new_slide, ph, drop_ph)  # noqa: E402
from plan import render_diagram                            # noqa: E402
import mmd2diagram                                         # noqa: E402

SRC = str(deck_theme.template_path())
OUT = os.path.join(HERE, "example.pptx")
MMD = os.path.join(HERE, "diagram.mmd")
LEGEND = [("solid", "запросы и данные"), ("dash", "аутентификация и LLM")]

text = open(MMD, encoding="utf-8-sig").read()
m, routes, leg = mmd2diagram.prepare(text, LEGEND)
block = mmd2diagram.to_block(m, routes, leg)
for w in m.warnings:
    print("WARN:", w)

prs = load_template(SRC)
_, L_CONT, _ = layouts(prs)
s = new_slide(prs, L_CONT, header="Схема подключений")
ph(s, 0).text_frame.text = "Схема подключений сервисов"
drop_ph(s, 1)
render_diagram(s, block)

for s in prs.slides:
    for idx in (1, 2):
        p = ph(s, idx)
        if p is not None and p.has_text_frame and not p.text_frame.text.strip():
            drop_ph(s, idx)

prs.save(OUT)
print("saved:", OUT)
