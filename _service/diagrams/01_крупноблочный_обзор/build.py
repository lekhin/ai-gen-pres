# -*- coding: utf-8 -*-
r"""Эталон 01: крупноблочный обзор системы для руководителей (слайд 16:9).

Собран библиотекой pptx_diagram.py. Показывает: периметры (group: сплошной
контейнер + пунктирная зона), роли узлов (people/data/model/app), метки связей
в зазорах, «ёлочку» в обход блока, двунаправленный пунктир OIDC, легенду,
вывод-подпись для руководителей.

Пересборка:  python build.py   → example.pptx рядом (см. _service\diagrams\README.md).
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

from pptx.enum.text import PP_ALIGN                        # noqa: E402

import deck_theme
from deck_paths import service_file        # noqa: E402
from pptx_helpers import (load_template, layouts, new_slide, ph, drop_ph)  # noqa: E402
from pptx_diagram import (node, edge, edge_label, group,   # noqa: E402
                          caption, legend, row_tops)

SRC = str(deck_theme.template_path())
OUT = os.path.join(HERE, "example.pptx")

prs = load_template(SRC)
L_TITL, L_CONT, L_IMG = layouts(prs)

s = new_slide(prs, L_CONT, header="Крупноблочный обзор")
ph(s, 0).text_frame.text = "Крупноблочный обзор системы"
drop_ph(s, 1)

# сетка: две колонки стенда, три ряда
C1, C2 = 2.2, 10.8
W1, W2 = 5.9, 6.4
R1, R2, R3 = row_tops(8.75, 3, 2.05, 0.8)
BH = 2.05
X1, X2 = C1 + W1 / 2, C2 + W2 / 2

# акторы и периметры
node(s, 1.5, 5.85, 9.6, 1.35, "Сотрудники предприятия", "пилотная группа · вход по SSO",
     role="people", shape="pill", title_size=10, body_size=7.5)
group(s, 1.5, 7.6, 19.5, 9.8, label="Стенд приложения — docker compose (внутренний контур)")
group(s, 23.3, 7.7, 9.07, 9.7, label="Внутренний контур · вне стенда",
      dash="dash", line_w=1.0, label_align=PP_ALIGN.CENTER)

# узлы
node(s, C1, R1, W1, BH, "Веб-интерфейс", "Nuxt 4: чат, аудио, презентации")
node(s, C2, R1, W2, BH, "Файловое хранилище", "документы, аудио, pptx · TTL-очистка",
     role="data")
node(s, C1, R2, W1, BH, "Бэкенд приложения", "FastAPI: SSO-доступ, история, файлы, записи, презентации")
node(s, C2, R2, W2, BH, "Шлюз моделей", "подключение моделей, квоты, учёт расходов",
     role="model")
node(s, C1, R3, W1, BH, "Postgres", "история чатов, записи, заявки, ключи, расходы",
     role="data")
node(s, C2, R3, W2, BH, "Генератор презентаций", "LibreOffice + фабрика корпоративных шаблонов")
node(s, 23.8, 8.5, 8.1, 2.2, "Корпоративная аутентификация", "Keycloak + Active Directory")
node(s, 23.8, 11.2, 8.1, 3.0, "LLM-эндпоинт · OpenAI-совместимый",
     "чат-модель + расшифровка речи (ASR)", role="model", body_size=8.5)
caption(s, 23.8, 14.7, 8.1, 2.4,
        [("Данные не покидают периметр: ", True),
         ("промпты, файлы и аудио обрабатываются моделями внутреннего контура.", False)])

# связи (метки — в зазорах, не на рамках)
edge(s, [(X1, 7.2), (X1, R1)])
edge_label(s, 1.85, 7.82, 3.1, "HTTPS · единая точка входа", size=7, align=PP_ALIGN.RIGHT)
edge(s, [(X1, R1 + BH), (X1, R2)])
edge(s, [(X1, R2 + BH), (X1, R3)])
edge(s, [(C1 + W1, R2 + BH / 2), (C2 - 0.05, R2 + BH / 2)])
edge_label(s, 8.15, 11.82, 2.55, "запросы к модели", align=PP_ALIGN.CENTER)
edge(s, [(7.5, R2 + BH), (12.55, R3)])
edge_label(s, 5.3, 13.78, 2.2, "задания на генерацию")
edge(s, [(X2, R3), (X2, R2 + BH + 0.02)])
edge_label(s, 14.35, 13.78, 3.0, "запросы к модели")
edge(s, [(C2 + W2, R3 + BH / 2), (18.9, R3 + BH / 2),
         (18.9, R1 + BH / 2), (C2 + W2 + 0.05, R1 + BH / 2)])
edge(s, [(7.6, R2), (12.9, R1 + BH + 0.02)])
edge(s, [(C2 + W2, R2 + BH / 2 + 0.05), (23.75, 12.7)])
edge_label(s, 19.15, 12.08, 4.3, "OpenAI-совместимый API", align=PP_ALIGN.CENTER)
edge(s, [(5.5, R2), (23.75, 10.5)], dash="dash")
edge_label(s, 19.15, 9.95, 4.4, "вход и роли (OIDC)")

legend(s, 2.2, 17.0, [("solid", "запросы и данные"), ("dash", "аутентификация (OIDC)")])

for s in prs.slides:
    for idx in (1, 2):
        p = ph(s, idx)
        if p is not None and p.has_text_frame and not p.text_frame.text.strip():
            drop_ph(s, idx)

prs.save(OUT)
print("saved:", OUT)
