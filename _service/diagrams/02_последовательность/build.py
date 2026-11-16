# -*- coding: utf-8 -*-
r"""Эталон 02: диаграмма последовательности (вход через SSO, паттерн BFF).

Содержание — mermaid sequenceDiagram из docs/ARCHITECTURE.md ai-chat, раздел 3.
Показывает: sequence() c самоо́бходом… (сжат для слайда), note, alt-ветвление,
колонку-пояснение. Пересборка: python build.py → example.pptx рядом.
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
from pptx_diagram import sequence, caption                 # noqa: E402

SRC = str(deck_theme.template_path())
OUT = os.path.join(HERE, "example.pptx")

prs = load_template(SRC)
L_TITL, L_CONT, L_IMG = layouts(prs)

s = new_slide(prs, L_CONT, header="Вход через SSO")
ph(s, 0).text_frame.text = "Вход через SSO: OAuth2 Code + PKCE"
drop_ph(s, 1)

actors = [
    {"id": "U", "name": "Браузер", "sub": "сотрудник"},
    {"id": "A", "name": "Бэкенд (BFF)", "sub": "FastAPI"},
    {"id": "K", "name": "Keycloak", "sub": "сервер SSO"},
    {"id": "D", "name": "Postgres"},
]

items = [
    {"from": "U", "to": "A", "text": "GET /auth/login?next=/chat/ (через nginx, TLS)"},
    {"from": "A", "to": "U", "text": "302 на Keycloak + cookie state/PKCE", "style": "dash"},
    {"from": "U", "to": "K", "text": "/authorize (code + S256, клиент: веб-приложение)"},
    {"note": "Аутентификация корпоративным SSO; роль кладётся в access-токен",
     "over": ["K", "K"]},
    {"from": "K", "to": "U", "text": "302 /auth/callback?code&state", "style": "dash"},
    {"from": "U", "to": "A", "text": "callback (code, state)"},
    {"from": "A", "to": "K", "text": "POST /token (code + verifier + secret)"},
    {"from": "K", "to": "A", "text": "access 15 мин + refresh", "style": "dash"},
    {"alt": [
        {"when": "роль пилотной группы подтверждена", "messages": [
            {"from": "A", "to": "D", "text": "сессия на сервере: токены в БД, JIT-пользователь"},
            {"from": "A", "to": "U", "text": "302 next + session-cookie (httpOnly, 8–12 ч)", "style": "dash"},
        ]},
        {"when": "роли нет и заявка не подтверждена", "messages": [
            {"from": "A", "to": "U", "text": "страница заявки в фокус-группу", "style": "dash"},
        ]},
    ]},
]

sequence(s, 1.5, 5.6, 23.9, actors, items, msg_step=0.64, msg_size=6.5)

caption(s, 26.2, 5.75, 6.1, 8.5,
        [("Токены не покидают серверов. ", True),
         ("Браузер получает только session-cookie (httpOnly, secure). Проверка "
          "подписи, iss, azp и роли — на бэкенде по публичным ключам SSO. "
          "Дальнейшие вызовы — только по cookie; токены SSO живут в серверной БД.",
          False)])

prs.save(OUT)
print("saved:", OUT)
