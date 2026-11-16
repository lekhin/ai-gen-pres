# -*- coding: utf-8 -*-
"""Декларативный план презентации → pptx по шаблону активной темы.

Сборка:  python deck.py build plan.json [-o out.pptx]
План (JSON, лежит в <Проект>\\_build\\plan.json) отделяет ДАННЫЕ слайдов от
кода: правки итераций — это дифф JSON, а не python. Экзотические слайды
по-прежнему допустимо писать кодом (build.py между метками) — гибрид.

Схема плана:
{
  "theme": "minimal",                   # opt; minimal|dark|academic (см. themes\\);
                                        #   перебивается DECK_THEME и --theme CLI
  "template": "_service\\themes\\dark\\template.pptx",  # opt; явный шаблон поверх
                                        #   темы (относ. путь ищется от корня
                                        #   фабрики, затем от папки плана)
  "output": "Имя_Проекта\\Имя_Проекта.pptx",  # opt; иначе <имя_плана>.pptx рядом
  "slides": [
    {
      "layout": "content",          # title | content | image
      "header": "НАЗВАНИЕ-В-ШАПКЕ", # opt (дубль заголовка)
      "title": "ЗАГОЛОВОК",         # opt → плейсхолдер idx0
      "subtitle": "Подзаголовок",   # opt → idx1, если не задан — drop_ph(1)
      "date": true,                 # opt, поле даты в подвале
      "background": "black",        # opt: тёмная плашка + перекраска колонтитулов
      "blocks": [ … ]
    }
  ]
}

Блоки (координаты: "col" 0..5 + "span" по сетке ИЛИ явные x/y/w/h в см;
y по умолчанию = CONTENT_TOP, см):
  text     {"paras":[{"text"|"runs":[{text,size,font,bold,italic,color}],
                     lvl,align,space_after,line,bullet}]  или "text":"строки\\nпо\\nсмайлику"}
           поля size/font/color на блоке — значения по умолчанию для абзацев.
           "line" — МНОЖИТЕЛЬ интерлиньяжа (1.2 = 120%), не пункты: 26 = 26×!
           "space_after" — уже в пунктах (pt).
  stat     {"value":"155","label":"тыс. т","note":"…","size":48,"color":"red"}
  chart    {"kind":"bar|hbar|line|pie|doughnut","cats":[…],
            "series":[{"name":"…","values":[…]}],  или [[…значения…]],
            "num_fmt":"0","labels":true,"legend":null,"y_max":null,"colors":["red",…]}
  image    {"path":"gantt.png","w":null,"caption":"подпись (Inter semibold italic)"}
           путь относительно файла плана; PNG deck_mpl вставляется 1:1 по w
  table    {"rows":[["год","2024"],…],"header":true,"widths":[см,…]}
  columns  {"cols":[{"span":3,"gap":0.3,"blocks":[…вложенные…]}, …]}
  rect     {"fill":"gray","line":null,"w_pt":1}  — плашка-карточка
  rule     {"color":"gray","w_pt":1}        — горизонтальная линия;
           "w"/col+span — длина линии в см, ТОЛЩИНА только в "w_pt" (pt)
  вертикальная линия — rect малой ширины: {"w":0.06,"h":…,"fill":"black"}
  diagram  {"groups":[{"x","y","w","h","label","dash","line_w","label_align"}],
            "nodes":[{"id","x","y","w","h","title","body","role","shape",
                      "title_size","body_size"}],
            "edges":[{"from","to","from_side","to_side","points":[[x,y],…],
                      "dash","w_pt","label":{"text","x","y","w","size","align"}}],
            "labels":[{"x","y","w","text","size","align"}],
            "legend":{"x","y","items":[["solid"|"dash","текст"]],"gap"}}
           роль: app|data|model|people|plain|planned;
           форма: rect|pill|db|diamond|io; координаты см; движок — pptx_diagram.
           Простые схемы декларативно; разветвлённые маршруты — "points".
Цвета: имена из палитры темы (accent ink paper line muted good bad + любые
собственные ключи theme.json) + nav1..navN + исторические алиасы
(red black white gray royal amethyst apple) + hex "2563EB".
Шрифты: "head" | "body" | "cg" | "inter" | явное имя ("Georgia").
"""
import json
import os
import sys
from pathlib import Path

from pptx.util import Pt, Cm
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

import deck_theme
from deck_paths import resolve_asset, repo_root
from pptx_helpers import (BLACK, WHITE, GRAY, RED, INK, ACCENT, LINE, MUTED,
                          NAV, CG, INTER, HEAD, BODY, SLIDE_W, SLIDE_H, COL_W,
                          GRID_W, MARGIN, CONTENT_TOP, FOOTER_TOP, R, chart,
                          chart_image, col_x, drop_ph, hline, layouts,
                          load_template, new_slide, ph, place, rect,
                          recolor_ph, send_to_back, set_text, textbox)

THEME = deck_theme.active()


def _build_colors():
    """Имя → RGBColor: палитра темы + nav1..N + исторические алиасы."""
    out = {}
    for k, v in THEME["palette"].items():
        if isinstance(v, str):
            out[k] = RGBColor.from_string(v.upper())
        elif isinstance(v, list):
            for i, c in enumerate(v, 1):
                out[f"nav{i}"] = RGBColor.from_string(c.upper())
    ds = THEME.get("dark_slide") or {}
    if ds.get("background"):
        out["dark_bg"] = RGBColor.from_string(ds["background"].upper())
    out.update({"red": ACCENT, "black": BLACK, "white": WHITE, "gray": GRAY,
                "accent": ACCENT, "ink": INK, "line": LINE, "muted": MUTED})
    if NAV:
        out.setdefault("royal", NAV[0])
        out.setdefault("amethyst", NAV[1] if len(NAV) > 1 else INK)
        out.setdefault("apple", NAV[2] if len(NAV) > 2 else LINE)
        out.setdefault("cinnabar", ACCENT)
    return out


COLORS = _build_colors()
LOGO = THEME.get("logo")
LOGO = resolve_asset(LOGO) if LOGO else None   # путь к логотипу темы (обычно нет)
FONTS = {"head": HEAD, "body": BODY, "cg": HEAD, "inter": BODY}
ALIGNS = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER,
          "right": PP_ALIGN.RIGHT, "justify": PP_ALIGN.JUSTIFY}
ANCHORS = {"top": MSO_ANCHOR.TOP, "middle": MSO_ANCHOR.MIDDLE,
           "bottom": MSO_ANCHOR.BOTTOM}


def color(v, default=BLACK):
    if v is None:
        return default
    c = COLORS.get(str(v).lower())
    if c is None:
        try:
            c = RGBColor.from_string(str(v).lstrip("#").upper())
        except (ValueError, TypeError):
            raise ValueError(f"неизвестный цвет {v!r}")
    return c


def font(v, default=INTER):
    return FONTS.get(str(v).lower(), v or default)


def geom(b, y_default=CONTENT_TOP, h_default=None):
    """Геометрия блока: col/span по сетке или явные x/y/w/h (см)."""
    if "col" in b or "span" in b:
        x = col_x(b.get("col", 0))
        w = COL_W * b.get("span", 1)
    elif "x" in b and "w" in b:
        x, w = b["x"], b["w"]
    else:
        raise KeyError(
            f"блок {b.get('type', '?')!r}: нет координат — задайте 'col' (0..5) "
            f"+ по желанию 'span' по сетке ИЛИ 'x' и 'w' в см; ключи блока: "
            f"{sorted(k for k in b if k != 'type')}; схема — докстринг plan.py / SKILL.md")
    y = b.get("y", y_default)
    h = b.get("h", h_default if h_default is not None else FOOTER_TOP - y - 0.6)
    return x, y, w, h


def _runs(par, block):
    """Список run-кортежей для set_text.R из JSON-спецификации абзаца."""
    base_size = par.get("size", block.get("size", 14))
    base_font = font(par.get("font", block.get("font", "inter")))
    base_color = color(par.get("color", block.get("color")), BLACK)
    rs = par.get("runs") or [{"text": par.get("text", "")}]
    return [(r.get("text", ""), font(r.get("font", base_font)),
             r.get("size", base_size), r.get("bold", False),
             r.get("italic", False), color(r.get("color"), base_color))
            for r in rs]


def render_text(slide, b, base_y=CONTENT_TOP):
    x, y, w, h = geom(b, base_y)
    tb = textbox(slide, x, y, w, h, anchor=ANCHORS.get(b.get("anchor"), MSO_ANCHOR.TOP))
    if b.get("text") and "paras" not in b:
        paras = [{"text": t} for t in str(b["text"]).split("\n")]
    else:
        paras = b.get("paras", [])
    items = [{"runs": _runs(p, b),
              "lvl": p.get("lvl", 0),
              "align": ALIGNS.get(p.get("align")),
              "space_after": p.get("space_after"),
              "line": p.get("line"),
              "bullet": p.get("bullet")}
             for p in paras]
    set_text(tb, items)
    return tb


def render_stat(slide, b, base_y=CONTENT_TOP):
    x, y, w, h = geom(b, base_y, 4.0)
    tb = textbox(slide, x, y, w, h)
    items = [{"runs": [(str(b.get("value", "")), CG, b.get("size", 48), True, False,
                        color(b.get("color"), RED))],
              "space_after": 0, "bullet": False}]
    if b.get("label"):
        items.append({"runs": [(b["label"], CG, b.get("label_size", 14), False,
                                False, color(b.get("label_color"), BLACK))],
                      "space_after": 2, "bullet": False})
    if b.get("note"):
        items.append({"runs": [(b["note"], INTER, 10, False, False,
                                color(b.get("note_color"), BLACK))],
                      "bullet": False})
    set_text(tb, items)
    return tb


def render_chart(slide, b, base_y=CONTENT_TOP):
    x, y, w, h = geom(b, base_y, 10.5)
    series = []
    for s in b.get("series", []):
        if isinstance(s, dict):
            series.append((s.get("name", ""), tuple(s["values"])))
        else:
            series.append(("Серия", tuple(s)))
    colors = [color(c) for c in b["colors"]] if b.get("colors") else None
    return chart(slide, b["kind"], x, y, w, h, b.get("cats", []), series,
                 num_fmt=b.get("num_fmt", "General"),
                 y_max=b.get("y_max"), y_min=b.get("y_min"),
                 colors=colors, legend=b.get("legend"),
                 labels=b.get("labels", True))


def render_image(slide, b, base_y=CONTENT_TOP, plan_dir=None):
    p = Path(b["path"])
    if not p.is_absolute() and plan_dir:
        p = plan_dir / p
    x, y, w, h = geom(b, base_y)
    if b.get("h") is not None and "x" in b:      # полный ручной режим
        pic = chart_image(slide, p, x, y, w=b.get("w", w), h=b["h"])
    else:  # пропорции PNG (deck_mpl = 1:1 к ячейке) сохраняются
        pic = chart_image(slide, p, x, y, w=b.get("w", w))
    if b.get("caption"):
        cap_y = y + pic.height / 360000.0 + 0.2
        tb = textbox(slide, x, cap_y, w, 1.0)
        set_text(tb, [{"runs": [R(b["caption"], 10, INTER, False, True)],
                       "bullet": False}])
    return pic


def render_table(slide, b, base_y=CONTENT_TOP):
    x, y, w, h = geom(b, base_y)
    rows, cols_n = b["rows"], len(b["rows"][0])
    gfx = slide.shapes.add_table(rows=len(rows), cols=cols_n,
                                 left=Cm(x), top=Cm(y), width=Cm(w), height=Cm(h))
    t = gfx.table
    t.first_row = False
    t.horz_banding = False
    for i, cw in enumerate(b.get("widths", [])):
        t.columns[i].width = Cm(cw)
    for ri, row in enumerate(rows):
        for ci, val in enumerate(row):
            cell = t.cell(ri, ci)
            cell.margin_left = cell.margin_right = Cm(0.2)
            cell.margin_top = cell.margin_bottom = Cm(0.1)
            if ri == 0 and b.get("header", True):
                cell.fill.solid(); cell.fill.fore_color.rgb = BLACK
                run = (str(val), CG, 11, True, False, WHITE)
            else:
                cell.fill.solid()
                cell.fill.fore_color.rgb = WHITE if ri % 2 else GRAY
                run = (str(val), INTER, b.get("size", 11), False, False, BLACK)
            set_text(cell, [{"runs": [run],
                             "align": ALIGNS.get(b.get("align", "left" if ci == 0 else "right")),
                             "bullet": False}])
    return gfx


def render_rect(slide, b, base_y=CONTENT_TOP):
    x, y, w, h = geom(b, base_y)
    return rect(slide, x, y, w, h,
                fill=color(b["fill"]) if b.get("fill") else None,
                line=color(b["line"]) if b.get("line") else None,
                line_w=b.get("w_pt", 1.0))


def render_rule(slide, b, base_y=CONTENT_TOP):
    x, y, w, _ = geom(b, base_y, 0)
    # толщина — только в w_pt; ключ "w" у rule означает длину (геометрию),
    # иначе w=30.87 превращал линию в полосу 31pt толщиной
    return hline(slide, x, y, w, color=color(b.get("color"), GRAY),
                 w_pt=b.get("w_pt", 1.0))


def render_columns(slide, b, base_y=CONTENT_TOP, plan_dir=None):
    """Колонки: span'ы суммируются ≤6, gap между колонками (см); ширина
    колонки = (наборная полоса − gap'ы), распределённая пропорционально span."""
    gap = b.get("gap", 0.3)
    total_span = sum(c.get("span", 1) for c in b["cols"])
    unit = (GRID_W - gap * (len(b["cols"]) - 1)) / total_span
    cur = MARGIN
    for cb in b["cols"]:
        w = unit * cb.get("span", 1)
        for inner in cb.get("blocks", []):
            ib = dict(inner)
            ib.pop("col", None)
            ib["x"], ib["w"] = cur, w
            ib.setdefault("y", base_y)
            render_block(slide, ib, base_y, plan_dir)
        cur += w + gap


def _node_anchor(n, side, other=None):
    """Точка входа/выхода связи у грани узла (side=auto — грань к другому узлу)."""
    x, y, w, h = n["x"], n["y"], n["w"], n["h"]
    if side in (None, "auto"):
        if other is not None:
            dx = (other["x"] + other["w"] / 2) - (x + w / 2)
            dy = (other["y"] + other["h"] / 2) - (y + h / 2)
            side = ("right" if dx >= 0 else "left") if abs(dx) >= abs(dy) \
                else ("bottom" if dy >= 0 else "top")
        else:
            side = "bottom"
    return {"left": (x - 0.03, y + h / 2), "right": (x + w + 0.03, y + h / 2),
            "top": (x + w / 2, y - 0.03), "bottom": (x + w / 2, y + h + 0.03)}[side]


def render_diagram(slide, b, base_y=CONTENT_TOP):
    from pptx_diagram import node as _n, edge as _e, edge_label as _l, \
        group as _g, legend as _lg
    nodes = {n.get("id", "n%d" % i): n for i, n in enumerate(b.get("nodes", []))}
    for g in b.get("groups", []):
        _g(slide, g["x"], g["y"], g["w"], g["h"], g.get("label"),
           dash=g.get("dash"), line_w=g.get("line_w", 1.5),
           label_align=ALIGNS.get(g.get("label_align"), PP_ALIGN.RIGHT))
    for n in b.get("nodes", []):
        _n(slide, n["x"], n["y"], n["w"], n["h"], n.get("title", ""), n.get("body"),
           role=n.get("role", "app"), shape=n.get("shape", "rect"),
           title_size=n.get("title_size", 10), body_size=n.get("body_size", 8))
    for e in b.get("edges", []):
        nf, nt = nodes.get(e.get("from")), nodes.get(e.get("to"))
        pts = [tuple(p) for p in e.get("points", [])]
        if nf is not None:
            a = _node_anchor(nf, e.get("from_side", "auto"), nt)
            pts = [a] + pts if pts else [a]
        if nt is not None:
            z = _node_anchor(nt, e.get("to_side", "auto"), nf)
            pts = pts + [z] if pts else ([z] if nf is None else pts + [z])
        if len(pts) >= 2:
            _e(slide, pts, dash=e.get("dash"), w_pt=e.get("w_pt", 1.25))
        lb = e.get("label")
        if lb:
            _l(slide, lb["x"], lb["y"], lb.get("w", 3.0), lb["text"],
               size=lb.get("size", 6.5), align=ALIGNS.get(lb.get("align"), PP_ALIGN.LEFT))
    for lb in b.get("labels", []):
        _l(slide, lb["x"], lb["y"], lb.get("w", 3.0), lb["text"],
           size=lb.get("size", 6.5), align=ALIGNS.get(lb.get("align"), PP_ALIGN.LEFT))
    lg = b.get("legend")
    if lg:
        _lg(slide, lg["x"], lg["y"], [tuple(i) for i in lg["items"]],
            gap=lg.get("gap", 5.9))


def render_block(slide, b, base_y=CONTENT_TOP, plan_dir=None):
    kind = b.get("type")
    if kind == "text":
        return render_text(slide, b, base_y)
    if kind == "stat":
        return render_stat(slide, b, base_y)
    if kind == "chart":
        return render_chart(slide, b, base_y)
    if kind == "image":
        return render_image(slide, b, base_y, plan_dir)
    if kind == "table":
        return render_table(slide, b, base_y)
    if kind == "rect":
        return render_rect(slide, b, base_y)
    if kind == "rule":
        return render_rule(slide, b, base_y)
    if kind == "columns":
        return render_columns(slide, b, base_y, plan_dir)
    if kind == "diagram":
        return render_diagram(slide, b, base_y)
    raise ValueError(f"неизвестный тип блока {kind!r}")


def dark_bg(slide, logo=None):
    """Тёмная подложка + перекраска колонтитулов (цвета dark_slide из theme.json;
    в самой тёмной теме — просто чуть более глубокий фон).
    logo — путь к файлу логотипа темы: кладётся на белую подложку (тёмный фон)."""
    ds = THEME.get("dark_slide") or {}
    bg = RGBColor.from_string(ds.get("background", "111827").upper())
    fg = RGBColor.from_string(ds.get("text", "F9FAFB").upper())
    send_to_back(rect(slide, 0, 0, SLIDE_W, SLIDE_H, fill=bg))
    if logo:
        rect(slide, 0.4, 0.31, 5.39, 2.72,
             fill=RGBColor.from_string(ds.get("plate", "FFFFFF").upper()))
        slide.shapes.add_picture(str(logo), Cm(0.50), Cm(0.41), Cm(5.19), Cm(2.52))
    for idx in (10, 11, 12):
        p = ph(slide, idx)
        if p is not None:
            recolor_ph(p, fg)


DARK_FG = RGBColor.from_string(
    (THEME.get("dark_slide") or {}).get("text", "F9FAFB").upper())


def apply_plan(prs, plan, plan_dir=None):
    L_TITLE, L_CONT, L_IMG = layouts(prs)
    lmap = {"title": L_TITLE, "content": L_CONT, "image": L_IMG}
    for n, sd in enumerate(plan.get("slides", []), 1):
        try:
            layout = lmap[sd.get("layout", "content")]
        except KeyError:
            raise ValueError(f"слайд {n}: layout {sd['layout']!r} (нужен title/content/image)")
        s = new_slide(prs, layout, header=sd.get("header"), with_date=sd.get("date", True))
        if sd.get("background") in ("black", "dark"):
            dark_bg(s, LOGO if sd.get("logo", True) and LOGO else None)
        dark = bool(sd.get("background"))
        t = ph(s, 0)
        if sd.get("title"):
            t.text_frame.text = sd["title"]
            if dark:
                # runs обычного текста наследуют цвет из layout (recolor_ph их
                # не находит, когда явного rPr нет) — задаём цвет каждому run
                for p_ in t.text_frame.paragraphs:
                    for r_ in p_.runs:
                        r_.font.color.rgb = DARK_FG
        elif layout is not L_TITLE:
            drop_ph(s, 0)   # пустой title-плейсхолдер = заглушка в PowerPoint
        sub = ph(s, 1)
        if sd.get("subtitle"):
            sub.text_frame.text = sd["subtitle"]
            if dark:
                c = DARK_FG if sd.get("subtitle_color", "white") == "white" else MUTED
                for p_ in sub.text_frame.paragraphs:
                    for r_ in p_.runs:
                        r_.font.color.rgb = c
        else:
            drop_ph(s, 1)
        p2 = ph(s, 2)   # правый бокс «с рисунком»: пустой = заглушка в PowerPoint
        if p2 is not None and p2.has_text_frame and not p2.text_frame.text.strip():
            drop_ph(s, 2)
        for blk in sd.get("blocks", []):
            render_block(s, blk, CONTENT_TOP, plan_dir)
    return prs


def build(plan_path, out=None, root=None):
    """Собрать pptx из плана. root — корень фабрики (для template/output);
    по умолчанию ищется вверх от плана, вне такой фабрики — папка плана и
    её предки (разрешение относительных путей — deck_paths.resolve_asset).

    Тема: поле "theme" плана; deck.py выставляет DECK_THEME по плану ДО импорта
    helper'ов. Здесь сверяем: импорт helper'ов уже заморозил активную тему —
    при рассогласовании предупреждаем и берём шаблон из темы плана."""
    plan_path = Path(plan_path)
    plan = json.loads(plan_path.read_text(encoding="utf-8-sig"))  # BOM (notepad/PS) допустим
    plan_dir = plan_path.resolve().parent
    theme = plan.get("theme")
    env_theme = os.environ.get("DECK_THEME")
    theme = env_theme or theme or deck_theme.DEFAULT_THEME
    if theme != THEME["name"]:
        print(f"WARN: тема плана {theme!r}, но helper'ы импортированы под "
              f"{THEME['name']!r}: шаблон будет из {theme!r}, а цвета/шрифты кода — "
              f"из {THEME['name']!r}. Собирай через `deck.py build` (он ставит "
              f"DECK_THEME до импорта) или задай os.environ['DECK_THEME'] заранее.")
    if root is None:
        root = repo_root(plan_dir) or plan_dir
    root = Path(root)
    if plan.get("template"):
        tpl = resolve_asset(plan["template"], root, plan_dir)
    else:
        tpl = deck_theme.template_path(theme)
    if out:
        out = Path(out)
        if not out.is_absolute():
            out = root / out
    elif plan.get("output"):
        out = resolve_asset(plan["output"], root, plan_dir)
    else:
        out = plan_dir / (plan_path.stem + ".pptx")
    prs = load_template(tpl)
    apply_plan(prs, plan, plan_dir=plan_path.parent)
    out.parent.mkdir(parents=True, exist_ok=True)
    prs.save(out)
    return out
