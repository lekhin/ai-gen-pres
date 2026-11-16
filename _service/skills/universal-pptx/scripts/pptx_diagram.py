# -*- coding: utf-8 -*-
"""Схемы нативными фигурами PowerPoint: узлы, связи, контейнеры, метки, легенда.

Надстройка над pptx_helpers (палитра/шрифты/сетка наследуются отт же). Цель —
схема на слайде остаётся РЕДАКТИРУЕМОЙ в PowerPoint (как chart()), а координаты
считаются один раз в декларативной части build-скрипта, а не в метрах и миллиметрах.

Типовой приём (эталон — _service\\diagrams\\01_крупноблочный_обзор):
    from pptx_diagram import *
    g = group(s, 1.5, 7.6, 19.5, 9.8, label="Стенд chk0-svt-ai1 — docker compose")
    fe  = node(s, C1, R1, W, H, "Веб-интерфейс", "Nuxt 4: чат, аудио, презентации")
    db  = node(s, C1, R3, W, H, "Postgres", "история чатов, заявки", role="data", shape="db")
    edge(s, [(X1, yA), (X1, yB)])                      # прямая стрелка
    edge(s, [(x1, y1), (18.9, y1), (18.9, y2), (x2, y2)])   # обход блока «ёлочкой»
    edge_label(s, x, y, w, "запросы к модели")
    legend(s, 2.2, 17.0, [("solid", "запросы и данные"), ("dash", "аутентификация")])

Семантические роли окраски (цвета — строго палитра активной темы):
    app     обычный компонент (фон/текст темы)     data   хранилище (серая плашка)
    model   путь модели (рамка акцентным цветом)   people люди/внешние акторы (тёмн. плашка)
    plain   без фона (подложки под текст)         planned целевое/планируемое (пунктирная рамка)

Грабли (проверено, см. также SKILL.md «Схемы»):
- python-pptx не умеет наконечники стрелок — только XML a:tailEnd/a:headEnd
  (стрелка — свойство линии, не фигуры);
- «ёлочку» вокруг препятствий собирайте явным списком точек: авто-elbow
  ломается на вертикальных/горизонтальных парах точек (begin_x == end_x);
- метку ребра ставьте в межблочный зазор (>= 2 см по горизонтали) или внутрь
  контейнера, но НЕ на рамки — на рендере это первый «дефект глазами»;
- одиночные пересечения пунктира с сплошными допустимы, сплошная-со-сплошной — нет.
"""
from pptx.util import Cm, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.dml import MSO_LINE
from pptx.oxml.ns import qn

from pptx_helpers import BLACK, RED, WHITE, GRAY, INTER, set_text, textbox, rect, R

# ---------------- наконечники и стили линий ----------------
SHAPE_PRESETS = {
    "rect": MSO_SHAPE.RECTANGLE,
    "pill": MSO_SHAPE.RECTANGLE,      # заменяется на ROUNDED_RECTANGLE ниже
    "db": MSO_SHAPE.CAN,
    "diamond": MSO_SHAPE.DIAMOND,
    "io": MSO_SHAPE.PARALLELOGRAM,
}

# роль -> стилизация плашки (см. docstring). role задает дефолты, их можно
# переопределить аргументами fill/line/dash node().
ROLES = {
    "app":    dict(fill=WHITE, line=BLACK, line_w=1.25, tc=BLACK, bc=BLACK),
    "data":   dict(fill=GRAY,  line=BLACK, line_w=1.25, tc=BLACK, bc=BLACK),
    "model":  dict(fill=WHITE, line=RED,   line_w=1.75, tc=BLACK, bc=BLACK),
    "people": dict(fill=BLACK, line=None,  line_w=None, tc=WHITE, bc=WHITE),
    "plain":  dict(fill=None,  line=None,  line_w=None, tc=BLACK, bc=BLACK),
    "planned": dict(fill=WHITE, line=BLACK, line_w=1.25, tc=BLACK, bc=BLACK, dash="dash"),
}


def arrowhead(conn, end="tail", typ="triangle"):
    """Наконечник на коннекторе: end = 'tail' (конец) или 'head' (начало)."""
    ln = conn.line._get_or_add_ln()
    tag = "a:tailEnd" if end == "tail" else "a:headEnd"
    ln.append(ln.makeelement(qn(tag), {"type": typ, "w": "med", "len": "med"}))
    return conn


def _style_line(fmt, color=BLACK, w_pt=1.25, dash=None):
    fmt.color.rgb = color
    fmt.width = Pt(w_pt)
    if dash:
        fmt.dash_style = {"dash": MSO_LINE.DASH, "dot": MSO_LINE.ROUND_DOT,
                          "dashdot": MSO_LINE.DASH_DOT}[dash]


# ---------------- узлы ----------------
def node(slide, x, y, w, h, title, body=None, role="app", shape="rect",
         fill=None, line=None, line_w=None, dash=None, title_color=None,
         body_color=None, title_size=10, body_size=8):
    """Плашка-узел: заголовок (Inter bold, центр) + подпись. role — см. ROLES."""
    st = dict(ROLES[role])
    if fill is not None:
        st["fill"] = fill
    if line is not None:
        st["line"] = line
    if line_w is not None:
        st["line_w"] = line_w
    if dash is not None:
        st["dash"] = dash
    if title_color is not None:
        st["tc"] = title_color
    if body_color is not None:
        st["bc"] = body_color

    ms = SHAPE_PRESETS[shape]
    if shape == "pill":
        ms = MSO_SHAPE.ROUNDED_RECTANGLE
    sp = slide.shapes.add_shape(ms, Cm(x), Cm(y), Cm(w), Cm(h))
    sp.shadow.inherit = False
    if shape == "pill":
        sp.adjustments[0] = 0.5
    if st["fill"] is None:
        sp.fill.background()
    else:
        sp.fill.solid()
        sp.fill.fore_color.rgb = st["fill"]
    if st["line"] is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = st["line"]
        sp.line.width = Pt(st["line_w"])
        if st.get("dash"):
            sp.line.dash_style = {"dash": MSO_LINE.DASH, "dot": MSO_LINE.ROUND_DOT,
                                  "dashdot": MSO_LINE.DASH_DOT}[st["dash"]]

    tf = sp.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = tf.margin_right = Cm(0.5 if shape == "diamond" else 0.18)
    tf.margin_top = tf.margin_bottom = Cm(0.08)
    items = [{"runs": [R(title, title_size, INTER, True, color=st["tc"])],
              "align": PP_ALIGN.CENTER, "bullet": False, "space_after": 2}]
    if body:
        items.append({"runs": [R(body, body_size, INTER, color=st["bc"])],
                      "align": PP_ALIGN.CENTER, "bullet": False, "line": 1.05})
    set_text(tf, items)
    return sp


# ---------------- связи ----------------
def edge(slide, points, color=BLACK, w_pt=1.25, dash=None, head=True,
         tail=False):
    """Ломаная-связь по списку точек [(x, y), ...] в см. Стрелка — на последнем
    звене (head); tail=True добавляет стрелку на первое (двунаправленная)."""
    conns = []
    for (x1, y1), (x2, y2) in zip(points, points[1:]):
        c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                       Cm(x1), Cm(y1), Cm(x2), Cm(y2))
        _style_line(c.line, color, w_pt, dash)
        c.shadow.inherit = False
        conns.append(c)
    if conns:
        if head:
            arrowhead(conns[-1])
        if tail:
            arrowhead(conns[0], end="head")
    return conns


def _text_h_cm(text, size, w_cm):
    """Оценка высоты текста (для обтягивающих метку боксов — тогда qa может
    честно проверять «линия сквозь текст», а не по пустому боксту)."""
    cpl = max(6, w_cm * 28.35 / (0.5 * size))
    lines = max(1, -(-len(text) // int(cpl)))
    return lines * size * 1.3 / 28.35 + 0.06


def _text_w_cm(text, size):
    return len(text) * 0.5 * size / 28.35 * 1.15


def edge_label(slide, x, y, w, text, size=6.5, align=PP_ALIGN.LEFT, bold=False,
               color=BLACK):
    """Подпись ребра (бокс обтягивает текст по высоте). Ставить в межблочный
    зазор >= 2 см или внутрь контейнера, не на рамки (см. docstring модуля)."""
    tb = textbox(slide, x, y, w, _text_h_cm(text, size, w))
    set_text(tb.text_frame, [{"runs": [R(text, size, INTER, bold, color=color)],
                              "align": align, "bullet": False, "line": 1.05}])
    return tb


# ---------------- контейнеры ----------------
def group(slide, x, y, w, h, label=None, dash=None, line=BLACK, line_w=1.5,
          label_align=PP_ALIGN.RIGHT, label_size=9.5, label_inset=0.15):
    """Контур группы (подсистема/периметр/зона доверия). label — ярлык внутри
    у верхней грани: справа (по умолчанию) или слева (label_align=LEFT).
    Вложенные группы — обычные group() внутри координат внешней."""
    sp = rect(slide, x, y, w, h, fill=None, line=line, line_w=line_w)
    if dash:
        sp.line.dash_style = {"dash": MSO_LINE.DASH, "dot": MSO_LINE.ROUND_DOT,
                              "dashdot": MSO_LINE.DASH_DOT}[dash]
    if label:
        # ярлык — бокс по ширине текста (не на всю грань), иначе qa ложно
        # флагует каждую связь, входящую в верхнюю рамку контейнера
        lw = min(w - 2 * label_inset, _text_w_cm(label, label_size) + 0.2)
        if label_align == PP_ALIGN.RIGHT:
            tx = x + w - label_inset - lw
        elif label_align == PP_ALIGN.CENTER:
            tx = x + (w - lw) / 2
        else:
            tx = x + label_inset
        tb = textbox(slide, tx, y + 0.15, lw, _text_h_cm(label, label_size, lw))
        set_text(tb.text_frame, [{"runs": [R(label, label_size, INTER, True)],
                                  "align": label_align, "bullet": False}])
    return sp


def caption(slide, x, y, w, h, runs, size=8.5, align=PP_ALIGN.LEFT, line=1.15):
    """Пояснение/вывод рядом со схемой: runs=[(текст, bold?), ...] или строка."""
    tb = textbox(slide, x, y, w, h)
    if isinstance(runs, str):
        runs = [(runs, False)]
    set_text(tb.text_frame, [{"runs": [R(t, size, INTER, b) for (t, b) in runs],
                              "align": align, "bullet": False, "line": line}])
    return tb


# ---------------- легенда и сетка ----------------
def legend(slide, x, y, items, gap=5.9):
    """Легенда связей: items = [("solid"|"dash", "текст"), ...] горизонтально."""
    cx = x
    for kind, text in items:
        edge(slide, [(cx, y), (cx + 1.2, y)], dash="dash" if kind == "dash" else None,
             head=True)
        edge_label(slide, cx + 1.35, y - 0.28, gap - 1.35, text, size=7)
        cx += gap


def row_tops(y0, n, h, gap):
    """Верхние края n рядов блоков высотой h с зазором gap."""
    return [y0 + i * (h + gap) for i in range(n)]


def col_lefts(x0, n, w, gap):
    """Левые края n колонок шириной w с зазором gap."""
    return [x0 + i * (w + gap) for i in range(n)]


# ---------------- последовательность (UML sequence) ----------------
def _note_h(text, size, nw):
    return _text_h_cm(text, size, nw - 0.2) + 0.24


def _note_w(it, axes):
    a, b = it.get("over", [])
    xa, xb = axes.get(a, axes.get(b)), axes.get(b, axes.get(a))
    return max(4.0, abs(xb - xa))


def _seq_h(items, step, axes, note_size):
    """Предварительный подсчёт высоты содержимого секвенции."""
    h = 0.0
    for it in items:
        if "alt" in it:
            h += 0.5
            for p in it["alt"]:
                h += 0.5 + _seq_h(p.get("messages", []), step, axes, note_size) + 0.5
            h += 0.2
        elif "note" in it:
            h += _note_h(it["note"], note_size, _note_w(it, axes)) + 0.12
        else:
            h += step
    return h


def _seq_render(slide, axes, items, y, step, msg_size, note_size):
    """Прогон списка элементов вниз от y; возвращает итоговый y."""
    for it in items:
        if "alt" in it:
            top = y
            sp = list(axes.values())[0]
            ep = list(axes.values())[-1]
            x1, x2 = min(sp, ep) - 0.3, max(sp, ep) + 0.3
            y += 0.5
            for j, p in enumerate(it["alt"]):
                if j:
                    edge(slide, [(x1, y), (x2, y)], dash="dash", w_pt=0.75, head=False)
                    edge_label(slide, x1 + 0.15, y + 0.04, x2 - x1 - 0.3,
                               p.get("when", "иначе"), size=note_size, bold=True)
                    y += 0.5
                else:
                    edge_label(slide, x1 + 0.15, y - 0.42, 2.5,
                               p.get("when", ""), size=note_size, bold=True)
                y = _seq_render(slide, axes, p.get("messages", []), y, step,
                                msg_size, note_size)
                y += 0.5
            rect(slide, x1, top, x2 - x1, y - top, fill=None, line=BLACK, line_w=0.75)
            y += 0.2
        elif "note" in it:
            nw = _note_w(it, axes)
            a, b = it.get("over", [])
            xa, xb = axes.get(a, axes.get(b)), axes.get(b, axes.get(a))
            nx = min(xa, xb) if abs(xb - xa) > 1 else xa - nw / 2
            nh = _note_h(it["note"], note_size, nw)
            rect(slide, nx, y, nw, nh, fill=GRAY, line=None)   # подложка — под текст
            tb = textbox(slide, nx + 0.1, y + 0.1, nw - 0.2, nh - 0.16)
            set_text(tb.text_frame, [{"runs": [R(it["note"], note_size, INTER, False, True)],
                                      "bullet": False, "line": 1.05}])
            y += nh + 0.12
        else:
            xa, xb = axes[it["from"]], axes[it["to"]]
            yy = y + step / 2
            if it["from"] == it["to"]:
                edge(slide, [(xa, yy - step * 0.28), (xa + 1.1, yy - step * 0.28),
                             (xa + 1.1, yy + step * 0.22), (xa + 0.03, yy + step * 0.22)])
                edge_label(slide, xa + 1.25, yy - 0.3, 4.5,
                           it.get("text", ""), size=msg_size)
            else:
                edge(slide, [(xa, yy), (xb, yy)],
                     dash="dash" if it.get("style") == "dash" else None)
                lx = min(xa, xb)
                lw = abs(xb - xa) - 0.2
                lh = _text_h_cm(it.get("text", ""), msg_size, lw)
                edge_label(slide, lx + 0.1, yy - lh - 0.1, lw,
                           it.get("text", ""), size=msg_size, align=PP_ALIGN.CENTER)
            y += step
    return y


def sequence(slide, x, y, w, actors, items, actor_h=1.15, msg_step=0.78,
             actor_size=9, sub_size=7, msg_size=6.5, note_size=6.5):
    """Диаграмма последовательности: actors=[{id,name,sub?}], items — сообщения
    {"from","to","text",style?"dash"}, самоо́бходы from==to, {"note","over":[a,b]},
    ветвление {"alt":[{"when":"…","messages":[…]}, …]}. Линии жизни рисуются
    до конца контента; возвращает итоговую высоту блока (см)."""
    n = len(actors)
    axes = {}
    aw = min(4.6, w / n - 0.5)
    for i, a in enumerate(actors):
        cx = x + w * (i + 0.5) / n
        axes[a["id"]] = cx
    total = _seq_h(items, msg_step, axes, note_size)
    bottom = y + actor_h + 0.15 + total + 0.25
    for a in actors:                     # линии жизни — под всё остальное
        edge(slide, [(axes[a["id"]], y + actor_h + 0.05), (axes[a["id"]], bottom)],
             dash="dot", w_pt=0.75, head=False)
    for a in actors:
        cx = axes[a["id"]]
        node(slide, cx - aw / 2, y, aw, actor_h, a["name"], a.get("sub"),
             title_size=actor_size, body_size=sub_size)
    _seq_render(slide, axes, items, y + actor_h + 0.2, msg_step, msg_size, note_size)
    return bottom - y


# ---------------- флоучарт (линейный + возврат «нет») ----------------
def flow(slide, cx, y, steps, w=6.4, h=1.05, gap=0.85, dec_h=1.7, loop_off=1.7,
         title_size=9, body_size=7.5, yes="да", no="нет", label_size=6.5):
    """Вертикальный алгоритм: steps=[{kind:"start|proc|io|dec|end", text, sub?,
    no_to? — индекс шага, куда возврат с подписью «нет»}]. Стрелки вниз — «да»
    после ромбов. Возвращает итоговый y (см)."""
    boxes = []
    yy = y
    for st in steps:
        kind = st.get("kind", "proc")
        if kind == "dec":
            wd = w + 2.6
            sp = node(slide, cx - wd / 2, yy, wd, dec_h, st["text"], st.get("sub"),
                      shape="diamond", title_size=title_size, body_size=body_size)
            boxes.append((cx - wd / 2, yy, wd, dec_h, kind))
            yy += dec_h
        else:
            shape = {"start": "pill", "end": "pill", "io": "io"}.get(kind, "rect")
            ww = w * 0.8 if kind in ("start", "end") else w
            node(slide, cx - ww / 2, yy, ww, h, st["text"], st.get("sub"),
                 shape=shape, title_size=title_size, body_size=body_size)
            boxes.append((cx - ww / 2, yy, ww, h, kind))
            yy += h
        yy += gap
    for i, (bx, by, bw, bh, kind) in enumerate(boxes):
        st = steps[i]
        if i + 1 < len(boxes):
            nx, ny = boxes[i + 1][0], boxes[i + 1][1]
            edge(slide, [(cx, by + bh), (cx, ny - 0.03)])
            if kind == "dec":
                edge_label(slide, cx + 0.12, by + bh + (ny - by - bh) / 2 - 0.15,
                           1.2, yes, size=label_size)
        if st.get("no_to") is not None:
            tx, ty, tw, th, _ = boxes[st["no_to"]]
            ry = by + bh / 2 if kind != "dec" else by + bh / 2
            rx = bx + bw if kind != "dec" else cx + bw / 2
            loop_x = max(bx + bw, tx + tw) + loop_off
            edge(slide, [(rx, ry), (loop_x, ry), (loop_x, ty + th / 2),
                         (tx + tw + 0.03, ty + th / 2)])
            edge_label(slide, rx + 0.15, ry - 0.42, 1.2, no, size=label_size)
    return yy - gap
