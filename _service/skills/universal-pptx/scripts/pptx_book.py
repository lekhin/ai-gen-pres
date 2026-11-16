# -*- coding: utf-8 -*-
"""Книжная (и вообще «плотная полосная») вёрстка для python-pptx: полоса по
сетке темы, заголовки групп с линейками, чек-листы и нумерованные списки со
висячим отступом и построчными линейками, раздача остатка высоты по отбивкам,
две колонки с линейкой-разделителем, выноски у поля, вертикальная «лента»
шагов, пирамида уровней.

Геометрия НЕ хардкодится: MX/W/TOP/BOT считаются из theme.json (slide,
grid.margin_cm, grid.content_top_cm, grid.footer_top_cm) — правка темы
перекраивает вёрстку автоматически.

Типовой лист:
    import os; os.environ["DECK_THEME"] = "book"       # до импорта helpers
    from pptx_helpers import load_template, layouts
    from pptx_book import Band, page, sect, lead, note, spread_cols, box, num

    prs = load_template(template_path())
    band = Band()                                      # MX, W, TOP, BOT из темы
    s = page(prs, layouts(prs)[1], band, "Приложение 1.2 · Чек-лист перед КТ",
             "Прил. 1.2 · Чек-лист перед КТ")
    spread_cols(s, band, [("Клиническая часть", [box("сформулирована задача КТ")]),
                          ("Протокол", [box("область минимально необходима")])])

Все списки набираются висячим отступом (первая строка выступающая) и, если
пунктов несколько, построчными линейками — печатная форма. Метрика — из
pptx_metrics, поэтому кегль подбирается под реальный набор, а не «на глаз».
"""
import os

from pptx.util import Cm, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

import deck_theme
import pptx_metrics as M
from pptx_helpers import (ACCENT, BAD, BODY, GOOD, HEAD, INK, LINE, MUTED, NAV, PAPER,
                          R, drop_ph, hline, new_slide, ph, rect, set_text, textbox,
                          vline)

_PAL = deck_theme.active()["palette"]


def _rgb(hexv):
    from pptx.dml.color import RGBColor
    return RGBColor.from_string(str(hexv).upper()) if hexv else None


TINT = _rgb(_PAL.get("accent_tint"))       # акцент приглушённый (плашки, схемы)
SOFT = _rgb(_PAL.get("accent_soft"))       # светлая тонировка акцентом
FSOFT = _rgb(_PAL.get("fill_soft"))        # нейтральная светлая плашка

HANG = 0.55            # висячий отступ списка, см
GAP_MIN = 0.9          # минимальный зазор между группами, см
NOTE_PT = 9            # кегль выносок
GAP_PT = 4.5           # базовая отбивка между пунктами, pt


class Band:
    """Печатная полоса активной темы: поля, ширина, верх/низ области набора."""

    def __init__(self, margin=None):
        t = deck_theme.active()
        g = t["grid"]
        self.mx = margin if margin is not None else g["margin_cm"]
        self.w = t["slide"]["w_cm"] - 2 * self.mx
        self.top = g["content_top_cm"]
        self.bottom = g["footer_top_cm"] - 0.3

    # точки остановки колонки: с выноской у поля / без выноски
    @property
    def end(self):
        return self.bottom - 2.2

    @property
    def end_flat(self):
        return self.bottom - 1.0

    def colw(self, gut=1.0, cols=2):
        return (self.w - gut * (cols - 1)) / cols


# ─────────────────────────────────── каркас листа ────────────────────────────
def page(prs, layout, band, title, header=None, title_pt=20):
    """Слайд-полоса: заголовок в плейсхолдер, лишние плейсхолдеры снести."""
    s = new_slide(prs, layout, header=header)
    set_text(ph(s, 0), [{"runs": [R(title, title_pt, HEAD, True, False, INK)],
                         "bullet": False}])
    drop_ph(s, 1)
    return s


def sect(s, band, x, y, w, txt, size=11.5, face="head"):
    """Заголовок группы: шрифт HEAD, цвет акцента, линейка под ним. → низ."""
    hh = M.para_h([(txt, size, face, True, False)], w)
    set_text(textbox(s, x, y, w, hh + 0.1),
             [{"runs": [R(txt, size, HEAD, True, False, ACCENT)], "bullet": False}])
    hline(s, x, y + hh + 0.12, w, color=LINE, w_pt=0.75)
    return y + hh + 0.42


def sect_h(txt, w, size=11.5, face="head"):
    return M.para_h([(txt, size, face, True, False)], w) + 0.42


def lead(s, band, x, y, w, txt, size=9.5):
    """Вводная строка-курсив под колодкой заголовка. → низ."""
    hh = M.para_h([(txt, size, "body", False, True)], w)
    set_text(textbox(s, x, y, w, hh + 0.15),
             [{"runs": [R(txt, size, BODY, False, True, MUTED)], "bullet": False}])
    return y + hh + 0.35


def note(s, band, x, y, w, paras, marker=None):
    """Выноска: вертикальная линейка + курсивный набор. → низ."""
    if isinstance(paras, str):
        paras = [paras]
    hh = M.runs_h([[(t, NOTE_PT, "body", False, True)] for t in paras], w - 0.45)
    vline(s, x, y + 0.05, hh + 0.1, color=marker or ACCENT, w_pt=1.4)
    set_text(textbox(s, x + 0.45, y, w - 0.45, hh + 0.15),
             [{"runs": [R(t, NOTE_PT, BODY, False, True, INK)], "space_after": 2,
               "bullet": False} for t in paras])
    return y + hh + 0.3


def band_note(s, band, w_paras, x=None, y=None):
    """Выноска у нижнего поля полосы (типовое место)."""
    return note(s, band, x if x is not None else band.mx,
                y if y is not None else band.bottom - 1.35, band.w, w_paras)


def para(s, x, y, w, paras, size=10, italic=False, color=None):
    """Обычный абзацный блок по ширине w (без списковых штучек). → низ."""
    items = [{"runs": [R(t, size, BODY, False, italic, color or INK)],
              "space_after": 2, "bullet": False} for t in ([paras] if isinstance(paras, str) else paras)]
    hh = M.space_h([[(t, size, "body", False, italic)] for t in
                    ([paras] if isinstance(paras, str) else paras)], w, 2)
    set_text(textbox(s, x, y, w, hh + 0.15), items)
    return y + hh + 0.3


def plate(s, x, y, w, txt, size=10, italic=False, fill=None, face="body"):
    """Плашка-правило: фраза на тонированной полосе. → низ после плашки."""
    hh = M.para_h([(txt, size, face, False, italic)], w - 0.7)
    font = HEAD if face == "head" else BODY
    rect(s, x, y, w, hh + 0.6, fill=SOFT if fill is None else fill)
    set_text(textbox(s, x + 0.35, y + 0.28, w - 0.7, hh + 0.1),
             [{"runs": [R(txt, size, font, False, italic, INK)], "bullet": False}])
    return y + hh + 0.95


# ─────────────────────────────────── списки ──────────────────────────────────
def hang(tf, marL=HANG):
    """Висячий отступ абзацев (книжный набор списков) поверх bullet=False."""
    for p in tf.paragraphs:
        pPr = p._p.get_or_add_pPr()
        pPr.set("marL", str(int(marL * 360000)))
        pPr.set("indent", str(-int(marL * 360000)))


def box(text, note_=None):
    """Пункт чек-листа с квадратиком."""
    return ("☐", text, note_, None)


def num(i, text, note_=None):
    """Нумерованный пункт (01, 02, …)."""
    return (f"{i:02d}", text, note_, None)


def list_items(s, x, y, w, rows, size=10.0, gap=GAP_PT, fill_bottom=None,
               rules=True):
    """Список rows=[(метка, текст, заметка|None, цвет_метки|None)].

    Каждый пункт — отдельный текстовый блок, между пунктами тонкая линейка
    (печатная форма). fill_bottom — раздать остаток высоты в межпунктные
    отбивки, чтобы список кончался у этой отметки. → высота блока, см.
    """
    hs, paras = [], []
    for lbl, txt, nt, *_ in rows:
        para_ = [(lbl + " ", size, "body", True, False), (txt, size, "body", False, False)]
        if nt:
            para_.append((" — " + nt, size - 1.5, "body", False, True))
        paras.append(para_)
        hs.append(M.para_h(para_, w, HANG))
    gap_cm = gap * M.CM_PER_PT
    h_nat = sum(hs) + len(rows) * gap_cm
    extra = 0.0
    if fill_bottom and fill_bottom > y + h_nat:
        extra = (fill_bottom - y - h_nat) / len(rows)
    yy = y
    for (lbl, txt, nt, *rest), ph_, para_ in zip(rows, hs, paras):
        col = rest[0] if rest else None
        runs = [R(lbl + " ", size, BODY, True, False, col or ACCENT),
                R(txt, size, BODY, False, False, INK)]
        if nt:
            runs.append(R(" — " + nt, size - 1.5, BODY, False, True, MUTED))
        tb = textbox(s, x, yy, w, ph_ + 0.1)
        set_text(tb, [{"runs": runs, "bullet": False}])
        hang(tb.text_frame)
        if rules and len(rows) > 1:
            hline(s, x, yy + ph_ + (gap_cm + extra) / 2, w, color=LINE, w_pt=0.5)
        yy += ph_ + gap_cm + extra
    return yy - gap_cm - extra - y


def list_h(rows, w, size=10.0, gap=GAP_PT):
    """Оценка высоты списка (та же метрика, что у list_items)."""
    return M.list_h(rows, w, size, gap, HANG)


# ─────────────────────────── две колонки с разделителем ──────────────────────
def spread_cols(s, band, groups, end_y=None, divider=True, extra_cap=2.2,
                gut=1.0, size_max=13.0, size_min=9.0):
    """groups=[(заголовок|None, rows), …] — раскладка по образцу кничного
    чек-листа: группы делятся по двум колонкам (баланс по высоте, порядок
    чтения сохраняется), колонки доводятся до end_y и разделяются линейкой.
    Кегль списка подбирается под объём полосы (от size_max к size_min)."""
    end_y = band.end if end_y is None else end_y
    colw = band.colw(gut)

    def gh(gr, size):
        gt, rw = gr
        return (sect_h(gt, colw, size + 1.5) if gt else 0.2) + list_h(rw, colw, size)

    def split_of(ests):
        tot = sum(ests)
        return min(range(1, len(ests)), key=lambda k: abs(tot - 2 * sum(ests[:k])))

    sizes = [c for c in (size_max, 12, 11.5, 11, 10.5, 10, 9.5, size_min) if c >= size_min]
    size = sizes[-1]
    for cand in dict.fromkeys(sizes):
        e = [gh(g, cand) for g in groups]
        k = split_of(e) if len(e) > 1 else 1
        col_max = max(sum(e[:k]), sum(e[k:])) if len(e) > 1 else sum(e)
        if col_max + GAP_MIN <= end_y - (band.top + 0.2):
            size = cand
            break
    ests = [gh(g, size) for g in groups]
    k = split_of(ests) if len(ests) > 1 else 1
    cols = [groups[:k], groups[k:]] if len(groups) > 1 else [groups, []]
    ends, extras = [], []
    for col, grs in enumerate(cols):
        gx = band.mx if col == 0 else band.mx + colw + gut
        gy = band.top + 0.2
        if not grs:
            ends.append(gy)
            extras.append(0.0)
            continue
        e = [gh(g, size) for g in grs]
        n_row = sum(len(rw) for _, rw in grs)
        extra = min(max(0.0, (end_y - (band.top + 0.2) - sum(e) - GAP_MIN * (len(grs) - 1))
                        / n_row), extra_cap)
        extras.append(extra)
        for (gt, rw), eh in zip(grs, e):
            body_h = eh - (sect_h(gt, colw, size + 1.5) if gt else 0.2)
            gy_l = sect(s, band, gx, gy, colw, gt, size=size + 1.5) if gt else gy + 0.2
            hh = list_items(s, gx, gy_l, colw, rw, size=size,
                            fill_bottom=gy_l + body_h + len(rw) * extra)
            gy = gy_l + hh + GAP_MIN
        ends.append(gy - GAP_MIN)
    if divider:
        tail = (GAP_PT * M.CM_PER_PT + max(extras)) / 2
        bottom = min(max(ends) + tail, band.bottom - 0.5)
        vline(s, band.mx + colw + gut / 2, band.top + 0.25,
              bottom - (band.top + 0.25), color=LINE, w_pt=0.75)
    return ends


# ────────────────────────── лента шагов и пирамида уровней ───────────────────
def steps_ribbon(s, band, steps, x=None, w=None, y_top=None, y_bot=None,
                 col_split=0.42, head=None, pictos=None):
    """Вертикальная «лента» шагов сверху вниз: номер на оси-таймлайне,
    пиктограмма, название шага слева, что звучит/делается — справа.

    steps=[(название, текст), …]; pictos=[функция(slide, cx, cy, sc)|None, …].
    → нижняя отметка ленты."""
    x = band.mx if x is None else x
    w = band.w if w is None else w
    y_top = band.top if y_top is None else y_top
    y_bot = band.bottom - 0.4 if y_bot is None else y_bot
    x_num = x + 0.35
    x_icon = x + 1.85
    x_div = x + w * col_split
    qw = w - (x_div - x) - 0.45
    if head:
        set_text(textbox(s, x_icon, y_top + 0.05, x_div - x_icon, 0.45),
                 [{"runs": [R(head[0], 8, BODY, True, False, MUTED)], "bullet": False}])
        set_text(textbox(s, x_div + 0.45, y_top + 0.05, qw, 0.45),
                 [{"runs": [R(head[1], 8, BODY, True, False, MUTED)], "bullet": False}])
        hline(s, x, y_top + 0.62, w, color=ACCENT, w_pt=1.1)
        y_top += 0.85
    row_h = (y_bot - y_top) / len(steps)
    for i in range(len(steps) - 1):                  # таймлайн: badge → badge
        y1 = y_top + i * row_h + row_h / 2 + 0.4
        y2 = y_top + (i + 1) * row_h + row_h / 2 - 0.4
        vline(s, x_num, y1, y2 - y1, color=SOFT, w_pt=1.1)
    vline(s, x_div, y_top + 0.1, len(steps) * row_h - 0.4, color=LINE, w_pt=0.75)
    for i, (name, txt) in enumerate(steps):
        yc = y_top + i * row_h + row_h / 2
        if i:
            hline(s, x, y_top + i * row_h - 0.02, w, color=LINE, w_pt=0.5)
        badge = s.shapes.add_shape(MSO_SHAPE.OVAL, Cm(x_num - 0.36), Cm(yc - 0.36),
                                   Cm(0.72), Cm(0.72))
        badge.fill.solid()
        badge.fill.fore_color.rgb = ACCENT
        badge.line.fill.background()
        badge.shadow.inherit = False
        tfb = badge.text_frame
        tfb.margin_left = tfb.margin_right = tfb.margin_top = tfb.margin_bottom = 0
        set_text(tfb, [{"runs": [R(str(i + 1), 9.5, HEAD, True, False, PAPER)],
                        "align": PP_ALIGN.CENTER, "bullet": False}])
        if pictos and i < len(pictos) and pictos[i]:
            pictos[i](s, x_icon, yc, 0.78)
        set_text(textbox(s, x_icon + 1.1, yc - 0.5, x_div - x_icon - 1.6, 1.0),
                 [{"runs": [R(name, 9.5, HEAD, True, False, INK)], "bullet": False}])
        set_text(textbox(s, x_div + 0.45, yc - 0.55, qw, 1.25),
                 [{"runs": [R(txt, 8.5, BODY, False, True, INK)], "bullet": False}])
    return y_top + len(steps) * row_h


def pyramid(s, band, levels, x=None, w=None, y_top=None, right_w=None,
            minw=4.4, maxw=9.0):
    """Пирамида уровней слева + реквизиты уровней справа (одна высота строки).

    levels=[(метка, короткое_имя, полное_имя, документы, оговорка|None, fill), …]
    сверху вниз от верхнего уровня. → нижняя отметка пирамиды."""
    x = band.mx if x is None else x
    w = 9.0 if w is None else w
    y_top = band.top if y_top is None else y_top
    right_w = band.w - w - 0.75 if right_w is None else right_w
    rx = x + w + 0.75
    DOC, NT, NM = 8, 7.5, 9
    lh = max(
        M.space_h([[(n2, NM, "head", True, False)],
                   [(docs[0], DOC, "body", False, False)]]
                  + ([[(nt, NT, "body", False, True)]] if nt else []),
                  right_w, 2) + 0.7
        for _, _, n2, docs, nt, _ in levels)
    lh = max(lh, 2.4)
    gap = 0.22
    for i, (lbl, short, full, docs, nt, fill) in enumerate(levels):
        wi = minw + (maxw - minw) * i / max(1, len(levels) - 1)
        xx = x + (w - wi) / 2
        yy = y_top + i * (lh + gap)
        sp = s.shapes.add_shape(MSO_SHAPE.TRAPEZOID, Cm(xx), Cm(yy), Cm(wi), Cm(lh))
        sp.fill.solid()
        sp.fill.fore_color.rgb = fill
        sp.line.color.rgb = LINE if fill in (SOFT, FSOFT) else fill
        sp.line.width = Pt(0.75)
        sp.shadow.inherit = False
        ink = PAPER if fill in (ACCENT, TINT) else INK
        tf = sp.text_frame
        tf.margin_left = tf.margin_right = Cm(0.15)
        tf.margin_top = tf.margin_bottom = Cm(0.05)
        set_text(tf, [{"runs": [R(lbl, 7.5, BODY, False, False, ink)],
                       "align": PP_ALIGN.CENTER, "bullet": False},
                      {"runs": [R(short, 9, HEAD, True, False, ink)],
                       "align": PP_ALIGN.CENTER, "bullet": False}])
        hline(s, rx - 0.35, yy + 0.06, right_w + 0.35, color=LINE, w_pt=0.75)
        items = [{"runs": [R(full, NM, HEAD, True, False, INK)], "space_after": 2,
                  "bullet": False},
                 {"runs": [R(docs[0], DOC, BODY, False, False, INK)], "space_after": 2,
                  "bullet": False}]
        if nt:
            items.append({"runs": [R(nt, NT, BODY, False, True, MUTED)], "bullet": False})
        set_text(textbox(s, rx, yy + 0.2, right_w, lh - 0.25), items)
    return y_top + len(levels) * (lh + gap)


def picto(kind):
    """Фабрика пиктограмм из нативных фигур: returns f(slide, cx, cy, sc).
    Набор — вопрос, шеврон, предупреждение, «уменьшаем», развилка, запрет,
    вопрос-облачко, диалог, решение."""
    def f(s, cx, cy, sc=1.0):
        k = 1.5 * sc

        def sh(shape, x, y, w, h, fill, rot=0):
            sp = s.shapes.add_shape(shape, Cm(x), Cm(y), Cm(w), Cm(h))
            sp.fill.solid()
            sp.fill.fore_color.rgb = fill
            sp.line.fill.background()
            sp.shadow.inherit = False
            if rot:
                sp.rotation = rot
            return sp

        def put(sp, ch, size, color):
            tf = sp.text_frame
            tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
            set_text(tf, [{"runs": [R(ch, size, HEAD, True, False, color)],
                           "align": PP_ALIGN.CENTER, "bullet": False}])

        if kind == "question":
            put(sh(MSO_SHAPE.OVAL, cx - k/2, cy - k/2, k, k, ACCENT), "?", k*22, PAPER)
        elif kind == "chevron":
            sh(MSO_SHAPE.CHEVRON, cx - k/2, cy - k*0.33, k, k*0.66, TINT)
        elif kind == "warning":
            put(sh(MSO_SHAPE.ISOSCELES_TRIANGLE, cx - k/2, cy - k/2 + 0.04, k, k*0.92,
                   BAD), "!", k*17, PAPER)
        elif kind == "down":
            sh(MSO_SHAPE.DOWN_ARROW, cx - k*0.31, cy - k/2, k*0.62, k, GOOD)
        elif kind == "fork":
            sh(MSO_SHAPE.RECTANGLE, cx - k*0.05, cy + k*0.06, k*0.1, k*0.44, MUTED)
            sh(MSO_SHAPE.UP_ARROW, cx - k*0.46, cy - k*0.46, k*0.3, k*0.62, MUTED, -28)
            sh(MSO_SHAPE.UP_ARROW, cx + k*0.16, cy - k*0.46, k*0.3, k*0.62, ACCENT, 28)
        elif kind == "prohibition":
            sh(MSO_SHAPE.NO_SYMBOL, cx - k/2, cy - k/2, k, k, ACCENT)
        elif kind == "bubble_q":
            put(sh(MSO_SHAPE.ROUNDED_RECTANGULAR_CALLOUT, cx - k/2, cy - k/2 - 0.04,
                   k, k, NAV[0]), "?", k*20, PAPER)
        elif kind == "dialog":
            sh(MSO_SHAPE.ROUNDED_RECTANGLE, cx - k*0.52, cy - k*0.5, k*0.5, k*0.4, MUTED)
            put(sh(MSO_SHAPE.ROUNDED_RECTANGULAR_CALLOUT, cx - k*0.2, cy - k*0.22,
                   k*0.86, k*0.72, ACCENT), "…", k*20, PAPER)
        elif kind == "decision":
            put(sh(MSO_SHAPE.DIAMOND, cx - k/2, cy - k/2, k, k, GOOD), "✓", k*15, PAPER)
    return f


if __name__ == "__main__":
    print(f"полоса активной темы ({deck_theme.active()['name']}):",
          vars(Band()))
    print("PIL:", M.have_pil())
    if os.environ.get("DECK_DEBUG"):
        print("METRIC", M.selftest())
