# -*- coding: utf-8 -*-
"""Хелперы сборки презентаций python-pptx по шаблону активной ТЕМЫ
(_service\\themes\\<тема>\\template.pptx + theme.json).

Использование в build-скрипте:
    from pptx_helpers import *
    prs = load_template(TPL)            # TPL = deck_theme.template_path()
    s = new_slide(prs, L_CONT, header="Название слайда")
    ph(s, 0).text_frame.text = "Название слайда"
    set_text(ph(s, 1).text_frame, [{"runs": [R("Текст", 14)]}])
    prs.save(OUT)

Палитра/шрифты/сетка/имена layout'ов берутся из theme.json активной темы
(см. deck_theme.py; выбор темы — ДО импорта этого модуля).

Ключевые грабли, учтённые здесь:
- add_slide() не копирует плейсхолдеры даты/колонтитула/номера: new_slide
  клонирует их <p:sp> из эталонной заготовки шаблона;
- изменение только width/height у плейсхолдера роняет его в (0,0): используйте
  place(sh, left, top, w, h) — задаёт все четыре координаты;
- буллеты унаследованного стиля убираются set_text(..., {"bullet": False})
  (a:buNone + marL/indent = 0);
- у всех автофигур/линий отключается shadow.inherit (иначе лезет тема Office);
- фоновая плашка на весь слайд добавляется вместе с send_to_back, иначе
  ляжет поверх текста;
- перекраска полей даты/номера на тёмном фоне — только через recolor_ph
  (в <a:fld> текст без <a:r>, p.runs пуст);
- неиспользуемые плейсхолдеры удаляются drop_ph (иначе PowerPoint рисует
  заглушку «Текст слайда»).
"""
import copy

from pptx import Presentation
from pptx.util import Cm, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.oxml.ns import qn

# ---------------- палитра и шрифты активной темы ----------------
from deck_theme import active as _active_theme   # noqa: E402

_T = _active_theme()
_P = _T["palette"]


def _c(hexv):
    return RGBColor.from_string(str(hexv).upper())


ACCENT = _c(_P["accent"])              # акцент (цифры, выделения, модель на схемах)
INK = _c(_P["ink"])                    # основной текст/структура
PAPER = _c(_P["paper"])                # фон слайда темы (белый/тёмный)
LINE = _c(_P["line"])                  # разделители шаблона
MUTED = _c(_P.get("muted", _P["line"]))
GOOD = _c(_P["good"]) if "good" in _P else INK
BAD = _c(_P["bad"]) if "bad" in _P else ACCENT
NAV = [_c(v) for v in _P.get("nav", [])]

# семантические алиасы (исторические имена кодовой базы — не переизобретать):
BLACK, WHITE, RED, GRAY = INK, PAPER, ACCENT, LINE
ROYAL_BLUE = NAV[0] if len(NAV) > 0 else ACCENT
AMETHYST = NAV[1] if len(NAV) > 1 else INK
APPLE_GREEN = NAV[2] if len(NAV) > 2 else LINE
CINNABAR = ACCENT

HEAD = _T["fonts"]["head"]             # заголовки/колонтитулы/пагинация
BODY = _T["fonts"]["body"]             # набор/даты/подписи
CG, INTER = HEAD, BODY                 # алиасы старой кодовой базы

# ---------------- сетка слайда (из темы) ----------------
_G = _T["grid"]
SLIDE_W = _T["slide"]["w_cm"]          # см
SLIDE_H = _T["slide"]["h_cm"]          # см
MARGIN = _G["margin_cm"]               # охранное поле от края, см
GRID_W = SLIDE_W - 2 * MARGIN          # наборная полоса
COLS = _G["cols"]
COL_W = GRID_W / COLS                  # ширина колонки сетки

CONTENT_TOP = _G["content_top_cm"]     # верх области контента (под заголовком), см
FOOTER_TOP = _G["footer_top_cm"]       # низовые колонтитулы, см

HEADER_ON = bool(_T.get("header", True))       # шапка-дубль в верхнем колонтитуле
DARK_SLIDE = _T.get("dark_slide") or {}        # параметры тёмного слайда в светлой теме


def col_x(i):
    """Левый край i-й колонки (i = 0..COLS-1), см."""
    return MARGIN + i * COL_W


def load_template(path):
    """Открыть шаблон и вырезать эталонные слайды, предварительно забрав из
    второго слайда элементы низовых колонтитулов (idx 10/11/12)."""
    prs = Presentation(path)
    ref = prs.slides[1] if len(prs.slides) > 1 else prs.slides[0]
    sps = {}
    for p in ref.placeholders:
        i = p.placeholder_format.idx
        if i in (10, 11, 12):
            sps[i] = copy.deepcopy(p._element)
    xml_slides = prs.slides._sldIdLst
    for sldId in list(xml_slides):
        prs.part.drop_rel(sldId.rId)
        xml_slides.remove(sldId)
    prs._deck_footer_sps = sps
    return prs


def layouts(prs):
    """(Титул, Контент, Контент с рисунком) по именам layout'ов из theme.json."""
    names = _T["layouts"]
    ls = {l.name: l for l in prs.slide_masters[0].slide_layouts}
    try:
        return ls[names["title"]], ls[names["content"]], ls[names["image"]]
    except KeyError as e:
        raise RuntimeError(
            f"в шаблоне нет layout'а {e.args[0]!r}; есть: {sorted(ls)}; "
            f"ждёт имена из theme.json темы {_T['name']!r}") from None


def ph(slide, idx):
    for p in slide.placeholders:
        if p.placeholder_format.idx == idx:
            return p
    return None


def new_slide(prs, layout, header=None, with_date=True):
    """Добавить слайд + живые поля даты/номера + шапку-дубль заголовка."""
    s = prs.slides.add_slide(layout)
    sps = getattr(prs, "_deck_footer_sps", {})
    if 12 in sps:
        s.shapes._spTree.append(copy.deepcopy(sps[12]))
    if with_date and 10 in sps:
        s.shapes._spTree.append(copy.deepcopy(sps[10]))
    if header is not None:
        if 11 not in sps:
            raise RuntimeError("В шаблоне не найден эталонный колонтитул (idx 11)")
        s.shapes._spTree.append(copy.deepcopy(sps[11]))
        ph(s, 11).text_frame.text = header
    return s


def R(text, size=14, font=INTER, bold=False, italic=False, color=BLACK):
    """run-спецификация для set_text."""
    return (text, font, size, bold, italic, color)


def set_text(tb, items):
    """Заполнить текст. tb — shape или TextFrame.
    items: [{"runs": [R(...)], lvl, align, space_after, line, bullet}]
    bullet=False снимает маркеры и отступы унаследованного стиля."""
    tf = tb.text_frame if hasattr(tb, "text_frame") else tb
    tf.word_wrap = True
    tf.clear()
    first = True
    for it in items:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.level = it.get("lvl", 0)
        if it.get("align"):
            p.alignment = it["align"]
        if it.get("space_after") is not None:
            p.space_after = Pt(it["space_after"])
        if it.get("line") is not None:
            p.line_spacing = it["line"]
        if it.get("bullet") is False:
            pPr = p._p.get_or_add_pPr()
            pPr.set("marL", "0")
            pPr.set("indent", "0")
            pPr.append(pPr.makeelement(qn("a:buNone"), {}))
        for (txt, fname, size, bold, italic, color) in it["runs"]:
            r = p.add_run()
            r.text = txt
            r.font.name = fname
            r.font.size = Pt(size)
            r.font.bold = bold
            r.font.italic = italic
            r.font.color.rgb = color
    return tf


def place(sh, x, y, w, h):
    """Переместить/изменить фигуру: ВСЕ четыре координаты, иначе (0,0)."""
    sh.left, sh.top, sh.width, sh.height = Cm(x), Cm(y), Cm(w), Cm(h)


def textbox(slide, x, y, w, h, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(Cm(x), Cm(y), Cm(w), Cm(h))
    tf = tb.text_frame
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    return tb


def rect(slide, x, y, w, h, fill=None, line=None, line_w=1.0):
    sp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(x), Cm(y), Cm(w), Cm(h))
    sp.shadow.inherit = False
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid()
        sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line
        sp.line.width = Pt(line_w)
    return sp


def send_to_back(sh):
    """Переместить фигуру в самый низ Z-порядка слайда (под плейсхолдеры).
    Нужен для фоновых плашек на всю площадь: добавленный rect иначе ляжет
    поверх текста. В spTree индексы 0-1 заняты nvGrpSpPr/grpSpPr."""
    sp = sh._element
    spTree = sp.getparent()
    spTree.remove(sp)
    spTree.insert(2, sp)
    return sh


def recolor_ph(shape, color=WHITE):
    """Перекрасить ВЕСЬ текст плейсхолдера (для тёмных фонов).
    Поля даты/номера содержат <a:fld> со своим <a:rPr> и без <a:r> —
    перекраска только p.runs их не берёт, поэтому идём по XML-элементам
    форматирования a:rPr/a:defRPr/a:endParaRPr."""
    hexv = "%02X%02X%02X" % (color[0], color[1], color[2])
    for tag in ("a:rPr", "a:defRPr", "a:endParaRPr"):
        for rPr in shape.text_frame._txBody.iter(qn(tag)):
            for sf in rPr.findall(qn("a:solidFill")):
                rPr.remove(sf)
            sf = rPr.makeelement(qn("a:solidFill"), {})
            sf.append(rPr.makeelement(qn("a:srgbClr"), {"val": hexv}))
            rPr.insert(0, sf)
    return shape


def drop_ph(slide, idx):
    """Удалить неиспользуемый плейсхолдер с элемента слайда.
    Пустой плейсхолдер из layout в режиме правки PowerPoint показывает
    заглушку «Текст слайда» — на сохранённых слайдах их быть не должно
    (проверка: scripts/pptx_check_ph.py)."""
    p = ph(slide, idx)
    if p is not None:
        p._element.getparent().remove(p._element)
    return p is not None


def hline(slide, x, y, w, color=BLACK, w_pt=1.0):
    ln = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Cm(x), Cm(y), Cm(x + w), Cm(y))
    ln.line.color.rgb = color
    ln.line.width = Pt(w_pt)
    ln.shadow.inherit = False
    return ln


def vline(slide, x, y, h, color=BLACK, w_pt=1.0):
    ln = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Cm(x), Cm(y), Cm(x), Cm(y + h))
    ln.line.color.rgb = color
    ln.line.width = Pt(w_pt)
    ln.shadow.inherit = False
    return ln


# ---------------- нативные диаграммы в стиле темы ----------------
# Порядок цветов серий/секторов: акцент, затем навигационная палитра темы.
CHART_COLORS = [ACCENT] + NAV + [INK, LINE]


def chart(slide, kind, x, y, w, h, cats, series, *, num_fmt="General",
          y_max=None, y_min=None, colors=None, legend=None, labels=True,
          label_size=9, axis_size=9):
    """Нативная диаграмма PowerPoint в стиле темы (редактируется заказчиком).

    kind: "bar" | "hbar" | "line" | "pie" | "doughnut".
    cats: подписи категорий; series: [значения] или [(имя, [значения]), ...].
    colors: переопределение палитры серий; legend: None = авто (внизу, если
    серий > 1); labels: подписи значений; num_fmt: формат чисел ('0,0', '0%'...).
    Грабли: у кольцевой диаграммы позиции подписей CENTER (OUTSIDE_END ломает
    файл); единая точка данных у pie — один сектор на весь круг.
    Возвращает graphic frame (можно place(gf, ...) для тонкой правки позиции).
    """
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION

    kinds = {"bar": XL_CHART_TYPE.COLUMN_CLUSTERED,
             "hbar": XL_CHART_TYPE.BAR_CLUSTERED,
             "line": XL_CHART_TYPE.LINE,
             "pie": XL_CHART_TYPE.PIE,
             "doughnut": XL_CHART_TYPE.DOUGHNUT}
    if kind not in kinds:
        raise ValueError(f"kind должен быть одним из {sorted(kinds)}")
    if kind in ("pie", "doughnut"):
        if series and isinstance(series[0], tuple):
            if len(series) > 1:
                raise ValueError("pie/doughnut принимают одну серию значений")
        elif series and isinstance(series[0], (list, tuple)):
            if len(series) > 1:
                raise ValueError("pie/doughnut принимают одну серию значений")
            series = [("Данные", tuple(series[0]))]
        else:
            series = [("Данные", list(series))]
    elif series and not isinstance(series[0], tuple):
        series = [("Серия", series)]

    data = CategoryChartData()
    data.categories = list(cats)
    for name, vals in series:
        data.add_series(name, tuple(vals))

    gf = slide.shapes.add_chart(kinds[kind], Cm(x), Cm(y), Cm(w), Cm(h), data)
    ch = gf.chart
    pal = list(colors) if colors else list(CHART_COLORS)
    multi = len(series) > 1

    ch.has_title = False
    ch.font.name = INTER
    ch.font.size = Pt(axis_size)
    ch.font.color.rgb = BLACK

    # серии: цвет и линии
    for i, ser in enumerate(ch.series):
        c = pal[i % len(pal)]
        if kind == "line":
            ser.format.line.color.rgb = c
            ser.format.line.width = Pt(2.2)
            try:
                ser.smooth = False
            except (AttributeError, TypeError):
                pass
        else:
            if kind in ("pie", "doughnut"):
                for pi, pt in enumerate(ser.points):
                    pt.format.fill.solid()
                    pt.format.fill.fore_color.rgb = pal[pi % len(pal)]
                    pt.format.line.color.rgb = WHITE
                    pt.format.line.width = Pt(1.5)
            else:
                ser.format.fill.solid()
                ser.format.fill.fore_color.rgb = c
                ser.format.line.fill.background()

    plot = ch.plots[0]
    if kind in ("bar", "hbar"):
        plot.gap_width = 60
        plot.vary_by_categories = not multi

    # подписи значений
    if labels:
        plot.has_data_labels = True
        dl = plot.data_labels
        dl.font.name = INTER
        dl.font.size = Pt(label_size)
        dl.font.bold = True
        dl.font.color.rgb = BLACK
        dl.number_format = num_fmt
        dl.number_format_is_linked = False
        if kind in ("bar", "hbar"):
            dl.position = XL_LABEL_POSITION.OUTSIDE_END
        elif kind == "line":
            dl.position = XL_LABEL_POSITION.ABOVE  # OUTSIDE_END line-чарт ломает
        elif kind == "pie":
            dl.position = XL_LABEL_POSITION.BEST_FIT
        else:  # doughnut — единственный допустимый вариант
            dl.position = XL_LABEL_POSITION.CENTER

    # легенда: только для нескольких серий (pie/doughnut — по категориям-секторам)
    if legend is None:
        legend = multi or kind in ("pie", "doughnut")
    ch.has_legend = legend
    if legend:
        ch.legend.position = XL_LEGEND_POSITION.BOTTOM
        ch.legend.include_in_layout = False
        ch.legend.font.name = INTER
        ch.legend.font.size = Pt(axis_size)
        ch.legend.font.color.rgb = BLACK

    # оси (нет у pie/doughnut)
    if kind not in ("pie", "doughnut"):
        for axis in (ch.category_axis, ch.value_axis):
            axis.has_major_gridlines = False
            axis.format.line.color.rgb = BLACK
            axis.format.line.width = Pt(0.75)
            tl = axis.tick_labels
            tl.font.name = INTER
            tl.font.size = Pt(axis_size)
            tl.font.color.rgb = BLACK
        va = ch.value_axis
        va.has_major_gridlines = True
        va.major_gridlines.format.line.color.rgb = GRAY
        va.major_gridlines.format.line.width = Pt(0.5)
        if y_max is not None:
            va.maximum_scale = y_max
        if y_min is not None:
            va.minimum_scale = y_min
    return gf


def chart_image(slide, path, x, y, w=None, h=None):
    """Вставить PNG графики (matplotlib/deck_mpl) на слайд.
    Фигуры deck_mpl делаются ровно под ячейку слайда в см, поэтому обычно
    достаточно w (высота сохранит пропорции PNG = нужный размер). Переданные
    w и h вместе исказят картинку — передавайте только один."""
    kw = {}
    if w is not None:
        kw["width"] = Cm(w)
    if h is not None:
        kw["height"] = Cm(h)
    return slide.shapes.add_picture(str(path), Cm(x), Cm(y), **kw)
