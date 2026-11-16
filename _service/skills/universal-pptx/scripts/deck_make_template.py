# -*- coding: utf-8 -*-
"""Генерирует template.pptx (носитель стилей) из theme.json темы.

python deck_make_template.py [имя-или-путь-темы ...]   # без аргументов — все темы

Контракт шаблона (см. SKILL.md «Ключевые параметры шаблона темы»):
  * мастер с txStyles: заголовки — fonts.head, текст — fonts.body, цвета из palette;
  * layout'ы с именами из theme.json["layouts"]:
      title   — центр. заголовок + подзаголовок + акцентная полоса;
      content — заголовок + линия-разделитель с акцентным сегментом + текст;
      image   — то же + правый бокс (idx 2) под рисунок/схему;
    в каждом — плейсхолдеры даты(10)/колонтитула(11)/номера(12);
  * слайд 1 — титул; слайд 2 — ЭТАЛОННЫЙ контентный (с него new_slide берёт
    клоны ph 10/11/12, шапка-дубль idx 11 вверху справа); слайд 3 — image.

Геометрия строится из theme.json (slide.w_cm/h_cm + grid), поэтому тема может
быть и портретной (книжный A4) — раскладка считается от размеров страницы.
После сборки перегенерируется inspect.md темы.
"""
import contextlib
import io
import json
import sys
from pathlib import Path

import pptx_inspect
from pptx import Presentation
from pptx.util import Cm, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from pptx.oxml import parse_xml

NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'


def _c(h):
    return RGBColor.from_string(h.upper())


def def_rpr(sz, font, color, bold=False):
    b = ' b="1"' if bold else ''
    return (f'<a:defRPr sz="{int(sz*100)}"{b} xmlns:a="http://schemas.openxmlformats.org/'
            f'drawingml/2006/main"><a:solidFill><a:srgbClr val="{color.upper()}"/></a:solidFill>'
            f'<a:latin typeface="{font}"/><a:cs typeface="{font}"/></a:defRPr>')


def set_ph_default(ph, sz, font, color, bold=False, align=None, anchor=None):
    """Мастер-стиль плейсхолдера уровня layout: defRPr в lstStyle layout'а."""
    tx = ph.text_frame._txBody
    for el in tx.findall(qn('a:lstStyle')):
        tx.remove(el)
    lst = parse_xml(f'<a:lstStyle {NS}/>')
    bodyPr = tx.find(qn('a:bodyPr'))
    bodyPr.addnext(lst)
    pPr = parse_xml(f'<a:lvl1pPr {NS}' + (f' algn="{align}"' if align else '') + '/>')
    pPr.append(parse_xml(def_rpr(sz, font, color, bold)))
    lst.append(pPr)
    if anchor:
        bodyPr.set('anchor', anchor)
        bodyPr.set('anchorCtr', 'true' if anchor == 'ctr' else 'false')


def master_txstyles(master, head, body, ink, sz_title, sz_body):
    """txStyles мастера: чем наследуются все плейсхолдеры без своего стиля."""
    tx = master.element.find(qn('p:txStyles'))
    if tx is None:
        return
    for tag, (sz, fname) in (('a:titleStyle', (sz_title, head)), ('a:bodyStyle', (sz_body, body)),
                             ('a:otherStyle', (sz_body, body))):
        el = tx.find(qn(tag))
        if el is None:
            continue
        pPr = el.find(qn('a:lvl1pPr'))
        if pPr is None:
            pPr = parse_xml(f'<a:lvl1pPr {NS}/>')
            el.insert(0, pPr)
        for old in pPr.findall(qn('a:defRPr')):
            pPr.remove(old)
        pPr.append(parse_xml(def_rpr(sz, fname, ink, bold=(tag == 'a:titleStyle'))))


def solid(shape, color_hex):
    shape.fill.solid()
    shape.fill.fore_color.rgb = _c(color_hex)
    shape.line.fill.background()
    shape.shadow.inherit = False
    return shape


def move(sh, x, y, w, h):
    sh.left, sh.top, sh.width, sh.height = Cm(x), Cm(y), Cm(w), Cm(h)


class _LayoutCanvas:
    """add_shape на layout: фигура создаётся на служебном слайде и переносится в layout."""

    def __init__(self, prs, layout):
        self.prs, self.layout = prs, layout

    def add_shape(self, st, x, y, w, h):
        tmp = self.prs.slides.add_slide(self.layout)
        sp = tmp.shapes.add_shape(st, x, y, w, h)
        sp._element.getparent().remove(sp._element)
        self.layout.shapes._spTree.append(sp._element)
        xml_slides = self.prs.slides._sldIdLst
        sldId = list(xml_slides)[-1]
        self.prs.part.drop_rel(sldId.rId)
        xml_slides.remove(sldId)
        return sp


def clone_phs(slide, layout, idxs):
    """add_slide НЕ переносит ph даты/колонтитула/номера — клонируем вручную."""
    import copy
    spTree = slide.shapes._spTree
    for p in layout.placeholders:
        if p.placeholder_format.idx in idxs:
            spTree.append(copy.deepcopy(p._element))


def drop_layouts(prs, keep_idx):
    """Оставить layout'ы с индексами keep_idx (привязка master→layout по sldLayoutIdLst)."""
    master = prs.slide_masters[0]
    lst = master.element.find(qn('p:sldLayoutIdLst'))
    for i, lid in enumerate(list(lst)):
        if i in keep_idx:
            continue
        master.part.drop_rel(lid.get(qn('r:id')))
        lst.remove(lid)


def build(theme_dir: Path):
    theme = json.loads((theme_dir / "theme.json").read_text(encoding="utf-8-sig"))
    g, pal, fnt = theme["grid"], theme["palette"], theme["fonts"]
    head, body = fnt["head"], fnt["body"]
    names = theme["layouts"]
    PW, PH = theme["slide"]["w_cm"], theme["slide"]["h_cm"]
    MX = g["margin_cm"]
    CW = PW - 2 * MX
    COL_W = CW / g["cols"]
    CTOP, FTOP = g["content_top_cm"], g["footer_top_cm"]
    Y_RULE = round(CTOP - 0.73, 2)          # линия под заголовком (landscape: 4.87)
    SZ_TITLE = g.get("title_pt", 28)
    SZ_COVER = g.get("cover_pt", 40)
    SZ_SUB = g.get("cover_sub_pt", 15)
    SZ_BODY = g.get("body_pt", 16)

    prs = Presentation()
    prs.slide_width, prs.slide_height = Cm(PW), Cm(PH)
    drop_layouts(prs, {0, 1, 3})            # Title Slide / Title+Content / Two Content
    lay_cover, lay_content, lay_image = prs.slide_layouts
    lay_cover.name, lay_content.name, lay_image.name = names["title"], names["content"], names["image"]
    master = prs.slide_masters[0]
    master.background.fill.solid()
    master.background.fill.fore_color.rgb = _c(pal["paper"])
    master_txstyles(master, head, body, pal["ink"], SZ_TITLE, SZ_BODY)

    # --- низовые колонтитулы во всех layout'ах: дата(10) / колонтитул(11) / номер(12) ---
    for lay in (lay_cover, lay_content, lay_image):
        for p in lay.placeholders:
            idx = p.placeholder_format.idx
            if idx == 10:
                move(p, MX, FTOP, 10.0, 0.75)
            elif idx == 11:
                move(p, MX + 12.0, FTOP, CW - 16.0, 0.75)
            elif idx == 12:
                move(p, PW - MX - 10.0, FTOP, 10.0, 0.75)
                p.text_frame.paragraphs[0].alignment = PP_ALIGN.RIGHT

    # --- layout «Титул» ---
    ph_t = lay_cover.placeholders[0]
    ph_s = lay_cover.placeholders[1]
    move(ph_t, MX + 1.0, PH * 0.31, CW - 2.0, PH * 0.23)
    move(ph_s, MX + 3.0, PH * 0.572, CW - 6.0, PH * 0.12)
    set_ph_default(ph_t, SZ_COVER, head, pal["ink"], bold=True, align="ctr", anchor="ctr")
    set_ph_default(ph_s, SZ_SUB, body, pal["muted"], align="ctr", anchor="t")
    bar = solid(_LayoutCanvas(prs, lay_cover).add_shape(
        MSO_SHAPE.RECTANGLE, Cm(0), Cm(0), Cm(5.0), Cm(0.16)), pal["accent"])
    move(bar, (PW - 5.0) / 2, PH * 0.572 - 0.4, 5.0, 0.16)
    bar.name = "Акцентная полоса"

    # --- layout'ы «Контент» / «Контент с рисунком»: линия-разделитель + сегмент ---
    for lay in (lay_content, lay_image):
        canvas = _LayoutCanvas(prs, lay)
        t = lay.placeholders[0]
        move(t, MX, 2.3, CW, 1.9)
        set_ph_default(t, SZ_TITLE, head, pal["ink"], bold=True)
        ln = solid(canvas.add_shape(MSO_SHAPE.RECTANGLE, Cm(0), Cm(0), Cm(CW), Cm(0.06)),
                   pal["line"])
        move(ln, MX, Y_RULE, CW, 0.06)
        ln.name = "Линия-разделитель"
        seg = solid(canvas.add_shape(MSO_SHAPE.RECTANGLE, Cm(0), Cm(0), Cm(COL_W), Cm(0.22)),
                    pal["accent"])
        move(seg, MX, Y_RULE - 0.08, COL_W, 0.22)
        seg.name = "Акцентный сегмент"
        for p in lay.placeholders:
            if p.placeholder_format.idx == 1:
                move(p, MX, CTOP, CW, FTOP - CTOP - 0.6)
                set_ph_default(p, SZ_BODY, body, pal["ink"])
    ph_img = [p for p in lay_image.placeholders if p.placeholder_format.idx == 2][0]
    half = (CW - 1.1) / 2
    move(lay_image.placeholders[1], MX, CTOP, half, FTOP - CTOP - 0.6)
    move(ph_img, MX + half + 1.1, CTOP, half, FTOP - CTOP - 0.6)
    set_ph_default(ph_img, 12, body, pal["muted"], align="ctr", anchor="ctr")

    # --- слайд 1: титул ---
    s1 = prs.slides.add_slide(lay_cover)
    s1.placeholders[0].text_frame.text = "Заголовок презентации"
    s1.placeholders[1].text_frame.text = "Подзаголовок / контекст"

    # --- слайд 2: эталонный контентный (источник клонов ph 10/11/12) ---
    s2 = prs.slides.add_slide(lay_content)
    clone_phs(s2, lay_content, {10, 11, 12})
    s2.placeholders[0].text_frame.text = "Заголовок слайда"
    tf = s2.placeholders[1].text_frame
    tf.text = "Первый абзац."
    tf.add_paragraph().text = "Второй абзац."
    # шапка-дубль: колонтитул idx 11 переносим в правый верхний угол
    dub = [p for p in s2.placeholders if p.placeholder_format.idx == 11][0]
    move(dub, PW - 3.0 - 10.37, 0.5, 10.37, 0.85)
    dub.name = "Шапка-дубль"
    dub.text_frame.paragraphs[0].alignment = PP_ALIGN.RIGHT
    set_ph_default(dub, 10, head, pal["muted"], bold=True, align="r")
    d = [p for p in s2.placeholders if p.placeholder_format.idx == 10][0]
    move(d, MX, FTOP, 10.0, 0.75)
    d.name = "Дата"
    n = [p for p in s2.placeholders if p.placeholder_format.idx == 12][0]
    move(n, PW - MX - 10.0, FTOP, 10.0, 0.75)
    n.name = "Номер слайда"
    n.text_frame.paragraphs[0].alignment = PP_ALIGN.RIGHT

    # --- слайд 3: image ---
    s3 = prs.slides.add_slide(lay_image)
    s3.placeholders[0].text_frame.text = "Заголовок слайда"
    s3.placeholders[1].text_frame.text = "Текст слева"
    ph_img_slide = [p for p in s3.placeholders if p.placeholder_format.idx == 2][0]
    ph_img_slide.text_frame.text = "Место для рисунка"

    out = theme_dir / "template.pptx"
    prs.save(out)

    # --- inspect.md (пересобирается вместе с шаблоном) ---
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        pptx_inspect.main(out, False)
    (theme_dir / "inspect.md").write_text(
        f"# Инспекция шаблона темы {theme['name']} (сгенерирован deck_make_template.py)\n\n"
        f"```\n{buf.getvalue()}\n```\n", encoding="utf-8")
    print(f"{out}: {PW:.2f}×{PH:.2f} см, layout'ы: "
          f"{names['title']} / {names['content']} / {names['image']}")


def main(dirs=None):
    here = Path(__file__).resolve().parents[3]   # корень фабрики
    themes_dir = here / "themes"
    if not dirs:
        dirs = sorted(p.name for p in themes_dir.iterdir() if (p / "theme.json").exists())
    for name in dirs:
        d = Path(name)
        if not (d / "theme.json").exists():
            d = themes_dir / name
        if not (d / "theme.json").exists():
            print(f"тема не найдена: {name}")
            return 1
        build(d)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
