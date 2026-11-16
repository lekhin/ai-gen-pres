# -*- coding: utf-8 -*-
"""Единый пред-рендер аудит собранного .pptx (структура + плейсхолдеры + текст).

Запуск:  python scripts/pptx_qa.py <file.pptx>
Проверяет без рендера то, что глаза по PNG не поймают:
  FAIL (код возврата 1):
    * битовый OOXML-пакет (zip, обязательные части, висячие ссылки slide-rels,
      непарсящийся XML слайдов);
    * пустые плейсхолдеры-заглушки («Текст слайда» в режиме правки);
    * мусор шаблонов оформления (lorem ipsum, TODO, XXX/ХХХ, [вставьте…]);
  WARN (код не меняют, но требуют взгляда):
    * шрифты вне разрешённых темой (theme.json, поле qa_fonts);
    * цвета вне палитры темы (без -t тема/имя файла неизвестны — сверяется
      с объединением всех тем, ложных предупреждений меньше);
    * двойные пробелы в тексте;
    * геометрия: фигура вылезла за границы слайда; текстовый блок в зоне
      шапки/логотипа; оценка высоты текста сильно больше блока;
      line_spacing-множитель >3 (частая ошибка: пункты вместо множителя);
    * схемы: связь (коннектор) проходит сквозь текст подписи/метки.
Код возврата: 0 — чисто/только предупреждения, 1 — есть ошибки.
"""
import math
import re
import sys
import zipfile

from lxml import etree
from pptx import Presentation

# палитра/шрифты для предупреждений берутся из темы (deck_theme.qa_allowlists)

# мусор невычитанного текста: (регексп, человекочитаемое описание)
JUNK_PATTERNS = [
    (re.compile(r"(lorem|ipsum)", re.I), "lorem ipsum"),
    (re.compile(r"\b(TODO|TBD)\b", re.I), "TODO/TBD"),
    (re.compile(r"\b[ХXхx]{3,}\b"), "заглушка XXX/ХХХ"),
    (re.compile(r"[\[{](встав|назван|подпис|замен|текст сюд)", re.I), "квадратная скобка-подсказка"),
    (re.compile(r"\{\{.*?\}\}"), "неподставленная {{переменная}}"),
]

_a = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


def check_package(path, err, warn):
    """Структурная целостность OOXML-пакета."""
    if not zipfile.is_zipfile(path):
        err("это не zip-архив — файл не является pptx")
        return
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        for need in ("[Content_Types].xml", "_rels/.rels", "ppt/presentation.xml"):
            if need not in names:
                err(f"в пакете нет обязательная часть {need}")
        slides = sorted(n for n in names if re.fullmatch(r"ppt/slides/slide\d+\.xml", n))
        if not slides:
            err("в пакете нет ни одного слайда ppt/slides/slideN.xml")
        for sp in slides:
            try:
                etree.fromstring(z.read(sp))
            except etree.XMLSyntaxError as e:
                err(f"{sp}: XML не парсится ({e.msg})")
            rp = f"ppt/slides/_rels/{sp.rsplit('/', 1)[1]}.rels"
            if rp not in names:
                continue
            try:
                root = etree.fromstring(z.read(rp))
            except etree.XMLSyntaxError as e:
                err(f"{rp}: XML не парсится ({e.msg})")
                continue
            base = sp.rsplit("/", 1)[0]
            for rel in root:
                if rel.get("TargetMode") == "External":
                    continue
                tgt = rel.get("Target", "").replace("../", "ppt/")
                if not tgt.startswith("ppt/") and not tgt.startswith("docProps"):
                    tgt = f"{base}/{rel.get('Target', '')}"
                clean = "/".join(p for p in tgt.split("/") if p not in ("", "."))
                if clean not in names:
                    err(f"{rp}: ссылка {rel.get('Target')!r} ведёт в никуда")


def check_content(path, err, warn, theme=None):
    """Плейсхолдеры, текст, шрифты, цвета."""
    from deck_theme import qa_allowlists
    PALETTE, BRAND_FONTS = qa_allowlists(theme)
    prs = Presentation(path)
    empty_ph, junk, dblspace = [], [], []
    fonts, colors = set(), set()

    def feed_text(slide_no, where, text):
        for rx, desc in JUNK_PATTERNS:
            if rx.search(text):
                junk.append((slide_no, where, desc, text.strip()[:60]))
        for line in text.splitlines():
            if "  " in line.strip():
                dblspace.append((slide_no, where, line.strip()[:60]))
                break

    def walk(shapes, slide_no):
        for sh in shapes:
            if sh.shape_type == 6:  # GROUP
                walk(sh.shapes, slide_no)
                continue
            if sh.is_placeholder and sh.has_text_frame and not sh.text_frame.text.strip():
                empty_ph.append((slide_no, sh.placeholder_format.idx))
            if sh.has_text_frame:
                feed_text(slide_no, sh.name, sh.text_frame.text)
                for p in sh.text_frame.paragraphs:
                    for r in p.runs:
                        if r.font.name:
                            fonts.add(r.font.name)
            if getattr(sh, "has_table", False) and sh.has_table:
                for row in sh.table.rows:
                    for c in row.cells:
                        feed_text(slide_no, sh.name, c.text)

    a_srgb = _a + "srgbClr"
    for i, s in enumerate(prs.slides, 1):
        walk(s.shapes, i)
        for el in s._element.iter(a_srgb):
            v = (el.get("val") or "").upper()
            if v:
                colors.add(v)

    for i, idx in empty_ph:
        err(f"слайд {i}: пустой плейсхолдер idx {idx} (PowerPoint покажет «Текст слайда»)")
    for i, w_, desc, frag in junk:
        err(f"слайд {i} [{w_}]: {desc}: {frag!r}")
    for i, w_, frag in dblspace:
        warn(f"слайд {i} [{w_}]: двойной пробел: {frag!r}")
    bad_fonts = fonts - BRAND_FONTS
    if bad_fonts:
        warn(f"шрифты вне разрешённых темой: {', '.join(sorted(bad_fonts))}")
    off_pal = {c for c in colors if c.upper() not in PALETTE}
    if off_pal:
        warn(f"цвета вне палитры темы: {', '.join(sorted(off_pal))}")


EMU_CM = 360000          # EMU в сантиметре
PT_EMU = 12700           # EMU в пункте
TOL_EMU = 108000         # допуск геометрии 3 мм
HEADER_BOTTOM_CM = 4.4   # значение по умолчанию (16:9); берётся из сетки темы


def _header_bottom_cm(theme):
    """Нижняя граница зоны шапки из сетки темы: контент начинается с
    grid.content_top_cm, линейка колодки — на ~1,2 см выше."""
    try:
        import deck_theme
        grid = deck_theme.load(theme)["grid"]
        return round(grid["content_top_cm"] - 1.2, 2), grid.get("footer_top_cm")
    except Exception:
        return HEADER_BOTTOM_CM, None


def _cm(emu):
    return round(emu / EMU_CM, 2)


def _est_overflow(slide_no, sh, warn):
    """Грубая оценка: влезает ли текст в текстовый блок (не плейсхолдер)."""
    w_pt = sh.width / PT_EMU
    box_pt = sh.height / PT_EMU
    h_pt, bad_ls = 0.0, False
    for p in sh.text_frame.paragraphs:
        if not p.runs:
            continue
        size = max((r.font.size.pt if r.font.size else 12) for r in p.runs)
        text = "".join(r.text for r in p.runs)
        mult = 1.0
        if isinstance(p.line_spacing, (int, float)) and p.line_spacing > 0:
            mult = float(p.line_spacing)
            if mult > 3 and not bad_ls:  # line_spacing — МНОЖИТЕЛЬ, не пункты
                bad_ls = True
                warn(f"слайд {slide_no} [{sh.name}]: line_spacing={mult:g} — "
                     f"вероятно пункты вместо множителя (надо ~1.2)")
        cpl = max(8, int(w_pt / (0.52 * size)))  # средняя ширина глифы ~0.52em
        h_pt += max(1, math.ceil(len(text) / cpl)) * size * 1.25 * mult
        if p.space_after is not None:
            h_pt += p.space_after.pt
    if box_pt > 0 and h_pt > 1.6 * box_pt and h_pt - box_pt > 40:
        warn(f"слайд {slide_no} [{sh.name}]: высота текста ~{h_pt:.0f}pt против "
             f"блока {box_pt:.0f}pt — вылезет за блок (интерлиньяж/переносы?)")


def check_geometry(path, warn, theme=None):
    """Границы слайда, зона шапки, переполнение текстовых блоков.
    Только верхнеуровневые фигуры: у фигур внутри групп координаты относительные.
    Границы зон берутся из сетки темы (портрет A4 и 16:9 живут по-разному)."""
    prs = Presentation(path)
    header_bottom, footer_top = _header_bottom_cm(theme)
    SW, SH = prs.slide_width, prs.slide_height
    for i, s in enumerate(prs.slides, 1):
        for sh in s.shapes:
            if None in (sh.left, sh.top, sh.width, sh.height):
                continue
            full = sh.width >= 0.97 * SW and sh.height >= 0.97 * SH  # фон
            if full:
                continue
            if (sh.left < -TOL_EMU or sh.top < -TOL_EMU
                    or sh.left + sh.width > SW + TOL_EMU
                    or sh.top + sh.height > SH + TOL_EMU):
                warn(f"слайд {i} [{sh.name}]: за границами слайда: "
                     f"x={_cm(sh.left)} y={_cm(sh.top)} w={_cm(sh.width)} h={_cm(sh.height)} см")
            if (not sh.is_placeholder and sh.has_text_frame
                    and sh.text_frame.text.strip()
                    and sh.top < header_bottom * EMU_CM):
                warn(f"слайд {i} [{sh.name}]: текстовый блок в зоне шапки/логотипа "
                     f"(y={_cm(sh.top)} см < {header_bottom} см)")
            if (footer_top and not sh.is_placeholder and sh.has_text_frame
                    and sh.text_frame.text.strip()
                    and sh.top > (footer_top + 0.3) * EMU_CM):
                warn(f"слайд {i} [{sh.name}]: набор под колонтитулами "
                     f"(y={_cm(sh.top)} см > {footer_top} см)")
            if not sh.is_placeholder and sh.has_text_frame and sh.text_frame.text.strip():
                _est_overflow(i, sh, warn)


def _seg_hits_box(x1, y1, x2, y2, l, t, r, b):
    """Пересекает ли отрезок прямоугольник (Liang-Barsky)."""
    if l < x1 < r and t < y1 < b and l < x2 < r and t < y2 < b:
        return True  # оба конца внутри (или касаются) — считаем пересечением
    dx, dy = x2 - x1, y2 - y1
    t0, t1 = 0.0, 1.0
    for p, q in ((-dx, x1 - l), (dx, r - x1), (-dy, y1 - t), (dy, b - y1)):
        if p == 0:
            if q < 0:
                return False
            continue
        t_ = q / p
        if p < 0:
            if t_ > t1:
                return False
            if t_ > t0:
                t0 = t_
        else:
            if t_ < t0:
                return False
            if t_ < t1:
                t1 = t_
    return t0 <= t1


def _chains(conns):
    """Сегменты-коннекторы → полилинии (цепи). Возвращает [(сегменты, концы)].
    Концом цепи (точка «прикрепления» связи) считается точка, к которой подходит
    ровно один сегмент; внутренние вершины «ёлочек» концами не являются."""
    from collections import defaultdict
    q = lambda v: round(v / 10000)
    adj = defaultdict(set)
    for i, (x1, y1, x2, y2) in enumerate(conns):
        adj[(q(x1), q(y1))].add(i)
        adj[(q(x2), q(y2))].add(i)
    seen, out = set(), []
    for i in range(len(conns)):
        if i in seen:
            continue
        stack, chain = [i], set()
        while stack:
            j = stack.pop()
            if j in chain:
                continue
            chain.add(j)
            x1, y1, x2, y2 = conns[j]
            for p in ((q(x1), q(y1)), (q(x2), q(y2))):
                stack.extend(adj[p] - chain)
        seen |= chain
        cnt = defaultdict(int)
        for j in chain:
            x1, y1, x2, y2 = conns[j]
            cnt[(q(x1), q(y1))] += 1
            cnt[(q(x2), q(y2))] += 1
        out.append((chain, {p for p, c in cnt.items() if c == 1}))
    return out


def check_diagram(path, warn):
    """Схемы: метка/подпись насквозь пробита связью (коннектором).
    Концы связей, прикреплённые к узлу (внешние концы полилиний у границ
    боксов), дефектом не считаются; внутренние вершины «ёлочек» — считаются."""
    prs = Presentation(path)
    tol = int(0.08 * EMU_CM)        # допуск «проходит сквозь»
    dock = int(0.35 * EMU_CM)       # конец связи у края = прикрепление к узлу
    q = lambda v: round(v / 10000)
    for i, s in enumerate(prs.slides, 1):
        conns, texts = [], []
        for sh in s.shapes:
            if type(sh).__name__ == "Connector":
                if None in (sh.left, sh.top, sh.width, sh.height) or \
                        sh.width == 0 and sh.height == 0:
                    continue
                try:
                    from pptx.enum.dml import MSO_LINE
                    if sh.line.dash_style == MSO_LINE.ROUND_DOT:
                        continue   # точечные направляющие (линии жизни) — не связи
                except (AttributeError, TypeError):
                    pass
                try:
                    conns.append((sh.begin_x, sh.begin_y, sh.end_x, sh.end_y))
                except Exception:
                    continue
            elif (not sh.is_placeholder and sh.has_text_frame
                    and sh.text_frame.text.strip()
                    and None not in (sh.left, sh.top, sh.width, sh.height)):
                texts.append((sh, sh.name, sh.text_frame.text.strip()[:40]))
        chains = _chains(conns)
        for sh, name, frag in texts:
            l = sh.left - tol
            t = sh.top - tol
            r = sh.left + sh.width + tol
            b = sh.top + sh.height + tol
            dl, dt, dr, db = l - dock, t - dock, r + dock, b + dock
            hit = None
            for chain, ends in chains:
                for j in chain:
                    x1, y1, x2, y2 = conns[j]
                    docked = ((dl < x1 < dr and dt < y1 < db
                               and (q(x1), q(y1)) in ends)
                              or (dl < x2 < dr and dt < y2 < db
                                  and (q(x2), q(y2)) in ends))
                    if docked:
                        continue
                    if _seg_hits_box(x1, y1, x2, y2, l, t, r, b):
                        hit = frag
                        break
                if hit:
                    break
            if hit:
                warn(f"слайд {i} [{name}]: связь проходит сквозь текст "
                     f"{hit!r} — сдвинь метку или маршрут")


def main(path, theme=None):
    errors, warnings = [], []
    err = lambda m: errors.append(m)      # noqa: E731
    warn = lambda m: warnings.append(m)   # noqa: E731
    check_package(path, err, warn)
    if not errors:
        try:
            check_content(path, err, warn, theme=theme)
            check_geometry(path, warn, theme=theme)
            check_diagram(path, warn)
        except Exception as e:  # pptx не открылся вовсе
            err(f"python-pptx не смог открыть файл: {e}")
    for m in errors:
        print(f"FAIL {m}")
    for m in warnings:
        print(f"WARN {m}")
    if errors:
        print(f"ИТОГ: ошибок {len(errors)}, предупреждений {len(warnings)}")
        return 1
    print(f"OK: ошибок нет (предупреждений {len(warnings)})")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    argv = sys.argv[1:]
    theme = None
    for flag in ("-t", "--theme"):
        if flag in argv:
            i = argv.index(flag)
            theme = argv[i + 1]
            del argv[i:i + 2]
    if not argv:
        print(__doc__)
        sys.exit(2)
    sys.exit(main(argv[0], theme=theme))
