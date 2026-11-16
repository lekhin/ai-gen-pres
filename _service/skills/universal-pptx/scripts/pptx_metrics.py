# -*- coding: utf-8 -*-
"""Метрика набора по РЕАЛЬНЫМ шрифтам темы: сколько строк и сколько сантиметров
займёт текст в блоке заданной ширины. Нужна, чтобы подбирать кегль и растягивать
отбивки ДО сборки, а не угадывать по рендеру.

Зачем, если есть грубая оценка в pptx_qa:
  qa отвечает на «влезает ли вообще», метрика — на «во сколько строк влезет и
  каким кеглем добить полосу». Ошибка грубой оценки — 1–2 строки на пункт, чего
  хватает, чтобы список вылез за нижнее поле или оставил под собой дыру.

Шрифты берутся из theme.json активной темы (fonts.head / fonts.body) и ищутся:
  * Inter — в `_service\\fonts` (variable: оси wght/opsz задаются через
    set_variation_by_axes, иначе bold считается по Regular-метрике);
  * Georgia и прочие системные — в папке шрифтов Windows / Linux;
  * нет файла — фолбэк на Inter Regular (ошибка ~3 % на кириллице);
  * нет PIL — аппроксимация (0,52em на знак, межстрочный 1,2em), тот же
    порядок, что в pptx_qa.

API (всм — сантиметры, кегль — пункты):
    line_h(pt, face, bold)                  высота строки
    n_lines(para, w, hang)                  число строк абзаца
    para_h(para, w, hang)                   высота абзаца
    runs_h(paras, w, hang)                  высота набора абзацев
    list_h(rows, w, size, gap)              высота чек-листа/списка
    space_h(paras, w)                       высота обычного текстового блока

para = [(текст, кегль, face, bold, italic), ...]; face: "body"|"head"|имя шрифта.
rows = [(метка, текст, заметка|None, ...)] — как в pptx_book.list_items.
"""
from functools import lru_cache

import deck_theme
from deck_paths import service_file

DPI = 300                      # raster-метрика: чем выше, тем точнее
CM_PER_PT = 0.03528            # 1 pt в см
CHAR_EM = 0.52                 # аппроксимация: средняя ширина знака, em
LEAD_EM = 1.20                 # аппроксимация: межстрочный, em

_FONTS = deck_theme.active().get("fonts", {})
_FDIR = service_file("fonts")
_WIN = "C:/Windows/Fonts"
_LINUX = "/usr/share/fonts"

# имя шрифта (в нижнем регистре) → кандидаты на (regular, bold, italic, bold-italic)
_FILES = {
    "inter": (f"{_FDIR}/Inter-VariableFont_opsz,wght.ttf",
              f"{_FDIR}/Inter-VariableFont_opsz,wght.ttf",
              f"{_FDIR}/Inter-Italic-VariableFont_opsz,wght.ttf",
              f"{_FDIR}/Inter-Italic-VariableFont_opsz,wght.ttf"),
    "georgia": (f"{_WIN}/georgia.ttf", f"{_WIN}/georgiab.ttf",
                f"{_WIN}/georgiai.ttf", f"{_WIN}/georgiaz.ttf"),
    "times new roman": (f"{_WIN}/times.ttf", f"{_WIN}/timesbd.ttf",
                        f"{_WIN}/timesi.ttf", f"{_WIN}/timesbi.ttf"),
    "arial": (f"{_WIN}/arial.ttf", f"{_WIN}/arialbd.ttf",
              f"{_WIN}/ariali.ttf", f"{_WIN}/arialbi.ttf"),
}
_FILES["liberation serif"] = (f"{_LINUX}/liberation/LiberationSerif-Regular.ttf",
                              f"{_LINUX}/liberation/LiberationSerif-Bold.ttf",
                              f"{_LINUX}/liberation/LiberationSerif-Italic.ttf",
                              f"{_LINUX}/liberation/LiberationSerif-BoldItalic.ttf")

try:
    from PIL import ImageFont
except Exception:                                       # pragma: no cover
    ImageFont = None


def face_name(face):
    """«head»/«body» → имя шрифта из темы; имя шрифта → как есть."""
    return _FONTS.get(face, face if isinstance(face, str) else "Inter")


def _path(face, bold, italic):
    name = face_name(face).lower()
    key = (bold << 1) | italic
    table = _FILES.get(name)
    if table and table[key]:
        return table[key], name
    return _FILES["inter"][key], name       # фолбэк: Inter меткой не испортить


@lru_cache(maxsize=None)
def _font(pt, face, bold, italic):
    """PIL-фонт в растере DPI; кэш по (кегль, шрифт, начертание)."""
    if ImageFont is None:
        return None
    path, _ = _path(face, bold, italic)
    px = max(1, int(round(pt * DPI / 72)))
    try:
        f = ImageFont.truetype(path, px)
    except Exception:                            # нет файла — метрика с фолбэка
        try:
            f = ImageFont.truetype(_FILES["inter"][(bold << 1) | italic], px)
        except Exception:
            return None
    if "inter" in path.lower() and hasattr(f, "set_variation_by_axes"):
        try:                            # variable-шрифт: вес и оптический размер
            f.set_variation_by_axes([700 if bold else 400, px])
        except Exception:
            pass
    return f


def have_pil():
    return ImageFont is not None


# ────────────────────────────────── измерения ────────────────────────────────
def line_h(pt, face="body", bold=False):
    """Высота строки (подъём + спуск) в см — так же её считает PowerPoint."""
    f = _font(pt, face, bold, False)
    if f is None:
        return pt * LEAD_EM * CM_PER_PT
    a, d = f.getmetrics()
    return (a + d) * 2.54 / DPI


def text_w(text, pt, face="body", bold=False, italic=False):
    """Ширина строки текста в см (без переносов)."""
    f = _font(pt, face, bold, italic)
    if f is None:
        return len(text) * CHAR_EM * pt * CM_PER_PT
    return f.getlength(text) * 2.54 / DPI


def _tokens(para):
    """[(текст, кегль, face, bold, italic)] → ([слова…], самый крупный кегль,
    его шрифт и жирность — они задают высоту строки абзаца)."""
    toks, maxpt, face0, bold0 = [], 0.0, "body", False
    for (txt, pt, face, bold, ital) in para:
        f = _font(pt, face, bold, ital)
        for i, word in enumerate(txt.split(" ")):
            if i:
                toks.append((" ", f))
            if word:
                toks.append((word, f))
        if pt > maxpt:
            maxpt, face0, bold0 = pt, face, bold
    return toks, maxpt, face0, bold0


def n_lines(para, w_cm, hang_cm=0.0):
    """Сколько строк займёт абзац шириной w_cm. hang_cm — висячий отступ
    (первая строка шире последующих на hang_cm)."""
    toks, _, _, _ = _tokens(para)
    if not toks:
        return 1
    lines, cur, wlim = 1, 0.0, w_cm
    for piece, f in toks:
        if f is None:                       # нет PIL — аппроксимация по знакам
            width = len(piece) * CHAR_EM * 10 * CM_PER_PT
        else:
            width = f.getlength(piece) * 2.54 / DPI
        if piece == " " and cur + width <= wlim:
            cur += width
            continue
        if cur + width > wlim and cur > 0:
            lines, cur, wlim = lines + 1, width, w_cm - hang_cm
        else:
            cur += width
    return lines


def para_h(para, w_cm, hang_cm=0.0):
    """Высота абзаца (все строки) в см."""
    _, maxpt, face0, bold0 = _tokens(para)
    return n_lines(para, w_cm, hang_cm) * line_h(maxpt, face0, bold0)


def runs_h(paras, w_cm, hang_cm=0.0):
    """Высота набора абзацев в см (без отбивок между ними)."""
    return sum(para_h(p, w_cm, hang_cm) for p in paras)


def space_h(paras, w_cm, gap_pt=0.0, hang_cm=0.0):
    """Высота блока абзацев с одинаковой отбивкой gap_pt между ними."""
    return runs_h(paras, w_cm, hang_cm) + len(paras) * gap_pt * CM_PER_PT


def list_h(rows, w_cm, size=10.0, gap_pt=4.5, hang_cm=0.55, note_delta=1.5):
    """Высота чек-листа rows=[(метка, текст, заметка|None, …)] — набор + отбивки.
    Заметка ставится на 1,5 pt меньше основного кегля (как в pptx_book)."""
    paras = [[(lbl + " ", size, "body", True, False), (txt, size, "body", False, False)]
             + ([(" — " + nt, size - note_delta, "body", False, True)] if nt else [])
             for lbl, txt, nt, *_ in rows]
    return space_h(paras, w_cm, gap_pt, hang_cm)


def selftest():
    """Самопроверка: сравнить метрику с грубой оценкой и показать цифры."""
    w = 8.0
    rows = [("☐", "пациент понимает, что КТ использует рентгеновское излучение", None),
            ("☐", "CTDIvol и DLP не выдаются за индивидуальный риск", None)]
    print(f"PIL: {have_pil()}  тема: {deck_theme.active()['name']}")
    for size in (9, 10, 11, 12, 13):
        print(f"  {size}pt: list_h={list_h(rows, w, size):.2f} см, "
              f"строка={line_h(size):.3f} см, "
              f"100знак={text_w('а' * 100, size):.2f} см")
    return True


if __name__ == "__main__":
    selftest()
