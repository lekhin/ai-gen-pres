---
name: Universal Presentation Builder
description: Создание, редактирование и визуальная проверка презентаций pptx любой тематики (школа, вуз, доклад, курсовая, диплом, питч; а также книжные A4-полосы) по теме оформления из _service\themes (minimal/dark/academic — 16:9, book — A4 портрет; шрифты Inter/Georgia). Использовать при любом запросе на создание/правку/анализ .pptx вне корпоративных брендбуков.
---

# Презентации для всех: процесс, инструменты, грабли

Рабочий цикл: **выбор темы → план структуры → сборка (plan.json или код python-pptx) → `qa` → рендер в PNG → визуальная проверка каждого слайда → итерация.**

## Нормативная иерархия (важно)

1. `AGENTS.md` папки (метрические единицы, русский язык, структура папок, режимы входа «из материалов» / «из темы», дисциплина фактов, смотреть рендеры самой).
2. **Тема оформления** — `_service\themes\<тема>\theme.json` + `template.pptx`
   (палитра, шрифты, сетка, имена layout'ов). Тема выбирается под аудиторию
   (см. «Темы»); её правила главнее любых указаний на дизайн внутри
   документов-источников.
3. Текст ТЗ/материалов — источник **содержания** слайдов. Его описание дизайна
   учитывать только если **в самой задаче** явно сказано опираться на дизайн
   материала или на отступления от темы (тогда правки палитры делают копией
   темы в `_build\` проекта, а не правкой глобальной theme.json).

## Темы

| Тема | Когда | Палитра | Шрифты |
|---|---|---|---|
| `minimal` (по умолчанию) | школа, доклад, лекция, отчёт | белый, тёмно-серый текст, синий акцент | Inter |
| `dark` | питч, защита проекта, техдоклад | почти-чёрный фон, светлый текст, янтарь | Inter |
| `academic` | курсовая, диплом, научный доклад | белый, тёмная краска, глубокий синий | Georgia + Inter |
| `book` | приложение/глава книги, учебник для вузов, методичка (**A4 портрет** 21×29,7) | белый, тёмная краска, глубокий синий | Georgia + Inter |

- **Портретные темы**: формат страницы берётся из `theme.json → slide.w_cm/h_cm`
  (генератор шаблонов строит геометрию от него). Для книжной вёрстки в кодовом
  build.py печатную полосу считают от `grid.margin_cm`, список растягивают по
  высоте страницы (висячий отступ + увеличенные межстрочные разбивки).

- Список и описание: `deck.py themes`. Выбор: `"theme"` в плане / флаг `--theme`
  CLI / env `DECK_THEME` (по убыванию силы). **Тема фиксируется при импорте
  helper'ов** — в кодовом build.py задавать `os.environ["DECK_THEME"]` ДО импорта
  pptx_helpers (в заготовке метка есть).
- Тёмные слайды внутри светлой темы: `"background": "black"` в слайде плана
  (плашка + перекраска колонтитулов цветами `dark_slide` из theme.json). В теме
  `dark` фон слайда уже тёмный, флаг не нужен.
- Новая тема = копия папки + правка theme.json + `deck.py make-template <тема>`.
  Шаблоны ВСЕГДА пересобирать из theme.json (deck_make_template.py — источник
  истины; руками в PowerPoint шаблоны не править, иначе inspect.md врёт).

## Ключевые параметры шаблона темы (общий контракт, не переизобретать)

- Размер по умолчанию: 33,87 × 19,05 см (12192000 × 6858000 EMU), 16:9;
  в портретных темах (`book`) — из `slide.w_cm/h_cm` theme.json (A4 21×29,7).
- Поля 1,5 см; 6-колоночная сетка: колонка 5,144 см; контент 30,87 см.
  Верх области контента y=5,6 см (ниже линии-разделителя заголовка); колонтитулы y=17,55 см.
- Layout'ы (имена из theme.json): `Титул` (центр. заголовок 40pt bold + подзаголовок
  15pt + акцентная полоса), `Контент` (заголовок 28pt bold слева + серая линия с
  акцентным сегментом + текст), `Контент с рисунком` (заголовок + текст слева +
  правый бокс idx2 под картинку/схему).
- Плейсхолдеры на слайде: idx 0 заголовок, 1 текст/подзаголовок, 2 правый бокс
  (image-layout), 10 дата (10pt, слева внизу), 11 шапка-дубль (10pt bold, справа
  вверху, только когда задан `header`), 12 пагинация (справа внизу).
- Декор (полоса, линия) — на layout'ах, наследуется слайдами автоматически.
- Детали и позиции: `_service\themes\<тема>\inspect.md` (читать вместо запуска
  `inspect`; пересоздаётся `deck.py make-template`).

## Шаги

0. **Папка проекта**: одна презентация = папка. Режим «из материалов» — папка с
   исходниками уже есть, pptx кладётся в её корень; режим «из темы» — создать
   папку `<Тема_презентации>\` и положить туда план. Всё служебное — в
   `<Имя_Проекта>\_build\`: `plan.json` и/или `build.py`, `render\slideN.png`,
   `facts.md`, вспомогательные скрипты. Общие темы/гайд/инструменты — в `_service\`.
1. **Тема и шаблон**: выбрать тему (см. таблицу); параметры — в разделе выше,
   детали — в `themes\<тема>\inspect.md`. Шаблон менять не нужно, если не менялась theme.json.
2. **План структуры**: описать структуру слайдов и сетку ДО сборки (можно сразу
   в `plan.json`); для режима «из темы» — сначала `facts.md` (см. AGENTS.md).
3. **Сборка**: типовые слайды — декларативный `plan.json` → `deck.py build`
   (см. «Декларативный план»); экзотика — кодом (заготовка `_service\build.py`,
   хелперы `pptx_helpers.py`: палитра темы, `chart()`, mpl-графика `deck_mpl.py`
   + `chart_image()`). Схемы с mermaid-исходником — каркас из mermaid:
   `deck.py mmd <file.mmd>`, дальше обычный цикл.
4. **Аудит ДО рендера**: `deck.py qa <project.pptx> [-t тема]` (пакет+плейсхолдеры+
   текст+геометрия). Повторять после каждой пересборки.
5. **Рендер**: `python _service\skills\universal-pptx\scripts\deck.py render <pptx> -o <папка проекта>\_build\render`
   (PNG ≤1500 px, имена `slideN.png`; Windows — PowerPoint COM `pptx_render.ps1`,
   вне Windows — soffice→pdf→pdftoppm/PyMuPDF).
6. **Визуальная проверка**: смотреть каждый PNG самой (`read`), искать наезды/уезды
   текста, наложения, контраст (в тёмных темах — особенно: тёмный текст на тёмном).
   Обзор всех слайдов — `deck.py thumb`.
7. Чек-лист: цвета/шрифты только темы (qa подскажет), заголовки читаются с 3 м,
   на финальном слайде — контакты/источники, если требовались.

## Грабли python-pptx / Windows (проверено, экономит цикл отладки)

- `add_slide()` **не копирует** плейсхолдеры даты/колонтитула/номера слайда
  (латентные типы ph). Клонировать `<p:sp>` с эталонного слайда шаблона
  (делает `pptx_helpers.new_slide`; эталонный слайд — №2 в template.pptx).
- Установка только `shape.width`/`.height` у плейсхолдера **ломает наследование позиции**
  → текст прыгает в (0,0). Всегда задавать все четыре: left, top, width, height.
- Маркеры у заимствованных стилей: для «чистых» параграфов добавлять `a:buNone` +
  `marL="0" indent="0"` в pPr (делает `set_text(bullet=False)`).
- У autoshape/коннекторов ставить `shape.shadow.inherit = False`, иначе въезжает тема Office с тенью.
- .ps1 с кириллицей сохранять в **UTF-8 с BOM**, иначе PowerShell 5.1 ломает парсинг.
- Консоль Windows кракозябрами выводит русские имена файлов — только отображение;
  для проверки состава `Get-ChildItem | ForEach-Object { $_.Name }` или python.
- Плейсхолдеры-«заглушки»: неиспользуемый плейсхолдер из layout показывает в
  PowerPoint подсказку «Текст слайда» (в PNG не видна!). Удалять `drop_ph`
  (apply_plan делает это сам для idx 1/2; проверять `deck.py check`).
- PowerShell 5.1: нет `&&` (разделять `;`); `python -c "..."` с кавычками ломается —
  писать временный `.py`.
- У `prs.slide_layouts[i]` нет `.title`; имена layout'ов — из theme.json (`layouts(prs)`).
- matplotlib **не принимает `RGBColor` из pptx_helpers** — цвета там hex-строки:
  брать константы `deck_mpl` (ACCENT/INK/…) или писать `"2563EB"`.
- Центрально-анкоренный заголовок «Титула»: крупный текст выезжает за box
  симметрично; кегль/высоту подбирать по фактическому числу строк.
- План `plan.json`: у блока `rule` ключ `"w"` — ДЛИНА (см), толщина только `"w_pt"`.
  Верт. линия — `rect` шириной 0.05–0,08 см. `"line"` — МНОЖИТЕЛЬ интерлиньяжа
  (1.2 = 120%): `26` = 26× и текст уезжает за слайд; `"space_after"` — в пунктах.
- Тёмная тема: стандартные «чёрный текст/белый фон» инстинкты инвертированы —
  BLACK=чернила темы (светлые!), WHITE=фон темы. Не писать RGBColor «в лоб»,
  брать константы. Ошибку ловит qa («цвета вне палитры темы») + рендер.

## Инструменты (путь к скриптам: `_service\skills\universal-pptx\scripts\`)

Запускать можно **из любой папки**: служебная папка (`_service\` с темами и
шрифтами) и скрипты скилла находятся автоматически (поиск вверх от скриптов и от
текущей папки — `deck_paths.py`). Если скилл установлен вне `_service\` —
переменные окружения `DECK_SERVICE_DIR`, `DECK_SKILL_SCRIPTS`.

Единая точка входа — CLI `deck.py` (команды работают и по отдельным скриптам):

```
python _service\skills\universal-pptx\scripts\deck.py themes
python _service\skills\universal-pptx\scripts\deck.py inspect <pptx> [--full]
python _service\skills\universal-pptx\scripts\deck.py build   plan.json [-o out.pptx] [--theme dark]
python _service\skills\universal-pptx\scripts\deck.py mmd     diag.mmd [-o diag.json] [--emit py] [--preview out.pptx]
python _service\skills\universal-pptx\scripts\deck.py qa      <pptx> [-t тема]
python _service\skills\universal-pptx\scripts\deck.py dump    <docx|pdf|pptx|md> [-o dump.txt] [--media [папка]]
python _service\skills\universal-pptx\scripts\deck.py facts   <pptx> --facts facts.md [--strict]
python _service\skills\universal-pptx\scripts\deck.py check   <pptx>
python _service\skills\universal-pptx\scripts\deck.py render  <pptx> [-o dir] [-W 1500] [--trim] [--svg]
                                                            # Win: COM; Linux: soffice+pdftoppm
                                                            # --fast: headless ArtCraft DeckCraft
                                                            # --fast --transparent: PNG без фона
python _service\skills\universal-pptx\scripts\deck.py thumb   <pptx | _build\render> [-o grid.png] [-c 4]
python _service\skills\universal-pptx\scripts\deck.py make-template [тема …]
```

Зависимости: `python-pptx`, `lxml`, `Pillow`, `matplotlib` (`pip install`);
PowerPoint — только для рендера (COM). Шрифты из `_service\fonts\` для matplotlib
не ставить в систему — регистрируются автоматически; для PowerPoint/рассылаемого
pptx шрифты темы должны быть установлены (Inter — OFL, лежит в `_service\fonts\`).

| Скрипт | Назначение |
|---|---|
| `deck.py` | Единый CLI: `themes / inspect / build / mmd / qa / check / render / thumb / make-template / dump / facts`. `render` кроссплатформенный: Windows — PowerPoint COM, Linux — soffice+pdftoppm/PyMuPDF, на выходе всегда `_build\render\slideN.png` ≤1500 px; `--trim` — обрезать поля до контента (иллюстрация-файл, с `-W 2400` — печатное качество), `--svg` — ещё и SVG по слайдам (если PowerPoint умеет). |
| `deck_dump.py` | Читатель материалов (`deck.py dump`): docx/pdf/pptx/md/txt → текст с НОМЕРАМИ СТРОК (на них ссылается facts.md), строки таблиц одной строкой; `--media` — извлечение картинок из docx. Дамп кладётся в `<проект>\_build\` и живёт с проектом. |
| `deck_facts.py` | Сверка данных (`deck.py facts`): числовые токены слайдов (без плейсхолдеров, номерных маркеров и бейджей; мета-счётчики «10 пунктов» в facts не в счёт) ↔ значения facts.md; печатает расхождения в обе стороны и ГОТОВУЮ приёмочную строку AGENTS.md (часть «источники присутствуют среди материалов» подтверждает агент). `--strict` — rc=1 при незааявленных числах. |
| `deck_theme.py` | Загрузка theme.json активной темы; список тем; allowlist цветов/шрифтов для qa. |
| `deck_make_template.py` | Генератор template.pptx + inspect.md из theme.json. Пересобирать после изменения theme.json (`deck.py make-template`). |
| `plan.py` | Декларативный план → pptx. Схема JSON — в разделе «Декларативный план plan.json» ниже. Блоки: text/stat/chart/image/table/columns/rect/rule/diagram. |
| `deck_mpl.py` | Стиль matplotlib темы: `figure(w_cm, h_cm)` — фигура ровно под ячейку, шрифты из `_service\fonts\`, палитра темы, `save()` — PNG 300 dpi прозрачный. Вставка `chart_image()`. Для диаграмм, которых нет в PowerPoint (Гант, воронки, heatmap). |
| `pptx_inspect.py` | Дамп структуры pptx: размеры, мастер, layout'ы, плейсхолдеры, формы, шрифты/кегли/цвета. |
| `pptx_render.ps1` | Движок рендера Windows: все слайды в PNG через PowerPoint COM. `-Src file.pptx -Dst dir [-Width 1500]`. |
| `pptx_helpers.py` | Библиотека сборки: палитра/шрифты/сетка темы, `load_template`, `new_slide` (живые поля даты/номера и шапка), `set_text`, `textbox/rect/hline/vline`, `send_to_back`, `recolor_ph`, `drop_ph`, `chart` (нативные диаграммы в цвете темы), `chart_image`. |
| `pptx_diagram.py` | Схемы нативными фигурами: `node` (роли app/data/model/people/planned, формы rect/pill/db/diamond/io), `edge`, `edge_label`, `group`, `caption`, `legend`, `sequence`, `flow`, `row_tops/col_lefts`. Окраска — из палитры активной темы. Эталоны — `_service\diagrams\`. |
| `mmd2diagram.py` | mermaid flowchart → схема: парсер + ранговая автораскладка; вывод — json-блок `diagram` для плана / каркас кода (`--emit py`) / превью. РАСКЛАДКА = КАРКАС, обязательны qa+рендер+правка. |
| `pptx_metrics.py` | Метрика набора по РЕАЛЬНЫМ шрифтам темы: сколько строк и сантиметров займёт текст блоком ширины w (`n_lines/para_h/space_h/list_h/line_h/text_w`). Нужна для подбора кегля и растяжки списков ДО сборки; qa оценивает переполнение лишь грубо («влезает ли вообще»). |
| `pptx_book.py` | Книжная («полосная») вёрстка: `Band` (геометрия полосы из theme.json), `page/sect/lead/note/plate/para`, чек-листы `box()/num()` + `list_items` (висячий отступ, построчные линейки, растяжка до низа), `spread_cols` (две колонки: баланс групп, автоподбор кегля, линейка-разделитель), `steps_ribbon` (вертикальная лента шагов), `pyramid` (пирамида уровней + реквизит), `picto` (пиктограммы из нативных фигур). Тема `book` и любая плотная полосная вёрстка. |
| `smoke_book.py` | Смоук/регрессия книжных модулей: `python …\scripts\smoke_book.py out.pptx` собирает 4 полосы (чек-листы, список, лента, пирамида) — прогнать qa после правок pptx_book/pptx_metrics/qa. |
| `pptx_qa.py` | Пред-рендер аудит: целостность пакета + пустые плейсхолдеры + мусор текста (lorem/TODO/…) + шрифты/цвета вне темы (`-t тема`; без неё — сверка с объединением всех тем) + геометрия (фигура за слайдом, блок в зоне шапки, набор под колонтитулами, переполнение, line_spacing>3); границы зон — из сетки темы (портрет A4 и 16:9 живут по-разному). Код возврата 1 = FAIL. |
| `pptx_check_ph.py` | Аудит плейсхолдеров: `deck.py check file.pptx` — idx/тип/наличие текста; rc=1 при заглушках. |
| `pptx_thumb.py` | Монтажная сетка всех слайдов ОДНОЙ PNG — быстрый визуальный обзор дешевле N рендеров. `-o` — путь к **файлу** PNG (не папке). |

### Декларативный план plan.json (`deck.py build`)

План отделяет ДАННЫЕ от кода: `slides[]` → `layout/header/title/subtitle/
background/blocks[]`; блоки `text|stat|chart|image|table|columns|rect|rule|diagram`,
координаты `"col"+"span"` (сетка) или `x/y/w/h` (см). Лежит в
`<Проект>\_build\plan.json`; картинка `path` — относительно плана.

```jsonc
{ "theme": "dark",                                 // opt; minimal|dark|academic
  "template": "_service\\themes\\dark\\template.pptx",  // opt; явный шаблон поверх темы
  "output": "Проект\\Проект.pptx",                 // opt; без него — <имя_плана>.pptx рядом
  "slides": [ { "layout": "content",       // title | content | image
      "header": "в шапке справа вверху",   // opt (дубль заголовка)
      "title": "…",                        // opt → ph idx0
      "subtitle": "…",                     // opt → idx1; если нет — drop_ph(1)
      "date": true, "background": "black", // opt: тёмный слайд (цвета dark_slide темы)
      "blocks": [ /* координаты: "col" 0..5 + "span" ИЛИ x/y/w/h в см; y по умолч. CONTENT_TOP */
        { "type": "text", "size": 14, "font": "body", "color": "ink",  // default'ы для абзацев
          "paras": [ { "text": "…" } ],                        // или "runs":[{text,size,font,bold,italic,color}]
          /* параграф: lvl, align, space_after(pt!), line(МНОЖИТЕЛЬ: 1.2=120%!), bullet */
          /* или "text": "строки\\nпо\\nсмайлику" */ },
        { "type": "stat", "value": "155", "label": "тыс. т", "note": "…", "size": 48, "color": "accent" },
        { "type": "chart", "kind": "bar|hbar|line|pie|doughnut", "cats": ["…"],
          "series": [ { "name": "…", "values": [1, 2] } ],      // или [[значения…]]
          "num_fmt": "0", "labels": true, "legend": null, "y_max": null, "colors": ["nav1"] },
        { "type": "image", "path": "gantt.png", "w": 15.4, "caption": "…" },  // w ИЛИ h — один, иначе растяжение
        { "type": "table", "rows": [["год", "2024"]], "header": true, "widths": [8, 8] },  // ширины в см
        { "type": "columns", "cols": [ { "span": 3, "gap": 0.3, "blocks": [/*вложенные*/] } ] },
        { "type": "rect", "fill": "line", "line": null, "w_pt": 1 },   // плашка-карточка
        { "type": "rule", "color": "line", "w_pt": 1 },  // гор. линия: "w"/col+span = ДЛИНА см, толщина ТОЛЬКО w_pt
        /* верт. линия: {"type":"rect","w":0.06,"h":…,"fill":"ink"} */
        { "type": "diagram", "groups": [], "nodes": [], "edges": [], "labels": [], "legend": null }
        /* diagram: роли app|data|model|people|plain|planned, формы rect|pill|db|diamond|io;
           детали — «Схемы» / mmd2diagram */
      ] } ] },
  /* цвета: имена палитры темы (accent ink paper line muted good bad + nav1..navN,
     исторические red/black/white/gray тоже работают) | hex "2563EB";
     шрифты: "head" | "body" | явное имя */ }
```

Правила:
- правки итераций — дифф JSON, пересборка `deck.py build`; экзотические слайды —
  код в build.py (гибрид, а не замена);
- `"background": "black"` — канонический тёмный слайд светлой темы: плашка +
  перекраска колонтитулов + белизаголовка по run'ам (recolor_ph наследуемый
  цвет runs не берёт);
- поле `"output"` — относительный путь: ищется от корня фабрики (ближайший
  предок с `_service\`), затем от папки плана; там же ищется `"template"`
  (`deck_paths.resolve_asset`).

### Схемы: `pptx_diagram.py` (узлы/связи/периметры)

Схемы рисовать **нативными фигурами** (заказчик правит в PowerPoint), не картинкой.
Эталоны — `_service\diagrams\01…05` (`build.py`/`plan.json` + preview).

```python
from pptx_diagram import node, edge, edge_label, group, caption, legend, row_tops, col_lefts
R1, R2, R3 = row_tops(8.75, 3, 2.05, 0.8)      # сетку рядов/колонок — хелперами
group(s, x, y, w, h, label="Периметр")          # dash="dash" — зона/план
node(s, x, y, w, h, "Заголовок", "подпись", role="model|data|people|app|plain|planned",
     shape="rect|pill|db|diamond|io")
edge(s, [(x1, y1), (x2, y2)], dash="dash")      # ломаная по точкам = обход блоков
edge_label(s, x, y, w, "подпись связи")         # в зазор >= 2 см или внутрь группы
caption(s, x, y, w, h, [("Вывод: ", True), ("пояснение", False)])
legend(s, x, y, [("solid", "запросы"), ("dash", "wi-fi")])
sequence(s, x, y, w, actors, items)   # UML-секвенции: alt-фреймы, заметки (эталон 02)
flow(s, cx, y, steps)                  # флоучарт: start|proc|io|dec|end (эталон 03)
```

Роли дают окраску из активной темы (app — фон/текст темы, data — серая плашка,
model — акцентная рамка, people — тёмная плашка со светлым, planned — пунктир);
`fill/line/dash` переопределяют. Секвенции/алгоритмы — `sequence()`/`flow()`;
простые схемы декларативно — тип блока `diagram` в плане.

Грабли схем:
- наконечники — только XML `a:tailEnd/a:headEnd` (делает модуль); стрелка —
  свойство ЛИНИИ, не фигуры;
- авто-elbow python-pptx ломается на вертикальных/горизонтальных парах точек:
  обход задавать явным списком точек в `edge()`;
- метку связи ставить в межблочный зазор >= 2 см или внутрь контейнера — не на рамки;
- пересечения: пунктир×сплошная одиночные допустимы; сплошная×сплошная — нет;
- `qa` проверяет «связь проходит сквозь текст» — предупреждения разбирать;
- кегль на схемах — исключение из шкалы: узлы 8–10 pt, связи 6,5–7 pt;
- mermaid-цвета classDef игнорировать: окраска из ролей палитры темы.

Плотные схемы (≥12 узлов): автораскладка `mmd` — каркас; правка в колоночную
(лановую) сетку со свободными коридорами разводки, ортогональные связи явными
точками без from/to, концы ровно на грани узла; эталон `_service\diagrams\05`.

### Импорт из mermaid: `mmd2diagram.py` (`deck.py mmd`)

```
deck.py mmd diag.mmd                 # json-блок diagram для plan.json
deck.py mmd diag.mmd --emit py       # каркас кода для build.py
deck.py mmd diag.mmd --preview p.pptx --title "…" --legend "solid=запросы,dash=wi-fi"
```

Отображение: `flowchart LR|TD`; `subgraph` → `group()`; формы `[(..)]`→db,
`{..}`→ромб, `[/../]`→io; роли классами `id:::model` (app/data/model/people/plain/planned);
`-.->` → пунктир; `|текст|` → метки. Границы: раскладка = каркас (обязательны
qa+рендер+правка); `sub1 --> sub2` не маршрутизируется; секвенции не импортируются.

### Диаграммы: `chart()` из pptx_helpers

Нативные диаграммы PowerPoint (заказчик правит данные), окрашены из темы:
серии из `CHART_COLORS` (акцент → навигационная палитра), сетки серые `line`,
легенда снизу при нескольких сериях.

```python
chart(s, "bar", x, y, w, h, cats, [("Продажи", [124, 132])], num_fmt="0")
chart(s, "doughnut", x, y, w, h, ["А", "Б", "В"], [58, 31, 11], num_fmt='0"%"')
```

`kind`: `bar | hbar | line | pie | doughnut`. Грабли:
- подписи значений у `line` — только `ABOVE`; у `doughnut` — только `CENTER`
  (иначе PowerPoint не открывает файл);
- после добавления графика файл обязательно проверять COM-рендером (`render`):
  python-pptx/LibreOffice битый чарт-XML открывают, PowerPoint — отказывает;
- размеры кратно колонкам сетки (`COL_W`), от `CONTENT_TOP`.

### График «несуществующего» типа: deck_mpl

```python
from deck_mpl import figure, save          # в скрипте _build\ проекта
fig, ax = figure(15.4, 8.0)                # см — ровно ячейка слайда
ax.barh(...)
save(fig, "gantt.png")                      # 300 dpi, прозрачный фон
```

Вставка: `chart_image(slide, "gantt.png", x, y, w=15.4)` — один из w/h.
Нативные чарты (chart()) предпочтительнее: картинка не редактируется заказчиком.

### Метрика набора: `pptx_metrics.py`

qa отвечает на «влезает ли вообще», метрика — на «во сколько строк влезет и каким
кеглем добить полосу». Считает по реальным шрифтам темы: Inter variable из
`_service\fonts\` (вес через `set_variation_by_axes`), Georgia и прочие системные —
из папки шрифтов ОС. Фолбэки: нет файла шрифта → Inter (ошибка ~3 % на кириллице);
нет PIL → аппроксимация 0,52em/1,2em (вне Windows Georgia считается с фолбэка —
кегль для книжной полосы там лучше не поднимать).

```python
import pptx_metrics as M          # тема активна (DECK_THEME задан до импорта)
M.line_h(10, "body")              # высота строки, см
M.n_lines([("текст", 10, "body", False, False)], w_cm=8.0, hang_cm=0.55)
M.space_h(paras, w_cm, gap_pt=2)  # высота блока абзацев с отбивкой
M.list_h(rows, w_cm, size=10)     # высота чек-листа rows=[(метка, текст, заметка|None)]
```

`para` = `[(текст, кегль, face, bold, italic)]`, `face`: `"body"|"head"|имя`.
Самопроверка: `python _service\skills\universal-pptx\scripts\pptx_metrics.py`.

### Книжная вёрстка: `pptx_book.py` (тема book, вообще — «полосная»)

Вся геометрия — из theme.json (`slide`, `grid.margin_cm/content_top_cm/footer_top_cm`),
в проекте не хардкодится. План `plan.py` книжных блоков не знает (чек-лист/лента/
пирамида) — такие полосы собираются кодом через этот модуль:

```python
import os; os.environ["DECK_THEME"] = "book"          # до импорта helper'ов
from deck_theme import template_path
from pptx_helpers import load_template, layouts
from pptx_book import Band, page, lead, sect, plate, note, spread_cols, box

band = Band()                                          # mx / w / top / bottom (+ end, end_flat)
prs = load_template(template_path())
s = page(prs, layouts(prs)[1], band, "Приложение 1.2 · Чек-лист перед КТ", "Прил. 1.2")
lead(s, band, band.mx, band.top - 0.55, band.w, "вводная курсивом")   # → низ
spread_cols(s, band, [("Клиническая часть", [box("сформулирована задача КТ"), ...]),
                      ("Протокол", [...])])                          # end_y=band.end*
```

- `spread_cols` — раскладка по образцу книжного чек-листа: группы делятся на две
  колонки с балансом по высоте (порядок чтения сохраняется), кегль подбирается
  9–13 pt под объём, списки растягиваются в отбивки до `end_y`, колонки
  разделяются линейкой до последней построчной линейки. Беззаголовочные группы
  (одна большая) остаются в одной колонке.
- `list_items(s, x, y, w, rows, fill_bottom=…)` — нижний уровень: висячий отступ
  (`hang`), построчные линейки-«печатная форма», раздача остатка в отбивки
  (cap `extra_cap`, см) → возвращает ФАКТИЧЕСКУЮ высоту относительно `y`.
- `sect` (заголовок группы с линейкой), `lead`, `note`/`band_note` (выноска у
  нижнего поля), `plate` (плашка-правило), `para` — каркас полосы, все → низ.
- `steps_ribbon(s, band, steps, head=("ШАГ","ЧТО ЗВУЧИТ…"), pictos=[…])` —
  вертикальная лента шагов: номер-бейдж на оси-таймлайне, пиктограмма, название
  слева, «что звучит» справа.
- `pyramid(s, band, levels)` — трапеции уровней слева, реквизиты справа, высота
  строки = максимум по реквизитам (метрика с отбивками) — текст не перелезает
  через линейку следующего уровня.
- `picto(kind)` — пиктограммы нативными фигурами: `question chevron warning down
  fork prohibition bubble_q dialog decision`.
- Грабли: QA считает пересечение ЛЮБОЙ линии с текстовым блоком — таймлайн ленты
  рисуй badge→badge, линейку колонок — до последней построчной линейки; для
  треугольника — `MSO_SHAPE.ISOSCELES_TRIANGLE`.
- Регрессия после правки модулей: `python _service\skills\universal-pptx\scripts
  \smoke_book.py %TEMP%\smoke.pptx` → `deck.py qa smoke.pptx -t book` (ждём 0/0)
  → рендер и визуальный смотр.

### Рецепты мелких мест (без новых компонентов)

- **пиктограммы вне книжной темы**: `picto(kind)` из `pptx_book` геометрически
  независима — те же вопрос/шеврон/предупреждение/развилка/запрет/облачко/диалог/
  решение идут в любой теме: `picto("warning")(s, cx, cy, sc)` — центр и размер в см;
- **формулы**: matplotlib mathtext через `deck_mpl` — `ax.text(x, y, r"$E=mc^2$",
  usetex=False)`, PNG 300 dpi вставляется `chart_image()`; кириллица в формулах —
  только вне `$…$`;
- **xlsx/csv как материалы**: читать на месте (`openpyxl.load_workbook(...,
  data_only=True)`), в `facts.md` источник — «файл, лист, диапазон ячеек»;
  сводные числа — статус derived с формулой;
- **дампы материалов живут с проектом**: `deck.py dump … -o <проект>\_build\dump.txt`,
  ссылки в facts.md — номера строк этого файла (temp-дампы недопустимы — сгорают);
- **иллюстрация как файл**: однослайдовый pptx со схемой → `render --trim -W 2400`
  (обрезка до контента, печатное разрешение); без фона — `--fast --transparent`
  (headless-рендер ArtCraft, RGBA); для веб — `-W 1500` без trim.

## Опционально: комплект ArtCraft (github.com/storytold)

Headless-CLI приложений на Rust (`*-cli.exe`, ищутся через PATH / `DECK_ARTCRAFT_DIR`):
- **deckcraft-cli** — открытый рендер .pptx→PNG без Office за доли секунды (~100×
  быстрее COM) — это `render --fast` для внутреннего цикла; **финальная приёмка —
  COM/soffice** (чистая комната, расхождения с PowerPoint возможны); есть
  `convert deck.pptx deck.pdf` (отдача PDF без Office), `.md`-очертание, MCP-сервер
  `mcp` (stdio) и JSON-команды правки `run --cmd 'slide.new={...}'`;
- **pdfcraft-cli** — `text|info|render|combine|extract|check|ocr`: им пользуется
  `deck.py dump` для pdf (крепче pypdf), пригоден для приёма pdf-материалов;
- **wordcraft-cli** — `text|info|convert docx→md/pdf` с `--strict` (падает, если
  глиф не рисуется или картинка бита — бесплатная проверка материалов);
- photocraft/vectorcraft/lightcraft — ручная доводка растровых/векторных
  материалов человеком, в конвейер не встраивать.

Расположение скилла: каноническая копия — `_service\skills\universal-pptx\`;
`.opencode\skills\` и `.claude\skills\` — junction-ссылки на неё (пересоздавать
при переносе: `New-Item -ItemType Junction -Path .opencode\skills\universal-pptx
-Target <корень>\_service\skills\universal-pptx`). Скилл переносим в любую папку:
пути ищутся автоматически, для нестандартной раскладки — env `DECK_SERVICE_DIR` /
`DECK_SKILL_SCRIPTS`. Все команды независимы от харнесса и рабочей папки.
