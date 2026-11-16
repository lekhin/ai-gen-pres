# -*- coding: utf-8 -*-
"""mermaid flowchart → схема движком pptx_diagram.py (стиль — активная тема).

mermaid в документах проекта (ARCHITECTURE.md и т.п.) — источник СОДЕРЖАНИЯ
схем; конвертер парсит flowchart, раскладывает ранговой сеткой и выдаёт:
  json    (по умолчанию) — блок {"type":"diagram",…} для plan.json (см. plan.py);
  --emit py              — каркас кода group()/node()/edge() для build.py;
  --preview Ф.pptx       — готовый однослайдовый pptx под быструю проверку render'ом.

РАСКЛАДКА — КАРКАС, не финал: после сборки обязательны qa + рендер + визуальная
проверка (SKILL.md «Схемы»); метки в межблочные зазоры >= 2 см и обходы
сплошная×сплошная править вручную (json/код для этого и выдаются).

Поддерживается: flowchart|graph LR|RL|TD|TB|BT; subgraph (вложенные, связи
между ними); формы id[..] id(..) id([..]) id[(..)] id{..} id{{..}} id[[..]]
id[/../] ; связи --> --- -.-> -.- ==> <--> с подписями |текст| и -- текст -->;
цепочки `A & B --> C & D` (fan-in/out); роли узлов классами: id:::data или
`class a,b model` (app|data|model|people|plain|planned, синонимы user/onprem/
llm/db/future… → см. CLASS_ALIAS); %%комментарии. [(..)] без класса — роль data.
Игнорируется: цвета classDef/style/linkStyle (окраска — из ролей палитры),
init-директивы, direction внутри subgraph, click.

Запуск:
  python mmd2diagram.py diag.mmd [-o diag.json] [--title "Схема"] [--legend "solid=запросы,dash=OIDC"]
  python mmd2diagram.py diag.mmd --emit py
  python mmd2diagram.py diag.mmd --preview out.pptx
  deck.py mmd diag.mmd …
"""
import argparse
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pptx_diagram import _text_w_cm  # noqa: E402  оценка ширины текста (см)

ROLES = ("app", "data", "model", "people", "plain", "planned")
# частые синонимы классов в mermaid-исходниках → роли палитры
CLASS_ALIAS = {"user": "people", "users": "people", "person": "people",
               "people": "people", "client": "people", "operator": "people",
               "external": "plain", "obs": "plain", "observ": "plain",
               "future": "planned", "plan": "planned", "todo": "planned",
               "db": "data", "storage": "data", "store": "data",
               "llm": "model", "onprem": "app", "internal": "app"}
MARGIN, CONTENT_TOP, FOOTER_TOP = 1.5, 5.6, 17.55
GRID_W = 33.866666 - 2 * MARGIN
GAP_X, GAP_Y = 3.4, 0.9          # зазоры между рангами/узлами внутри ранга, см
PAD = 0.6                        # отступ периметра от узлов, см
PAD_TOP = 1.3                    # сверху периметра — место под ярлык
V_DIR = ("TB", "TD", "BT")       # вертикальные направления mermaid

NODE_RE = re.compile(r'''
 (?P<id>[A-Za-zА-Яа-яЁё_][\w\-.]*)
 (?:\s*:::\s*(?P<cls>[\w\-]+))?
 (?:\s*(?:
    \[\((?P<db>.*?)\)\]
   |\[\[(?P<sub>[^\]]*?)\]\]
   |\(\[(?P<stad>[^\]]*?)\]\)
   |\(\((?P<circ>[^()]*?)\)\)
   |\{\{(?P<hex>[^{}]*?)\}\}
   |\{(?P<dia>[^{}]*?)\}
   |\[(?P<io>/[^/\]]*?/|\\[^\\\]]*?\\|(?P<rect>[^\]]*?))\]
   |\((?P<round>[^()]*?)\)
 )?)''', re.VERBOSE)

OP_RE = re.compile(
    r'(?P<op><-->|<\-\.->|-->|-\.->|---|--|-\.|==>|~~~)'
    r'(?:\s*\|\s*(?P<label>[^|]*?)\s*\|)?')

SHAPE = {"db": "db", "sub": "pill", "stad": "pill", "circ": "pill",
         "hex": "rect", "dia": "diamond", "io": "io", "rect": "rect",
         "round": "pill"}


def clean(text):
    """Метка mermaid → (title, body): кавычки, <br>, &quot;."""
    if text is None:
        return None, None
    t = text.strip().strip('"').strip("'").strip()
    t = t.replace("&quot;", '"').replace("&amp;", "&").replace("&lt;", "<") \
         .replace("&gt;", ">")
    parts = [p.strip() for p in re.split(r"<br\s*/?>", t) if p.strip()]
    if not parts:
        return "", None
    return parts[0], (" · ".join(parts[1:]) or None) if len(parts) > 1 else None


def clean_edge_label(t):
    """Подпись связи: кавычки, <br> → пробел (edge_label — однострочный)."""
    if not t:
        return None
    t = t.strip().strip('"').strip("'")
    t = re.sub(r"<br\s*/?>", " ", t)
    t = re.sub(r"\s+", " ", t).strip()
    return t or None


class Model:
    def __init__(self):
        self.direction = "LR"
        self.nodes = {}        # id -> {title,body,shape,role,group}
        self.groups = []       # {id,title,parent,dash}
        self.edges = []        # {src,dst,dash,w_pt,label,bi}
        self.classdefs = set()
        self.warnings = []

    def group(self, gid):
        for g in self.groups:
            if g["id"] == gid:
                return g


def role_of(cls):
    """Имя mermaid-класса → роль палитры (или None: цвета classDef игнорируем)."""
    if not cls:
        return None
    c = cls.lower()
    if c in ROLES:
        return c
    return CLASS_ALIAS.get(c)


def parse(text, block=None):
    m = Model()
    text, ws = extract_mermaid(text, block)
    m.warnings.extend(ws)
    stack = []                                   # вложенность subgraph
    gcnt = 0

    def add_node(nid, kind=None, label=None, cls=None):
        if label is None and kind is None and nid not in m.nodes and m.group(nid):
            return          # id уже subgraph — это ссылка на группу в связи, не узел
        n = m.nodes.setdefault(nid, {"title": nid.replace("_", " "),
                                     "body": None, "shape": "rect",
                                     "role": "app", "group": stack[-1] if stack else None})
        if label is not None:
            n["title"], n["body"] = clean(label)
        if kind:
            n["shape"] = SHAPE[kind]
            if kind == "db" and n["role"] == "app":
                n["role"] = "data"
        r = role_of(cls)
        if r:
            n["role"] = r

    for raw in text.splitlines():
        line = raw.strip()
        if not line or line.startswith(("%%", "init", "---")):
            continue
        line = line.split("%%", 1)[0].strip()
        line = re.sub(r'-\.\s+([^|]+?)\s+\.\->', r'-.->|\1|', line)
        line = re.sub(r'(?<=[\w\)\]\}])\s--\s+(?!\|)(.+?)\s+-->', r' -->|\1|', line)

        mo = re.match(r'(?:flowchart|graph)\s+(LR|RL|TB|TD|BT)\b', line)
        if mo:
            m.direction = mo.group(1)
            continue
        if line in ("end", "end;"):
            if stack:
                stack.pop()
            continue
        mo = re.match(r'subgraph\s+(?:(?P<gid>[\w\-.]+)\s*(?:\[([^\]]*)\]|\(([^)]*)\))|(?P<title>.+))', line)
        if mo:
            gcnt += 1
            gid = mo.group("gid") or f"_g{gcnt}"
            title = (mo.group(2) or mo.group(3) or mo.group("title") or gid).strip()
            m.groups.append({"id": gid, "title": clean(title)[0],
                             "parent": stack[-1] if stack else None})
            stack.append(gid)
            continue
        mo = re.match(r'classDef\s+([\w\-]+)', line)
        if mo:
            m.classdefs.add(mo.group(1))
            continue
        mo = re.match(r'class\s+([\w\-,\s]+?)\s+([\w\-]+)\s*;?$', line)
        if mo:
            cls = role_of(mo.group(2))
            for nid in re.split(r'[,\s]+', mo.group(1)):
                if nid in m.nodes and cls:
                    m.nodes[nid]["role"] = cls
                elif nid not in m.nodes:
                    m.warnings.append(f"class для неизвестного узла {nid}")
            continue
        if re.match(r'(style|linkStyle|click|direction)\s', line):
            continue

        pos, seq = 0, []          # seq: ('n',nid) | ('e',атрибуты) — для цепочек &
        while pos < len(line):
            if line[pos].isspace() or line[pos] in ";&":
                pos += 1
                continue
            mo = OP_RE.match(line, pos)
            if mo:
                op = mo.group("op")
                seq.append(("e", {"dash": "dash" if "-." in op else None,
                                  "w_pt": 2.5 if "==" in op else 1.25,
                                  "bi": "<" in op, "hidden": op == "~~~",
                                  "label": clean_edge_label(mo.group("label"))}))
                pos = mo.end()
                continue
            mo = NODE_RE.match(line, pos)
            if mo and mo.group("id"):
                kind = next((k for k in SHAPE if mo.group(k) is not None
                             and not (k == "io" and mo.group("rect") is not None)), None)
                if kind == "io" and mo.group("rect") is None:
                    label = mo.group("io").strip("/\\ ")
                else:
                    label = mo.group(kind) if kind else None
                nid = mo.group("id")
                add_node(nid, kind, label, mo.group("cls"))
                pos = mo.end()
                mc = re.match(r'\s*:::\s*([\w\-]+)', line[pos:])  # :::class ПОСЛЕ формы
                if mc:
                    r = role_of(mc.group(1))
                    if r:
                        m.nodes[nid]["role"] = r
                    pos += mc.end()
                seq.append(("n", nid))
                continue
            m.warnings.append(f"не разобрано: {line[pos:pos+30]!r}")
            break

        # seq → группы узлов, разделённые операторами; каждый оператор связывает
        # все узлы группы слева со всеми справа (mermaid `A & B --> C & D`)
        groups, links, cur = [], [], []
        for tok in seq:
            if tok[0] == "n":
                cur.append(tok[1])
            else:
                if cur:
                    groups.append(cur); cur = []
                links.append(tok[1])
        if cur:
            groups.append(cur)
        for i in range(min(len(links), len(groups) - 1)):
            if links[i].get("hidden"):        # ~~~ — разделитель, связи не создаём
                continue
            for s in groups[i]:
                for d in groups[i + 1]:
                    if s == d:
                        m.warnings.append(f"само́связь {s} пропущена")
                        continue
                    if d in m.nodes or m.group(d):
                        m.edges.append(dict(links[i], src=s, dst=d))
    # связи между subgraph: mermaid разрешает source подграфов — не умеем
    for e in list(m.edges):
        for k in ("src", "dst"):
            if e[k] not in m.nodes and m.group(e[k]):
                m.warnings.append(
                    f"связь {e['src']}→{e['dst']} между subgraph — править вручную")
    # mermaid-прокладки pad[" "]:::pad (пустые метки) не рисуем, связи с ними не нужны
    for nid in [i for i, n in m.nodes.items() if not (n["title"] or "").strip()]:
        del m.nodes[nid]
        m.edges = [e for e in m.edges if e["src"] != nid and e["dst"] != nid]
    return m


# ---------------- ранжирование и порядок в рангах ----------------
def ranks_edges(m):
    """Ранги узлов (longest path) + рёбра ранжирования (без обратных)."""
    ids = list(m.nodes)
    color = {i: 0 for i in ids}
    fwd = []

    def dfs(u):
        color[u] = 1
        for e in m.edges:
            for v, is_fwd_edge in ((e["dst"], True),):
                if e["src"] != u:
                    continue
                if v not in m.nodes:
                    continue
                if color[v] == 1:
                    m.warnings.append(f"цикл {u}→{v}: ранг по обходному ребру")
                    continue
                fwd.append((u, v))
                if color[v] == 0:
                    dfs(v)
        color[u] = 2

    for i in ids:
        if color[i] == 0:
            dfs(i)
    rank = {i: 0 for i in ids}
    for _ in range(len(ids)):                      # релаксация longest path
        upd = False
        for u, v in fwd:
            if rank[v] < rank[u] + 1:
                rank[v], upd = rank[u] + 1, True
        if not upd:
            break
    # изолированные узлы (без рёбер) лежат в ранге 0 и раздувают первую колонку:
    # ставим их к медианному рангу соседей по subgraph (если они там есть)
    touched = {u for u, v in fwd} | {v for u, v in fwd}
    grp_ranks = {}
    for i, n in m.nodes.items():
        g = n.get("group")
        if g and i in touched:
            grp_ranks.setdefault(g, []).append(rank[i])
    import statistics
    for i in ids:
        g = m.nodes[i].get("group")
        if i not in touched and g and grp_ranks.get(g):
            rank[i] = int(statistics.median(grp_ranks[g]))
    return rank, fwd


def order_ranks(m, rank, fwd):
    nr = max(rank.values(), default=0) + 1
    order = [[] for _ in range(nr)]
    for i in m.nodes:
        order[rank[i]].append(i)
    pos = {i: k for r in order for k, i in enumerate(r)}
    for sweep in range(3):                         # barycenter, вперёд-назад
        pairs = fwd if sweep % 2 == 0 else [(v, u) for u, v in fwd]
        for r in range(1, nr):
            bary = {}
            for i in order[r]:
                ps = [pos[u] for u, v in pairs if v == i]
                bary[i] = sum(ps) / len(ps) if ps else pos[i]
            order[r].sort(key=lambda i: (bary[i], pos[i]))
            for k, i in enumerate(order[r]):
                pos[i] = k
    return order


# ---------------- раскладка ----------------
def measure(m):
    for n in m.nodes.values():
        w = _text_w_cm(n["title"], 10) + 0.9
        if n["body"]:
            w = max(w, _text_w_cm(n["body"], 8) / 1.9 + 1.0)
        n["w"] = min(8.0, max(3.0, w))
        n["h"] = 1.2 if not n["body"] else 1.9
        if n["shape"] == "diamond":
            n["w"] += 2.6          # у ромба текстовая область ~в 2 раза уже описанного
            n["h"] = max(n["h"], 2.2)
        n["w_min"], n["h_min"] = n["w"], n["h"]


def layout_lr(m, order):
    colw = [max((m.nodes[i]["w"] for i in r), default=3.0) for r in order]
    colh = [sum(m.nodes[i]["h"] for i in r) + GAP_Y * (len(r) - 1) for r in order]
    maxh = max(colh, default=0)
    x = 0.0
    xs = []
    for w in colw:
        xs.append(x)
        x += w + GAP_X
    for r, nodes in enumerate(order):
        y = (maxh - colh[r]) / 2
        cx = xs[r] + (colw[r] - 0) / 2
        for i in nodes:
            n = m.nodes[i]
            n["x"] = xs[r] + (colw[r] - n["w"]) / 2
            n["y"] = y
            y += n["h"] + GAP_Y
    return xs, colw


def group_bounds(m):
    """Периметры по узлам/вложенным периметрам; узлы внутри group id."""
    for g in reversed(m.groups):
        bb = [n for n in m.nodes.values() if n["group"] == g["id"]]
        bb += [q for q in m.groups
               if q.get("parent") == g["id"] and "x" in q]
        if not bb:
            m.warnings.append(f"subgraph {g['id']} без узлов — периметр пропущен")
            g["empty"] = True
            continue
        g["x"] = min(b["x"] for b in bb) - PAD
        g["y"] = min(b["y"] for b in bb) - PAD_TOP
        g["w"] = max(b["x"] + b["w"] for b in bb) - g["x"] + PAD
        g["h"] = max(b["y"] + b["h"] for b in bb) - g["y"] + PAD


def box_of(m, key):
    n = m.nodes.get(key)
    if n:
        return n
    g = m.group(key)
    if g and not g.get("empty"):
        return g
    return None


# ---------------- расталкивание периметров ----------------
def anc_chain(m, gid):
    out = set()
    g = m.group(gid)
    while g and g.get("parent"):
        out.add(g["parent"])
        g = m.group(g["parent"])
    return out


def _overlap(a, b):
    return (a["x"] < b["x"] + b["w"] - 0.05 and a["x"] + a["w"] > b["x"] + 0.05
            and a["y"] < b["y"] + b["h"] - 0.05 and a["y"] + a["h"] > b["y"] + 0.05)


def _move_group(m, g, dy):
    """Сдвиг периметра и всего его содержимого вниз на dy."""
    inside = [q for q in m.groups
              if g["id"] in anc_chain(m, q["id"]) or q["id"] == g["id"]]
    for q in inside:
        q["y"] += dy
    for n in m.nodes.values():
        if n["group"] and (n["group"] == g["id"]
                           or g["id"] in anc_chain(m, n["group"])):
            n["y"] += dy


def resolve_collisions(m):
    """Внешние узлы и соседние периметры не должны лежать на чужом периметре."""
    for _ in range(8):
        for g in m.groups:
            g.pop("x", None)
        group_bounds(m)
        gs = [g for g in m.groups if not g.get("empty")]
        moved = False
        for n in m.nodes.values():
            keep = ({n["group"]} | anc_chain(m, n["group"])) if n["group"] else set()
            for g in gs:
                if g["id"] in keep or not _overlap(n, g):
                    continue
                n["y"] = g["y"] + g["h"] + 0.45
                moved = True
        for i, g1 in enumerate(gs):
            for g2 in gs[i + 1:]:
                rel = (g2["id"] in anc_chain(m, g1["id"])
                       or g1["id"] in anc_chain(m, g2["id"]))
                if not rel and _overlap(g1, g2):
                    _move_group(m, g2, g1["y"] + g1["h"] + 0.6 - g2["y"])
                    moved = True
        if not moved:
            return


# ---------------- маршрутизация связей (в LR-пространстве) ----------------
def est_w(text):
    return min(4.5, _text_w_cm(text or "", 6.5) + 0.2) if text else 0


def _fan(routes):
    """Веера: разносим якоря выхода/входа по грани узла и стаггерим локти."""
    outs, ins = {}, {}
    for r in routes:
        if r["kind"] == "fwd":
            outs.setdefault(id(r["a"]), []).append(r)
            ins.setdefault(id(r["b"]), []).append(r)
    for lst in outs.values():
        lst.sort(key=lambda r: r["b"]["y"] + r["b"]["h"] / 2)
        n = len(lst)
        for i, r in enumerate(lst):
            t = 0.5 if n == 1 else 0.3 + 0.4 * i / (n - 1)
            r["ay"] = r["a"]["y"] + r["a"]["h"] * t
            r["stag"] = 0.0 if n == 1 else (i - (n - 1) / 2) * 0.35
    for lst in ins.values():
        lst.sort(key=lambda r: r["a"]["y"] + r["a"]["h"] / 2)
        n = len(lst)
        for i, r in enumerate(lst):
            t = 0.5 if n == 1 else 0.3 + 0.4 * i / (n - 1)
            r["by"] = r["b"]["y"] + r["b"]["h"] * t


def _segs(routes, skip=None):
    segs = []
    for i, r in enumerate(routes):
        if skip is not None and i == skip:
            continue
        for (x1, y1), (x2, y2) in zip(r["pts"], r["pts"][1:]):
            segs.append((x1, y1, x2, y2))
    return segs


def _hits(segs, boxes, x, y, w, h=0.35):
    for x1, y1, x2, y2 in segs:
        if y1 == y2:
            lo, hi = sorted((x1, x2))
            if y - 0.02 <= y1 <= y + h + 0.02 and lo < x + w and hi > x:
                return True
        else:
            lo, hi = sorted((y1, y2))
            if x - 0.02 <= x1 <= x + w + 0.02 and lo < y + h and hi > y:
                return True
    for b in boxes:
        if _overlap({"x": x, "y": y, "w": w, "h": h}, b):
            return True
    return False


def _pick(cands, segs, boxes):
    for c in cands:
        if not _hits(segs, boxes, c["x"], c["y"], c["w"]):
            return c
    return cands[0]


def route(m):
    """Маршруты связей. Логика пишется в LR-пространстве; для TB узлы
    предварительно меняются x/y местами, а точки возвращаются обратно."""
    V = m.direction in V_DIR

    def sbox(b):
        return {"x": b["y"], "y": b["x"], "w": b["h"], "h": b["w"]} if V else b

    def spt(p):
        return (p[1], p[0]) if V else p

    out = []
    for e in m.edges:
        a, b = box_of(m, e["src"]), box_of(m, e["dst"])
        if not a or not b:
            continue
        a, b = sbox(a), sbox(b)
        kind = ("fwd" if a["x"] + a["w"] < b["x"] - 0.1 else
                "back" if b["x"] + b["w"] < a["x"] - 0.1 else "side")
        out.append(dict(e, a=a, b=b, kind=kind))
    _fan(out)
    for r in out:
        a, b = r["a"], r["b"]
        if r["kind"] == "fwd":
            ax = a["x"] + a["w"] + 0.03
            bx = b["x"] - 0.03
            ay = r.get("ay", a["y"] + a["h"] / 2)
            by = r.get("by", b["y"] + b["h"] / 2)
            if abs(ay - by) < 0.2:
                pts = [(ax, ay), (bx, by)]
            else:
                mid = (ax + bx) / 2 + r.get("stag", 0.0)
                pts = [(ax, ay), (mid, ay), (mid, by), (bx, by)]
        elif r["kind"] == "back":
            p0 = (a["x"] + a["w"] / 2, a["y"] - 0.03)
            p1 = (b["x"] + b["w"] / 2, b["y"] - 0.03)
            lane = min(p0[1], p1[1]) - 0.8
            pts = [p0, (p0[0], lane), (p1[0], lane), p1]
        else:
            p0 = (a["x"] + a["w"] + 0.03, a["y"] + a["h"] / 2)
            p1 = (b["x"] + b["w"] + 0.03, b["y"] + b["h"] / 2)
            lane = max(p0[0], p1[0]) + 1.4
            pts = [p0, (lane, p0[1]), (lane, p1[1]), p1]
        r["pts"] = [spt(p) for p in pts]
    return out


def place_labels(m, routes):
    """Метки связей — в последнее чистое место ПОСЛЕ финальной геометрии
    (refit/transpose), иначе скейл заново посадит метки на связи."""
    boxes = list(m.nodes.values()) + [g for g in m.groups if not g.get("empty")]
    for i, r in enumerate(routes):
        if not r["label"]:
            continue
        lab_w = est_w(r["label"])
        (x1, y1), (bx, by) = r["pts"][-2], r["pts"][-1]
        ax0, ay0 = r["pts"][0]
        segs = _segs(routes)       # свои сегменты тоже: метка не должна лезть
        #                            и на собственную вертикальную стойку
        if abs(y1 - by) < 0.02:                          # вход по горизонтали
            cands = [
                {"x": bx - lab_w - 0.2, "y": by - 0.55, "w": lab_w},
                {"x": (x1 + bx) / 2 - lab_w / 2, "y": by - 0.55, "w": lab_w},
                {"x": (ax0 + bx) / 2 - lab_w / 2, "y": min(ay0, by) - 0.55,
                 "w": lab_w},
                {"x": (ax0 + bx) / 2 - lab_w / 2, "y": max(ay0, by) + 0.2,
                 "w": lab_w},
            ]
        else:                                            # вход по вертикали (TB)
            cands = [
                {"x": bx + 0.15, "y": (y1 + by) / 2 - 0.17, "w": lab_w},
                {"x": bx + 0.15, "y": min(y1, by) - 0.5, "w": lab_w},
                {"x": bx - lab_w - 0.15, "y": (y1 + by) / 2 - 0.17, "w": lab_w},
            ]
        r["label"] = dict(_pick(cands, segs, boxes), text=r["label"])


# ---------------- подгонка в область контента слайда ----------------
def fit(m, legend_h=0.0):
    """Единый скейл+центровка ГЕОМЕТРИИ (узлы/периметры). Маршрутизация —
    ПОСЛЕ этого: связи строим по финальным боксам, иначе план пересчитает
    якоря по финальным узлам и появятся диагонали сквозь текст."""
    objs = list(m.nodes.values()) + [g for g in m.groups if not g.get("empty")]
    if not objs:
        return 1.0
    x0 = min(o["x"] for o in objs)
    y0 = min(o["y"] for o in objs)
    x1 = max(o["x"] + o["w"] for o in objs)
    y1 = max(o["y"] + o["h"] for o in objs)
    aw = GRID_W - 3.2        # запас на обводные левые справа и метки к ним
    ah = FOOTER_TOP - 0.35 - CONTENT_TOP - legend_h - 1.6   # на тыловые левые
    s = min(aw / (x1 - x0), ah / (y1 - y0), 1.4)
    dx = MARGIN + (aw - (x1 - x0) * s) / 2 - x0 * s
    dy = CONTENT_TOP + (ah - (y1 - y0) * s) / 2 - y0 * s
    # плотная схема: если геометрия сжимается (<×1.0), текст узлов мельчает
    # пропорционально — иначе узлы упираются в минимум под 10pt и наезжают друг на друга
    tfs = max(6.0, round(min(1.0, s) * 10 * 10) / 10)   # title кегль, pt
    bfs = max(5.0, round(tfs * 0.8 * 10) / 10)                 # body кегль, pt
    kfs = tfs / 10.0
    for o in objs:
        o["x"], o["y"] = o["x"] * s + dx, o["y"] * s + dy
        o["w"], o["h"] = o["w"] * s, o["h"] * s
    for n in m.nodes.values():
        n["title_size"], n["body_size"] = tfs, bfs
        # текст фиксированного кегля: узел не уменьшать ниже измеренного минимума
        dw = max(0.0, n["w_min"] * kfs - n["w"]) / 2
        dh = max(0.0, n["h_min"] * (bfs / 8.0) - n["h"]) / 2
        n["x"], n["y"] = n["x"] - dw, n["y"] - dh
        n["w"], n["h"] = n["w"] + 2 * dw, n["h"] + 2 * dh
    for _ in range(6):            # расширение могло столкнуть соседей — раздвигаем
        moved = False
        ns = list(m.nodes.values())
        for a in ns:
            for b in ns:
                if a is b or not _overlap(a, b):
                    continue
                ov_x = min(a["x"] + a["w"], b["x"] + b["w"]) - max(a["x"], b["x"])
                ov_y = min(a["y"] + a["h"], b["y"] + b["h"]) - max(a["y"], b["y"])
                if ov_x < ov_y:   # разводим по оси меньшего наезда
                    b["x"] = (a["x"] + a["w"] + 0.3) if b["x"] >= a["x"] \
                        else a["x"] - b["w"] - 0.3
                else:
                    b["y"] = (a["y"] + a["h"] + 0.3) if b["y"] >= a["y"] \
                        else a["y"] - b["h"] - 0.3
                moved = True
        if not moved:
            break
    gs = [g for g in m.groups if not g.get("empty")]
    for g in reversed(gs):        # периметры — изнутри наружу: дети, затем родители
        kids = [n for n in m.nodes.values() if n["group"] and
                (n["group"] == g["id"] or g["id"] in anc_chain(m, n["group"]))]
        kids += [q for q in gs if g["id"] in anc_chain(m, q["id"])]
        for b in kids:
            g["x"] = min(g["x"], b["x"] - PAD)
            g["y"] = min(g["y"], b["y"] - PAD_TOP)
            r_, b_ = max(g["x"] + g["w"], b["x"] + b["w"] + PAD), \
                     max(g["y"] + g["h"], b["y"] + b["h"] + PAD)
            g["w"], g["h"] = r_ - g["x"], b_ - g["y"]
    if s < 0.75:
        m.warnings.append(f"схема сжата ×{s:.2f} — проверяй читаемость, дробили схему?")
    return s


def transpose(m):
    """LR-раскладка → вертикальная (TB): обмен x/y у узлов и периметров.
    Делается ДО подгонки и маршрутизации — те работают с финальной геометрией."""
    for o in list(m.nodes.values()) + [g for g in m.groups if not g.get("empty")]:
        o["x"], o["y"] = o["y"], o["x"]
        o["w"], o["h"] = o["h"], o["w"]
    for g in m.groups:               # ярлык периметра всегда сверху — освобождает
        if not g.get("empty"):       # сверху место уже в финальной ориентации
            g["y"] -= PAD_TOP - PAD
            g["h"] += PAD_TOP - PAD


def refit(m, routes, legend_h=0.0):
    """Финальная центровка/скейл ПОСЛЕ маршрутизации и transpose: учитывает
    левые связей и метки, чтобы ничего не вылезло за область контента."""
    objs = list(m.nodes.values()) + [g for g in m.groups if not g.get("empty")]
    if not objs:
        return
    xs, ys = [], []
    for o in objs:
        xs += [o["x"], o["x"] + o["w"]]
        ys += [o["y"], o["y"] + o["h"]]
    for r in routes:
        for x, y in r["pts"]:
            xs.append(x)
            ys.append(y)
    x0, x1, y0, y1 = min(xs), max(xs), min(ys), max(ys)
    aw = GRID_W - 1.0
    ah = FOOTER_TOP - 0.35 - CONTENT_TOP - legend_h
    s = min(aw / (x1 - x0), ah / (y1 - y0), 1.4)
    if s >= 1:
        return               # связи влезли в отступы fit — геометрию не трогаем
    dx = MARGIN + (aw - (x1 - x0) * s) / 2 - x0 * s
    dy = CONTENT_TOP + (ah - (y1 - y0) * s) / 2 - y0 * s
    for o in objs:
        o["x"], o["y"] = o["x"] * s + dx, o["y"] * s + dy
        o["w"], o["h"] = o["w"] * s, o["h"] * s
    for r in routes:
        r["pts"] = [(px * s + dx, py * s + dy) for px, py in r["pts"]]
    if s < 0.75:
        m.warnings.append(f"схема сжата ×{s:.2f} — проверяй читаемость, дробили схему?")


def prepare(mmd_text, legend_items=None, direction=None, block=None):
    """parse → layout → transpose → fit → route → refit → метки."""
    m = parse(mmd_text, block)
    if direction:
        m.direction = direction      # переопределение TB/LR (глубокие схемы в TB сплющивает)
    measure(m)
    rank, fwd = ranks_edges(m)
    order = order_ranks(m, rank, fwd)
    global GAP_X, GAP_Y
    if len(order) >= 6 or len(m.nodes) >= 12:   # плотная схема — экономим место
        GAP_X, GAP_Y = 2.2, 0.55
    layout_lr(m, order)
    resolve_collisions(m)          # свежие периметры пересчитываются внутри
    if m.direction in V_DIR:
        transpose(m)               #orientation фиксируем до подгонки и связей
    fit(m, legend_h=1.0 if legend_items else 0.0)
    routes = route(m)              # по финальным боксам (ось учитывает)
    refit(m, routes, legend_h=1.0 if legend_items else 0.0)
    place_labels(m, routes)        # метки — по финальной геометрии
    return m, routes, (legend_items or [])


# ---------------- вывод ----------------
def rnd(v):
    return round(v + 1e-9, 2)


def to_block(m, routes, legend_items):
    nodes = [{"id": i, "x": rnd(n["x"]), "y": rnd(n["y"]), "w": rnd(n["w"]),
              "h": rnd(n["h"]), "title": n["title"], "role": n["role"],
              "shape": n["shape"], **({"body": n["body"]} if n["body"] else {}),
              **({"title_size": n["title_size"]} if n.get("title_size", 10) < 10 else {}),
              **({"body_size": n["body_size"]} if n.get("body_size", 8) < 8 else {})}
             for i, n in m.nodes.items()]
    groups = [{"x": rnd(g["x"]), "y": rnd(g["y"]), "w": rnd(g["w"]),
               "h": rnd(g["h"]), "label": g["title"]}
              for g in m.groups if not g.get("empty")]
    edges = []
    for r in routes:
        e = {}
        if r["dash"]:
            e["dash"] = r["dash"]
        if r["w_pt"] != 1.25:
            e["w_pt"] = r["w_pt"]
        # ПОЛНАЯ полилиния с концами: план по from/to якорит связи в центры
        # граней, а конвертер разносит якоря по грани — с from/to были бы диагонали
        e["points"] = [[rnd(x), rnd(y)] for x, y in r["pts"]]
        if r["label"]:
            lb = r["label"]
            e["label"] = {"text": lb["text"], "x": rnd(lb["x"]),
                          "y": rnd(lb["y"]), "w": rnd(lb["w"])}
        edges.append(e)
    block = {"type": "diagram"}
    if groups:
        block["groups"] = groups
    block["nodes"] = nodes
    block["edges"] = edges
    if legend_items:
        block["legend"] = {"x": rnd(MARGIN), "y": rnd(FOOTER_TOP - 0.55),
                           "items": [list(i) for i in legend_items]}
    return block


def to_py(m, routes, legend_items):
    L = ["# --- сгенерировано mmd2diagram.py — каркас, править по канонам "
         "SKILL.md «Схемы\" ---",
         "from pptx_diagram import node, edge, edge_label, group, legend"]
    var = {i: re.sub(r'\W', '_', i) for i in m.nodes}
    for g in [g for g in m.groups if not g.get("empty")]:
        L.append(f'group(s, {rnd(g["x"])}, {rnd(g["y"])}, {rnd(g["w"])}, '
                 f'{rnd(g["h"])}, {json.dumps(g["title"], ensure_ascii=False)})')
    for i, n in m.nodes.items():
        body = json.dumps(n["body"], ensure_ascii=False) if n["body"] else "None"
        L.append(f'{var[i]} = node(s, {rnd(n["x"])}, {rnd(n["y"])}, {rnd(n["w"])}, '
                 f'{rnd(n["h"])}, {json.dumps(n["title"], ensure_ascii=False)}, '
                 f'{body}, role="{n["role"]}", shape="{n["shape"]}")')
    for r in routes:
        pts = ", ".join(f"({rnd(x)}, {rnd(y)})" for x, y in r["pts"])
        kw = ""
        if r["dash"]:
            kw += ', dash="dash"'
        if r["w_pt"] != 1.25:
            kw += f', w_pt={r["w_pt"]}'
        L.append(f"edge(s, [{pts}]{kw})")
        if r["label"]:
            lb = r["label"]
            L.append(f'edge_label(s, {rnd(lb["x"])}, {rnd(lb["y"])}, '
                     f'{rnd(lb["w"])}, {json.dumps(lb["text"], ensure_ascii=False)}, '
                     f'size=6.5)')
    if legend_items:
        items = ", ".join(f'("{k}", {json.dumps(t, ensure_ascii=False)})'
                          for k, t in legend_items)
        L.append(f'legend(s, {rnd(MARGIN)}, {rnd(FOOTER_TOP - 0.55)}, [{items}])')
    return "\n".join(L)


def preview_pptx(block, out, title):
    from deck_paths import service_file
    import deck_theme
    from pptx_helpers import load_template, layouts, new_slide, ph, drop_ph
    from plan import render_diagram
    prs = load_template(deck_theme.template_path())
    _, L_CONT, _ = layouts(prs)
    s = new_slide(prs, L_CONT, header=title or "Схема")
    ph(s, 0).text_frame.text = title or "Схема"
    drop_ph(s, 1)
    render_diagram(s, block)
    for idx in (1, 2):
        p = ph(s, idx)
        if p is not None and p.has_text_frame and not p.text_frame.text.strip():
            drop_ph(s, idx)
    prs.save(out)
    return out


def parse_legend(s):
    if not s:
        return None
    items = []
    for part in s.split(","):
        kind, _, text = part.partition("=")
        items.append((kind.strip(), text.strip()))
    return items


def extract_mermaid(text, block=None):
    """Из .md/.mdc взять mermaid-блок flowchart/graph (sequenceDiagram и
    прочие пропускаем). block=None — первый; N (1-й) — N-й подходящий.
    Возвращает (text, [предупреждения])."""
    blocks = re.findall(r"```(?:mermaid|flowchart)\s*\n(.*?)```", text, re.S | re.I)
    if not blocks:
        return text, []
    flow = [b for b in blocks
            if re.search(r"^\s*(flowchart|graph)\s+(LR|RL|TB|TD|BT)", b, re.M | re.I)]
    if block is not None:
        if not (1 <= block <= len(flow)):
            raise SystemExit(f"--block {block}: в файле flowchart-блоков "
                             f"всего {len(flow)} (mermaid-блоков {len(blocks)})")
        return flow[block - 1], []
    if flow:
        return flow[0], []
    return blocks[0], ["в файле нет mermaid flowchart/graph (только seq/др.) — "
                       f"взят первый блок из {len(blocks)}, проверьте выбор"]


def main(argv=None):
    p = argparse.ArgumentParser(
        prog="mmd2diagram.py", description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("file", help="mermaid-файл (или - для stdin)")
    p.add_argument("-o", "--out", help="куда писать (json/py; stdout по умолчанию)")
    p.add_argument("--emit", choices=("json", "py"), default="json")
    p.add_argument("--preview", metavar="PPTX", help="собрать однослайдовый pptx")
    p.add_argument("--title", help="заголовок слайда для --preview")
    p.add_argument("--legend", help='легенда: "solid=запросы,dash=OIDC"')
    p.add_argument("--dir", dest="direction", choices=("LR", "RL", "TB", "TD", "BT"),
                   help="переопределить направление (глубокие TB-схемы часто "
                        "читаемее в LR)")
    p.add_argument("--block", type=int, metavar="N",
                   help="N-й flowchart-блок mermaid в .md (1-й по счёту), "
                        "когда в документе несколько схем")
    a = p.parse_args(argv)

    text = sys.stdin.read() if a.file == "-" else open(
        a.file, encoding="utf-8-sig").read()
    legend = parse_legend(a.legend)
    m, routes, leg = prepare(text, legend, a.direction, a.block)
    block = to_block(m, routes, leg)

    for w in m.warnings:
        print(f"WARN: {w}", file=sys.stderr)
    if a.preview:
        print(f"preview: {preview_pptx(block, a.preview, a.title)}")
        return 0
    if a.emit == "py":
        res = to_py(m, routes, leg)
    else:
        res = json.dumps(block, ensure_ascii=False, indent=2)
    if a.out:
        open(a.out, "w", encoding="utf-8").write(res + "\n")
        print(f"saved: {a.out}")
    else:
        print(res)
    return 0


if __name__ == "__main__":
    sys.exit(main())
