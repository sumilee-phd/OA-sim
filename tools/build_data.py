# -*- coding: utf-8 -*-
"""실기 시뮬레이터 과제 데이터 생성(2026-11-11).

읽기만 하는 원본
  20_원고/2권/ch10~ch16.md            지시문 표·작업 순서·따라 하기 단계
  90_자산/그림/ch16/build_mock.py      모의 1~4세트 SETS(정의), 슬라이드 원안 함수
  90_자산/그림/ch14/verify_ch14_ppt.py 14장 슬라이드 원안 함수
  90_자산/그림/ch1x/verify_*.json, ch16/mockN_results.json  실측 기대값

쓰는 파일
  data/scenarios.json (사람이 읽는 정본), data/scenarios.js (window.OA_DATA, file:// 실행용)
  assets/slides/*.png (원안·견본 그림 사본)

검사: 장비 점검 세트의 자료를 원고 표(10-1·12-1·12-2)와, 모의 세트의 자료를 16장 표와 대조한다. 어긋나면 쓰지 않고 종료 코드 1.
실행: python tools/build_data.py   (--check 는 기존 파일과 비교만, 같으면 0·다르면 2)
"""
import ast
import hashlib
import json
import os
import re
import shutil
import sys
import types

sys.stdout.reconfigure(encoding="utf-8")
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)                                   # 실기시뮬레이터
ASSET = os.path.dirname(ROOT)                                  # 90_자산
BOOK = os.path.dirname(ASSET)                                  # 13_사무자동화산업기사
MS = os.path.join(BOOK, "20_원고", "2권")
FIG = os.path.join(ASSET, "그림")
ERR = []


def err(msg):
    ERR.append(msg)
    print("오류:", msg)


def load_json(rel):
    with open(os.path.join(FIG, rel), encoding="utf-8") as f:
        return json.load(f)


# =====================================================================================
# 원고 md 읽기
# =====================================================================================
def read_md(name):
    with open(os.path.join(MS, name), encoding="utf-8") as f:
        return f.read().splitlines()


def strip_md(s):
    return s.replace("**", "").strip()


def parse_row(line):
    cells = line.strip().strip("|").split("|")
    return [strip_md(c) for c in cells]


def md_table_at(lines, i):
    """lines[i] 가 표 머리 행이면 (head, rows, 다음 줄 번호)."""
    head = parse_row(lines[i])
    j = i + 2
    rows = []
    while j < len(lines) and lines[j].startswith("|"):
        rows.append(parse_row(lines[j]))
        j += 1
    return head, rows, j


def find_table(lines, caption_prefix):
    for i, ln in enumerate(lines):
        if ln.startswith("> " + caption_prefix):
            k = i + 1
            while k < len(lines) and not lines[k].startswith("|"):
                k += 1
            head, rows, _ = md_table_at(lines, k)
            return ln[2:].strip(), head, rows
    err(f"표 없음: {caption_prefix}")
    return None, [], []


def blocks_between(lines, start_pat, stop_pats):
    """start_pat 줄 다음부터 stop_pats 중 하나가 나올 때까지 블록(p·table·box)을 만든다."""
    out = []
    i = 0
    while i < len(lines) and not re.match(start_pat, lines[i]):
        i += 1
    if i >= len(lines):
        err(f"구간 시작 없음: {start_pat}")
        return out
    i += 1
    caption = None
    while i < len(lines):
        ln = lines[i]
        if any(re.match(p, ln) for p in stop_pats):
            break
        if ln.startswith("> "):
            caption = ln[2:].strip() if ln.startswith("> 표") else None   # 그림 캡션은 버린다
            i += 1
            continue
        if ln.startswith("|"):
            head, rows, j = md_table_at(lines, i)
            if caption is None and len(head) == 1:                 # 상자(작업 순서·주의 등)
                if head[0] != "작업 순서":                         # 작업 순서는 order 로 따로 둔다
                    out.append({"t": "box", "title": head[0], "lines": [r[0] for r in rows]})
            else:
                out.append({"t": "table", "caption": caption or "", "head": head, "rows": rows})
            caption = None
            i = j
            continue
        if ln.strip() and not ln.startswith("@@") and not ln.startswith("[[") and not ln.startswith("#"):
            out.append({"t": "p", "text": strip_md(ln)})
        if ln.startswith("#### ") or ln.startswith("### "):
            out.append({"t": "h", "text": strip_md(ln.lstrip("#"))})
        i += 1
    return out


def follow_steps(lines):
    """### 따라 하기 구간의 번호 단계."""
    out = []
    sec = ""
    cur = None
    for ln in lines:
        m = re.match(r"^## (\d+\.\d+) (.+)", ln)
        if m:
            sec = m.group(1) + " " + m.group(2)
            cur = None
            continue
        m = re.match(r"^### 따라 하기: (.+)", ln)
        if m:
            cur = {"sec": sec, "title": m.group(1).strip(), "items": []}
            out.append(cur)
            continue
        if ln.startswith("## ") or (ln.startswith("### ") and not ln.startswith("### 따라 하기")):
            cur = None
            continue
        if cur is not None:
            m = re.match(r"^(\d+)\. (.+)", ln)
            if m:
                cur["items"].append({"n": int(m.group(1)), "text": strip_md(m.group(2))})
    return out


def order_box(lines):
    for i, ln in enumerate(lines):
        if ln.strip() == "| 작업 순서 |":
            _, rows, _ = md_table_at(lines, i)
            return [r[0] for r in rows]
    err("작업 순서 상자 없음")
    return []


# =====================================================================================
# build_mock.py 의 SETS: dict(...) 호출과 리터럴만 있는 식을 AST 로 평가
# =====================================================================================
def eval_node(n):
    if isinstance(n, ast.Constant):
        return n.value
    if isinstance(n, ast.Dict):
        return {eval_node(k): eval_node(v) for k, v in zip(n.keys, n.values)}
    if isinstance(n, (ast.Tuple, ast.List)):
        return [eval_node(e) for e in n.elts]
    if isinstance(n, ast.UnaryOp) and isinstance(n.op, ast.USub):
        return -eval_node(n.operand)
    if isinstance(n, ast.Call) and isinstance(n.func, ast.Name) and n.func.id == "dict":
        return {kw.arg: eval_node(kw.value) for kw in n.keywords}
    raise ValueError("평가할 수 없는 식: " + ast.dump(n)[:80])


def load_sets():
    src = open(os.path.join(FIG, "ch16", "build_mock.py"), encoding="utf-8").read()
    tree = ast.parse(src)
    for node in tree.body:
        if isinstance(node, ast.Assign) and any(isinstance(t, ast.Name) and t.id == "SETS" for t in node.targets):
            return eval_node(node.value)
    raise SystemExit("SETS 없음")


# =====================================================================================
# 슬라이드 원안: COM 호출을 기록하는 가짜 Shapes 로 원안 함수를 실행
# =====================================================================================
PT = 28.3465
KIND = {1: "rect", 5: "roundRect", 9: "ellipse", 4: "diamond", 10: "hexagon", 69: "flowTerminator", 64: "flowData",
        67: "flowDocument", 71: "flowManualInput", 83: "flowStoredData", 86: "flowMagneticDisk", 102: "scrollH",
        179: "cloud", 33: "rightArrow", 34: "leftArrow", 35: "upArrow", 36: "downArrow", 31: "leftBrace", 32: "rightBrace",
        105: "calloutRect", 106: "calloutRoundRect", 107: "calloutOval", 61: "flowProcess", 63: "flowDecision"}


class Bag:
    def __init__(self, shape, path):
        object.__setattr__(self, "_shape", shape)
        object.__setattr__(self, "_path", path)

    def __getattr__(self, name):
        if name.startswith("__"):
            raise AttributeError(name)
        b = Bag(self._shape, self._path + "." + name)
        object.__setattr__(self, name, b)
        return b

    def __setattr__(self, name, value):
        self._shape.props[self._path + "." + name] = getattr(value, "value", value)

    def __setitem__(self, k, value):
        self._shape.props[f"{self._path}[{k}]"] = getattr(value, "value", value)

    def __call__(self, *args):
        self._shape.calls.append((self._path, args))
        return None


class FakeShape:
    def __init__(self, coll, kind, geom):
        object.__setattr__(self, "props", {})
        object.__setattr__(self, "calls", [])
        object.__setattr__(self, "coll", coll)
        object.__setattr__(self, "kind", kind)
        object.__setattr__(self, "geom", geom)
        object.__setattr__(self, "children", [])
        self.props["Name"] = f"도형{len(coll.items) + 1}"

    def __getattr__(self, name):
        if name.startswith("__"):
            raise AttributeError(name)
        if name == "Name":
            return self.props["Name"]
        if name == "ConnectorFormat":
            return Conn(self)
        b = Bag(self, name)
        object.__setattr__(self, name, b)
        return b

    def __setattr__(self, name, value):
        self.props[name] = getattr(value, "value", value)

    def ZOrder(self, cmd):
        items = self.coll.items
        items.remove(self)
        if cmd == 1:                       # msoSendToBack
            items.insert(0, self)
        else:
            items.append(self)


class Conn:
    def __init__(self, shp):
        self.shp = shp

    def BeginConnect(self, other, site):
        self.shp.props["begin"] = (other, site)

    def EndConnect(self, other, site):
        self.shp.props["end"] = (other, site)


class FakeRange:
    def __init__(self, coll, names):
        self.coll, self.names = coll, names

    def Group(self):
        members = [s for s in self.coll.items if s.props["Name"] in self.names]
        idx = min(self.coll.items.index(s) for s in members)
        for s in members:
            self.coll.items.remove(s)
        g = FakeShape(self.coll, "group", None)
        object.__setattr__(g, "children", members)
        self.coll.items.insert(idx, g)
        return g


class FakeShapes:
    def __init__(self):
        self.items = []

    def _add(self, kind, geom):
        s = FakeShape(self, kind, geom)
        self.items.append(s)
        return s

    def AddShape(self, kind, l, t, w, h):
        return self._add(KIND.get(kind, f"auto{kind}"), (l, t, w, h))

    def AddTextbox(self, orient, l, t, w, h):
        return self._add("textbox", (l, t, w, h))

    def AddConnector(self, kind, x1, y1, x2, y2):
        return self._add("connStraight" if kind == 1 else "connElbow", (x1, y1, x2, y2))

    def AddLine(self, x1, y1, x2, y2):
        return self._add("line", (x1, y1, x2, y2))

    def Range(self, names):
        return FakeRange(self, names)

    @property
    def Count(self):
        return len(self.items)


class FakeSlide:
    def __init__(self):
        self.Shapes = FakeShapes()


def cm(v):
    return round(v / PT, 3)


def rgb_hex(v):
    if v is None:
        return None
    v = int(v)
    r, g, b = v & 255, (v >> 8) & 255, (v >> 16) & 255
    return "#%02X%02X%02X" % (r, g, b)


def shape_spec(s):
    p = s.props
    d = {"name": p.get("Name"), "kind": s.kind}
    if s.kind == "group":
        d["children"] = [shape_spec(c) for c in s.children]
        return d
    if s.kind in ("line", "connStraight", "connElbow"):
        x1, y1, x2, y2 = s.geom
        if s.kind == "line":
            d.update(x1=cm(x1), y1=cm(y1), x2=cm(x2), y2=cm(y2))
        else:                                                   # AddConnector 는 (시작 x, 시작 y, 끝 x, 끝 y)
            d.update(x1=cm(x1), y1=cm(y1), x2=cm(x2), y2=cm(y2))
        if "begin" in p:
            d["begin"] = {"shape": p["begin"][0].props["Name"], "site": p["begin"][1]}
        if "end" in p:
            d["end"] = {"shape": p["end"][0].props["Name"], "site": p["end"][1]}
    else:
        l, t, w, h = s.geom
        d.update(x=cm(l), y=cm(t), w=cm(w), h=cm(h))
    fill_vis = p.get("Fill.Visible", -1 if s.kind not in ("textbox",) else 0)
    d["fill"] = rgb_hex(p.get("Fill.ForeColor.RGB", 0xFFFFFF)) if fill_vis else None
    line_vis = p.get("Line.Visible", -1)
    d["line"] = None if not line_vis else {
        "color": rgb_hex(p.get("Line.ForeColor.RGB", 0)), "weight": p.get("Line.Weight", 0.75),
        "dash": {None: "solid", 1: "solid", 4: "dash"}.get(p.get("Line.DashStyle"), "solid"),
        "begin": "arrow" if p.get("Line.BeginArrowheadStyle") == 2 else None,
        "end": "arrow" if p.get("Line.EndArrowheadStyle") == 2 else None}
    if p.get("Shadow.Visible"):
        d["shadow"] = {"color": rgb_hex(p.get("Shadow.ForeColor.RGB", 0x808080)), "dx": p.get("Shadow.OffsetX", 4),
                       "dy": p.get("Shadow.OffsetY", 4), "blur": p.get("Shadow.Blur", 0), "transparency": p.get("Shadow.Transparency", 0)}
    if "TextFrame.TextRange.Text" in p:
        d["text"] = p["TextFrame.TextRange.Text"]
        d["size"] = p.get("TextFrame.TextRange.Font.Size", 18)
        d["bold"] = bool(p.get("TextFrame.TextRange.Font.Bold", 0))
        d["align"] = {1: "left", 2: "center", 3: "right"}.get(p.get("TextFrame.TextRange.ParagraphFormat.Alignment", 1), "left")
        d["anchor"] = {1: "top", 3: "middle", 4: "bottom"}.get(p.get("TextFrame.VerticalAnchor", 3), "middle")
        d["font"] = p.get("TextFrame.TextRange.Font.Name", "맑은 고딕")
    adj = {int(k[len("Adjustments["):-1]): v for k, v in p.items() if k.startswith("Adjustments[")}
    if adj:
        d["adj"] = [adj.get(1), adj.get(2)]
    return d


def ref_slides():
    sys.path[:0] = [os.path.join(FIG, sub) for sub in ("ch10", "ch11", "ch12", "ch13", "ch14", "ch16")]
    import verify_ch14_ppt as V14   # noqa: E402  (pywin32 모듈을 읽기만 하고 COM 은 실행하지 않는다)
    import build_mock as BM         # noqa: E402
    out = {}

    def run(fn, *a):
        sl = FakeSlide()
        fn(sl, *a)
        return [shape_spec(s) for s in sl.Shapes.items]
    out["equip"] = [run(V14.slide1), run(V14.slide2)]
    for n in (1, 2, 3, 4):
        f1, f2 = BM.SLIDES[n]
        pt = BM.SETS[n]["pt"]
        out[f"mock{n}"] = [run(f1, pt["title1"]), run(f2, pt["title2"])]
    return out


# =====================================================================================
# 세트 정의
# =====================================================================================
def equip_sp():
    """10·11장 장비 점검 작업표(원고 표 10-1~10-4, 11-2)."""
    return dict(
        title="장비 점검 비용 산출",
        head=["장비명", "담당자", "구분코드", "사용연수", "점검시간", "할증률", "설치위치", "점검비용", "할증금액", "청구금액"],
        data=[["스위치", "김민수", 1, 3, 2, 0.00], ["서버", "박지영", 2, 5, 4, 0.10], ["라우터", "이준호", 1, 2, 3, 0.05],
              ["서버", "김서연", 1, 7, 6, 0.10], ["스위치", "최현우", 2, 4, 1, 0.00], ["라우터", "정다은", 2, 6, 2, 0.05],
              ["서버", "김태호", 2, 3, 3, 0.00], ["스위치", "한지민", 1, 8, 5, 0.10], ["라우터", "윤성민", 1, 1, 1, 0.00],
              ["스위치", "오수빈", 2, 5, 3, 0.05], ["서버", "김하늘", 1, 2, 2, 0.10], ["라우터", "장우진", 2, 9, 4, 0.10]],
        code_names={"1": "본관", "2": "별관"},
        f_h='=IF(D{r}=1,"본관","별관")',
        f_i='=F{r}*IF(B{r}="라우터",50000,IF(B{r}="스위치",30000,80000))',
        f_j="=I{r}*G{r}", f_k="=I{r}+J{r}",
        t6="점검시간이 3시간 미만인 장비들의 점검비용 합계", f6='=SUMIF(F5:F16,"<3",I5:I16)',
        t8="담당자 성명에 '김'이 포함되어 있는 사람의 수", f8='=COUNTIF(C5:C16,"*김*")&"명"',
        t10="할증금액이 10000 이상 20000 미만인 장비의 청구금액 합", f10='=SUMIFS(K5:K16,J5:J16,">=10000",J5:J16,"<20000")',
        sort_key_desc=False,
        chart_group="별관", chart_title="별관 장비 점검 비용", x_title="담당자", y_title="금액",
        pivot=dict(kind="group_items", row1="사용연수", row2="담당자", vals=["할증금액", "청구금액"], caption="연식 구분",
                   groups={"신규": [1, 2, 3], "중기": [4, 5, 6], "노후": [7, 8, 9]}, adj1_collapse=["중기"], adj2_expand_item=3),
    )


def equip_db():
    """12·13장 구분·점검 테이블(원고 표 12-1~12-7, 13-1~13-6)."""
    return dict(
        gubun=[["R", "라우터"], ["S", "스위치"], ["W", "무선AP"], ["F", "방화벽"]],
        gname="구분명", tname="점검", key="점검번호", who="장비명", days="점검시간", fee="점검비용",
        fields=[["점검번호", "숫자"], ["장비명", "짧은 텍스트"], ["구분코드", "짧은 텍스트"], ["점검월", "숫자"], ["점검시간", "숫자"], ["점검비용", "통화"]],
        rows=[[1001, "코어스위치1", "S", 3, 40, 120000], [1002, "코어스위치2", "S", 3, 35, 120000], [1003, "층스위치A", "S", 4, 20, 60000],
              [1004, "층스위치B", "S", 5, 25, 60000], [1005, "본관라우터", "R", 3, 50, 200000], [1006, "별관라우터", "R", 6, 45, 180000],
              [1007, "무선AP-1", "W", 4, 15, 40000], [1008, "무선AP-2", "W", 5, 20, 40000], [1009, "무선AP-3", "W", 6, 10, 30000],
              [1010, "백본라우터", "R", 5, 60, 250000], [1011, "무선AP-4", "W", 6, 12, None]],
        q_codes=["R", "W"], q_fee_lt=200000, qname="장비점검조회",
        form_title="구분코드가 R 또는 W이면서 점검비용이 200,000원 미만인 장비 점검 현황",
        base={"R": 5000, "S": 3000, "W": 2000}, base_label="시간단가", calc="인건비", disc="할인액", total="청구액",
        disc_rule=[[40, 20000], [20, 10000], [0, 5000]], disc_sign=-1, nz=True,
        report_title="장비 점검비 산출 현황", rname_q="보고서용",
    )


def mock_sp(sp):
    d = dict(sp)
    d["data"] = [list(r) for r in sp["data"]]
    d["code_names"] = {str(k): v for k, v in sp["code_names"].items()}
    del d["rate"]
    pv = dict(sp["pivot"])
    if pv["kind"] == "group_rate":
        d["pivot"] = dict(kind="group_items", row1="할인율", row2=pv["row2"], vals=list(pv["vals"]), caption=pv["caption"],
                          groups={k: list(v) for k, v in pv["groups"].items()}, adj1_collapse=list(pv["adj1_collapse"]),
                          adj2_expand_item=pv["adj2_expand_rate"])
    else:
        d["pivot"] = dict(kind="plain", row1=pv["row1"], row2=pv["row2"], vals=list(pv["vals"]),
                          adj1_collapse=list(pv["adj1_collapse"]), adj2_expand_item=pv["adj2_expand_only"])
    return d


def mock_db(db):
    fields = [[db["key"], "숫자"], [db["who"], "짧은 텍스트"], ["구분코드", "짧은 텍스트"], [db["days"], "숫자"], [db["fee"], "통화"]]
    return dict(
        gubun=[list(g) for g in db["gubun"]], gname="종류", tname=db["tname"], key=db["key"], who=db["who"], days=db["days"], fee=db["fee"],
        fields=fields, rows=[list(r) for r in db["rows"]], q_codes=list(db["q_codes"]), q_fee_lt=db["q_fee_lt"], qname="조회",
        form_title=db["form_title"], base=db["base"], base_label=db["base_label"], calc=db["calc"], disc=db["disc"], total=db["total"],
        disc_rule=[list(x) for x in db["disc_rule"]], disc_sign=db.get("disc_sign", -1), nz=False,
        report_title=db["report_title"], rname_q="보고서용",
    )


# =====================================================================================
# 원고 표와 자료 대조
# =====================================================================================
def num(s):
    s = s.replace(",", "").strip()
    if s.endswith("%"):
        return round(float(s[:-1]) / 100, 6)
    if s in ("", "(비어 있음)"):
        return None
    try:
        return int(s)
    except ValueError:
        try:
            return float(s)
        except ValueError:
            return s


def check_sp(lines, cap, sp, label):
    _, head, rows = find_table(lines, cap)
    if not rows:
        return
    if [c for c in rows[0][1:]] != sp["head"][:6]:
        err(f"{label} 머리글 불일치 {rows[0][1:]} / {sp['head'][:6]}")
    got = [[num(c) for c in r[1:]] for r in rows[1:]]
    want = [[v if not isinstance(v, float) else round(v, 6) for v in r] for r in sp["data"]]
    if got != want:
        err(f"{label} 자료 불일치\n  원고 {got}\n  정의 {want}")


def check_db(lines, cap_g, cap_t, db, label):
    _, hg, rg = find_table(lines, cap_g)
    if [[r[0], r[1]] for r in rg] != db["gubun"]:
        err(f"{label} 구분 테이블 불일치 {rg}")
    _, ht, rt = find_table(lines, cap_t)
    if ht != [f[0] for f in db["fields"]]:
        err(f"{label} 필드 불일치 {ht}")
    got = [[num(c) for c in r] for r in rt]
    if got != db["rows"]:
        err(f"{label} 행 불일치\n  원고 {got}\n  정의 {db['rows']}")


# =====================================================================================
# 기대값
# =====================================================================================
def exp_equip():
    v10 = load_json("ch10/verify_results.json")["실측"]
    v11 = load_json("ch11/verify_results.json")["실측"]
    v12 = load_json("ch12/verify_results.json")["실측"]
    v13 = load_json("ch13/verify_results.json")["실측"]
    v14 = load_json("ch14/verify_results.json")["실측"]
    sp = {"평균": [v10["I17"], v10["J17"], v10["K17"]], "⑥": v10["K18"], "⑧": v10["K20"], "⑩": v10["K22"],
          "정렬_사람": v10["정렬_담당자"], "차트_합": v10["별관_청구"], "차트_기본": v10["별관_점검비용"],
          "피벗_그룹": v11["그룹"], "피벗_총합": v11["총합계"], "피벗_연수": v11["연수"], "피벗_확장항목_사람": v11["3년_담당자"],
          "피벗_행": [4, 14, 18, 29], "_출처": "ch10·ch11 verify_results.json 실측, 피벗 행은 11.4절 검산(조정1 11행·조정2 12행)"}
    db = {"조회": v12, "보고서_행": v13["행"], "보고서_그룹": v13["그룹"], "보고서_평균": v13["평균"],
          "_출처": "ch12·ch13 verify_results.json 실측"}
    pt = {"도형 수": v14["도형 수"], "슬라이드 크기 cm": v14["슬라이드 크기 cm"], "회색 채우기": v14["회색 채우기"],
          "유인물": v14["유인물"], "_출처": "ch14 verify_results.json 실측"}
    return {"sp": sp, "db": db, "pt": pt}


def exp_mock(n):
    d = load_json(f"ch16/mock{n}_results.json")
    e_sp, e_db = d["기대"]["sp"], d["기대"]["db"]
    x = d["excel"]["실측"]
    a = d["access"]["실측"]
    bad = [k for sec in ("excel", "access") for k, v in d[sec]["일치"].items() if not v]
    if bad:
        err(f"mock{n} 실측 JSON 에 불일치 항목 {bad}")
    sp = {"평균": x["평균"], "⑥": x["⑥"], "⑧": x["⑧"], "⑩": x["⑩"], "정렬_사람": x["정렬_사람"], "차트_합": x["차트_합"],
          "차트_기본": x["차트_기본"], "피벗_그룹": x["피벗1_그룹"], "피벗_총합": x["피벗1_총합"], "피벗_행": x["피벗_행"],
          "피벗_그룹항목_원래이름": x.get("피벗_그룹항목_원래이름"), "피벗_그룹항목": x.get("피벗_그룹항목"),
          "_출처": f"ch16/mock{n}_results.json excel 실측"}
    db = {"조회": a["조회"], "보고서_행": a["보고서_행"], "보고서_그룹": a["보고서_그룹"], "보고서_평균": a["보고서_평균"],
          "SQL": d["access_sql"], "_출처": f"ch16/mock{n}_results.json access 실측"}
    pt = {"도형 수": d["ppt"]["도형 수"], "도형 종류": d["ppt"]["도형 종류"], "유인물": d["ppt"]["유인물"],
          "_출처": f"ch16/mock{n}_results.json ppt 실측"}
    for k in ("평균", "⑥", "⑧", "⑩", "정렬_사람"):
        a1, b1 = sp[k], e_sp[k]
        same = (all(abs(p - q) < 0.01 for p, q in zip(a1, b1)) if k == "평균" else (a1 == b1 or (isinstance(a1, float) and abs(a1 - b1) < 0.01)))
        if not same:
            err(f"mock{n} 기대·실측 차이 {k}: {a1} / {b1}")
    return {"sp": sp, "db": db, "pt": pt}


def exp_ch15():
    v = load_json("ch15/verify_ch15_p2.json")
    a = load_json("ch15/verify_results.json")
    return {"p2": v["실측"], "p2_일치": v["일치"], "abc": a, "_출처": "ch15/verify_ch15_p2.json·verify_results.json"}


# =====================================================================================
# 지시문
# =====================================================================================
def chapter_task(md, ch, app, sheet, start_from, title):
    lines = read_md(md)
    task = blocks_between(lines, r"^### 과제", [r"^## \d+\.1 "])
    return {"ch": ch, "app": app, "sheet": sheet, "startFrom": start_from, "title": title,
            "task": task, "order": order_box(lines), "steps": follow_steps(lines)}


def mock_tasks(lines, n):
    head = {1: "16.1", 2: "16.2", 3: "16.3", 4: "16.4"}[n]
    nxt = {1: "16.2", 2: "16.3", 3: "16.4", 4: "16.5"}[n]
    i0 = next(i for i, ln in enumerate(lines) if ln.startswith(f"## {head} "))
    i1 = next(i for i, ln in enumerate(lines) if ln.startswith(f"## {nxt} "))
    seg = lines[i0:i1]
    intro = blocks_between(seg, rf"^## {head} ", [r"^### 가\."])
    sp = blocks_between(seg, r"^### 가\. 표계산", [r"^### 나\."])
    db = blocks_between(seg, r"^### 나\. 자료처리", [r"^### 다\."])
    pt = blocks_between(seg, r"^### 다\. 시상", [r"^## "])
    return intro, sp, db, pt


def ch15_task():
    lines = read_md("ch15.md")
    i0 = next(i for i, ln in enumerate(lines) if ln.startswith("## 15.3 "))
    seg = lines[i0:]
    blocks = []
    for cap in ("표 15-6.", "표 15-7.", "표 15-10.", "표 15-12.", "표 15-14."):
        c, h, r = find_table(seg, cap)
        blocks.append({"t": "table", "caption": c, "head": h, "rows": r})
    formulas = []
    for cap in ("표 15-8.", "표 15-9.", "표 15-11.", "표 15-13.", "표 15-15."):
        c, h, r = find_table(seg, cap)
        formulas.append({"t": "table", "caption": c, "head": h, "rows": r})
    return blocks, formulas


def copy_assets():
    dst = os.path.join(ROOT, "assets", "slides")
    os.makedirs(dst, exist_ok=True)
    pairs = [("ch14/fig14_01_slide1.png", "equip_1.png"), ("ch14/fig14_02_slide2.png", "equip_2.png"),
             ("ch10/fig10_01_print.png", "equip_s1.png"), ("ch11/fig11_01_print.png", "equip_s2.png"),
             ("ch12/fig12_01_print.png", "equip_s3.png"), ("ch13/fig13_01_print.png", "equip_s4.png"),
             ("ch14/fig14_03_print.png", "equip_s5.png")]
    for n in (1, 2, 3, 4):
        pairs += [(f"ch16/fig16_s{n}_slide1.png", f"mock{n}_1.png"), (f"ch16/fig16_s{n}_slide2.png", f"mock{n}_2.png"),
                  (f"ch16/fig16_s{n}_sheet.png", f"mock{n}_sheet.png")]
    for src, name in pairs:
        shutil.copyfile(os.path.join(FIG, src), os.path.join(dst, name))
    return len(pairs)


def main():
    check = "--check" in sys.argv
    SETS = load_sets()
    ch16 = read_md("ch16.md")
    slides = ref_slides()

    # ---------- 장비 점검 세트
    sp, db = equip_sp(), equip_db()
    check_sp(read_md("ch10.md"), "표 10-1.", sp, "장비 점검 작업표")
    check_db(read_md("ch12.md"), "표 12-1.", "표 12-2.", db, "장비 점검 DB")
    equip = {
        "id": "equip", "title": "장비 점검 세트", "sub": "10~14장 과제, 출력물 5매",
        "chapters": [
            chapter_task("ch10.md", 10, "excel", "5-1", None, "10장 작업표와 혼합 그래프"),
            chapter_task("ch11.md", 11, "excel", "5-2", 10, "11장 피벗테이블과 집계 검증"),
            chapter_task("ch12.md", 12, "access", "5-3", None, "12장 데이터베이스 조회화면"),
            chapter_task("ch13.md", 13, "access", "5-4", 12, "13장 데이터베이스 보고서"),
            chapter_task("ch14.md", 14, "ppt", "5-5", None, "14장 프레젠테이션과 유인물"),
        ],
        "sp": sp, "db": db,
        "pt": {"titles": ["장비 점검 자료의 처리 흐름", "본관·별관 네트워크 장비 구성"], "images": ["assets/slides/equip_1.png", "assets/slides/equip_2.png"],
               "ref": slides["equip"]},
        "samples": [f"assets/slides/equip_s{i}.png" for i in range(1, 6)],
        "expected": exp_equip(),
    }
    sets = [equip]

    # ---------- 15장
    t15, f15 = ch15_task()
    sets.append({"id": "ch15", "title": "15장 유형 과제", "sub": "과제 A~F 수식 연습, 출력물 없음",
                 "chapters": [{"ch": 15, "app": "excel", "sheet": None, "startFrom": None, "title": "15.3 유형별 대표 과제",
                               "task": t15, "formulas": f15, "order": [], "steps": []}],
                 "ch15": True, "expected": exp_ch15()})

    # ---------- 모의 1~4
    caps = {1: ("표 16-2.", "표 16-5.", "표 16-6."), 2: ("표 16-9.", "표 16-12.", "표 16-13."),
            3: ("표 16-15.", "표 16-18.", "표 16-19."), 4: ("표 16-21.", "표 16-24.", "표 16-25.")}
    for n in (1, 2, 3, 4):
        S = SETS[n]
        msp, mdb = mock_sp(S["sp"]), mock_db(S["db"])
        c_sp, c_g, c_t = caps[n]
        check_sp(ch16, c_sp, msp, f"모의 {n} 작업표")
        check_db(ch16, c_g, c_t, mdb, f"모의 {n} DB")
        intro, tsp, tdb, tpt = mock_tasks(ch16, n)
        sets.append({
            "id": f"mock{n}", "title": f"모의 {n}세트", "sub": S["name"] + ", 2시간, 출력물 5매",
            "intro": intro,
            "chapters": [
                {"ch": 16, "app": "excel", "sheet": "5-1·5-2", "startFrom": None, "title": "가. 표계산 작업", "task": tsp, "order": [], "steps": []},
                {"ch": 16, "app": "access", "sheet": "5-3·5-4", "startFrom": None, "title": "나. 자료처리 작업", "task": tdb, "order": [], "steps": []},
                {"ch": 16, "app": "ppt", "sheet": "5-5", "startFrom": None, "title": "다. 시상 작업", "task": tpt, "order": [], "steps": []},
            ],
            "sp": msp, "db": mdb,
            "pt": {"titles": [S["pt"]["title1"], S["pt"]["title2"]], "images": [f"assets/slides/mock{n}_1.png", f"assets/slides/mock{n}_2.png"],
                   "ref": slides[f"mock{n}"]},
            "samples": [f"assets/slides/mock{n}_sheet.png"],
            "expected": exp_mock(n),
        })

    # 슬라이드 원안 도형 수 = 실측 도형 수
    for s in sets:
        if "pt" in s:
            got = [len(x) for x in s["pt"]["ref"]]
            want = s["expected"]["pt"]["도형 수"]
            if got != want:
                err(f"{s['id']} 원안 도형 수 {got} / 실측 {want}")

    if ERR:
        print(f"오류 {len(ERR)}건, 파일을 쓰지 않습니다.")
        sys.exit(1)

    # 부록 표 A-1 출력물 5매 점검표
    a1 = read_md("a1_점검표.md")
    _, ah, arows = find_table(a1, "표 A-1.")
    checklist, cur = [], None
    for r in arows:
        if r[0]:
            cur = r[0]
        checklist.append({"sheet": cur, "item": r[1], "sec": r[2]})
    if len(checklist) != 30:
        err(f"점검표 항목 수 {len(checklist)}(30 이어야 함)")
    if ERR:
        sys.exit(1)

    body = {"title": "사무자동화산업기사 실기 시뮬레이터 과제", "sets": sets, "checklist": checklist}
    raw = json.dumps(body, ensure_ascii=False, sort_keys=True)
    body["version"] = hashlib.sha256(raw.encode("utf-8")).hexdigest()[:12]
    text = json.dumps(body, ensure_ascii=False, indent=1)
    js = "/* 생성 파일: tools/build_data.py. 직접 고치지 않습니다. */\nwindow.OA_DATA = " + json.dumps(body, ensure_ascii=False) + ";\n"
    pj, pjs = os.path.join(ROOT, "data", "scenarios.json"), os.path.join(ROOT, "data", "scenarios.js")
    if check:
        same = os.path.exists(pj) and open(pj, encoding="utf-8").read() == text
        print("같음" if same else "다름", body["version"])
        sys.exit(0 if same else 2)
    open(pj, "w", encoding="utf-8", newline="\n").write(text)
    open(pjs, "w", encoding="utf-8", newline="\n").write(js)
    n = copy_assets()
    print(f"세트 {len(sets)}, 그림 {n}장, version {body['version']}")
    for s in sets:
        print(" ", s["id"], s["title"], "장", [c["title"] for c in s["chapters"]],
              "도형", [len(x) for x in s["pt"]["ref"]] if "pt" in s else "-")


if __name__ == "__main__":
    main()
