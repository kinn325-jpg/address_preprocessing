# -*- coding: utf-8 -*-
"""
배송 송장 주소 자동 정제 프로그램
================================
사용법:
    python 주소정제.py [원본폴더] [결과폴더]

    - 원본폴더 기본값: 이 프로그램 폴더의 상위 폴더 (예: 배송자료 폴더)
    - 결과폴더 기본값: 원본폴더\\주소정제_결과

동작:
    1) 원본폴더의 모든 .xlsx 를 읽어 '도로명 사전'을 학습합니다
       (같은 우편번호에서 자주 쓰인 정상 도로명 표기를 기준으로 삼음).
    2) 각 파일의 Address1 / Address2 / Postcode 를 자동 수정한 뒤
       결과폴더에 '<원본파일명>_정제.xlsx' 로 저장합니다. (원본은 건드리지 않음)
         - 수정된 칸: 노란색
         - M열 '수정내역', N열 '확인필요'(빨간색) 추가
    3) 결과폴더에 '_주소정제_결과보고.xlsx' (요약 / 확인필요 목록 / 수정내역) 생성

필요: Python 3.8+, openpyxl  (pip install openpyxl)
"""
import os
import re
import sys
import json
import glob
import difflib
import collections
import datetime

try:
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment
except ImportError:
    print("openpyxl 이 필요합니다.  명령 프롬프트에서:  pip install openpyxl")
    sys.exit(1)

# ---------------------------------------------------------------------------
# 설정
# ---------------------------------------------------------------------------
COL_AWB, COL_A1, COL_A2, COL_POST = 3, 9, 10, 11     # C, I, J, K
COL_NOTE, COL_FLAG = 13, 14                          # M, N (새로 추가)
TRUNC_LEN = 34            # 이 길이 이상이면 '35자 잘림' 가능성이 있는 칸으로 봄
OUT_SUFFIX = "_정제"

FILL_CHANGED = PatternFill("solid", fgColor="FFF2CC")   # 노랑: 수정된 칸
FILL_FLAG = PatternFill("solid", fgColor="F8CBAD")      # 빨강: 확인 필요
FILL_HEAD = PatternFill("solid", fgColor="D9E1F2")
FILL_WHITE = PatternFill("solid", fgColor="FFFFFF")
FONT_BLACK = Font(color="000000")

# 행정구역 (이름, 종류, 우편번호 범위들)
DISTRICTS = [
    ("Songpa", "gu", [(5500, 5899)]),
    ("Gangnam", "gu", [(6000, 6399)]),
    ("Gangdong", "gu", [(5200, 5499)]),
    ("Seocho", "gu", [(6500, 6899)]),
    ("Gwangjin", "gu", [(4900, 5199)]),
    ("Bundang", "gu", [(13100, 13699)]),
    ("Sujeong", "gu", [(13100, 13699)]),
    ("Jungwon", "gu", [(13100, 13699)]),
    ("Hanam", "si", [(12900, 13099)]),
    ("Seongnam", "si", [(13100, 13699)]),
    ("Gwangju", "si", [(12700, 12899)]),
    ("Gyeonggi", "do", [(10000, 18999)]),
]
SEOUL_GU = {"Songpa", "Gangnam", "Gangdong", "Seocho", "Gwangjin"}
SEONGNAM_GU = {"Bundang", "Sujeong", "Jungwon"}
NOT_ROAD_NAMES = {"beon", "bun", "beun", "ben", "beongil", "bungil", "dae", "daer", "dea"}
# 도로명 앞 단어로 붙이면 안 되는 단어들
PRE_STOP = {"gu", "si", "do", "gil", "ro", "f", "fl", "floor", "dong", "ho", "bldg", "building",
            "tower", "apt", "the", "and", "of", "seoul", "korea", "b", "unit", "room", "rm", "no",
            "ga", "ri", "eup", "myeon", "center", "centre", "plaza", "city", "daero", "st", "nd",
            "rd", "th", "kr", "south", "republic", "suite", "office", "gil", "bd", "rep"}

# 자동으로 잡기 어려운 발음식 표기 (기본 별칭, 수동사전.json 으로 추가 가능)
DEFAULT_ALIASES = {
    "Seonneung-ro": "Seolleung-ro", "Sunneung-ro": "Seolleung-ro",
    "Seonreung-ro": "Seolleung-ro", "Sunreung-ro": "Seolleung-ro",
    "Heonneung-ro": "Heolleung-ro", "Heonreung-ro": "Heolleung-ro", "Hunreung-ro": "Heolleung-ro",
}


def strict_similar(a, b, raw_a=None, raw_b=None):
    """두 철자 차이가 '오타' 수준인지 (2글자 이상 삽입/삭제는 다른 도로로 봄: Balibong / Balibongbuk)"""
    if raw_a and raw_b:
        ra = re.sub(r"[^a-z]", "", raw_a.lower())
        rb = re.sub(r"[^a-z]", "", raw_b.lower())
        for tag, i1, i2, j1, j2 in difflib.SequenceMatcher(None, ra, rb).get_opcodes():
            if tag in ("insert", "delete") and max(i2 - i1, j2 - j1) >= 2:
                return False
    sm = difflib.SequenceMatcher(None, a, b)
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag in ("insert", "delete") and max(i2 - i1, j2 - j1) >= 2:
            return False
        if tag == "replace" and max(i2 - i1, j2 - j1) > 2:
            return False
    return sm.ratio()


# ---------------------------------------------------------------------------
# 로마자 표기 변형을 같은 값으로 접는 함수 (Samsung/Samseong, Youngdong/Yeongdong …)
# ---------------------------------------------------------------------------
def fold(s):
    s = re.sub(r"[^a-z]", "", s.lower())
    s = re.sub(r"y(?![aeiou])", "i", s)          # Olympic / Olimpic
    for a, b in (("yeo", "yu"), ("you", "yu"), ("eo", "u"), ("oo", "u"), ("ou", "u"),
                 ("wh", "w"), ("sh", "s"), ("ee", "i"), ("k", "g"),
                 ("p", "b"), ("t", "d"), ("l", "r"), ("ch", "j"), ("c", "g")):
        s = s.replace(a, b)
    return re.sub(r"(.)\1+", r"\1", s)


def pad_post(v):
    """우편번호 값을 5자리 문자열로. 유효하지 않으면 None"""
    if v is None:
        return None
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    d = re.sub(r"\D", "", str(v))
    if len(d) == 4:
        d = "0" + d
    return d if len(d) == 5 else None


def match_case(src, canon):
    letters = re.sub(r"[^A-Za-z]", "", src)
    if letters and letters.isupper() and len(letters) > 1:
        return canon.upper()
    return canon


ROAD_RE = re.compile(
    r"(?<![A-Za-z0-9])(?:(?P<pre>[A-Za-z]+) )?(?P<name>(?!dae-?ro(?![A-Za-z]))[A-Za-z]+)"
    r"(?P<sep>\s*-\s*|\s+|)(?P<suf>dae-ro|daero|ro|lo|gil)(?![A-Za-z])", re.I)

DIST_RE = re.compile(
    r"(?<![A-Za-z])(?P<name>[A-Za-z]{4,})(?P<sep>\s*-\s*|\s+|)(?P<suf>gu|ku|si|shi|city|do)(?![A-Za-z])",
    re.I)

TEXT_POST_RE = re.compile(r"(?<![\d-])(0[4-6]\d{3}|1[2-3]\d{3})(?![\d-])")


# ---------------------------------------------------------------------------
# 도로명 사전
# ---------------------------------------------------------------------------
class RoadDict:
    def __init__(self):
        self.spell = collections.defaultdict(collections.Counter)   # key -> 표기 -> 횟수
        self.total = collections.Counter()                           # key -> 전체 횟수
        self.post_roads = collections.defaultdict(collections.Counter)  # 우편번호5 -> key
        self.road_post3 = collections.defaultdict(collections.Counter)  # key -> 우편번호 앞3
        self.canon = {}
        self.manual = {"치환": {}, "도로명_추가": []}
        self._fuzzy_cache = {}

    # --- 학습 ---
    def observe(self, text, post5):
        if not text:
            return
        for m in ROAD_RE.finditer(text):
            name, sep, suf = m.group("name"), m.group("sep"), m.group("suf")
            if name.lower() in NOT_ROAD_NAMES or len(name) < 3:
                continue
            if suf.lower().startswith("dae") and name.lower().endswith("dae") and len(name) > 5:
                name = name[:-3]          # Songpadae-daero → Songpa-daero
            k = fold(name + suf)
            self.total[k] += 1
            if "-" in sep:
                self.spell[k][name.capitalize() + "-" + suf.lower()] += 1
            if post5:
                self.post_roads[post5][k] += 1
                self.road_post3[k][post5[:3]] += 1

    def build(self):
        self.canon = {}
        for k, c in self.spell.items():
            if self.total[k] >= 3 and c:
                self.canon[k] = c.most_common(1)[0][0]
        for r in self.manual.get("도로명_추가", []):
            k = fold(r)
            self.canon[k] = r
            self.total[k] = max(self.total[k], 50)
        self.known_keys = list(self.canon)
        self._fuzzy_cache = {}

        # (a) 오타 변형 → 대표 도로명 (예: Sonpa-daero → Songpa-daero)
        self.alias = {}
        for k in self.known_keys:
            best = self._best_target(k, self.total[k], raw=self.canon[k])
            if best:
                self.alias[k] = best
        aliases = dict(DEFAULT_ALIASES)
        aliases.update(self.manual.get("도로명_별칭", {}))
        for src, dst in aliases.items():
            dk = fold(dst)
            if dk not in self.canon:
                self.canon[dk] = dst
                self.total[dk] = max(self.total[dk], 50)
            if fold(src) != dk:
                self.alias[fold(src)] = dk
        self.known_keys = list(self.canon)

        # (b) 띄어쓰기로 생긴 조각 (예: 'Hwang Saeul-ro' 의 'Saeul-ro')
        self.fragment = set()
        by_len = sorted(self.known_keys, key=len, reverse=True)
        for k in self.known_keys:
            for K in by_len:
                if len(K) <= len(k) + 1:
                    break
                if K.endswith(k) and self.total[K] >= 5 * self.total[k]:
                    self.fragment.add(k)
                    break

        self.main_keys = {k for k in self.known_keys if k not in self.alias and k not in self.fragment}
        self.vocab_roads = [self.canon[k] for k in self.main_keys if self.total[k] >= 10]

    def _best_target(self, k, cnt, post5=None, need_evidence=False, raw=None):
        best = None
        for cand in difflib.get_close_matches(k, self.known_keys, n=5, cutoff=0.9):
            if cand == k:
                continue
            tc = self.total[cand]
            if tc < 20 or tc < 10 * max(cnt, 1):
                continue
            ratio = strict_similar(k, cand, raw, self.canon[cand])
            if not ratio or ratio < 0.9:
                continue
            if need_evidence:
                seen_here = post5 and self.post_roads.get(post5, {}).get(cand, 0) > 0
                if not (seen_here or ratio >= 0.92):
                    continue
            if best is None or (ratio, tc) > best[0]:
                best = ((ratio, tc), cand)
        return best[1] if best else None

    # --- 조회 ---
    def lookup(self, name, suf, post5, loose=True):
        """(표준표기, 방식) 또는 (None, None)"""
        if suf.lower().startswith("dae") and name.lower().endswith("dae") and len(name) > 5:
            name = name[:-3]
        k = fold(name + suf)
        if k in self.alias:
            t = self.alias[k]
            while t in self.alias and self.alias[t] != k:
                t = self.alias[t]
            return self.canon[t], "fuzzy"
        if k in self.canon:
            return self.canon[k], "exact"
        if not loose:
            return None, None
        ck = (k, post5)
        if ck not in self._fuzzy_cache:
            self._fuzzy_cache[ck] = self._best_target(k, self.total.get(k, 0), post5, True, raw=name + suf)
        cand = self._fuzzy_cache[ck]
        if cand:
            cand = self.alias.get(cand, cand)
            return self.canon[cand], "fuzzy"
        return None, None

    def roads_at(self, post5):
        return {self.canon[k] for k in self.post_roads.get(post5, {}) if k in self.main_keys}

    # --- 저장 / 불러오기 ---
    def to_json(self):
        return {
            "spell": {k: dict(v) for k, v in self.spell.items()},
            "total": dict(self.total),
            "post_roads": {k: dict(v) for k, v in self.post_roads.items()},
            "road_post3": {k: dict(v) for k, v in self.road_post3.items()},
        }

    def load_json(self, d):
        for k, v in d.get("spell", {}).items():
            self.spell[k].update(v)
        self.total.update(d.get("total", {}))
        for k, v in d.get("post_roads", {}).items():
            self.post_roads[k].update(v)
        for k, v in d.get("road_post3", {}).items():
            self.road_post3[k].update(v)


# ---------------------------------------------------------------------------
# 주소 한 칸 수정
# ---------------------------------------------------------------------------
def district_fix(text, notes):
    names = [d[0] for d in DISTRICTS]
    folded = {fold(n): n for n in names}
    kinds = {d[0]: d[1] for d in DISTRICTS}

    def repl(m):
        raw, name = m.group(0), m.group("name")
        suf = m.group("suf").lower()
        kind = "gu" if suf in ("gu", "ku") else "do" if suf == "do" else "si"
        fn = fold(name)
        hit = folded.get(fn)
        if not hit:
            same_kind = [k for k, n in folded.items() if kinds[n] == kind]
            cm = difflib.get_close_matches(fn, same_kind, n=1, cutoff=0.84)
            if cm and len(fn) >= 5 and (strict_similar(fn, cm[0], name, folded[cm[0]]) or 0) >= 0.84:
                hit = folded[cm[0]]
        if not hit:
            return raw
        canon = match_case(raw, hit + "-" + kinds[hit])
        if canon.lower() != raw.lower():
            notes.append(f"행정구역 표기 '{raw}'→'{canon}'")
            return canon
        return raw
    return DIST_RE.sub(repl, text)


def road_fix(text, post5, rd, notes):
    def repl(m):
        raw = m.group(0)
        pre, name, sep, suf = m.group("pre"), m.group("name"), m.group("sep"), m.group("suf")
        if name.lower() in NOT_ROAD_NAMES:
            return raw
        cands = []
        if pre:
            st = m.start("pre")
            before = m.string[st - 1] if st > 0 else " "
            if pre.lower() not in PRE_STOP and before not in "-." and len(pre) >= 2:
                cands.append((pre + name, True))
        cands.append((name, False))
        for nm, uses_pre in cands:
            loose = bool(sep.strip()) or len(nm) >= 6
            canon, how = rd.lookup(nm, suf, post5, loose=loose)
            if not canon:
                continue
            if uses_pre and how == "fuzzy" and not fold(canon).startswith(fold(pre)[:3]):
                continue
            src = raw if uses_pre else raw[len(pre) + 1:] if pre else raw
            new = match_case(src, canon)
            if new.lower() == src.lower():
                return raw
            tag = "도로명 오타" if how == "fuzzy" else "도로명 표기"
            notes.append(f"{tag} '{src}'→'{new}'")
            return new if uses_pre or not pre else pre + " " + new
        return raw
    return ROAD_RE.sub(repl, text)


def expected_next(text):
    """텍스트 끝에 올 것으로 기대되는 행정구역"""
    last = None
    for m in re.finditer(r"(Songpa|Gangnam|Gangdong|Seocho|Gwangjin|Bundang|Sujeong|Jungwon)-gu|"
                         r"(Hanam|Seongnam|Gwangju)-si|(Gyeonggi)-do|(Seoul)", text, re.I):
        last = m
    if not last:
        return []
    if last.group(1):
        g = last.group(1).capitalize()
        return ["Seoul"] if g in SEOUL_GU else ["Seongnam-si"]
    if last.group(2):
        return ["Gyeonggi-do"]
    return ["Korea", "Republic of Korea", "South Korea"]


def complete_tail(text, post5, rd, notes, flags_holder, has_road_elsewhere):
    """35자에서 잘린 마지막 단어 복원"""
    if not text or not re.search(r"[A-Za-z0-9-]$", text):
        return text
    # 1) 숫자 조각 → 우편번호
    m = re.search(r"(?:Seoul|Korea|-do|-si|-gu|Gyeonggi|code|\()[\s,:]*(\d{1,4})$", text, re.I)
    if m and post5 and post5.startswith(m.group(1)) and not re.search(r"\d{5}", text):
        frag = m.group(1)
        notes.append(f"잘린 우편번호 '{frag}'→'{post5}'")
        return text[: len(text) - len(frag)] + post5
    m = re.search(r"(\d)\s*(b|be|beo|beon|beon-|beon-g|beon-gi|beong|beongi|bun|bun-|bun-g|bun-gi)$", text, re.I)
    if m:
        frag = text[m.start(2):]
        new = "BEON-GIL" if frag.isupper() else "beon-gil"
        notes.append(f"잘린 단어 '{frag}'→'{new}'")
        return text[:m.start(2)] + new
    m = re.search(r"(\d)-(g|gi)$", text, re.I)
    if m:
        frag = text[m.start(2):]
        new = "GIL" if frag.isupper() else "gil"
        notes.append(f"잘린 단어 '-{frag}'→'-{new}'")
        return text[:m.start(2)] + new
    m = re.search(r"([A-Za-z][A-Za-z\-]*)$", text)
    if not m:
        return text
    tok = m.group(1)
    head = text[: len(text) - len(tok)]
    tu = tok.upper()
    # 2) 국가명
    if tu in ("KO", "KOR", "KORE") or (tu == "K" and re.search(r"(of|South|,)\s*$", head, re.I)):
        new = match_case(tok, "Korea") if len(tok) > 1 else ("KOREA" if head.isupper() else "Korea")
        notes.append(f"잘린 단어 '{tok}'→'{new}'")
        return head + new
    # 3) 기대 행정구역 (…Songpa-gu, SE → Seoul)
    for e in expected_next(head):
        if e.upper().startswith(tu) and e.upper() != tu and (len(tok) >= 2 or head.rstrip().endswith(",")):
            new = match_case(tok, e) if len(tok) > 1 else (e.upper() if head.isupper() else e)
            notes.append(f"잘린 단어 '{tok}'→'{new}'")
            return head + new
    if len(tok) < 3:
        return text
    # 4) 도로명 / 행정구역 사전에서 앞부분 일치
    dist_vocab = [d[0] + "-" + d[1] for d in DISTRICTS]
    if any(v.upper() == tu for v in rd.vocab_roads + dist_vocab):
        return text
    ft = fold(tok)

    def pref(v, loose):
        vu = v.upper()
        if vu.startswith(tu) or vu.replace("-", "").startswith(tu.replace("-", "")):
            return True
        return (loose and len(tok) >= 4 and fold(v).startswith(ft)
                and vu.replace("-", "")[:3] == tu.replace("-", "")[:3])
    road_c = [v for v in rd.vocab_roads if pref(v, True)]
    dist_c = [v for v in dist_vocab if pref(v, False)]
    if post5:
        pn = int(post5)
        rng = {d[0] + "-" + d[1]: d[2] for d in DISTRICTS}
        dist_c = [v for v in dist_c if any(lo <= pn <= hi for lo, hi in rng[v])]
    if has_road_elsewhere:
        # 이미 도로명이 있는 주소라면, 도로명 접미사가 잘린 경우(-dae, -r …)나
        # 도로 이름 자체가 잘린 경우만 도로명으로 복원
        def truncated_road(v):
            namepart = re.split(r"-(?=[a-z]+$)", v)[0]
            if fold(namepart) == ft:
                return False
            return bool(re.search(r"-[A-Za-z]{0,4}$", tok)) or len(tok) < len(namepart)
        road_c = [v for v in road_c if truncated_road(v)]
    choice = None
    # 구/시 복원은 쉼표·도로명·번지·다른 행정구역 뒤에 올 때만 (건물명 끝 'BUND' 등 제외)
    if dist_c and not re.search(r"(,|\d|-ro|-gil|-daero|ro|gil|-gu|-si|-do|dong)\s*$", head, re.I):
        dist_c = []
    if len(road_c) + len(dist_c) == 1:
        choice = (road_c + dist_c)[0]
    elif dist_c and len(dist_c) == 1 and has_road_elsewhere:
        choice = dist_c[0]
    elif road_c:
        here = rd.roads_at(post5) if post5 else set()
        rc = [v for v in road_c if v in here]
        if len(rc) == 1 and not (dist_c and has_road_elsewhere):
            choice = rc[0]
    if choice:
        new = match_case(tok, choice)
        notes.append(f"잘린 단어 '{tok}'→'{new}'")
        return head + new
    if (road_c or dist_c) and not has_road_elsewhere:
        flags_holder.append(f"주소 끝 잘림 '{tok}…' (복원 불가)")
    return text


def clean_qmarks(text, notes, flags):
    if text is None:
        return None
    if text.strip().upper() == "#NAME?":
        notes.append("엑셀 오류값 '#NAME?' 삭제")
        flags.append("원본 주소가 '#NAME?' 오류로 소실됨")
        return ""
    if "?" not in text:
        return text
    if re.search(r"\?{2,}", text) or text.count("?") >= 0.3 * max(len(text.replace(" ", "")), 1):
        flags.append("한글 깨짐('???') - 원문 확인 필요")
        return text
    new = re.sub(r"\s*\?\s*", " ", text).strip()
    new = re.sub(r"\s+,", ",", new)
    if new != text:
        notes.append("깨진 문자 '?' → 공백")
    return new


def tidy(text):
    text = re.sub(r"[ \t ]+", " ", text)
    text = re.sub(r"\s+,", ",", text)
    text = re.sub(r",{2,}", ",", text)
    return text.strip()


def fix_address(a1, a2, post_raw, rd):
    """return: new_a1, new_a2, new_post(int or original), notes[list], flags[list]"""
    notes, flags = [], []
    a1 = "" if a1 is None else str(a1)
    a2 = "" if a2 is None else str(a2)
    o1, o2 = a1, a2
    post5 = pad_post(post_raw)

    a1 = clean_qmarks(a1, notes, flags)
    a2 = clean_qmarks(a2, notes, flags)

    for src, dst in rd.manual.get("치환", {}).items():
        for_i = re.compile(re.escape(src), re.I)
        if for_i.search(a1) or for_i.search(a2):
            a1, a2 = for_i.sub(dst, a1), for_i.sub(dst, a2)
            notes.append(f"수동사전 '{src}'→'{dst}'")

    # 35자 경계에서 끊긴 단어 연결 (A1 끝 + A2 앞)
    mt = re.search(r"([A-Za-z][A-Za-z\-]*)$", a1)
    mh = re.match(r"(-?\s*[A-Za-z][A-Za-z\-]*)", a2)
    if mt and mh:
        t, h = mt.group(1), mh.group(1)
        merged = t + h.replace(" ", "")
        dist_vocab = {(d[0] + "-" + d[1]).upper(): d[0] + "-" + d[1] for d in DISTRICTS}
        dist_vocab.update({"SEOUL": "Seoul", "KOREA": "Korea"})
        road_up = {v.upper(): v for v in rd.vocab_roads}
        hit = dist_vocab.get(merged.upper()) or road_up.get(merged.upper())
        if not hit:
            mm = re.match(r"([A-Za-z]{3,})-?(daero|ro|gil)$", merged, re.I)
            if mm:
                c, how = rd.lookup(mm.group(1), mm.group(2), post5, loose=False)
                hit = c
        split_like = (len(o1) >= TRUNC_LEN or t.endswith("-") or h.startswith("-")
                      or re.fullmatch(r"(dae-?ro|ro|gil|gu|si|do)", h, re.I))
        if hit and split_like:
            new = match_case(merged, hit)
            a1 = a1[: len(a1) - len(t)] + new
            a2 = a2[len(h):].lstrip(" ,")
            notes.append(f"끊긴 단어 연결 '{t}'+'{h}'→'{new}'")

    # 도로명 / 행정구역 표기 수정
    a1 = district_fix(road_fix(a1, post5, rd, notes), notes)
    a2 = district_fix(road_fix(a2, post5, rd, notes), notes)

    def has_road(s):
        return bool(re.search(r"([A-Za-z]{2,}|\d)\s*-?\s*(daero|ro|gil|lo)(?![A-Za-z])|beon-?gil", s, re.I))
    road_any = has_road(a1) or has_road(a2)

    # 잘린 끝 단어 복원
    tail_flags = []
    if len(o1) >= TRUNC_LEN:
        a1 = complete_tail(a1, post5, rd, notes, tail_flags, road_any)
    if len(o2) >= TRUNC_LEN:
        a2 = complete_tail(a2, post5, rd, notes, tail_flags, road_any)
    flags += tail_flags

    # A1 이 도로명으로 끝났는데 A2 가 '-daero', 'ro' 등 접미사 조각으로 시작하면 조각 삭제
    ms = re.match(r"\s*-?\s*(dae-?ro|ro|gil)(?![A-Za-z])[\s,]*", a2, re.I)
    if ms and re.search(r"[A-Za-z]{3,}-?" + ms.group(1).replace("-", "-?") + r"$", a1.rstrip(" ,"), re.I):
        notes.append(f"Address2 앞 중복 조각 '{a2[:ms.end()].strip()}' 삭제")
        a2 = a2[ms.end():]

    # 공백 정리 (수정된 칸만)
    if a1 != o1:
        a1 = tidy(a1)
    if a2 != o2:
        a2 = tidy(a2)

    # Address1 / Address2 중복
    n1 = re.sub(r"[^a-z0-9]", "", a1.lower())
    n2 = re.sub(r"[^a-z0-9]", "", a2.lower())
    if n1 and n2 and len(min(n1, n2, key=len)) >= 8 and (n1.startswith(n2) or n2.startswith(n1)):
        if len(n2) > len(n1):
            a1 = a2
        a2 = ""
        notes.append("Address1/Address2 중복 제거")

    # 우편번호
    full = (a1 + " " + a2).strip()
    new_post = post_raw
    tp = [p for p in TEXT_POST_RE.findall(full)]
    if not post5:
        if tp:
            new_post = int(tp[-1])
            post5 = tp[-1]
            notes.append(f"우편번호 보완(주소에서 추출) {tp[-1]}")
        else:
            flags.append("우편번호 없음/오류")
    elif tp and all(p != post5 for p in tp):
        flags.append(f"우편번호 불일치(주소 내 {tp[-1]}, 우편번호칸 {post5})")

    # 검사
    if full and not re.search(r"\d", full):
        flags.append("번지(숫자) 없음")
    jibun = re.search(r"[A-Za-z]{3,}[- ]?(dong|ri|eup|myeon)(?![A-Za-z])", full, re.I)
    if full and not has_road(full) and not jibun and "한글 깨짐('???') - 원문 확인 필요" not in flags:
        flags.append("도로명/지번 없음")
    if not full:
        flags.append("주소 없음")
    if post5:
        pn = int(post5)
        for nm, kind, ranges in DISTRICTS:
            if nm == "Gyeonggi":
                continue
            if re.search(nm + r"-" + kind + r"\b", full, re.I):
                if not any(lo <= pn <= hi for lo, hi in ranges):
                    flags.append(f"구/시({nm}-{kind}) ↔ 우편번호({post5}) 불일치")
                    break
        for m in ROAD_RE.finditer(full):
            k = fold(m.group("name") + m.group("suf"))
            k = rd.alias.get(k, k)
            c = rd.road_post3.get(k)
            if c and sum(c.values()) >= 30 and k in rd.main_keys:
                if c.get(post5[:3], 0) == 0:
                    flags.append(f"도로명({rd.canon[k]}) ↔ 우편번호({post5}) 불일치 의심")
                    break

    return (a1 if a1 != o1 else o1), (a2 if a2 != o2 else o2), new_post, notes, flags


# ---------------------------------------------------------------------------
# 파일 처리
# ---------------------------------------------------------------------------
def list_inputs(src_dir, out_dir):
    files = []
    for f in sorted(glob.glob(os.path.join(src_dir, "*.xlsx"))):
        b = os.path.basename(f)
        if b.startswith("~$") or b.startswith("_") or OUT_SUFFIX + ".xlsx" in b:
            continue
        files.append(f)
    return files


def iter_rows_ro(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        for ws in wb.worksheets:
            empty = 0
            for r in ws.iter_rows(min_row=1, values_only=True):
                if len(r) < COL_POST:
                    continue
                awb = r[COL_AWB - 1]
                if awb is None:
                    empty += 1
                    if empty > 200:
                        break
                    continue
                empty = 0
                if str(awb).strip().upper() == "AWB":
                    continue
                yield r[COL_A1 - 1], r[COL_A2 - 1], r[COL_POST - 1]
    finally:
        wb.close()


def learn(files, rd):
    for i, f in enumerate(files, 1):
        try:
            for a1, a2, p in iter_rows_ro(f):
                p5 = pad_post(p)
                rd.observe(str(a1) if a1 else "", p5)
                rd.observe(str(a2) if a2 else "", p5)
        except Exception as e:
            print(f"  [학습 건너뜀] {os.path.basename(f)}: {e}")
        if i % 20 == 0 or i == len(files):
            print(f"  사전 학습 {i}/{len(files)}")
    rd.build()


def process_file(path, out_path, rd, cache, report_rows, flag_rows):
    wb = openpyxl.load_workbook(path)
    base = os.path.basename(path)
    stat = {"file": base, "rows": 0, "changed": 0, "flagged": 0}
    bold = Font(bold=True)
    for ws in wb.worksheets:
        empty = 0
        maxr = ws.max_row
        header_seen = False
        for r in range(1, maxr + 1):
            awb = ws.cell(r, COL_AWB).value
            if awb is None:
                empty += 1
                if empty > 200:
                    break
                continue
            empty = 0
            if str(awb).strip().upper() == "AWB":
                for col, title in ((COL_NOTE, "수정내역"), (COL_FLAG, "확인필요")):
                    c = ws.cell(r, col, title)
                    c.font = bold
                    c.fill = FILL_HEAD
                header_seen = True
                continue
            c1, c2, cp = ws.cell(r, COL_A1), ws.cell(r, COL_A2), ws.cell(r, COL_POST)
            key = (c1.value, c2.value, cp.value)
            if key not in cache:
                cache[key] = fix_address(c1.value, c2.value, cp.value, rd)
            n1, n2, np_, notes, flags = cache[key]
            flags = list(flags)
            if ws.title.strip().upper() == "NA":
                flags.append("배송구역 미지정(NA 시트)")
            stat["rows"] += 1
            o1, o2 = c1.value, c2.value
            changed = False
            if (o1 or "") != (n1 or ""):
                c1.value = n1 if n1 else None
                c1.fill = FILL_CHANGED
                changed = True
            if (o2 or "") != (n2 or ""):
                c2.value = n2 if n2 else None
                c2.fill = FILL_CHANGED
                changed = True
            if np_ != cp.value:
                cp.value = np_
                cp.fill = FILL_CHANGED
                changed = True
            if isinstance(cp.value, (int, float)):
                cp.number_format = "00000"
            nc, fc = ws.cell(r, COL_NOTE), ws.cell(r, COL_FLAG)
            nc.fill, fc.fill = FILL_WHITE, FILL_WHITE          # 원본 양식의 배경색 덮기
            nc.font, fc.font = FONT_BLACK, FONT_BLACK
            if notes and changed:
                nc.value = " / ".join(dict.fromkeys(notes))
                stat["changed"] += 1
            if flags:
                fc.value = " / ".join(dict.fromkeys(flags))
                fc.fill = FILL_FLAG
                stat["flagged"] += 1
            if (notes and changed) or flags:
                row = [base, ws.title, r, awb,
                       ws.cell(r, 1).value, ws.cell(r, 7).value, ws.cell(r, 8).value,
                       o1, o2, c1.value, c2.value, pad_post(cp.value) or cp.value,
                       " / ".join(dict.fromkeys(notes)) if changed else "",
                       " / ".join(dict.fromkeys(flags))]
                if notes and changed:
                    report_rows.append(row)
                if flags:
                    flag_rows.append(row)
        if header_seen:
            ws.column_dimensions["M"].width = 48
            ws.column_dimensions["N"].width = 42
    wb.save(out_path)
    return stat


def write_report(path, stats, report_rows, flag_rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "요약"
    ws.append(["주소 정제 결과", datetime.datetime.now().strftime("%Y-%m-%d %H:%M")])
    ws.append([])
    ws.append(["파일명", "전체 건수", "자동 수정 건수", "확인필요 건수"])
    for s in stats:
        ws.append([s["file"], s["rows"], s["changed"], s["flagged"]])
    ws.append(["합계", sum(s["rows"] for s in stats), sum(s["changed"] for s in stats),
               sum(s["flagged"] for s in stats)])
    ws.column_dimensions["A"].width = 50
    for c in "BCD":
        ws.column_dimensions[c].width = 14
    for c in ws[3]:
        c.font = Font(bold=True)
        c.fill = FILL_HEAD

    # 유형별 집계
    cnt = collections.Counter()
    for r in flag_rows:
        for f in r[13].split(" / "):
            cnt[re.sub(r"\s+", " ", re.sub(r"\([^)]*\)|'[^']*'", "", f)).strip()] += 1
    ws.append([])
    ws.append(["확인필요 유형", "건수"])
    ws.cell(ws.max_row, 1).font = Font(bold=True)
    for k, v in cnt.most_common():
        ws.append([k, v])

    heads = ["파일", "시트", "행", "AWB", "Route", "Receiver", "Tel",
             "원본 Address1", "원본 Address2", "수정 Address1", "수정 Address2",
             "Postcode", "수정내역", "확인필요 사유"]
    widths = [34, 6, 6, 12, 7, 18, 14, 36, 36, 38, 38, 8, 50, 45]
    for title, rows in (("확인필요", flag_rows), ("수정내역", report_rows)):
        w = wb.create_sheet(title)
        w.append(heads)
        for c in w[1]:
            c.font = Font(bold=True)
            c.fill = FILL_HEAD
        for row in rows:
            w.append(row)
        for i, wd in enumerate(widths, 1):
            w.column_dimensions[openpyxl.utils.get_column_letter(i)].width = wd
        w.freeze_panes = "A2"
        w.auto_filter.ref = f"A1:N{max(len(rows) + 1, 2)}"
    wb.save(path)


def main():
    here = os.path.dirname(os.path.abspath(__file__))
    src = os.path.abspath(sys.argv[1]) if len(sys.argv) > 1 else os.path.dirname(here)
    out = os.path.abspath(sys.argv[2]) if len(sys.argv) > 2 else os.path.join(src, "주소정제_결과")
    os.makedirs(out, exist_ok=True)
    files = list_inputs(src, out)
    print(f"원본 폴더: {src}\n결과 폴더: {out}\n대상 파일: {len(files)}개")
    if not files:
        return

    rd = RoadDict()
    dict_path = os.path.join(here, "도로명사전.json")
    manual_path = os.path.join(here, "수동사전.json")
    if os.path.exists(manual_path):
        with open(manual_path, encoding="utf-8") as fp:
            rd.manual.update(json.load(fp))
    if os.path.exists(dict_path) and "--relearn" not in sys.argv:
        with open(dict_path, encoding="utf-8") as fp:
            saved = json.load(fp)
        rd.load_json(saved.get("dict", {}))
        learned = set(saved.get("files", []))
    else:
        learned = set()
    new_files = [f for f in files if os.path.basename(f) not in learned]
    print(f"[1/3] 도로명 사전 학습 (새 파일 {len(new_files)}개)")
    learn(new_files, rd)
    try:
        with open(dict_path, "w", encoding="utf-8") as fp:
            json.dump({"files": sorted(learned | {os.path.basename(f) for f in files}),
                       "dict": rd.to_json()}, fp, ensure_ascii=False)
    except OSError:
        pass
    print(f"  도로명 {len(rd.canon)}개 학습됨")

    print("[2/3] 주소 수정 및 저장")
    cache, stats, report_rows, flag_rows = {}, [], [], []
    for i, f in enumerate(files, 1):
        stem = os.path.splitext(os.path.basename(f))[0].rstrip()
        op = os.path.join(out, stem + OUT_SUFFIX + ".xlsx")
        try:
            s = process_file(f, op, rd, cache, report_rows, flag_rows)
            stats.append(s)
            print(f"  ({i}/{len(files)}) {stem}  수정 {s['changed']} / 확인필요 {s['flagged']}")
        except Exception as e:
            print(f"  ({i}/{len(files)}) [실패] {stem}: {e}")

    print("[3/3] 결과보고서 작성")
    write_report(os.path.join(out, "_주소정제_결과보고.xlsx"), stats, report_rows, flag_rows)
    print(f"완료: 자동 수정 {sum(s['changed'] for s in stats)}건, 확인필요 {sum(s['flagged'] for s in stats)}건")


if __name__ == "__main__":
    main()
