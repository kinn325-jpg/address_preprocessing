# [설명] 이 줄은 "이 파일에는 한글이 들어 있어요" 하고 컴퓨터에게 알려 주는 줄이에요.
# -*- coding: utf-8 -*-
# [설명] 아래 따옴표 세 개(""") 사이의 글은 '설명서'예요. 컴퓨터는 실행하지 않고, 사람이 읽으라고 적어 둔 거예요.
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

# ===========================================================================
# [전체 흐름 한눈에 보기]  ← 코드를 읽기 전에 이 지도부터 보세요
# ===========================================================================
# 프로그램은 맨 아래 main() 에서 시작해서 3단계로 진행됩니다.
#
#   main()
#    │
#    ├─ [1/3] 학습   learn() → RoadDict.observe() → RoadDict.build()
#    │         모든 엑셀의 주소를 훑어 "어떤 도로명이, 어떤 우편번호에서, 몇 번 나왔는지" 센다.
#    │         많이 나온 표기를 '정답 표기'로 정하고 → 도로명사전.json 에 저장
#    │
#    ├─ [2/3] 수정   process_file() → 송장 1건마다 fix_address()
#    │         fix_address() 가 이 프로그램의 핵심. 주소 1건을 받아
#    │         (고친 주소1, 고친 주소2, 우편번호, notes, flags) 를 돌려준다.
#    │           notes(자동 수정 내역)  → 노란 칸 + M열 '수정내역'
#    │           flags(사람 확인 사유)  → 빨간 칸 N열 '확인필요'
#    │
#    └─ [3/3] 보고   write_report()
#              파일별 건수 요약 / 확인필요 목록 / 수정내역 목록을 엑셀 1개로 정리
#
# [읽는 순서 추천]  main() → process_file() → fix_address() → 나머지 부품 함수들
#
# [꼭 알아야 할 용어 3가지]
#   - fold(접기)     : 로마자 표기 차이를 없앤 '비교용 열쇠(key)'. Samsung 과 Samseong 을 같은 값으로 만든다
#   - canon(표준표기) : 데이터에서 가장 많이 쓰인 표기. 고칠 때의 '정답'
#   - notes / flags  : notes = 자동으로 고친 내용,  flags = 사람이 확인해야 하는 사유
#
# [설계 원칙]  "잘못 고치는 것이 안 고치는 것보다 위험하다"
#   → 코드 곳곳의 '20회 이상', '유사도 0.9 이상' 같은 조건은 전부
#     "확실할 때만 고치고, 애매하면 손대지 않기" 위한 안전장치다.



"""
# [설명] ==========================================================================
# [설명] [먼저 알아 두면 좋은 말]
# [설명]   - 변수   : 값을 담아 두는 '이름표 붙은 상자'예요.  예) a = 3  → a 상자에 3을 넣기
# [설명]   - 함수   : 'def 이름(...)' 으로 만드는 '일하는 방법 묶음'이에요. 이름을 부르면 그 일을 해 줘요.
# [설명]   - return : 함수가 일을 끝내고 "결과는 이거야" 하고 돌려주는 말이에요.
# [설명]   - if     : "만약 ~라면" 이에요. 맞을 때만 아래 줄을 실행해요. (elif = 아니고 만약, else = 그것도 아니면)
# [설명]   - for    : "하나씩 꺼내서 똑같은 일을 반복해" 예요.
# [설명]   - 리스트 : [ ] 안에 여러 값을 줄 세워 담은 것이에요.
# [설명]   - 딕셔너리: { } 안에 '이름: 값' 짝을 담은 것이에요. 사전처럼 이름으로 값을 찾아요.
# [설명]   - 정규식 : re.어쩌고(...) 로 쓰는 '글자 모양 찾기 틀'이에요. 예) "숫자 5개가 붙어 있는 곳을 찾아줘"
# [설명]   - '# [설명]' 으로 시작하는 줄은 공부용 설명이고, 코드는 원본 '주소정제.py' 와 한 글자도 다르지 않아요.
# [설명] ==========================================================================
# [설명] import 는 "다른 사람이 미리 만들어 둔 도구 상자를 꺼내 올게요" 라는 말이에요.
# [설명] os : 폴더와 파일을 다루는 도구 (폴더 만들기, 파일 이름 알아내기)

import os
# [설명] re : 글자 속에서 원하는 모양을 찾고 바꾸는 도구 (정규식)
import re
# [설명] sys : 프로그램 자체를 다루는 도구 (프로그램 끝내기, 실행할 때 같이 적어 준 말 읽기)
import sys
# [설명] json : 사전 내용을 파일로 저장하거나 다시 읽어 오는 도구
import json
# [설명] glob : 폴더 안에서 "이런 이름의 파일 다 찾아줘" 할 때 쓰는 도구
import glob
# [설명] difflib : 두 글자가 얼마나 비슷한지 재 주는 도구 (오타 찾기에 써요)
import difflib
# [설명] collections : 개수를 세어 주는 똑똑한 상자(Counter)가 들어 있는 도구
import collections
# [설명] datetime : 오늘 날짜와 지금 시간을 알려 주는 도구
import datetime

# [설명] try 는 "일단 해 봐" 예요. 안 되면 아래 except 로 넘어가요.
try:
    # [설명] openpyxl : 엑셀 파일을 열고, 고치고, 저장하는 도구
    import openpyxl
    # [설명] 엑셀 칸에 색 칠하기(PatternFill), 글꼴(Font), 정렬(Alignment) 도구도 꺼내 와요.
    from openpyxl.styles import PatternFill, Font, Alignment
# [설명] 만약 openpyxl 이 컴퓨터에 깔려 있지 않아서 꺼내 오지 못했다면
except ImportError:
    # [설명] "이걸 먼저 설치하세요" 하고 화면에 알려 주고
    print("openpyxl 이 필요합니다.  명령 프롬프트에서:  pip install openpyxl")
    # [설명] 프로그램을 여기서 끝내요.
    sys.exit(1)

# ---------------------------------------------------------------------------
# 설정
# ---------------------------------------------------------------------------
# [설명] 엑셀에서 왼쪽부터 몇 번째 칸인지 적어 둔 숫자예요. 3번째(C칸)=송장번호(AWB), 9번째(I칸)=주소1, 10번째(J칸)=주소2, 11번째(K칸)=우편번호
COL_AWB, COL_A1, COL_A2, COL_POST = 3, 9, 10, 11     # C, I, J, K
# [설명] 13번째(M칸)에는 '수정내역', 14번째(N칸)에는 '확인필요' 를 새로 적을 거예요.
COL_NOTE, COL_FLAG = 13, 14                          # M, N (새로 추가)
# [설명] 송장 주소는 35글자에서 뚝 잘려요. 그래서 34글자 이상인 칸은 "혹시 잘렸나?" 하고 의심해요.
TRUNC_LEN = 34            # 이 길이 이상이면 '35자 잘림' 가능성이 있는 칸으로 봄
# [설명] 결과 파일 이름 뒤에 붙일 말이에요. 예) 배송자료.xlsx → 배송자료_정제.xlsx
OUT_SUFFIX = "_정제"

# [설명] 노란색 물감: 프로그램이 고친 칸에 칠해요.
FILL_CHANGED = PatternFill("solid", fgColor="FFF2CC")   # 노랑: 수정된 칸
# [설명] 빨간색(살구색) 물감: 사람이 꼭 확인해야 하는 칸에 칠해요.
FILL_FLAG = PatternFill("solid", fgColor="F8CBAD")      # 빨강: 확인 필요
# [설명] 하늘색 물감: 제목 줄에 칠해요.
FILL_HEAD = PatternFill("solid", fgColor="D9E1F2")
# [설명] 흰색 물감: 원래 있던 색을 지울 때 써요.
FILL_WHITE = PatternFill("solid", fgColor="FFFFFF")
# [설명] 검은색 글씨예요.
FONT_BLACK = Font(color="000000")

# 행정구역 (이름, 종류, 우편번호 범위들)
# [설명] DISTRICTS 는 '동네 이름표 목록'이에요. 한 줄이 (이름, 종류, 우편번호 범위) 예요.
DISTRICTS = [
    # [설명] 예) 송파(Songpa)는 '구(gu)'이고, 우편번호는 05500 ~ 05899 사이예요.
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
# [설명] 서울에 있는 구 이름 모음이에요.
SEOUL_GU = {"Songpa", "Gangnam", "Gangdong", "Seocho", "Gwangjin"}
# [설명] 성남시에 있는 구 이름 모음이에요.
SEONGNAM_GU = {"Bundang", "Sujeong", "Jungwon"}
# [설명] '-gil' 앞에 오지만 도로 이름은 아닌 단어들이에요. (예: 12beon-gil 의 beon 은 '번'이라는 뜻)
NOT_ROAD_NAMES = {"beon", "bun", "beun", "ben", "beongil", "bungil", "dae", "daer", "dea"}
# 도로명 앞 단어로 붙이면 안 되는 단어들
# [설명] 도로 이름 앞에 붙이면 안 되는 단어들이에요. (층, 호, 빌딩, 서울 같은 말은 도로 이름의 일부가 아니니까요)
PRE_STOP = {"gu", "si", "do", "gil", "ro", "f", "fl", "floor", "dong", "ho", "bldg", "building",
            "tower", "apt", "the", "and", "of", "seoul", "korea", "b", "unit", "room", "rm", "no",
            "ga", "ri", "eup", "myeon", "center", "centre", "plaza", "city", "daero", "st", "nd",
            "rd", "th", "kr", "south", "republic", "suite", "office", "gil", "bd", "rep"}

# 자동으로 잡기 어려운 발음식 표기 (기본 별칭, 수동사전.json 으로 추가 가능)
# [설명] 소리 나는 대로 적어서 자주 틀리는 표기예요. '틀린 표기: 올바른 표기' 짝이에요. (선릉로는 Seolleung-ro 가 맞아요)
DEFAULT_ALIASES = {
    "Seonneung-ro": "Seolleung-ro", "Sunneung-ro": "Seolleung-ro",
    "Seonreung-ro": "Seolleung-ro", "Sunreung-ro": "Seolleung-ro",
    "Heonneung-ro": "Heolleung-ro", "Heonreung-ro": "Heolleung-ro", "Hunreung-ro": "Heolleung-ro",
}


# [설명] [함수] 두 글자의 차이가 '오타' 정도로 작은지 알아봐요. 차이가 크면 "다른 도로야"(False) 라고 답해요.
def strict_similar(a, b, raw_a=None, raw_b=None):
    """두 철자 차이가 '오타' 수준인지 (2글자 이상 삽입/삭제는 다른 도로로 봄: Balibong / Balibongbuk)"""
    # [설명] 원래 글자(접기 전)를 둘 다 받았다면, 먼저 그걸로 비교해요.
    if raw_a and raw_b:
        # [설명] 첫 번째 글자를 소문자로 바꾸고, 알파벳이 아닌 것(-, 띄어쓰기)은 지워요.
        ra = re.sub(r"[^a-z]", "", raw_a.lower())
        # [설명] 두 번째 글자도 똑같이 해요.
        rb = re.sub(r"[^a-z]", "", raw_b.lower())
        # [설명] 두 글자를 나란히 놓고 "어디가 더 들어갔고(insert), 어디가 빠졌는지(delete)" 하나씩 살펴봐요.
        for tag, i1, i2, j1, j2 in difflib.SequenceMatcher(None, ra, rb).get_opcodes():
            # [설명] 글자가 2개 이상 통째로 더 있거나 빠져 있으면
            if tag in ("insert", "delete") and max(i2 - i1, j2 - j1) >= 2:
                # [설명] 오타가 아니라 다른 도로예요. (Balibong 과 Balibongbuk 은 다른 길)
                return False
    # [설명] 이번에는 a 와 b 를 비교할 준비를 해요.
    sm = difflib.SequenceMatcher(None, a, b)
    # [설명] 다른 곳을 하나씩 살펴봐요.
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        # [설명] 글자가 2개 이상 통째로 더 있거나 빠져 있으면
        if tag in ("insert", "delete") and max(i2 - i1, j2 - j1) >= 2:
            # [설명] 다른 도로예요.
            return False
        # [설명] 글자가 3개 이상 다른 글자로 바뀌어 있어도
        if tag == "replace" and max(i2 - i1, j2 - j1) > 2:
            # [설명] 다른 도로예요.
            return False
    # [설명] 여기까지 왔으면 오타 수준이에요. 얼마나 비슷한지 점수(0 ~ 1, 1이면 똑같음)를 돌려줘요.
    return sm.ratio()


# ---------------------------------------------------------------------------
# 로마자 표기 변형을 같은 값으로 접는 함수 (Samsung/Samseong, Youngdong/Yeongdong …)
# ---------------------------------------------------------------------------
# [설명] [함수] 글자를 '접어서' 비슷한 소리끼리 같은 모양으로 만들어요. Samsung 과 Samseong 이 같은 길인 걸 알아보려고요.
def fold(s):
    # [설명] 소문자로 바꾸고 알파벳만 남겨요.
    s = re.sub(r"[^a-z]", "", s.lower())
    # [설명] y 뒤에 모음(a,e,i,o,u)이 없으면 i 로 바꿔요. (Olympic 과 Olimpic 을 같게)
    s = re.sub(r"y(?![aeiou])", "i", s)          # Olympic / Olimpic
    # [설명] 비슷한 소리 짝꿍 목록이에요. 왼쪽 글자를 오른쪽 글자로 통일해요. (eo→u, k→g, p→b ...)
    for a, b in (("yeo", "yu"), ("you", "yu"), ("eo", "u"), ("oo", "u"), ("ou", "u"),
                 ("wh", "w"), ("sh", "s"), ("ee", "i"), ("k", "g"),
                 ("p", "b"), ("t", "d"), ("l", "r"), ("ch", "j"), ("c", "g")):
        # [설명] 짝꿍대로 글자를 바꿔요.
        s = s.replace(a, b)
    # [설명] 같은 글자가 연달아 있으면 하나로 줄여서 돌려줘요. (ss → s)
    return re.sub(r"(.)\1+", r"\1", s)


# [설명] [함수] 우편번호를 '5글자 숫자'로 맞춰요. 이상한 값이면 None(없음)을 돌려줘요.
def pad_post(v):
    """우편번호 값을 5자리 문자열로. 유효하지 않으면 None"""
    # [설명] 값이 비어 있으면
    if v is None:
        # [설명] "없음" 이라고 답해요.
        return None
    # [설명] 엑셀이 5854.0 처럼 소수로 준 경우에는
    if isinstance(v, float) and v.is_integer():
        # [설명] 5854 처럼 정수로 바꿔요.
        v = int(v)
    # [설명] 숫자가 아닌 글자는 모두 지워요.
    d = re.sub(r"\D", "", str(v))
    # [설명] 숫자가 4개뿐이면 (엑셀이 05854 의 맨 앞 0 을 지워 버린 경우예요)
    if len(d) == 4:
        # [설명] 앞에 0 을 다시 붙여요.
        d = "0" + d
    # [설명] 숫자가 딱 5개면 그걸 돌려주고, 아니면 "없음" 이라고 답해요.
    return d if len(d) == 5 else None


# [설명] [함수] 원래 글자가 전부 대문자였으면, 고친 글자도 대문자로 맞춰 줘요. (GANGNAM-GU 처럼)
def match_case(src, canon):
    # [설명] 원래 글자에서 알파벳만 뽑아요.
    letters = re.sub(r"[^A-Za-z]", "", src)
    # [설명] 알파벳이 2글자 이상이고 전부 대문자라면
    if letters and letters.isupper() and len(letters) > 1:
        # [설명] 고친 글자도 대문자로 바꿔서 돌려줘요.
        return canon.upper()
    # [설명] 아니면 고친 글자를 그대로 돌려줘요.
    return canon


# [설명] ROAD_RE 는 '도로 이름 찾기 틀'이에요. 주소에서 "앞단어(pre) + 이름(name) + 사이기호(sep) + 끝말(suf: daero, ro, gil)" 모양을 찾아요. 예) Teheran-ro
ROAD_RE = re.compile(
    r"(?<![A-Za-z0-9])(?:(?P<pre>[A-Za-z]+) )?(?P<name>(?!dae-?ro(?![A-Za-z]))[A-Za-z]+)"
    r"(?P<sep>\s*-\s*|\s+|)(?P<suf>dae-ro|daero|ro|lo|gil)(?![A-Za-z])", re.I)

# [설명] DIST_RE 는 '구·시·도 찾기 틀'이에요. "이름(4글자 이상) + 끝말(gu, ku, si, city, do)" 모양을 찾아요. 예) Gangnam-gu
DIST_RE = re.compile(
    r"(?<![A-Za-z])(?P<name>[A-Za-z]{4,})(?P<sep>\s*-\s*|\s+|)(?P<suf>gu|ku|si|shi|city|do)(?![A-Za-z])",
    re.I)

# [설명] TEXT_POST_RE 는 '주소 글자 속 우편번호 찾기 틀'이에요. 04~06 또는 12~13 으로 시작하는 숫자 5개를 찾아요.
TEXT_POST_RE = re.compile(r"(?<![\d-])(0[4-6]\d{3}|1[2-3]\d{3})(?![\d-])")


# ---------------------------------------------------------------------------
# 도로명 사전
# ---------------------------------------------------------------------------
# [설명] [설계도] RoadDict 는 '도로명 사전'이에요. class 는 "이런 물건을 만들 거야" 하고 적어 두는 설계도예요.
class RoadDict:
    # [설명] __init__ 은 사전을 처음 만들 때 딱 한 번 실행돼요. 빈 상자들을 준비해요. (self 는 '이 사전 자신'이라는 뜻)
    def __init__(self):
        # [설명] spell : 도로마다 "어떤 표기가 몇 번 나왔는지" 세는 상자
        self.spell = collections.defaultdict(collections.Counter)   # key -> 표기 -> 횟수
        # [설명] total : 도로마다 "전부 몇 번 나왔는지" 세는 상자
        self.total = collections.Counter()                           # key -> 전체 횟수
        # [설명] post_roads : 우편번호마다 "그 동네에 어떤 도로가 나왔는지" 세는 상자
        self.post_roads = collections.defaultdict(collections.Counter)  # 우편번호5 -> key
        # [설명] road_post3 : 도로마다 "우편번호 앞 3자리가 무엇이었는지" 세는 상자
        self.road_post3 = collections.defaultdict(collections.Counter)  # key -> 우편번호 앞3
        # [설명] canon : 도로마다 '대표 표기(가장 올바른 표기)'를 적어 두는 상자
        self.canon = {}
        # [설명] manual : 사람이 직접 적어 준 사전(수동사전)을 담는 상자
        self.manual = {"치환": {}, "도로명_추가": []}
        # [설명] _fuzzy_cache : 한 번 찾아본 답을 적어 두는 메모장 (같은 걸 또 찾지 않으려고요)
        self._fuzzy_cache = {}

    # --- 학습 ---
    # [설명] [일하는 방법] 주소 한 칸을 보고, 그 안의 도로 이름을 배워요.
    def observe(self, text, post5):
        # [설명] 주소가 비어 있으면
        if not text:
            # [설명] 할 일이 없으니 그냥 끝내요.
            return
        # [설명] 주소에서 도로 이름 모양을 전부 찾아서 하나씩 꺼내요.
        for m in ROAD_RE.finditer(text):
            # [설명] 찾은 것에서 이름(name), 사이기호(sep), 끝말(suf)을 꺼내요. 예) Teheran / - / ro
            name, sep, suf = m.group("name"), m.group("sep"), m.group("suf")
            # [설명] 도로 이름이 아닌 단어이거나, 이름이 3글자보다 짧으면
            if name.lower() in NOT_ROAD_NAMES or len(name) < 3:
                # [설명] 이건 건너뛰고 다음 것으로 가요. (continue = 건너뛰기)
                continue
            # [설명] 'Songpadae-daero' 처럼 dae 가 두 번 들어간 경우라면
            if suf.lower().startswith("dae") and name.lower().endswith("dae") and len(name) > 5:
                # [설명] 이름 끝의 dae 3글자를 잘라 내요.
                name = name[:-3]          # Songpadae-daero → Songpa-daero
            # [설명] 이름과 끝말을 붙여서 접어요. 이게 이 도로의 '열쇠(k)'예요. 표기가 조금 달라도 열쇠는 같아져요.
            k = fold(name + suf)
            # [설명] 이 도로를 한 번 봤다고 +1 해요.
            self.total[k] += 1
            # [설명] 이름과 끝말 사이에 '-' 가 들어간 반듯한 표기일 때만
            if "-" in sep:
                # [설명] "이 표기를 한 번 봤어요" 하고 +1 해요. (첫 글자는 대문자, 끝말은 소문자로 맞춰서)
                self.spell[k][name.capitalize() + "-" + suf.lower()] += 1
            # [설명] 우편번호가 있으면
            if post5:
                # [설명] "이 우편번호 동네에서 이 도로를 봤어요" 하고 +1 해요.
                self.post_roads[post5][k] += 1
                # [설명] "이 도로는 우편번호 앞 3자리가 이거였어요" 하고 +1 해요.
                self.road_post3[k][post5[:3]] += 1

    # [설명] [일하는 방법] 지금까지 배운 것을 정리해서 사전을 완성해요.
    def build(self):
        # [설명] 대표 표기 상자를 비우고 새로 시작해요.
        self.canon = {}
        # [설명] 도로 열쇠(k)와 그 표기 개수표(c)를 하나씩 꺼내요.
        for k, c in self.spell.items():
            # [설명] 3번 이상 본 도로만 믿어요. (한두 번 나온 건 오타일 수 있으니까요)
            if self.total[k] >= 3 and c:
                # [설명] 가장 많이 쓰인 표기를 그 도로의 '대표 표기'로 정해요.
                self.canon[k] = c.most_common(1)[0][0]
        # [설명] 사람이 수동사전에 직접 추가한 도로 이름을 하나씩 꺼내요.
        for r in self.manual.get("도로명_추가", []):
            # [설명] 그 도로의 열쇠를 만들고
            k = fold(r)
            # [설명] 대표 표기로 넣어요.
            self.canon[k] = r
            # [설명] 사람이 알려 준 거니까 50번 본 것만큼 믿어 줘요.
            self.total[k] = max(self.total[k], 50)
        # [설명] 지금 알고 있는 도로 열쇠 목록을 만들어요.
        self.known_keys = list(self.canon)
        # [설명] 메모장을 비워요.
        self._fuzzy_cache = {}

        # (a) 오타 변형 → 대표 도로명 (예: Sonpa-daero → Songpa-daero)
        # [설명] alias 는 '틀린 열쇠 → 올바른 열쇠' 를 적는 상자예요.
        self.alias = {}
        # [설명] 알고 있는 도로를 하나씩 꺼내요.
        for k in self.known_keys:
            # [설명] "이 도로와 아주 비슷한데 훨씬 자주 나오는 도로가 있나?" 찾아봐요.
            best = self._best_target(k, self.total[k], raw=self.canon[k])
            # [설명] 그런 도로가 있으면
            if best:
                # [설명] "이건 저 도로의 오타예요" 하고 적어 둬요.
                self.alias[k] = best
        # [설명] 처음부터 넣어 둔 기본 별칭을 복사해요.
        aliases = dict(DEFAULT_ALIASES)
        # [설명] 사람이 수동사전에 적어 준 별칭도 더해요.
        aliases.update(self.manual.get("도로명_별칭", {}))
        # [설명] (틀린 표기, 올바른 표기) 짝을 하나씩 꺼내요.
        for src, dst in aliases.items():
            # [설명] 올바른 표기의 열쇠를 만들어요.
            dk = fold(dst)
            # [설명] 그 도로가 아직 사전에 없으면
            if dk not in self.canon:
                # [설명] 대표 표기로 넣고
                self.canon[dk] = dst
                # [설명] 50번 본 것만큼 믿어 줘요.
                self.total[dk] = max(self.total[dk], 50)
            # [설명] 틀린 표기의 열쇠가 올바른 열쇠와 다르면
            if fold(src) != dk:
                # [설명] "틀린 열쇠 → 올바른 열쇠" 로 이어 줘요.
                self.alias[fold(src)] = dk
        # [설명] 도로 열쇠 목록을 다시 만들어요.
        self.known_keys = list(self.canon)

        # (b) 띄어쓰기로 생긴 조각 (예: 'Hwang Saeul-ro' 의 'Saeul-ro')
        # [설명] fragment 는 '도로 이름 조각' 모음이에요. 띄어 써서 이름이 반만 남은 것들이에요.
        self.fragment = set()
        # [설명] 도로 열쇠를 긴 것부터 차례로 줄 세워요.
        by_len = sorted(self.known_keys, key=len, reverse=True)
        # [설명] 도로를 하나씩 꺼내서(k)
        for k in self.known_keys:
            # [설명] 그것보다 긴 도로(K)와 비교해요.
            for K in by_len:
                # [설명] 길이 차이가 1글자 이하가 되면 더 볼 필요가 없으니
                if len(K) <= len(k) + 1:
                    # [설명] 비교를 멈춰요. (break = 반복 그만하기)
                    break
                # [설명] 긴 도로가 짧은 것으로 끝나고(예: hwangsaeulro 와 saeulro), 긴 도로가 5배 이상 자주 나오면
                if K.endswith(k) and self.total[K] >= 5 * self.total[k]:
                    # [설명] 짧은 것은 '조각'이라고 적어 두고
                    self.fragment.add(k)
                    # [설명] 비교를 멈춰요.
                    break

        # [설명] main_keys : 오타도 아니고 조각도 아닌 '진짜 도로' 열쇠 모음이에요.
        self.main_keys = {k for k in self.known_keys if k not in self.alias and k not in self.fragment}
        # [설명] vocab_roads : 진짜 도로 중에서 10번 이상 나온 것들의 대표 표기 목록이에요. 잘린 주소를 되살릴 때 써요.
        self.vocab_roads = [self.canon[k] for k in self.main_keys if self.total[k] >= 10]

    # [설명] [일하는 방법] 이 열쇠(k)가 '어떤 도로의 오타인지' 가장 그럴듯한 후보를 찾아요.
    def _best_target(self, k, cnt, post5=None, need_evidence=False, raw=None):
        # [설명] 아직 찾은 후보가 없어요.
        best = None
        # [설명] 알고 있는 도로 중에서 90% 이상 비슷한 것을 5개까지 꺼내서 하나씩 봐요.
        for cand in difflib.get_close_matches(k, self.known_keys, n=5, cutoff=0.9):
            # [설명] 자기 자신이면
            if cand == k:
                # [설명] 건너뛰어요.
                continue
            # [설명] 후보 도로가 몇 번 나왔는지 꺼내요.
            tc = self.total[cand]
            # [설명] 후보가 20번도 안 나왔거나, 나보다 10배 이상 많이 나오지 않았으면
            if tc < 20 or tc < 10 * max(cnt, 1):
                # [설명] 믿기 어려우니 건너뛰어요.
                continue
            # [설명] 차이가 정말 '오타' 수준인지 확인해요.
            ratio = strict_similar(k, cand, raw, self.canon[cand])
            # [설명] 오타 수준이 아니거나 90% 보다 덜 비슷하면
            if not ratio or ratio < 0.9:
                # [설명] 건너뛰어요.
                continue
            # [설명] "증거가 필요해" 라고 했을 때는
            if need_evidence:
                # [설명] 같은 우편번호 동네에서 그 후보 도로를 본 적이 있는지 확인해요.
                seen_here = post5 and self.post_roads.get(post5, {}).get(cand, 0) > 0
                # [설명] 본 적도 없고, 아주 많이(92% 이상) 비슷하지도 않으면
                if not (seen_here or ratio >= 0.92):
                    # [설명] 건너뛰어요.
                    continue
            # [설명] 지금까지 찾은 것보다 더 좋은 후보라면 (더 비슷하거나 더 자주 나오면)
            if best is None or (ratio, tc) > best[0]:
                # [설명] 이 후보를 기억해요.
                best = ((ratio, tc), cand)
        # [설명] 가장 좋은 후보를 돌려줘요. 없으면 "없음" 이라고 답해요.
        return best[1] if best else None

    # --- 조회 ---
    # [설명] [일하는 방법] 도로 이름을 사전에서 찾아서 (올바른 표기, 어떻게 찾았는지) 를 돌려줘요.
    def lookup(self, name, suf, post5, loose=True):
        """(표준표기, 방식) 또는 (None, None)"""
        # [설명] dae 가 두 번 들어간 이름이면
        if suf.lower().startswith("dae") and name.lower().endswith("dae") and len(name) > 5:
            # [설명] 끝의 dae 를 잘라 내요.
            name = name[:-3]
        # [설명] 열쇠를 만들어요.
        k = fold(name + suf)
        # [설명] 이 열쇠가 '오타 목록'에 있으면
        if k in self.alias:
            # [설명] 올바른 열쇠로 바꿔요.
            t = self.alias[k]
            # [설명] 그 올바른 열쇠도 또 다른 것의 오타라면 끝까지 따라가요. (서로 가리켜서 빙빙 도는 건 멈춰요)
            while t in self.alias and self.alias[t] != k:
                # [설명] 한 칸 더 따라가요.
                t = self.alias[t]
            # [설명] 올바른 표기와 "fuzzy(비슷해서 고쳤어요)" 를 돌려줘요.
            return self.canon[t], "fuzzy"
        # [설명] 이 열쇠가 사전에 그대로 있으면
        if k in self.canon:
            # [설명] 대표 표기와 "exact(딱 맞았어요)" 를 돌려줘요.
            return self.canon[k], "exact"
        # [설명] "느슨하게 찾지 마" 라고 했으면
        if not loose:
            # [설명] 여기서 "못 찾았어요" 하고 끝내요.
            return None, None
        # [설명] (열쇠, 우편번호) 짝을 메모장 이름으로 써요.
        ck = (k, post5)
        # [설명] 메모장에 아직 답이 없으면
        if ck not in self._fuzzy_cache:
            # [설명] 비슷한 도로를 찾아서 메모장에 적어 둬요.
            self._fuzzy_cache[ck] = self._best_target(k, self.total.get(k, 0), post5, True, raw=name + suf)
        # [설명] 메모장에서 답을 꺼내요.
        cand = self._fuzzy_cache[ck]
        # [설명] 후보가 있으면
        if cand:
            # [설명] 그 후보가 오타 목록에 있으면 올바른 것으로 바꾸고
            cand = self.alias.get(cand, cand)
            # [설명] 올바른 표기와 "fuzzy" 를 돌려줘요.
            return self.canon[cand], "fuzzy"
        # [설명] 끝까지 못 찾았으면 (없음, 없음) 을 돌려줘요.
        return None, None

    # [설명] [일하는 방법] 이 우편번호 동네에서 본 '진짜 도로' 이름들을 모아서 돌려줘요.
    def roads_at(self, post5):
        return {self.canon[k] for k in self.post_roads.get(post5, {}) if k in self.main_keys}

    # --- 저장 / 불러오기 ---
    # [설명] [일하는 방법] 배운 내용을 파일에 저장할 수 있는 모양(딕셔너리)으로 바꿔요.
    def to_json(self):
        # [설명] 상자 네 개를 이름표와 함께 묶어서 돌려줘요.
        return {
            "spell": {k: dict(v) for k, v in self.spell.items()},
            "total": dict(self.total),
            "post_roads": {k: dict(v) for k, v in self.post_roads.items()},
            "road_post3": {k: dict(v) for k, v in self.road_post3.items()},
        }

    # [설명] [일하는 방법] 파일에 저장해 둔 사전 내용을 다시 읽어서 상자에 채워 넣어요.
    def load_json(self, d):
        # [설명] 저장된 '표기 개수표'를 하나씩 꺼내서
        for k, v in d.get("spell", {}).items():
            # [설명] 내 상자에 더해요.
            self.spell[k].update(v)
        # [설명] '전체 횟수'도 더해요.
        self.total.update(d.get("total", {}))
        # [설명] '우편번호별 도로'를 하나씩 꺼내서
        for k, v in d.get("post_roads", {}).items():
            # [설명] 내 상자에 더해요.
            self.post_roads[k].update(v)
        # [설명] '도로별 우편번호 앞 3자리'를 하나씩 꺼내서
        for k, v in d.get("road_post3", {}).items():
            # [설명] 내 상자에 더해요.
            self.road_post3[k].update(v)


# ---------------------------------------------------------------------------
# 주소 한 칸 수정
# ---------------------------------------------------------------------------
# [설명] [함수] 주소에 적힌 구·시·도 이름을 올바른 표기로 고쳐요. (Kangnam-ku → Gangnam-gu)
def district_fix(text, notes):
    # [설명] 동네 이름표에서 이름만 뽑아요.
    names = [d[0] for d in DISTRICTS]
    # [설명] '접은 이름 → 진짜 이름' 표를 만들어요.
    folded = {fold(n): n for n in names}
    # [설명] '이름 → 종류(gu, si, do)' 표를 만들어요.
    kinds = {d[0]: d[1] for d in DISTRICTS}

    # [설명] [작은 함수] 주소에서 찾은 구·시·도 하나를 "무엇으로 바꿀지" 정해요.
    def repl(m):
        # [설명] 찾은 글자 전체(raw)와 이름 부분(name)을 꺼내요.
        raw, name = m.group(0), m.group("name")
        # [설명] 끝말을 소문자로 꺼내요.
        suf = m.group("suf").lower()
        # [설명] 끝말이 gu 나 ku 면 '구', do 면 '도', 나머지(si, shi, city)는 '시'로 봐요.
        kind = "gu" if suf in ("gu", "ku") else "do" if suf == "do" else "si"
        # [설명] 이름을 접어요.
        fn = fold(name)
        # [설명] 이름표에서 찾아봐요.
        hit = folded.get(fn)
        # [설명] 이름표에 없으면
        if not hit:
            # [설명] 같은 종류(구는 구끼리, 시는 시끼리) 이름만 모으고
            same_kind = [k for k, n in folded.items() if kinds[n] == kind]
            # [설명] 그중에서 84% 이상 비슷한 것을 하나 찾아요.
            cm = difflib.get_close_matches(fn, same_kind, n=1, cutoff=0.84)
            # [설명] 찾았고, 이름이 5글자 이상이고, 차이가 오타 수준이면
            if cm and len(fn) >= 5 and (strict_similar(fn, cm[0], name, folded[cm[0]]) or 0) >= 0.84:
                # [설명] 그 이름으로 정해요.
                hit = folded[cm[0]]
        # [설명] 그래도 못 찾았으면
        if not hit:
            # [설명] 고치지 않고 원래 글자를 그대로 돌려줘요.
            return raw
        # [설명] 올바른 표기를 만들어요. 예) Gangnam + - + gu (원래 대문자였으면 대문자로)
        canon = match_case(raw, hit + "-" + kinds[hit])
        # [설명] 원래 글자와 다르면
        if canon.lower() != raw.lower():
            # [설명] "이렇게 고쳤어요" 하고 수정내역에 적고
            notes.append(f"행정구역 표기 '{raw}'→'{canon}'")
            # [설명] 고친 글자를 돌려줘요.
            return canon
        # [설명] 같으면 원래 글자를 그대로 돌려줘요.
        return raw
    # [설명] 주소 안의 모든 구·시·도 모양에 위의 작은 함수를 적용해서 돌려줘요.
    return DIST_RE.sub(repl, text)


# [설명] [함수] 주소에 적힌 도로 이름을 사전의 올바른 표기로 고쳐요. (Taeheran-ro → Teheran-ro)
def road_fix(text, post5, rd, notes):
    # [설명] [작은 함수] 주소에서 찾은 도로 이름 하나를 "무엇으로 바꿀지" 정해요.
    def repl(m):
        # [설명] 찾은 글자 전체를 꺼내요.
        raw = m.group(0)
        # [설명] 앞단어, 이름, 사이기호, 끝말을 꺼내요.
        pre, name, sep, suf = m.group("pre"), m.group("name"), m.group("sep"), m.group("suf")
        # [설명] 도로 이름이 아닌 단어면
        if name.lower() in NOT_ROAD_NAMES:
            # [설명] 고치지 않고 그대로 돌려줘요.
            return raw
        # [설명] '이게 도로 이름일지도 몰라' 하는 후보 목록이에요.
        cands = []
        # [설명] 앞단어가 있으면 (예: 'Hwang Saeul-ro' 의 Hwang)
        if pre:
            # [설명] 앞단어가 주소의 몇 번째 글자부터인지 알아내고
            st = m.start("pre")
            # [설명] 그 바로 앞 글자가 무엇인지 봐요.
            before = m.string[st - 1] if st > 0 else " "
            # [설명] 앞단어가 금지 단어가 아니고, 바로 앞이 - 나 . 이 아니고, 2글자 이상이면
            if pre.lower() not in PRE_STOP and before not in "-." and len(pre) >= 2:
                # [설명] '앞단어+이름' 을 붙인 것(Hwangsaeul)을 첫 번째 후보로 넣어요.
                cands.append((pre + name, True))
        # [설명] 이름만 있는 것도 후보로 넣어요.
        cands.append((name, False))
        # [설명] 후보를 하나씩 꺼내서 확인해요.
        for nm, uses_pre in cands:
            # [설명] 사이에 기호가 있거나 이름이 6글자 이상일 때만 '비슷한 것 찾기'를 허락해요. (짧은 이름은 잘못 고치기 쉬워서요)
            loose = bool(sep.strip()) or len(nm) >= 6
            # [설명] 사전에서 찾아봐요.
            canon, how = rd.lookup(nm, suf, post5, loose=loose)
            # [설명] 사전에 없으면
            if not canon:
                # [설명] 다음 후보로 넘어가요.
                continue
            # [설명] 앞단어를 붙여서 '비슷한 것'으로 찾았는데, 앞 3글자 소리가 다르면
            if uses_pre and how == "fuzzy" and not fold(canon).startswith(fold(pre)[:3]):
                # [설명] 믿지 않고 다음 후보로 넘어가요.
                continue
            # [설명] 주소에서 '고칠 부분'만 잘라 내요. (앞단어를 안 썼으면 앞단어는 빼고)
            src = raw if uses_pre else raw[len(pre) + 1:] if pre else raw
            # [설명] 올바른 표기로 바꾸되, 대문자·소문자는 원래대로 맞춰요.
            new = match_case(src, canon)
            # [설명] 바꿔도 똑같으면
            if new.lower() == src.lower():
                # [설명] 고칠 게 없으니 그대로 돌려줘요.
                return raw
            # [설명] 비슷해서 고쳤으면 '도로명 오타', 딱 맞는 걸 반듯하게 고쳤으면 '도로명 표기' 라고 이름 붙여요.
            tag = "도로명 오타" if how == "fuzzy" else "도로명 표기"
            # [설명] 수정내역에 적어요.
            notes.append(f"{tag} '{src}'→'{new}'")
            # [설명] 고친 글자를 돌려줘요. (앞단어를 안 썼으면 앞단어를 다시 앞에 붙여서)
            return new if uses_pre or not pre else pre + " " + new
        # [설명] 맞는 후보가 하나도 없으면 그대로 돌려줘요.
        return raw
    # [설명] 주소 안의 모든 도로 이름 모양에 위의 작은 함수를 적용해서 돌려줘요.
    return ROAD_RE.sub(repl, text)


# [설명] [함수] 주소가 중간에 잘렸을 때 "그 다음에 무슨 말이 올 차례였을까?" 를 맞혀요.
def expected_next(text):
    """텍스트 끝에 올 것으로 기대되는 행정구역"""
    # [설명] 마지막으로 찾은 것을 기억할 상자예요.
    last = None
    # [설명] 주소에서 구 이름, 시 이름, 경기도, 서울을 전부 찾아요.
    for m in re.finditer(r"(Songpa|Gangnam|Gangdong|Seocho|Gwangjin|Bundang|Sujeong|Jungwon)-gu|"
                         r"(Hanam|Seongnam|Gwangju)-si|(Gyeonggi)-do|(Seoul)", text, re.I):
        # [설명] 찾을 때마다 덮어써서, 끝나면 '가장 마지막 것'만 남아요.
        last = m
    # [설명] 하나도 못 찾았으면
    if not last:
        # [설명] "모르겠어요" (빈 목록) 하고 답해요.
        return []
    # [설명] 마지막이 '구'였다면
    if last.group(1):
        # [설명] 구 이름의 첫 글자만 대문자로 만들고
        g = last.group(1).capitalize()
        # [설명] 서울의 구면 다음은 Seoul, 아니면 다음은 Seongnam-si 예요.
        return ["Seoul"] if g in SEOUL_GU else ["Seongnam-si"]
    # [설명] 마지막이 '시'였다면
    if last.group(2):
        # [설명] 다음은 Gyeonggi-do 예요.
        return ["Gyeonggi-do"]
    # [설명] 마지막이 '도'나 '서울'이었다면 다음은 나라 이름(Korea)이에요.
    return ["Korea", "Republic of Korea", "South Korea"]


# [설명] [함수] 35글자에서 잘려 버린 주소의 '마지막 단어'를 되살려요. 확실할 때만 고쳐요.
def complete_tail(text, post5, rd, notes, flags_holder, has_road_elsewhere):
    """35자에서 잘린 마지막 단어 복원"""
    # [설명] 주소가 비었거나, 끝이 글자·숫자·- 가 아니면 (마침표나 쉼표로 끝나면 안 잘린 거예요)
    if not text or not re.search(r"[A-Za-z0-9-]$", text):
        # [설명] 그대로 돌려줘요.
        return text
    # 1) 숫자 조각 → 우편번호
    # [설명] 끝이 'Seoul 058' 처럼 숫자 조각으로 끝나는지 봐요.
    m = re.search(r"(?:Seoul|Korea|-do|-si|-gu|Gyeonggi|code|\()[\s,:]*(\d{1,4})$", text, re.I)
    # [설명] 그 숫자 조각이 우편번호 앞부분과 같고, 주소에 숫자 5개짜리가 없으면
    if m and post5 and post5.startswith(m.group(1)) and not re.search(r"\d{5}", text):
        # [설명] 조각을 꺼내서
        frag = m.group(1)
        # [설명] "잘린 우편번호를 이었어요" 하고 적고
        notes.append(f"잘린 우편번호 '{frag}'→'{post5}'")
        # [설명] 조각 자리에 우편번호 전체를 넣어서 돌려줘요.
        return text[: len(text) - len(frag)] + post5
    # [설명] 끝이 숫자 뒤에 'b', 'beo', 'beon-g' 처럼 'beon-gil(번길)'이 잘린 모양인지 봐요.
    m = re.search(r"(\d)\s*(b|be|beo|beon|beon-|beon-g|beon-gi|beong|beongi|bun|bun-|bun-g|bun-gi)$", text, re.I)
    # [설명] 그렇다면
    if m:
        # [설명] 잘린 조각을 꺼내고
        frag = text[m.start(2):]
        # [설명] 원래 대문자였으면 BEON-GIL, 아니면 beon-gil 로 정해요.
        new = "BEON-GIL" if frag.isupper() else "beon-gil"
        # [설명] 수정내역에 적고
        notes.append(f"잘린 단어 '{frag}'→'{new}'")
        # [설명] 조각 자리에 넣어서 돌려줘요.
        return text[:m.start(2)] + new
    # [설명] 끝이 '12-g' 처럼 '-gil(길)'이 잘린 모양인지 봐요.
    m = re.search(r"(\d)-(g|gi)$", text, re.I)
    # [설명] 그렇다면
    if m:
        # [설명] 잘린 조각을 꺼내고
        frag = text[m.start(2):]
        # [설명] GIL 또는 gil 로 정해요.
        new = "GIL" if frag.isupper() else "gil"
        # [설명] 수정내역에 적고
        notes.append(f"잘린 단어 '-{frag}'→'-{new}'")
        # [설명] 조각 자리에 넣어서 돌려줘요.
        return text[:m.start(2)] + new
    # [설명] 주소 맨 끝의 영어 단어를 찾아요.
    m = re.search(r"([A-Za-z][A-Za-z\-]*)$", text)
    # [설명] 없으면
    if not m:
        # [설명] 그대로 돌려줘요.
        return text
    # [설명] tok : 맨 끝 단어 (잘렸을지도 모르는 조각)
    tok = m.group(1)
    # [설명] head : 그 단어 앞까지의 주소
    head = text[: len(text) - len(tok)]
    # [설명] tu : 조각을 대문자로 바꾼 것 (비교하기 쉽게)
    tu = tok.upper()
    # 2) 국가명
    # [설명] 조각이 KO, KOR, KORE 이거나, 'of K' 처럼 K 하나만 남은 경우라면
    if tu in ("KO", "KOR", "KORE") or (tu == "K" and re.search(r"(of|South|,)\s*$", head, re.I)):
        # [설명] Korea 로 정해요. (대문자·소문자는 원래대로)
        new = match_case(tok, "Korea") if len(tok) > 1 else ("KOREA" if head.isupper() else "Korea")
        # [설명] 수정내역에 적고
        notes.append(f"잘린 단어 '{tok}'→'{new}'")
        # [설명] 고쳐서 돌려줘요.
        return head + new
    # 3) 기대 행정구역 (…Songpa-gu, SE → Seoul)
    # [설명] "다음에 올 차례인 말"(예: 송파구 다음엔 Seoul)을 하나씩 꺼내서 봐요.
    for e in expected_next(head):
        # [설명] 그 말이 조각으로 시작하고, 조각이 2글자 이상이면 (쉼표 바로 뒤라면 1글자도 괜찮아요)
        if e.upper().startswith(tu) and e.upper() != tu and (len(tok) >= 2 or head.rstrip().endswith(",")):
            # [설명] 그 말로 정해요.
            new = match_case(tok, e) if len(tok) > 1 else (e.upper() if head.isupper() else e)
            # [설명] 수정내역에 적고
            notes.append(f"잘린 단어 '{tok}'→'{new}'")
            # [설명] 고쳐서 돌려줘요.
            return head + new
    # [설명] 조각이 3글자보다 짧으면
    if len(tok) < 3:
        # [설명] 맞히기 어려우니 고치지 않아요.
        return text
    # 4) 도로명 / 행정구역 사전에서 앞부분 일치
    # [설명] 구·시 이름 목록을 'Songpa-gu' 같은 모양으로 만들어요.
    dist_vocab = [d[0] + "-" + d[1] for d in DISTRICTS]
    # [설명] 조각이 이미 완성된 도로 이름이나 구 이름과 똑같으면
    if any(v.upper() == tu for v in rd.vocab_roads + dist_vocab):
        # [설명] 잘린 게 아니니 그대로 돌려줘요.
        return text
    # [설명] 조각을 접어요.
    ft = fold(tok)

    # [설명] [작은 함수] 어떤 이름(v)이 이 조각으로 '시작하는지' 알아봐요.
    def pref(v, loose):
        # [설명] 이름을 대문자로 바꾸고
        vu = v.upper()
        # [설명] 이름이 조각으로 시작하면 ('-' 는 빼고 봐도 돼요)
        if vu.startswith(tu) or vu.replace("-", "").startswith(tu.replace("-", "")):
            # [설명] "맞아요" 하고 답해요.
            return True
        # [설명] 느슨하게 봐도 될 때는: 조각이 4글자 이상이고, 접은 소리가 같게 시작하고,
        return (loose and len(tok) >= 4 and fold(v).startswith(ft)
                # [설명] 앞 3글자가 똑같으면 "맞아요" 하고 답해요.
                and vu.replace("-", "")[:3] == tu.replace("-", "")[:3])
    # [설명] 조각으로 시작하는 도로 이름 후보를 모아요.
    road_c = [v for v in rd.vocab_roads if pref(v, True)]
    # [설명] 조각으로 시작하는 구·시 이름 후보를 모아요.
    dist_c = [v for v in dist_vocab if pref(v, False)]
    # [설명] 우편번호가 있으면
    if post5:
        # [설명] 우편번호를 숫자로 바꾸고
        pn = int(post5)
        # [설명] '구 이름 → 우편번호 범위' 표를 만들어서
        rng = {d[0] + "-" + d[1]: d[2] for d in DISTRICTS}
        # [설명] 이 우편번호가 들어가는 구·시만 후보로 남겨요.
        dist_c = [v for v in dist_c if any(lo <= pn <= hi for lo, hi in rng[v])]
    # [설명] 주소의 다른 곳에 이미 도로 이름이 있다면
    if has_road_elsewhere:
        # 이미 도로명이 있는 주소라면, 도로명 접미사가 잘린 경우(-dae, -r …)나
        # 도로 이름 자체가 잘린 경우만 도로명으로 복원
        # [설명] [작은 함수] 이 도로 후보가 '정말 잘린 도로 이름인지' 알아봐요.
        def truncated_road(v):
            # [설명] 도로 표기에서 이름 부분만 떼어 내요. (Songpa-daero → Songpa)
            namepart = re.split(r"-(?=[a-z]+$)", v)[0]
            # [설명] 조각이 이미 이름 전체와 같으면
            if fold(namepart) == ft:
                # [설명] 도로가 잘린 게 아니에요.
                return False
            # [설명] 조각이 '-da' 처럼 끝말이 잘린 모양이거나, 이름보다 짧으면 잘린 거예요.
            return bool(re.search(r"-[A-Za-z]{0,4}$", tok)) or len(tok) < len(namepart)
        # [설명] 정말 잘린 것 같은 도로 후보만 남겨요.
        road_c = [v for v in road_c if truncated_road(v)]
    # [설명] 아직 고른 답이 없어요.
    choice = None
    # 구/시 복원은 쉼표·도로명·번지·다른 행정구역 뒤에 올 때만 (건물명 끝 'BUND' 등 제외)
    # [설명] 구·시 후보는 그 앞이 쉼표·숫자·도로 이름·동네 이름으로 끝날 때만 믿어요. (건물 이름 끝의 'BUND' 를 분당으로 잘못 고치지 않으려고요)
    if dist_c and not re.search(r"(,|\d|-ro|-gil|-daero|ro|gil|-gu|-si|-do|dong)\s*$", head, re.I):
        # [설명] 그렇지 않으면 구·시 후보를 버려요.
        dist_c = []
    # [설명] 후보가 통틀어 딱 하나뿐이면
    if len(road_c) + len(dist_c) == 1:
        # [설명] 그것을 답으로 골라요.
        choice = (road_c + dist_c)[0]
    # [설명] 구·시 후보가 하나이고, 도로 이름이 이미 다른 곳에 있으면
    elif dist_c and len(dist_c) == 1 and has_road_elsewhere:
        # [설명] 구·시를 답으로 골라요.
        choice = dist_c[0]
    # [설명] 도로 후보가 여러 개 남았다면
    elif road_c:
        # [설명] 이 우편번호 동네에서 실제로 본 도로들을 꺼내서
        here = rd.roads_at(post5) if post5 else set()
        # [설명] 후보 중에 그 동네에 있는 것만 남겨요.
        rc = [v for v in road_c if v in here]
        # [설명] 딱 하나만 남으면 (그리고 구·시 후보와 헷갈릴 일이 없으면)
        if len(rc) == 1 and not (dist_c and has_road_elsewhere):
            # [설명] 그것을 답으로 골라요.
            choice = rc[0]
    # [설명] 답을 골랐으면
    if choice:
        # [설명] 대문자·소문자를 맞추고
        new = match_case(tok, choice)
        # [설명] 수정내역에 적고
        notes.append(f"잘린 단어 '{tok}'→'{new}'")
        # [설명] 고쳐서 돌려줘요.
        return head + new
    # [설명] 후보는 있는데 하나로 고르지 못했다면 (그리고 다른 곳에 도로 이름도 없으면)
    if (road_c or dist_c) and not has_road_elsewhere:
        # [설명] 함부로 고치지 않고 "끝이 잘렸는데 되살리지 못했어요" 하고 확인필요에 적어요.
        flags_holder.append(f"주소 끝 잘림 '{tok}…' (복원 불가)")
    # [설명] 고치지 않고 그대로 돌려줘요.
    return text


# [설명] [함수] 깨진 글자 '?' 와 엑셀 오류값 '#NAME?' 를 정리해요.
def clean_qmarks(text, notes, flags):
    # [설명] 값이 아예 없으면
    if text is None:
        # [설명] "없음" 을 그대로 돌려줘요.
        return None
    # [설명] 칸에 '#NAME?' 만 적혀 있으면 (엑셀이 주소를 계산식으로 착각해서 주소가 사라진 경우예요)
    if text.strip().upper() == "#NAME?":
        # [설명] 수정내역에 "오류값을 지웠어요" 라고 적고
        notes.append("엑셀 오류값 '#NAME?' 삭제")
        # [설명] 확인필요에 "원래 주소가 사라졌어요" 라고 적고
        flags.append("원본 주소가 '#NAME?' 오류로 소실됨")
        # [설명] 빈 글자를 돌려줘요.
        return ""
    # [설명] '?' 가 하나도 없으면
    if "?" not in text:
        # [설명] 고칠 게 없으니 그대로 돌려줘요.
        return text
    # [설명] '??' 처럼 물음표가 연달아 있거나, 글자의 30% 이상이 물음표라면 (한글이 통째로 깨진 경우예요)
    if re.search(r"\?{2,}", text) or text.count("?") >= 0.3 * max(len(text.replace(" ", "")), 1):
        # [설명] 되살릴 수 없으니 확인필요에 적고
        flags.append("한글 깨짐('???') - 원문 확인 필요")
        # [설명] 고치지 않고 그대로 돌려줘요.
        return text
    # [설명] 물음표가 하나씩만 있으면 띄어쓰기로 바꿔요.
    new = re.sub(r"\s*\?\s*", " ", text).strip()
    # [설명] 쉼표 앞에 생긴 띄어쓰기는 지워요.
    new = re.sub(r"\s+,", ",", new)
    # [설명] 바뀐 게 있으면
    if new != text:
        # [설명] 수정내역에 적어요.
        notes.append("깨진 문자 '?' → 공백")
    # [설명] 고친 글자를 돌려줘요.
    return new


# [설명] [함수] 띄어쓰기와 쉼표를 깔끔하게 정리해요.
def tidy(text):
    # [설명] 띄어쓰기가 여러 개 붙어 있으면 하나로 줄여요.
    text = re.sub(r"[ \t ]+", " ", text)
    # [설명] 쉼표 앞의 띄어쓰기를 지워요.
    text = re.sub(r"\s+,", ",", text)
    # [설명] 쉼표가 여러 개 붙어 있으면 하나로 줄여요.
    text = re.sub(r",{2,}", ",", text)
    # [설명] 맨 앞과 맨 뒤의 띄어쓰기를 떼고 돌려줘요.
    return text.strip()


# [설명] [가장 중요한 함수] 송장 한 건의 주소1, 주소2, 우편번호를 받아서 고쳐요. 위에서 만든 함수들을 차례로 불러요.
def fix_address(a1, a2, post_raw, rd):
    """return: new_a1, new_a2, new_post(int or original), notes[list], flags[list]"""
    # [설명] notes : "무엇을 고쳤는지" 적는 목록,  flags : "사람이 확인해야 하는 이유" 를 적는 목록
    notes, flags = [], []
    # [설명] 주소1이 비어 있으면 빈 글자로, 아니면 글자로 바꿔요.
    a1 = "" if a1 is None else str(a1)
    # [설명] 주소2도 똑같이 해요.
    a2 = "" if a2 is None else str(a2)
    # [설명] 원래 주소를 o1, o2 에 따로 보관해요. 나중에 "바뀌었나?" 비교하려고요.
    o1, o2 = a1, a2
    # [설명] 우편번호를 5글자로 맞춰요.
    post5 = pad_post(post_raw)

    # [설명] 주소1의 깨진 글자를 정리해요.
    a1 = clean_qmarks(a1, notes, flags)
    # [설명] 주소2의 깨진 글자를 정리해요.
    a2 = clean_qmarks(a2, notes, flags)

    # [설명] 사람이 수동사전의 '치환'에 적어 둔 (틀린 말, 바꿀 말) 짝을 하나씩 꺼내요.
    for src, dst in rd.manual.get("치환", {}).items():
        # [설명] 대문자·소문자를 가리지 않고 찾는 틀을 만들어요.
        for_i = re.compile(re.escape(src), re.I)
        # [설명] 주소1이나 주소2에 그 틀린 말이 있으면
        if for_i.search(a1) or for_i.search(a2):
            # [설명] 바꿀 말로 바꿔요.
            a1, a2 = for_i.sub(dst, a1), for_i.sub(dst, a2)
            # [설명] 수정내역에 적어요.
            notes.append(f"수동사전 '{src}'→'{dst}'")

    # 35자 경계에서 끊긴 단어 연결 (A1 끝 + A2 앞)
    # [설명] 주소1의 '맨 끝 단어'를 찾아요.
    mt = re.search(r"([A-Za-z][A-Za-z\-]*)$", a1)
    # [설명] 주소2의 '맨 앞 단어'를 찾아요.
    mh = re.match(r"(-?\s*[A-Za-z][A-Za-z\-]*)", a2)
    # [설명] 둘 다 있으면 (단어가 두 칸에 걸쳐 끊겼을 수 있어요)
    if mt and mh:
        # [설명] 끝 단어(t)와 앞 단어(h)를 꺼내요.
        t, h = mt.group(1), mh.group(1)
        # [설명] 둘을 붙여 봐요. 예) GAN + GNAM-GU → GANGNAM-GU
        merged = t + h.replace(" ", "")
        # [설명] 구·시 이름표를 만들어요. (대문자 표기 → 올바른 표기)
        dist_vocab = {(d[0] + "-" + d[1]).upper(): d[0] + "-" + d[1] for d in DISTRICTS}
        # [설명] Seoul 과 Korea 도 이름표에 넣어요.
        dist_vocab.update({"SEOUL": "Seoul", "KOREA": "Korea"})
        # [설명] 도로 이름표도 만들어요.
        road_up = {v.upper(): v for v in rd.vocab_roads}
        # [설명] 붙인 말이 구·시 이름표나 도로 이름표에 있는지 찾아봐요.
        hit = dist_vocab.get(merged.upper()) or road_up.get(merged.upper())
        # [설명] 없으면
        if not hit:
            # [설명] 붙인 말이 '이름 + daero/ro/gil' 모양인지 보고
            mm = re.match(r"([A-Za-z]{3,})-?(daero|ro|gil)$", merged, re.I)
            # [설명] 그 모양이면
            if mm:
                # [설명] 도로명 사전에서 '딱 맞는 것만' 찾아봐요.
                c, how = rd.lookup(mm.group(1), mm.group(2), post5, loose=False)
                # [설명] 찾은 것을 답으로 써요.
                hit = c
        # [설명] 정말 끊긴 것 같은지 확인해요: 주소1이 34글자 이상이거나, '-' 로 끝나거나 시작하거나,
        split_like = (len(o1) >= TRUNC_LEN or t.endswith("-") or h.startswith("-")
                      # [설명] 주소2 앞 단어가 ro, gil, gu 같은 끝말뿐일 때예요.
                      or re.fullmatch(r"(dae-?ro|ro|gil|gu|si|do)", h, re.I))
        # [설명] 붙인 말이 아는 이름이고, 정말 끊긴 것 같으면
        if hit and split_like:
            # [설명] 대문자·소문자를 맞추고
            new = match_case(merged, hit)
            # [설명] 주소1의 끝 단어를 완성된 말로 바꾸고
            a1 = a1[: len(a1) - len(t)] + new
            # [설명] 주소2에서는 앞 단어를 떼어 내요.
            a2 = a2[len(h):].lstrip(" ,")
            # [설명] 수정내역에 적어요.
            notes.append(f"끊긴 단어 연결 '{t}'+'{h}'→'{new}'")

    # 도로명 / 행정구역 표기 수정
    # [설명] 주소1의 도로 이름을 먼저 고치고, 그다음 구·시·도 이름을 고쳐요.
    a1 = district_fix(road_fix(a1, post5, rd, notes), notes)
    # [설명] 주소2도 똑같이 해요.
    a2 = district_fix(road_fix(a2, post5, rd, notes), notes)

    # [설명] [작은 함수] 글자 안에 도로 이름(ro, gil, daero, beon-gil)이 있는지 알아봐요.
    def has_road(s):
        return bool(re.search(r"([A-Za-z]{2,}|\d)\s*-?\s*(daero|ro|gil|lo)(?![A-Za-z])|beon-?gil", s, re.I))
    # [설명] 주소1이나 주소2 어느 한쪽에라도 도로 이름이 있는지 적어 둬요.
    road_any = has_road(a1) or has_road(a2)

    # 잘린 끝 단어 복원
    # [설명] 잘린 주소를 살펴보다 나온 '확인필요'를 잠깐 담아 둘 목록이에요.
    tail_flags = []
    # [설명] 원래 주소1이 34글자 이상이면 (잘렸을 수 있으니)
    if len(o1) >= TRUNC_LEN:
        # [설명] 주소1의 끝 단어를 되살려 봐요.
        a1 = complete_tail(a1, post5, rd, notes, tail_flags, road_any)
    # [설명] 원래 주소2가 34글자 이상이면
    if len(o2) >= TRUNC_LEN:
        # [설명] 주소2의 끝 단어를 되살려 봐요.
        a2 = complete_tail(a2, post5, rd, notes, tail_flags, road_any)
    # [설명] 거기서 나온 확인필요를 전체 목록에 더해요.
    flags += tail_flags

    # A1 이 도로명으로 끝났는데 A2 가 '-daero', 'ro' 등 접미사 조각으로 시작하면 조각 삭제
    # [설명] 주소2가 '-daero', 'ro' 같은 끝말 조각으로 시작하는지 봐요.
    ms = re.match(r"\s*-?\s*(dae-?ro|ro|gil)(?![A-Za-z])[\s,]*", a2, re.I)
    # [설명] 그런데 주소1이 이미 그 끝말로 끝나 있으면 (같은 말이 두 번 적힌 거예요)
    if ms and re.search(r"[A-Za-z]{3,}-?" + ms.group(1).replace("-", "-?") + r"$", a1.rstrip(" ,"), re.I):
        # [설명] 수정내역에 적고
        notes.append(f"Address2 앞 중복 조각 '{a2[:ms.end()].strip()}' 삭제")
        # [설명] 주소2에서 그 조각을 떼어 내요.
        a2 = a2[ms.end():]

    # 공백 정리 (수정된 칸만)
    # [설명] 주소1이 원래와 달라졌으면
    if a1 != o1:
        # [설명] 띄어쓰기와 쉼표를 정리해요.
        a1 = tidy(a1)
    # [설명] 주소2가 원래와 달라졌으면
    if a2 != o2:
        # [설명] 띄어쓰기와 쉼표를 정리해요.
        a2 = tidy(a2)

    # Address1 / Address2 중복
    # [설명] 주소1에서 글자와 숫자만 남겨요. (띄어쓰기·쉼표가 달라도 같은 내용인지 비교하려고요)
    n1 = re.sub(r"[^a-z0-9]", "", a1.lower())
    # [설명] 주소2도 똑같이 해요.
    n2 = re.sub(r"[^a-z0-9]", "", a2.lower())
    # [설명] 둘 다 내용이 있고, 8글자 이상이고, 한쪽이 다른 쪽으로 시작하면 (같은 주소가 두 칸에 적힌 거예요)
    if n1 and n2 and len(min(n1, n2, key=len)) >= 8 and (n1.startswith(n2) or n2.startswith(n1)):
        # [설명] 주소2 쪽이 더 길면
        if len(n2) > len(n1):
            # [설명] 더 긴 것을 주소1에 넣고
            a1 = a2
        # [설명] 주소2는 비워요.
        a2 = ""
        # [설명] 수정내역에 적어요.
        notes.append("Address1/Address2 중복 제거")

    # 우편번호
    # [설명] 주소1과 주소2를 이어 붙여 전체 주소를 만들어요.
    full = (a1 + " " + a2).strip()
    # [설명] 새 우편번호는 일단 원래 값으로 해 둬요.
    new_post = post_raw
    # [설명] 주소 글자 속에 적혀 있는 우편번호를 전부 찾아요.
    tp = [p for p in TEXT_POST_RE.findall(full)]
    # [설명] 우편번호 칸이 비었거나 이상한 값이었다면
    if not post5:
        # [설명] 주소 글자 속에 우편번호가 있으면
        if tp:
            # [설명] 그 숫자를 우편번호로 쓰고
            new_post = int(tp[-1])
            # [설명] 5글자 우편번호도 채워요.
            post5 = tp[-1]
            # [설명] 수정내역에 적어요.
            notes.append(f"우편번호 보완(주소에서 추출) {tp[-1]}")
        # [설명] 주소 속에도 없으면
        else:
            # [설명] 확인필요에 "우편번호가 없어요" 라고 적어요.
            flags.append("우편번호 없음/오류")
    # [설명] 우편번호 칸은 멀쩡한데, 주소 속에 적힌 우편번호가 전부 그것과 다르면
    elif tp and all(p != post5 for p in tp):
        # [설명] 확인필요에 "두 우편번호가 달라요" 라고 적어요.
        flags.append(f"우편번호 불일치(주소 내 {tp[-1]}, 우편번호칸 {post5})")

    # 검사
    # [설명] 주소에 숫자가 하나도 없으면
    if full and not re.search(r"\d", full):
        # [설명] 확인필요에 "번지가 없어요" 라고 적어요.
        flags.append("번지(숫자) 없음")
    # [설명] '동·리·읍·면' 으로 끝나는 옛날식(지번) 주소인지 찾아봐요.
    jibun = re.search(r"[A-Za-z]{3,}[- ]?(dong|ri|eup|myeon)(?![A-Za-z])", full, re.I)
    # [설명] 도로 이름도 없고 지번도 없으면 (건물 이름만 적힌 주소예요. 한글이 깨진 건은 빼고요)
    if full and not has_road(full) and not jibun and "한글 깨짐('???') - 원문 확인 필요" not in flags:
        # [설명] 확인필요에 "도로명/지번이 없어요" 라고 적어요.
        flags.append("도로명/지번 없음")
    # [설명] 주소가 아예 비어 있으면
    if not full:
        # [설명] 확인필요에 "주소가 없어요" 라고 적어요.
        flags.append("주소 없음")
    # [설명] 우편번호가 있으면 주소와 서로 맞는지 검사해요.
    if post5:
        # [설명] 우편번호를 숫자로 바꿔요.
        pn = int(post5)
        # [설명] 동네 이름표를 하나씩 꺼내요. (이름, 종류, 우편번호 범위)
        for nm, kind, ranges in DISTRICTS:
            # [설명] 경기도는 범위가 너무 넓어서
            if nm == "Gyeonggi":
                # [설명] 건너뛰어요.
                continue
            # [설명] 주소에 그 구·시 이름이 적혀 있으면
            if re.search(nm + r"-" + kind + r"\b", full, re.I):
                # [설명] 그런데 우편번호가 그 구·시의 범위 안에 들어 있지 않으면
                if not any(lo <= pn <= hi for lo, hi in ranges):
                    # [설명] 확인필요에 "구와 우편번호가 서로 안 맞아요" 라고 적고
                    flags.append(f"구/시({nm}-{kind}) ↔ 우편번호({post5}) 불일치")
                    # [설명] 더 찾지 않고 멈춰요.
                    break
        # [설명] 주소 속의 도로 이름을 하나씩 꺼내요.
        for m in ROAD_RE.finditer(full):
            # [설명] 도로의 열쇠를 만들고
            k = fold(m.group("name") + m.group("suf"))
            # [설명] 오타 목록에 있으면 올바른 열쇠로 바꿔요.
            k = rd.alias.get(k, k)
            # [설명] 그 도로가 지금까지 나온 '우편번호 앞 3자리' 기록을 꺼내요.
            c = rd.road_post3.get(k)
            # [설명] 30번 이상 본 진짜 도로일 때만 믿고 검사해요.
            if c and sum(c.values()) >= 30 and k in rd.main_keys:
                # [설명] 이 우편번호 앞 3자리에서는 그 도로를 한 번도 본 적이 없으면
                if c.get(post5[:3], 0) == 0:
                    # [설명] 확인필요에 "도로와 우편번호가 안 맞는 것 같아요" 라고 적고
                    flags.append(f"도로명({rd.canon[k]}) ↔ 우편번호({post5}) 불일치 의심")
                    # [설명] 더 찾지 않고 멈춰요.
                    break

    # [설명] (주소1, 주소2, 우편번호, 고친 내용, 확인할 이유) 다섯 가지를 돌려줘요. 안 바뀐 칸은 원래 값을 그대로 돌려줘요.
    return (a1 if a1 != o1 else o1), (a2 if a2 != o2 else o2), new_post, notes, flags


# ---------------------------------------------------------------------------
# 파일 처리
# ---------------------------------------------------------------------------
# [설명] [함수] 원본 폴더에서 '정제할 엑셀 파일' 목록을 만들어요.
def list_inputs(src_dir, out_dir):
    # [설명] 빈 목록으로 시작해요.
    files = []
    # [설명] 폴더 안의 .xlsx 파일을 이름 순서대로 하나씩 꺼내요.
    for f in sorted(glob.glob(os.path.join(src_dir, "*.xlsx"))):
        # [설명] 폴더 길은 떼고 파일 이름만 꺼내요.
        b = os.path.basename(f)
        # [설명] 엑셀이 잠깐 만든 임시 파일(~$ 로 시작), 결과보고 파일(_ 로 시작), 이미 정제된 파일(_정제.xlsx)이면
        if b.startswith("~$") or b.startswith("_") or OUT_SUFFIX + ".xlsx" in b:
            # [설명] 건너뛰어요.
            continue
        # [설명] 아니면 목록에 넣어요.
        files.append(f)
    # [설명] 다 모은 목록을 돌려줘요.
    return files


# [설명] [함수] 엑셀 파일을 '읽기만 하는 빠른 방식'으로 열어서, 한 줄씩 (주소1, 주소2, 우편번호) 를 건네줘요.
def iter_rows_ro(path):
    # [설명] 엑셀 파일을 열어요. (read_only = 읽기만, data_only = 계산식 말고 결과 값만)
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    # [설명] try ... finally 는 "중간에 무슨 일이 생겨도 마지막 정리는 꼭 해" 라는 뜻이에요.
    try:
        # [설명] 엑셀 안의 시트(쪽)를 하나씩 꺼내요.
        for ws in wb.worksheets:
            # [설명] 빈 줄을 몇 번 만났는지 세는 숫자예요.
            empty = 0
            # [설명] 시트의 줄을 위에서부터 하나씩 꺼내요.
            for r in ws.iter_rows(min_row=1, values_only=True):
                # [설명] 칸 수가 모자란 줄이면
                if len(r) < COL_POST:
                    # [설명] 건너뛰어요.
                    continue
                # [설명] 송장번호(AWB) 칸의 값을 꺼내요. (파이썬은 0부터 세기 때문에 1을 빼요)
                awb = r[COL_AWB - 1]
                # [설명] 송장번호가 비어 있으면
                if awb is None:
                    # [설명] 빈 줄 수를 1 늘리고
                    empty += 1
                    # [설명] 빈 줄이 200번 넘게 이어지면
                    if empty > 200:
                        # [설명] 이 시트는 끝났다고 보고 멈춰요.
                        break
                    # [설명] 아니면 다음 줄로 가요.
                    continue
                # [설명] 내용이 있는 줄을 만났으니 빈 줄 수를 0으로 되돌려요.
                empty = 0
                # [설명] 송장번호 칸에 'AWB' 라고 적힌 줄은 제목 줄이니
                if str(awb).strip().upper() == "AWB":
                    # [설명] 건너뛰어요.
                    continue
                # [설명] yield 는 "하나 건네주고 잠깐 기다릴게" 예요. 주소1, 주소2, 우편번호를 건네줘요.
                yield r[COL_A1 - 1], r[COL_A2 - 1], r[COL_POST - 1]
    # [설명] 끝나면 (또는 중간에 문제가 생겨도)
    finally:
        # [설명] 엑셀 파일을 꼭 닫아요.
        wb.close()


# [설명] [함수] 엑셀 파일들을 읽으면서 도로명 사전을 배워요.
def learn(files, rd):
    # [설명] 파일을 하나씩 꺼내요. i 는 1부터 세는 순서 번호예요.
    for i, f in enumerate(files, 1):
        # [설명] 일단 해 봐요.
        try:
            # [설명] 파일에서 한 줄씩 주소1, 주소2, 우편번호를 받아서
            for a1, a2, p in iter_rows_ro(f):
                # [설명] 우편번호를 5글자로 맞추고
                p5 = pad_post(p)
                # [설명] 주소1에서 도로 이름을 배우고
                rd.observe(str(a1) if a1 else "", p5)
                # [설명] 주소2에서도 도로 이름을 배워요.
                rd.observe(str(a2) if a2 else "", p5)
        # [설명] 파일을 읽다가 문제가 생기면
        except Exception as e:
            # [설명] "이 파일은 건너뛰었어요" 하고 화면에 알려 줘요. (프로그램은 멈추지 않아요)
            print(f"  [학습 건너뜀] {os.path.basename(f)}: {e}")
        # [설명] 20개마다, 그리고 마지막 파일에서
        if i % 20 == 0 or i == len(files):
            # [설명] "여기까지 배웠어요" 하고 화면에 알려 줘요.
            print(f"  사전 학습 {i}/{len(files)}")
    # [설명] 다 배웠으면 사전을 정리해서 완성해요.
    rd.build()


# [설명] [함수] 엑셀 파일 하나를 열어서 주소를 고치고, 색을 칠하고, 새 이름으로 저장해요.
def process_file(path, out_path, rd, cache, report_rows, flag_rows):
    # [설명] 엑셀 파일을 열어요. (이번에는 고쳐야 하니까 보통 방식으로 열어요)
    wb = openpyxl.load_workbook(path)
    # [설명] 파일 이름만 꺼내요.
    base = os.path.basename(path)
    # [설명] 이 파일의 성적표예요. 전체 몇 건, 고친 것 몇 건, 확인필요 몇 건인지 세요.
    stat = {"file": base, "rows": 0, "changed": 0, "flagged": 0}
    # [설명] 굵은 글씨예요.
    bold = Font(bold=True)
    # [설명] 시트를 하나씩 꺼내요.
    for ws in wb.worksheets:
        # [설명] 빈 줄 세는 숫자예요.
        empty = 0
        # [설명] 시트의 마지막 줄 번호예요.
        maxr = ws.max_row
        # [설명] 제목 줄을 봤는지 적어 두는 표시예요. (아직 못 봤어요)
        header_seen = False
        # [설명] 1번 줄부터 마지막 줄까지 하나씩 봐요.
        for r in range(1, maxr + 1):
            # [설명] 송장번호 칸의 값을 꺼내요.
            awb = ws.cell(r, COL_AWB).value
            # [설명] 비어 있으면
            if awb is None:
                # [설명] 빈 줄 수를 1 늘리고
                empty += 1
                # [설명] 200번 넘게 이어지면
                if empty > 200:
                    # [설명] 이 시트는 끝이에요.
                    break
                # [설명] 아니면 다음 줄로 가요.
                continue
            # [설명] 빈 줄 수를 0으로 되돌려요.
            empty = 0
            # [설명] 제목 줄(AWB 라고 적힌 줄)이면
            if str(awb).strip().upper() == "AWB":
                # [설명] M칸에는 '수정내역', N칸에는 '확인필요' 라는 제목을 쓸 거예요.
                for col, title in ((COL_NOTE, "수정내역"), (COL_FLAG, "확인필요")):
                    # [설명] 그 칸에 제목을 쓰고
                    c = ws.cell(r, col, title)
                    # [설명] 굵게 하고
                    c.font = bold
                    # [설명] 하늘색을 칠해요.
                    c.fill = FILL_HEAD
                # [설명] "제목 줄을 봤어요" 하고 표시하고
                header_seen = True
                # [설명] 다음 줄로 가요.
                continue
            # [설명] 이 줄의 주소1 칸, 주소2 칸, 우편번호 칸을 꺼내요.
            c1, c2, cp = ws.cell(r, COL_A1), ws.cell(r, COL_A2), ws.cell(r, COL_POST)
            # [설명] 세 값을 묶어서 '이름표(key)'로 써요.
            key = (c1.value, c2.value, cp.value)
            # [설명] 이 주소를 처음 본 거라면
            if key not in cache:
                # [설명] 고쳐서 그 결과를 기억해 둬요. (똑같은 주소가 또 나오면 다시 계산하지 않고 꺼내 써요)
                cache[key] = fix_address(c1.value, c2.value, cp.value, rd)
            # [설명] 기억해 둔 결과를 꺼내요: 새 주소1, 새 주소2, 새 우편번호, 고친 내용, 확인할 이유
            n1, n2, np_, notes, flags = cache[key]
            # [설명] 확인할 이유 목록을 복사해요. (기억해 둔 원본이 바뀌지 않게요)
            flags = list(flags)
            # [설명] 시트 이름이 'NA' 면 (배송 구역이 정해지지 않은 시트예요)
            if ws.title.strip().upper() == "NA":
                # [설명] 확인필요에 그 이유를 더해요.
                flags.append("배송구역 미지정(NA 시트)")
            # [설명] 전체 건수를 1 늘려요.
            stat["rows"] += 1
            # [설명] 원래 주소1, 주소2를 따로 적어 둬요.
            o1, o2 = c1.value, c2.value
            # [설명] "이 줄이 바뀌었나?" 표시예요. (아직은 아니에요)
            changed = False
            # [설명] 주소1이 원래와 달라졌으면
            if (o1 or "") != (n1 or ""):
                # [설명] 새 주소를 칸에 넣고
                c1.value = n1 if n1 else None
                # [설명] 노란색을 칠하고
                c1.fill = FILL_CHANGED
                # [설명] "바뀌었어요" 하고 표시해요.
                changed = True
            # [설명] 주소2가 원래와 달라졌으면
            if (o2 or "") != (n2 or ""):
                # [설명] 새 주소를 칸에 넣고
                c2.value = n2 if n2 else None
                # [설명] 노란색을 칠하고
                c2.fill = FILL_CHANGED
                # [설명] "바뀌었어요" 하고 표시해요.
                changed = True
            # [설명] 우편번호가 원래와 달라졌으면
            if np_ != cp.value:
                # [설명] 새 우편번호를 칸에 넣고
                cp.value = np_
                # [설명] 노란색을 칠하고
                cp.fill = FILL_CHANGED
                # [설명] "바뀌었어요" 하고 표시해요.
                changed = True
            # [설명] 우편번호가 숫자면
            if isinstance(cp.value, (int, float)):
                # [설명] 엑셀에서 05854 처럼 5자리로 보이게 해요. (맨 앞의 0 이 사라지지 않게요)
                cp.number_format = "00000"
            # [설명] 이 줄의 M칸(수정내역)과 N칸(확인필요)을 꺼내요.
            nc, fc = ws.cell(r, COL_NOTE), ws.cell(r, COL_FLAG)
            # [설명] 두 칸의 배경을 흰색으로 칠하고
            nc.fill, fc.fill = FILL_WHITE, FILL_WHITE          # 원본 양식의 배경색 덮기
            # [설명] 글씨를 검은색으로 해요.
            nc.font, fc.font = FONT_BLACK, FONT_BLACK
            # [설명] 고친 내용이 있고 정말 바뀌었으면
            if notes and changed:
                # [설명] M칸에 고친 내용을 ' / ' 로 이어서 적어요. (dict.fromkeys 는 똑같은 말이 두 번 나오지 않게 해 줘요)
                nc.value = " / ".join(dict.fromkeys(notes))
                # [설명] 고친 건수를 1 늘려요.
                stat["changed"] += 1
            # [설명] 확인할 이유가 있으면
            if flags:
                # [설명] N칸에 이유를 ' / ' 로 이어서 적고
                fc.value = " / ".join(dict.fromkeys(flags))
                # [설명] 빨간색을 칠하고
                fc.fill = FILL_FLAG
                # [설명] 확인필요 건수를 1 늘려요.
                stat["flagged"] += 1
            # [설명] 고쳤거나 확인이 필요한 줄이면, 결과보고서에 넣을 한 줄을 만들어요.
            if (notes and changed) or flags:
                # [설명] 파일 이름, 시트 이름, 줄 번호, 송장번호,
                row = [base, ws.title, r, awb,
                       # [설명] 1번째·7번째·8번째 칸의 값(배송구역, 받는 사람, 전화번호),
                       ws.cell(r, 1).value, ws.cell(r, 7).value, ws.cell(r, 8).value,
                       # [설명] 원래 주소1·2, 고친 주소1·2, 우편번호,
                       o1, o2, c1.value, c2.value, pad_post(cp.value) or cp.value,
                       # [설명] 고친 내용,
                       " / ".join(dict.fromkeys(notes)) if changed else "",
                       # [설명] 확인할 이유를 한 줄로 모아요.
                       " / ".join(dict.fromkeys(flags))]
                # [설명] 고친 줄이면
                if notes and changed:
                    # [설명] '수정내역' 목록에 넣고
                    report_rows.append(row)
                # [설명] 확인이 필요한 줄이면
                if flags:
                    # [설명] '확인필요' 목록에 넣어요.
                    flag_rows.append(row)
        # [설명] 이 시트에 제목 줄이 있었으면
        if header_seen:
            # [설명] M칸의 너비를 넓히고
            ws.column_dimensions["M"].width = 48
            # [설명] N칸의 너비도 넓혀요. (글이 잘 보이게요)
            ws.column_dimensions["N"].width = 42
    # [설명] 새 이름으로 저장해요. (원본 파일은 그대로 남아요)
    wb.save(out_path)
    # [설명] 성적표를 돌려줘요.
    return stat


# [설명] [함수] 모든 파일의 결과를 모아 '결과보고서' 엑셀을 만들어요.
def write_report(path, stats, report_rows, flag_rows):
    # [설명] 새 엑셀 파일을 만들어요.
    wb = openpyxl.Workbook()
    # [설명] 첫 번째 시트를 꺼내서
    ws = wb.active
    # [설명] 이름을 '요약' 으로 지어요.
    ws.title = "요약"
    # [설명] 첫 줄에 제목과 지금 날짜·시간을 적어요.
    ws.append(["주소 정제 결과", datetime.datetime.now().strftime("%Y-%m-%d %H:%M")])
    # [설명] 한 줄 비워요.
    ws.append([])
    # [설명] 표의 제목 줄을 적어요.
    ws.append(["파일명", "전체 건수", "자동 수정 건수", "확인필요 건수"])
    # [설명] 파일별 성적표를 하나씩 꺼내서
    for s in stats:
        # [설명] 한 줄씩 적어요.
        ws.append([s["file"], s["rows"], s["changed"], s["flagged"]])
    # [설명] 마지막에 전체 합계를 적어요.
    ws.append(["합계", sum(s["rows"] for s in stats), sum(s["changed"] for s in stats),
               sum(s["flagged"] for s in stats)])
    # [설명] A칸의 너비를 넓히고
    ws.column_dimensions["A"].width = 50
    # [설명] B, C, D 칸은
    for c in "BCD":
        # [설명] 조금만 넓혀요.
        ws.column_dimensions[c].width = 14
    # [설명] 3번째 줄(표의 제목 줄)의 칸을 하나씩 꺼내서
    for c in ws[3]:
        # [설명] 굵게 하고
        c.font = Font(bold=True)
        # [설명] 하늘색을 칠해요.
        c.fill = FILL_HEAD

    # 유형별 집계
    # [설명] 확인필요 이유별로 개수를 셀 상자예요.
    cnt = collections.Counter()
    # [설명] 확인필요 줄을 하나씩 꺼내서
    for r in flag_rows:
        # [설명] 그 줄의 이유들을 ' / ' 로 나눠 하나씩 꺼내고
        for f in r[13].split(" / "):
            # [설명] 괄호 안과 따옴표 안의 내용(주소마다 다른 부분)을 지운 다음, 같은 종류끼리 +1 해요.
            cnt[re.sub(r"\s+", " ", re.sub(r"\([^)]*\)|'[^']*'", "", f)).strip()] += 1
    # [설명] 한 줄 비우고
    ws.append([])
    # [설명] 작은 표의 제목을 적고
    ws.append(["확인필요 유형", "건수"])
    # [설명] 굵게 해요.
    ws.cell(ws.max_row, 1).font = Font(bold=True)
    # [설명] 많이 나온 이유부터 차례로
    for k, v in cnt.most_common():
        # [설명] (이유, 건수) 를 적어요.
        ws.append([k, v])

    # [설명] '확인필요' 시트와 '수정내역' 시트에 쓸 제목 줄이에요.
    heads = ["파일", "시트", "행", "AWB", "Route", "Receiver", "Tel",
             "원본 Address1", "원본 Address2", "수정 Address1", "수정 Address2",
             "Postcode", "수정내역", "확인필요 사유"]
    # [설명] 칸마다 너비를 얼마로 할지 적어 둔 숫자예요.
    widths = [34, 6, 6, 12, 7, 18, 14, 36, 36, 38, 38, 8, 50, 45]
    # [설명] 시트 두 개를 차례로 만들어요: '확인필요' 와 '수정내역'
    for title, rows in (("확인필요", flag_rows), ("수정내역", report_rows)):
        # [설명] 새 시트를 만들고
        w = wb.create_sheet(title)
        # [설명] 제목 줄을 적어요.
        w.append(heads)
        # [설명] 제목 줄의 칸을 하나씩 꺼내서
        for c in w[1]:
            # [설명] 굵게 하고
            c.font = Font(bold=True)
            # [설명] 하늘색을 칠해요.
            c.fill = FILL_HEAD
        # [설명] 모아 둔 줄을 하나씩 꺼내서
        for row in rows:
            # [설명] 시트에 적어요.
            w.append(row)
        # [설명] 칸마다 순서 번호와 너비를 꺼내서
        for i, wd in enumerate(widths, 1):
            # [설명] 그 칸의 너비를 맞춰요. (get_column_letter 는 1→A, 2→B 로 바꿔 줘요)
            w.column_dimensions[openpyxl.utils.get_column_letter(i)].width = wd
        # [설명] 아래로 내려도 제목 줄이 보이게 첫 줄을 고정해요.
        w.freeze_panes = "A2"
        # [설명] 제목 줄에 '골라 보기(필터)' 단추를 달아요.
        w.auto_filter.ref = f"A1:N{max(len(rows) + 1, 2)}"
    # [설명] 결과보고서를 저장해요.
    wb.save(path)


# [설명] [시작 함수] 프로그램이 실행되면 여기서부터 차례로 일을 해요.
def main():
    # [설명] 이 프로그램 파일이 들어 있는 폴더의 위치를 알아내요.
    here = os.path.dirname(os.path.abspath(__file__))
    # [설명] 실행할 때 원본 폴더를 적어 줬으면 그 폴더를, 안 적었으면 프로그램 폴더의 한 칸 위 폴더를 써요.
    src = os.path.abspath(sys.argv[1]) if len(sys.argv) > 1 else os.path.dirname(here)
    # [설명] 결과 폴더도 적어 줬으면 그걸, 안 적었으면 원본 폴더 안의 '주소정제_결과' 를 써요.
    out = os.path.abspath(sys.argv[2]) if len(sys.argv) > 2 else os.path.join(src, "주소정제_결과")
    # [설명] 결과 폴더가 없으면 만들어요. (이미 있어도 괜찮아요)
    os.makedirs(out, exist_ok=True)
    # [설명] 정제할 파일 목록을 만들어요.
    files = list_inputs(src, out)
    # [설명] 폴더 위치와 파일 개수를 화면에 보여 줘요.
    print(f"원본 폴더: {src}\n결과 폴더: {out}\n대상 파일: {len(files)}개")
    # [설명] 파일이 하나도 없으면
    if not files:
        # [설명] 여기서 끝내요.
        return

    # [설명] 빈 도로명 사전을 하나 만들어요.
    rd = RoadDict()
    # [설명] '도로명사전.json'(프로그램이 배운 내용) 파일의 위치예요.
    dict_path = os.path.join(here, "도로명사전.json")
    # [설명] '수동사전.json'(사람이 적어 준 내용) 파일의 위치예요.
    manual_path = os.path.join(here, "수동사전.json")
    # [설명] 수동사전 파일이 있으면
    if os.path.exists(manual_path):
        # [설명] 파일을 열어서 (with 는 다 쓰면 알아서 닫아 줘요)
        with open(manual_path, encoding="utf-8") as fp:
            # [설명] 내용을 읽어 사전에 넣어요.
            rd.manual.update(json.load(fp))
    # [설명] 전에 저장해 둔 도로명사전이 있고, "처음부터 다시 배워(--relearn)" 라고 하지 않았으면
    if os.path.exists(dict_path) and "--relearn" not in sys.argv:
        # [설명] 파일을 열어서
        with open(dict_path, encoding="utf-8") as fp:
            # [설명] 저장된 내용을 읽고
            saved = json.load(fp)
        # [설명] 사전에 채워 넣고
        rd.load_json(saved.get("dict", {}))
        # [설명] 전에 이미 배운 파일 이름들을 꺼내요.
        learned = set(saved.get("files", []))
    # [설명] 그게 아니면
    else:
        # [설명] 배운 파일이 하나도 없는 걸로 해요.
        learned = set()
    # [설명] 아직 배우지 않은 새 파일만 골라요.
    new_files = [f for f in files if os.path.basename(f) not in learned]
    # [설명] "1단계: 사전을 배워요" 하고 화면에 알려 주고
    print(f"[1/3] 도로명 사전 학습 (새 파일 {len(new_files)}개)")
    # [설명] 새 파일들로 사전을 배워요.
    learn(new_files, rd)
    # [설명] 일단 해 봐요.
    try:
        # [설명] 도로명사전 파일을 쓰기용으로 열어서
        with open(dict_path, "w", encoding="utf-8") as fp:
            # [설명] 배운 파일 이름들과
            json.dump({"files": sorted(learned | {os.path.basename(f) for f in files}),
                       # [설명] 사전 내용을 저장해요. (ensure_ascii=False 는 한글을 한글 그대로 저장하라는 뜻이에요)
                       "dict": rd.to_json()}, fp, ensure_ascii=False)
    # [설명] 저장하지 못했으면 (예: 파일이 잠겨 있을 때)
    except OSError:
        # [설명] 그냥 넘어가요. (pass = 아무것도 안 하기)
        pass
    # [설명] 도로를 몇 개 배웠는지 화면에 보여 줘요.
    print(f"  도로명 {len(rd.canon)}개 학습됨")

    # [설명] "2단계: 주소를 고쳐서 저장해요" 하고 알려 줘요.
    print("[2/3] 주소 수정 및 저장")
    # [설명] 빈 상자 네 개를 준비해요: 기억 상자, 성적표 목록, 수정내역 목록, 확인필요 목록
    cache, stats, report_rows, flag_rows = {}, [], [], []
    # [설명] 파일을 하나씩 꺼내요.
    for i, f in enumerate(files, 1):
        # [설명] 파일 이름에서 '.xlsx' 를 뗀 이름을 만들어요.
        stem = os.path.splitext(os.path.basename(f))[0].rstrip()
        # [설명] 결과 파일의 위치를 만들어요. 예) 결과폴더/배송자료_정제.xlsx
        op = os.path.join(out, stem + OUT_SUFFIX + ".xlsx")
        # [설명] 일단 해 봐요.
        try:
            # [설명] 파일 하나를 고쳐서 저장하고 성적표를 받아요.
            s = process_file(f, op, rd, cache, report_rows, flag_rows)
            # [설명] 성적표를 목록에 넣고
            stats.append(s)
            # [설명] "몇 번째 파일, 고친 것 몇 건, 확인필요 몇 건" 을 화면에 보여 줘요.
            print(f"  ({i}/{len(files)}) {stem}  수정 {s['changed']} / 확인필요 {s['flagged']}")
        # [설명] 문제가 생기면
        except Exception as e:
            # [설명] "이 파일은 실패했어요" 하고 알려 주고 다음 파일로 넘어가요.
            print(f"  ({i}/{len(files)}) [실패] {stem}: {e}")

    # [설명] "3단계: 결과보고서를 만들어요" 하고 알려 줘요.
    print("[3/3] 결과보고서 작성")
    # [설명] 결과보고서 엑셀을 만들어요.
    write_report(os.path.join(out, "_주소정제_결과보고.xlsx"), stats, report_rows, flag_rows)
    # [설명] 전체 합계를 화면에 보여 주고 끝내요.
    print(f"완료: 자동 수정 {sum(s['changed'] for s in stats)}건, 확인필요 {sum(s['flagged'] for s in stats)}건")


# [설명] "이 파일을 직접 실행했을 때만" 아래를 해요. (다른 프로그램이 이 파일을 도구로 불러 쓸 때는 안 해요)
if __name__ == "__main__":
    # [설명] 시작 함수(main)를 불러요. 여기서 프로그램이 진짜로 출발해요!
    main()
