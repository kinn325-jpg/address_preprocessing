# -*- coding: utf-8 -*-
"""포트폴리오 시연용 '가짜' 배송 자료를 만듭니다.
실제 배송 자료에는 개인정보가 있어 공개하지 않고,
실제로 자주 나온 오류 유형만 본떠 가상의 수취인/연락처/주소로 만들었습니다.
사용법: python 샘플생성.py  ->  input/샘플_배송자료.xlsx 생성
"""
import os
import openpyxl

HEAD = ["Pcs", "Code", "AWB", "Weight", None, "Company", "Receiver", "Tel",
        "Address1", "Address2", "Postcode", "EDD"]

# (Address1, Address2, Postcode)  — 정상 표기 (사전 학습용)
NORMAL = [
    ("513 Yeongdong-daero", "Gangnam-gu Seoul", "06164"),
    ("521 Teheran-ro", "Gangnam-gu Seoul", "06164"),
    ("300 Olympic-ro", "Songpa-gu Seoul", "05551"),
    ("268 Songpa-daero", "Songpa-gu Seoul", "05719"),
    ("212 Samseong-ro", "Gangnam-gu Seoul", "06284"),
    ("235 Pangyoyeok-ro", "Bundang-gu Seongnam-si Gyeonggi-do", "13494"),
]
# 오류가 섞인 표기 (실제 자료에서 자주 나온 유형을 본뜬 것)
ERRORS = [
    ("Room 1203, 12F, 513 Yeongdong-daer", "Gangnam-gu Seoul", "06164"),   # 35자 잘림
    ("Unit 801, 8F Tower A, 268 Son", "Seoul", "05719"),                   # 도로명 잘림
    ("521 Taeheran-ro, Seoul GAN", "GNAM-GU", "06164"),                     # 단어가 두 칸으로 끊김
    ("212 Samsungro", "Kangnam-ku Seoul", "06284"),                         # 도로명 오타 + 구 표기
    ("268 Sonpa daero", "Songpa-ku Seoul", "05719"),                        # 띄어쓰기 + 구 표기
    ("235 Pangyoyeok-ro, Bundang-gu", "Sungnam city Gyeonggido", "13494"),  # 시/도 표기
    ("300 Olympic-ro?Songpa-gu", "Seoul", "5551"),                          # 깨진 문자 + 우편번호 앞자리 0
    ("513 Yeongdong-daero", "513 Yeongdong-daero Gangnam-gu", "06164"),     # Address1/2 중복
    ("Lotte World Tower 25F", "Songpa-gu Seoul", "05551"),                  # 건물명만 -> 확인필요
    ("#NAME?", "Gangnam-gu Seoul", "06164"),                                # 원본 소실 -> 확인필요
    ("212 Samseong-ro", "Songpa-gu Seoul", "06284"),                        # 구-우편번호 불일치 -> 확인필요
    ("521 Teheran-ro Gangnam-gu 06164", "Seoul", None),                     # 우편번호 칸 비어 있음
]

def rows(data, start):
    out = []
    for i, (a1, a2, p) in enumerate(data):
        n = start + i
        out.append([1, "S", f"AWB{n:06d}", 1.0, None, "샘플상사", f"수취인{n:03d}",
                    "010-0000-0000", a1, a2, p, "2026-01-01"])
    return out

def main():
    here = os.path.dirname(os.path.abspath(__file__))
    os.makedirs(os.path.join(here, "input"), exist_ok=True)
    wb = openpyxl.Workbook()
    # 프로그램은 같은 우편번호에서 자주 쓰인 표기를 기준으로 삼기 때문에(20회 이상),
    # 정상 표기를 충분히 넣어 실제 자료와 비슷한 학습 조건을 만듭니다.
    sheets = {"1팀": NORMAL * 12 + ERRORS[:6], "2팀": NORMAL * 12 + ERRORS[6:], "NA": [NORMAL[0]]}
    first = True
    n = 1
    for name, data in sheets.items():
        ws = wb.active if first else wb.create_sheet()
        first = False
        ws.title = name
        ws.append(["샘플 배송 자료 (가상 데이터)"])
        ws.append([])
        ws.append(HEAD)
        for r in rows(data, n):
            ws.append(r)
        n += len(data)
    path = os.path.join(here, "input", "샘플_배송자료.xlsx")
    wb.save(path)
    print("생성:", path)

if __name__ == "__main__":
    main()
