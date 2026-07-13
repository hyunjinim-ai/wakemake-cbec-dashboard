# -*- coding: utf-8 -*-
"""
상품명 매핑 테이블 (중국어 원문 ↔ 한글 상품명 ↔ 내부 SKU 코드)
================================================================
플랫폼마다 상품명이 중국어로, 표기 방식이 다르게(약어·띄어쓰기·신구버전) 나오므로
동일 상품을 SKU로 정확히 매칭하기 위한 단일 소스.  build.py 와 data.json 이 이 표를 사용한다.

- status "확정": 자동 합산 대상.
- status "검토": 매칭이 애매 → 자동 합산 금지, 대시보드 "미매칭 검토" 목록에 노출.
- SKU 코드는 잠정값(WM-*/CG-*). 실제 상품마스터 코드가 확보되면 이 표(또는 구글시트 `상품매핑` 탭)에서 교체.
"""

# ── 정규화: 별칭 매칭 시 표기 차이를 흡수 ──────────────────────────────
import re
_STRIP_MARKS = ["新品", "AD", "구형", "임박", "기획", "특가", "한정", "限量", "预售",
                "赠品", "정품", "본품", "신규", "데일리"]
def norm_cn(s):
    s = str(s or "")
    s = re.sub(r"[\s　]", "", s)                          # 공백(전각 포함) 제거
    s = re.sub(r"[【】\[\]（）()《》<>·・,，、\-_/|~!\.．]", "", s)   # 괄호·구분자 제거
    for mk in _STRIP_MARKS:
        s = s.replace(mk, "")
    return s.lower()

# ── 매핑 테이블 (canonical SKU 1행 = 동일 상품) ────────────────────────
#   tmallIds : 티몰 상품ID(신뢰 앵커, 여러 변형ID 허용)
#   douyin/xhs : 각 플랫폼 중문 상품명/카테고리 별칭(여러 표기 허용)
MAPPING = [
 {"sku":"WM-EYE-SOFT","ko":"소프트 블러링 아이팔레트","brand":"웨이크메이크","status":"확정",
  "tmallIds":["802291666650"], "douyin":[], "xhs":[],
  "note":"티몰 히어로 SKU(6월 매출 51%)"},
 {"sku":"WM-EYE-16","ko":"16색 아이섀도우 팔레트(데일리)","brand":"웨이크메이크","status":"확정",
  "tmallIds":["930390492188"], "douyin":[], "xhs":[]},
 {"sku":"WM-EYE-MULTI","ko":"6색 멀티 팔레트","brand":"웨이크메이크","status":"확정",
  "tmallIds":["915599957043"], "douyin":[], "xhs":[]},
 {"sku":"WM-CUSHION","ko":"워터 글로우 코팅 쿠션","brand":"웨이크메이크","status":"확정",
  "tmallIds":["930412381224","979107752691","834612380188","1010846509316","937100575838","916056804484"],
  "douyin":[], "xhs":["粉糖气垫","气垫种草","气垫测评"],
  "note":"샤오홍슈 聚光 '气垫(쿠션)' 계열 합산"},
 {"sku":"WM-CUSHION-GRAY","ko":"매트 그레이 쿠션","brand":"웨이크메이크","status":"확정",
  "tmallIds":["955630790429"], "douyin":[], "xhs":[]},
 {"sku":"WM-CONCEAL","ko":"디파이닝 커버 컨실러","brand":"웨이크메이크","status":"확정",
  "tmallIds":["991053980918","943692099726","1025567503245"], "douyin":[], "xhs":[]},
 {"sku":"WM-FND","ko":"디파이닝 커버 파운데이션","brand":"웨이크메이크","status":"확정",
  "tmallIds":["955175851181"], "douyin":[], "xhs":[]},
 {"sku":"WM-PACT","ko":"퓨어 소이 파우더 팩트","brand":"웨이크메이크","status":"확정",
  "tmallIds":["802234879735"], "douyin":[], "xhs":[]},
 {"sku":"WM-TINT-WG","ko":"워터풀 글로우 틴트","brand":"웨이크메이크","status":"확정",
  "tmallIds":["944149219488","970326378569"], "douyin":[], "xhs":[]},
 {"sku":"WM-TINT-LAYER","ko":"워터 블러링 레이어 틴트","brand":"웨이크메이크","status":"확정",
  "tmallIds":["944132243773"], "douyin":[], "xhs":[]},
 {"sku":"WM-TINT-BOLD","ko":"볼드 립 블러 틴트","brand":"웨이크메이크","status":"확정",
  "tmallIds":["1056920657682"], "douyin":[], "xhs":[]},
 {"sku":"WM-TINT-LOVELY","ko":"러비 미광 립 틴트","brand":"웨이크메이크","status":"확정",
  "tmallIds":["923655776577"], "douyin":[], "xhs":[]},
 {"sku":"WM-BALM-STICK","ko":"헬시 글로우 밤 스틱","brand":"웨이크메이크","status":"확정",
  "tmallIds":["956917552339"], "douyin":[], "xhs":[]},
 {"sku":"WM-BALM-SOFT","ko":"소프트 블러링 밤 스틱","brand":"웨이크메이크","status":"확정",
  "tmallIds":["979918030260"], "douyin":[], "xhs":[]},
 {"sku":"WM-SHADE","ko":"믹스 블러링 볼륨 쉐딩","brand":"웨이크메이크","status":"확정",
  "tmallIds":["802325822279"], "douyin":[], "xhs":[]},
 {"sku":"WM-CHEEK","ko":"쉐이킹 블러 치크","brand":"웨이크메이크","status":"확정",
  "tmallIds":["981293620333"], "douyin":[], "xhs":[]},
 {"sku":"WM-MIST","ko":"스테이 픽서 세팅 미스트","brand":"웨이크메이크","status":"확정",
  "tmallIds":["1055872138792"], "douyin":[], "xhs":[]},
 {"sku":"WM-BROW","ko":"눈썹연필·아이브로우","brand":"웨이크메이크","status":"확정",
  "tmallIds":["994394931225"], "douyin":[], "xhs":[]},
 {"sku":"WM-BRUSH","ko":"와이드 파운데이션 브러시","brand":"웨이크메이크","status":"확정",
  "tmallIds":["923641001774"], "douyin":[], "xhs":[]},
 {"sku":"WM-PAL-CONTOUR","ko":"윤곽 형광 조색 팔레트","brand":"웨이크메이크","status":"확정",
  "tmallIds":["802226967628"], "douyin":[], "xhs":[]},
 {"sku":"WM-PAL-CONCEAL","ko":"정형 컨실러 조색 팔레트","brand":"웨이크메이크","status":"확정",
  "tmallIds":["802231575085"], "douyin":[], "xhs":[]},
]

# ── 애매(검토) 별칭: 자동 합산 금지, 검토 목록으로 노출 ────────────────
AMBIGUOUS = [
 {"platform":"샤오홍슈","cn":"十六色眼影","candidates":["WM-EYE-16","WM-EYE-SOFT"],
  "reason":"'十六色(16색)'은 문자상 16색 데일리 팔레트지만, 6월 聚光 최대 집행은 히어로 '소프트 블러링 아이팔레트'일 가능성 — 실제 프로모션 SKU 판별 필요"},
 {"platform":"샤오홍슈","cn":"眼影妆教","candidates":["WM-EYE-16","WM-EYE-SOFT"],
  "reason":"콘텐츠 유형(妆教=메이크업 튜토리얼) 태그 — 특정 아이섀도우 SKU 불명"},
 {"platform":"샤오홍슈","cn":"眼影测评","candidates":["WM-EYE-16","WM-EYE-SOFT"],
  "reason":"콘텐츠 유형(测评=리뷰) 태그 — 특정 SKU 불명"},
 {"platform":"샤오홍슈","cn":"眼影种草","candidates":["WM-EYE-16","WM-EYE-SOFT"],
  "reason":"콘텐츠 유형(种草=제품 소개) 태그 — 특정 SKU 불명"},
 {"platform":"샤오홍슈","cn":"眼影合集","candidates":["WM-EYE-16","WM-EYE-SOFT"],
  "reason":"콘텐츠 유형(合集=모음) 태그 — 특정 SKU 불명"},
]
