#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
웨이크메이크 중국 역직구 대시보드 — 데이터 빌드 스크립트 (v2)
================================================================
데이터 소스(구글시트 / 원본 엑셀) → data.json → index.html 자동 생성.
3개 뷰: ① 일자별 매출(브랜드별)  ② 플랫폼 통합(티몰·도우인·샤오홍슈)  ③ 티몰 월별(기존).

사용법:
  python build.py --gsheet [시트URL|ID]      # ★구글시트('중국역직구 플랫폼 판매데이터') → 티몰 월별+일자별 동기화(매월)
  python build.py --seed                    # 확보된 원본에서 실데이터로 data.json 시드(최초 1회)
  python build.py --sheet [sheet.config.json]   # 구글시트(내 스키마 탭별 공개 CSV) → data.json 재생성
  python build.py --render-only             # data.json 만 바꾼 뒤 index.html 재생성
  python build.py --seed-csv                # 현재 data.json → 구글시트 붙여넣기용 스타터 CSV 출력
  python build.py <티몰원본.xlsx>            # (기존) 티몰 6시트 엑셀 → 티몰 파트 재계산
  python build.py --from-raw <index_2.html>  # (기존) RAW 대시보드에서 티몰 파트 복원

환율 1元 = 220원 · 매출은 증정품 제외 결제금액 기준.
"""
import sys, os, json, re, csv, io

try:
    sys.stdout.reconfigure(encoding="utf-8")  # Windows 콘솔(cp949) 한글/기호 출력 보장
except Exception:
    pass

from mapping import MAPPING, AMBIGUOUS, norm_cn

FX = 220
HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(HERE, "template.html")
DATA_JSON = os.path.join(HERE, "data.json")
OUT_HTML = os.path.join(HERE, "index.html")
SHEET_CFG = os.path.join(HERE, "sheet.config.json")
SEED_DIR = os.path.join(HERE, "sheets_seed")
# 도우인 일자별 + 샤오홍슈 聚光 원본(로컬, 시드 전용 · 커밋 안 함)
RAW_MULTI = r"C:\Users\user\Downloads\로우 확인중.xlsx"

# ---------- 상품 한글명 매핑 (없으면 코드 뒤 6자리) ----------
NAME_MAP = {
 "802291666650":"소프트 블러링 아이팔레트","916056804484":"워터 글로우 코팅밤(구형)",
 "979107752691":"워터 글로우 코팅밤AD(쿠션)","802325822279":"믹스 블러링 볼륨 쉐딩",
 "937100575838":"[임박] 워터 글로우 코팅 쿠션","991053980918":"디파이닝 커버 컨실러",
 "944149219488":"워터풀 글로우 틴트","923641001774":"와이드 파운데이션 브러시",
 "981293620333":"쉐이킹 블러 치크","834612380188":"[한정특가] 워터 글로우 코팅 쿠션",
 "943692099726":"[임박] 디파이닝 커버 컨실러","930412381224":"워터 글로우 코팅 쿠션",
 "1025567503245":"[기획] 디파이닝 커버 컨실러","956917552339":"헬시 글로우 밤 스틱",
 "979918030260":"소프트 블러링 밤 스틱","944132243773":"워터 블러링 레이어 틴트",
 "1056920657682":"볼드 립 블러 틴트","1055872138792":"스테이 픽서 세팅 미스트",
 "930390492188":"16색 아이섀도우 팔레트(데일리)",
 "955175851181":"디파이닝 커버 파운데이션(브러시증정)","1010846509316":"워터 글로우 코팅 쿠션(브러시증정)",
 "970326378569":"워터풀 글로우 틴트(신규)","994394931225":"눈썹연필·아이브로우",
 "915599957043":"6색 멀티 팔레트","955630790429":"매트 그레이 쿠션",
 "923655776577":"[특가] 러비 미광 립 틴트","802226967628":"윤곽 형광 조색 팔레트",
 "802231575085":"정형 컨실러 조색 팔레트","802234879735":"퓨어 소이 파우더 팩트",
}
# 유입경로 세부 채널 한글명 (l4 코드 → 한글)
TR = {"搜索":"검색","推荐":"추천 피드","淘宝直播":"타오바오 라이브","天猫榜单":"티몰 랭킹",
      "淘金币及相关场域":"타오진비(포인트)","一淘":"이타오(캐시백)","关键词推广":"키워드 광고(무계)",
      "淘宝客":"타오바오커(제휴)","购物车":"장바구니","我的淘宝":"마이 타오바오",
      "手淘拍立淘":"이미지 검색","淘口令分享":"외부 공유 링크","消息":"메시지 알림",
      "天猫APP":"티몰 APP","逛逛":"타오바오 SNS 피드","自主访问":"직접 방문"}
GMAP = {"经营优势":"자연 유입(무료)","付费推广":"유료 광고","主动回访":"자발적 재방문"}

# 마케팅비 항목 → 플랫폼 귀속(플랫폼 통합 뷰 비용 분해용)
COST_PLATFORM = {
    "티몰 유료광고(내부광고 소모액)": "티몰",
    "도우인 라이브방송(Still 부담 수수료)": "도우인",
    "라이브방송 올리브영 지원분": "도우인",
    "샤오홍씽 CID 효과광고(聚光)": "샤오홍슈",
    "KOL/KOC/SNS 바이럴 마케팅": "샤오홍슈",
}

def monthNum(m):
    mm = re.sub(r"[^0-9]", "", str(m))
    return int(mm) if mm else 0
def kname(pid, cn):
    return NAME_MAP.get(pid) or ("코드 " + str(pid)[-6:])
def is_gift(p):
    cn = str(p.get("cn") or "")
    if ("赠品" in cn) and (p.get("payCNY") or 0) == 0: return True
    if cn.startswith("【赠품】") or cn.startswith("【赠品】"): return True
    return False
def is_new(p):
    return "新品" in str(p.get("cn") or "")
def numf(v):
    try: return float(v)
    except (TypeError, ValueError): return 0.0

# =========================================================================
#  A. 티몰 원본 엑셀 파서 (원본 대시보드 parseWorkbook 의 Python 이식)  — 기존 유지
# =========================================================================
def parse_workbook(path):
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True, read_only=True)
    def sheet(name):
        if name not in wb.sheetnames: return []
        return [list(r) for r in wb[name].iter_rows(values_only=True)]
    return _parse_tmall_sheets(sheet)

def _parse_tmall_sheets(sheet):
    """티몰 시트들(엑셀 or 구글시트 CSV) → 티몰 파트 원자료. sheet(name) -> rows(list of list)."""
    def norm(m):
        return re.sub(r"\s", "", "" if m is None else str(m)).replace("月", "월")
    def num(v):
        if v is None: return 0.0
        if isinstance(v, (int, float)): return float(v)
        s = str(v).replace(",", "").replace("%", "").replace("元", "").replace("₩", "").strip()
        try: return float(s)
        except ValueError: return 0.0
    def pct(v):  # 전환율: 분수(0.05)·퍼센트("5%"/5) 혼용 흡수 → 퍼센트 숫자
        s = str(v).strip()
        if s.endswith("%"): return round(num(s), 2)
        f = num(s)
        return round(f * 100, 2) if 0 < f <= 1.5 else round(f, 2)

    tmall_actual = []
    for row in sheet("티몰글로벌"):
        if len(row) > 8 and str(row[8]).strip() == "티몰":
            tmall_actual = [num(x) for x in row[9:15]]; break

    def idx(header, name):
        try: return header.index(name)
        except ValueError:
            for i, h in enumerate(header):
                if str(h).strip() == name: return i
            return -1
    def cell(r, header, name):
        i = idx(header, name)
        return r[i] if (0 <= i < len(r)) else None

    ps = sheet("웨이크메이크_상품별판매데이터")
    products = {}
    if ps:
        h = ps[0]
        for r in ps[1:]:
            if not r: continue
            m = norm(cell(r, h, "통계 일자"))
            if not m or m == "null": continue
            pid = cell(r, h, "상품 ID")
            if pid is None: continue
            pid = str(int(pid) if isinstance(pid, float) else pid)
            products.setdefault(m, []).append({
                "id": pid, "cn": cell(r, h, "상품명"), "status": cell(r, h, "상품 상태"),
                "uv": int(num(cell(r, h, "상품 방문자 수"))),
                "cart": int(num(cell(r, h, "장바구니 담기 인원 수"))),
                "payCNY": round(num(cell(r, h, "결제 완료 금액"))),
                "conv": pct(cell(r, h, "상품 결제 전환율")),
                "sConv": pct(cell(r, h, "검색 유입 결제 전환율")),
                "sUV": int(num(cell(r, h, "검색 유입 방문자 수"))),
            })

    ts = sheet("웨이크메이크_티몰 점포 유입 현황")
    traffic = {}
    if ts:
        h = ts[0]
        for r in ts[1:]:
            if not r: continue
            m = norm(cell(r, h, "일자"))
            if not m or m == "null": continue
            vis = int(num(cell(r, h, "방문자 수"))); buy = int(num(cell(r, h, "결제 완료 구매자 수")))
            traffic.setdefault(m, []).append({
                "l1": cell(r, h, "1차 유입 경로"), "l2": cell(r, h, "2차 유입 경로"),
                "l3": cell(r, h, "3차 유입 경로"), "l4": cell(r, h, "4차 유입 경로"),
                "visitors": vis, "buyers": buy,
                "conv": round(buy / vis * 10000) / 100 if vis else 0})

    tmall_ad = {}
    ads = sheet("웨이크메이크_티몰 내부 광고 현황")
    if ads:
        h = ads[0]
        for r in ads[1:]:
            if not r: continue
            m = norm(cell(r, h, "통계 일자"))
            if m and m != "null":
                tmall_ad[m] = tmall_ad.get(m, 0) + num(cell(r, h, "광고 소모액(마케팅 비용)"))

    def p2m(p):
        if not p: return None
        s = str(p).split("-")[0].strip()
        mm = re.sub(r"[^0-9].*$", "", s.split(".")[0])
        return (mm + "월") if mm else None
    live_still, live_oy = {}, {}
    ls = sheet("라이브방송 광고 내역")
    if ls:
        h = ls[0]
        for r in ls[1:]:
            if not r: continue
            if str(cell(r, h, "브랜드") or "").lower() != "wakemake": continue
            m = p2m(cell(r, h, "기간 범위"))
            if not m: continue
            live_still[m] = live_still.get(m, 0) + num(cell(r, h, "Still 부담 수수료"))
            live_oy[m] = live_oy.get(m, 0) + num(cell(r, h, "올리브영 지원 내역"))

    xhs_cost, xhs_contrib = {}, {}
    xs = sheet("샤오홍씽-광고")
    if xs:
        h = xs[0]
        for r in xs[1:]:
            if not r: continue
            m = norm(cell(r, h, "월"))
            if not m or m == "null": continue
            c = xhs_cost.setdefault(m, {"ad": 0, "kk": 0})
            c["ad"] += num(cell(r, h, "효과광고(聚光) 비용"))
            c["kk"] += num(cell(r, h, "KOL 마케팅비")) + num(cell(r, h, "KOC 마케팅비")) + num(cell(r, h, "SNS 마케팅비"))
            k = xhs_contrib.setdefault(m, {"visit": 0, "newUV": 0, "buy": 0, "gmv": 0, "ad": 0})
            k["visit"] += num(cell(r, h, "매장 방문 UV")); k["newUV"] += num(cell(r, h, "매장 신규 방문자수"))
            k["buy"] += num(cell(r, h, "전체 매장 구매 UV")); k["gmv"] += num(cell(r, h, "전체 매장 GMV(元)"))
            k["ad"] += num(cell(r, h, "효과광고(聚光) 비용"))

    months = sorted(products.keys(), key=monthNum)
    cost = {}
    for m in months:
        cost[m] = []
        if tmall_ad.get(m): cost[m].append({"owner": "스틸", "item": "티몰 유료광고(내부광고 소모액)", "cny": round(tmall_ad[m])})
        if live_still.get(m): cost[m].append({"owner": "스틸", "item": "도우인 라이브방송(Still 부담 수수료)", "cny": round(live_still[m])})
        if xhs_cost.get(m, {}).get("ad"): cost[m].append({"owner": "OY", "item": "샤오홍씽 CID 효과광고(聚光)", "cny": round(xhs_cost[m]["ad"])})
        if xhs_cost.get(m, {}).get("kk"): cost[m].append({"owner": "OY", "item": "KOL/KOC/SNS 바이럴 마케팅", "cny": round(xhs_cost[m]["kk"])})
        if live_oy.get(m): cost[m].append({"owner": "OY", "item": "라이브방송 올리브영 지원분", "cny": round(live_oy[m])})

    monthly = []
    for i, m in enumerate(months):
        rows = products.get(m, [])
        monthly.append({"month": m, "salesCNY": sum(p["payCNY"] for p in rows),
                        "targetKRW": round(tmall_actual[i]) if i < len(tmall_actual) else 0,
                        "uv": sum(p["uv"] for p in rows), "pv": 0})
    xhs = {}
    for m in months:
        c = xhs_contrib.get(m)
        if not c: continue
        total_uv = next((x["uv"] for x in monthly if x["month"] == m), 0)
        xhs[m] = {"adKRW": round(c["ad"] * FX), "storeVisitUV": round(c["visit"]), "newUV": round(c["newUV"]),
                  "buyUV": round(c["buy"]), "gmvKRW": round(c["gmv"] * FX),
                  "shareOfTotalUV": round(c["visit"] / total_uv * 1000) / 10 if total_uv else 0,
                  "roas": round(c["gmv"] / c["ad"] * 10) / 10 if c["ad"] else 0}
    return {"months": months, "monthly": monthly, "cost": cost, "products": products, "traffic": traffic, "xhs": xhs}

# =========================================================================
#  A-2. 구글시트 판독기 (탭명=엑셀 시트명) + 티몰 일자별  (신규)
# =========================================================================
def _sid(url_or_id):
    m = re.search(r"/spreadsheets/d/([A-Za-z0-9_\-]+)", str(url_or_id))
    return m.group(1) if m else str(url_or_id).strip()

def gsheet_reader(url_or_id):
    """구글시트(공개) → sheet(name)->rows. gviz CSV(탭명 지정)로 탭을 읽어 캐시."""
    import urllib.request, urllib.parse
    sid = _sid(url_or_id); cache = {}
    def sheet(name):
        if name in cache: return cache[name]
        url = "https://docs.google.com/spreadsheets/d/%s/gviz/tq?tqx=out:csv&sheet=%s" % (sid, urllib.parse.quote(name))
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            raw = urllib.request.urlopen(req, timeout=30).read().decode("utf-8-sig", "replace")
            rows = list(csv.reader(io.StringIO(raw)))
        except Exception:
            rows = []
        cache[name] = rows
        return rows
    return sheet

def parse_tmall_daily(sheet, sheet_name="웨이크메이크_티몰 일자별매출"):
    """티몰 일자별 매출 탭 → 브랜드별 일자 series (현재 시트엔 WAKEMAKE海外旗舰店 = 웨이크메이크만)."""
    rows = sheet(sheet_name)
    if not rows: return {}
    h = rows[0]
    def ci(*names):
        for n in names:
            if n in h: return h.index(n)
        return -1
    iDate, iStore = ci("통계 일자", "일자"), ci("상점명", "상점", "스토어")
    iPay, iBuy, iOrd = ci("결제 완료 금액"), ci("결제 완료 구매자 수"), ci("결제 완료 하위 주문 건수", "주문 수량")
    iUV = ci("상품 방문자 수(UV)", "방문자 수")
    def num(v):
        s = str(v).replace(",", "").replace("元", "").strip()
        try: return float(s)
        except ValueError: return 0.0
    out = {}
    for r in rows[1:]:
        if iDate < 0 or iDate >= len(r): continue
        d = str(r[iDate]).strip()[:10]
        if not re.match(r"\d{4}-\d{2}-\d{2}", d): continue
        store = str(r[iStore]) if 0 <= iStore < len(r) else ""
        brand = "컬러그램" if ("COLORGRAM" in store.upper() or "colorgram" in store) else "웨이크메이크"
        pay = num(r[iPay]) if 0 <= iPay < len(r) else 0
        out.setdefault(brand, []).append({
            "date": d, "salesKRW": round(pay * FX),
            "orders": int(num(r[iOrd])) if 0 <= iOrd < len(r) else 0,
            "buyers": int(num(r[iBuy])) if 0 <= iBuy < len(r) else 0,
            "uv": int(num(r[iUV])) if 0 <= iUV < len(r) else 0,
        })
    for b in out: out[b].sort(key=lambda x: x["date"])
    return out

# =========================================================================
#  B. 기존 대시보드(RAW)에서 티몰 파트 시드 — 엑셀 없이 검증용  — 기존 유지
# =========================================================================
def seed_from_raw(html_path):
    txt = open(html_path, "r", encoding="utf-8").read()
    m = re.search(r"const RAW\s*=\s*`(.+?)`", txt, re.S)
    if not m: raise SystemExit("index_2.html 에서 const RAW 를 찾지 못했습니다")
    d = json.loads(m.group(1))
    def unpack(rows, keys): return [dict(zip(keys, r)) for r in rows]
    products = {mo: unpack(rows, d["pkeys"]) for mo, rows in d["products"].items()}
    traffic = {mo: unpack(rows, d["tkeys"]) for mo, rows in d["traffic"].items()}
    return {"months": d["months"], "monthly": d["monthly"], "cost": d["cost"],
            "products": products, "traffic": traffic, "xhs": d.get("xhs", {})}

# =========================================================================
#  C. 티몰 파트 계산 (series + byMonth)  — 기존 유지
# =========================================================================
def compute_tmall(u):
    months = u["months"]
    monthly_by = {x["month"]: x for x in u["monthly"]}
    series = []
    for m in months:
        rows = [p for p in u["products"].get(m, []) if not is_gift(p)]
        sales = sum(p["payCNY"] for p in rows) * FX
        uv = sum(p["uv"] for p in rows)
        cs = u["cost"].get(m, [])
        costK = sum(c["cny"] for c in cs) * FX
        oy = sum(c["cny"] for c in cs if c["owner"] == "OY") * FX
        st = sum(c["cny"] for c in cs if c["owner"] == "스틸") * FX
        xh = u["xhs"].get(m, {})
        mo = monthly_by.get(m, {})
        series.append({
            "month": m, "salesKRW": round(sales), "targetKRW": round(mo.get("targetKRW", 0)),
            "achv": round(sales / mo["targetKRW"] * 1000) / 10 if mo.get("targetKRW") else 0,
            "uv": uv, "pv": mo.get("pv", 0) or 0,
            "costKRW": round(costK), "oyKRW": round(oy), "stKRW": round(st),
            "roas": round(sales / costK * 100) / 100 if costK else 0,
            "xhsAdKRW": round(xh.get("adKRW", 0)), "xhsGmvKRW": round(xh.get("gmvKRW", 0)),
            "xhsRoas": xh.get("roas", 0), "xhsShare": xh.get("shareOfTotalUV", 0), "xhsUV": xh.get("storeVisitUV", 0),
        })

    by_month = {}
    for i, m in enumerate(months):
        rows = [p for p in u["products"].get(m, []) if not is_gift(p)]
        rows.sort(key=lambda p: p.get("payCNY") or 0, reverse=True)
        tot = sum(p["payCNY"] for p in rows)
        products = [{
            "name": kname(p["id"], p.get("cn")), "id": str(p.get("id") or ""),
            "status": "온라인" if p.get("status") == "当前在线" else "품절/내림",
            "new": is_new(p), "uv": p.get("uv") or 0,
            "payKRW": round((p.get("payCNY") or 0) * FX), "payCNY": p.get("payCNY") or 0,
            "conv": p.get("conv") or 0, "share": round((p["payCNY"] / tot * 1000)) / 10 if tot else 0,
        } for p in rows]

        cur = u["cost"].get(m, []); prev = u["cost"].get(months[i - 1], []) if i > 0 else []
        def find(rows_, o, it):
            return next((c["cny"] * FX for c in rows_ if c["owner"] == o and c["item"] == it), 0)
        keys = []
        for c in cur + prev:
            k = (c["owner"], c["item"])
            if k not in keys: keys.append(k)
        cost_items = [{"owner": o, "item": it, "cur": round(find(cur, o, it)),
                       "prev": round(find(prev, o, it)), "diff": round(find(cur, o, it) - find(prev, o, it))}
                      for o, it in keys]
        cost_items.sort(key=lambda x: x["cur"], reverse=True)

        tr = u["traffic"].get(m, [])
        groups = []
        for g in ["经营优势", "付费推广", "主动回访"]:
            row = next((r for r in tr if r["l1"] == g and r["l2"] == "汇总"), None)
            if row: groups.append({"name": GMAP[g], "visitors": row["visitors"], "buyers": row["buyers"], "conv": row["conv"]})
        subs_raw = [r for r in tr if r["l4"] in TR and r["l4"] != "汇总"]
        seen = {}
        for r in subs_raw:
            nm = TR[r["l4"]]
            if nm not in seen or r["visitors"] > seen[nm]["visitors"]:
                seen[nm] = {"name": nm, "visitors": r["visitors"], "buyers": r["buyers"], "conv": r["conv"]}
        subs = sorted(seen.values(), key=lambda x: x["visitors"], reverse=True)[:10]

        by_month[m] = {"products": products, "productTotalKRW": round(tot * FX),
                       "costItems": cost_items, "traffic": {"groups": groups, "subs": subs}}
    return {"months": months, "series": series, "byMonth": by_month}

# =========================================================================
#  D. 도우인 · 샤오홍슈 원본 파서  (신규)
# =========================================================================
DOUYIN_STORE = {
    "WAKEMAKE海外旗舰店": ("웨이크메이크", "cross"),
    "WAKE MAKE唯可魅旗舰店": ("웨이크메이크", "domestic"),
    "COLORGRAM海外旗舰店": ("컬러그램", "cross"),
    "COLORGRAM旗舰店": ("컬러그램", "domestic"),
}
def parse_douyin(path):
    """도우인 일자별(스토어x브랜드) 원본 → 정규화 rows."""
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True, read_only=True)
    if "도우인 웨메 매출" not in wb.sheetnames:
        wb.close(); return []
    ws = wb["도우인 웨메 매출"]
    rows = list(ws.iter_rows(values_only=True)); h = rows[0]
    def ci(name): return h.index(name) if name in h else -1
    iDate, iStore = ci("日期"), ci("载体类型")
    iGmv, iPay, iOrd, iBuy, iAov = ci("成交金额"), ci("用户支付金额"), ci("成交订单数"), ci("成交人数"), ci("客单价")
    iExp, iClk = ci("商品曝光人数"), ci("商品点击人数")
    out = []
    for r in rows[1:]:
        if r[iDate] is None: continue
        store = str(r[iStore])
        brand, stype = DOUYIN_STORE.get(store, (store, "?"))
        out.append({
            "date": str(r[iDate])[:10], "store": store, "brand": brand, "stype": stype,
            "gmvCNY": round(numf(r[iGmv]), 2), "payCNY": round(numf(r[iPay]), 2),
            "orders": int(numf(r[iOrd])), "buyers": int(numf(r[iBuy])), "aov": round(numf(r[iAov]), 2),
            "expUV": int(numf(r[iExp])) if iExp >= 0 else 0, "clickUV": int(numf(r[iClk])) if iClk >= 0 else 0,
        })
    wb.close()
    return out

def parse_xhs_juguang(path):
    """샤오홍슈 聚光(효과광고) 일자·카테고리별 집행 원본 → 월/카테고리 집계."""
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True, read_only=True)
    if "쥐광 내역" not in wb.sheetnames:
        wb.close(); return {"byMonth": {}, "catByMonth": {}}
    ws = wb["쥐광 내역"]
    rows = list(ws.iter_rows(values_only=True)); h = rows[0]
    def ci(name): return h.index(name) if name in h else -1
    iM, iCost, iExp, iClk, iCat = ci("月份"), ci("消费"), ci("曝光/展现"), ci("点击量"), ci("商品品类")
    byMonth, catByMonth = {}, {}
    for r in rows[1:]:
        if iM < 0 or r[iM] is None: continue
        m = str(int(numf(r[iM]))) if numf(r[iM]) else str(r[iM])
        cost, exp, clk = numf(r[iCost]), int(numf(r[iExp])), int(numf(r[iClk]))
        b = byMonth.setdefault(m, {"cost": 0.0, "exp": 0, "clk": 0})
        b["cost"] += cost; b["exp"] += exp; b["clk"] += clk
        cat = str(r[iCat]) if iCat >= 0 and r[iCat] is not None else "(미지정)"
        cm = catByMonth.setdefault(m, {})
        c = cm.setdefault(cat, {"cost": 0.0, "exp": 0, "clk": 0})
        c["cost"] += cost; c["exp"] += exp; c["clk"] += clk
    wb.close()
    return {"byMonth": byMonth, "catByMonth": catByMonth}

# =========================================================================
#  E. 매핑 인덱스 + 미매칭 판정  (신규)
# =========================================================================
def build_map_index():
    """정규화 별칭 → SKU 인덱스, 티몰ID → SKU 인덱스, 애매 별칭 set."""
    by_alias = {}   # (platform, norm) -> sku
    by_tmall = {}   # tmall id -> sku
    for m in MAPPING:
        for pid in m.get("tmallIds", []):
            by_tmall[str(pid)] = m["sku"]
        for plat, key in (("도우인", "douyin"), ("샤오홍슈", "xhs")):
            for a in m.get(key, []):
                by_alias[(plat, norm_cn(a))] = m["sku"]
    ambiguous = {(a["platform"], norm_cn(a["cn"])): a for a in AMBIGUOUS}
    sku_of = {m["sku"]: m for m in MAPPING}
    return {"alias": by_alias, "tmall": by_tmall, "ambiguous": ambiguous, "sku": sku_of}

def match_platform_name(platform, cn, idx):
    """플랫폼 상품명/카테고리 → (sku or None, ambiguous_info or None)."""
    key = (platform, norm_cn(cn))
    if key in idx["alias"]:
        return idx["alias"][key], None
    if key in idx["ambiguous"]:
        return None, idx["ambiguous"][key]
    return None, None

# =========================================================================
#  F. 일자별(페이지①) 계산  (신규)
# =========================================================================
def compute_daily(douyin, month="6월", store_type="cross"):
    prefix = "2026-%02d" % monthNum(month)
    brands = ["웨이크메이크", "컬러그램"]
    series = {b: [] for b in brands}
    for d in douyin:
        if d["stype"] != store_type: continue
        if not d["date"].startswith(prefix): continue
        if d["brand"] not in series: continue
        series[d["brand"]].append({
            "date": d["date"], "salesKRW": round(d["payCNY"] * FX), "gmvKRW": round(d["gmvCNY"] * FX),
            "orders": d["orders"], "buyers": d["buyers"], "aov": round(d["aov"] * FX),
        })
    for b in brands:
        series[b].sort(key=lambda x: x["date"])
    dates = sorted({x["date"] for b in brands for x in series[b]})
    return {
        "month": month, "brands": brands,
        "platformLabel": "도우인(抖音) 역직구 · 海外旗舰店",
        "metricNote": "일 결제금액(用户支付金额) · 증정 제외 · 환율 220원",
        "coverage": (dates[0] + " ~ " + dates[-1]) if dates else "",
        "dates": dates, "series": series,
        "prev": {b: None for b in brands},   # 5월 일자별 미확보 → 시트 입력 시 채움
    }

def daily_status_sample():
    """페이지① 상단 '7월 현황' 샘플 — 새 실시간 시트(1B_7…) 연결 전 레이아웃용.
    브랜드×플랫폼 목표·MTD실적·UV(전월 동기 대비). 실데이터 연결 시 이 함수를 파서로 교체."""
    asof, dim = 15, 31
    plats = ["티몰", "도우인", "샤오홍슈"]
    WV = [0.9, 1.0, 1.1, 0.8, 0.72, 1.2, 1.35, 1.05, 0.95, 1.05, 1.18, 0.85, 0.78, 1.12, 1.25]  # 15일 가중(주말 등락)
    def series(mtd):
        s = sum(WV); return [round(mtd * w / s) for w in WV]
    seed = {
        "웨이크메이크": {"target": {"티몰": 155000000, "도우인": 52000000, "샤오홍슈": 21000000},
                    "mtd": {"티몰": 71800000, "도우인": 27300000, "샤오홍슈": 8400000},
                    "uv": {"cur": 63200, "prev": 61050}},
        "컬러그램": {"target": {"티몰": 42000000, "도우인": 34000000, "샤오홍슈": 9000000},
                  "mtd": {"티몰": 16400000, "도우인": 17600000, "샤오홍슈": 2600000},
                  "uv": {"cur": 38900, "prev": 35200}},
    }
    data = {}
    for b, v in seed.items():
        per = {p: series(v["mtd"][p]) for p in plats}
        daily = []
        for i in range(asof):
            day = {"date": "2026-07-%02d" % (i + 1)}
            for p in plats: day[p] = per[p][i]
            day["total"] = sum(day[p] for p in plats)
            daily.append(day)
        data[b] = {"target": v["target"], "actual": v["mtd"], "uv": v["uv"], "daily": daily}
    return {"sample": True, "asOf": "2026-07-%02d" % asof, "month": "7월", "prevMonth": "6월",
            "dayOfMonth": asof, "daysInMonth": dim, "brands": list(seed.keys()), "platforms": plats, "data": data}

def daily_struct(series_by_brand, label, metricNote, month=None):
    """임의 브랜드별 일자 series → 페이지① 데이터 구조(월 필터 옵션)."""
    brands = ["웨이크메이크", "컬러그램"]
    series = {b: list(series_by_brand.get(b, [])) for b in brands}
    if month:
        pre = "2026-%02d" % monthNum(month)
        series = {b: [x for x in series[b] if x["date"].startswith(pre)] for b in brands}
    for b in brands: series[b].sort(key=lambda x: x["date"])
    dates = sorted({x["date"] for b in brands for x in series[b]})
    if month:
        mlabel = month
    else:
        mos = sorted({int(d[5:7]) for d in dates})
        mlabel = "%d–%d월" % (mos[0], mos[-1]) if len(mos) > 1 else ("%d월" % mos[0] if mos else "전체")
    return {"month": mlabel, "brands": brands, "platformLabel": label,
            "metricNote": metricNote, "coverage": (dates[0] + " ~ " + dates[-1]) if dates else "",
            "dates": dates, "series": series, "prev": {b: None for b in brands}}

# =========================================================================
#  G. 플랫폼 통합(페이지②) 계산  (신규 · 브랜드 차원)
# =========================================================================
WM_BRAND = "웨이크메이크"

def _brand_products(prodmap, brand, platform, idx):
    """옵션 상품별 입력(시트 도우인_상품별/샤오홍슈_상품별) → 매핑 반영 상품 리스트."""
    out = []
    for p in (prodmap or {}).get(brand, []):
        cn = p.get("name") or p.get("cn") or ""
        sku, amb = match_platform_name(platform, cn, idx)
        tid = idx["tmall"].get(str(p.get("tmallId", "")))
        sku = sku or tid
        out.append({"name": cn, "sku": sku, "ko": idx["sku"][sku]["ko"] if sku else cn,
                    "matched": bool(sku), "payKRW": round(numf(p.get("payKRW") or (numf(p.get("payCNY")) * FX))),
                    "orders": int(numf(p.get("orders"))), "uv": int(numf(p.get("uv")))})
    out.sort(key=lambda x: x["payKRW"], reverse=True)
    return out

def compute_platforms(tmall, douyin, xhs, month="6월", douyin_products=None, xhs_products=None):
    """브랜드별(웨이크메이크·컬러그램) 플랫폼 통합 구조."""
    brands = [WM_BRAND, "컬러그램"]
    byBrand = {b: _platforms_for_brand(tmall, douyin, xhs, month, b,
                                       douyin_products=douyin_products, xhs_products=xhs_products) for b in brands}
    return {"brands": brands, "byBrand": byBrand}

def _platforms_for_brand(tmall, douyin, xhs, month, brand, store_type="cross",
                         douyin_products=None, xhs_products=None):
    idx = build_map_index()
    prefix = "2026-%02d" % monthNum(month)
    S = {s["month"]: s for s in tmall["series"]}
    BM = tmall["byMonth"]
    is_wm = (brand == WM_BRAND)
    s6 = S.get(month, {}) if is_wm else {}
    bm6 = BM.get(month, {}) if is_wm else {}

    tmall_cost_items = bm6.get("costItems", [])
    def plat_cost(plat):
        return sum(c["cur"] for c in tmall_cost_items if COST_PLATFORM.get(c["item"]) == plat)
    tm_cost = plat_cost("티몰")

    # ---- 티몰 ---- (웨이크메이크만 로우데이터 확보)
    if is_wm and bm6:
        tmall_p = {
            "label": "티몰글로벌(天猫国际)", "hasData": True, "hasProductDetail": True,
            "kpi": {"salesKRW": s6.get("salesKRW", 0), "uv": s6.get("uv", 0), "costKRW": tm_cost,
                    "roas": round(s6.get("salesKRW", 0) / tm_cost, 2) if tm_cost else 0,
                    "orders": None, "buyers": None},
            "products": bm6.get("products", []), "productTotalKRW": bm6.get("productTotalKRW", 0),
            "costItems": [c for c in tmall_cost_items if COST_PLATFORM.get(c["item"]) == "티몰"],
            "traffic": bm6.get("traffic", {"groups": [], "subs": []}),
            "note": "상품별 sell-out·유입경로 전체 확보(로우데이터)",
        }
    else:
        tmall_p = {
            "label": "티몰글로벌(天猫国际)", "hasData": False, "hasProductDetail": False,
            "kpi": {"salesKRW": 0, "uv": 0, "costKRW": 0, "roas": 0, "orders": None, "buyers": None},
            "products": [], "productTotalKRW": 0, "costItems": [], "traffic": {"groups": [], "subs": []},
            "note": "티몰 " + brand + " 데이터 미확보 — 시트 `플랫폼_티몰` 입력 대기",
        }

    # ---- 도우인 ---- (두 브랜드 모두 일자별 확보)
    drows = [d for d in douyin if d["stype"] == store_type and d["brand"] == brand and d["date"].startswith(prefix)]
    dsum = {"payKRW": 0, "gmvKRW": 0, "orders": 0, "buyers": 0, "expUV": 0, "clickUV": 0}
    daily = []
    for d in sorted(drows, key=lambda x: x["date"]):
        dsum["payKRW"] += d["payCNY"] * FX; dsum["gmvKRW"] += d["gmvCNY"] * FX
        dsum["orders"] += d["orders"]; dsum["buyers"] += d["buyers"]
        dsum["expUV"] += d["expUV"]; dsum["clickUV"] += d["clickUV"]
        daily.append({"date": d["date"], "salesKRW": round(d["payCNY"] * FX),
                      "orders": d["orders"], "buyers": d["buyers"]})
    dy_cost = plat_cost("도우인")
    dy_sales = round(dsum["payKRW"])
    dprods = _brand_products(douyin_products, brand, "도우인", idx)
    douyin_p = {
        "label": "도우인(抖音) 역직구", "hasData": bool(drows), "hasProductDetail": bool(dprods),
        "kpi": {"salesKRW": dy_sales, "gmvKRW": round(dsum["gmvKRW"]), "costKRW": dy_cost,
                "roas": round(dy_sales / dy_cost, 2) if dy_cost else None,
                "orders": dsum["orders"], "buyers": dsum["buyers"],
                "aov": round(dy_sales / dsum["orders"]) if dsum["orders"] else 0,
                "uv": dsum["expUV"]},
        "products": dprods, "daily": daily,
        "funnel": {"expUV": dsum["expUV"], "clickUV": dsum["clickUV"], "buyers": dsum["buyers"]},
        "note": "스토어 단위 일자 집계" + ("" if dprods else " · 상품별 sell-out 은 시트 `도우인_상품별` 입력 시 표시"),
    }

    # ---- 샤오홍슈 (聚光 집행 + 기여 GMV) ---- (웨이크메이크만 로우 확보)
    xm = str(monthNum(month))
    xj = xhs.get("byMonth", {}).get(xm, {"cost": 0, "exp": 0, "clk": 0}) if is_wm else {"cost": 0, "exp": 0, "clk": 0}
    xcat = xhs.get("catByMonth", {}).get(xm, {}) if is_wm else {}
    xhs_ad_juguang = round(xj["cost"] * FX)
    xhs_cost_all = plat_cost("샤오홍슈")  # 聚光 + KOL/KOC/SNS
    xhs_gmv = s6.get("xhsGmvKRW", 0)
    xprods = _brand_products(xhs_products, brand, "샤오홍슈", idx)
    # 카테고리 매핑
    cats, unmatched = [], []
    for cn, v in sorted(xcat.items(), key=lambda x: -x[1]["cost"]):
        sku, amb = match_platform_name("샤오홍슈", cn, idx)
        adK = round(v["cost"] * FX)
        cats.append({"cn": cn, "adKRW": adK, "exp": v["exp"], "clk": v["clk"],
                     "sku": sku, "ko": idx["sku"][sku]["ko"] if sku else None, "matched": bool(sku)})
        if not sku:
            reason = amb["reason"] if amb else "매핑 테이블에 없음 — 상품/콘텐츠 구분 검토"
            cand = [idx["sku"][c]["ko"] for c in (amb["candidates"] if amb else []) if c in idx["sku"]]
            unmatched.append({"platform": "샤오홍슈", "cnName": cn, "metricLabel": "聚光 집행",
                              "valueKRW": adK, "reason": reason, "candidates": cand})
    xhs_p = {
        "label": "샤오홍슈(小红书) 聚光", "hasData": is_wm, "hasProductDetail": bool(xprods),
        "kpi": {"salesKRW": xhs_gmv, "gmvKRW": xhs_gmv, "adKRW": xhs_ad_juguang, "costKRW": xhs_cost_all,
                "roas": round(xhs_gmv / xhs_cost_all, 2) if xhs_cost_all else 0,
                "uv": s6.get("xhsUV", 0), "exp": xj["exp"], "clk": xj["clk"]},
        "categories": cats, "products": xprods,
        "note": ("聚光 로우 기준 집행 상세(카테고리별) · GMV/UV 는 기여 GMV 소스 기준(탭③과 소폭 상이)"
                 if is_wm else "샤오홍슈 " + brand + " 데이터 미확보 — 시트 `플랫폼_샤오홍슈` 입력 대기"),
    }

    byPlatform = {"티몰": tmall_p, "도우인": douyin_p, "샤오홍슈": xhs_p}

    # ---- 통합(합산) ----
    comb_sales = tmall_p["kpi"]["salesKRW"] + douyin_p["kpi"]["salesKRW"] + xhs_p["kpi"]["salesKRW"]
    comb_cost = tm_cost + dy_cost + xhs_cost_all
    by_sales = [{"platform": p, "salesKRW": byPlatform[p]["kpi"]["salesKRW"]} for p in ["티몰", "도우인", "샤오홍슈"]]
    # 매핑 합산 상품(티몰 매출 + 도우인 상품 매출 + 샤오홍슈 확정매칭 聚光비)
    prod_by_sku = {}
    def bucket(sku, ko):
        return prod_by_sku.setdefault(sku, {"sku": sku, "ko": ko, "티몰": 0, "도우인": 0, "샤오홍슈_ad": 0})
    for p in tmall_p["products"]:
        sku = idx["tmall"].get(p.get("id", ""))
        ko = idx["sku"][sku]["ko"] if sku else p["name"]
        bucket(sku or ("TM:" + (p.get("id") or p["name"])), ko)["티몰"] += p["payKRW"]
    for p in douyin_p["products"]:
        sku = p["sku"] or ("DY:" + p["name"])
        bucket(sku, p["ko"])["도우인"] += p["payKRW"]
    for c in cats:
        if c["matched"]:
            bucket(c["sku"], idx["sku"][c["sku"]]["ko"])["샤오홍슈_ad"] += c["adKRW"]
    comb_products = sorted(prod_by_sku.values(), key=lambda x: x["티몰"] + x["도우인"], reverse=True)

    combined = {
        "kpi": {"salesKRW": round(comb_sales), "costKRW": round(comb_cost),
                "roas": round(comb_sales / comb_cost, 2) if comb_cost else 0,
                "orders": douyin_p["kpi"]["orders"]},
        "bySales": by_sales, "products": comb_products,
        "note": "매출=티몰 결제+도우인 결제+샤오홍슈 기여GMV · 비용=채널 귀속 마케팅비 합",
    }

    return {"brand": brand, "month": month, "storeType": store_type,
            "order": ["티몰", "도우인", "샤오홍슈"], "byPlatform": byPlatform,
            "combined": combined, "unmatched": unmatched}

def mapping_out():
    return {"skus": [{"sku": m["sku"], "ko": m["ko"], "brand": m.get("brand", ""),
                      "status": m.get("status", "확정"), "tmallIds": m.get("tmallIds", []),
                      "douyin": m.get("douyin", []), "xhs": m.get("xhs", []),
                      "note": m.get("note", "")} for m in MAPPING],
            "ambiguous": AMBIGUOUS}

# =========================================================================
#  H. 조립 · 렌더 · 시트 어댑터 · 시드 CSV  (신규/확장)
# =========================================================================
def assemble(tmall_u, douyin, xhs, month=None):
    tmall = compute_tmall(tmall_u)
    if month is None:
        month = tmall["months"][-1] if tmall["months"] else "6월"
    data = {"fx": FX}
    data.update(tmall)  # months, series, byMonth
    data["daily"] = compute_daily(douyin, month=month)
    data["platforms"] = compute_platforms(tmall, douyin, xhs, month=month)
    data["mapping"] = mapping_out()
    return data

def render(data):
    if not os.path.exists(TEMPLATE):
        raise SystemExit("template.html 이 없습니다.")
    tpl = open(TEMPLATE, "r", encoding="utf-8").read()
    if "__DATA__" not in tpl:
        raise SystemExit("template.html 에 __DATA__ 자리표시자가 없습니다.")
    html = tpl.replace("__DATA__", json.dumps(data, ensure_ascii=False, separators=(",", ":")))
    open(OUT_HTML, "w", encoding="utf-8").write(html)

def tmall_from_datajson():
    """기존 data.json 에서 티몰 파트(months/series/byMonth) 로드 — 원본 엑셀 없이 재조립."""
    d = json.load(open(DATA_JSON, "r", encoding="utf-8"))
    return {"months": d["months"], "series": d["series"], "byMonth": d["byMonth"]}

# ---- 구글시트 어댑터 (탭별 공개 CSV) ----
def fetch_csv(url):
    import urllib.request
    with urllib.request.urlopen(url, timeout=30) as r:
        raw = r.read().decode("utf-8-sig")
    return list(csv.DictReader(io.StringIO(raw)))

def _read_products_tab(url):
    """플랫폼_도우인/샤오홍슈 '상품별' 탭 → {brand: [ {name,payCNY,orders,uv,tmallId} ]}."""
    out = {}
    for r in fetch_csv(url):
        name = r.get("상품명(중문)") or r.get("상품명") or r.get("商品名") or r.get("상품") or ""
        if not name: continue
        out.setdefault(r.get("브랜드") or "", []).append({
            "name": name, "tmallId": r.get("티몰_상품ID") or "",
            "payCNY": numf(r.get("결제금액(元)") or r.get("결제금액")), "payKRW": numf(r.get("결제금액(₩)")),
            "orders": numf(r.get("주문수")), "uv": numf(r.get("UV") or r.get("방문자"))})
    return out

def build_from_sheet(cfg_path):
    """sheet.config.json 의 탭별 CSV URL → data.json 재생성.
    신규 데이터(일자별/도우인/샤오홍슈/마케팅/매핑/상품별)를 시트에서, 티몰 파트는
    티몰 탭 미구현 시 기존 data.json 재사용. (스키마는 README 참조)"""
    if not os.path.exists(cfg_path):
        raise SystemExit("설정 파일이 없습니다: %s  (README의 sheet.config.json 예시 참조)" % cfg_path)
    cfg = json.load(open(cfg_path, "r", encoding="utf-8"))
    tabs = cfg.get("tabs", {})
    tmall = tmall_from_datajson()   # 티몰 파트: 기존 유지(시트 티몰 탭 미구현 시)
    # 도우인 일자별
    douyin = []
    if tabs.get("플랫폼_도우인"):
        for r in fetch_csv(tabs["플랫폼_도우인"]):
            store = r.get("스토어", "")
            brand = r.get("브랜드") or DOUYIN_STORE.get(store, ("", ""))[0]
            stype = "cross" if ("海外" in store or r.get("구분") == "역직구") else "domestic"
            douyin.append({"date": str(r.get("일자", ""))[:10], "store": store, "brand": brand, "stype": stype,
                           "gmvCNY": numf(r.get("GMV(元)") or r.get("GMV")), "payCNY": numf(r.get("결제금액(元)") or r.get("결제금액")),
                           "orders": int(numf(r.get("주문수"))), "buyers": int(numf(r.get("구매자수"))),
                           "aov": numf(r.get("객단가")), "expUV": int(numf(r.get("노출UV"))), "clickUV": int(numf(r.get("클릭UV")))})
    # 샤오홍슈 聚光
    xhs = {"byMonth": {}, "catByMonth": {}}
    if tabs.get("플랫폼_샤오홍슈"):
        for r in fetch_csv(tabs["플랫폼_샤오홍슈"]):
            m = re.sub(r"[^0-9]", "", str(r.get("월", "")))
            if not m: continue
            cost, exp, clk = numf(r.get("消费(元)") or r.get("消费")), int(numf(r.get("曝光"))), int(numf(r.get("点击")))
            b = xhs["byMonth"].setdefault(m, {"cost": 0.0, "exp": 0, "clk": 0})
            b["cost"] += cost; b["exp"] += exp; b["clk"] += clk
            cat = r.get("商品品类") or r.get("카테고리") or "(미지정)"
            c = xhs["catByMonth"].setdefault(m, {}).setdefault(cat, {"cost": 0.0, "exp": 0, "clk": 0})
            c["cost"] += cost; c["exp"] += exp; c["clk"] += clk
    dprods = _read_products_tab(tabs["도우인_상품별"]) if tabs.get("도우인_상품별") else None
    xprods = _read_products_tab(tabs["샤오홍슈_상품별"]) if tabs.get("샤오홍슈_상품별") else None
    # 시트에 신규 데이터가 없으면 기존 data.json 의 daily/platforms 유지
    d0 = json.load(open(DATA_JSON, "r", encoding="utf-8")) if os.path.exists(DATA_JSON) else {}
    if not douyin and not xhs["catByMonth"]:
        data = {"fx": FX}; data.update(tmall)
        data["daily"] = d0.get("daily", compute_daily([]))
        data["platforms"] = d0.get("platforms", compute_platforms(tmall, [], xhs, douyin_products=dprods, xhs_products=xprods))
        data["mapping"] = mapping_out()
    else:
        data = assemble_from_parts(tmall, douyin, xhs, douyin_products=dprods, xhs_products=xprods)
    _write_data(data)
    render(data)
    return data

def assemble_from_parts(tmall, douyin, xhs, month=None, douyin_products=None, xhs_products=None):
    if month is None:
        month = tmall["months"][-1] if tmall["months"] else "6월"
    data = {"fx": FX}; data.update(tmall)
    data["dailyStatus"] = daily_status_sample()   # 페이지① 상단 7월 현황(샘플 · 새 시트 연결 시 교체)
    data["daily"] = compute_daily(douyin, month=month)
    data["platforms"] = compute_platforms(tmall, douyin, xhs, month=month,
                                          douyin_products=douyin_products, xhs_products=xhs_products)
    data["mapping"] = mapping_out()
    return data

# ---- 시드 CSV(구글시트 붙여넣기용) ----
def write_seed_csv(data):
    os.makedirs(SEED_DIR, exist_ok=True)
    def w(name, header, rows):
        with open(os.path.join(SEED_DIR, name), "w", encoding="utf-8-sig", newline="") as f:
            wr = csv.writer(f); wr.writerow(header)
            for r in rows: wr.writerow(r)
    # 일자별매출(도우인 역직구, 브랜드별)
    d = data["daily"]; rows = []
    for b in d["brands"]:
        for x in d["series"][b]:
            rows.append([x["date"], b, "도우인", round(x["salesKRW"] / FX, 2), x["orders"], x["buyers"]])
    w("일자별매출.csv", ["일자", "브랜드", "플랫폼", "결제금액(元)", "주문수", "구매자수"], rows)
    # 상품매핑
    rows = [[m["sku"], m["ko"], m.get("brand", ""), ";".join(m.get("tmallIds", [])),
             ";".join(m.get("douyin", [])), ";".join(m.get("xhs", [])), m.get("status", "확정"), m.get("note", "")]
            for m in MAPPING]
    for a in AMBIGUOUS:
        rows.append(["", "", "", "", "", a["cn"], "검토", a["reason"]])
    w("상품매핑.csv", ["SKU코드", "한글상품명", "브랜드", "티몰_상품ID(;)", "도우인_별칭(;)", "샤오홍슈_별칭(;)", "상태", "비고"], rows)
    # 마케팅비(6월, 플랫폼 귀속)
    wm = data["platforms"]["byBrand"][WM_BRAND]; m6 = wm["month"]; bm = data["byMonth"].get(m6, {})
    rows = [[COST_PLATFORM.get(c["item"], "기타"), c["owner"], c["item"], round(c["cur"] / FX), m6]
            for c in bm.get("costItems", [])]
    w("마케팅비.csv", ["플랫폼", "주체", "항목", "금액(元)", "월"], rows)
    # 샤오홍슈 聚光 카테고리
    rows = [[m6, c["cn"], round(c["adKRW"] / FX, 2), c["exp"], c["clk"], "확정" if c["matched"] else "검토"]
            for c in wm["byPlatform"]["샤오홍슈"]["categories"]]
    w("플랫폼_샤오홍슈.csv", ["월", "商品品类", "消费(元)", "曝光", "点击", "매핑상태"], rows)
    # 도우인 일자별(스토어별 원본 형식 · 붙여넣기용)
    dd = data["daily"]
    rows = [[x["date"], b, "海外旗舰店", round(x.get("gmvKRW", x["salesKRW"]) / FX, 2), round(x["salesKRW"] / FX, 2),
             x["orders"], x["buyers"]] for b in dd["brands"] for x in dd["series"][b]]
    w("플랫폼_도우인.csv", ["일자", "브랜드", "스토어", "GMV(元)", "결제금액(元)", "주문수", "구매자수"], rows)
    # 상품별 입력 틀(빈 템플릿 + 예시 1행) — 채우면 상품 단위 통합에 자동 반영
    w("플랫폼_도우인_상품별.csv", ["월", "브랜드", "상품명(중문)", "티몰_상품ID", "결제금액(元)", "주문수", "UV"],
      [[m6, "웨이크메이크", "(예시)水光气垫", "930412381224", "", "", ""]])
    w("플랫폼_샤오홍슈_상품별.csv", ["월", "브랜드", "상품명(중문)", "티몰_상품ID", "결제금액(元)", "주문수", "UV"],
      [[m6, "웨이크메이크", "(예시)粉糖气垫", "930412381224", "", "", ""]])
    print("[시드 CSV] %s 에 생성: 일자별매출 / 플랫폼_도우인 / 플랫폼_샤오홍슈 / 상품매핑 / 마케팅비 / *_상품별(입력틀)" % SEED_DIR)

# =========================================================================
#  I. CLI
# =========================================================================
def _write_data(data):
    """data.json 저장 — dailyTmall(구글시트 전용)은 다른 모드에서 기존값 보존."""
    if "dailyTmall" not in data and os.path.exists(DATA_JSON):
        try:
            old = json.load(open(DATA_JSON, "r", encoding="utf-8"))
            if old.get("dailyTmall"): data["dailyTmall"] = old["dailyTmall"]
        except Exception: pass
    json.dump(data, open(DATA_JSON, "w", encoding="utf-8"), ensure_ascii=False, indent=1)

def _gsheet_id_from_cfg():
    if os.path.exists(SHEET_CFG):
        try: return json.load(open(SHEET_CFG, "r", encoding="utf-8")).get("gsheetId")
        except Exception: return None
    return None

def _summary(data):
    ms = data["months"]
    print("[완료] 티몰 월: %s (최신 %s)" % (ms, ms[-1] if ms else "-"))
    for key, lab in (("daily", "일자별·도우인"), ("dailyTmall", "일자별·티몰")):
        d = data.get(key)
        if d:
            print("       [%s] %s · %s · %s" % (lab, d["platformLabel"], d["coverage"],
                  " / ".join("%s %d일" % (b, len(d["series"][b])) for b in d["brands"])))
    for brand in data["platforms"]["brands"]:
        pf = data["platforms"]["byBrand"][brand]; um = pf["unmatched"]
        line = "       <%s> " % brand + " · ".join(
            "%s ₩%s" % (p, format(pf["byPlatform"][p]["kpi"].get("salesKRW", 0), ",")) for p in pf["order"])
        print(line)
        print("           통합 매출 ₩%s · ROAS %sx · 미매칭 %d건" % (
              format(pf["combined"]["kpi"]["salesKRW"], ","), pf["combined"]["kpi"]["roas"], len(um)))

def _local_douyin_xhs():
    douyin = parse_douyin(RAW_MULTI) if os.path.exists(RAW_MULTI) else []
    xhs = parse_xhs_juguang(RAW_MULTI) if os.path.exists(RAW_MULTI) else {"byMonth": {}, "catByMonth": {}}
    return douyin, xhs

def main():
    args = sys.argv[1:]
    if not args:
        print(__doc__); sys.exit(0)
    cmd = args[0]

    if cmd == "--render-only":
        data = json.load(open(DATA_JSON, "r", encoding="utf-8"))
        render(data); print("[완료] index.html 재생성"); return

    if cmd == "--seed-csv":
        data = json.load(open(DATA_JSON, "r", encoding="utf-8"))
        write_seed_csv(data); return

    if cmd == "--gsheet":
        # 구글시트('중국역직구 플랫폼 판매데이터') → 티몰 월별 + 티몰 일자별 동기화
        src = args[1] if len(args) > 1 else _gsheet_id_from_cfg()
        if not src: raise SystemExit("사용법: python build.py --gsheet <시트URL 또는 ID>  (또는 sheet.config.json 의 gsheetId)")
        sh = gsheet_reader(src)
        tmall = compute_tmall(_parse_tmall_sheets(sh))
        if not tmall["months"]:
            raise SystemExit("시트에서 티몰 데이터를 읽지 못했습니다 — 공유(링크 보기)·탭명 확인.")
        douyin, xhs = _local_douyin_xhs()   # 도우인·샤오홍슈는 아직 시트에 없음 → 로컬 원본 유지
        data = assemble_from_parts(tmall, douyin, xhs)
        data["dailyTmall"] = daily_struct(parse_tmall_daily(sh),
                                          "티몰글로벌(天猫国际) · 海外旗舰店",
                                          "티몰 일 결제금액 · 증정 제외 · 환율 220원")
        _write_data(data); render(data); _summary(data); return

    if cmd == "--sheet":
        cfg = args[1] if len(args) > 1 else SHEET_CFG
        data = build_from_sheet(cfg); _summary(data); return

    if cmd == "--seed":
        src = args[1] if len(args) > 1 else RAW_MULTI
        if not os.path.exists(src): raise SystemExit("원본이 없습니다: %s" % src)
        tmall = tmall_from_datajson()
        douyin = parse_douyin(src); xhs = parse_xhs_juguang(src)
        data = assemble_from_parts(tmall, douyin, xhs)
        _write_data(data); render(data); _summary(data); return

    if cmd == "--from-raw":
        if len(args) < 2: raise SystemExit("사용법: python build.py --from-raw <index_2.html>")
        tmall = compute_tmall(seed_from_raw(args[1]))
        douyin, xhs = _local_douyin_xhs()
        data = assemble_from_parts(tmall, douyin, xhs)
        _write_data(data); render(data); _summary(data); return

    # 기본: 티몰 원본 엑셀
    path = cmd
    if not os.path.exists(path): raise SystemExit("엑셀 파일을 찾을 수 없습니다: %s" % path)
    tmall = compute_tmall(parse_workbook(path))
    douyin, xhs = _local_douyin_xhs()
    data = assemble_from_parts(tmall, douyin, xhs)
    _write_data(data); render(data); _summary(data)

if __name__ == "__main__":
    main()
