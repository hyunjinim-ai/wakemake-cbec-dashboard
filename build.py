#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
웨이크메이크 중국 역직구 대시보드 — 월별 데이터 빌드 스크립트
================================================================
티몰 원본 엑셀(다중 시트) → data.json → index.html 자동 생성.

사용법:
  python build.py <원본엑셀.xlsx>          # 매월: 엑셀에서 데이터 생성 후 대시보드 재생성
  python build.py --from-raw <index_2.html> # 시드: 기존 대시보드(RAW)에서 data.json 복원(검증용)
  python build.py --render-only             # data.json 만 바꾼 뒤 index.html 재생성

산출물: data.json (계산된 데이터), index.html (template.html + data.json)
환율 1元 = 220원 · 매출은 증정품 제외 결제금액 기준.
"""
import sys, os, json, re

try:
    sys.stdout.reconfigure(encoding="utf-8")  # Windows 콘솔(cp949) 한글/기호 출력 보장
except Exception:
    pass

FX = 220
HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(HERE, "template.html")
DATA_JSON = os.path.join(HERE, "data.json")
OUT_HTML = os.path.join(HERE, "index.html")

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

def monthNum(m):
    mm = re.sub(r"[^0-9]", "", str(m))
    return int(mm) if mm else 0
def kname(pid, cn):
    return NAME_MAP.get(pid) or ("코드 " + str(pid)[-6:])
def is_gift(p):
    cn = str(p.get("cn") or "")
    if ("赠品" in cn) and (p.get("payCNY") or 0) == 0: return True
    if cn.startswith("【赠品】"): return True
    return False
def is_new(p):
    return "新品" in str(p.get("cn") or "")

# =========================================================================
#  A. 티몰 원본 엑셀 파서 (원본 대시보드 parseWorkbook 의 Python 이식)
# =========================================================================
def parse_workbook(path):
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True, read_only=True)
    def sheet(name):
        if name not in wb.sheetnames: return []
        return [list(r) for r in wb[name].iter_rows(values_only=True)]
    def norm(m):
        return re.sub(r"\s", "", "" if m is None else str(m)).replace("月", "월")
    def num(v):
        try: return float(v)
        except (TypeError, ValueError): return 0.0

    # 목표매출 (티몰글로벌): col[8]=="티몰" 행의 9~14열 = 6개월 목표
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

    # 상품별 판매
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
                "conv": round(num(cell(r, h, "상품 결제 전환율")) * 10000) / 100,
                "sConv": round(num(cell(r, h, "검색 유입 결제 전환율")) * 10000) / 100,
                "sUV": int(num(cell(r, h, "검색 유입 방문자 수"))),
            })

    # 점포 유입
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

    # 티몰 내부광고
    tmall_ad = {}
    ads = sheet("웨이크메이크_티몰 내부 광고 현황")
    if ads:
        h = ads[0]
        for r in ads[1:]:
            if not r: continue
            m = norm(cell(r, h, "통계 일자"))
            if m and m != "null":
                tmall_ad[m] = tmall_ad.get(m, 0) + num(cell(r, h, "광고 소모액(마케팅 비용)"))

    # 라이브방송 광고
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

    # 샤오홍슈 광고
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
        monthly.append({"month": m,
                        "salesCNY": sum(p["payCNY"] for p in rows),
                        "targetKRW": round(tmall_actual[i]) if i < len(tmall_actual) else 0,
                        "uv": sum(p["uv"] for p in rows),
                        "pv": 0})
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
#  B. 기존 대시보드(RAW)에서 시드 — 엑셀 없이 검증용
# =========================================================================
def seed_from_raw(html_path):
    txt = open(html_path, "r", encoding="utf-8").read()
    m = re.search(r"const RAW\s*=\s*`(.+?)`", txt, re.S)
    if not m: raise SystemExit("index_2.html 에서 const RAW 를 찾지 못했습니다")
    d = json.loads(m.group(1))
    def unpack(rows, keys): return [dict(zip(keys, r)) for r in rows]
    products = {mo: unpack(rows, d["pkeys"]) for mo, rows in d["products"].items()}
    traffic = {mo: unpack(rows, d["tkeys"]) for mo, rows in d["traffic"].items()}
    # monthly(pv 포함)·cost·xhs 는 RAW 에 이미 계산돼 있음
    return {"months": d["months"], "monthly": d["monthly"], "cost": d["cost"],
            "products": products, "traffic": traffic, "xhs": d.get("xhs", {})}

# =========================================================================
#  C. data.json 계산 (series + 월별 파생 데이터)
# =========================================================================
def compute_data(u):
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
            "name": kname(p["id"], p.get("cn")),
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

    return {"fx": FX, "months": months, "series": series, "byMonth": by_month}

# =========================================================================
#  D. 렌더 & CLI
# =========================================================================
def render(data):
    if not os.path.exists(TEMPLATE):
        raise SystemExit("template.html 이 없습니다 — 먼저 template.html 을 준비하세요.")
    tpl = open(TEMPLATE, "r", encoding="utf-8").read()
    if "__DATA__" not in tpl:
        raise SystemExit("template.html 에 __DATA__ 자리표시자가 없습니다.")
    html = tpl.replace("__DATA__", json.dumps(data, ensure_ascii=False, separators=(",", ":")))
    open(OUT_HTML, "w", encoding="utf-8").write(html)

def main():
    args = sys.argv[1:]
    if not args:
        print(__doc__); sys.exit(0)
    if args[0] == "--render-only":
        data = json.load(open(DATA_JSON, "r", encoding="utf-8"))
    elif args[0] == "--from-raw":
        if len(args) < 2: raise SystemExit("사용법: python build.py --from-raw <index_2.html>")
        data = compute_data(seed_from_raw(args[1]))
        json.dump(data, open(DATA_JSON, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    else:
        path = args[0]
        if not os.path.exists(path): raise SystemExit(f"엑셀 파일을 찾을 수 없습니다: {path}")
        data = compute_data(parse_workbook(path))
        json.dump(data, open(DATA_JSON, "w", encoding="utf-8"), ensure_ascii=False, indent=1)

    render(data)
    ms = data["months"]
    print(f"[완료] 월: {ms}  (최신: {ms[-1] if ms else '없음'})")
    print(f"       data.json / index.html 생성됨. 다음: git add -A && git commit -m \"{ms[-1] if ms else ''} 데이터\" && git push")

if __name__ == "__main__":
    main()
