# 웨이크메이크·컬러그램 중국 역직구 통합 대시보드

티몰글로벌(天猫国际) · 도우인(抖音) · 샤오홍슈(小红书) **3개 플랫폼 통합** 성과·마케팅 대시보드.
외부 의존성 없는 **단일 정적 HTML**(`index.html`) — 인라인 CSS/JS + 직접 렌더링 SVG 차트, 데이터 내장.
환율 1元 = 220원 · 매출은 증정품 제외 결제금액 기준 · 라이트/다크 · 모바일 반응형.

## 3개 탭(뷰)
- **① 일자별 매출** — 6월 일자별 매출을 **웨이크메이크/컬러그램 브랜드별로 분리**. 추세·피크/급락일·주중주말·브랜드 비교 코멘트 **자동 생성**. (현재 소스: 도우인 역직구 로우데이터)
- **② 플랫폼 통합** — 티몰/도우인/샤오홍슈를 한 페이지에. **플랫폼 셀렉터**로 개별/합산 전환. SKU **매핑 합산**과 **미매칭 검토** 목록, 플랫폼 귀속 마케팅비, 전문가 제안 포함.
- **③ 티몰 월별** — 기존 티몰글로벌 6개월 추이·상품 sell-out·유입경로·마케팅비 대시보드(월 선택형).

---

## 📊 데이터 소스 = Google Sheets (빌드 타임 연동)

데이터는 코드에 하드코딩하지 않고 **구글시트에서 값을 가져와** `data.json`으로 계산 후 `index.html`에 구워넣습니다.
배포 사이트에는 시트 링크·인증정보가 **전혀 노출되지 않습니다**(정적 파일에 결과만 포함).

### ⭐ 빠른 동기화: `--gsheet` (권장)
**"중국역직구 플랫폼 판매데이터"** 시트(원본 티몰 워크북 + 티몰 일자별 탭 구조)를 그대로 읽습니다.
```bash
# sheet.config.json 에 gsheetId 를 넣은 뒤 (또는 URL 직접 전달):
python build.py --gsheet
python build.py --gsheet "https://docs.google.com/spreadsheets/d/<ID>/edit"
```
- 읽는 탭: `티몰글로벌`·`웨이크메이크_상품별판매데이터`·`웨이크메이크_티몰 내부 광고 현황`·`웨이크메이크_티몰 점포 유입 현황`·`라이브방송 광고 내역`·`샤오홍씽-광고` → **티몰 월별(③)**, 그리고 `웨이크메이크_티몰 일자별매출` → **티몰 일자별(①)**.
- 시트 공유는 **'링크가 있는 모든 사용자 - 뷰어'** 면 충분(gviz CSV로 읽음). 콤마·% 서식 자동 흡수.
- 도우인·샤오홍슈 sell-out 은 아직 이 시트에 없어 로컬 원본(`로우 확인중.xlsx`) 값을 유지합니다. 시트에 추가하면 확장됩니다.

### 시트 탭 스키마
| 탭 | 컬럼(핵심) | 용도 |
|---|---|---|
| `일자별매출` | 일자 · 브랜드 · 플랫폼 · 결제금액(元) · 주문수 · 구매자수 | 페이지① |
| `플랫폼_도우인` | 일자 · 스토어 · 브랜드 · GMV(元) · 결제금액(元) · 주문수 · 구매자수 · 노출UV · 클릭UV | ①·② 도우인 |
| `플랫폼_샤오홍슈` | 월 · 商品品类 · 消费(元) · 曝光 · 点击 · 매핑상태 | ② 샤오홍슈(聚光) |
| `마케팅비` | 플랫폼 · 주체(OY/스틸) · 항목 · 금액(元) · 월 | ②·③ 비용 |
| `유입경로` | 플랫폼 · 경로 · 방문자 · 구매자 | ② 유입 |
| `상품매핑` | SKU코드 · 한글상품명 · 브랜드 · 티몰_상품ID(;) · 도우인_별칭(;) · 샤오홍슈_별칭(;) · 상태(확정/검토) · 비고 | 매핑 |

> **스타터 CSV 제공**: `python build.py --seed-csv` 를 실행하면 위 스키마에 **실데이터가 채워진** CSV가 `sheets_seed/`에 생성됩니다.
> 각 파일을 구글시트의 동일 탭에 **붙여넣기**하면 시트가 바로 채워집니다. (이 폴더는 대외비라 커밋되지 않습니다.)

### 연동 절차 (최초 1회)
1. `python build.py --seed-csv` → `sheets_seed/*.csv` 를 구글시트 각 탭에 붙여넣기.
2. 각 탭을 **파일 → 공유 → 웹에 게시 → CSV** 로 게시하고 URL 복사.
3. `sheet.config.example.json` 을 `sheet.config.json` 으로 복사 후 탭별 URL 을 채움. (이 파일은 커밋 안 됨)
4. `python build.py --sheet` → 시트에서 값을 받아 `data.json` + `index.html` 재생성.

### 매월 갱신
```bash
# 구글시트에서 이번 달 데이터만 수정한 뒤:
python build.py --sheet         # 시트 → data.json → index.html
git add -A && git commit -m "7월 데이터" && git push   # Cloudflare 자동 재배포
```

> **더 강한 보안이 필요하면(비공개 시트):** 공개 CSV 대신 **서비스 계정**(GCP 서비스계정 JSON 키를 로컬 `*.sa.json`(gitignore)로 보관, 시트를 서비스계정 이메일에만 공유)으로 전환할 수 있습니다. `build.py` 의 `fetch_csv`(어댑터)만 Sheets API 호출로 교체하면 됩니다.

---

## 🧩 상품명 매핑 (중국어 원문 ↔ 한글 ↔ SKU)

플랫폼마다 상품명이 중국어로, 표기가 다르게(약어·띄어쓰기·신구버전) 나옵니다. 동일 상품을 정확히 합산하려고 매핑 테이블을 둡니다.
- **소스**: `mapping.py` 의 `MAPPING`(초기값) 또는 구글시트 `상품매핑` 탭.
- **정규화**: 공백·전각/반각·괄호·마커(新品/AD/구형/[임박]…) 제거 후 별칭 매칭. 티몰은 상품ID가 앵커.
- **애매하면 자동 합산 금지**: 어느 SKU에도 확신 매칭이 안 되는 항목(예: `十六色眼影` — 16색 데일리 vs 소프트블러링 판별 필요)은 **합산에서 제외**하고 페이지② **"⚠ 검토 필요"** 목록에 원문·후보·금액을 표시합니다. 시트에서 확정 별칭을 넣으면 다음 빌드부터 합산됩니다.

---

## 🛠 build.py 사용법
| 명령 | 설명 |
|---|---|
| `python build.py --gsheet [URL\|ID]` | **★'중국역직구 플랫폼 판매데이터' 시트 → 티몰 월별+일자별 동기화** (매월 권장) |
| `python build.py --sheet [sheet.config.json]` | 내 커스텀 스키마(탭별 공개 CSV) → data.json + index.html |
| `python build.py --seed-csv` | 현재 `data.json` → 시트 붙여넣기용 스타터 CSV(`sheets_seed/`) |
| `python build.py --render-only` | `data.json` 만 수정한 뒤 `index.html` 재생성 |
| `python build.py --seed [로우.xlsx]` | 원본 엑셀에서 실데이터로 최초 시드(도우인/샤오홍슈) |
| `python build.py --from-raw <index_2.html>` | 기존 티몰 RAW 에서 티몰 파트 복원(+도우인/샤오홍슈 흡수) |
| `python build.py <티몰원본.xlsx>` | 티몰 6시트 엑셀 → 티몰 파트 재계산 |

파서 라이브러리: `python -m pip install openpyxl` (엑셀 파싱 시).

## 파일 구조
| 파일 | 역할 |
|---|---|
| `index.html` | **배포 대상** — build.py 생성물(template + data.json). Cloudflare가 서빙 |
| `template.html` | 대시보드 골격(HTML/CSS/JS). `__DATA__` 자리에 데이터 주입 |
| `data.json` | build.py 계산 결과(티몰 series/byMonth · daily · platforms · mapping) |
| `build.py` | 시트/엑셀 → data.json → index.html 빌드 · 어댑터 구조 |
| `mapping.py` | 상품명 매핑 테이블(중문↔한글↔SKU) · 정규화 · 애매 별칭 |
| `sheet.config.example.json` | 시트 CSV URL 설정 예시(→ `sheet.config.json` 로 복사) |

## 배포 (Cloudflare)
GitHub 저장소 → **Cloudflare 정적 자산 Worker** 자동 배포. Build command 없음 · output `/` · branch `main`.
`main` push 시 약 1분 후 라이브 반영: **https://wakemake-cbec-dashboard.hyunjin-im.workers.dev/**

## ⚠️ 보안 (중요)
실제 매출·마케팅비 등 **대외비 데이터가 배포 페이지(HTML)에 그대로 포함**되며, 현재 URL·GitHub 저장소는 **공개** 상태입니다.
- 원본 엑셀(`*.xlsx/*.csv`)·시트 URL(`sheet.config.json`)·서비스계정 키(`*.sa.json`)·`sheets_seed/` 는 `.gitignore` 로 커밋 제외.
- 검색엔진 비노출(`<meta name="robots" content="noindex,nofollow">`) 적용됨.
- **강력 권고: [Cloudflare Access](https://developers.cloudflare.com/cloudflare-one/policies/access/) 로 사이트를 이메일 허용목록 뒤로 두세요.** 페이지에 박힌 수치를 보호하는 실질적 방법입니다. (미적용 시 URL 을 아는 누구나 열람 가능)
