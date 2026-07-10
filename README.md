# 웨이크메이크 중국 역직구 대시보드 (티몰글로벌)

웨이크메이크(WAKEMAKE) 티몰글로벌 자영점의 **중국 역직구 통합 성과 · 마케팅 비용 대시보드**입니다.
외부 의존성 없는 **단일 정적 HTML**(`index.html`) — 인라인 CSS/JS + 직접 렌더링한 SVG 차트, 데이터 내장.
**월 선택형**: 상단 드롭다운으로 월을 바꾸면 KPI·인사이트·상품·비용·유입경로가 해당 월로 재계산됩니다.

## 구성
- **① 상반기 매출 · ROAS 추이** (선택월 강조)
- **② 월별 상품 Sell-out** (상품 랭킹 · 점유율 · 상세 테이블 · 고전환 SKU)
- **③ 월별 마케팅 비용 상세** (주체별/항목별 · 전월 대비 · 샤오홍슈 CID 성과)
- **④ 월별 점포 유입경로**

라이트/다크 테마 · 모바일 반응형 · 환율 1元 = 220원 기준(증정품 제외 결제금액).

---

## 📅 매월 갱신 방법 (자동화)

원본 티몰 엑셀 한 개만 있으면 됩니다. 3단계로 끝납니다.

```bash
# 0) (최초 1회) 엑셀 파서 라이브러리 설치
python -m pip install openpyxl

# 1) 티몰 원본 엑셀로 데이터·대시보드 재생성
python build.py 원본엑셀.xlsx
#   → data.json 과 index.html 이 새로 만들어지고, 최신 월이 기본 선택됩니다.

# 2) 배포 (GitHub 푸시 → Cloudflare 자동 재배포)
git add -A
git commit -m "7월 데이터"
git push
```
푸시 후 약 1분이면 라이브 URL에 반영됩니다.

### build.py 사용법
| 명령 | 설명 |
|---|---|
| `python build.py 파일.xlsx` | 티몰 원본 엑셀 → `data.json` + `index.html` 재생성 (매월 사용) |
| `python build.py --from-raw index_2.html` | 기존 대시보드 RAW에서 데이터 복원(검증용) |
| `python build.py --render-only` | `data.json`을 직접 수정한 뒤 `index.html`만 다시 렌더 |

### 원본 엑셀이 읽는 시트 (원본 대시보드와 동일)
`티몰글로벌`(목표) · `웨이크메이크_상품별판매데이터` · `웨이크메이크_티몰 점포 유입 현황`
· `웨이크메이크_티몰 내부 광고 현황` · `라이브방송 광고 내역` · `샤오홍씽-광고`
> 시트명·컬럼명이 바뀌면 `build.py`의 파서를 함께 수정해야 합니다.

---

## 파일 구조
| 파일 | 역할 |
|---|---|
| `index.html` | **배포 대상** — build.py가 생성 (template + data.json). Cloudflare가 이 파일을 서빙 |
| `template.html` | 대시보드 골격(HTML/CSS/JS). `__DATA__` 자리에 데이터가 주입됨 |
| `data.json` | build.py가 계산한 월별 데이터 (사람이 읽고 수정 가능) |
| `build.py` | 엑셀 → data.json → index.html 빌드 스크립트 |

## 배포 설정 (Cloudflare)
GitHub 저장소 → **Cloudflare (정적 자산 Worker/Pages)** 자동 배포.
- Build command: **없음** · Build output directory: **`/`** · Production branch: **`main`**
- `main`에 push하면 자동 재배포.

## ⚠️ 보안
실제 매출·마케팅비 등 **대외비 데이터**가 포함됩니다.
- 원본 엑셀(`*.xlsx`)은 `.gitignore`로 커밋 제외 (계산된 `data.json`만 저장소에 올라감)
- 접근 제한이 필요하면 **Cloudflare Access**로 특정 이메일만 허용 권장
- 검색엔진 비노출(`<meta name="robots" content="noindex,nofollow">`) 적용됨
