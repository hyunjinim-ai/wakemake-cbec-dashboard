# 웨이크메이크 중국 역직구 대시보드 (티몰글로벌)

웨이크메이크(WAKEMAKE) 티몰글로벌 자영점의 **중국 역직구 통합 성과 · 마케팅 비용 대시보드**입니다.
외부 의존성 없는 **단일 정적 HTML**(`index.html`) — 인라인 CSS/JS + 직접 렌더링한 SVG 차트, 데이터 내장.

## 구성
- **① 상반기(1~6월) 매출 · ROAS 추이**
- **② 6월 상품별 Sell-out** (상품 랭킹 · 점유율 · 상세 테이블)
- **③ 6월 마케팅 비용 상세** (주체별/항목별 · 전월 대비 · 샤오홍슈 CID 성과)
- **④ 6월 점포 유입경로**

라이트/다크 테마 대응 · 모바일 반응형 · 환율 1元 = 220원 기준(증정품 제외 결제금액).

## 배포
GitHub 저장소 → **Cloudflare Pages** 자동 배포.
- Framework preset: **None**
- Build command: **(비움)**
- Build output directory: **`/`**

`main` 브랜치에 푸시하면 Cloudflare Pages가 자동으로 재배포합니다.

## 수정 방법
데이터/디자인은 모두 `index.html` 한 파일 안에 있습니다.
- 데이터: `<script>` 상단 `const DATA = { ... }`
- 스타일/컬러 토큰: `<style>` 상단 `:root` (아틀란티스 그린 `#9BCE26` · 코랄 `#FF7878`)

## ⚠️ 보안
실제 매출·마케팅비 등 **대외비 데이터**가 포함됩니다.
- GitHub 저장소는 **Private** 유지
- 배포 URL은 **Cloudflare Access**로 접근 제한 권장 (`<meta name="robots" content="noindex,nofollow">` 적용됨)
