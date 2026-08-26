# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 프로젝트 개요

**TOOLSPIA WMS** — 사내 소규모(약 20명) 사용을 위한 웹 기반 재고관리 시스템.  
소스 코드는 `C:\Users\home\Desktop\WMS\Claude_Code\wms-cc\` 에 있다(현재 git 저장소). `C:\Users\home\Desktop\WMS\CURSOR_재고관리\` 는 옛 소스 위치이며 더 이상 사용하지 않는다.

## 실행 방법

```bash
# 의존성 설치 (최초 1회)
npm install

# 서버 실행 (프로덕션)
npm start
# 또는 bat 파일로 실행
WMS-실행.bat

# 개발 모드 (파일 변경 시 자동 재시작)
npm run dev
```

브라우저에서 `http://localhost:3000` 접속.

## 기술 스택

- **런타임**: Node.js 18+ (외부 프레임워크 없음 — `http` 모듈 직접 사용)
- **프론트엔드**: Vanilla JS + HTML/CSS (SPA 구조, 빌드 도구 없음)
- **데이터**: `data/db.json` 파일 단일 JSON DB (메모리에 올려두지 않고 매 요청마다 읽기/쓰기)
- **유일한 npm 패키지**: `xlsx` (엑셀 파싱/생성)
- **이카운트 ERP 연동**: HTTP API 호출 (`/api/ecount/...` 엔드포인트)

## 아키텍처

### 서버 (`server.js` — 약 5,000줄 단일 파일)

모든 백엔드 로직이 한 파일에 있다.

**계층 구조:**

```
요청 → pathname 직접 매칭 (if/else 체인) → 비즈니스 로직 함수 → readDb/writeDb → db.json
```

**주요 함수 그룹:**
- `readDb / writeDb` — JSON DB 읽기/쓰기. 파일 잠금 오류 시 재시도 로직 포함. 쓰기 전 자동 백업(`maybeBackupBeforeWrite`)
- `normalizeDb(db)` — DB 로드 시 구버전 호환 필드 보정. 새 필드 추가 시 여기에 기본값 초기화 코드 추가 필요
- `upsertProduct` — 상품 등록/수정 (code 기준 중복 체크)
- `addMovement / cancelMovement` — 입출고 이동 추가/취소
- `fetchEcountSession / callJsonApi` — 이카운트 API 세션 인증 및 호출
- `parseOutboundOrderRowsByPartner` — 파트너사(다이소/이마트/롯데)별 출고 주문서 파싱

**DB 스키마 (`data/db.json`):**
```json
{
  "products": [],        // 상품 마스터
  "movements": [],       // 입출고 이력
  "warehouses": [],      // 창고 목록
  "partners": { "inbound": [], "outbound": [], "purchase": [] },
  "managers": [],
  "seq": 1,
  "inboundSlipConfirmLog": [],
  "inboundPlanUpload": { "lines": [], "slipWorkflow": {} },
  "outboundOrderUpload": { "lines": [], "uploadedRows": [], "batches": [], "appliedBatchByPartner": {}, "slipWorkflow": {}, "confirmedLists": {} },
  "outboundCodeMasters": []
}
```

**주요 API 엔드포인트:**
- `GET/POST /api/products` — 상품 단건
- `POST /api/products/bulk` — 상품 엑셀 일괄 업로드
- `GET /api/stock` — 재고 현황
- `POST /api/movements`, `POST /api/movements/bulk` — 입출고
- `GET /api/history` — 입출고 이력
- `GET/POST /api/partners` — 거래처
- `GET/POST /api/warehouses` — 창고
- `/api/inbound-plan-upload/*` — 입고예정 업로드 워크플로우
- `/api/outbound-order-upload/*` — 출고 주문서 업로드 워크플로우 (다이소/이마트/롯데)
- `/api/ecount/*` — 이카운트 ERP 연동
- `/api/wcs-test/*` — WCS 자동화창고 API 테스트

### 프론트엔드 (`public/`)

- `index.html` — SPA 셸. 사이드바 버튼 + `<section id="view-*">` 컨테이너만 선언
- `app.js` — 모든 UI 로직 (단일 파일). `state` 전역 객체로 데이터 관리. 각 뷰는 `render*()` 함수로 innerHTML 갱신
- `styles.css` — 전역 스타일

**뷰 목록:** `dashboard`, `master`, `products`, `stock`, `inbound`, `inbound-plan`, `inbound-plan-2`, `outbound`, `outbound-plan`, `outbound-upload-daiso`, `outbound-upload-emart`, `outbound-upload-lotte`, `adjust`, `history`, `wcs-test`, `alert`

## 환경 변수 (`.env`)

이카운트 연동에 필요한 값들:
- `ECOUNT_COM_CODE`, `ECOUNT_USER_ID`, `ECOUNT_USER_PW`, `ECOUNT_API_KEY`
- `ECOUNT_OUTBOUND_SALE_CUST_EMART` — 이마트 거래처 코드
- `PORT`, `HOST` — 서버 포트/호스트 (기본: 3000, 0.0.0.0)
- `WMS_DB_BACKUP_MIN_MS` — 자동 백업 최소 간격 ms (기본 90000)
- `WMS_DB_BACKUP_KEEP` — 백업 파일 보관 개수 (기본 80)

## 주의사항

- **DB 동시성**: 단일 프로세스 단일 파일 DB. 다중 프로세스로 실행하면 충돌 위험.
- **`normalizeDb` 필수**: DB 구조 변경(필드 추가) 시 `normalizeDb()` 함수에 마이그레이션 코드를 추가해야 한다. 이 함수는 `readDb()` 호출마다 실행된다.
- **백업**: `data/db-backups/` 에 자동 백업 저장. 수동 백업은 `data/db.json` 복사.
- **엑셀 파싱**: 상품/입출고 업로드는 첫 번째 시트 기준. xlsx/xls/csv 지원.
