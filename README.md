# 주간 시간당 채산이익 보고서

## 파일 구조

```
weekly-profit-app/
├── index.html          # 메인 HTML (구조)
├── style.css           # 스타일
├── config.js           # ⭐ 설정값 (여기서만 수정)
├── app.js              # 앱 로직
└── weekly_profit_script.gs  # 구글 Apps Script (시트 연동)
```

## 빠른 시작

1. VS Code에서 폴더 열기
2. `index.html`을 브라우저에서 열기 (Live Server 확장 추천)
3. 또는 `open index.html` 명령으로 바로 열기

## 설정 변경 (`config.js`)

| 항목 | 설명 |
|------|------|
| `SCRIPT_URL` | 구글 Apps Script 웹훅 URL |
| `SHEET_URL` | 구글 시트 URL |
| `CHANNELS` | 채널 목록 |
| `CH_DEFAULTS` | 채널별 기본 목표값 |
| `VAR_RATES` | 채널별 기본 변동비율 |
| `DEFAULT_FIXED` | 기본 고정비 |
| `DEFAULT_MAPPING` | 사업소 → 채널 매핑 |
| `COL` | 실적 파일 컬럼명 |

## 채널 구조

| 채널 | 포함 사업소 |
|------|------------|
| 옴채운 | 유통 + 공식몰 (각 권역, 자사몰, 관계사) |
| 온라인 | 온라인 |
| 특수채널 | 폐쇄몰 + 사업개발 |

## 변동비 계산

```
변동비 = 판매수수료 + 물류비
판매수수료 = 채널별 매출 × 판매수수료율(%)
물류비     = 채널별 매출 × 물류비율(%)
```

## 구글 시트 연동

1. 구글 시트 열기
2. 확장 프로그램 → Apps Script
3. `weekly_profit_script.gs` 내용 붙여넣기
4. 배포 → 웹 앱으로 배포
5. 생성된 URL을 `config.js`의 `SCRIPT_URL`에 입력

## VS Code 추천 확장

- **Live Server** — 저장 시 브라우저 자동 새로고침
- **Prettier** — 코드 자동 포맷
