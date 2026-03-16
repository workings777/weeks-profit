// =============================================
// config.js — 설정값 모음 (여기서만 수정하면 됩니다)
// =============================================

const CONFIG = {

  // 구글 Apps Script 웹훅 URL
  SCRIPT_URL: 'https://script.google.com/macros/s/AKfycbyH673n3hM7nEekhuL3CuJ0SDo89eqbU8e6o9k_gU12DuRyItvnsKdfePQuE3YRgQ4kvg/exec',

  // 구글 시트 URL (과거 주차 탭 링크)
  SHEET_URL: 'https://docs.google.com/spreadsheets/d/1gwxAe3RMLaa1rROeWSk0yBAnxdumAkh6J-du65zzcoI',

  // BEP 시간당 채산이익 기준값
  BEP_HOURLY: 40294,

  // 채널 정의
  CHANNELS: ['옴채운', '온라인', '특수채널'],

  // 채널별 기본 목표값 [매출, 매출원가] (백만원)
  CH_DEFAULTS: {
    옴채운:   [1300, 675],
    온라인:   [600,  312],
    특수채널: [1040, 565],
  },

  // 채널별 기본 변동비율 (%)
  VAR_RATES: {
    옴채운:   { commission: 5.0, logistics: 2.0 },
    온라인:   { commission: 8.0, logistics: 3.0 },
    특수채널: { commission: 3.0, logistics: 2.0 },
  },

  // 기본 고정비 (백만원)
  DEFAULT_FIXED: 391,

  // 기본 영업외손익 (백만원)
  DEFAULT_OTHER: -53,

  // 기본 근무시간 (시간)
  DEFAULT_HOURS: 2735,

  // 사업소 → 채널 기본 매핑
  DEFAULT_MAPPING: {
    '온라인':     '온라인',
    '자사몰사업소': '옴채운',
    '관계사':     '옴채운',
    '부산권역':   '옴채운',
    '북서권역':   '옴채운',
    '대전권역':   '옴채운',
    '남부권역(A)': '옴채운',
    '서부권역':   '옴채운',
    '동부권역(B)': '옴채운',
    '대구권역':   '옴채운',
    '북동권역':   '옴채운',
    '동부권역(A)': '옴채운',
    '광주권역':   '옴채운',
    '중부권역':   '옴채운',
    '남부권역(B)': '옴채운',
    '대리점기타': '옴채운',
    '폐쇄몰사업소': '특수채널',
    '사업개발B':  '특수채널',
    '하자조치내수': '제외',
    '쌤플':       '제외',
    '수출사업소': '제외',
    '북미사업소': '제외',
  },

  // 실적 파일 필터링 — 제외할 매출유형
  EXCLUDE_TYPES: ['정상수주반품', 'A/S용'],

  // 실적 파일 컬럼명
  COL: {
    BIZ:       '사업소',       // 사업소 컬럼
    BIZ_ALT:   '실적사업소',   // 대체 사업소 컬럼
    SALES:     '매출단가*수량', // 매출 컬럼
    COGS:      'ingoga x qty', // 원가 컬럼 (입고단가×수량)
    SALE_TYPE: '매출유형',     // 매출유형 컬럼
  },
};
