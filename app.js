// =============================================
// app.js — 주간 시간당 채산이익 보고서 앱
// =============================================

const App = (() => {

  // ---- 상태 ----
  let weekOffset = 0;
  let actualData = {};  // { 채널: { sales, cogs } }
  let weeklyData = {};  // { weekKey: { label, goal, actual, saved } }
  const _saved = localStorage.getItem('varRates');
  let varRates = _saved ? JSON.parse(_saved) : JSON.parse(JSON.stringify(CONFIG.VAR_RATES));
  const _cogsHist = localStorage.getItem('channelCogsHistory');
  let channelCogsHistory = _cogsHist ? JSON.parse(_cogsHist) : {};  // { 채널: [rate, ...] }

  function makeDefaultGoal() {
    const sales   = CONFIG.CHANNELS.reduce((s, c) => s + CONFIG.CH_DEFAULTS[c][0], 0);
    const cogs    = CONFIG.CHANNELS.reduce((s, c) => s + CONFIG.CH_DEFAULTS[c][1], 0);
    const varComm = CONFIG.CHANNELS.reduce((s, c) => s + CONFIG.CH_DEFAULTS[c][0] * (varRates[c]?.commission || 0) / 100, 0);
    const varLogi = CONFIG.CHANNELS.reduce((s, c) => s + CONFIG.CH_DEFAULTS[c][0] * (varRates[c]?.logistics  || 0) / 100, 0);
    const varC    = varComm + varLogi;
    const fixed   = CONFIG.DEFAULT_FIXED;
    const other   = CONFIG.DEFAULT_OTHER;
    const hours   = CONFIG.DEFAULT_HOURS;
    const contrib = sales - cogs - varC;
    const profit  = contrib - fixed + other;
    return { sales, cogs, varC, varComm, varLogi, fixed, other, hours, contrib, profit };
  }
  let currentGoal = JSON.parse(localStorage.getItem('currentGoal') || 'null') || makeDefaultGoal();

  // ---- 유틸 ----
  const fm  = n => Math.round(n).toLocaleString();
  const fmp = (n, d = 1) => isFinite(n) && !isNaN(n) ? n.toFixed(d) + '%' : '-';
  const gv  = id => parseFloat(document.getElementById(id)?.value) || 0;
  const $   = id => document.getElementById(id);
  const tag = c => `<span class="tag tag--${c}">${c}</span>`;

  // ---- 주차 계산 ----
  function getWeekInfo(off) {
    const now = new Date();
    const day = now.getDay();
    const diff = day === 0 ? -6 : 1 - day;
    const mon = new Date(now);
    mon.setDate(now.getDate() + diff + off * 7);
    const sun = new Date(mon);
    sun.setDate(mon.getDate() + 6);
    const yr = mon.getFullYear();
    const jan1 = new Date(yr, 0, 1);
    const wn = Math.ceil(((mon - jan1) / 86400000 + jan1.getDay() + 1) / 7);
    const f = d => `${d.getMonth() + 1}/${d.getDate()}`;
    return { year: yr, week: wn, label: `${yr}년 ${wn}주차 (${f(mon)}~${f(sun)})`, key: `${yr}-W${wn}` };
  }

  function updateWeekBar() {
    const wi = getWeekInfo(weekOffset);
    $('wk-label').textContent = wi.label;
    const b = $('wk-badge');
    if (weekOffset === 0) {
      b.textContent = '금주'; b.style.background = '#e6f1fb'; b.style.color = '#0c447c';
    } else if (weekOffset === 1) {
      b.textContent = '익주'; b.style.background = '#e1f5ee'; b.style.color = '#085041';
    } else if (weekOffset < 0) {
      b.textContent = `${Math.abs(weekOffset)}주 전`; b.style.background = '#f1efe8'; b.style.color = '#5f5e5a';
    } else {
      b.textContent = `${weekOffset}주 후`; b.style.background = '#faeeda'; b.style.color = '#633806';
    }
  }

  function ofw(n) {
    weekOffset = n === 0 ? 0 : weekOffset + n;
    updateWeekBar();
    renderReport();
  }

  // ---- 데이터 계산 ----
  function calcTotalVarCost() {
    return CONFIG.CHANNELS.reduce((s, c) => {
      const sales = gv(`g-sales-${c}`);
      return s + sales * (varRates[c]?.commission || 0) / 100
               + sales * (varRates[c]?.logistics  || 0) / 100;
    }, 0);
  }

  function getGoalData() {
    const sales  = CONFIG.CHANNELS.reduce((s, c) => s + gv(`g-sales-${c}`), 0);
    const cogs   = CONFIG.CHANNELS.reduce((s, c) => s + gv(`g-cogs-${c}`), 0);
    const varC   = calcTotalVarCost();
    const fixed  = gv('g-fixed');
    const other  = gv('g-other');
    const hours  = gv('g-hours');
    const contrib = sales - cogs - varC;
    const profit  = contrib - fixed + other;
    return { sales, cogs, varC, fixed, other, hours, contrib, profit };
  }

  function getActualData() {
    const sales  = CONFIG.CHANNELS.reduce((s, c) => s + (actualData[c]?.sales || 0), 0) / 1e6;
    const cogs   = CONFIG.CHANNELS.reduce((s, c) => s + (actualData[c]?.cogs  || 0), 0) / 1e6;
    const varC   = CONFIG.CHANNELS.reduce((s, c) => {
      const actSales = (actualData[c]?.sales || 0) / 1e6;
      return s + actSales * (varRates[c]?.commission || 0) / 100
               + actSales * (varRates[c]?.logistics  || 0) / 100;
    }, 0);
    const fixed  = gv('g-fixed');
    const other  = gv('g-other');
    const hours  = gv('a-hours');
    const contrib = sales - cogs - varC;
    const profit  = contrib - fixed + other;
    return { sales, cogs, varC, fixed, other, hours, contrib, profit };
  }

  // ---- 빌드: 입력폼 ----
  function buildInputs() {
    ['sales', 'cogs'].forEach(type => {
      const cont = $(`${type}-inputs`);
      if (!cont) return;
      let html = '';
      CONFIG.CHANNELS.forEach(c => {
        const def = CONFIG.CH_DEFAULTS[c][type === 'sales' ? 0 : 1];
        const hint = type === 'cogs'
          ? `<span class="field__hint" id="g-cogs-hint-${c}" style="font-size:11px;color:#888"></span>`
          : '';
        html += `<div class="field">
          <label>${tag(c)}${hint}</label>
          <input type="number" id="g-${type}-${c}" value="${def}" class="input" oninput="App.recalc()">
          <span class="field__unit">백만원</span>
        </div>`;
      });
      html += `<div class="field">
        <label style="font-weight:600">합계</label>
        <input type="number" id="g-${type}-total" readonly class="input">
        <span class="field__unit">백만원</span>
      </div>`;
      cont.innerHTML = html;
    });
  }

  // ---- 빌드: 변동비 테이블 ----
  function buildVarCostTable() {
    const tbody = $('varCost-table-body');
    if (!tbody) return;
    let totC = 0, totL = 0, totV = 0, totS = 0;
    const rows = CONFIG.CHANNELS.map(c => {
      const s = gv(`g-sales-${c}`);
      const cr = varRates[c]?.commission || 0;
      const lr = varRates[c]?.logistics  || 0;
      const ca = s * cr / 100, la = s * lr / 100, va = ca + la;
      totC += ca; totL += la; totV += va; totS += s;
      return `<tr>
        <td class="th-left">${tag(c)}</td>
        <td style="text-align:right" id="vc-sales-${c}">${fm(s)}</td>
        <td><input type="number" step="0.1" value="${cr}" id="vcr-comm-${c}" oninput="App.onVarRateChange('${c}')"></td>
        <td style="text-align:right" id="vc-comm-${c}">${fm(ca)}</td>
        <td><input type="number" step="0.1" value="${lr}" id="vcr-logi-${c}" oninput="App.onVarRateChange('${c}')"></td>
        <td style="text-align:right" id="vc-logi-${c}">${fm(la)}</td>
        <td style="text-align:right;font-weight:600" id="vc-total-${c}">${fm(va)}</td>
      </tr>`;
    }).join('');
    tbody.innerHTML = rows + `<tr class="total-row">
      <td class="th-left">합계</td>
      <td style="text-align:right" id="vc-sales-total">${fm(totS)}</td>
      <td></td>
      <td style="text-align:right" id="vc-comm-total">${fm(totC)}</td>
      <td></td>
      <td style="text-align:right" id="vc-logi-total">${fm(totL)}</td>
      <td style="text-align:right" id="vc-var-total">${fm(totV)}</td>
    </tr>`;
  }

  function onVarRateChange(c) {
    varRates[c].commission = parseFloat($(`vcr-comm-${c}`)?.value) || 0;
    varRates[c].logistics  = parseFloat($(`vcr-logi-${c}`)?.value) || 0;
    updateVarCostRow(c);
    updateVarCostTotals();
    recalc();
    syncFeeSettingTable();
  }

  function updateVarCostRow(c) {
    const s = gv(`g-sales-${c}`);
    const ca = s * (varRates[c]?.commission || 0) / 100;
    const la = s * (varRates[c]?.logistics  || 0) / 100;
    const se = $(`vc-sales-${c}`); if (se) se.textContent = fm(s);
    const ce = $(`vc-comm-${c}`); if (ce) ce.textContent = fm(ca);
    const le = $(`vc-logi-${c}`); if (le) le.textContent = fm(la);
    const te = $(`vc-total-${c}`); if (te) te.textContent = fm(ca + la);
  }

  function updateVarCostTotals() {
    let totC = 0, totL = 0, totS = 0;
    CONFIG.CHANNELS.forEach(c => {
      const s = gv(`g-sales-${c}`);
      totS += s;
      totC += s * (varRates[c]?.commission || 0) / 100;
      totL += s * (varRates[c]?.logistics  || 0) / 100;
    });
    const se = $('vc-sales-total'); if (se) se.textContent = fm(totS);
    const ce = $('vc-comm-total'); if (ce) ce.textContent = fm(totC);
    const le = $('vc-logi-total'); if (le) le.textContent = fm(totL);
    const te = $('vc-var-total');  if (te) te.textContent = fm(totC + totL);
  }

  // ---- recalc ----
  function recalc() {
    ['sales', 'cogs'].forEach(type => {
      const total = CONFIG.CHANNELS.reduce((s, c) => s + gv(`g-${type}-${c}`), 0);
      const el = $(`g-${type}-total`); if (el) el.value = total;
    });
    updateVarCostTotals();
    CONFIG.CHANNELS.forEach(c => updateVarCostRow(c));
    renderReport();
    renderGoalSummary();
  }

  // ---- 목표 요약 ----
  function renderGoalSummary() {
    const g = getGoalData();
    const el = $('goal-summary'); if (!el) return;
    const varComm = CONFIG.CHANNELS.reduce((s, c) => s + gv(`g-sales-${c}`) * (varRates[c]?.commission || 0) / 100, 0);
    const varLogi = CONFIG.CHANNELS.reduce((s, c) => s + gv(`g-sales-${c}`) * (varRates[c]?.logistics  || 0) / 100, 0);
    el.innerHTML = [
      { l: '매출',          v: `${fm(g.sales)}백만` },
      { l: '변동비 (수수료+물류)', v: `${fm(g.varC)}백만`, s: `수수료 ${fm(varComm)} / 물류 ${fm(varLogi)}` },
      { l: '채산이익',      v: `${fm(g.profit)}백만`, s: fmp(g.profit / g.sales * 100) },
      { l: '시간당 채산이익', v: `${fm(g.hours > 0 ? Math.round(g.profit * 1e6 / g.hours) : 0)}원` },
    ].map(m => `<div class="metric">
      <div class="metric__label">${m.l}</div>
      <div class="metric__value">${m.v}</div>
      <div class="metric__sub">${m.s || ''}</div>
    </div>`).join('');
  }

  // ---- 보고서 렌더 ----
  function renderReport() {
    const cg   = currentGoal;       // 금주 목표 (고정)
    const gNext = getGoalData();    // 익주 목표 (현재 입력값)
    const hasActual = Object.keys(actualData).length > 0;
    const a = hasActual ? getActualData() : null;
    const gB = cg.sales, aB = a?.sales || 1;

    const fmt2 = (v, base, unit) =>
      unit === '원' ? fm(v) + '원' :
      unit === 'h'  ? fm(v) + 'h'  :
      base > 0      ? `${fm(v)} (${fmp(v / base * 100)})` : fm(v);

    const achStr = (gVal, aVal) => {
      if (!hasActual || aVal == null) return '-';
      const r = aVal / gVal * 100;
      return `<span class="${r >= 100 ? 'pos' : 'neg'}">${fmp(r)}</span>`;
    };

    // 익주목표 변동비 breakdown (nv)
    const varComm = CONFIG.CHANNELS.reduce((s, c) => s + gv(`g-sales-${c}`) * (varRates[c]?.commission || 0) / 100, 0);
    const varLogi = CONFIG.CHANNELS.reduce((s, c) => s + gv(`g-sales-${c}`) * (varRates[c]?.logistics  || 0) / 100, 0);
    // 금주실적 변동비 breakdown (av)
    const aVarComm = hasActual ? CONFIG.CHANNELS.reduce((s, c) => s + (actualData[c]?.sales || 0) / 1e6 * (varRates[c]?.commission || 0) / 100, 0) : null;
    const aVarLogi = hasActual ? CONFIG.CHANNELS.reduce((s, c) => s + (actualData[c]?.sales || 0) / 1e6 * (varRates[c]?.logistics  || 0) / 100, 0) : null;

    const rows = [
      { l: '매출',           gv: cg.sales,        av: a?.sales,   nv: gNext.sales,        pct: true,  cls: 'hl' },
      { l: '(-) 매출원가',   gv: cg.cogs,         av: a?.cogs,    nv: gNext.cogs,         pct: true,  ind: true },
      { l: '(-) 변동비',     gv: cg.varC,         av: a?.varC,    nv: gNext.varC,         pct: true,  ind: true },
      { l: '    판매수수료', gv: cg.varComm ?? 0, av: aVarComm,   nv: varComm,            pct: true,  ind2: true },
      { l: '    물류비',     gv: cg.varLogi ?? 0, av: aVarLogi,   nv: varLogi,            pct: true,  ind2: true },
      { sep: true },
      { l: '공헌이익',       gv: cg.contrib,      av: a?.contrib, nv: gNext.contrib,      pct: true,  cls: 'hl' },
      { l: '(-) 고정비',     gv: cg.fixed,        av: a?.fixed,   nv: gNext.fixed,        pct: true,  ind: true },
      { l: '(±) 영업외손익', gv: cg.other,        av: a?.other,   nv: gNext.other,        pct: true,  ind: true },
      { sep: true },
      { l: '채산이익',       gv: cg.profit,       av: a?.profit,  nv: gNext.profit,       pct: true,  cls: 'hl' },
      { l: '(-) 근무시간', hourRow: true, gv: cg.hours, av: hasActual ? a?.hours : null, nv: gNext.hours },
      { sep: true },
      {
        l: '시간당 채산이익',
        gv: cg.hours > 0 ? Math.round(cg.profit * 1e6 / cg.hours) : 0,
        av: a && a.hours > 0 ? Math.round(a.profit * 1e6 / a.hours) : null,
        nv: gNext.hours > 0 ? Math.round(gNext.profit * 1e6 / gNext.hours) : 0,
        unit: '원', cls: 'hl'
      },
    ];

    const hrInput = (id, val) =>
      `<input type="number" id="${id}" value="${val ?? ''}" style="width:70px;text-align:right">h`;

    $('rep-body').innerHTML = rows.map(r => {
      if (r.sep) return '<tr><td colspan="5" style="padding:3px 0;border-bottom:none"></td></tr>';
      if (r.hourRow) return `<tr>
        <td>${r.l}</td>
        <td>${hrInput('rep-hours-cg', r.gv)}</td>
        <td>${r.av == null ? '-' : hrInput('rep-hours-a', r.av)}</td>
        <td>${hrInput('rep-hours-next', r.nv)}</td>
        <td>-</td>
      </tr>`;
      const cls = (r.cls || '') + (r.ind ? ' ind' : '') + (r.ind2 ? ' ind2' : '');
      return `<tr class="${cls}">
        <td>${r.l}</td>
        <td>${fmt2(r.gv, gB, r.unit)}</td>
        <td>${r.av == null ? '-' : fmt2(r.av, aB, r.unit)}</td>
        <td>${fmt2(r.nv, gB, r.unit)}</td>
        <td>${achStr(r.gv, r.av)}</td>
      </tr>`;
    }).join('');

    if (hasActual) {
      $('m-sales').textContent = fm(a.sales) + '백만';
      const gap = a.sales - cg.sales;
      $('m-sales-gap').innerHTML = `목표 대비 <span class="${gap >= 0 ? 'pos' : 'neg'}">${gap >= 0 ? '+' : ''}${fm(gap)}백만</span>`;
      $('m-profit').textContent = fm(a.profit) + '백만';
      $('m-profit-rate').textContent = fmp(a.profit / a.sales * 100);
      $('m-hourly').textContent = fm(a.hours > 0 ? Math.round(a.profit * 1e6 / a.hours) : 0) + '원';
      $('channel-breakdown').style.display = 'block';
      $('ch-body').innerHTML = CONFIG.CHANNELS.map(c => {
        const d = actualData[c] || { sales: 0, cogs: 0 };
        const s = d.sales / 1e6, co = d.cogs / 1e6;
        return `<tr>
          <td>${tag(c)}</td>
          <td>${fm(s)}백만</td>
          <td>${fm(co)}백만</td>
          <td>${s > 0 ? fmp(co / s * 100) : '-'}</td>
        </tr>`;
      }).join('');
    }
  }

  // ---- 변동비 설정 탭 ----
  function buildFeeSettingTable() {
    const tbody = $('fee-setting-body'); if (!tbody) return;
    tbody.innerHTML = CONFIG.CHANNELS.map(c => {
      const cr = varRates[c]?.commission || 0;
      const lr = varRates[c]?.logistics  || 0;
      return `<tr>
        <td class="th-left">${tag(c)}</td>
        <td><input type="number" step="0.1" value="${cr}" id="fs-comm-${c}" oninput="App.onFeeSettingChange('${c}')"></td>
        <td><input type="number" step="0.1" value="${lr}" id="fs-logi-${c}" oninput="App.onFeeSettingChange('${c}')"></td>
        <td style="text-align:center" id="fs-total-${c}">${(cr + lr).toFixed(1)}%</td>
      </tr>`;
    }).join('');
  }

  function onFeeSettingChange(c) {
    const cr = parseFloat($(`fs-comm-${c}`)?.value) || 0;
    const lr = parseFloat($(`fs-logi-${c}`)?.value) || 0;
    const te = $(`fs-total-${c}`); if (te) te.textContent = (cr + lr).toFixed(1) + '%';
  }

  function syncFeeSettingTable() {
    CONFIG.CHANNELS.forEach(c => {
      const ce = $(`fs-comm-${c}`); if (ce) ce.value = varRates[c]?.commission || 0;
      const le = $(`fs-logi-${c}`); if (le) le.value = varRates[c]?.logistics  || 0;
      const te = $(`fs-total-${c}`); if (te) te.textContent = ((varRates[c]?.commission || 0) + (varRates[c]?.logistics || 0)).toFixed(1) + '%';
    });
  }

  function saveFeeSettings() {
    CONFIG.CHANNELS.forEach(c => {
      varRates[c].commission = parseFloat($(`fs-comm-${c}`)?.value) || 0;
      varRates[c].logistics  = parseFloat($(`fs-logi-${c}`)?.value) || 0;
      const vcrc = $(`vcr-comm-${c}`); if (vcrc) vcrc.value = varRates[c].commission;
      const vcrl = $(`vcr-logi-${c}`); if (vcrl) vcrl.value = varRates[c].logistics;
      onFeeSettingChange(c);
    });
    localStorage.setItem('varRates', JSON.stringify(varRates));
    recalc();
    alert('변동비율 저장 완료!');
  }

  // ---- 파일 업로드 ----
  function handleFile(input) {
    const file = input.files[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = e => {
      const wb = XLSX.read(e.target.result, { type: 'array', cellDates: true });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(ws, { defval: '' });
      const byBiz = {};
      data.forEach(row => {
        const 유형 = row[CONFIG.COL.SALE_TYPE] || '';
        if (CONFIG.EXCLUDE_TYPES.includes(유형)) return;
        const biz = row[CONFIG.COL.BIZ] || row[CONFIG.COL.BIZ_ALT] || '';
        if (!biz) return;
        if (!byBiz[biz]) byBiz[biz] = { sales: 0, cogs: 0, cnt: 0 };
        byBiz[biz].sales += parseFloat(row[CONFIG.COL.SALES] || 0);
        byBiz[biz].cogs  += parseFloat(row[CONFIG.COL.COGS]  || 0);
        byBiz[biz].cnt++;
      });

      const mapCont = $('map-rows');
      mapCont.innerHTML = Object.entries(byBiz)
        .sort((a, b) => b[1].sales - a[1].sales)
        .map(([biz, d]) => {
          const def = CONFIG.DEFAULT_MAPPING[biz] || '제외';
          const opts = [...CONFIG.CHANNELS, '제외'].map(c =>
            `<option value="${c}"${c === def ? ' selected' : ''}>${c}</option>`
          ).join('');
          return `<div class="map-row">
            <span class="map-row__name">${biz}</span>
            <span class="map-row__sales">${fm(d.sales / 1e6)}백만</span>
            <select id="map-${biz}">${opts}</select>
            <span class="map-row__cnt">${d.cnt}건</span>
          </div>`;
        }).join('');

      $('mapping-card').style.display = 'block';
      window._byBiz = byBiz;
    };
    reader.readAsArrayBuffer(file);
  }

  function applyMapping() {
    const byBiz = window._byBiz || {};
    actualData = {};
    CONFIG.CHANNELS.forEach(c => actualData[c] = { sales: 0, cogs: 0 });
    Object.keys(byBiz).forEach(biz => {
      const sel = $('map-' + biz);
      const ch  = sel ? sel.value : '제외';
      if (ch === '제외' || !actualData[ch]) return;
      actualData[ch].sales += byBiz[biz].sales;
      actualData[ch].cogs  += byBiz[biz].cogs;
    });

    const tbody = $('agg-body');
    let totS = 0, totC = 0, totN = 0;
    tbody.innerHTML = CONFIG.CHANNELS.map(c => {
      const d = actualData[c];
      const cnt = Object.entries(byBiz)
        .filter(([b]) => { const s = $('map-' + b); return s && s.value === c; })
        .reduce((s, [, v]) => s + v.cnt, 0);
      totS += d.sales; totC += d.cogs; totN += cnt;
      return `<tr>
        <td>${tag(c)}</td>
        <td>${fm(d.sales / 1e6)}백만</td>
        <td>${fm(d.cogs  / 1e6)}백만</td>
        <td>${d.sales > 0 ? fmp(d.cogs / d.sales * 100) : '-'}</td>
        <td>${cnt}건</td>
      </tr>`;
    }).join('') + `<tr class="hl">
      <td>합계</td>
      <td>${fm(totS / 1e6)}백만</td>
      <td>${fm(totC / 1e6)}백만</td>
      <td>${totS > 0 ? fmp(totC / totS * 100) : '-'}</td>
      <td>${totN}건</td>
    </tr>`;

    // 채널별 원가율 기록
    CONFIG.CHANNELS.forEach(c => {
      const d = actualData[c];
      if (d.sales > 0) {
        if (!channelCogsHistory[c]) channelCogsHistory[c] = [];
        channelCogsHistory[c].push(d.cogs / d.sales);
        if (channelCogsHistory[c].length > 12) channelCogsHistory[c].shift();
      }
    });
    localStorage.setItem('channelCogsHistory', JSON.stringify(channelCogsHistory));
    applyAvgCogsRates();

    $('agg-result').style.display = 'block';
    renderReport();
    alert('실적이 반영되었습니다! 주간 보고서 탭에서 확인하세요.');
  }

  // ---- 채널 평균 원가율 자동계산 ----
  function applyAvgCogsRates() {
    let applied = false;
    CONFIG.CHANNELS.forEach(c => {
      const hist = channelCogsHistory[c];
      if (!hist || hist.length === 0) return;
      const avgRate = hist.reduce((s, v) => s + v, 0) / hist.length;
      const salesEl = $(`g-sales-${c}`);
      const cogsEl  = $(`g-cogs-${c}`);
      if (salesEl && cogsEl) {
        cogsEl.value = Math.round(parseFloat(salesEl.value || 0) * avgRate);
        applied = true;
      }
      // hint 업데이트
      const hintEl = $(`g-cogs-hint-${c}`);
      if (hintEl) hintEl.textContent = `평균 ${fmp(avgRate * 100)} (${hist.length}주)`;
    });
    if (applied) recalc();
  }

  // ---- 구글 시트 연동 ----
  function setSyncStatus(state, text) {
    $('sync-dot').className = 'sync-dot ' + state;
    $('sync-text').textContent = text;
  }

  async function saveWeek() {
    const wi = getWeekInfo(weekOffset);
    const g  = currentGoal;
    const hasA = Object.keys(actualData).length > 0;
    const a  = hasA ? getActualData() : null;
    const btn = $('save-btn');
    btn.textContent = '저장 중...'; btn.disabled = true;
    setSyncStatus('loading', '저장 중...');
    try {
      const res  = await fetch(CONFIG.SCRIPT_URL, {
        method: 'POST',
        body: JSON.stringify({ action: 'save', weekKey: wi.key, weekLabel: wi.label, goal: g, actual: a || {} })
      });
      const json = await res.json();
      if (json.status === 'ok') {
        weeklyData[wi.key] = { label: wi.label, goal: g, actual: a, saved: new Date().toLocaleString('ko-KR') };
        setSyncStatus('ok', '저장 완료 — ' + wi.label);
        alert('구글 시트에 저장 완료!');
      } else {
        setSyncStatus('err', '저장 실패');
        alert('저장 실패: ' + json.message);
      }
    } catch (e) {
      setSyncStatus('err', '연결 오류');
      alert('오류: ' + e.message);
    }
    btn.textContent = '구글 시트에 저장'; btn.disabled = false;
  }

  async function loadFromSheets() {
    setSyncStatus('loading', '불러오는 중...');
    try {
      const resp = await fetch(CONFIG.SCRIPT_URL);
      const json = await resp.json();
      if (json.status === 'ok' && json.data) {
        weeklyData = {};
        json.data.forEach(row => {
          const key = row['주차키']; if (!key) return;
          weeklyData[key] = {
            label: row['주차라벨'],
            saved: row['저장일시'],
            goal: {
              sales:   +row['목표_매출']      || 0,
              cogs:    +row['목표_매출원가']   || 0,
              varC:    +row['목표_변동비']     || 0,
              fixed:   +row['목표_고정비']     || 0,
              other:   +row['목표_영업외손익'] || 0,
              hours:   +row['목표_근무시간']   || 0,
              contrib: +row['목표_공헌이익']   || 0,
              profit:  +row['목표_채산이익']   || 0,
            },
            actual: +row['실적_매출'] > 0 ? {
              sales:   +row['실적_매출']      || 0,
              cogs:    +row['실적_매출원가']   || 0,
              varC:    +row['실적_변동비']     || 0,
              fixed:   +row['실적_고정비']     || 0,
              other:   +row['실적_영업외손익'] || 0,
              hours:   +row['실적_근무시간']   || 0,
              contrib: +row['실적_공헌이익']   || 0,
              profit:  +row['실적_채산이익']   || 0,
            } : null,
          };
        });
        renderHistory();
        applyAvgCogsRates();
        setSyncStatus('ok', `${json.data.length}개 주차 불러옴`);
      } else {
        setSyncStatus('err', '불러오기 실패');
      }
    } catch (e) {
      setSyncStatus('err', '연결 오류');
    }
  }

  function renderHistory() {
    const el   = $('hist-list');
    const keys = Object.keys(weeklyData).sort().reverse();
    if (!keys.length) {
      el.innerHTML = '<p class="empty-msg">저장된 데이터가 없습니다.</p>';
      return;
    }
    el.innerHTML = keys.map(k => {
      const d = weeklyData[k];
      const hasA = d.actual && d.actual.sales > 0;
      return `<div class="hist-card">
        <div class="hist-card__header">
          <span class="hist-card__title">${d.label}</span>
          <span class="hist-card__date">${d.saved}</span>
        </div>
        <div class="hist-card__body">
          목표 ${fm(d.goal.sales)}백만 / 채산이익 ${fm(d.goal.profit)}백만
          ${hasA ? `<br>실적 ${fm(d.actual.sales)}백만 / 채산이익 ${fm(d.actual.profit)}백만 (달성률 ${fmp(d.actual.profit / d.goal.profit * 100)})` : ''}
        </div>
      </div>`;
    }).join('');
  }

  // ---- 기타 액션 ----
  function applyGoal() {
    const g = getGoalData();
    const varComm = CONFIG.CHANNELS.reduce((s, c) => s + gv(`g-sales-${c}`) * (varRates[c]?.commission || 0) / 100, 0);
    const varLogi = CONFIG.CHANNELS.reduce((s, c) => s + gv(`g-sales-${c}`) * (varRates[c]?.logistics  || 0) / 100, 0);
    currentGoal = { ...g, varComm, varLogi };
    localStorage.setItem('currentGoal', JSON.stringify(currentGoal));
    renderReport();
    renderGoalSummary();
    alert('금주 목표로 적용 완료!');
  }

  function recalcHours() {
    const cgH  = parseFloat($('rep-hours-cg')?.value)   || 0;
    const aH   = parseFloat($('rep-hours-a')?.value);
    const nxtH = parseFloat($('rep-hours-next')?.value) || 0;

    currentGoal.hours = cgH;
    localStorage.setItem('currentGoal', JSON.stringify(currentGoal));

    const aEl = $('a-hours'); if (aEl && !isNaN(aH)) aEl.value = aH;
    const nEl = $('g-hours'); if (nEl) nEl.value = nxtH;

    renderReport();
    renderGoalSummary();
  }

  function aiAnalysis() {
    const g = getGoalData();
    const a = Object.keys(actualData).length > 0 ? getActualData() : null;
    const msg = a
      ? `이번 주 손익 보고서를 분석해줘.\n목표 매출 ${fm(g.sales)}백만 / 실적 매출 ${fm(a.sales)}백만\n목표 채산이익 ${fm(g.profit)}백만 / 실적 채산이익 ${fm(a.profit)}백만\n목표 대비 실적 차이 원인과 익주 목표 달성 전략을 알려줘.`
      : '이번 주 손익 목표를 분석해줘. 익주 목표 달성 전략을 알려줘.';
    console.log('[AI 분석 요청]', msg);
    alert('AI 분석은 Claude.ai 웹앱에서만 동작합니다.');
  }

  function exportExcel() {
    const cg    = currentGoal;
    const gNext = getGoalData();
    const hasA  = Object.keys(actualData).length > 0;
    const a     = hasA ? getActualData() : null;

    const wiCur  = getWeekInfo(weekOffset);
    const wiNext = getWeekInfo(weekOffset + 1);

    const fmn  = n => Math.round(n).toLocaleString();
    const fpct = (val, base) => base > 0 ? `${fmn(val)}(${(val / base * 100).toFixed(1)}%)` : fmn(val);
    const av   = (val, base) => val != null ? fpct(val, base) : '-';

    const cgB = cg.sales, aB = a?.sales || 1, nB = gNext.sales;
    const cgH = cg.hours > 0 ? Math.round(cg.profit * 1e6 / cg.hours) : 0;
    const aH  = a && a.hours > 0 ? Math.round(a.profit * 1e6 / a.hours) : null;
    const nH  = gNext.hours > 0 ? Math.round(gNext.profit * 1e6 / gNext.hours) : 0;

    const rows = [
      [`BEP 시간당 채산이익 : ${CONFIG.BEP_HOURLY.toLocaleString()}원`, `${wiCur.label} 목표`, `${wiCur.label} 실적`, `${wiNext.label} 목표`, '목표 대비 실적에 대한 분석 / 금주 목표를 달성하기 위한 계획'],
      ['시간당채산이익(원)', fmn(cgH), aH != null ? fmn(aH) : '-', fmn(nH), ''],
      ['', '', '', '', ''],
      ['매출',              fpct(cg.sales,   cgB), av(a?.sales,   aB), fpct(gNext.sales,   nB), ''],
      ['(-) 매출원가',      fpct(cg.cogs,    cgB), av(a?.cogs,    aB), fpct(gNext.cogs,    nB), ''],
      ['(-) 변동비(원가외)', fpct(cg.varC,   cgB), av(a?.varC,    aB), fpct(gNext.varC,    nB), ''],
      ['', '', '', '', ''],
      ['공헌이익',          fpct(cg.contrib, cgB), av(a?.contrib, aB), fpct(gNext.contrib, nB), ''],
      ['(-) 고정비',        fpct(cg.fixed,   cgB), av(a?.fixed,   aB), fpct(gNext.fixed,   nB), ''],
      ['(±) 영업외손익',    fpct(cg.other,   cgB), av(a?.other,   aB), fpct(gNext.other,   nB), ''],
      ['', '', '', '', ''],
      ['채산이익',          fpct(cg.profit,  cgB), av(a?.profit,  aB), fpct(gNext.profit,  nB), ''],
      ['(+) 근무시간(시간)', fmn(cg.hours), a ? fmn(a.hours) : '-', fmn(gNext.hours), ''],
      ['', '', '', '', ''],
      ['(=) 시간당채산이익(원)', fmn(cgH), aH != null ? fmn(aH) : '-', fmn(nH), ''],
    ];

    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws['!cols'] = [{ wch: 20 }, { wch: 22 }, { wch: 22 }, { wch: 22 }, { wch: 50 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '주간보고서');
    XLSX.writeFile(wb, `주간보고서_${wiCur.key}.xlsx`);
  }

  // ---- 탭 전환 ----
  function initTabs() {
    document.querySelectorAll('.tab').forEach(btn => {
      btn.addEventListener('click', () => {
        document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
        document.querySelectorAll('.section').forEach(s => s.classList.remove('active'));
        btn.classList.add('active');
        const tabId = 'tab-' + btn.dataset.tab;
        document.getElementById(tabId).classList.add('active');
        if (btn.dataset.tab === 'history') loadFromSheets();
      });
    });
  }

  // ---- 파일 인풋 이벤트 ----
  function initFileInput() {
    const fi = $('fi');
    if (fi) fi.addEventListener('change', () => handleFile(fi));
  }

  // ---- 초기화 ----
  function init() {
    initTabs();
    initFileInput();
    buildInputs();
    buildVarCostTable();
    buildFeeSettingTable();
    updateWeekBar();
    recalc();
    applyAvgCogsRates();

    // 구글 시트 링크 업데이트
    const links = document.querySelectorAll('a[href*="spreadsheets"]');
    links.forEach(a => a.href = CONFIG.SHEET_URL);
  }

  document.addEventListener('DOMContentLoaded', init);

  // ---- Public API ----
  return {
    ofw, recalc,
    onVarRateChange, onFeeSettingChange,
    saveFeeSettings, applyGoal, applyMapping,
    saveWeek, loadFromSheets,
    aiAnalysis, exportExcel,
    recalcHours, applyAvgCogsRates,
  };

})();
