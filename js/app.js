(() => {
  'use strict';

  const C = window.SCDCore;
  const $ = id => document.getElementById(id);
  const palette = ['#2f7d4a', '#d97706', '#4f8f72', '#a7c957', '#5b8f5a', '#f59e0b', '#3f6f52', '#9fbd75'];
  const phaseBasePalette = new Map([['A', '#2f7d4a'], ['B', '#d97706'], ['C', '#4f8f72'], ['D', '#a7c957']]);
  const fontMap = {
    'noto-sans': '"Noto Sans JP", "Yu Gothic", Meiryo, sans-serif',
    'times-new-roman': '"Times New Roman", Times, serif',
    'ms-gothic': '"MS Gothic", "ＭＳ ゴシック", monospace',
    'ms-mincho': '"MS Mincho", "ＭＳ 明朝", serif',
    meiryo: 'Meiryo, "メイリオ", sans-serif',
    'yu-gothic': '"Yu Gothic", "游ゴシック", sans-serif',
    'yu-mincho': '"Yu Mincho", "游明朝", serif',
    'noto-serif': '"Noto Serif JP", serif',
    aptos: 'Aptos, Calibri, sans-serif',
    calibri: 'Calibri, Arial, sans-serif',
    arial: 'Arial, sans-serif',
    helvetica: 'Helvetica, Arial, sans-serif',
    'segoe-ui': '"Segoe UI", sans-serif',
    cambria: 'Cambria, Georgia, serif',
    georgia: 'Georgia, serif',
    verdana: 'Verdana, sans-serif',
    'courier-new': '"Courier New", monospace'
  };

  const state = {
    workbook: null,
    fileName: '',
    baseFileName: '',
    headers: [],
    rows: [],
    types: [],
    selection: null,
    dragging: false,
    dragAnchor: null,
    phases: [],
    phaseColors: {},
    slots: Array.from({ length: 4 }, (_, index) => ({ name: ['A1', 'B1', 'A2', 'B2'][index], range: '', values: [], excluded: 0 })),
    lastChartSize: { width: 1000, height: 620 }
  };

  function escapeHtml(value) {
    return String(value ?? '').replace(/[&<>'"]/g, char => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' }[char]));
  }

  function setStatus(id, text, type = 'neutral') {
    const el = $(id);
    if (!el) return;
    el.textContent = text;
    el.className = `status ${type}`;
  }

  function showMessage(id, text, type = 'error') {
    const el = $(id);
    if (!el) return;
    el.textContent = text;
    el.className = `message ${type}`;
  }

  function hideMessage(id) {
    const el = $(id);
    if (!el) return;
    el.className = 'message hidden';
    el.textContent = '';
  }

  function uniqueHeaders(rawHeaders, width) {
    const used = new Map();
    return Array.from({ length: width }, (_, index) => {
      const base = String(rawHeaders[index] ?? '').trim() || `Column ${C.columnNumberToLetters(index + 1)}`;
      const count = (used.get(base) || 0) + 1;
      used.set(base, count);
      return count === 1 ? base : `${base}_${count}`;
    });
  }

  function inferColumnType(index) {
    const nonMissing = state.rows.map(row => row[index]).filter(value => value !== null && value !== undefined && String(value).trim() !== '');
    if (!nonMissing.length) return 'empty';
    const numericCount = nonMissing.filter(C.isFiniteNumber).length;
    if (numericCount === nonMissing.length) return 'numeric';
    if (numericCount > 0) return 'mixed';
    return 'text';
  }

  function loadMatrix(matrix, fileName) {
    if (!Array.isArray(matrix) || !matrix.length) throw new Error('データが空です．');
    const width = Math.max(...matrix.map(row => Array.isArray(row) ? row.length : 0));
    if (!width) throw new Error('列を認識できませんでした．');
    state.headers = uniqueHeaders(matrix[0] || [], width);
    state.rows = matrix.slice(1).map(row => Array.from({ length: width }, (_, index) => row?.[index] ?? null));
    state.types = state.headers.map((_, index) => inferColumnType(index));
    state.fileName = fileName;
    state.selection = null;
    state.phases = [];
    state.phaseColors = {};
    renderDataSummary();
    renderGrid();
    populateColumnSelectors();
    renderSlots();
    setStatus('globalStatus', `${state.rows.length}行 × ${state.headers.length}列`, 'success');
    setStatus('visualStatus', '解析可能', 'success');
    setStatus('statsStatus', '範囲を登録してください', 'neutral');
    renderVisual();
  }

  function renderDataSummary() {
    const missing = state.rows.flat().filter(value => value === null || value === undefined || String(value).trim() === '').length;
    const mixed = state.types.filter(type => type === 'mixed').length;
    const el = $('dataSummary');
    el.classList.remove('hidden');
    el.innerHTML = `<span class="summary-pill">${escapeHtml(state.fileName)}</span><span class="summary-pill">${state.rows.length}データ行</span><span class="summary-pill">${state.headers.length}列</span><span class="summary-pill">欠損セル ${missing}</span><span class="summary-pill">混合型列 ${mixed}</span>`;
  }

  function renderGrid() {
    const wrap = $('dataGridWrap');
    if (!state.headers.length) {
      wrap.innerHTML = '<div class="empty-state">ファイルを読み込むと，ここにデータ表が表示されます．</div>';
      return;
    }
    const displayRows = Math.min(state.rows.length, 5000);
    let html = '<table class="data-grid"><thead><tr><th class="row-number">#</th>';
    state.headers.forEach((header, index) => {
      html += `<th title="${escapeHtml(header)}">${escapeHtml(header)}<span class="data-type">${escapeHtml(state.types[index])} · ${C.columnNumberToLetters(index + 1)}</span></th>`;
    });
    html += '</tr></thead><tbody>';
    for (let row = 0; row < displayRows; row += 1) {
      html += `<tr><td class="row-number">${row + 2}</td>`;
      for (let col = 0; col < state.headers.length; col += 1) {
        const value = state.rows[row][col];
        const missing = value === null || value === undefined || String(value).trim() === '';
        const invalid = !missing && state.types[col] === 'mixed' && !C.isFiniteNumber(value);
        html += `<td data-row="${row}" data-col="${col}" class="${missing ? 'missing' : invalid ? 'invalid' : ''}" title="${escapeHtml(value ?? '')}">${missing ? '—' : escapeHtml(value)}</td>`;
      }
      html += '</tr>';
    }
    html += '</tbody></table>';
    if (state.rows.length > displayRows) html += `<div class="message error">表示は先頭${displayRows}行までです．解析には全行を使用します．</div>`;
    wrap.innerHTML = html;
    bindGridSelection();
  }

  function bindGridSelection() {
    const wrap = $('dataGridWrap');
    wrap.querySelectorAll('td[data-row]').forEach(cell => {
      cell.addEventListener('mousedown', event => {
        event.preventDefault();
        const point = { row: Number(cell.dataset.row), col: Number(cell.dataset.col) };
        state.dragging = true;
        state.dragAnchor = point;
        setSelection(point, point);
      });
      cell.addEventListener('mouseenter', () => {
        if (!state.dragging) return;
        setSelection(state.dragAnchor, { row: Number(cell.dataset.row), col: Number(cell.dataset.col) });
      });
    });
  }

  function setSelection(a, b) {
    state.selection = {
      startRow: Math.min(a.row, b.row),
      endRow: Math.max(a.row, b.row),
      startCol: Math.min(a.col, b.col),
      endCol: Math.max(a.col, b.col)
    };
    updateSelectionHighlight();
  }

  function selectionToA1(selection) {
    if (!selection) return '';
    return `${C.columnNumberToLetters(selection.startCol + 1)}${selection.startRow + 2}:${C.columnNumberToLetters(selection.endCol + 1)}${selection.endRow + 2}`;
  }

  function updateSelectionHighlight() {
    document.querySelectorAll('.data-grid td.selected').forEach(cell => cell.classList.remove('selected'));
    if (!state.selection) return;
    document.querySelectorAll('.data-grid td[data-row]').forEach(cell => {
      const row = Number(cell.dataset.row);
      const col = Number(cell.dataset.col);
      if (row >= state.selection.startRow && row <= state.selection.endRow && col >= state.selection.startCol && col <= state.selection.endCol) cell.classList.add('selected');
    });
    const range = selectionToA1(state.selection);
    $('directRange').value = range;
    $('selectionLabel').textContent = ` 選択中：${range}（${state.selection.endRow - state.selection.startRow + 1}行 × ${state.selection.endCol - state.selection.startCol + 1}列）`;
  }

  function applyDirectRange() {
    try {
      if (!state.headers.length) throw new Error('先にデータを読み込んでください．');
      const parsed = C.parseA1Range($('directRange').value);
      if (parsed.endCol >= state.headers.length || parsed.endRow >= state.rows.length) throw new Error('指定範囲がデータ表の範囲を超えています．');
      state.selection = parsed;
      updateSelectionHighlight();
      document.querySelector(`td[data-row="${parsed.startRow}"][data-col="${parsed.startCol}"]`)?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
    } catch (error) {
      setStatus('globalStatus', error.message, 'error');
    }
  }

  function populateColumnSelectors() {
    ['timeColumn', 'phaseColumn', 'dataColumn'].forEach(id => {
      $(id).innerHTML = '<option value="">選択してください</option>' + state.headers.map((header, index) => `<option value="${index}">${escapeHtml(header)} (${C.columnNumberToLetters(index + 1)}列・${state.types[index]})</option>`).join('');
    });
    const findHeader = patterns => state.headers.findIndex(header => patterns.some(pattern => pattern.test(header)));
    let time = findHeader([/^day$/i, /time/i, /session/i, /測定/i, /日/i]);
    let phase = findHeader([/phase/i, /フェーズ/i, /条件/i]);
    if (time < 0) time = state.types.findIndex(type => type === 'numeric');
    if (phase < 0) phase = state.types.findIndex(type => type === 'text');
    let data = state.types.findIndex((type, index) => type === 'numeric' && index !== time);
    if (data < 0) data = state.types.findIndex(type => type === 'numeric');
    if (time >= 0) $('timeColumn').value = String(time);
    if (phase >= 0) $('phaseColumn').value = String(phase);
    if (data >= 0) $('dataColumn').value = String(data);
  }

  function phaseBaseKey(phase) {
    const match = String(phase.raw || phase.label || '').trim().match(/^[A-Za-z]+/);
    return (match?.[0] || String(phase.raw || phase.label || '')).toUpperCase();
  }

  function defaultPhaseColor(phase, index) {
    const key = phaseBaseKey(phase);
    return phaseBasePalette.get(key) || palette[index % palette.length];
  }

  function ensurePhaseColors(phases) {
    phases.forEach((phase, index) => {
      if (!state.phaseColors[phase.label]) {
        const color = defaultPhaseColor(phase, index);
        state.phaseColors[phase.label] = { point: color, line: color };
      }
    });
    renderPhaseColorControls(phases);
  }

  function renderPhaseColorControls(phases) {
    const el = $('phaseColorControls');
    const labels = phases.map(phase => phase.label);
    const existing = [...el.querySelectorAll('.phase-color-row')].map(row => row.dataset.phase);
    if (JSON.stringify(labels) === JSON.stringify(existing)) return;
    el.innerHTML = phases.map(phase => `<div class="phase-color-row" data-phase="${escapeHtml(phase.label)}"><strong>${escapeHtml(phase.label)}</strong><label>点<input type="color" data-color-type="point" value="${state.phaseColors[phase.label].point}"></label><label>線<input type="color" data-color-type="line" value="${state.phaseColors[phase.label].line}"></label></div>`).join('');
    el.querySelectorAll('input[type=color]').forEach(input => input.addEventListener('input', event => {
      const row = event.target.closest('.phase-color-row');
      state.phaseColors[row.dataset.phase][event.target.dataset.colorType] = event.target.value;
      renderVisual();
    }));
  }

  function getVisualDataset() {
    const timeCol = Number($('timeColumn').value);
    const phaseCol = Number($('phaseColumn').value);
    const dataCol = Number($('dataColumn').value);
    if (![timeCol, phaseCol, dataCol].every(Number.isInteger)) throw new Error('時間列，フェーズ列，データ列をすべて選択してください．');
    const phases = C.splitConsecutivePhases(state.rows.map(row => row[phaseCol]));
    phases.forEach(phase => {
      const points = phase.indices.map(rowIndex => ({ rowIndex, x: C.toNumber(state.rows[rowIndex][timeCol]), y: C.toNumber(state.rows[rowIndex][dataCol]) })).filter(point => point.x !== null && point.y !== null);
      phase.points = points;
      phase.x = points.map(point => point.x);
      phase.y = points.map(point => point.y);
      phase.excluded = phase.indices.length - points.length;
    });
    state.phases = phases;
    ensurePhaseColors(phases);
    return { phases, timeCol, phaseCol, dataCol, timeName: state.headers[timeCol], dataName: state.headers[dataCol] };
  }

  function optionalNumber(id) {
    const text = $(id).value.trim();
    if (text === '') return null;
    const value = Number(text);
    return Number.isFinite(value) ? value : null;
  }

  function readChartSettings() {
    return {
      showPoints: $('showPoints').checked,
      showLines: $('showLines').checked,
      pointShape: $('pointShape').value,
      pointSize: Number($('pointSizeNumber').value),
      lineWidth: Number($('lineWidthNumber').value),
      showTrend: $('showTrend').checked,
      trendColor: $('trendColor').value,
      trendWidth: Number($('trendWidthNumber').value),
      trendDash: $('trendDash').value,
      showBounds: $('showBounds').checked,
      sdColor: $('sdColor').value,
      sdWidth: Number($('sdWidthNumber').value),
      sdDash: $('sdDash').value,
      vaiorColor: $('vaiorColor').value,
      vaiorWidth: Number($('vaiorWidthNumber').value),
      vaiorDash: $('vaiorDash').value,
      showPhaseBoundaries: $('showPhaseBoundaries').checked,
      boundaryColor: $('boundaryColor').value,
      boundaryWidth: Number($('boundaryWidthNumber').value),
      chartFont: fontMap[$('chartFont').value] || fontMap['noto-sans'],
      backgroundColor: $('backgroundColor').value,
      axisColor: $('axisColor').value,
      showGridLines: $('showGridLines').checked,
      gridColor: $('gridColor').value,
      yMin: optionalNumber('yAxisMin'),
      yMax: optionalNumber('yAxisMax'),
      yTick: optionalNumber('yAxisTick')
    };
  }

  function fmt(value, digits = Number($('visualDigits').value)) {
    if (value === null || value === undefined || !Number.isFinite(Number(value))) return '算出不能';
    return Number(value).toFixed(digits);
  }

  function fmtPercent(value, digits = Number($('visualDigits').value)) {
    if (value === null || !Number.isFinite(value)) return '算出不能';
    return `${(value * 100).toFixed(digits)}％`;
  }

  function fmtStat(value, digits) {
    return value === null || value === undefined || !Number.isFinite(Number(value)) ? '算出不能' : Number(value).toFixed(digits);
  }

  function pFmt(value, digits = Number($('statsDigits').value)) {
    if (value === null || value === undefined || !Number.isFinite(value)) return '算出不能';
    if (value < 0.001) return 'p＜0.001';
    return `p＝${value.toFixed(digits)}`;
  }

  function ciFmt(ci, digits = Number($('statsDigits').value)) {
    if (!ci) return '算出不能';
    return `${ci.lower.toFixed(digits)}～${ci.upper.toFixed(digits)}`;
  }

  function equation(slope, intercept, digits = Number($('visualDigits').value)) {
    if (![slope, intercept].every(Number.isFinite)) return '算出不能';
    return `Y ＝ ${slope.toFixed(digits)} × X ${intercept >= 0 ? '＋' : '－'} ${Math.abs(intercept).toFixed(digits)}`;
  }

  function table(headers, rows) {
    return `<div class="table-scroll"><table class="result-table"><thead><tr>${headers.map(header => `<th>${escapeHtml(header)}</th>`).join('')}</tr></thead><tbody>${rows.map(row => `<tr>${row.map(cell => `<td>${cell}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`;
  }

  function metric(label, value) {
    return `<div class="metric"><span>${escapeHtml(label)}</span><strong>${value}</strong></div>`;
  }

  function statusBadge(result) {
    if (!result || result.status === 'indeterminate') return `<span class="badge unknown">判定不能${result?.reason ? `：${escapeHtml(result.reason)}` : ''}</span>`;
    return result.status === 'effective' ? '<span class="badge good">効果あり</span>' : '<span class="badge bad">効果なし</span>';
  }

  function makeYAxis(data, settings) {
    const allY = data.phases.flatMap(phase => phase.y);
    if (!allY.length) return {};
    let min = settings.yMin;
    let max = settings.yMax;
    if (min !== null && max !== null && min >= max) throw new Error('縦軸の最小値は最大値より小さくしてください．');
    const dataMin = Math.min(...allY);
    const dataMax = Math.max(...allY);
    const span = Math.max(dataMax - dataMin, Math.abs(dataMax || 1) * 0.1, 1);
    if (min !== null && max === null) max = Math.max(dataMax + span * 0.08, min + span);
    if (max !== null && min === null) min = Math.min(dataMin - span * 0.08, max - span);
    const axis = {
      title: data.dataName,
      showline: true,
      linecolor: settings.axisColor,
      showgrid: settings.showGridLines,
      gridcolor: settings.gridColor,
      zeroline: false
    };
    if (min !== null || max !== null) axis.range = [min, max];
    if (settings.yTick !== null) {
      if (settings.yTick <= 0) throw new Error('縦軸の目盛間隔は0より大きい値にしてください．');
      axis.dtick = settings.yTick;
    }
    return axis;
  }

  function renderVisual() {
    hideMessage('visualValidation');
    if (!state.headers.length) return;
    try {
      setStatus('visualStatus', '計算中', 'loading');
      const data = getVisualDataset();
      const method = $('visualMethod').value;
      const directionNeeded = method === 'twoSD' || method === 'celeration' || method === 'vaior';
      $('visualDirectionWrap').classList.toggle('hidden', !directionNeeded);
      const direction = $('visualDirection').value;
      const settings = readChartSettings();
      const traces = [];
      const shapes = [];
      const mode = settings.showPoints && settings.showLines ? 'lines+markers' : settings.showPoints ? 'markers' : settings.showLines ? 'lines' : 'none';

      data.phases.forEach(phase => {
        if (!phase.points.length) return;
        traces.push({
          x: phase.x,
          y: phase.y,
          type: 'scatter',
          mode,
          name: phase.label,
          marker: { color: state.phaseColors[phase.label].point, size: settings.pointSize, symbol: settings.pointShape },
          line: { color: state.phaseColors[phase.label].line, width: settings.lineWidth },
          hovertemplate: `${escapeHtml(phase.label)}<br>${escapeHtml(data.timeName)}=%{x}<br>${escapeHtml(data.dataName)}=%{y}<extra></extra>`
        });
      });

      if (settings.showPhaseBoundaries) {
        for (let index = 0; index < data.phases.length - 1; index += 1) {
          const left = data.phases[index].x;
          const right = data.phases[index + 1].x;
          if (!left.length || !right.length) continue;
          const boundary = (left[left.length - 1] + right[0]) / 2;
          shapes.push({ type: 'line', x0: boundary, x1: boundary, y0: 0, y1: 1, yref: 'paper', line: { color: settings.boundaryColor, width: settings.boundaryWidth, dash: 'dash' } });
        }
      }

      let analysisHtml = '';
      if (method === 'ols' || method === 'theil') analysisHtml = renderPhaseTrendAnalysis(data, method, traces, settings);
      else if (method === 'twoSD') analysisHtml = renderTwoSD(data, direction, traces, settings);
      else if (method === 'celeration') analysisHtml = renderCelerationAB(data, direction, traces, settings);
      else analysisHtml = renderVaior(data, direction, traces, settings);

      const allStats = C.descriptiveStats(state.rows.map(row => row[data.dataCol]));
      const statRows = [['データ全体', fmt(allStats.mean), fmt(allStats.sd), fmt(allStats.median), fmt(allStats.max), fmt(allStats.min)]];
      data.phases.forEach(phase => {
        const stats = C.descriptiveStats(phase.y);
        statRows.push([escapeHtml(phase.label), fmt(stats.mean), fmt(stats.sd), fmt(stats.median), fmt(stats.max), fmt(stats.min)]);
      });
      $('visualResults').innerHTML = `<div class="result-block"><h3>記述統計量</h3>${table(['対象', '平均値', '標準偏差', '中央値', '最大値', '最小値'], statRows)}</div>${analysisHtml}`;
      $('chartCaption').textContent = `${data.timeName} × ${data.dataName}｜${data.phases.map(phase => phase.label).join(' → ')}`;

      const layout = {
        margin: { l: 70, r: 24, t: 34, b: 62 },
        paper_bgcolor: settings.backgroundColor,
        plot_bgcolor: settings.backgroundColor,
        font: { family: settings.chartFont, color: '#2b2b2b' },
        hovermode: 'closest',
        showlegend: true,
        legend: { orientation: 'h', y: 1.08, x: 0 },
        xaxis: { title: data.timeName, showline: true, linecolor: settings.axisColor, showgrid: settings.showGridLines, gridcolor: settings.gridColor, zeroline: false },
        yaxis: makeYAxis(data, settings),
        shapes
      };
      Plotly.react('chart', traces, layout, { responsive: true, displayModeBar: false, displaylogo: false, scrollZoom: false });
      const chartEl = $('chart');
      state.lastChartSize = { width: Math.max(320, Math.round(chartEl.clientWidth)), height: Math.max(240, Math.round(chartEl.clientHeight)) };
      setStatus('visualStatus', '計算完了', 'success');
    } catch (error) {
      $('chart').innerHTML = '<div class="empty-state">グラフを表示できません．</div>';
      $('visualResults').innerHTML = `<div class="message error">${escapeHtml(error.message)}</div>`;
      showMessage('visualValidation', error.message);
      setStatus('visualStatus', '計算不能', 'error');
    }
  }

  function renderPhaseTrendAnalysis(data, method, traces, settings) {
    const rows = [];
    const notices = [];
    data.phases.forEach(phase => {
      try {
        if (phase.x.length < 2) throw new Error('有効データが2点未満です．');
        const result = method === 'ols' ? C.linearRegression(phase.x, phase.y) : C.theilSen(phase.x, phase.y);
        if (settings.showTrend) {
          const xMin = Math.min(...phase.x);
          const xMax = Math.max(...phase.x);
          traces.push({ x: [xMin, xMax], y: [result.predict(xMin), result.predict(xMax)], type: 'scatter', mode: 'lines', showlegend: false, line: { color: settings.trendColor, width: settings.trendWidth, dash: settings.trendDash }, hoverinfo: 'skip' });
        }
        rows.push([escapeHtml(phase.label), escapeHtml(equation(result.slope, result.intercept)), fmt(result.slope), fmt(result.intercept)]);
        if (method === 'theil' && result.excludedSameX > 0) notices.push(`${phase.label}：同一Xの組み合わせを${result.excludedSameX}ペア除外`);
      } catch (error) {
        const count = 4;
        rows.push([escapeHtml(phase.label), `<span class="badge unknown">${escapeHtml(error.message)}</span>`, ...Array(count - 2).fill('—')]);
      }
    });
    const title = method === 'ols' ? 'Trend（最小二乗法）' : 'Trend（Theil–Sen法）';
    const body = table(['フェーズ', '回帰式', 'Slope', 'Intercept'], rows);
    const note = notices.length ? `<p class="result-note">${notices.map(escapeHtml).join('／')}</p>` : '';
    return `<div class="result-block"><h3>${title}</h3>${body}${note}</div>`;
  }

  function phasePairsWithData(data) {
    return C.pairBaselineIntervention(data.phases).map(pair => ({ baseline: pair.baseline, intervention: pair.intervention }));
  }

  function renderTwoSD(data, direction, traces, settings) {
    const pairs = phasePairsWithData(data);
    if (!pairs.length) throw new Error('A期の直後にB期が続く組み合わせがありません．');
    const rows = [];
    pairs.forEach(({ baseline, intervention }) => {
      try {
        const result = C.twoSDAnalysis(baseline.x, baseline.y, intervention.x, intervention.y, direction);
        const xAll = [...baseline.x, ...intervention.x];
        if (settings.showTrend) traces.push({ x: xAll, y: xAll.map(() => result.baselineMean), type: 'scatter', mode: 'lines', showlegend: false, line: { color: settings.trendColor, width: settings.trendWidth, dash: settings.trendDash }, hoverinfo: 'skip' });
        if (settings.showBounds) {
          traces.push({ x: xAll, y: xAll.map(() => result.upper), type: 'scatter', mode: 'lines', showlegend: false, line: { color: settings.sdColor, width: settings.sdWidth, dash: settings.sdDash }, hoverinfo: 'skip' });
          traces.push({ x: xAll, y: xAll.map(() => result.lower), type: 'scatter', mode: 'lines', showlegend: false, line: { color: settings.sdColor, width: settings.sdWidth, dash: settings.sdDash }, hoverinfo: 'skip' });
        }
        rows.push([`${escapeHtml(baseline.label)} → ${escapeHtml(intervention.label)}`, fmt(result.baselineMean), fmt(result.sd), fmt(result.twoSD), fmt(result.upper), fmt(result.lower), `${result.improvedCount}/${result.interventionN}（${fmtPercent(result.improvedRate)}）`]);
      } catch (error) {
        rows.push([`${escapeHtml(baseline.label)} → ${escapeHtml(intervention.label)}`, `<span class="badge unknown">${escapeHtml(error.message)}</span>`, '—', '—', '—', '—', '—']);
      }
    });
    return `<div class="result-block"><h3>2SD法</h3>${table(['比較', 'A期平均', 'SD', '2SD', '上限', '下限', '改善側境界を超えたB期データ'], rows)}</div>`;
  }

  function renderCelerationAB(data, direction, traces, settings) {
    const pairs = phasePairsWithData(data);
    if (!pairs.length) throw new Error('A期の直後にB期が続く組み合わせがありません．');
    const rows = [];
    pairs.forEach(({ baseline, intervention }) => {
      try {
        const result = C.celerationABAnalysis(baseline.x, baseline.y, intervention.x, intervention.y, direction);
        const xAll = [...baseline.x, ...intervention.x];
        if (settings.showTrend) traces.push({ x: xAll, y: xAll.map(result.predict), type: 'scatter', mode: 'lines', showlegend: false, line: { color: settings.trendColor, width: settings.trendWidth, dash: settings.trendDash }, hoverinfo: 'skip' });
        const judgment = result.effective ? '<span class="badge good">効果あり</span>' : '<span class="badge bad">効果ありとは判定されない</span>';
        rows.push([
          `${escapeHtml(baseline.label)} → ${escapeHtml(intervention.label)}`,
          `(${fmt(result.firstPoint.x)}，${fmt(result.firstPoint.y)})`,
          `(${fmt(result.secondPoint.x)}，${fmt(result.secondPoint.y)})`,
          escapeHtml(equation(result.slope, result.intercept)),
          `${result.improvedCount}/${result.testN}`,
          String(result.oppositeCount),
          String(result.tiedCount),
          pFmt(result.pValue, Number($('visualDigits').value)),
          judgment
        ]);
      } catch (error) {
        rows.push([`${escapeHtml(baseline.label)} → ${escapeHtml(intervention.label)}`, `<span class="badge unknown">${escapeHtml(error.message)}</span>`, '—', '—', '—', '—', '—', '—', '—']);
      }
    });
    return `<div class="result-block"><h3>Celeration Line分析</h3>${table(['比較', '前半中央値座標', '後半中央値座標', '回帰式', '改善側／検定対象', '反対側', '同値', '二項検定（片側）', '判定'], rows)}<p class="result-note">二項検定は，介入期の各実測値をベースライン期から延長したCeleration Lineの予測値と比較し，改善方向に位置する点が偶然の50％を上回るかを正確片側検定で確認します．線上の点は検定から除外します．</p></div>`;
  }

  function renderVaior(data, direction, traces, settings) {
    const pairs = phasePairsWithData(data);
    if (!pairs.length) throw new Error('A期の直後にB期が続く組み合わせがありません．');
    const blocks = [];
    pairs.forEach(({ baseline, intervention }) => {
      try {
        const result = C.vaiorAnalysis(baseline.x, baseline.y, intervention.x, intervention.y, direction);
        const xAll = [...baseline.x, ...intervention.x];
        if (settings.showTrend) traces.push({ x: xAll, y: xAll.map(result.predict), type: 'scatter', mode: 'lines', showlegend: false, line: { color: settings.trendColor, width: settings.trendWidth, dash: settings.trendDash }, hoverinfo: 'skip' });
        if (settings.showBounds) {
          traces.push({ x: xAll, y: xAll.map(x => result.predict(x) + result.mad), type: 'scatter', mode: 'lines', showlegend: false, line: { color: settings.vaiorColor, width: settings.vaiorWidth, dash: settings.vaiorDash }, hoverinfo: 'skip' });
          traces.push({ x: xAll, y: xAll.map(x => result.predict(x) - result.mad), type: 'scatter', mode: 'lines', showlegend: false, line: { color: settings.vaiorColor, width: settings.vaiorWidth, dash: settings.vaiorDash }, hoverinfo: 'skip' });
        }
        const criterion = result.criterion > 1 ? `${fmtPercent(result.criterion)}（100％を超えるため到達不能）` : fmtPercent(result.criterion);
        blocks.push(`<section class="compact-result-section"><h4>${escapeHtml(baseline.label)} → ${escapeHtml(intervention.label)}</h4><div class="result-grid">${metric('Theil–Sen回帰式', escapeHtml(equation(result.slope, result.intercept)))}${metric('変動幅（絶対偏差中央値）', fmt(result.mad))}${metric('A期範囲外', `${result.baselineOutsideCount}/${result.baselineN}（${fmtPercent(result.baselineOutsideRate)}）`)}${metric('判定基準値', criterion)}${metric('B期改善側範囲外', `${result.interventionImprovedCount}/${result.interventionN}（${fmtPercent(result.interventionImprovedRate)}）`)}${metric('即時効果', statusBadge(result.immediate))}${metric('漸進的・遅延的効果', statusBadge(result.delayed))}${metric('介入期全体', statusBadge(result.overall))}</div></section>`);
      } catch (error) {
        blocks.push(`<div class="message error">${escapeHtml(baseline.label)} → ${escapeHtml(intervention.label)}：${escapeHtml(error.message)}</div>`);
      }
    });
    return `<div class="result-block"><h3>VAIOR（Visual Aid Implying an Objective Rule）</h3>${blocks.join('')}</div>`;
  }

  function renderSlots() {
    const el = $('dataSlots');
    el.innerHTML = state.slots.map((slot, index) => `<article class="slot-card" data-slot="${index}"><h3>データ${index + 1}</h3><label class="field">データ名<input class="slot-name" value="${escapeHtml(slot.name)}"></label><label class="field">セル範囲<input class="slot-range" value="${escapeHtml(slot.range)}" placeholder="例：C2:C6"></label><div class="slot-meta">${slot.range ? `<strong>${escapeHtml(slot.range)}</strong><div>数値データ：${slot.values.length - slot.excluded}点</div><div>除外：${slot.excluded}点</div><div class="preview-values">${slot.values.slice(0, 10).map(value => escapeHtml(value)).join(', ') || '数値データなし'}</div>` : '未登録'}</div><div class="slot-actions"><button class="button secondary register-slot" type="button">現在の選択を登録</button><button class="button danger clear-slot" type="button">解除</button></div></article>`).join('');
    el.querySelectorAll('.slot-name').forEach((input, index) => input.addEventListener('input', () => {
      state.slots[index].name = input.value || `データ${index + 1}`;
      renderComparisonChecks();
    }));
    el.querySelectorAll('.slot-range').forEach((input, index) => input.addEventListener('change', () => registerSlot(index, input.value)));
    el.querySelectorAll('.register-slot').forEach((button, index) => button.addEventListener('click', () => registerSlot(index, state.selection ? selectionToA1(state.selection) : button.closest('.slot-card').querySelector('.slot-range').value)));
    el.querySelectorAll('.clear-slot').forEach((button, index) => button.addEventListener('click', () => {
      state.slots[index] = { name: ['A1', 'B1', 'A2', 'B2'][index], range: '', values: [], excluded: 0 };
      renderSlots();
    }));
    renderComparisonChecks();
  }

  function suggestedPhaseName(parsed) {
    const phaseCol = Number($('phaseColumn').value);
    if (!Number.isInteger(phaseCol)) return null;
    const phases = C.splitConsecutivePhases(state.rows.map(row => row[phaseCol]));
    return phases.find(phase => parsed.startRow >= phase.start && parsed.endRow <= phase.end)?.label || null;
  }

  function registerSlot(index, rangeText) {
    hideMessage('statsValidation');
    try {
      if (!state.headers.length) throw new Error('先にデータを読み込んでください．');
      const parsed = C.parseA1Range(rangeText);
      if (parsed.startCol !== parsed.endCol) throw new Error('統計解析の各データ欄には，1列の範囲を指定してください．');
      if (parsed.endCol >= state.headers.length || parsed.endRow >= state.rows.length) throw new Error('指定範囲がデータ表の範囲を超えています．');
      const rawValues = [];
      for (let row = parsed.startRow; row <= parsed.endRow; row += 1) rawValues.push(state.rows[row][parsed.startCol]);
      const cleaned = C.cleanNumeric(rawValues);
      if (!cleaned.values.length) throw new Error('指定範囲に数値データがありません．');
      state.slots[index] = { name: suggestedPhaseName(parsed) || state.slots[index].name || `データ${index + 1}`, range: parsed.normalized, values: rawValues, excluded: cleaned.excludedCount };
      renderSlots();
      setStatus('statsStatus', `${state.slots.filter(slot => slot.range).length}データ登録済み`, 'success');
    } catch (error) {
      showMessage('statsValidation', error.message);
      setStatus('statsStatus', '登録エラー', 'error');
    }
  }

  function renderComparisonChecks() {
    const el = $('comparisonChecks');
    const selected = [...el.querySelectorAll('input:checked')].map(input => Number(input.value));
    el.innerHTML = state.slots.map((slot, index) => `<label class="comparison-item"><input type="checkbox" value="${index}" ${selected.includes(index) ? 'checked' : ''} ${slot.range ? '' : 'disabled'}><strong>${escapeHtml(slot.name)}</strong><span>${slot.range ? escapeHtml(slot.range) : '未登録'}</span></label>`).join('');
  }

  function renderStatsDefinition() {
    const method = $('statsMethod').value;
    const tau = method === 'tauU';
    $('tauCorrectionWrap').classList.toggle('hidden', !tau);
    const notes = {
      pnd: 'PND：A期最大値をB期データが厳密に上回った割合です．同値は含めません．',
      nap: 'NAP：A×Bの全ペアを比較し，B>Aを1，同値を0.5，B<Aを0として平均します．',
      tauU: $('tauCorrection').checked
        ? 'Tau-U：A–B間のSからA期内の傾向Sを差し引き，A×Bのペア数で割ります．'
        : 'Tau：A×Bの全ペアについて，B>Aを＋1，同値を0，B<Aを−1として平均します．'
    };
    $('statsDefinitionNote').textContent = notes[method];
  }

  function compactAnalysisInfo(result, baseline, intervention) {
    const excludedA = result.excludedBaseline ?? C.cleanNumeric(baseline.values).excludedCount;
    const excludedB = result.excludedIntervention ?? C.cleanNumeric(intervention.values).excludedCount;
    const nA = baseline.values.length - excludedA;
    const nB = intervention.values.length - excludedB;
    const excluded = excludedA + excludedB > 0 ? `／除外：A期 ${excludedA}，B期 ${excludedB}` : '';
    const inference = result.inferenceMethod ? `<br>推論法：${escapeHtml(result.inferenceMethod)}` : '';
    const caution = result.inferenceCaution ? `<br>${escapeHtml(result.inferenceCaution)}` : '';
    return `<p class="result-note">使用データ：${escapeHtml(baseline.name)} n=${nA}，${escapeHtml(intervention.name)} n=${nB}${excluded}${inference}${caution}<br>統計解析はB期が高い方向を正方向として計算します．目安は便宜的な区分であり，グラフの水準・傾向・ばらつきと併せて解釈してください．</p>`;
  }

  function calculateStats() {
    hideMessage('statsValidation');
    try {
      setStatus('statsStatus', '計算中', 'loading');
      const selected = [...document.querySelectorAll('#comparisonChecks input:checked')].map(input => Number(input.value));
      if (selected.length !== 2) throw new Error('比較するデータを2つだけ選択してください．上側のデータをベースライン期，下側のデータを介入期として扱います．');
      selected.sort((a, b) => a - b);
      const baseline = state.slots[selected[0]];
      const intervention = state.slots[selected[1]];
      const method = $('statsMethod').value;
      const digits = Number($('statsDigits').value);
      const comparison = `${baseline.name}–${intervention.name}`;
      let result;
      let html;

      if (method === 'pnd') {
        result = C.pnd(baseline.values, intervention.values, 'higher');
        html = table(['比較', 'A期最大値', '基準を超えたB期データ', 'PND', '目安'], [[escapeHtml(comparison), fmtStat(result.bestBaseline, digits), `${result.improvedCount}/${result.interventionN}`, `${fmtStat(result.pnd * 100, digits)}％`, escapeHtml(result.interpretation)]]);
      } else if (method === 'nap') {
        result = C.nap(baseline.values, intervention.values, 'higher');
        html = table(['比較', 'NAP', '90％CI', '95％CI', 'P値', '目安', 'ペア内訳'], [[escapeHtml(comparison), fmtStat(result.nap, digits), ciFmt(result.ci90, digits), ciFmt(result.ci95, digits), pFmt(result.pValue, digits), escapeHtml(result.interpretation), `B>A ${result.improved}／同値 ${result.tied}／B<A ${result.worsened}`]]);
      } else {
        const corrected = $('tauCorrection').checked;
        if (corrected) {
          result = C.tauU(baseline.values, intervention.values, 'higher', { correctBaseline: true });
          html = table(['比較', 'Tau-U', '90％CI', '95％CI', 'P値', '目安'], [[escapeHtml(comparison), fmtStat(result.tauU, digits), ciFmt(result.ci90, digits), ciFmt(result.ci95, digits), pFmt(result.pValue, digits), escapeHtml(result.interpretation)]]);
          html += `<details class="calculation-details"><summary>計算過程を表示</summary><div class="result-grid">${metric('補正前のTau', fmtStat(result.tauAB, digits))}${metric('ベースライン傾向 S_A', fmtStat(result.trend.score, digits))}${metric('計算式', `(${fmtStat(result.cross.score, digits)} − ${fmtStat(result.trend.score, digits)}) ÷ ${result.denominator}`)}</div></details>`;
        } else {
          result = C.tau(baseline.values, intervention.values, 'higher');
          html = table(['比較', 'Tau', '90％CI', '95％CI', 'P値', '目安', 'ペア内訳'], [[escapeHtml(comparison), fmtStat(result.tau, digits), ciFmt(result.ci90, digits), ciFmt(result.ci95, digits), pFmt(result.pValue, digits), escapeHtml(result.interpretation), `B>A ${result.improved}／同値 ${result.tied}／B<A ${result.worsened}`]]);
        }
      }
      $('statsResults').innerHTML = html + compactAnalysisInfo(result, baseline, intervention);
      setStatus('statsStatus', '計算完了', 'success');
    } catch (error) {
      showMessage('statsValidation', error.message);
      setStatus('statsStatus', '計算不能', 'error');
    }
  }

  async function handleFile(file) {
    try {
      setStatus('globalStatus', '読込中', 'loading');
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: 'array', cellDates: false });
      state.workbook = workbook;
      state.baseFileName = file.name;
      $('sheetSelect').innerHTML = workbook.SheetNames.map(name => `<option value="${escapeHtml(name)}">${escapeHtml(name)}</option>`).join('');
      $('sheetControl').classList.toggle('hidden', workbook.SheetNames.length <= 1);
      loadSheet(workbook.SheetNames[0], state.baseFileName);
    } catch (error) {
      setStatus('globalStatus', error.message, 'error');
    }
  }

  function loadSheet(name, fileName = state.baseFileName || state.fileName) {
    if (!state.workbook) return;
    const matrix = XLSX.utils.sheet_to_json(state.workbook.Sheets[name], { header: 1, defval: null, raw: true });
    loadMatrix(matrix, `${fileName} / ${name}`);
  }

  function resetChartSettings() {
    const values = {
      showPoints: true, showLines: true, pointShape: 'circle', pointSizeRange: 8, pointSizeNumber: 8,
      lineWidthRange: 2, lineWidthNumber: 2, showTrend: true, trendColor: '#2b2b2b', trendWidthRange: 2,
      trendWidthNumber: 2, trendDash: 'solid', showBounds: true, sdColor: '#5faf68', sdWidthRange: 1.5,
      sdWidthNumber: 1.5, sdDash: 'dash', vaiorColor: '#93c47d', vaiorWidthRange: 1.5, vaiorWidthNumber: 1.5,
      vaiorDash: 'dot', showPhaseBoundaries: true, boundaryColor: '#aab8ac', boundaryWidthRange: 1,
      boundaryWidthNumber: 1, chartFont: 'noto-sans', backgroundColor: '#ffffff', axisColor: '#415047',
      showGridLines: true, gridColor: '#dfeadf', yAxisMin: '', yAxisMax: '', yAxisTick: ''
    };
    Object.entries(values).forEach(([id, value]) => {
      const el = $(id);
      if (!el) return;
      if (typeof value === 'boolean') el.checked = value;
      else el.value = value;
    });
    state.phaseColors = {};
    renderVisual();
  }

  function bindDual(rangeId, numberId) {
    const range = $(rangeId);
    const number = $(numberId);
    range.addEventListener('input', () => { number.value = range.value; renderVisual(); });
    number.addEventListener('input', () => { range.value = number.value; renderVisual(); });
  }

  function openDownloadDialog() {
    if (!state.headers.length || !$('chart').data) {
      showMessage('visualValidation', '先にグラフを表示してください．');
      return;
    }
    const width = Math.max(320, state.lastChartSize.width || 1000);
    const height = Math.max(240, state.lastChartSize.height || 620);
    $('downloadWidth').value = width;
    $('downloadHeight').value = height;
    $('downloadAspect').checked = true;
    $('downloadDialog').dataset.aspect = String(width / height);
    if (typeof $('downloadDialog').showModal === 'function') $('downloadDialog').showModal();
    else $('downloadDialog').setAttribute('open', '');
  }

  function closeDownloadDialog() {
    if (typeof $('downloadDialog').close === 'function') $('downloadDialog').close();
    else $('downloadDialog').removeAttribute('open');
  }

  function syncDownloadDimension(changed) {
    if (!$('downloadAspect').checked) return;
    const aspect = Number($('downloadDialog').dataset.aspect) || 1.6;
    if (changed === 'width') $('downloadHeight').value = Math.max(1, Math.round(Number($('downloadWidth').value) / aspect));
    else $('downloadWidth').value = Math.max(1, Math.round(Number($('downloadHeight').value) * aspect));
  }

  async function downloadChart() {
    const format = $('downloadFormat').value;
    const width = Number($('downloadWidth').value);
    const height = Number($('downloadHeight').value);
    if (!Number.isFinite(width) || !Number.isFinite(height) || width < 100 || height < 100) {
      $('downloadError').textContent = '幅と高さは100以上の数値で入力してください．';
      $('downloadError').classList.remove('hidden');
      return;
    }
    $('downloadError').classList.add('hidden');
    $('confirmDownloadButton').disabled = true;
    $('confirmDownloadButton').textContent = '作成中…';
    try {
      const dataUrl = await Plotly.toImage('chart', { format, width: Math.round(width), height: Math.round(height), scale: 1 });
      const link = document.createElement('a');
      link.href = dataUrl;
      link.download = `SCD_graph.${format === 'jpeg' ? 'jpg' : format}`;
      document.body.appendChild(link);
      link.click();
      link.remove();
      closeDownloadDialog();
    } catch (error) {
      $('downloadError').textContent = `画像を作成できませんでした：${error.message}`;
      $('downloadError').classList.remove('hidden');
    } finally {
      $('confirmDownloadButton').disabled = false;
      $('confirmDownloadButton').textContent = 'ダウンロード';
    }
  }

  function initializeEvents() {
    $('fileInput').addEventListener('change', event => { const file = event.target.files[0]; if (file) handleFile(file); });
    $('sheetSelect').addEventListener('change', event => loadSheet(event.target.value));
    $('applyRangeButton').addEventListener('click', applyDirectRange);
    $('directRange').addEventListener('keydown', event => { if (event.key === 'Enter') applyDirectRange(); });
    document.addEventListener('mouseup', () => { state.dragging = false; });

    document.querySelectorAll('.tab').forEach(tab => tab.addEventListener('click', () => {
      document.querySelectorAll('.tab').forEach(item => item.classList.toggle('active', item === tab));
      document.querySelectorAll('.analysis-panel').forEach(panel => panel.classList.toggle('active', panel.id === tab.dataset.tab));
      const statsActive = tab.dataset.tab === 'statsPanel';
      $('inlineStatsRegistration').classList.toggle('hidden', !statsActive);
      document.querySelector('.data-workflow-card').classList.toggle('stats-mode', statsActive);
      if (tab.dataset.tab === 'visualPanel' && state.headers.length) renderVisual();
    }));

    ['timeColumn', 'phaseColumn', 'dataColumn', 'visualMethod', 'visualDirection', 'visualDigits'].forEach(id => $(id).addEventListener('change', renderVisual));
    ['showPoints', 'showLines', 'pointShape', 'showTrend', 'trendColor', 'trendDash', 'showBounds', 'sdColor', 'sdDash', 'vaiorColor', 'vaiorDash', 'showPhaseBoundaries', 'boundaryColor', 'chartFont', 'backgroundColor', 'axisColor', 'showGridLines', 'gridColor', 'yAxisMin', 'yAxisMax', 'yAxisTick'].forEach(id => $(id).addEventListener('input', renderVisual));
    bindDual('pointSizeRange', 'pointSizeNumber');
    bindDual('lineWidthRange', 'lineWidthNumber');
    bindDual('trendWidthRange', 'trendWidthNumber');
    bindDual('sdWidthRange', 'sdWidthNumber');
    bindDual('vaiorWidthRange', 'vaiorWidthNumber');
    bindDual('boundaryWidthRange', 'boundaryWidthNumber');
    $('resetChartButton').addEventListener('click', resetChartSettings);

    $('customizationToggle').addEventListener('change', event => {
      $('customizationBody').classList.toggle('hidden', !event.target.checked);
      event.target.closest('.collapse-toggle').querySelector('span:first-child').textContent = event.target.checked ? '設定を閉じる' : '設定を開く';
    });

    $('statsMethod').addEventListener('change', renderStatsDefinition);
    $('tauCorrection').addEventListener('change', renderStatsDefinition);
    $('calculateStatsButton').addEventListener('click', calculateStats);

    $('openDownloadDialogButton').addEventListener('click', openDownloadDialog);
    $('cancelDownloadButton').addEventListener('click', closeDownloadDialog);
    $('cancelDownloadFooterButton').addEventListener('click', closeDownloadDialog);
    $('confirmDownloadButton').addEventListener('click', downloadChart);
    $('downloadWidth').addEventListener('input', () => syncDownloadDimension('width'));
    $('downloadHeight').addEventListener('input', () => syncDownloadDimension('height'));
    $('downloadDialog').addEventListener('click', event => {
      if (event.target === $('downloadDialog')) closeDownloadDialog();
    });
  }

  initializeEvents();
  renderSlots();
  renderStatsDefinition();
})();
