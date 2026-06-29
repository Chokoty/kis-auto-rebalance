/**
 * AI 시장 분석 및 비중 제안 (Gemini API + Google Search Grounding)
 * 컬럼 레이아웃: A=종목코드, B=종목명, C=기준비율(고정), D=운용비율(수정가능), E=유형
 */

// ─────────────────────────────────────────────────────────────
// 시스템 프롬프트
// ─────────────────────────────────────────────────────────────

/**
 * 시스템 기본 프롬프트 반환 (커스텀 저장된 것이 있으면 그것을 반환)
 */
function getSystemPrompt() {
  const customPrompt = PropertiesService.getUserProperties().getProperty('CUSTOM_AI_PROMPT');
  if (customPrompt) return customPrompt;

  return `너는 10년차 프로 자산 배분 전략가이자 글로벌 시장 분석가야.
아래 포트폴리오 데이터와 Google Search로 최신 뉴스·시장 동향을 종합해서 비중 조정을 제안해줘.

[분석 방법]:
- 현재 포트폴리오 비중과 목표 비중의 차이 확인
- Google Search로 각 자산군(채권, 금, 달러, 국내주식, 해외주식 등)의 최신 뉴스·시장 상황 파악
- 수익률 갭(목표 미달/초과)과 시장 상황을 종합해 전체 배분 방향 결정

[포트폴리오 원칙]:
- 단일 종목 최대 비중: 40% 이하 (분산 원칙 유지)
- 합산 반드시 100% 유지
- 점진적 조정 선호 (급격한 변경 지양)
- **기준비중 drift 제한 (최우선)**: 각 종목 제안 비중은 기준비중 ±15%p 이내로 제한. 현금은 기준비중의 3배 초과 금지 (예: 기준 5% → 최대 15%)
- 분석 결과에 반드시 포함: 각 종목의 현재 운용비중이 기준비중 대비 얼마나 drift(이탈)되었는지
- **용어 구분**:
    • **현금**: 포트폴리오 비중 카테고리 (JSON에 code "CASH", 본문에 "현금")
    • **예수금**: 실제 계좌 금액 — 비중 조정 제안에는 사용 금지

[출력 형식]:
**중요: 응답은 아래 양식을 정확히 지켜서 마크다운으로 깔끔하게 출력하고, 문서는 가급적 짧고 굵게 써!**
마지막에는 반드시 지정된 형식의 JSON 데이터를 포함해야 해 (전체 합산 100%).

[출력 양식]:
### 🌍 시장 요약
- 시장 뷰 요약 1줄
- 시장 뷰 요약 1줄

### 📐 기준비중 대비 drift 현황
- **[종목명]**: 기준 X% → 현재목표 Y% (drift: ±Z%p)

### 🎯 비중 조정 제안
- **[종목명]**: 현재 X% ➡️ **조정 Y%** | 근거 1줄

[ALLOCATION_START] {"ratios": [{"code": "종목코드(현금은 CASH)", "name": "종목명", "ratio": 숫자, "rationale": "간략한 이유"}], "summary": "전체 요약"} [ALLOCATION_END]

"이것은 투자 조언이 아닙니다. 스스로 판단하세요."`;
}

/**
 * 프롬프트 저장 (사용자 속성)
 */
function saveSystemPrompt(prompt) {
  PropertiesService.getUserProperties().setProperty('CUSTOM_AI_PROMPT', prompt);
}

/**
 * 프롬프트 초기화
 */
function resetSystemPrompt() {
  PropertiesService.getUserProperties().deleteProperty('CUSTOM_AI_PROMPT');
}

/**
 * 사용자 프롬프트 설정창 열기
 */
function openAIPromptSettings() {
  const currentPrompt = getSystemPrompt();
  const html = `
    <html>
      <head>
        <style>
          body { font-family: 'Malgun Gothic', sans-serif; padding: 20px; color: #3c4043; }
          textarea { width: 100%; height: 400px; font-family: monospace; font-size: 13px; padding: 12px; border: 1px solid #dadce0; border-radius: 4px; line-height: 1.5; }
          .btn-container { display: flex; gap: 10px; margin-top: 20px; }
          .btn { flex: 1; padding: 12px; border: none; border-radius: 4px; cursor: pointer; font-weight: bold; }
          .btn-save { background: #4285F4; color: white; }
          .btn-reset { background: white; border: 1px solid #dadce0; color: #5f6368; }
          .hint { font-size: 12px; color: #5f6368; margin-bottom: 12px; }
        </style>
      </head>
      <body>
        <h3>🤖 AI 시스템 프롬프트 설정</h3>
        <p class="hint">분석 및 비중 제안 시 AI에게 전달되는 지침입니다.</p>
        <textarea id="promptArea">${currentPrompt.replace(/</g,'&lt;').replace(/>/g,'&gt;')}</textarea>
        <div class="btn-container">
          <button class="btn btn-reset" onclick="reset()">초기화</button>
          <button class="btn btn-save" onclick="save()">저장하기</button>
        </div>
        <script>
          function save() {
            const prompt = document.getElementById('promptArea').value;
            google.script.run
              .withSuccessHandler(() => { alert('저장되었습니다.'); google.script.host.close(); })
              .saveSystemPrompt(prompt);
          }
          function reset() {
            if (!confirm('기본 프롬프트로 초기화하시겠습니까?')) return;
            google.script.run
              .withSuccessHandler(() => { alert('초기화되었습니다.'); google.script.host.close(); })
              .resetSystemPrompt();
          }
        <\/script>
      </body>
    </html>
  `;
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(600).setHeight(620),
    ' '
  );
}

// ─────────────────────────────────────────────────────────────
// 포트폴리오 헬퍼
// ─────────────────────────────────────────────────────────────

/**
 * 포트폴리오설정 C열(기준비율)이 비어 있으면 D열(운용비율)로 초기화.
 */
function ensureInitialRatios() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('📋 포트폴리오설정');
  if (!sheet) return false;
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return false;
  // col: 0=종목코드, 1=종목명, 2=기준비율(C), 3=운용비율(D), 4=유형(E)
  const data = sheet.getRange(3, 1, lastRow - 2, 5).getValues();
  const needsInit = data.some(row => !(typeof row[2] === 'number' && row[2] > 0));
  if (!needsInit) return false;
  const updates = data.map(row => [
    (typeof row[2] === 'number' && row[2] > 0) ? row[2] : row[3]
  ]);
  sheet.getRange(3, 3, updates.length, 1).setValues(updates);
  return true;
}

/**
 * 포트폴리오설정 시트에서 종목코드 → 기준비율 맵을 반환
 */
function getInitialRatioMap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('📋 포트폴리오설정');
  const map = {};
  if (!sheet) return map;
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return map;
  // col: 0=종목코드, 1=종목명, 2=기준비율(C), 3=운용비율(D), 4=유형(E)
  const data = sheet.getRange(3, 1, lastRow - 2, 5).getValues();
  data.forEach(row => {
    const code = String(row[0]).trim();
    const name = String(row[1]).trim();
    const key = code || 'CASH';
    const val = (typeof row[2] === 'number' && row[2] > 0) ? row[2] : row[3];
    map[key] = val;
    if (name === '현금') map['CASH'] = val;
  });
  return map;
}

// ─────────────────────────────────────────────────────────────
// 대시보드 상태 요약
// ─────────────────────────────────────────────────────────────

/**
 * 미리보기용 경량 포트폴리오 현황 (시트 값만 읽음)
 */
function getDashboardBasicState() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dash = ss.getSheetByName('📊 대시보드');
  if (!dash) return '대시보드 데이터를 찾을 수 없습니다.';

  const totalEval    = dash.getRange('B3').getValue();
  const cash         = dash.getRange('B4').getValue();
  const lastUpdate   = dash.getRange('B6').getDisplayValue();
  const currentReturn = parseFloat((dash.getRange('H5').getValue() || '0').toString().replace('%', ''));
  const targetYield  = parseFloat((dash.getRange('H6').getValue() || 0)) * 100;
  const gap = targetYield - currentReturn;

  const lastRow = dash.getLastRow();
  if (lastRow < 9) return `총 평가액: ${totalEval}, 업데이트: ${lastUpdate}`;

  const data = dash.getRange(9, 1, lastRow - 8, 10).getValues();
  const initialRatioMap = getInitialRatioMap();

  let report = `[포트폴리오 현황 (기준: ${lastUpdate})]\n`;
  report += `총 평가액: ${totalEval.toLocaleString()}원 | 예수금: ${cash.toLocaleString()}원\n`;
  report += `현재 수익률: ${currentReturn >= 0 ? '+' : ''}${currentReturn.toFixed(2)}% | 목표: ${targetYield.toFixed(1)}% | 갭: ${gap >= 0 ? '+' : ''}${gap.toFixed(2)}%p\n`;
  report += `\n[종목별 현황] (현재비중 → 운용비중 | 기준비중)\n`;

  data.forEach(row => {
    const code = String(row[0]).trim();
    const name = row[1];
    if (!code) return;
    const evalAmount   = row[5];
    const currentRatio = row[6];
    const targetRatio  = row[7];
    const diff         = row[9];
    const initialRatio = initialRatioMap[code];
    const initialStr   = initialRatio != null ? ` | 기준: ${initialRatio}%` : '';
    report += `▶ ${name} (${code}): 현재 ${currentRatio} → 목표 ${targetRatio}${initialStr} (차이 ${diff}) | ${typeof evalAmount === 'number' ? evalAmount.toLocaleString() : evalAmount}원\n`;
  });

  return report;
}

/**
 * AI 분석용 포트폴리오 현황 요약 (수익률 갭 포함)
 */
function getDashboardStateForAI() {
  return getDashboardBasicState();
}

// ─────────────────────────────────────────────────────────────
// AI 분석 실행
// ─────────────────────────────────────────────────────────────

/**
 * AI 시장 분석 및 비중 제안 실행 (메뉴 진입점 — 바로 실행)
 * @param {string} [customPrompt] — 커스텀 프롬프트 (없으면 저장된 프롬프트 사용)
 */
function runAIBriefing(customPrompt) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getConfig();

  if (!config.geminiApiKey) {
    SpreadsheetApp.getUi().alert('Gemini API Key가 필요한 기능입니다.\n메뉴 > ⚙️ 설정 및 관리 > 🛡️ API 키 보안 설정에서 입력해 주세요.');
    return;
  }

  // 기준비율(C열) 미설정 시 현재 운용비율로 자동 초기화
  const wasInitialized = ensureInitialRatios();
  if (wasInitialized) {
    ss.toast('📋 포트폴리오설정 C열(기준비율)이 비어 있어 현재 운용비율로 초기화했습니다.', '⚠️ 기준비율 초기화', 8);
    Utilities.sleep(2000);
  }

  // 당일 추천 이미 존재하면 교체 여부 확인
  if (hasTodayRecommendation()) {
    const ui = SpreadsheetApp.getUi();
    const res = ui.alert('⚠️ 오늘 이미 AI 추천이 있습니다', '다시 실행하면 오늘 추천 내역을 교체합니다.\n계속하시겠습니까?', ui.ButtonSet.YES_NO);
    if (res !== ui.Button.YES) return;
  }

  ss.toast('Gemini AI가 포트폴리오와 시장 뉴스를 분석 중입니다...', '🤖 AI 분석 시작', -1);

  try {
    const dashState = getDashboardStateForAI();
    const result = getGeminiAnalysis(config, dashState, customPrompt);

    // 당일 기존 추천 삭제 후 새 추천 기록
    if (result.json && result.json.ratios) {
      try { deleteTodayRecommendations(); recordAIRatios(result.json, '추천'); } catch (e) { Logger.log('추천 기록 실패: ' + e.message); }
    }

    ss.toast('분석이 완료되었습니다.', '✅', 2);
    showAIBriefingOutput(result.text, result.json);
  } catch (e) {
    Logger.log('AI 분석 오류: ' + e.toString());
    ss.toast('AI 분석 중 오류가 발생했습니다: ' + e.message, '❌ 오류');
  }
}

/**
 * Gemini API 호출 (Google Search Grounding 포함)
 */
function getGeminiAnalysis(config, dashState, customPrompt) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${config.geminiModelId}:generateContent?key=${config.geminiApiKey}`;
  const basePrompt = (customPrompt || getSystemPrompt()).replace('{{TARGET_YIELD}}', config.targetYield);
  const prompt = `${basePrompt}\n\n[현재 계좌 대시보드 상황]:\n${dashState}`;

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    tools: [{ googleSearch: {} }]
  };
  const options = {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify(payload), muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const resContent = response.getContentText();
  const resJson = JSON.parse(resContent);

  if (!resJson.candidates || !resJson.candidates[0] || !resJson.candidates[0].content) {
    throw new Error('AI 응답이 올바르지 않습니다: ' + resContent.substring(0, 300));
  }

  const fullText = resJson.candidates[0].content.parts[0].text;
  let proposedJson = null;
  let cleanText = fullText;

  // 1. 정상 포맷: [ALLOCATION_START]...[ALLOCATION_END]
  const tagged = fullText.match(/\[ALLOCATION_START\]\s*([\s\S]*?)\s*\[ALLOCATION_END\]/);
  if (tagged) {
    try {
      const raw = tagged[1].trim().replace(/```json\s*|```\s*/g, '').trim();
      proposedJson = JSON.parse(raw);
      cleanText = fullText.replace(/\*?\*?\[ALLOCATION_START\][\s\S]*?\[ALLOCATION_END\]\*?\*?/, '').trim();
    } catch (e) {
      Logger.log('[AI] tagged JSON 파싱 실패: ' + e.message);
    }
  }

  // 2. fallback: ratios 배열을 포함한 JSON 객체 직접 검색
  if (!proposedJson) {
    const objMatch = fullText.match(/\{\s*"ratios"\s*:\s*\[[\s\S]*?\]\s*\}/);
    if (objMatch) {
      try {
        proposedJson = JSON.parse(objMatch[0]);
        cleanText = fullText.replace(objMatch[0], '').trim();
      } catch (e) {
        Logger.log('[AI] fallback JSON 파싱 실패: ' + e.message);
      }
    }
  }

  return { text: cleanText, json: proposedJson };
}

/**
 * 분석 결과를 모달창으로 표시 (비중 비교 테이블 포함)
 */
function showAIBriefingOutput(content, proposedJson) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const jsonStr = proposedJson ? encodeURIComponent(JSON.stringify(proposedJson)) : '';

  // 비중 비교 테이블 데이터 구성
  const ratioMap = {};
  try {
    const settingSheet = ss.getSheetByName('📋 포트폴리오설정');
    if (settingSheet && settingSheet.getLastRow() >= 3) {
      settingSheet.getRange(3, 1, settingSheet.getLastRow() - 2, 4).getDisplayValues().forEach(function(row) {
        const code = String(row[0]).trim() || 'CASH';
        const name = String(row[1] || code).trim();
        const tgt  = parseFloat(String(row[3]).replace(/[^0-9.\-]/g, '')) || 0; // D열: 운용비율
        ratioMap[code] = { name: name, cur: 0, tgt: tgt };
      });
    }
    const dash = ss.getSheetByName('📊 대시보드');
    if (dash && dash.getLastRow() >= 9) {
      const numRows = dash.getLastRow() - 8;
      dash.getRange(9, 1, numRows, 8).getDisplayValues().forEach(function(row) {
        const code = String(row[0]).trim();
        if (!code) return;
        const cur = parseFloat(String(row[6] || '').replace(/[^0-9.\-]/g, '')) || 0;
        const tgt = parseFloat(String(row[7] || '').replace(/[^0-9.\-]/g, '')) || 0;
        if (ratioMap[code]) {
          ratioMap[code].cur = cur;
          if (tgt > 0) ratioMap[code].tgt = tgt;
        } else {
          ratioMap[code] = { name: String(row[1] || code).trim(), cur: cur, tgt: tgt };
        }
      });
      const totalEval = parseFloat(dash.getRange('B3').getValue()) || 0;
      const cashAmt   = parseFloat(dash.getRange('B4').getValue()) || 0;
      if (totalEval > 0) {
        if (!ratioMap['CASH']) ratioMap['CASH'] = { name: '현금', cur: 0, tgt: 5 };
        ratioMap['CASH'].cur = cashAmt / totalEval * 100;
      }
    }
  } catch (e) {
    Logger.log('[BRIEFING] ratioMap build error: ' + e.toString());
  }

  // AI 제안값 매핑
  let ratioRows = '';
  try {
    const toIntPct = function(v) { return (typeof v === 'number' && v > 0 && v <= 1) ? Math.round(v * 100) : (parseFloat(v) || 0); };
    const hasAI = !!(proposedJson && proposedJson.ratios && proposedJson.ratios.length);
    const aiByCode = {};
    let aiTotal = 0;
    let cashAiPct = null;
    if (hasAI) {
      proposedJson.ratios.forEach(function(p) {
        const code = String(p.code || '').trim();
        const name = String(p.name || '').trim();
        const isCash = code === 'CASH' || (!code && (!name || /현금|cash|예수/i.test(name))) || /현금|cash|예수금/i.test(name);
        const ai = toIntPct(p.ratio);
        aiTotal += ai;
        if (isCash) { cashAiPct = ai; return; }
        if (code) aiByCode[code] = ai;
      });
    }

    const stockCodes = Object.keys(ratioMap).filter(c => c !== 'CASH');
    stockCodes.forEach(function(code) {
      const r = ratioMap[code];
      const cells = [
        '<td class="nm">' + (r.name || code) + '</td>',
        '<td class="num">' + r.cur.toFixed(1) + '%</td>',
        '<td class="num">' + r.tgt.toFixed(1) + '%</td>'
      ];
      if (hasAI) {
        const aiVal = aiByCode[code];
        if (aiVal !== undefined) {
          const chg = aiVal - r.tgt;
          const cls = chg > 0.5 ? 'up' : chg < -0.5 ? 'dn' : 'nc';
          const arrow = chg > 0.5 ? '▲' : chg < -0.5 ? '▼' : '=';
          cells.push('<td class="' + cls + '">' + aiVal + '%</td>');
          cells.push('<td class="' + cls + '">' + arrow + (chg > 0 ? '+' : '') + chg.toFixed(0) + '%p</td>');
        } else {
          cells.push('<td class="num">—</td><td class="num">—</td>');
        }
      }
      ratioRows += '<tr>' + cells.join('') + '</tr>';
    });

    // 현금 행
    const cashRow = ratioMap['CASH'] || { name: '현금', cur: 0, tgt: 5 };
    const cashCells = [
      '<td class="nm">💵 현금</td>',
      '<td class="num">' + cashRow.cur.toFixed(1) + '%</td>',
      '<td class="num">' + cashRow.tgt.toFixed(1) + '%</td>'
    ];
    if (hasAI) {
      const cashAi = cashAiPct !== null ? cashAiPct : Math.max(0, 100 - aiTotal);
      if (cashAiPct === null) aiTotal += cashAi;
      const ccg = cashAi - cashRow.tgt;
      const ccls = ccg > 0.5 ? 'up' : ccg < -0.5 ? 'dn' : 'nc';
      const carrow = ccg > 0.5 ? '▲' : ccg < -0.5 ? '▼' : '=';
      cashCells.push('<td class="' + ccls + '">' + cashAi + '%</td>');
      cashCells.push('<td class="' + ccls + '">' + carrow + (ccg > 0 ? '+' : '') + ccg.toFixed(0) + '%p</td>');
    }
    ratioRows += '<tr style="background:#f8f9fa;font-style:italic;">' + cashCells.join('') + '</tr>';

    if (hasAI) {
      const aiTotalRounded = Math.round(aiTotal * 100) / 100;
      const totalCls = Math.abs(aiTotalRounded - 100) < 0.01 ? 'nc' : 'dn';
      const totalIcon = Math.abs(aiTotalRounded - 100) < 0.01 ? '✅' : '⚠️';
      ratioRows += '<tr style="background:#e8f0fe;font-weight:bold;">' +
        '<td class="nm">' + totalIcon + ' 합계</td><td class="num">100.0%</td><td class="num">100.0%</td>' +
        '<td class="' + totalCls + '">' + aiTotalRounded.toFixed(2) + '%</td><td class="num">—</td></tr>';
    }
  } catch (e) {
    Logger.log('[BRIEFING] table render error: ' + e.toString());
  }

  const headerCells = (proposedJson && proposedJson.ratios)
    ? '<th>현재%</th><th>목표%</th><th>AI 제안%</th><th>목표→제안</th>'
    : '<th>현재%</th><th>목표%</th>';
  const ratioTableHtml = ratioRows
    ? '<div class="section"><div class="st">📊 현재비율 / 목표비율 / AI 제안</div>' +
      '<table class="rt"><thead><tr><th style="text-align:left">종목명</th>' + headerCells + '</tr></thead>' +
      '<tbody>' + ratioRows + '</tbody></table></div>'
    : '';

  // 마크다운 렌더링
  function inlineMd(t) {
    return t.replace(/\*\*([^*\n]+)\*\*/g, '<strong>$1</strong>').replace(/\*([^*\n]+)\*/g, '<em>$1</em>');
  }
  const mdLines = content.split('\n');
  const mdParts = [];
  let inList = false;
  mdLines.forEach(function(line) {
    if (/^### /.test(line)) {
      if (inList) { mdParts.push('</ul>'); inList = false; }
      mdParts.push('<div class="h3">' + inlineMd(line.replace(/^### /, '')) + '</div>');
    } else if (/^## /.test(line)) {
      if (inList) { mdParts.push('</ul>'); inList = false; }
      mdParts.push('<div class="h2">' + inlineMd(line.replace(/^## /, '')) + '</div>');
    } else if (/^- /.test(line)) {
      if (!inList) { mdParts.push('<ul>'); inList = true; }
      mdParts.push('<li>' + inlineMd(line.replace(/^- /, '')) + '</li>');
    } else if (line.trim() === '') {
      if (inList) { mdParts.push('</ul>'); inList = false; }
      mdParts.push('<br>');
    } else {
      if (inList) { mdParts.push('</ul>'); inList = false; }
      mdParts.push('<p>' + inlineMd(line) + '</p>');
    }
  });
  if (inList) mdParts.push('</ul>');
  const mdHtml = mdParts.join('');

  const applyBtn = proposedJson
    ? '<input id="applyBtn" type="button" class="apply-btn" value="✅ 지금 즉시 비중 적용하기" onclick="applyRatios(\'' + jsonStr + '\')">'
    : '';

  const html = `<!DOCTYPE html><html><head><style>
*{box-sizing:border-box;}
body{font-family:'Malgun Gothic',sans-serif;line-height:1.6;color:#3c4043;margin:0;padding:0;font-size:13px;}
.wrap{padding:16px 20px;}
.title{font-size:17px;font-weight:bold;color:#4285F4;padding-bottom:10px;border-bottom:2px solid #e8eaed;margin-bottom:12px;}
.section{margin-bottom:12px;}
.st{font-weight:bold;font-size:12px;color:#5f6368;text-transform:uppercase;letter-spacing:.5px;margin-bottom:5px;}
.rt{width:100%;border-collapse:collapse;font-size:12px;}
.rt th{background:#e8f0fe;color:#1967d2;padding:5px 8px;text-align:center;font-weight:bold;}
.rt td{padding:4px 8px;border-bottom:1px solid #f1f3f4;}
.rt td.nm{font-weight:500;}
.rt td.num{text-align:right;color:#5f6368;}
.rt td.up{text-align:right;color:#137333;font-weight:bold;}
.rt td.dn{text-align:right;color:#c5221f;font-weight:bold;}
.rt td.nc{text-align:right;color:#9aa0a6;}
.md{max-height:360px;overflow-y:auto;padding-right:6px;}
.md .h3{font-size:13px;font-weight:bold;color:#1a73e8;margin:10px 0 4px;border-left:3px solid #4285F4;padding-left:7px;}
.md .h2{font-size:14px;font-weight:bold;color:#3c4043;margin:12px 0 4px;}
.md p{margin:2px 0;}
.md ul{margin:3px 0 3px 14px;padding:0;}
.md li{margin-bottom:4px;}
.md strong{color:#1a73e8;}
.md em{color:#5f6368;}
.apply-btn{display:block;width:100%;padding:11px;background:#4285F4;color:white;border:none;border-radius:6px;cursor:pointer;font-weight:bold;font-size:14px;margin-top:12px;}
.apply-btn:disabled{opacity:.5;cursor:not-allowed;}
.footer{margin-top:10px;text-align:right;border-top:1px solid #f1f3f4;padding-top:10px;}
.close-btn{padding:7px 14px;cursor:pointer;border:1px solid #dadce0;background:white;border-radius:4px;font-size:13px;}
</style><script>
function applyRatios(j){
  if(!confirm('AI가 제안한 비중으로 포트폴리오 설정을 변경하시겠습니까?'))return;
  var btn=document.getElementById('applyBtn');
  btn.value='적용 중...';btn.disabled=true;
  google.script.run
    .withSuccessHandler(function(){alert('✅ 운용비율이 업데이트되었습니다.\\n\\n대시보드 새로고침을 클릭하면 변경 내용을 확인할 수 있습니다.');google.script.host.close();})
    .withFailureHandler(function(e){alert('오류: '+e);btn.value='✅ 지금 즉시 비중 적용하기';btn.disabled=false;})
    .applyAIProposedRatiosManual(decodeURIComponent(j));
}
<\/script></head><body><div class="wrap">
<div class="title">🤖 Gemini AI 시장 분석 및 비중 제안</div>
${ratioTableHtml}
<div class="section"><div class="st">📝 분석 내용</div><div class="md">${mdHtml}</div></div>
${applyBtn}
<div class="footer"><button class="close-btn" onclick="google.script.host.close()">닫기</button></div>
</div></body></html>`;

  SpreadsheetApp.getUi().showModelessDialog(
    HtmlService.createHtmlOutput(html).setWidth(700).setHeight(900),
    ' '
  );
}

// ─────────────────────────────────────────────────────────────
// AI 빠른 질문
// ─────────────────────────────────────────────────────────────

/**
 * AI 빠른 질문 팝업 — 멀티턴 채팅 지원
 */
function openAIQuickQuestion() {
  const config = getConfig();
  if (!config.geminiApiKey) {
    SpreadsheetApp.getUi().alert('Gemini API Key가 필요한 기능입니다.');
    return;
  }
  const html = `<!DOCTYPE html><html><head><style>
*{box-sizing:border-box;}
html,body{height:100%;margin:0;padding:0;}
body{font-family:'Malgun Gothic',sans-serif;color:#3c4043;font-size:13px;line-height:1.5;display:flex;flex-direction:column;}
.hdr{font-size:16px;font-weight:bold;color:#4285F4;padding:14px 20px 10px;border-bottom:2px solid #e8eaed;flex-shrink:0;}
#inputSection{padding:14px 20px 16px;flex:1;display:flex;flex-direction:column;gap:6px;}
.lbl{font-size:11px;font-weight:bold;color:#5f6368;text-transform:uppercase;letter-spacing:.5px;}
.hint{font-size:11px;color:#9aa0a6;}
textarea{width:100%;border:1px solid #dadce0;border-radius:4px;padding:9px;font-family:inherit;font-size:13px;line-height:1.4;}
#q{flex:1;min-height:80px;resize:vertical;}
.cb-row{display:flex;align-items:center;gap:8px;font-size:13px;}
.btn-row{display:flex;gap:10px;margin-top:4px;}
.btn{flex:1;padding:10px;border:none;border-radius:6px;cursor:pointer;font-weight:bold;font-size:13px;}
.bp{background:#4285F4;color:white;}.bp:disabled{opacity:.5;cursor:not-allowed;}
.bs{background:white;border:1px solid #dadce0;color:#5f6368;}
#chatSection{display:none;flex:1;flex-direction:column;min-height:0;}
.chat-log{flex:1;overflow-y:auto;padding:10px 16px;background:#fafafa;min-height:0;}
.chat-user{margin:8px 0;padding:10px 12px;background:#e8f0fe;border-radius:8px;border-left:3px solid #4285F4;}
.chat-ai{margin:8px 0;padding:10px 12px;background:#f8f9fa;border-radius:8px;border-left:3px solid #34a853;}
.chat-lbl{font-size:11px;font-weight:bold;color:#5f6368;margin-bottom:4px;}
.chat-txt{white-space:pre-wrap;font-size:13px;line-height:1.6;}
.chat-foot{flex-shrink:0;border-top:1px solid #e8eaed;padding:10px 16px 14px;}
#followQ{width:100%;resize:none;margin-bottom:8px;}
</style></head><body>
<div class="hdr">💬 AI에게 질문하기</div>
<div id="inputSection">
  <div class="lbl">질문 내용</div>
  <textarea id="q" placeholder="예: 지금 S&P500 비중 늘려도 괜찮을까요? 요즘 금 시장 상황은?"></textarea>
  <div class="hint">간단한 시황 질문도 가능합니다.</div>
  <div class="cb-row"><input type="checkbox" id="inclData" checked><label for="inclData">📊 현재 포트폴리오 데이터 포함 (데이터 기반 질문 권장)</label></div>
  <div class="btn-row">
    <button class="btn bs" onclick="google.script.host.close()">닫기</button>
    <button class="btn bp" id="firstBtn" onclick="ask()">💬 질문하기</button>
  </div>
</div>
<div id="chatSection">
  <div class="chat-log" id="chatLog"></div>
  <div class="chat-foot">
    <textarea id="followQ" rows="2" placeholder="이어서 질문하기..."></textarea>
    <div class="btn-row">
      <button class="btn bs" onclick="google.script.host.close()">닫기</button>
      <button class="btn bp" id="sendBtn" onclick="send()">💬 전송</button>
    </div>
  </div>
</div>
<script>
var hist=[],inclFlag=false,MAX=10;
function ask(){
  var q=document.getElementById('q').value.trim();
  if(!q){alert('질문을 입력해주세요.');return;}
  inclFlag=document.getElementById('inclData').checked;
  hist=[{role:'user',text:q}];
  var btn=document.getElementById('firstBtn');
  btn.disabled=true;btn.textContent='⏳ 질문 중...';
  google.script.run
    .withSuccessHandler(function(r){hist.push({role:'model',text:r});showChat();})
    .withFailureHandler(function(e){btn.disabled=false;btn.textContent='💬 질문하기';alert('오류: '+e.message);})
    .runAIQuickQuestion(hist,inclFlag);
}
function showChat(){
  document.getElementById('inputSection').style.display='none';
  var cs=document.getElementById('chatSection');cs.style.display='flex';
  document.getElementById('chatLog').innerHTML='';
  bubble('user',hist[0].text);bubble('model',hist[1].text);scrollBot();
}
function send(){
  var q=document.getElementById('followQ').value.trim();
  if(!q)return;
  hist.push({role:'user',text:q});
  bubble('user',q);
  document.getElementById('followQ').value='';
  var btn=document.getElementById('sendBtn');
  btn.disabled=true;btn.textContent='⏳';
  var lid='ld_'+Date.now();
  bubble('model','⏳ 응답 중...',lid);scrollBot();
  var msgs=hist.slice(-MAX);
  while(msgs.length&&msgs[0].role!=='user')msgs=msgs.slice(1);
  google.script.run
    .withSuccessHandler(function(r){
      var el=document.getElementById(lid);if(el)el.remove();
      hist.push({role:'model',text:r});bubble('model',r);
      btn.disabled=false;btn.textContent='💬 전송';scrollBot();
    })
    .withFailureHandler(function(e){
      var el=document.getElementById(lid);if(el)el.remove();
      hist.pop();document.getElementById('followQ').value=q;btn.disabled=false;btn.textContent='💬 전송';alert('오류: '+e.message);
    })
    .runAIQuickQuestion(msgs,inclFlag);
}
function bubble(role,text,id){
  var log=document.getElementById('chatLog'),d=document.createElement('div');
  d.className=role==='user'?'chat-user':'chat-ai';if(id)d.id=id;
  d.innerHTML='<div class="chat-lbl">'+(role==='user'?'나':'🤖 AI')+'</div><div class="chat-txt">'+(role==='user'?esc(text):md(text))+'</div>';
  log.appendChild(d);
}
function esc(t){return t.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/\\n/g,'<br>');}
function md(t){
  t=t.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  t=t.replace(/\\*\\*([^*]+)\\*\\*/g,'<strong>$1</strong>');
  var lines=t.split('\\n');
  for(var i=0;i<lines.length;i++){
    var l=lines[i];
    if(/^#{1,3}\\s+/.test(l)){lines[i]='<strong>'+l.replace(/^#+\\s+/,'')+'</strong>';}
    else if(/^\\*\\s+/.test(l)){lines[i]='• '+l.replace(/^\\*\\s+/,'');}
    else if(/^-\\s+/.test(l)){lines[i]='• '+l.replace(/^-\\s+/,'');}
  }
  return lines.join('<br>');}
function scrollBot(){var l=document.getElementById('chatLog');l.scrollTop=l.scrollHeight;}
<\/script></body></html>`;
  SpreadsheetApp.getUi().showModelessDialog(
    HtmlService.createHtmlOutput(html).setWidth(620).setHeight(600),
    ' '
  );
}

/**
 * AI 빠른 질문 처리 (서버 사이드)
 * question이 배열인 경우 멀티턴으로 처리 (컨테이너 업데이트 없이도 동작)
 */
function runAIQuickQuestion(question, includeData) {
  if (Array.isArray(question)) {
    // 구버전 컨테이너는 inclData를 전달하지 않으므로 undefined → true로 기본값 처리
    return runAIQuickQuestionMultiTurn(question, includeData !== false);
  }

  const config = getConfig();
  if (!config.geminiApiKey) throw new Error('Gemini API Key가 설정되지 않았습니다.');

  const url = `https://generativelanguage.googleapis.com/v1beta/models/${config.geminiModelId}:generateContent?key=${config.geminiApiKey}`;

  const systemContext = `너는 한국 주식·ETF 포트폴리오 투자 전문가야. 사용자는 한국투자증권 API로 자동 리밸런싱을 운용 중이며, 국내주식·해외주식·채권·금·달러·현금을 분산하는 자산배분 포트폴리오를 관리하고 있어. 질문에 한국어로 간결하고 실용적으로 답변해줘. 투자 조언이 아닌 참고용 분석임을 명심해.`;

  let userPrompt = question;
  if (includeData) {
    const dashState = getDashboardStateForAI();
    userPrompt = `${question}\n\n[현재 포트폴리오 현황]:\n${dashState}`;
  }

  const payload = {
    system_instruction: { parts: [{ text: systemContext }] },
    contents: [{ parts: [{ text: userPrompt }] }],
    tools: [{ googleSearch: {} }]
  };
  const options = { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true };

  const response = UrlFetchApp.fetch(url, options);
  const resJson = JSON.parse(response.getContentText());
  if (resJson.candidates && resJson.candidates[0] && resJson.candidates[0].content) {
    return resJson.candidates[0].content.parts[0].text;
  }
  throw new Error('AI 응답이 올바르지 않습니다: ' + response.getContentText().substring(0, 200));
}

/**
 * AI 멀티턴 질문 처리 — 대화 히스토리를 Gemini contents 배열로 전달
 * @param {{role:'user'|'model', text:string}[]} messages - 최근 N개 메시지 배열 (클라이언트에서 트리밍)
 * @param {boolean} inclData - true면 첫 번째 user 메시지에 포트폴리오 현황 주입
 */
function runAIQuickQuestionMultiTurn(messages, inclData) {
  const config = getConfig();
  if (!config.geminiApiKey) throw new Error('Gemini API Key가 설정되지 않았습니다.');
  if (!messages || !Array.isArray(messages) || messages.length === 0) {
    throw new Error('메시지 배열이 비어있거나 유효하지 않습니다.');
  }

  const url = `https://generativelanguage.googleapis.com/v1beta/models/${config.geminiModelId}:generateContent?key=${config.geminiApiKey}`;

  const systemContext = `너는 한국 주식·ETF 포트폴리오 투자 전문가야. 사용자는 한국투자증권 API로 자동 리밸런싱을 운용 중이며, 국내주식·해외주식·채권·금·달러·현금을 분산하는 자산배분 포트폴리오를 관리하고 있어. 질문에 한국어로 간결하고 실용적으로 답변해줘. 투자 조언이 아닌 참고용 분석임을 명심해.`;

  const contents = messages.map(function(msg, idx) {
    let text = msg.text;
    if (inclData && idx === 0 && msg.role === 'user') {
      if (text.indexOf('[현재 포트폴리오 현황]') === -1) {
        text = text + '\n\n[현재 포트폴리오 현황]:\n' + getDashboardStateForAI();
      }
    }
    return { role: msg.role, parts: [{ text: text }] };
  });
  if (!contents[0] || contents[0].role !== 'user') {
    throw new Error('첫 번째 메시지는 반드시 사용자(role: user) 메시지여야 합니다.');
  }

  const payload = {
    system_instruction: { parts: [{ text: systemContext }] },
    contents: contents,
    tools: [{ googleSearch: {} }]
  };
  const options = {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify(payload), muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const resJson = JSON.parse(response.getContentText());
  if (resJson.candidates && resJson.candidates[0] && resJson.candidates[0].content) {
    return resJson.candidates[0].content.parts[0].text;
  }
  throw new Error('AI 응답이 올바르지 않습니다: ' + response.getContentText().substring(0, 200));
}

// ─────────────────────────────────────────────────────────────
// 비중 적용 (Layer 1 락 포함)
// ─────────────────────────────────────────────────────────────

const RATIO_LOCK_DAYS_DEFAULT = 14;

/**
 * AI 비중 변경 락 상태 조회
 */
function getRatioLockStatus() {
  const props = PropertiesService.getScriptProperties();
  const lastStr = props.getProperty('LAST_AI_RATIO_UPDATE');
  const lockDays = RATIO_LOCK_DAYS_DEFAULT;

  if (!lastStr) {
    return { locked: false, daysRemaining: 0, daysSinceUpdate: Infinity, lockDays, lastUpdate: null };
  }

  const last = new Date(lastStr);
  const now = new Date();
  const daysSinceUpdate = Math.floor((now - last) / 86400000);
  const daysRemaining = Math.max(0, lockDays - daysSinceUpdate);
  return { locked: daysSinceUpdate < lockDays, daysRemaining, daysSinceUpdate, lockDays, lastUpdate: last };
}

/**
 * 락 해제 (수동)
 */
function unlockRatioChange() {
  PropertiesService.getScriptProperties().deleteProperty('LAST_AI_RATIO_UPDATE');
  SpreadsheetApp.getUi().alert('🔓 AI 비중 변경 락이 해제되었습니다.');
}

/**
 * AI가 제안한 비중을 포트폴리오설정 D열(운용비율)에 적용
 */
function applyAIProposedRatios(jsonStr, options) {
  options = options || {};
  const data = typeof jsonStr === 'string' ? JSON.parse(jsonStr) : jsonStr;
  if (!data || !data.ratios) throw new Error('유효하지 않은 데이터입니다.');

  if (!options.force) {
    const lock = getRatioLockStatus();
    if (lock.locked) {
      Logger.log('[L1 락] 비중 변경 스킵 — ' + lock.daysRemaining + '일 남음');
      return { applied: false, locked: true, daysRemaining: lock.daysRemaining };
    }
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('📋 포트폴리오설정');
  if (!sheet) throw new Error('포트폴리오설정 시트를 찾을 수 없습니다.');

  const lastRow = sheet.getLastRow();
  // col: 0=종목코드(A), 1=종목명(B), 2=기준비율(C/고정), 3=운용비율(D) ← AI는 D열만 수정
  const range = sheet.getRange(3, 1, lastRow - 2, 4);
  const values = range.getValues();

  const toIntPct = v => (typeof v === 'number' && v > 0 && v <= 1) ? Math.round(v * 100) : v;

  const updatedValues = values.map(row => {
    const code = String(row[0]).trim();
    const name = String(row[1]).trim();
    const proposal = data.ratios.find(r =>
      String(r.code || '').trim() === code ||
      String(r.name || '').trim() === name ||
      (code === '' && (r.code === 'CASH' || /현금|cash/i.test(r.name || '')))
    );
    if (proposal) {
      row[3] = toIntPct(proposal.ratio);
    }
    return row;
  });

  range.setValues(updatedValues);
  PropertiesService.getScriptProperties().setProperty('LAST_AI_RATIO_UPDATE', new Date().toISOString());
  recordAIRatios(data, '적용한 비중');
  return { applied: true, locked: false };
}

/**
 * 수동 적용용 래퍼 — 다이얼로그 "지금 적용" 버튼에서 호출 (락 무시)
 */
function applyAIProposedRatiosManual(jsonStr) {
  return applyAIProposedRatios(jsonStr, { force: true });
}

// ─────────────────────────────────────────────────────────────
// 비중변경이력 기록
// ─────────────────────────────────────────────────────────────

/**
 * AI 비중 제안 내역을 '📝 비중변경이력' 시트에 누적 기록
 */
function recordAIRatios(data, status) {
  status = status || '추천';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupAIHistorySheet();
  const historySheet = ss.getSheetByName('📝 비중변경이력');

  const config = getConfig();
  const now = new Date();
  const toIntPct = v => (typeof v === 'number' && v > 0 && v <= 1) ? Math.round(v * 100) : (parseFloat(v) || 0);

  // 포트폴리오설정에서 현재 비중(before)·유형 읽기
  const settingSheet = ss.getSheetByName('📋 포트폴리오설정');
  const currentMap = {};
  if (settingSheet && settingSheet.getLastRow() >= 3) {
    settingSheet.getRange(3, 1, settingSheet.getLastRow() - 2, 5).getValues().forEach(row => {
      const code = String(row[0]).trim() || 'CASH';
      currentMap[code] = { name: row[1], ratio: parseFloat(row[3]) || 0, type: row[4] || '' }; // D열: 운용비율
    });
  }

  const rows = data.ratios.map(proposal => {
    const code = String(proposal.code || '').trim();
    const cur  = currentMap[code] || {};
    const assetName  = cur.name || (code === 'CASH' ? '현금' : proposal.name || code);
    const assetType  = cur.type || '';
    const beforeRatio = cur.ratio != null ? cur.ratio : '';
    const afterRatio  = toIntPct(proposal.ratio);
    return [now, assetName, assetType, beforeRatio, afterRatio, proposal.rationale || data.summary || 'AI 비중 제안', config.geminiModelId, status];
  });

  if (rows.length > 0) {
    const lastRow = historySheet.getLastRow();
    const startRow = Math.max(lastRow + 1, 2);
    historySheet.getRange(startRow, 1, rows.length, 8).setValues(rows);
  }
  trimExtraColumns(historySheet, 8);
}

function hasTodayRecommendation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('📝 비중변경이력');
  if (!sheet || sheet.getLastRow() < 2) return false;
  const tz = Session.getScriptTimeZone();
  const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
  return data.some(row => {
    if (!row[0]) return false;
    return Utilities.formatDate(new Date(row[0]), tz, 'yyyy-MM-dd') === todayStr && String(row[7]).trim() === '추천';
  });
}

function deleteTodayRecommendations() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('📝 비중변경이력');
  if (!sheet || sheet.getLastRow() < 2) return;
  const tz = Session.getScriptTimeZone();
  const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
  const toDelete = [];
  data.forEach((row, idx) => {
    if (!row[0]) return;
    if (Utilities.formatDate(new Date(row[0]), tz, 'yyyy-MM-dd') === todayStr && String(row[7]).trim() === '추천') {
      toDelete.push(idx + 2);
    }
  });
  for (let i = toDelete.length - 1; i >= 0; i--) sheet.deleteRow(toDelete[i]);
}

// ─────────────────────────────────────────────────────────────
// 추천 비중 반영
// ─────────────────────────────────────────────────────────────

/**
 * 가장 최근의 '추천' 상태인 비중 제안을 찾아 실제 포트폴리오에 반영
 */
function applyLatestRecommendation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let historySheet = ss.getSheetByName('📝 비중변경이력');
  if (!historySheet) {
    setupAIHistorySheet();
    historySheet = ss.getSheetByName('📝 비중변경이력');
  }

  const lastRow = historySheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('반영할 추천 내역이 없습니다.');
    return;
  }

  const data = historySheet.getRange(2, 1, lastRow - 1, 8).getValues();

  let latestTime = null;
  const recommendedRatios = [];
  const rowsToUpdate = [];

  for (let i = data.length - 1; i >= 0; i--) {
    const row = data[i];
    const time   = row[0].toString();
    const status = row[7];
    if (status === '추천') {
      if (latestTime === null) latestTime = time;
      if (time === latestTime) {
        recommendedRatios.push({ name: row[1], ratio: parseFloat(row[4]) });
        rowsToUpdate.push(i + 2);
      } else {
        break;
      }
    }
  }

  if (recommendedRatios.length === 0) {
    SpreadsheetApp.getUi().alert('적용할 수 있는 최근 추천 내역이 없습니다.');
    return;
  }

  const settingSheet = ss.getSheetByName('📋 포트폴리오설정');
  const lastSettingRow = settingSheet.getLastRow();
  // D열(운용비율) 업데이트
  const range = settingSheet.getRange(3, 1, lastSettingRow - 2, 4);
  const values = range.getValues();

  const updatedValues = values.map(row => {
    const code = String(row[0]).trim();
    const name = row[1];
    const match = recommendedRatios.find(r => r.name === code || r.name === name);
    if (match) row[3] = match.ratio; // D열
    return row;
  });

  range.setValues(updatedValues);

  rowsToUpdate.forEach(rowIdx => {
    historySheet.getRange(rowIdx, 8).setValue('적용됨');
  });

  PropertiesService.getScriptProperties().setProperty('LAST_AI_RATIO_UPDATE', new Date().toISOString());
  ss.toast('추천 비중이 성공적으로 반영되었습니다.', '✅ 반영 완료');
  updateDashboard();
}

// ─────────────────────────────────────────────────────────────
// 누적 AI 추천 평균 계산 (자동화용)
// ─────────────────────────────────────────────────────────────

/**
 * 비중변경이력 시트에서 직전 N일간 '추천' 평균 비중을 계산
 */
function getAveragedAIRatios(days) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('📝 비중변경이력');
  if (!sheet || sheet.getLastRow() < 2) return null;

  const props = PropertiesService.getScriptProperties();
  const lastUpdateStr = props.getProperty('LAST_AI_RATIO_UPDATE');
  let cutoffMs = Date.now() - days * 86400000;
  if (lastUpdateStr) {
    cutoffMs = Math.max(cutoffMs, new Date(lastUpdateStr).getTime());
  }

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
  const grouped = {};
  const seenTimes = {};
  let runCount = 0;

  data.forEach(row => {
    const time = row[0] instanceof Date ? row[0] : new Date(row[0]);
    if (!time || isNaN(time.getTime()) || time.getTime() < cutoffMs) return;
    const name = String(row[1] || '').trim();
    const ratio = parseFloat(row[4]) || 0;
    const status = String(row[7] || '').trim();
    if (status !== '추천' || !name) return;
    if (!grouped[name]) grouped[name] = [];
    grouped[name].push(ratio);
    const tk = time.getTime();
    if (!seenTimes[tk]) { seenTimes[tk] = true; runCount++; }
  });

  const names = Object.keys(grouped);
  if (names.length === 0 || runCount < 1) return null;

  const averages = {};
  names.forEach(n => {
    const arr = grouped[n];
    averages[n] = arr.reduce((s, v) => s + v, 0) / arr.length;
  });

  const settingSheet = ss.getSheetByName('📋 포트폴리오설정');
  if (!settingSheet) return null;
  const settings = settingSheet.getRange(3, 1, Math.max(1, settingSheet.getLastRow() - 2), 2).getValues();
  const ratios = [];
  settings.forEach(row => {
    const code = String(row[0]).trim() || 'CASH';
    const name = String(row[1] || '').trim();
    if (averages[name] !== undefined) {
      ratios.push({ code: code, name: name, ratio: Math.round(averages[name]) });
    }
  });

  if (ratios.length > 0) {
    const total = ratios.reduce((s, r) => s + r.ratio, 0);
    if (total !== 100) {
      let maxIdx = 0;
      for (let i = 1; i < ratios.length; i++) {
        if (ratios[i].ratio > ratios[maxIdx].ratio) maxIdx = i;
      }
      ratios[maxIdx].ratio += (100 - total);
    }
  }

  return {
    ratios: ratios,
    summary: `직전 ${days}일 ${runCount}회 AI 추천의 평균치`,
    runCount: runCount
  };
}
