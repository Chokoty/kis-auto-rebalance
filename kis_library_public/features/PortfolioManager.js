/**
 * 포트폴리오 종목 추가/관리 다이얼로그
 */
function openPortfolioManagerDialog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('📋 포트폴리오설정');

  const portfolio = [];
  if (sheet && sheet.getLastRow() >= 3) {
    // col: 0=종목코드, 1=종목명, 2=기준비율(C/고정), 3=운용비율(D), 4=유형(E)
    sheet.getRange(3, 1, sheet.getLastRow() - 2, 5).getValues().forEach(function(row) {
      if (row[1] || row[0]) {
        portfolio.push({
          code: String(row[0]).trim(),
          name: String(row[1]),
          initialRatio: parseFloat(row[2]) || 0,
          ratio: parseFloat(row[3]) || 0,
          type: String(row[4] || '')
        });
      }
    });
  }

  const initJson = JSON.stringify(portfolio);

  const html = '<!DOCTYPE html><html><head><style>' +
'*{box-sizing:border-box;}' +
'body{font-family:"Malgun Gothic",sans-serif;padding:0;margin:0;font-size:13px;color:#3c4043;}' +
'.wrap{padding:14px 16px;}' +
'.title{font-size:16px;font-weight:bold;color:#1a73e8;margin-bottom:12px;}' +
'.section{background:#f8f9fa;border-radius:8px;padding:12px;margin-bottom:10px;}' +
'.st{font-weight:bold;font-size:11px;color:#5f6368;text-transform:uppercase;letter-spacing:.5px;margin-bottom:8px;}' +
'.search-row{display:flex;gap:6px;align-items:center;}' +
'input[type=text]{flex:1;padding:7px 10px;border:1px solid #dadce0;border-radius:4px;font-size:13px;}' +
'.btn{padding:7px 12px;border:none;border-radius:4px;cursor:pointer;font-weight:bold;font-size:12px;}' +
'.btn:disabled{opacity:.5;cursor:not-allowed;}' +
'.btn-blue{background:#1a73e8;color:white;}' +
'.btn-green{background:#34a853;color:white;}' +
'.btn-outline{background:white;border:1px solid #dadce0;color:#3c4043;}' +
'.found{margin-top:8px;display:none;}' +
'.found-name{font-weight:bold;color:#1a73e8;font-size:13px;margin-bottom:6px;}' +
'.found-row{display:flex;gap:6px;align-items:center;}' +
'select{padding:5px 6px;border:1px solid #dadce0;border-radius:4px;font-size:12px;background:white;}' +
'.ri{width:58px;padding:5px 7px;border:1px solid #dadce0;border-radius:4px;font-size:13px;text-align:right;}' +
'.msg{font-size:12px;margin-top:6px;min-height:18px;}' +
'.msg.ok{color:#137333;}.msg.err{color:#c5221f;}.msg.info{color:#5f6368;}' +
'.total-bar{display:flex;justify-content:space-between;align-items:center;padding:4px 2px;margin-bottom:6px;}' +
'.total-lbl{font-size:12px;color:#5f6368;}' +
'.tv{font-weight:bold;font-size:14px;}' +
'.tv.ok{color:#137333;}.tv.warn{color:#c5221f;}' +
'.pt{width:100%;border-collapse:collapse;font-size:12px;}' +
'.pt th{background:#e8f0fe;color:#1967d2;padding:5px 6px;font-size:11px;}' +
'.pt td{padding:4px 5px;border-bottom:1px solid #f1f3f4;vertical-align:middle;}' +
'.pt td.nm{max-width:130px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}' +
'.pt td.cd{color:#9aa0a6;font-size:11px;text-align:center;}' +
'.pt input[type=number]{width:52px;padding:3px 5px;border:1px solid #dadce0;border-radius:3px;text-align:right;font-size:12px;}' +
'.pt select{padding:3px 4px;font-size:11px;}' +
'.db{background:none;border:none;cursor:pointer;color:#ea4335;font-size:13px;padding:2px 5px;border-radius:3px;}' +
'.db:hover{background:#fce8e6;}' +
'.cash-btn{font-size:11px;padding:4px 8px;margin-top:6px;}' +
'.footer{margin-top:10px;display:flex;justify-content:space-between;align-items:center;border-top:1px solid #e8eaed;padding-top:10px;}' +
'.sv{padding:10px 22px;background:#1a73e8;color:white;border:none;border-radius:6px;cursor:pointer;font-weight:bold;font-size:14px;}' +
'.sv:disabled{opacity:.5;cursor:not-allowed;}' +
'</style></head><body><div class="wrap">' +
'<div class="title">📋 포트폴리오 종목 관리</div>' +

// 종목 추가 섹션
'<div class="section">' +
'<div class="st">➕ 종목 추가</div>' +
'<div class="search-row">' +
'<input type="text" id="codeInput" placeholder="종목코드 입력 (예: 440650)" maxlength="10" onkeydown="if(event.key===\'Enter\')searchStock()">' +
'<button class="btn btn-blue" id="searchBtn" onclick="searchStock()">🔍 검색</button>' +
'</div>' +
'<div class="found" id="foundBox">' +
'<div class="found-name" id="foundName"></div>' +
'<div class="found-row">' +
'<span style="font-size:12px;color:#5f6368">유형</span>' +
'<select id="newType"></select>' +
'<span style="font-size:12px;color:#5f6368">운용비율</span>' +
'<input type="number" class="ri" id="newRatio" placeholder="%" min="0" max="100" step="1">' +
'<span style="font-size:12px;color:#5f6368">%</span>' +
'<button class="btn btn-green" onclick="addStock()">추가</button>' +
'</div>' +
'</div>' +
'<div class="msg info" id="searchMsg"></div>' +
'<button class="btn btn-outline cash-btn" onclick="addCash()">+ 현금 행 추가</button>' +
'</div>' +

// 포트폴리오 테이블
'<div class="section">' +
'<div class="st">📊 현재 포트폴리오</div>' +
'<div class="total-bar">' +
'<span class="total-lbl">운용비율 합계</span>' +
'<span class="tv" id="totalPct">-</span>' +
'</div>' +
'<div style="max-height:240px;overflow-y:auto;">' +
'<table class="pt"><thead><tr>' +
'<th style="text-align:left">종목명</th><th>코드</th><th>유형</th><th>비율%</th><th></th>' +
'</tr></thead><tbody id="ptBody"></tbody></table>' +
'</div></div>' +

'<div class="footer">' +
'<button class="btn btn-outline" onclick="google.script.host.close()">닫기</button>' +
'<button class="sv" id="saveBtn" onclick="saveAll()">💾 저장하기</button>' +
'</div></div>' +

'<script>' +
'var TYPES=["채권","금","달러","국내주식","해외주식","현금","기타"];' +
'var portfolio=' + initJson + ';' +
'var foundStock=null;' +

'function typeOpts(sel){return TYPES.map(function(t){return\'<option value="\'+t+\'"\'+( t===sel?\' selected\':\'\')+\'>\'+t+"</option>";}).join("");}' +

'function guessType(n){' +
'  n=n.toLowerCase();' +
'  if(/채권|bond|단기채|국채|회사채/.test(n))return "채권";' +
'  if(/골드|금선물|gold/.test(n))return "금";' +
'  if(/달러|usd|달러선물/.test(n))return "달러";' +
'  if(/s&p|나스닥|nasdaq|미국s&p|미국나스닥|미국 s|글로벌|선진국|us equity/.test(n))return "해외주식";' +
'  if(/미국/.test(n)&&!/달러/.test(n))return "해외주식";' +
'  if(/현금|cash/.test(n))return "현금";' +
'  return "국내주식";' +
'}' +

'function renderTable(){' +
'  var html=portfolio.map(function(r,i){' +
'    return "<tr>"' +
'      +"<td class=nm title=\\""+r.name+"\\">"+r.name+"</td>"' +
'      +"<td class=cd>"+(r.code||"-")+"</td>"' +
'      +"<td><select onchange=\\"portfolio["+i+"].type=this.value\\">"+typeOpts(r.type)+"</select></td>"' +
'      +"<td><input type=number value=\\""+r.ratio+"\\" min=0 max=100 step=1 '  +
'         onchange=\\"portfolio["+i+"].ratio=parseFloat(this.value)||0;updateTotal()\\"></td>"' +
'      +"<td><button class=db onclick=\\"removeStock("+i+")\\" title=삭제>✕</button></td>"' +
'      +"</tr>";' +
'  }).join("");' +
'  document.getElementById("ptBody").innerHTML=html;' +
'  updateTotal();' +
'}' +

'function updateTotal(){' +
'  var t=portfolio.reduce(function(s,r){return s+(parseFloat(r.ratio)||0);},0);' +
'  var el=document.getElementById("totalPct");' +
'  el.textContent=t.toFixed(0)+"%";' +
'  el.className="tv "+(Math.abs(t-100)<0.5?"ok":"warn");' +
'}' +

'function removeStock(i){' +
'  var r=portfolio[i];' +
'  if(r&&r.type!=="현금"&&parseFloat(r.ratio)>0){' +
'    if(confirm((r.name||r.code)+"을(를) 목표에서 제외합니다.\\n비율을 0%로 바꿔 대시보드에서 매도 후 삭제하거나,\\n[확인]=즉시 삭제 / [취소]=0%로 변경")){' +
'      portfolio.splice(i,1);' +
'    } else {' +
'      portfolio[i]=Object.assign({},r,{ratio:0});' +
'    }' +
'  } else {' +
'    portfolio.splice(i,1);' +
'  }' +
'  renderTable();' +
'}' +

'function showMsg(txt,cls){var el=document.getElementById("searchMsg");el.textContent=txt;el.className="msg "+(cls||"info");}' +

'function searchStock(){' +
'  var code=document.getElementById("codeInput").value.trim().replace(/[^0-9A-Za-z]/g,"");' +
'  if(!code){showMsg("종목코드를 입력하세요.","err");return;}' +
'  var btn=document.getElementById("searchBtn");' +
'  btn.disabled=true;btn.textContent="조회 중...";' +
'  document.getElementById("foundBox").style.display="none";' +
'  foundStock=null;' +
'  showMsg("KIS API에서 종목 정보 조회 중...","info");' +
'  google.script.run' +
'    .withSuccessHandler(function(res){' +
'      btn.disabled=false;btn.textContent="🔍 검색";' +
'      if(!res.success){showMsg("❌ "+(res.error||"종목을 찾을 수 없습니다."),"err");return;}' +
'      var t=guessType(res.name);' +
'      foundStock={code:code,name:res.name,type:t,ratio:0};' +
'      document.getElementById("foundName").textContent=res.name;' +
'      document.getElementById("newType").innerHTML=typeOpts(t);' +
'      document.getElementById("newRatio").value="";' +
'      document.getElementById("foundBox").style.display="block";' +
'      showMsg("","info");' +
'      document.getElementById("newRatio").focus();' +
'    })' +
'    .withFailureHandler(function(e){' +
'      btn.disabled=false;btn.textContent="🔍 검색";' +
'      showMsg("❌ 오류: "+e.message,"err");' +
'    })' +
'    .searchStockByCode(code);' +
'}' +

'function addStock(){' +
'  if(!foundStock){return;}' +
'  var ratio=parseFloat(document.getElementById("newRatio").value)||0;' +
'  if(ratio<=0){alert("운용비율을 입력하세요.");return;}' +
'  foundStock.type=document.getElementById("newType").value;' +
'  var dup=portfolio.find(function(r){return r.code&&r.code===foundStock.code;});' +
'  if(dup){alert("이미 포트폴리오에 있는 종목입니다.");return;}' +
'  portfolio.push({code:foundStock.code,name:foundStock.name,ratio:ratio,type:foundStock.type});' +
'  foundStock=null;' +
'  document.getElementById("codeInput").value="";' +
'  document.getElementById("foundBox").style.display="none";' +
'  showMsg("✅ 추가되었습니다. 저장 버튼을 눌러 확정하세요.","ok");' +
'  renderTable();' +
'}' +

'function addCash(){' +
'  var dup=portfolio.find(function(r){return r.name==="현금"||r.type==="현금";});' +
'  if(dup){alert("현금 행이 이미 있습니다.");return;}' +
'  portfolio.push({code:"",name:"현금",ratio:5,type:"현금"});' +
'  renderTable();' +
'}' +

'function saveAll(){' +
'  var total=portfolio.reduce(function(s,r){return s+r.ratio;},0);' +
'  if(Math.abs(total-100)>0.5){' +
'    if(!confirm("비율 합계가 "+total.toFixed(0)+"%입니다. (100%가 아님)\\n계속 저장하시겠습니까?"))return;' +
'  }' +
'  var btn=document.getElementById("saveBtn");' +
'  btn.disabled=true;btn.textContent="저장 중...";' +
'  google.script.run' +
'    .withSuccessHandler(function(){' +
'      alert("✅ 포트폴리오가 저장되었습니다.");' +
'      google.script.host.close();' +
'    })' +
'    .withFailureHandler(function(e){' +
'      alert("오류: "+e.message);' +
'      btn.disabled=false;btn.textContent="💾 저장하기";' +
'    })' +
'    .savePortfolioSettings(portfolio);' +
'}' +

'renderTable();' +
'<\/script></body></html>';

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(560).setHeight(680),
    ' '
  );
}

/**
 * KIS API로 종목코드 → 종목명 조회
 * search-stock-info(CTPF1604R) 우선 시도, 실패 시 inquire-price(FHKST01010100) fallback
 */
function searchStockByCode(code) {
  code = String(code).trim().replace(/[^0-9A-Za-z]/g, '');
  if (!code) return { success: false, error: '종목코드를 입력하세요.' };

  // 1차: search-stock-info (종목 기본정보 조회)
  try {
    const data = callKISAPI(
      '/uapi/domestic-stock/v1/quotations/search-stock-info',
      'CTPF1604R',
      { 'PRDT_TYPE_CD': '300', 'PDNO': code }
    );
    const out = data.output || {};
    const name = out.prdt_abrv_name || out.prdt_name || '';
    if (name) return { success: true, name: name.trim(), code: code };
  } catch(e) {
    Logger.log('[searchStockByCode] search-stock-info 실패: ' + e.message);
  }

  // 2차: inquire-price fallback (KOSPI 'J' → KOSDAQ 'Q' 순으로 시도)
  for (const mktCode of ['J', 'Q']) {
    try {
      const data = callKISAPI(
        '/uapi/domestic-stock/v1/quotations/inquire-price',
        'FHKST01010100',
        { 'FID_COND_MRKT_DIV_CODE': mktCode, 'FID_INPUT_ISCD': code }
      );
      const out = data.output || {};
      const name = out.hts_kor_isnm || out.prdt_abrv_name || '';
      if (name) return { success: true, name: name.trim(), code: code };
    } catch(e) {
      Logger.log('[searchStockByCode] inquire-price(' + mktCode + ') 실패: ' + e.message);
    }
  }

  return { success: false, error: '종목을 찾을 수 없습니다. 코드를 다시 확인해주세요.' };
}

/**
 * 포트폴리오설정 시트에 저장
 */
function savePortfolioSettings(rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('📋 포트폴리오설정');
  if (!sheet) throw new Error('포트폴리오설정 시트를 찾을 수 없습니다.');

  // 기존 기준비율(C열) 보존: 새 종목은 운용비율로 초기화
  const lastRow = sheet.getLastRow();
  const existingInitialRatios = {};
  if (lastRow >= 3) {
    sheet.getRange(3, 1, lastRow - 2, 3).getValues().forEach(function(row) {
      const code = String(row[0]).trim();
      if (code && typeof row[2] === 'number' && row[2] > 0) {
        existingInitialRatios[code] = row[2];
      }
    });
    sheet.getRange(3, 1, lastRow - 2, 5).clearContent();
  }

  if (rows.length > 0) {
    // col: 종목코드 | 종목명 | 기준비율(고정) | 운용비율 | 유형
    const values = rows.map(function(r) {
      const initialRatio = existingInitialRatios[r.code] || r.initialRatio || r.ratio || 0;
      return [r.code || '', r.name || '', initialRatio, r.ratio || 0, r.type || ''];
    });
    const range = sheet.getRange(3, 1, values.length, 5);
    range.setValues(values);
    range.setHorizontalAlignment('right');
    sheet.getRange(3, 2, values.length, 1).setHorizontalAlignment('left');
    sheet.getRange(3, 5, values.length, 1).setHorizontalAlignment('left');
  }

  return true;
}
