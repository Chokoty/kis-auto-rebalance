// 대시보드 시트 초기 설정 (목표 수익률 추가)
function setupDashboardSheet(sheet) {
  // 전체 삭제 대신 필요한 부분만 초기화 (UX 개선)
  // sheet.clear(); -> 삭제

  sheet.getRange('A1:N1').merge()
    .setValue('📊 포트폴리오 대시보드')
    .setFontWeight('bold')
    .setFontSize(16)
    .setHorizontalAlignment('center')
    .setBackground('#4285f4')
    .setFontColor('white');

  // 업데이트 알림 배너
  const updateMsg = (typeof checkVersionUpdate === 'function') ? checkVersionUpdate() : null;
  if (updateMsg) {
    sheet.getRange('A2:N2').merge()
      .setValue(updateMsg)
      .setFontWeight('bold')
      .setFontColor('white')
      .setBackground('#ea4335')
      .setHorizontalAlignment('center');
  } else {
    sheet.getRange('A2:N2').breakApart().clearContent().setBackground('white');
  }
  
  // 요약 정보 라벨 (A열)
  sheet.getRange('A3:A7').setValues([
    ['💰 총 평가액'],
    ['💵 예수금'],
    ['💵 현금 비율'],
    ['🕒 업데이트'],
    ['🛣️ 주행 상태']
  ]);
  
  sheet.getRange('A3:A7').setFontWeight('bold').setBackground('#f3f3f3').setHorizontalAlignment('center');
  sheet.getRange('B3:B7').setHorizontalAlignment('right');
  
  // 리밸런싱 요약 라벨 (D열)
  sheet.getRange('D3:D7').setValues([
    ['🔄 리밸런싱 대상'],
    ['🛒 총 매수 필요액'],
    ['📉 총 매도 예정액'],
    ['💸 예상 제비용'],
    ['💳 순 자산 변동']
  ]);
  sheet.getRange('D8:E8').clearContent().setBackground('white'); // 기존 위치 삭제

  // 투자 현황 및 자동화 상태 (G열)
  sheet.getRange('G3:G7').setValues([
    ['💵 총 투자금'],
    ['💰 평가손익'],
    ['📈 수익률'],
    ['🎯 목표 수익률(연)'],
    ['💡 월 인출 추천']
  ]);
  
  sheet.getRange('G3:G7').setFontWeight('bold').setBackground('#f3f3f3').setHorizontalAlignment('center');
  sheet.getRange('H3:H7').setHorizontalAlignment('right');
  
  // 기존 흔적 제거
  sheet.getRange('G8:H8').clearContent().setBackground('white');
  
  // 테이블 헤더 (8행으로 복구)
  sheet.getRange('A8:O8').setValues([[
    '종목코드', '종목명', '유형',
    '보유수량', '현재가', '평가액',
    '현재비율', '목표비율', '예상비중', '차이',
    '필요액션', '수량/금액',
    '실행수량', '실행가격', '평균단가'
  ]])
  .setFontWeight('bold')
  .setBackground('#34a853')
  .setFontColor('white')
  .setHorizontalAlignment('center');
  
  // 데이터 영역 기본값 Right 정렬
  sheet.getRange('A8:O').setHorizontalAlignment('right');
  
  // 실행용 컬럼 숨기기 (13, 14, 15 컬럼)
  // sheet.showColumns(1, 12); // 불필요하게 모든 컬럼을 보일 필요 없음
  sheet.hideColumns(13, 3);
  trimExtraColumns(sheet);
}

// 통합 대시보드 업데이트 (목표 수익률 추가)
function updateDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('📊 대시보드');
  
  if (!sheet) {
    sheet = ss.insertSheet('📊 대시보드', 0);
  }
  
  // 1. 헤더 및 스타일 초기화 (컬럼 추가/변경 대응)
  setupDashboardSheet(sheet);
  
  try {
    ss.toast('대시보드를 업데이트 중입니다...', '⏳ 처리 중', -1);
    
    // 1. 데이터 수집
    const balance = getBalance();
    const holdings = getHoldings();
    const targetPortfolio = getTargetPortfolio();
    
    Logger.log('전체 총평가액: ' + balance.totalEval);
    Logger.log('현재 예수금: ' + balance.cash);
    Logger.log('주문가능금액: ' + balance.buyPower);
    
    // 2. 전체 수익률 계산 (목표 종목만)
    const targetHoldings = holdings.filter(h => targetPortfolio[h.code]);
    
    let totalInvested = 0;
    let totalProfit = 0;
    
    targetHoldings.forEach(h => {
      totalInvested += h.avgPrice * h.quantity;
      totalProfit += h.profit;
    });
    
    const totalReturn = totalInvested > 0 ? (totalProfit / totalInvested) * 100 : 0;
    
    Logger.log('목표 종목 수: ' + targetHoldings.length);
    Logger.log('총 투자금액 (목표 종목만): ' + totalInvested);
    Logger.log('총 손익 (목표 종목만): ' + totalProfit);
    Logger.log('전체 수익률 (목표 종목만): ' + totalReturn.toFixed(2) + '%');
    
    // 3. 실질 운용 자산 계산 (목표외 종목 제외)
    let nonTargetTotal = 0;
    
    holdings.forEach(h => {
      if (!targetPortfolio[h.code]) {
        nonTargetTotal += h.evalAmount;
      }
    });
    
    const totalEval = balance.totalEval;
    const managedTotal = totalEval - nonTargetTotal;
    
    Logger.log(`목표외 자산: ${nonTargetTotal}`);
    Logger.log(`실질 운용 자산(리밸런싱 기준): ${managedTotal}`);
    
    const cashRatio = (balance.cash / managedTotal) * 100;
    
    // 현재 비율 계산 (분모를 managedTotal로)
    const currentRatios = {};
    holdings.forEach(h => {
      if (targetPortfolio[h.code]) {
        currentRatios[h.code] = (h.evalAmount / managedTotal) * 100;
      }
    });
    
    // 4. 요약 정보 업데이트
    sheet.getRange('B3').setValue(totalEval).setNumberFormat('#,##0');
    sheet.getRange('B4').setValue(balance.cash).setNumberFormat('#,##0');

    // 보호 예수금 상태 표시 (C4)
    const protInfo = getProtectedCash();
    const protCell = sheet.getRange('C4');
    if (protInfo.amount > 0) {
      protCell.setValue(`🔒 ${protInfo.amount.toLocaleString()}원 보호중 (${protInfo.daysLeft}일 후 해제)`)
        .setFontColor('#e65100').setFontSize(10).setFontWeight('bold').setBackground('#fff3e0');
    } else {
      protCell.clearContent().setBackground('white').setFontWeight('normal');
    }

    sheet.getRange('B5').setValue(cashRatio.toFixed(2) + '%');
    sheet.getRange('B6').setValue(new Date()).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    
    // 5. 수익률 정보 업데이트
    sheet.getRange('H3').setValue(totalInvested).setNumberFormat('#,##0');
    sheet.getRange('H4').setValue(totalProfit).setNumberFormat('#,##0');
    
    // 5-0. 목표 수익률 값 가져오기 (H6 위치)
    const config = getConfig(); 
    sheet.getRange('H6').setValue((config.targetYield || 0) / 100).setNumberFormat('0.0%');
    
    const profitCell = sheet.getRange('H4');
    if (totalProfit > 0) {
      profitCell.setFontColor('#cc0000').setFontWeight('bold');
    } else if (totalProfit < 0) {
      profitCell.setFontColor('#0000cc').setFontWeight('bold');
    } else {
      profitCell.setFontColor('#000000').setFontWeight('normal');
    }
    
    const returnCell = sheet.getRange('H5');
    returnCell.setValue(totalReturn.toFixed(2) + '%');
    if (totalReturn > 0) {
      returnCell.setFontColor('#cc0000').setFontWeight('bold');
    } else if (totalReturn < 0) {
      returnCell.setFontColor('#0000cc').setFontWeight('bold');
    } else {
      returnCell.setFontColor('#000000').setFontWeight('normal');
    }
    
    // 5-1. 주행 상태 실제 트리거 확인 및 동기화
    const props = PropertiesService.getScriptProperties();
    const triggers = ScriptApp.getProjectTriggers();
    const hasHighwayTrigger = triggers.some(t => t.getHandlerFunction() === 'scheduledBiWeeklyRebalance');
    props.setProperty('HIGHWAY_LANE_KEEPING', hasHighwayTrigger ? 'TRUE' : 'FALSE');

    const statusText = `차선유지:${hasHighwayTrigger ? 'ON' : 'OFF'}`;
    const statusCell = sheet.getRange('B7');
    statusCell.setValue(statusText);
    statusCell.setFontColor(hasHighwayTrigger ? '#137333' : '#5f6368').setFontWeight('bold');

    // 5-1. 추천 인출액 계산 (공식 적용)
    // 월 이론 인출액 = 총자산 × ((1 + r)^(1/12) - 1)
    const currentTargetRate = (config.targetYield || 0) / 100;
    const monthlyRate = Math.pow(1 + currentTargetRate, 1/12) - 1;
    const recommendedAmount = Math.floor(totalEval * monthlyRate);

    sheet.getRange('G7').setValue('💡 월 인출 추천').setFontWeight('bold').setBackground('#e8f5e9'); // 초록색 배경
    sheet.getRange('H7').setValue(recommendedAmount).setNumberFormat('#,##0');
    if (recommendedAmount > 0) {
      sheet.getRange('H7').setFontColor('#137333').setFontWeight('bold');
    } else {
      sheet.getRange('H7').setFontColor('#000000').setFontWeight('normal');
    }
    
    // ========================================
    // 6. 리밸런싱 계산 (통합 함수 호출)
    // ========================================
    
    // 목표 현금 비율 가져오기
    const portfolioSheet = ss.getSheetByName('📋 포트폴리오설정');
    let targetCashRatio = 5;
    if (portfolioSheet) {
      const lastRow = portfolioSheet.getLastRow();
      for (let i = 3; i <= lastRow; i++) {
        if (portfolioSheet.getRange(i, 2).getValue() === '현금') {
          targetCashRatio = parseFloat(portfolioSheet.getRange(i, 4).getValue()) || 5; // D: 운용비율
          break;
        }
      }
    }

    const rebalanceResult = calculateRebalancePlan(managedTotal, balance, holdings, targetPortfolio, {
      targetCashRatio: targetCashRatio,
      tolerance: config.rebalanceTolerance
    });

    const { allStocks, sellOrders, buyOrders, totalSellAmount, totalBuyNeeded, immediateBuyPossible, totalFees, availableCash } = rebalanceResult;
    const rebalanceCount = sellOrders.length + buyOrders.length;
    
    // 8. 리밸런싱 요약 업데이트
    sheet.getRange('E3').setValue(rebalanceCount);
    sheet.getRange('E4').setValue(totalBuyNeeded).setNumberFormat('#,##0');
    sheet.getRange('E5').setValue(totalSellAmount).setNumberFormat('#,##0');
    sheet.getRange('E6').setValue(totalFees).setNumberFormat('#,##0').setFontColor('#ea4335');
    sheet.getRange('E7').setValue(totalSellAmount - totalBuyNeeded).setNumberFormat('#,##0'); // 순수하게 계좌에 남거나 부족한 금액 (실제)
    
    // 즉시 매수 가능 알림 (기존 E5 자리에 표시되던 것을 toast로 더 강조하거나 셀 서식으로 대체 가능)
    if (immediateBuyPossible < totalBuyNeeded) {
      ss.toast(`⚠️ 즉시 매수 가능액은 ${immediateBuyPossible.toLocaleString()}원입니다. 나머지는 매도 체결 후 가능합니다.`, '현금 부족 알림');
    }
    
    // 9. 기존 데이터 삭제 (15개 컬럼 모두 삭제)
    const lastRow = sheet.getLastRow();
    if (lastRow >= 9) { // 헤더가 8행으로 복구되었으므로 9행부터 삭제
      sheet.getRange(9, 1, lastRow - 8, 15).clearContent();
    }
    
    // 10. 데이터 입력
    if (allStocks.length > 0) {
      allStocks.sort((a, b) => b.targetRatio - a.targetRatio);
      
      const data = allStocks.map(s => [
        s.code, s.name, s.type, s.quantity || 0, s.currentPrice || 0, s.evalAmount || 0,
        s.currentRatio.toFixed(2) + '%', s.targetRatio.toFixed(2) + '%', 
        s.expectedRatio.toFixed(2) + '%', s.diff.toFixed(2) + '%',
        s.action, s.actionAmount, s.actionQuantity || 0, s.actionPrice || 0,
        s.avgPrice || 0
      ]);
      
      sheet.getRange(9, 1, data.length, 15).setValues(data);
      sheet.getRange(9, 4, data.length, 1).setNumberFormat('#,##0');
      sheet.getRange(9, 5, data.length, 1).setNumberFormat('#,##0');
      sheet.getRange(9, 6, data.length, 1).setNumberFormat('#,##0');
      sheet.getRange(9, 13, data.length, 1).setNumberFormat('#,##0');
      sheet.getRange(9, 14, data.length, 1).setNumberFormat('#,##0');
      sheet.getRange(9, 15, data.length, 1).setNumberFormat('#,##0'); 
      // sheet.showColumns(1, 12);
      sheet.hideColumns(13, 3); // Hides columns 13, 14, 15
      
      // 색상 적용
      for (let i = 0; i < allStocks.length; i++) {
        const row = 9 + i;
        const stock = allStocks[i];
        const diffCell = sheet.getRange(row, 10); // Column J (Diff)
        const config = getConfig();
        if (Math.abs(stock.diff) > (config.rebalanceTolerance || 2.0)) {
          if (stock.diff > 0) diffCell.setBackground('#fce8e6').setFontColor('#c53929');
          else diffCell.setBackground('#e8f0fe').setFontColor('#1a73e8');
        } else {
          diffCell.setBackground('#e6f4ea').setFontColor('#137333');
        }
        
        const actionCell = sheet.getRange(row, 11); // Column K (Action)
        if (stock.action.includes('매수 (즉시)')) actionCell.setFontColor('#c53929').setFontWeight('bold').setBackground('#fff0f0');
        else if (stock.action.includes('매수 (대기)')) actionCell.setFontColor('#f9ab00').setFontWeight('bold');
        else if (stock.action.includes('매수')) actionCell.setFontColor('#c53929').setFontWeight('bold');
        else if (stock.action.includes('매도')) actionCell.setFontColor('#1a73e8').setFontWeight('bold');
        else if (stock.action.includes('수익실현')) actionCell.setFontColor('#cc0000').setFontWeight('bold');
        else if (stock.action.includes('재배분매도')) actionCell.setFontColor('#e65100').setFontWeight('bold').setBackground('#fff3e0');
        else if (stock.action.includes('신호대기')) actionCell.setFontColor('#616161').setFontWeight('bold').setBackground('#f5f5f5');
        else if (stock.action.includes('유지')) actionCell.setFontColor('#137333');
        else if (stock.action.includes('예산부족')) actionCell.setFontColor('#f9ab00').setFontWeight('bold');
        else if (stock.action.includes('가격오류')) actionCell.setFontColor('#ff00ff').setFontWeight('bold');
      }
    }
    
    if (buyOrders.length > 0) {
      Logger.log('=== 최종 매수 주문 목록 ===');
      buyOrders.forEach(order => Logger.log(`${order.name}: ${order.quantity}주 x ${order.price.toLocaleString()}원 = ${order.amount.toLocaleString()}원`));
    }
    
    const cashStatus = availableCash <= 0 
      ? `💵 현금 ${cashRatio.toFixed(1)}% (목표 ${targetCashRatio}% 유지 중)` 
      : `🛒 매수: ${buyOrders.length}종목 ${totalBuyNeeded.toLocaleString()}원`;
    
    ss.toast(`💰 운용자산: ${managedTotal.toLocaleString()}원 | 📈 수익률: ${totalReturn >= 0 ? '+' : ''}${totalReturn.toFixed(2)}% | ` + cashStatus, '✅ 대시보드 업데이트 완료!', 10);
    
  } catch (e) {
    Logger.log('updateDashboard 오류: ' + e.toString());
    Logger.log('오류 스택: ' + e.stack);
    ss.toast('오류: ' + e.message, '❌ 대시보드 업데이트 실패', 10);
  }
}

/**
 * 리밸런싱 계획 계산 통합 함수 (UI/자동화 공용)
 */
function calculateRebalancePlan(managedTotal, balance, holdings, targetPortfolio, options = {}) {
  const config = getConfig();
  const tolerance = options.tolerance || config.rebalanceTolerance || 2.0;
  const profitTakingThreshold = options.profitTakingThreshold || config.profitTakingThreshold || 40.0;
  const useTAForQty = false;
  const REDISTRIBUTE_THRESHOLD = 0.15;
  const allStocks = [];
  const priceCache = {};
  const scoreCache = {};

  holdings.forEach(h => priceCache[h.code] = h.currentPrice);
  Object.keys(targetPortfolio).forEach(code => {
    if (!priceCache[code] || priceCache[code] === 0) {
      priceCache[code] = getCurrentPrice(code);
      if(priceCache[code] > 0) Utilities.sleep(100);
    }
    scoreCache[code] = { score: 0, summary: '중립' };
  });

  const currentRatios = {};
  holdings.forEach(h => {
    if (targetPortfolio[h.code]) currentRatios[h.code] = (h.evalAmount / managedTotal) * 100;
  });

  let totalSellAmount = 0;
  const sellOrders = [];
  Object.keys(targetPortfolio).forEach(code => {
    const target = targetPortfolio[code];
    const holding = holdings.find(h => h.code === code);
    const current = currentRatios[code] || 0;
    const diff = target.ratio - current;
    const ta = scoreCache[code] || { score: 0, summary: '중립' };

    // 수익률 기반 공격적 매도 여부 판단
    const isHighProfit = holding && (holding.profitRate * 100 >= profitTakingThreshold);
    const baseTolerance = isHighProfit ? 0.1 : tolerance;

    // 기술적 지표 기반 매도 tolerance 축소 (TA 조정 활성 시에만)
    const sellToleranceMult = useTAForQty ? getSellToleranceMultiplier(fsdMode, ta.score) : 1.0;
    const effectiveTolerance = baseTolerance * sellToleranceMult;

    // 수익 재배분 매도: TA 조정 활성 + score < -0.5 AND 수익률 >= 15%
    const isRedistribute = !!(useTAForQty && holding && ta.score < -0.5 && holding.profitRate >= REDISTRIBUTE_THRESHOLD);

    // 매도 조건: 비중 초과 OR 재배분 조건
    const shouldSell = (diff < -effectiveTolerance && holding) || (isRedistribute && holding);

    if (shouldSell) {
      const availableToSell = holding.ordPsblQty !== undefined ? holding.ordPsblQty : holding.quantity;
      if (availableToSell <= 0) return;

      let targetAmount;
      if (isRedistribute && diff >= -effectiveTolerance) {
        // 비중 초과가 아닌 순수 재배분 매도: 목표의 redistributeRatio까지 추가 매도
        const redistributeRatio = getRedistributeRatio(fsdMode);
        targetAmount = managedTotal * (target.ratio / 100) * redistributeRatio;
      } else {
        targetAmount = managedTotal * (target.ratio / 100);
      }

      const excessAmount = holding.evalAmount - targetAmount;
      const currentPrice = holding.currentPrice;
      if (currentPrice > 0 && excessAmount > 0) {
        let quantity = Math.round(excessAmount / currentPrice);
        quantity = Math.min(quantity, availableToSell);

        if (quantity > 0) {
          const rawSellAmount = quantity * currentPrice;
          const actualSellProceeds = Math.floor(rawSellAmount * (1 - config.sellFeeRate));

          totalSellAmount += actualSellProceeds;
          sellOrders.push({
            code, name: target.name, type: target.type, quantity, price: currentPrice,
            amount: rawSellAmount,
            actualProceeds: actualSellProceeds,
            isProfitTaking: isHighProfit,
            isRedistribute,
            taScore: ta.score,
            avgPrice: holding.avgPrice
          });
        }
      }
    }
  });

  const targetCashRatio = options.targetCashRatio || 5;
  // 보호 예수금 차감 (수익실현 후 2주 보호 기간 동안 리밸런싱에서 제외)
  const protectedCashInfo = getProtectedCash();
  const effectiveBuyPower = Math.max(0, balance.buyPower - protectedCashInfo.amount);

  const actualCash = effectiveBuyPower + totalSellAmount; // 매도 후 가용 현금
  const targetCashAmount = managedTotal * (targetCashRatio / 100); // 목표 현금 보유액 (managedTotal은 이미 현금 포함 총액)

  // 리밸런싱용 가용 현금 (목표 현금을 제외한 여유분)
  let tradableCash = actualCash - targetCashAmount;
  let availableCash = Math.max(0, tradableCash);

  const buyTargets = [];
  Object.keys(targetPortfolio).forEach(code => {
    const target = targetPortfolio[code];
    const holding = holdings.find(h => h.code === code);
    const currentAmount = holding?.evalAmount || 0;

    // 이 종목이 목표로 하는 금액
    const newTargetAmount = managedTotal * (target.ratio / 100);
    const shortfall = newTargetAmount - currentAmount;
    const diffPct = target.ratio - (currentAmount / managedTotal * 100);

    // 매도와 동일하게 tolerance 이상 차이날 때만 매수 대상에 추가
    if (shortfall > 0 && diffPct > tolerance) {
       buyTargets.push({
         code, name: target.name, type: target.type, shortfall,
         diff: diffPct
       });
    }
  });

  let totalBuyNeeded = 0;
  let immediateBuyPossible = 0;
  const buyOrders = [];
  
  // 목표외 보유 종목 전량 매도 주문
  holdings.forEach(h => {
    if (targetPortfolio[h.code]) return;
    const availableToSell = h.ordPsblQty !== undefined ? h.ordPsblQty : h.quantity;
    if (availableToSell <= 0) return;
    const price = priceCache[h.code] || h.currentPrice;
    const rawAmount = availableToSell * price;
    const actualProceeds = Math.floor(rawAmount * (1 - config.sellFeeRate));
    totalSellAmount += actualProceeds;
    sellOrders.push({
      code: h.code, name: h.name, type: h.type || '기타',
      quantity: availableToSell, price,
      amount: rawAmount, actualProceeds,
      isRedistribute: false, isProfitTaking: false, taScore: 0, sellMult: 1
    });
  });

  // 1. 현재 즉시 가용 현금 (목표 현금 제외, 단 1주 매수 가능하도록 목표 현금의 20% 여유 허용)
  const cashSlack = targetCashAmount * 0.20;
  let currentTradableCash = Math.max(0, effectiveBuyPower - targetCashAmount + cashSlack);
  // 2. 매도 성공 시 추가될 가용 현금
  let pendingSellCash = totalSellAmount;

  if (buyTargets.length > 0) {
    // 부족한 비중(diff)이 큰 순서대로 정렬 (우선순위)
    buyTargets.sort((a, b) => b.diff - a.diff);

    buyTargets.forEach(target => {
      let currentPrice = priceCache[target.code];
      if (currentPrice > 0) {
        const idealQty = Math.round(target.shortfall / currentPrice);
        if (idealQty <= 0) return;

        // FSD 모드 × 기술적 지표 → 매수 배율 적용 (TA 조정 활성 시에만)
        const ta = scoreCache[target.code] || { score: 0, summary: '중립' };
        const buyMult = useTAForQty ? getBuyMultiplier(fsdMode, ta.score) : 1.0;
        if (buyMult === 0) {
          Logger.log('[TA] ' + target.code + ' 매수 건너뜀 (score=' + ta.score + ', mode=' + fsdMode + ')');
          return;
        }
        const adjustedIdealQty = useTAForQty ? Math.max(1, Math.round(idealQty * buyMult)) : idealQty;

        // 전체 가용 현금 내 최대 수량 (현재 현금 + 매도 예정 현금)
        const totalAvailableNow = currentTradableCash + pendingSellCash;
        const maxCanBuyTotal = Math.floor(totalAvailableNow / (currentPrice * (1 + config.buyFeeRate)));

        // 실제 수량 결정 (adjustedIdealQty 기준, 수수료 포함)
        const quantity = Math.min(adjustedIdealQty, maxCanBuyTotal);
        
        if (quantity > 0) {
          const rawBuyAmount = quantity * currentPrice;
          const buyAmountWithFee = Math.ceil(rawBuyAmount * (1 + config.buyFeeRate));
          
          // 이 매수 주문이 현재 현금으로 즉시 가능한지 판별
          let isImmediate = false;
          let immediateQty = 0;
          
          if (currentTradableCash >= currentPrice * (1 + config.buyFeeRate)) {
            // 현재 현금으로 살 수 있는 수량 계산
            immediateQty = Math.min(quantity, Math.floor(currentTradableCash / (currentPrice * (1 + config.buyFeeRate))));
            if (immediateQty > 0) {
              const immediateAmountWithFee = Math.ceil(immediateQty * currentPrice * (1 + config.buyFeeRate));
              immediateBuyPossible += (immediateQty * currentPrice); // 화면 표시용은 원금
              currentTradableCash -= immediateAmountWithFee;
              if (immediateQty === quantity) isImmediate = true;
            }
          }
          
          // 남은 대기 자금 차감 (매도 완료 후 살 수 있는 몫 - 수수료 포함)
          const pendingNeededForThis = buyAmountWithFee - Math.ceil(immediateQty * currentPrice * (1 + config.buyFeeRate));
          pendingSellCash -= pendingNeededForThis;

          totalBuyNeeded += rawBuyAmount; // 화면 표시용 원금
          buyOrders.push({
            code: target.code, name: target.name, type: target.type,
            quantity, idealQty, adjustedIdealQty, price: currentPrice, amount: rawBuyAmount,
            amountWithFee: buyAmountWithFee,
            isImmediate, immediateQty,
            taScore: ta.score, buyMult
          });
        }
      }
    });
  }

  Object.keys(targetPortfolio).forEach(code => {
    const target = targetPortfolio[code];
    const holding = holdings.find(h => h.code === code);
    const current = currentRatios[code] || 0;
    const diff = target.ratio - current;
    let action = '✅ 유지', actionQuantity = 0, actionPrice = 0, actionAmount = '', needsRebalance = false;

    const sellOrder = sellOrders.find(o => String(o.code) === String(code));
    if (sellOrder) {
      action = sellOrder.isRedistribute && !sellOrder.isProfitTaking ? '💸 재배분매도' : sellOrder.isProfitTaking ? '💰 수익실현' : '📉 매도';
      actionQuantity = sellOrder.quantity;
      actionPrice = sellOrder.price;
      const formattedAmount = String(sellOrder.amount).replace(/\B(?=(\d{3})+(?!\d))/g, ",");
      const scoreStr = sellOrder.taScore !== 0 ? (sellOrder.taScore > 0 ? ', ▲' : ', ▼') + Math.abs(sellOrder.taScore) : '';
      actionAmount = sellOrder.quantity + "주 (" + formattedAmount + "원" + scoreStr + ")";
      needsRebalance = true;
    }
    const buyOrder = buyOrders.find(o => String(o.code) === String(code));
    if (buyOrder) {
      if (buyOrder.isImmediate) {
        action = '📈 매수 (즉시)';
      } else if (buyOrder.immediateQty > 0) {
        action = '📈 매수 (일부즉시)';
      } else {
        action = '📈 매수 (대기)';
      }

      actionQuantity = buyOrder.quantity;
      actionPrice = buyOrder.price;
      const amtStr = String(buyOrder.amount);
      let formattedAmount = '';
      for (let i = 0; i < amtStr.length; i++) {
        if (i > 0 && (amtStr.length - i) % 3 === 0) formattedAmount += ',';
        formattedAmount += amtStr[i];
      }

      const scoreStr = buyOrder.taScore !== 0 ? (buyOrder.taScore > 0 ? ' ▲' : ' ▼') + Math.abs(buyOrder.taScore) : '';
      // 배율로 조정된 경우 원래 idealQty도 표시
      const isAdjusted = buyOrder.adjustedIdealQty && buyOrder.adjustedIdealQty < buyOrder.idealQty;
      const qtyStr = isAdjusted
        ? `${buyOrder.quantity}주 (목표 ${buyOrder.idealQty}주→${buyOrder.adjustedIdealQty}주${scoreStr})`
        : `${buyOrder.quantity}주 (목표 ${buyOrder.idealQty}주${scoreStr})`;

      if (buyOrder.immediateQty > 0 && !buyOrder.isImmediate) {
        actionAmount = qtyStr + ', 즉시 ' + buyOrder.immediateQty + '주';
      } else {
        actionAmount = qtyStr;
      }
      needsRebalance = true;
    }
    
    // 매수/매도 주문이 나가지 않은 경우 상태 결정
    if (!sellOrder && !buyOrder) {
      const checkPrice = priceCache[code];
      if (!checkPrice || checkPrice <= 0) {
        action = '❌ 가격오류'; actionAmount = '현재가 확인필요';
      } else if (diff > tolerance) {
        // 리밸런싱 기준 이상 부족한데 주문이 안 나간 경우
        const ta = scoreCache[code] || { score: 0 };
        const buyMult = useTAForQty ? getBuyMultiplier(fsdMode, ta.score) : 1.0;
        if (useTAForQty && buyMult === 0) {
          // 기술적 신호로 인해 매수 보류
          action = '⏸️ 신호대기';
          actionAmount = '매수보류 (▼' + Math.abs(ta.score) + ')';
        } else {
          const targetInBuyList = buyTargets.find(t => t.code === code);
          if (targetInBuyList) {
            const currentRemaining = Math.max(0, tradableCash) - totalBuyNeeded;
            if (currentRemaining < checkPrice || (targetInBuyList.shortfall < checkPrice * 0.5)) {
              action = '⚠️ 예산부족';
              actionAmount = `1주 ${checkPrice.toLocaleString()}원`;
            }
          }
        }
      } else if (actualCash <= 1000) {
        action = '✅ 유지';
      }
    }
    // 예상 비중 계산 (매수/매도 시뮬레이션 결과 반영)
    let expectedAmount = holding?.evalAmount || 0;
    if (sellOrder) expectedAmount -= sellOrder.amount;
    if (buyOrder) expectedAmount += buyOrder.amount;
    const expectedRatio = (expectedAmount / managedTotal) * 100;

    allStocks.push({
      code, name: target.name, type: target.type,
      quantity: holding?.quantity || 0, currentPrice: priceCache[code] || 0,
      evalAmount: holding?.evalAmount || 0,
      avgPrice: holding?.avgPrice || 0,
      currentRatio: current, targetRatio: target.ratio,
      expectedRatio, diff, action, actionQuantity, actionPrice, actionAmount, needsRebalance
    });
  });

  // 목표외 보유 종목 대시보드 표시
  holdings.forEach(h => {
    if (targetPortfolio[h.code]) return;
    const current = (h.evalAmount / managedTotal) * 100;
    const sellOrder = sellOrders.find(o => String(o.code) === String(h.code));
    let action = '📉 청산필요';
    let actionAmount = `전량 ${h.quantity}주`;
    if (sellOrder) {
      action = '📉 매도';
      actionAmount = sellOrder.quantity + '주 (' + String(sellOrder.amount).replace(/\B(?=(\d{3})+(?!\d))/g, ',') + '원)';
    }
    allStocks.push({
      code: h.code, name: h.name, type: h.type || '기타',
      quantity: h.quantity, currentPrice: h.currentPrice,
      evalAmount: h.evalAmount, avgPrice: h.avgPrice || 0,
      currentRatio: current, targetRatio: 0,
      expectedRatio: sellOrder ? 0 : current,
      diff: -current, action, actionQuantity: sellOrder?.quantity || 0,
      actionPrice: sellOrder?.price || h.currentPrice, actionAmount, needsRebalance: true
    });
  });

  let totalFees = 0;
  sellOrders.forEach(o => totalFees += (o.amount - o.actualProceeds));
  buyOrders.forEach(o => totalFees += (o.amountWithFee - o.amount));

  return { sellOrders, buyOrders, allStocks, totalSellAmount, totalBuyNeeded, immediateBuyPossible, availableCash, totalFees, targetCashRatio };
}

// 대시보드에서 리밸런싱 실행
function executeRebalanceFromDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName('📊 대시보드');
  
  if (!sheet) {
    ss.toast('먼저 "대시보드 새로고침"을 실행하세요.', '⚠️ 대시보드 없음', 5);
    return;
  }
  
  const response = ui.alert(
    '리밸런싱 실행',
    '실제로 주문을 실행하시겠습니까?\n\n⚠️ 이 작업은 실제 매매를 진행합니다.',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    ss.toast('리밸런싱이 취소되었습니다.', 'ℹ️ 취소', 3);
    return;
  }
  
  const lastRow = sheet.getLastRow();

  if (lastRow < 9) {
    ss.toast('먼저 "대시보드 새로고침"을 실행하세요.', '⚠️ 항목 없음', 5);
    return;
  }

  // 동시 실행 방지: 자동 스케줄러와 충돌하지 않도록 Lock 획득
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    ss.toast('자동 리밸런싱이 실행 중입니다. 잠시 후 다시 시도하세요.', '⚠️ 실행 중', 5);
    return;
  }

  try {
  ss.toast('리밸런싱을 실행 중입니다...', '⏳ 처리 중', -1);

  const results = [];
  const logSheet = ss.getSheetByName('📝 거래내역');
  if (!logSheet) {
    ss.toast('⚠️ "📝 거래내역" 시트를 찾을 수 없습니다. 시트명을 확인하세요.', '❌ 오류', 10);
    return;
  }
  
  // 매도/매수 주문 분리 수집
  const sellList = [];
  const buyList = [];
  
  for (let i = 9; i <= lastRow; i++) {
    const action = String(sheet.getRange(i, 11).getValue()).trim(); // Column K
    const code = String(sheet.getRange(i, 1).getValue()).trim();   // Column A
    const name = String(sheet.getRange(i, 2).getValue()).trim();   // Column B
    const quantityRaw = sheet.getRange(i, 13).getValue();          // Column M (Hidden)
    const priceRaw = sheet.getRange(i, 14).getValue();             // Column N (Hidden)
    
    if (!action.includes('매수') && !action.includes('매도') && !action.includes('수익실현')) continue;
    if (!code || !quantityRaw) continue;
    
    const quantity = parseInt(quantityRaw);
    const price = parseInt(priceRaw) || 0;
    const avgPrice = parseFloat(sheet.getRange(i, 15).getValue()) || 0; // Column O
    
    if (isNaN(quantity) || quantity <= 0) continue;
    
    const order = { code, name, quantity, price, action, avgPrice };
    
    if (action.includes('매도') || action.includes('수익실현')) {
      sellList.push(order);
    } else {
      buyList.push(order);
    }
  }
  
  Logger.log(`매도 ${sellList.length}건, 매수 ${buyList.length}건`);
  
  // 매도 먼저 실행
  for (const order of sellList) {
    ss.toast(`${order.name} 매도 주문 중...`, '📉 처리 중', 2);
    
    const result = placeOrder(order.code, 'sell', order.quantity, 0);
    
    results.push({
      action: '매도',
      code: order.code,
      name: order.name,
      quantity: order.quantity,
      price: order.price,
      success: result.success,
      message: result.message
    });
    
    logSheet.appendRow([
      new Date(), '매도', order.code, order.name,
      order.quantity, order.price, order.quantity * order.price,
      result.success ? '성공' : '실패', result.message
    ]);
    trimExtraColumns(logSheet, 9);

    // 수익 실현 기록 (매도 성공 시)
    if (result.success) {
      let profitSheet = ss.getSheetByName('📝 수익실현기록') || ss.getSheetByName('📝 수익 실현 기록');
      if (!profitSheet) {
        ss.toast('수익실현기록 시트를 찾는 중...', '🔍 확인 중');
        // 시트가 없으면 생성 시도
        try { setupSheets(); profitSheet = ss.getSheetByName('📝 수익실현기록'); } catch(e) {}
      }
      
      if (profitSheet) {
        const realizedProfit = (order.price - order.avgPrice) * order.quantity;
        const lastProfitRow = profitSheet.getLastRow();
        let cumulativeProfit = realizedProfit;
        if (lastProfitRow >= 1) {
          const lastVal = profitSheet.getRange(lastProfitRow, 8).getValue();
          cumulativeProfit += (typeof lastVal === 'number' ? lastVal : 0);
        }
        
        profitSheet.appendRow([
          new Date(), order.name, order.action, order.quantity, order.price, 
          order.quantity * order.price, realizedProfit, cumulativeProfit
        ]);
        
        // 포맷팅
        const newRow = profitSheet.getLastRow();
        profitSheet.getRange(newRow, 6, 1, 3).setNumberFormat('#,##0');
        profitSheet.getRange(newRow, 7, 1, 2).setFontColor(realizedProfit >= 0 ? '#d93025' : '#1967d2');
      }
    }
    
    Utilities.sleep(500);
  }
  
  // 매수 실행
  for (const order of buyList) {
    ss.toast(`${order.name} 매수 주문 중...`, '📈 처리 중', 2);
    
    const result = placeOrder(order.code, 'buy', order.quantity, 0);
    
    results.push({
      action: '매수',
      code: order.code,
      name: order.name,
      quantity: order.quantity,
      price: order.price,
      success: result.success,
      message: result.message
    });
    
    logSheet.appendRow([
      new Date(), '매수', order.code, order.name,
      order.quantity, order.price, order.quantity * order.price,
      result.success ? '성공' : '실패', result.message
    ]);
    trimExtraColumns(logSheet, 9);
    
    Utilities.sleep(500);
  }
  
  // 결과 표시
  const successCount = results.filter(r => r.success).length;
  const failCount = results.filter(r => !r.success).length;
  
  ss.toast(
    `✅ 성공: ${successCount}건 | ❌ 실패: ${failCount}건`,
    '🔄 리밸런싱 실행 완료',
    10
  );
  
  if (successCount > 0) {
    // KIS API가 체결 결과를 반영하는 데 시간이 걸림 (2초는 부족, 8초로 여유 확보)
    ss.toast('체결 확인 중... (8초 대기)', '⏳ 갱신 대기', 10);
    Utilities.sleep(8000);
    updateDashboard();
  }

  } finally {
    lock.releaseLock();
  }
}

/**
 * UI 없이 리밸런싱 실행 (자동화용)
 */
function executeRebalanceSilently() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('📊 대시보드');
  if (!sheet) return;
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 9) return;
  
  const sellList = [];
  const buyList = [];
  
  for (let i = 9; i <= lastRow; i++) {
    const action = String(sheet.getRange(i, 11).getValue()).trim();
    const code = String(sheet.getRange(i, 1).getValue()).trim();
    const name = String(sheet.getRange(i, 2).getValue()).trim();
    const quantityRaw = sheet.getRange(i, 13).getValue();
    const priceRaw = sheet.getRange(i, 14).getValue();
    const avgPrice = parseFloat(sheet.getRange(i, 15).getValue()) || 0;
    
    if (!action.includes('매수') && !action.includes('매도') && !action.includes('수익실현')) continue;
    if (!code || !quantityRaw) continue;
    
    const quantity = parseInt(quantityRaw);
    if (isNaN(quantity) || quantity <= 0) continue;
    
    const order = { code, name, quantity, price: parseInt(priceRaw) || 0, action, avgPrice };
    if (action.includes('매도') || action.includes('수익실현')) sellList.push(order);
    else buyList.push(order);
  }
  
  return performRebalanceOrders(sellList, buyList);
}

/**
 * 수집된 주문 목록을 실제로 한투 API에 전송
 */
function performRebalanceOrders(sellList, buyList) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName('📝 거래내역');
  if (!logSheet) {
    Logger.log('[오류] "📝 거래내역" 시트를 찾을 수 없어 주문을 중단합니다.');
    return 0;
  }
  const profitSheet = ss.getSheetByName('📝 수익실현기록');
  let successCount = 0;

  // 매도 우선
  sellList.forEach(order => {
    const result = placeOrder(order.code, 'sell', order.quantity, 0);
    if (result.success) {
      successCount++;
      if (profitSheet) {
        const realizedProfit = (order.price - order.avgPrice) * order.quantity;
        const lastRow = profitSheet.getLastRow();
        let cumulative = realizedProfit + (lastRow >= 1 ? (parseFloat(profitSheet.getRange(lastRow, 8).getValue()) || 0) : 0);
        profitSheet.appendRow([new Date(), order.name, order.action, order.quantity, order.price, order.quantity * order.price, realizedProfit, cumulative]);
      }
    }
    logSheet.appendRow([new Date(), '매도', order.code, order.name, order.quantity, order.price, order.quantity * order.price, result.success ? '성공' : '실패', result.message]);
    trimExtraColumns(logSheet, 9);
    Utilities.sleep(500);
  });

  // 매수
  buyList.forEach(order => {
    const result = placeOrder(order.code, 'buy', order.quantity, 0);
    if (result.success) successCount++;
    logSheet.appendRow([new Date(), '매수', order.code, order.name, order.quantity, order.price, order.quantity * order.price, result.success ? '성공' : '실패', result.message]);
    trimExtraColumns(logSheet, 9);
    Utilities.sleep(500);
  });
  
  if (successCount > 0) {
    Utilities.sleep(8000);
    updateDashboard();
  }
  return successCount;
}



// 거래내역 기록
function logTrades(results) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('📝 거래내역');
  
  if (results.length === 0) return;
  
  const data = results.map(r => [
    r.time,
    r.type,
    r.code,
    r.name,
    r.quantity,
    r.price,
    r.amount,
    r.status,
    r.message
  ]);
  
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, data.length, 9).setValues(data);
 }

/**
 * 트리거용 자동 새로고침 루틴 (onOpen 시 실행)
 * 대시보드와 계좌 현황을 최신 상태로 유지하며, 알림(Toast)을 표시합니다.
 */
function automatedRefreshRoutine() {
  try {
    Logger.log('자동 새로고침 시작 (대시보드 & 계좌현황)...');
    updateDashboard();    // 대시보드 업데이트 (알림 포함)
    SpreadsheetApp.flush(); // UI(시트 및 토스트 팝업) 즉시 반영
    Utilities.sleep(5000); // 5초 대기 (알림 확인용)
    updateAccountSheet(); // 계좌현황 업데이트 (알림 포함)
    Logger.log('자동 새로고침 완료.');
  } catch (e) {
    Logger.log('자동 새로고침 오류: ' + e.toString());
  }
}
