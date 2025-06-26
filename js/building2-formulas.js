// ===== LG Comp List 수식 처리 =====

// 수식 적용 메인 함수
function applyLGFormulas(worksheet, col) {
    // 1. 면적 변환 수식 (평 → m²)
    applyAreaConversionFormulas(worksheet, col);
    
    // 2. 임대 기준 계산 수식
    applyRentCalculationFormulas(worksheet, col);
    
    // 3. 임대기준 조정 수식
    applyAdjustedRentFormulas(worksheet, col);
    
    // 4. 예상비용 수식
    applyExpectedCostFormulas(worksheet, col);
}

// 면적 변환 수식 적용
function applyAreaConversionFormulas(worksheet, col) {
    // 28행: 임대면적 (m²) = 평 × 3.305785
    const cell28 = worksheet.getCell(`${col}28`);
    cell28.value = { formula: LG_TEMPLATE_CONFIG.formulas.pyToM2(col, 30) };
    cell28.numFmt = LG_TEMPLATE_CONFIG.numberFormats.area_m2;
    applyFormulaStyle(cell28);
    
    // 29행: 전용면적 (m²) = 평 × 3.305785
    const cell29 = worksheet.getCell(`${col}29`);
    cell29.value = { formula: LG_TEMPLATE_CONFIG.formulas.pyToM2(col, 31) };
    cell29.numFmt = LG_TEMPLATE_CONFIG.numberFormats.area_m2;
    applyFormulaStyle(cell29);
}

// 임대 기준 계산 수식 적용
function applyRentCalculationFormulas(worksheet, col) {
    // 35행: 월 평당 지출비용 = 임대료 + 관리비
    const cell35 = worksheet.getCell(`${col}35`);
    cell35.value = { formula: LG_TEMPLATE_CONFIG.formulas.monthlyExpense(col) };
    cell35.numFmt = LG_TEMPLATE_CONFIG.numberFormats.currency;
    applyFormulaStyle(cell35, 'right');
    
    // 36행: 총 보증금 = 평당 보증금 × 임대면적(평)
    const cell36 = worksheet.getCell(`${col}36`);
    cell36.value = { formula: LG_TEMPLATE_CONFIG.formulas.totalDeposit(col) };
    cell36.numFmt = LG_TEMPLATE_CONFIG.numberFormats.currency;
    applyFormulaStyle(cell36, 'right');
    
    // 37행: 월 임대료 총액 = 평당 임대료 × 임대면적(평)
    const cell37 = worksheet.getCell(`${col}37`);
    cell37.value = { formula: LG_TEMPLATE_CONFIG.formulas.totalRent(col) };
    cell37.numFmt = LG_TEMPLATE_CONFIG.numberFormats.currency;
    applyFormulaStyle(cell37, 'right');
    
    // 38행: 월 관리비 총액 = 평당 관리비 × 임대면적(평)
    const cell38 = worksheet.getCell(`${col}38`);
    cell38.value = { formula: LG_TEMPLATE_CONFIG.formulas.totalManagement(col) };
    cell38.numFmt = LG_TEMPLATE_CONFIG.numberFormats.currency;
    applyFormulaStyle(cell38, 'right');
    
    // 39행: 월 전용면적당 지출비용 = (임대료총액 + 관리비총액) / 전용면적(평)
    const cell39 = worksheet.getCell(`${col}39`);
    cell39.value = { formula: LG_TEMPLATE_CONFIG.formulas.expensePerDedicated(col) };
    cell39.numFmt = LG_TEMPLATE_CONFIG.numberFormats.currency;
    applyFormulaStyle(cell39, 'right');
}

// 임대기준 조정 수식 적용
function applyAdjustedRentFormulas(worksheet, col) {
    // 40행: 보증금 = 32행과 동일
    const cell40 = worksheet.getCell(`${col}40`);
    cell40.value = { formula: `${col}32` };
    cell40.numFmt = LG_TEMPLATE_CONFIG.numberFormats.currency;
    applyFormulaStyle(cell40, 'right');
    
    // 42행: 평균 임대료 = 임대료 - (임대료 × 렌트프리개월 / 12)
    const cell42 = worksheet.getCell(`${col}42`);
    cell42.value = { formula: LG_TEMPLATE_CONFIG.formulas.avgRent(col) };
    cell42.numFmt = LG_TEMPLATE_CONFIG.numberFormats.currency;
    applyFormulaStyle(cell42, 'right');
    
    // 43행: 관리비 = 34행과 동일
    const cell43 = worksheet.getCell(`${col}43`);
    cell43.value = { formula: `${col}34` };
    cell43.numFmt = LG_TEMPLATE_CONFIG.numberFormats.currency;
    applyFormulaStyle(cell43, 'right');
    
    // 44행: NOC = (평균임대료 + 관리비) × (임대면적/전용면적)
    const cell44 = worksheet.getCell(`${col}44`);
    cell44.value = { formula: LG_TEMPLATE_CONFIG.formulas.noc(col) };
    cell44.numFmt = LG_TEMPLATE_CONFIG.numberFormats.currency;
    applyFormulaStyle(cell44, 'center');
}

// 예상비용 수식 적용
function applyExpectedCostFormulas(worksheet, col) {
    // 46행: 보증금 = 조정된 보증금 × 임대면적(평)
    const cell46 = worksheet.getCell(`${col}46`);
    cell46.value = { formula: LG_TEMPLATE_CONFIG.formulas.expectedDeposit(col) };
    cell46.numFmt = LG_TEMPLATE_CONFIG.numberFormats.currency;
    applyFormulaStyle(cell46, 'right');
    
    // 47행: 평균 월 임대료 = 평균 임대료 × 임대면적(평)
    const cell47 = worksheet.getCell(`${col}47`);
    cell47.value = { formula: LG_TEMPLATE_CONFIG.formulas.expectedMonthlyRent(col) };
    cell47.numFmt = LG_TEMPLATE_CONFIG.numberFormats.currency;
    applyFormulaStyle(cell47, 'right');
    
    // 48행: 평균 월 관리비 = 관리비 × 임대면적(평)
    const cell48 = worksheet.getCell(`${col}48`);
    cell48.value = { formula: LG_TEMPLATE_CONFIG.formulas.expectedMonthlyManagement(col) };
    cell48.numFmt = LG_TEMPLATE_CONFIG.numberFormats.currency;
    applyFormulaStyle(cell48, 'right');
    
    // 49행: 월 (임대료 + 관리비) = 47행 + 48행
    const cell49 = worksheet.getCell(`${col}49`);
    cell49.value = { formula: LG_TEMPLATE_CONFIG.formulas.monthlyTotal(col) };
    cell49.numFmt = LG_TEMPLATE_CONFIG.numberFormats.currency;
    applyFormulaStyle(cell49, 'center');
    
    // 50행: 연 실제 부담 고정금액 = 49행 × 12
    const cell50 = worksheet.getCell(`${col}50`);
    cell50.value = { formula: LG_TEMPLATE_CONFIG.formulas.yearlyTotal(col) };
    cell50.numFmt = LG_TEMPLATE_CONFIG.numberFormats.currency;
    applyFormulaStyle(cell50, 'center');
}

// 수식 셀 스타일 적용
function applyFormulaStyle(cell, alignment = 'center') {
    // LG Smart Regular 폰트 강제 적용
    cell.font = {
        name: 'LG Smart Regular',
        size: 9
    };
    
    // 가운데 정렬 (A1-A4 제외)
    const cellRef = cell.address;
    if (!['A1', 'A2', 'A3', 'A4'].includes(cellRef)) {
        cell.alignment = {
            horizontal: alignment,
            vertical: 'middle'
        };
    }
    
    // 테두리
    cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };
}

// 수식 검증 함수
function validateFormulas(worksheet, col) {
    const errors = [];
    const formulaCells = [28, 29, 35, 36, 37, 38, 39, 40, 42, 43, 44, 46, 47, 48, 49, 50];
    
    formulaCells.forEach(row => {
        const cell = worksheet.getCell(`${col}${row}`);
        if (!cell.formula && !cell.value) {
            errors.push(`${col}${row}: 수식이 없습니다.`);
        }
    });
    
    return {
        isValid: errors.length === 0,
        errors: errors
    };
}

// 조건부 서식 적용 (선택사항)
function applyConditionalFormatting(worksheet, col) {
    // 예: NOC가 특정 값 이상이면 색상 변경
    const nocCell = worksheet.getCell(`${col}44`);
    
    // 연 실제 부담 금액이 높으면 강조
    const yearlyCell = worksheet.getCell(`${col}50`);
    
    // 이 부분은 필요에 따라 구현
}

// 수식 복사 함수 (여러 빌딩에 동일 수식 적용)
function copyFormulas(worksheet, fromCol, toCol) {
    const formulaRows = [28, 29, 35, 36, 37, 38, 39, 40, 42, 43, 44, 46, 47, 48, 49, 50];
    
    formulaRows.forEach(row => {
        const fromCell = worksheet.getCell(`${fromCol}${row}`);
        const toCell = worksheet.getCell(`${toCol}${row}`);
        
        if (fromCell.formula) {
            // 수식의 열 참조를 변경
            const adjustedFormula = fromCell.formula.replace(
                new RegExp(fromCol, 'g'), 
                toCol
            );
            toCell.value = { formula: adjustedFormula };
            toCell.numFmt = fromCell.numFmt;
            
            // 스타일 복사
            toCell.font = { ...fromCell.font };
            toCell.alignment = { ...fromCell.alignment };
            toCell.border = { ...fromCell.border };
        }
    });
}

// 수식 재계산 트리거 (필요시)
function recalculateFormulas(worksheet) {
    // ExcelJS는 자동으로 수식을 계산하므로 일반적으로 필요없음
    // 단, 강제 재계산이 필요한 경우 사용
// 전역 함수로 등록
window.applyLGFormulas = applyLGFormulas;