// ===== LG Comp List 데이터 처리 =====

// 빌딩 데이터를 워크시트에 입력 (LG 버전)
function fillBuildingDataLG(worksheet, building, buildingIndex) {
    if (buildingIndex >= 6) return; // 최대 6개
    
    const startCol = LG_TEMPLATE_CONFIG.buildingColumns[buildingIndex];
    const colIndex = LG_UTILS.getColumnIndex(startCol);
    
    // 기초정보 입력
    fillBasicInfo(worksheet, building, startCol);
    
    // 공실 현황 입력 (임시 - 실제로는 별도 데이터 필요)
    fillVacancyInfo(worksheet, building, startCol, colIndex);
    
    // 제안 정보 입력
    fillProposalInfo(worksheet, building, startCol);
    
    // 주차 정보 입력
    fillParkingInfo(worksheet, building, startCol);
    
    // 기본값 설정
    setDefaultValues(worksheet, startCol);
}

// 기초정보 입력
function fillBasicInfo(worksheet, building, col) {
    // 주소
    setCellValue(worksheet, `${col}18`, building.address || building.addressJibun || '');
    
    // 위치 (지하철역)
    setCellValue(worksheet, `${col}19`, building.station || '');
    
    // 준공일
    setCellValue(worksheet, `${col}20`, building.completionYear || '');
    
    // 규모 (층수)
    setCellValue(worksheet, `${col}21`, building.floors || '');
    
    // 연면적
    if (building.grossFloorAreaPy) {
        setCellValue(worksheet, `${col}22`, `${formatNumber(building.grossFloorAreaPy)}평`);
    }
    
    // 기준층 전용면적
    if (building.baseFloorAreaDedicatedPy) {
        setCellValue(worksheet, `${col}23`, `${formatNumber(building.baseFloorAreaDedicatedPy)}평`);
    }
    
    // 전용률
    if (building.dedicatedRate) {
        setCellValue(worksheet, `${col}24`, `${building.dedicatedRate}%`);
    }
    
    // 대지면적
    if (building.landAreaPy) {
        setCellValue(worksheet, `${col}25`, `${formatNumber(building.landAreaPy)}평`);
    }
}

// 공실 현황 입력
function fillVacancyInfo(worksheet, building, col, colIndex) {
    // 임시 데이터 (실제로는 층별 공실 정보 필요)
    // 예시: 1개 층만 표시
    const floorCol = col;
    const dedicatedCol = LG_UTILS.getColumnLetter(colIndex + 1);
    const rentCol = LG_UTILS.getColumnLetter(colIndex + 2);
    
    // 34행에 샘플 데이터
    if (building.vacancy) {
        // vacancy 정보에서 층 추출 시도
        const vacancyMatch = building.vacancy.match(/(\d+)층/);
        if (vacancyMatch) {
            setCellValue(worksheet, `${floorCol}34`, `${vacancyMatch[1]}층`);
        } else {
            setCellValue(worksheet, `${floorCol}34`, '10층');
        }
    } else {
        setCellValue(worksheet, `${floorCol}34`, '10층');
    }
    
    // 전용/임대 면적 (기준층 면적 사용)
    const dedicatedArea = building.baseFloorAreaDedicatedPy || 0;
    const rentArea = building.baseFloorAreaPy || 0;
    
    setCellValue(worksheet, `${dedicatedCol}34`, dedicatedArea);
    setCellValue(worksheet, `${rentCol}34`, rentArea);
    
    // 소계 수식
    setCellFormula(worksheet, `${dedicatedCol}39`, `=SUM(${dedicatedCol}34:${dedicatedCol}38)`);
    setCellFormula(worksheet, `${rentCol}39`, `=SUM(${rentCol}34:${rentCol}38)`);
}

// 제안 정보 입력
function fillProposalInfo(worksheet, building, col) {
    // 전용/임대면적 수식 (공실 테이블에서 가져옴)
    const colIndex = LG_UTILS.getColumnIndex(col);
    const dedicatedCol = LG_UTILS.getColumnLetter(colIndex + 1);
    const rentCol = LG_UTILS.getColumnLetter(colIndex + 2);
    
    setCellFormula(worksheet, `${col}43`, `=${dedicatedCol}34`);
    setCellFormula(worksheet, `${col}44`, `=${rentCol}34`);
    
    // 임대 조건 (평당 금액에서 숫자 추출)
    const depositPy = extractNumberFromPrice(building.depositPy);
    const rentPy = extractNumberFromPrice(building.rentPricePy);
    const managementPy = extractNumberFromPrice(building.managementFeePy);
    
    setCellValue(worksheet, `${col}45`, depositPy);
    setCellValue(worksheet, `${col}46`, rentPy);
    setCellValue(worksheet, `${col}47`, managementPy);
    
    // 실질 임대료 수식
    setCellFormula(worksheet, `${col}48`, `=${col}46*(12-${col}49)/12`);
    
    // 비용 계산 수식
    setCellFormula(worksheet, `${col}50`, `=${col}45*${col}44`);
    setCellFormula(worksheet, `${col}51`, `=${col}46*${col}44`);
    setCellFormula(worksheet, `${col}52`, `=${col}47*${col}44`);
    setCellFormula(worksheet, `${col}54`, `=${col}51+${col}52`);
    setCellFormula(worksheet, `${col}55`, `=${col}54*21`);
}

// 주차 정보 입력
function fillParkingInfo(worksheet, building, col) {
    // 총 주차대수
    setCellValue(worksheet, `${col}59`, building.parkingSpace || '');
    
    // 무료주차 제공대수 계산 수식
    setCellFormula(worksheet, `${col}61`, `=${col}44/${col}60`);
    
    // 유료주차비
    setCellValue(worksheet, `${col}62`, building.parkingFee || '');
}

// 기본값 설정
function setDefaultValues(worksheet, col) {
    // R.F 개월수
    setCellValue(worksheet, `${col}49`, 0);
    
    // 인테리어 기간
    setCellValue(worksheet, `${col}56`, '미제공');
    
    // 인테리어 지원금
    setCellValue(worksheet, `${col}57`, '미제공');
    
    // 무료주차 조건 (기본값)
    setCellValue(worksheet, `${col}60`, 50); // 50평당 1대
}

// 채권분석 수식 설정
function setBondAnalysisFormulas(worksheet, col) {
    // 담보율 계산
    setCellFormula(worksheet, `${col}30`, `=${col}29/${col}32`);
    
    // 토지가격 적용
    const colIndex = LG_UTILS.getColumnIndex(col);
    const landAreaCol = LG_UTILS.getColumnLetter(colIndex + 2); // 대지면적 열
    setCellFormula(worksheet, `${col}32`, `=${col}31*${landAreaCol}25`);
}

// 셀 값 설정 헬퍼
function setCellValue(worksheet, cellRef, value) {
    const cell = worksheet.getCell(cellRef);
    cell.value = value;
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
}

// 수식 설정 헬퍼
function setCellFormula(worksheet, cellRef, formula) {
    const cell = worksheet.getCell(cellRef);
    cell.value = { formula: formula };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
}

// 가격 문자열에서 숫자 추출
function extractNumberFromPrice(priceStr) {
    if (!priceStr) return 0;
    
    // "52만원", "5.20만원" 같은 형식에서 숫자 추출
    const match = priceStr.match(/([\d,]+\.?\d*)/);
    if (match) {
        return parseFloat(match[1].replace(/,/g, ''));
    }
    return 0;
}

// 숫자 포맷팅
function formatNumber(num) {
    if (!num) return '0';
    return Math.round(num).toLocaleString('ko-KR');
}

// 전역 함수로 등록
window.fillBuildingDataLG = fillBuildingDataLG;
