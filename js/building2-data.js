// ===== LG Comp List 데이터 처리 =====

// 빌딩 데이터를 LG 형식으로 변환
function processLGBuildingData(building) {
    const processedData = {};
    
    // 기본 데이터 매핑
    Object.entries(LG_TEMPLATE_CONFIG.dataMapping).forEach(([row, field]) => {
        if (building[field] !== undefined) {
            if (field === 'dedicatedRate') {
                // 전용률은 이미 퍼센트 값이므로 100으로 나누지 않음
                processedData[row] = building[field] || 0;
            } else {
                processedData[row] = building[field];
            }
        } else {
            processedData[row] = '';
        }
    });
    
    // 특수 처리가 필요한 필드들
    processedData.description = building.description || '';
    
    // buildings.json에 있는 추가 데이터 처리
    processedData.landAreaPy = building.landAreaPy || '';
    processedData.depositPy = building.depositPy || '';
    processedData.rentPricePy = building.rentPricePy || '';
    processedData.managementFeePy = building.managementFeePy || '';
    
    return processedData;
}

// 빌딩 데이터를 워크시트에 입력 (LG 버전)
function fillBuildingDataLG(worksheet, building, columnIndex) {
    const col = String.fromCharCode(64 + columnIndex); // D=4, E=5, F=6, G=7, H=8
    const processedData = processLGBuildingData(building);
    
    // 빌딩개요/일반 (6행)
    if (processedData.description) {
        setCellValue(worksheet, `${col}6`, processedData.description, {
            wrapText: true,
            alignment: 'left'
        });
    }
    
    // 기본 정보 입력 (7-23행) - buildings.json의 실제 데이터와 매핑
    setCellValue(worksheet, `${col}7`, building.addressJibun || '');
    setCellValue(worksheet, `${col}8`, building.address || '');
    setCellValue(worksheet, `${col}9`, building.station || '');
    setCellValue(worksheet, `${col}10`, building.floors || '');
    setCellValue(worksheet, `${col}11`, building.completionYear || '');
    
    // 전용률
    setCellValue(worksheet, `${col}12`, (building.dedicatedRate || 0) / 100, {
        format: LG_TEMPLATE_CONFIG.numberFormats.percentage
    });
    
    // 면적 정보 - buildings.json의 실제 필드명 사용
    setCellValue(worksheet, `${col}13`, building.baseFloorArea || 0, {
        format: LG_TEMPLATE_CONFIG.numberFormats.area_m2
    });
    setCellValue(worksheet, `${col}14`, building.baseFloorAreaPy || 0, {
        format: LG_TEMPLATE_CONFIG.numberFormats.area_py
    });
    setCellValue(worksheet, `${col}15`, building.baseFloorAreaDedicated || 0, {
        format: LG_TEMPLATE_CONFIG.numberFormats.area_m2
    });
    setCellValue(worksheet, `${col}16`, building.baseFloorAreaDedicatedPy || 0, {
        format: LG_TEMPLATE_CONFIG.numberFormats.area_py
    });
    
    // 빌딩 세부현황
    setCellValue(worksheet, `${col}17`, building.elevator || '');
    setCellValue(worksheet, `${col}18`, building.hvac || '');
    setCellValue(worksheet, `${col}19`, building.buildingUse || '');
    setCellValue(worksheet, `${col}20`, building.structure || '');
    setCellValue(worksheet, `${col}21`, building.parkingSpace || '');
    
    // 주차 관련
    setCellValue(worksheet, `${col}22`, building.parkingFee || '');
    setCellValue(worksheet, `${col}23`, building.parkingSpace || '');
    
    // 임차 제안 기본값 설정 (25-31행)
    setDefaultProposalValues(worksheet, col, building);
    
    // 임대 기준 - buildings.json의 실제 데이터 활용
    setRentValuesFromBuilding(worksheet, col, building);
    
    // 렌트프리 기본값 (41행)
    setCellValue(worksheet, `${col}41`, 0, {
        format: '0',
        alignment: 'center'
    });
}

// 임차 제안 기본값 설정
function setDefaultProposalValues(worksheet, col, building) {
    // 25: 최적 임차 층수
    setCellValue(worksheet, `${col}25`, '-');
    
    // 26: 입주 가능 시기
    setCellValue(worksheet, `${col}26`, '-');
    
    // 27: 거래유형
    setCellValue(worksheet, `${col}27`, '-');
    
    // 30: 임대면적 (평) - 기준층 임대면적 사용
    const rentAreaPy = building.baseFloorAreaPy || 100;
    setCellValue(worksheet, `${col}30`, rentAreaPy, {
        format: LG_TEMPLATE_CONFIG.numberFormats.area_py
    });
    
    // 31: 전용면적 (평) - 기준층 전용면적 사용
    const dedicatedAreaPy = building.baseFloorAreaDedicatedPy || rentAreaPy * 0.65;
    setCellValue(worksheet, `${col}31`, dedicatedAreaPy, {
        format: LG_TEMPLATE_CONFIG.numberFormats.area_py
    });
}

// buildings.json의 임대 정보를 사용하여 임대 기준 설정
function setRentValuesFromBuilding(worksheet, col, building) {
    // depositPy, rentPricePy, managementFeePy에서 숫자 추출
    const extractNumber = (str) => {
        if (!str) return 0;
        // "52만원", "5.20만원" 같은 형식에서 숫자 추출
        const match = str.match(/[\d,]+\.?\d*/);
        if (match) {
            return parseFloat(match[0].replace(/,/g, ''));
        }
        return 0;
    };
    
    // 32: 월 평당 보증금
    const depositPy = extractNumber(building.depositPy) * 10000; // 만원 단위를 원 단위로
    setCellValue(worksheet, `${col}32`, depositPy, {
        format: LG_TEMPLATE_CONFIG.numberFormats.currency,
        alignment: 'right'
    });
    
    // 33: 월 평당 임대료
    const rentPy = extractNumber(building.rentPricePy) * 10000; // 만원 단위를 원 단위로
    setCellValue(worksheet, `${col}33`, rentPy, {
        format: LG_TEMPLATE_CONFIG.numberFormats.currency,
        alignment: 'right'
    });
    
    // 34: 월 평당 관리비
    const managementPy = extractNumber(building.managementFeePy) * 10000; // 만원 단위를 원 단위로
    setCellValue(worksheet, `${col}34`, managementPy, {
        format: LG_TEMPLATE_CONFIG.numberFormats.currency,
        alignment: 'right'
    });
}

// 셀 값 설정 헬퍼 함수
function setCellValue(worksheet, cellRef, value, options = {}) {
    const cell = worksheet.getCell(cellRef);
    cell.value = value;
    
    // 폰트 설정 (LG Smart Regular 강제)
    cell.font = {
        name: 'LG Smart Regular',
        size: 9,
        bold: options.bold || false
    };
    
    // 정렬 설정 (A1-A4가 아니면 기본 가운데 정렬)
    if (!['A1', 'A2', 'A3', 'A4'].includes(cellRef)) {
        cell.alignment = {
            horizontal: options.alignment || 'center',
            vertical: 'middle',
            wrapText: options.wrapText || false
        };
    }
    
    // 숫자 포맷
    if (options.format) {
        cell.numFmt = options.format;
    }
    
    // 테두리
    cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };
}

// 데이터 검증 함수
function validateBuildingData(building) {
    const errors = [];
    
    // 필수 필드 확인
    if (!building.name) {
        errors.push('빌딩명이 없습니다.');
    }
    
    if (!building.address && !building.addressJibun) {
        errors.push('주소 정보가 없습니다.');
    }
    
    // 숫자 필드 검증
    const numericFields = [
        'dedicatedRate', 'baseFloorArea', 'baseFloorAreaPy',
        'baseFloorAreaDedicated', 'baseFloorAreaDedicatedPy'
    ];
    
    numericFields.forEach(field => {
        if (building[field] !== undefined && isNaN(building[field])) {
            errors.push(`${field} 필드가 숫자가 아닙니다.`);
        }
    });
    
    return {
        isValid: errors.length === 0,
        errors: errors
    };
}

// 데이터 정제 함수
function sanitizeBuildingData(building) {
    const sanitized = { ...building };
    
    // 숫자 필드 정제
    const numericFields = [
        'dedicatedRate', 'baseFloorArea', 'baseFloorAreaPy',
        'baseFloorAreaDedicated', 'baseFloorAreaDedicatedPy',
        'landAreaPy', 'completionYear'
    ];
    
    numericFields.forEach(field => {
        if (sanitized[field]) {
            // 문자열에서 숫자 추출
            const numValue = parseFloat(String(sanitized[field]).replace(/[^0-9.-]/g, ''));
            sanitized[field] = isNaN(numValue) ? 0 : numValue;
        }
    });
    
    // 텍스트 필드 정제
    const textFields = ['name', 'address', 'addressJibun', 'station', 'floors'];
    textFields.forEach(field => {
        if (sanitized[field]) {
            sanitized[field] = String(sanitized[field]).trim();
        }
    });
    
    return sanitized;
}

// 여러 빌딩 데이터 일괄 처리
function processBuildingsBatch(buildings) {
    return buildings.map(building => {
        const sanitized = sanitizeBuildingData(building);
        const validation = validateBuildingData(sanitized);
        
        if (!validation.isValid) {
            console.warn(`빌딩 ${building.name} 데이터 검증 실패:`, validation.errors);
        }
        
        return {
            original: building,
            processed: processLGBuildingData(sanitized),
            validation: validation
        };
    });
}

// 전역 함수로 등록
window.fillBuildingDataLG = fillBuildingDataLG;
window.processLGBuildingData = processLGBuildingData;
window.validateBuildingData = validateBuildingData;
