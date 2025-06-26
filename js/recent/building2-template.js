// ===== LG Comp List 템플릿 구조 생성 =====

// 워크시트 기본 설정
function setupWorksheet(worksheet) {
    // 열 너비 설정
    Object.entries(LG_TEMPLATE_CONFIG.columnWidths).forEach(([col, width]) => {
        const colIndex = LG_UTILS.getColumnIndex(col);
        worksheet.getColumn(colIndex).width = width;
    });
    
    // 행 높이 설정
    Object.entries(LG_TEMPLATE_CONFIG.rowHeights).forEach(([row, height]) => {
        worksheet.getRow(parseInt(row)).height = height;
    });
    
    // 셀 병합
    LG_TEMPLATE_CONFIG.mergedCells.forEach(range => {
        worksheet.mergeCells(range);
    });
}

// 템플릿 헤더 생성
function createTemplateHeader(worksheet, companyName, reportTitle) {
    // B3:C4 - PRESENT TO
    const b3 = worksheet.getCell('B3');
    const categoryConfig = LG_TEMPLATE_CONFIG.categories['B3'];
    b3.value = categoryConfig.text;
    b3.font = {
        name: 'LG Smart Regular',
        size: categoryConfig.fontSize,
        bold: categoryConfig.bold,
        color: { argb: categoryConfig.fontColor }
    };
    b3.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: categoryConfig.bgColor }
    };
    b3.alignment = { horizontal: 'center', vertical: 'middle' };
    applyBorder(b3);
    
    // B5:C5 - 로고 영역
    const b5 = worksheet.getCell('B5');
    b5.value = companyName || 'LG CNS';
    b5.font = { name: 'LG Smart Regular', size: 20, bold: true };
    b5.alignment = { horizontal: 'center', vertical: 'middle' };
    applyBorder(b5);
    
    // 보고서 제목 (필요시 추가 가능)
    // 현재 템플릿에는 별도 제목 영역이 없으므로 B6에 포함시킬 수 있음
}

// 카테고리 셀 생성
function createCategoryCell(worksheet, cellAddress) {
    const cell = worksheet.getCell(cellAddress);
    const config = LG_TEMPLATE_CONFIG.categories[cellAddress];
    
    if (config) {
        cell.value = config.text;
        cell.font = {
            name: 'LG Smart Regular',
            size: config.fontSize,
            bold: config.bold,
            color: { argb: config.fontColor }
        };
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: config.bgColor }
        };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        applyBorder(cell);
    }
}

// C열 라벨 생성
function createRowLabels(worksheet) {
    const labelStyle = LG_TEMPLATE_CONFIG.styles.labelCellStyle;
    
    Object.entries(LG_TEMPLATE_CONFIG.rowLabels).forEach(([row, label]) => {
        const cell = worksheet.getCell(`C${row}`);
        cell.value = label;
        cell.font = labelStyle.font;
        cell.alignment = labelStyle.alignment;
        applyBorder(cell);
    });
}

// 모든 카테고리 생성
function createAllCategories(worksheet) {
    Object.keys(LG_TEMPLATE_CONFIG.categories).forEach(cellAddress => {
        if (cellAddress !== 'B3') { // B3는 헤더에서 처리
            createCategoryCell(worksheet, cellAddress);
        }
    });
}

// 빌딩 열 헤더 생성 (D4, E4, F4, G4, H4)
function createBuildingHeaders(worksheet, buildings) {
    const headerStyle = LG_TEMPLATE_CONFIG.styles.headerStyle;
    
    buildings.forEach((building, index) => {
        if (index < 5) { // 최대 5개
            const col = String.fromCharCode(68 + index); // D, E, F, G, H
            const cell = worksheet.getCell(`${col}4`);
            
            cell.value = building.name;
            cell.font = headerStyle.font;
            cell.fill = headerStyle.fill;
            cell.alignment = headerStyle.alignment;
            applyBorder(cell);
        }
    });
}

// 용어 설명 추가
function addTerminology(worksheet) {
    // 용어 설명 헤더
    const headerCell = worksheet.getCell('B52');
    headerCell.value = LG_TEMPLATE_CONFIG.terminology[52];
    headerCell.font = { name: 'LG Smart Regular', size: 10, bold: true };
    headerCell.alignment = { horizontal: 'left', vertical: 'middle' };
    
    // 용어 설명 내용
    [53, 54, 55].forEach(row => {
        const cell = worksheet.getCell(`B${row}`);
        cell.value = LG_TEMPLATE_CONFIG.terminology[row];
        cell.font = { name: 'LG Smart Regular', size: 10 };
        cell.alignment = { horizontal: 'left', vertical: 'middle' };
    });
}

// 테두리 적용 헬퍼 함수
function applyBorder(cell) {
    cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };
}

// 전체 템플릿 생성 메인 함수
function createLGTemplate(workbook, worksheet, buildings, companyName, reportTitle) {
    // 1. 워크시트 기본 설정
    setupWorksheet(worksheet);
    
    // 2. 헤더 생성
    createTemplateHeader(worksheet, companyName, reportTitle);
    
    // 3. 모든 카테고리 생성
    createAllCategories(worksheet);
    
    // 4. C열 라벨 생성
    createRowLabels(worksheet);
    
    // 5. 빌딩 헤더 생성
    createBuildingHeaders(worksheet, buildings);
    
    // 6. 용어 설명 추가
    addTerminology(worksheet);
    
    // 7. A1-A4 제외 모든 셀에 기본 스타일 적용
    applyDefaultStyles(worksheet);
}

// 기본 스타일 적용 (A1-A4 제외 가운데 정렬)
function applyDefaultStyles(worksheet) {
    // 데이터가 입력될 영역 (D7:H50)에 기본 스타일 적용
    for (let row = 7; row <= 50; row++) {
        for (let colIndex = 4; colIndex <= 8; colIndex++) { // D부터 H까지
            const col = String.fromCharCode(64 + colIndex);
            const cellRef = `${col}${row}`;
            const cell = worksheet.getCell(cellRef);
            
            // 빈 셀에도 스타일 적용
            if (!cell.font) {
                cell.font = LG_TEMPLATE_CONFIG.styles.dataCellStyle.font;
            }
            if (!cell.alignment) {
                cell.alignment = LG_TEMPLATE_CONFIG.styles.dataCellStyle.alignment;
            }
            if (!cell.border) {
                applyBorder(cell);
            }
        }
    }
}

// 템플릿 검증 함수
function validateTemplate(worksheet) {
    // 필수 병합 셀 확인
    const requiredMerges = LG_TEMPLATE_CONFIG.mergedCells;
    let isValid = true;
    
    // 카테고리 셀 확인
    Object.keys(LG_TEMPLATE_CONFIG.categories).forEach(cellAddress => {
        const cell = worksheet.getCell(cellAddress);
        if (!cell.value) {
            console.warn(`카테고리 셀 ${cellAddress}가 비어있습니다.`);
            isValid = false;
        }
    });
    
    return isValid;
}
