// ===== LG Comp List 스타일 적용 =====

// 전체 스타일 적용 메인 함수
function applyLGStyles(worksheet) {
    // 1. 전체 시트에 기본 폰트 적용
    applyDefaultFont(worksheet);
    
    // 2. 카테고리 스타일 적용
    applyCategoryStyles(worksheet);
    
    // 3. 데이터 영역 스타일 적용
    applyDataAreaStyles(worksheet);
    
    // 4. 특수 영역 스타일 적용
    applySpecialAreaStyles(worksheet);
    
    // 5. 조건부 스타일 적용
    applyConditionalStyles(worksheet);
}

// 기본 폰트 적용 (LG Smart Regular)
function applyDefaultFont(worksheet) {
    // 사용 범위의 모든 셀에 기본 폰트 적용
    const maxRow = 55;
    const maxCol = 8; // H열까지
    
    for (let row = 1; row <= maxRow; row++) {
        for (let col = 1; col <= maxCol; col++) {
            const cellRef = LG_UTILS.getColumnLetter(col) + row;
            const cell = worksheet.getCell(cellRef);
            
            // 기존 폰트 설정이 없으면 기본값 적용
            if (!cell.font || !cell.font.name) {
                cell.font = {
                    ...cell.font,
                    name: 'LG Smart Regular',
                    size: cell.font?.size || 9
                };
            }
        }
    }
}

// 카테고리 스타일 적용
function applyCategoryStyles(worksheet) {
    Object.entries(LG_TEMPLATE_CONFIG.categories).forEach(([cellRef, config]) => {
        const cell = worksheet.getCell(cellRef);
        
        // 배경색 (그레이 계열)
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: config.bgColor }
        };
        
        // 폰트
        cell.font = {
            name: 'LG Smart Regular',
            size: config.fontSize,
            bold: config.bold,
            color: { argb: config.fontColor }
        };
        
        // 정렬 (가운데)
        cell.alignment = {
            horizontal: 'center',
            vertical: 'middle'
        };
        
        // 테두리
        applyBorderStyle(cell);
    });
}

// 데이터 영역 스타일 적용
function applyDataAreaStyles(worksheet) {
    // C열 라벨 스타일
    Object.keys(LG_TEMPLATE_CONFIG.rowLabels).forEach(row => {
        const cell = worksheet.getCell(`C${row}`);
        
        // 모든 라벨 가운데 정렬
        cell.alignment = {
            horizontal: 'center',
            vertical: 'middle'
        };
        
        cell.font = {
            name: 'LG Smart Regular',
            size: 9
        };
        
        applyBorderStyle(cell);
    });
    
    // D-H열 데이터 영역 (7-50행)
    for (let row = 7; row <= 50; row++) {
        for (let col = 4; col <= 8; col++) { // D부터 H까지
            const cellRef = LG_UTILS.getColumnLetter(col) + row;
            const cell = worksheet.getCell(cellRef);
            
            // A1-A4가 아닌 모든 셀 가운데 정렬
            if (!['A1', 'A2', 'A3', 'A4'].includes(cellRef)) {
                if (!cell.alignment) {
                    cell.alignment = {};
                }
                
                // 특정 행은 다른 정렬 적용
                if ([32, 33, 34, 36, 37, 38, 39, 40, 42, 43, 46, 47, 48].includes(row)) {
                    // 금액 관련 행은 우측 정렬
                    cell.alignment.horizontal = 'right';
                } else if ([44, 49, 50].includes(row)) {
                    // NOC, 월 합계, 연 합계는 가운데 정렬
                    cell.alignment.horizontal = 'center';
                } else {
                    // 나머지는 가운데 정렬
                    cell.alignment.horizontal = 'center';
                }
                
                cell.alignment.vertical = 'middle';
            }
            
            // 폰트
            if (!cell.font) {
                cell.font = {};
            }
            cell.font.name = 'LG Smart Regular';
            cell.font.size = cell.font.size || 9;
            
            // 테두리
            applyBorderStyle(cell);
        }
    }
}

// 특수 영역 스타일 적용
function applySpecialAreaStyles(worksheet) {
    // B5 로고 영역
    const logoCell = worksheet.getCell('B5');
    logoCell.font = {
        name: 'LG Smart Regular',
        size: 20,
        bold: true
    };
    logoCell.alignment = {
        horizontal: 'center',
        vertical: 'middle'
    };
    applyBorderStyle(logoCell);
    
    // 빌딩명 헤더 (D4-H4)
    for (let col = 4; col <= 8; col++) {
        const cell = worksheet.getCell(LG_UTILS.getColumnLetter(col) + '4');
        if (cell.value) {
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFCCCCCC' }
            };
            cell.font = {
                name: 'LG Smart Regular',
                size: 9,
                bold: true
            };
            cell.alignment = {
                horizontal: 'center',
                vertical: 'middle',
                wrapText: true
            };
            applyBorderStyle(cell);
        }
    }
    
    // 용어 설명 영역 (52-55행)
    const termHeaderCell = worksheet.getCell('B52');
    termHeaderCell.font = {
        name: 'LG Smart Regular',
        size: 10,
        bold: true
    };
    
    [53, 54, 55].forEach(row => {
        const cell = worksheet.getCell(`B${row}`);
        cell.font = {
            name: 'LG Smart Regular',
            size: 10
        };
        cell.alignment = {
            horizontal: 'left',
            vertical: 'middle'
        };
    });
}

// 조건부 스타일 적용
function applyConditionalStyles(worksheet) {
    // 음수 값 빨간색 표시
    const moneyRows = [32, 33, 34, 35, 36, 37, 38, 39, 40, 42, 43, 44, 46, 47, 48, 49, 50];
    
    for (let col = 4; col <= 8; col++) {
        moneyRows.forEach(row => {
            const cell = worksheet.getCell(LG_UTILS.getColumnLetter(col) + row);
            
            // 조건부 서식은 ExcelJS에서 직접 지원하지 않으므로
            // 값을 확인하여 스타일 적용
            if (cell.value && typeof cell.value === 'number' && cell.value < 0) {
                cell.font = {
                    ...cell.font,
                    color: { argb: 'FFFF0000' } // 빨간색
                };
            }
        });
    }
}

// 테두리 스타일 적용 헬퍼
function applyBorderStyle(cell, style = 'thin') {
    cell.border = {
        top: { style: style },
        left: { style: style },
        bottom: { style: style },
        right: { style: style }
    };
}

// 특정 범위에 스타일 일괄 적용
function applyStyleToRange(worksheet, startCell, endCell, style) {
    const startCol = startCell.charCodeAt(0) - 64;
    const startRow = parseInt(startCell.substring(1));
    const endCol = endCell.charCodeAt(0) - 64;
    const endRow = parseInt(endCell.substring(1));
    
    for (let row = startRow; row <= endRow; row++) {
        for (let col = startCol; col <= endCol; col++) {
            const cellRef = LG_UTILS.getColumnLetter(col) + row;
            const cell = worksheet.getCell(cellRef);
            
            // 스타일 적용
            if (style.font) cell.font = { ...cell.font, ...style.font };
            if (style.alignment) cell.alignment = { ...cell.alignment, ...style.alignment };
            if (style.fill) cell.fill = style.fill;
            if (style.border) cell.border = style.border;
            if (style.numFmt) cell.numFmt = style.numFmt;
        }
    }
}

// 인쇄 설정
function applyPrintSettings(worksheet) {
    // 페이지 설정
    worksheet.pageSetup = {
        paperSize: 9, // A4
        orientation: 'portrait',
        fitToPage: true,
        fitToWidth: 1,
        fitToHeight: 0,
        margins: {
            left: 0.7,
            right: 0.7,
            top: 0.75,
            bottom: 0.75,
            header: 0.3,
            footer: 0.3
        }
    };
    
    // 인쇄 영역 설정
    worksheet.pageSetup.printArea = 'A1:H55';
    
    // 반복 행 설정 (헤더)
    worksheet.pageSetup.printTitlesRow = '1:4';
}

// 스타일 검증
function validateStyles(worksheet) {
    const errors = [];
    
    // LG Smart Regular 폰트 확인
    for (let row = 1; row <= 55; row++) {
        for (let col = 1; col <= 8; col++) {
            const cellRef = LG_UTILS.getColumnLetter(col) + row;
            const cell = worksheet.getCell(cellRef);
            
            if (cell.value && (!cell.font || cell.font.name !== 'LG Smart Regular')) {
                errors.push(`${cellRef}: LG Smart Regular 폰트가 적용되지 않음`);
            }
        }
    }
    
    // A1-A4 제외 가운데 정렬 확인
    for (let row = 5; row <= 50; row++) {
        for (let col = 1; col <= 8; col++) {
            const cellRef = LG_UTILS.getColumnLetter(col) + row;
            const cell = worksheet.getCell(cellRef);
            
            if (cell.value && !cell.alignment) {
                errors.push(`${cellRef}: 정렬이 설정되지 않음`);
            }
        }
    }
    
    return {
        isValid: errors.length === 0,
        errors: errors
    };
}
