// ===== LG Comp List 스타일 적용 =====

// 전체 스타일 적용 메인 함수
function applyLGStyles(worksheet) {
    // 1. 열 너비 설정
    setColumnWidths(worksheet);
    
    // 2. 섹션별 스타일 적용
    applySectionStyles(worksheet);
    
    // 3. 빌딩명 스타일
    applyBuildingNameStyles(worksheet);
    
    // 4. 공실 테이블 스타일
    applyVacancyTableStyles(worksheet);
    
    // 5. 수식 셀 스타일
    applyFormulaStyles(worksheet);
    
    // 6. 전체 폰트 설정
    applyDefaultFont(worksheet);
}

// 열 너비 설정
function setColumnWidths(worksheet) {
    // A열
    worksheet.getColumn('A').width = 9.375;
    
    // B열
    worksheet.getColumn('B').width = 4.5;
    
    // C-D열
    worksheet.getColumn('C').width = 9.375;
    worksheet.getColumn('D').width = 9.375;
    
    // E-V열 (빌딩 데이터 영역)
    for (let i = 5; i <= 22; i++) { // E=5, V=22
        worksheet.getColumn(i).width = 10.625;
    }
}

// 섹션별 스타일 적용
function applySectionStyles(worksheet) {
    // 헤더 스타일 (1-4행)
    const titleCell = worksheet.getCell('A1');
    titleCell.font = { size: 14, bold: true };
    titleCell.alignment = { horizontal: 'left', vertical: 'top' };
    
    // 섹션 타이틀 스타일
    const sectionCells = [
        'A6', 'A7', 'A9', 'A18', 'A26', 'A33', 
        'A40', 'A45', 'A48', 'A50', 'A56', 'A59', 'A63'
    ];
    
    sectionCells.forEach(cellRef => {
        const cell = worksheet.getCell(cellRef);
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFE0E0E0' }
        };
        cell.font = { bold: true, size: 10 };
        cell.alignment = { 
            horizontal: 'center', 
            vertical: 'middle',
            wrapText: true
        };
    });
    
    // B열 라벨 배경색
    // 기준층 임대기준 (45-47행) - 연한 노란색
    for (let row = 45; row <= 47; row++) {
        const cell = worksheet.getCell(`B${row}`);
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFFF2CC' }
        };
    }
    
    // 실질 임대기준 (48-49행) - 연한 녹색
    for (let row = 48; row <= 49; row++) {
        const cell = worksheet.getCell(`B${row}`);
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFE2EFDA' }
        };
    }
}

// 빌딩명 스타일 적용
function applyBuildingNameStyles(worksheet) {
    LG_TEMPLATE_CONFIG.buildingColumns.forEach(col => {
        const cell = worksheet.getCell(`${col}6`);
        if (cell.value) {
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FF4472C4' }
            };
            cell.font = {
                size: 12,
                bold: true,
                color: { argb: 'FFFFFFFF' }
            };
            cell.alignment = {
                horizontal: 'center',
                vertical: 'middle'
            };
        }
    });
}

// 공실 테이블 스타일 적용
function applyVacancyTableStyles(worksheet) {
    LG_TEMPLATE_CONFIG.buildingColumns.forEach(startCol => {
        const colIndex = LG_UTILS.getColumnIndex(startCol);
        
        // 헤더 스타일 (33행)
        for (let offset = 0; offset < 3; offset++) {
            const col = LG_UTILS.getColumnLetter(colIndex + offset);
            const cell = worksheet.getCell(`${col}33`);
            
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFD9D9D9' }
            };
            cell.font = { bold: true, size: 10 };
            cell.alignment = {
                horizontal: 'center',
                vertical: 'middle'
            };
        }
        
        // 데이터 영역 스타일 (34-38행)
        for (let row = 34; row <= 38; row++) {
            for (let offset = 0; offset < 3; offset++) {
                const col = LG_UTILS.getColumnLetter(colIndex + offset);
                const cell = worksheet.getCell(`${col}${row}`);
                
                cell.alignment = {
                    horizontal: 'center',
                    vertical: 'middle'
                };
                
                // 숫자 포맷 (전용/임대 열)
                if (offset > 0) {
                    cell.numFmt = '#,##0';
                }
            }
        }
        
        // 소계 행 스타일 (39행)
        const totalCell = worksheet.getCell(`${startCol}39`);
        totalCell.font = { bold: true };
        totalCell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFF2F2F2' }
        };
    });
}

// 수식 셀 스타일 적용
function applyFormulaStyles(worksheet) {
    // 수식이 있는 행들
    const formulaRows = [30, 32, 39, 43, 44, 48, 50, 51, 52, 54, 55, 61];
    
    LG_TEMPLATE_CONFIG.buildingColumns.forEach(col => {
        formulaRows.forEach(row => {
            const cell = worksheet.getCell(`${col}${row}`);
            
            if (cell.formula) {
                // 수식 결과 정렬
                cell.alignment = {
                    horizontal: 'center',
                    vertical: 'middle'
                };
                
                // 특정 행 숫자 포맷
                if ([30, 61].includes(row)) {
                    // 비율
                    cell.numFmt = '0.00%';
                } else if ([50, 51, 52, 54, 55].includes(row)) {
                    // 금액
                    cell.numFmt = '#,##0"원"';
                } else if ([43, 44].includes(row)) {
                    // 면적
                    cell.numFmt = '#,##0"평"';
                }
            }
        });
    });
}

// 전체 기본 폰트 설정
function applyDefaultFont(worksheet) {
    // 사용 범위의 모든 셀에 기본 폰트 적용
    for (let row = 1; row <= 85; row++) {
        for (let col = 1; col <= 22; col++) { // A-V열
            const cell = worksheet.getCell(row, col);
            
            if (!cell.font || !cell.font.name) {
                cell.font = {
                    ...cell.font,
                    name: '맑은 고딕',
                    size: cell.font?.size || 10
                };
            }
        }
    }
}

// 인쇄 설정
function applyPrintSettings(worksheet) {
    worksheet.pageSetup = {
        paperSize: 9, // A4
        orientation: 'landscape', // 가로
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
    worksheet.pageSetup.printArea = 'A1:V85';
}

// 전역 함수로 등록
window.applyLGStyles = applyLGStyles;
window.applyPrintSettings = applyPrintSettings;
