// ===== LG Comp List 스타일 적용 =====

// 전체 스타일 적용 메인 함수
function applyLGStyles(worksheet) {
    // 1. 열 너비 설정
    setColumnWidths(worksheet);
    
    // 2. 전체 기본 폰트 먼저 설정 (요구사항2: LG스마트체 Regular)
    applyDefaultFont(worksheet);
    
    // 3. 섹션별 스타일 적용
    applySectionStyles(worksheet);
    
    // 4. 빌딩명 스타일
    applyBuildingNameStyles(worksheet);
    
    // 5. 공실 테이블 스타일
    applyVacancyTableStyles(worksheet);
    
    // 6. 수식 셀 스타일
    applyFormulaStyles(worksheet);
    
    // 7. 전체 테두리 적용
    applyAllBorders(worksheet);
    
    // 8. 색상 요구사항 적용 (요구사항 6-9)
    applyColorRequirements(worksheet);
}

// 열 너비 설정 (10개 빌딩 지원)
function setColumnWidths(worksheet) {
    // A열
    worksheet.getColumn('A').width = 9.375;
    
    // B열
    worksheet.getColumn('B').width = 4.5;
    
    // C-D열
    worksheet.getColumn('C').width = 9.375;
    worksheet.getColumn('D').width = 9.375;
    
    // E-AH열 (최대 10개 빌딩 데이터 영역)
    for (let i = 5; i <= 34; i++) { // E=5, AH=34
        worksheet.getColumn(i).width = 10.625;
    }
}

// 전체 기본 폰트 설정 (요구사항2: LG스마트체 Regular)
function applyDefaultFont(worksheet) {
    // 사용 범위의 모든 셀에 기본 폰트 적용
    for (let row = 1; row <= 85; row++) {
        for (let col = 1; col <= 34; col++) { // A-AH열
            const cell = worksheet.getCell(row, col);
            
            // 기본 폰트 설정
            if (!cell.font) {
                cell.font = {};
            }
            cell.font = {
                ...cell.font,
                name: 'LG Smart Regular',  // 요구사항2: LG스마트체 Regular
                size: cell.font.size || 10,
                color: cell.font.color || { argb: 'FF000000' }
            };
        }
    }
}

// 섹션별 스타일 적용
function applySectionStyles(worksheet) {
    // 헤더 스타일 (1-4행) - 요구사항3 반영
    const titleCell = worksheet.getCell('A1');
    titleCell.font = { 
        name: 'LG Smart Regular',
        size: 14, 
        bold: true,
        color: { argb: 'FF000000' }
    };
    titleCell.alignment = { horizontal: 'left', vertical: 'top' };
    
    // 섹션 타이틀 스타일
    const sectionCells = [
        'A6', 'A7', 'A9', 'A18', 'A26', 'A33', 
        'A40', 'A45', 'A48', 'A50', 'A56', 'A59', 'A63'
    ];
    
    sectionCells.forEach(cellRef => {
        const cell = worksheet.getCell(cellRef);
        
        // 배경색 설정 (요구사항6: A6는 주황 80% 밝게)
        if (cellRef === 'A6') {
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: LG_TEMPLATE_CONFIG.colors.a6 }  // 주황 80% 밝게
            };
        } else if (cellRef === 'A7') {
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: LG_TEMPLATE_CONFIG.colors.location }  // 검정 35% 밝게
            };
        } else {
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFE0E0E0' }
            };
        }
        
        // 폰트 설정
        cell.font = { 
            name: 'LG Smart Regular',
            bold: true, 
            size: 10,
            color: { argb: 'FF000000' }
        };
        
        // 정렬 설정
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
        // 폰트 재설정
        cell.font = {
            name: 'LG Smart Regular',
            size: 10,
            color: { argb: 'FF000000' }
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
        // 폰트 재설정
        cell.font = {
            name: 'LG Smart Regular',
            size: 10,
            color: { argb: 'FF000000' }
        };
    }
}

// 빌딩명 스타일 적용 (요구사항7: 녹색 80% 밝게)
function applyBuildingNameStyles(worksheet) {
    LG_TEMPLATE_CONFIG.buildingColumns.forEach(col => {
        const cell = worksheet.getCell(`${col}6`);
        if (cell.value) {
            // 배경색 설정 (요구사항7: 녹색 80% 밝게)
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: LG_TEMPLATE_CONFIG.colors.buildingName }  // 녹색 80% 밝게
            };
            
            // 폰트 설정
            cell.font = {
                name: 'LG Smart Regular',
                size: 12,
                bold: true,
                color: { argb: 'FF000000' }  // 검은색 텍스트
            };
            
            // 정렬
            cell.alignment = {
                horizontal: 'center',
                vertical: 'middle'
            };
            
            // 테두리
            cell.border = {
                top: { style: 'thin', color: { argb: 'FF000000' } },
                left: { style: 'thin', color: { argb: 'FF000000' } },
                bottom: { style: 'thin', color: { argb: 'FF000000' } },
                right: { style: 'thin', color: { argb: 'FF000000' } }
            };
        }
    });
}

// 색상 요구사항 적용 (요구사항 6-9)
function applyColorRequirements(worksheet) {
    LG_TEMPLATE_CONFIG.buildingColumns.forEach(col => {
        // 요구사항8: E8, H8, K8, N8, Q8, T8 등 - 파랑 80% 밝게
        const cell8 = worksheet.getCell(`${col}8`);
        cell8.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: LG_TEMPLATE_CONFIG.colors.proposal }  // 파랑 80% 밝게
        };
        cell8.font = {
            name: 'LG Smart Regular',
            size: 10,
            color: { argb: 'FF000000' }
        };
        
        // 요구사항9: E7, H7, K7, N7, Q7, T7 등 - 검정 35% 밝게
        const cell7 = worksheet.getCell(`${col}7`);
        cell7.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: LG_TEMPLATE_CONFIG.colors.location }  // 검정 35% 밝게
        };
        cell7.font = {
            name: 'LG Smart Regular',
            size: 10,
            color: { argb: 'FF000000' }
        };
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
            
            // 배경색
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFD9D9D9' }  // 연한 회색
            };
            
            // 폰트
            cell.font = { 
                name: 'LG Smart Regular',
                bold: true, 
                size: 10,
                color: { argb: 'FF000000' }
            };
            
            // 정렬
            cell.alignment = {
                horizontal: 'center',
                vertical: 'middle'
            };
            
            // 테두리
            cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
        }
        
        // 데이터 영역 스타일 (34-38행)
        for (let row = 34; row <= 38; row++) {
            for (let offset = 0; offset < 3; offset++) {
                const col = LG_UTILS.getColumnLetter(colIndex + offset);
                const cell = worksheet.getCell(`${col}${row}`);
                
                cell.font = { name: 'LG Smart Regular', size: 10 };
                cell.alignment = {
                    horizontal: 'center',
                    vertical: 'middle'
                };
                
                // 숫자 포맷 (전용/임대 열)
                if (offset > 0 && cell.value) {
                    cell.numFmt = '#,##0';
                }
                
                // 테두리
                cell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
            }
        }
        
        // 소계 행 스타일 (39행)
        for (let offset = 0; offset < 3; offset++) {
            const col = LG_UTILS.getColumnLetter(colIndex + offset);
            const cell = worksheet.getCell(`${col}39`);
            
            // 배경색
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFF2F2F2' }
            };
            
            // 폰트
            cell.font = { 
                name: 'LG Smart Regular',
                bold: true,
                size: 10,
                color: { argb: 'FF000000' }
            };
            
            // 테두리
            cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
        }
    });
}

// 수식 셀 스타일 적용
function applyFormulaStyles(worksheet) {
    // 수식이 있는 행들
    const formulaRows = [30, 32, 39, 43, 44, 48, 50, 51, 52, 54, 55, 61];
    
    LG_TEMPLATE_CONFIG.buildingColumns.forEach(col => {
        formulaRows.forEach(row => {
            const cell = worksheet.getCell(`${col}${row}`);
            
            // LG Smart Regular 폰트 적용
            cell.font = {
                name: 'LG Smart Regular',
                size: 10
            };
            
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
                cell.numFmt = '#,##0';
            } else if ([43, 44].includes(row)) {
                // 면적
                cell.numFmt = '#,##0';
            }
            
            // 테두리
            cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
        });
    });
}

// 전체 테두리 적용
function applyAllBorders(worksheet) {
    // 주요 영역에 테두리 적용 (1-85행, A-AH열)
    for (let row = 1; row <= 85; row++) {
        for (let col = 1; col <= 34; col++) { // AH=34
            const cell = worksheet.getCell(row, col);
            
            // 병합된 셀이거나 값이 있는 셀에만 테두리 적용
            if (cell.isMerged || cell.value !== null) {
                if (!cell.border) {
                    cell.border = {
                        top: { style: 'thin', color: { argb: 'FF000000' } },
                        left: { style: 'thin', color: { argb: 'FF000000' } },
                        bottom: { style: 'thin', color: { argb: 'FF000000' } },
                        right: { style: 'thin', color: { argb: 'FF000000' } }
                    };
                }
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
    
    // 인쇄 영역 설정 (최대 AH열까지)
    worksheet.pageSetup.printArea = 'A1:AH85';
}

// 조건부 서식 적용 (선택사항)
function applyConditionalFormatting(worksheet) {
    // 예: 특정 값 범위에 따른 색상 변경
    LG_TEMPLATE_CONFIG.buildingColumns.forEach(col => {
        // 실질 임대료가 높은 경우 강조
        const cell48 = worksheet.getCell(`${col}48`);
        
        // 월 납부액이 높은 경우 강조
        const cell54 = worksheet.getCell(`${col}54`);
        
        // 이 부분은 필요에 따라 구현
    });
}

// 스타일 검증 함수
function validateStyles(worksheet) {
    const errors = [];
    let isValid = true;
    
    // LG Smart Regular 폰트 검증
    for (let row = 1; row <= 85; row++) {
        for (let col = 1; col <= 34; col++) {
            const cell = worksheet.getCell(row, col);
            if (cell.value && cell.font && cell.font.name !== 'LG Smart Regular') {
                errors.push(`${cell.address}: 폰트가 LG Smart Regular가 아닙니다.`);
                isValid = false;
            }
        }
    }
    
    // 색상 요구사항 검증
    const a6 = worksheet.getCell('A6');
    if (a6.fill && a6.fill.fgColor && a6.fill.fgColor.argb !== LG_TEMPLATE_CONFIG.colors.a6) {
        errors.push('A6: 주황 80% 밝게 색상이 적용되지 않았습니다.');
        isValid = false;
    }
    
    return {
        isValid: isValid,
        errors: errors
    };
}

// 대량 스타일 적용 최적화
function optimizedApplyLGStyles(worksheet, buildingCount) {
    console.log('LG 스타일 적용 시작...');
    
    // 1. 기본 설정
    setColumnWidths(worksheet);
    
    // 2. 배치로 폰트 적용
    applyDefaultFont(worksheet);
    
    // 3. 선택된 빌딩 수만큼만 스타일 적용
    for (let i = 0; i < buildingCount && i < 10; i++) {
        const col = LG_TEMPLATE_CONFIG.buildingColumns[i];
        applyBuildingSpecificStyles(worksheet, col, i);
    }
    
    // 4. 나머지 스타일
    applySectionStyles(worksheet);
    applyVacancyTableStyles(worksheet);
    applyFormulaStyles(worksheet);
    applyAllBorders(worksheet);
    applyColorRequirements(worksheet);
    
    console.log('LG 스타일 적용 완료');
}

// 빌딩별 스타일 적용
function applyBuildingSpecificStyles(worksheet, col, index) {
    // 빌딩명 스타일
    const nameCell = worksheet.getCell(`${col}6`);
    if (nameCell.value) {
        nameCell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: LG_TEMPLATE_CONFIG.colors.buildingName }
        };
        nameCell.font = {
            name: 'LG Smart Regular',
            size: 12,
            bold: true,
            color: { argb: 'FF000000' }
        };
    }
    
    // 7행, 8행 색상 적용
    const cell7 = worksheet.getCell(`${col}7`);
    cell7.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: LG_TEMPLATE_CONFIG.colors.location }
    };
    
    const cell8 = worksheet.getCell(`${col}8`);
    cell8.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: LG_TEMPLATE_CONFIG.colors.proposal }
    };
}

// 성능 모니터링
function measureStylePerformance(worksheet, buildingCount) {
    const startTime = performance.now();
    
    optimizedApplyLGStyles(worksheet, buildingCount);
    
    const endTime = performance.now();
    console.log(`스타일 적용 시간: ${endTime - startTime}ms`);
    
    return endTime - startTime;
}

// 전역 함수로 등록
window.applyLGStyles = applyLGStyles;
window.applyPrintSettings = applyPrintSettings;
window.validateStyles = validateStyles;
window.optimizedApplyLGStyles = optimizedApplyLGStyles;
window.measureStylePerformance = measureStylePerformance;
    