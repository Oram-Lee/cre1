// building.js에 추가할 완전한 스타일 정의와 적용 함수

// 템플릿 스타일 정의
const templateStyles = {
    // PRESENT TO 스타일 (B3)
    presentTo: {
        font: { name: 'Noto Sans KR', size: 9, bold: true, color: 'FFFFFFFF' },
        fill: { fgColor: { rgb: 'FF2C2A2A' }, patternType: 'solid' },
        alignment: { horizontal: 'center', vertical: 'center' },
        border: { all: { style: 'thin', color: { rgb: 'FFB8B8B8' } } }
    },
    
    // 빌딩명 헤더 스타일 (D4~H4)
    buildingHeader: {
        font: { name: 'Noto Sans KR', size: 9, bold: true, color: 'FF000000' },
        fill: { fgColor: { rgb: 'FFCCCCCC' }, patternType: 'solid' },
        alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
        border: { all: { style: 'thin', color: { rgb: 'FFB8B8B8' } } }
    },
    
    // 카테고리 스타일 (B열 - 기본)
    category: {
        font: { name: 'Noto Sans KR', size: 9, bold: true, color: 'FF000000' },
        fill: { fgColor: { rgb: 'FFFFFFFF' }, patternType: 'solid' },
        alignment: { horizontal: 'center', vertical: 'center' },
        border: { all: { style: 'thin', color: { rgb: 'FFB8B8B8' } } }
    },
    
    // 특별한 카테고리 스타일들
    categoryYellow: {
        font: { name: 'Noto Sans KR', size: 9, bold: true, color: 'FF000000' },
        fill: { fgColor: { rgb: 'FFF9D6AE' }, patternType: 'solid' },
        alignment: { horizontal: 'center', vertical: 'center' },
        border: { all: { style: 'thin', color: { rgb: 'FFB8B8B8' } } }
    },
    
    categoryBlue: {
        font: { name: 'Noto Sans KR', size: 9, bold: true, color: 'FF000000' },
        fill: { fgColor: { rgb: 'FFD9ECF2' }, patternType: 'solid' },
        alignment: { horizontal: 'center', vertical: 'center' },
        border: { all: { style: 'thin', color: { rgb: 'FFB8B8B8' } } }
    },
    
    categoryBrightYellow: {
        font: { name: 'Noto Sans KR', size: 9, bold: true, color: 'FF000000' },
        fill: { fgColor: { rgb: 'FFFBCF3A' }, patternType: 'solid' },
        alignment: { horizontal: 'center', vertical: 'center' },
        border: { all: { style: 'thin', color: { rgb: 'FFB8B8B8' } } }
    },
    
    // 일반 데이터 셀
    dataCell: {
        font: { name: 'Noto Sans KR', size: 9, bold: false },
        alignment: { horizontal: 'center', vertical: 'center' },
        border: { all: { style: 'thin', color: { rgb: 'FFB8B8B8' } } }
    },
    
    // 안내 문구 스타일 (빨간색)
    warningText: {
        font: { name: 'Noto Sans KR', size: 10, bold: true, color: 'FFFF0000' },
        alignment: { horizontal: 'left', vertical: 'center' }
    },
    
    // 숫자 형식
    numberFormats: {
        percentage: '##0.00\\ "%"',
        squareMeter: '#,##0.000\\ "m²"',
        pyeong: '#,##0.000\\ "평"',
        currency: '\\₩* #,##0'
    }
};

// 정확한 열 너비 설정
function setColumnWidths(sheet) {
    sheet['!cols'] = [
        { wch: 2.6640625 },   // A열
        { wch: 13.21875 },    // B열
        { wch: 24.5546875 },  // C열
        { wch: 26.33203125 }, // D열
        { wch: 26.33203125 }, // E열
        { wch: 26.33203125 }, // F열
        { wch: 26.33203125 }, // G열
        { wch: 26.33203125 }  // H열
    ];
}

// 정확한 행 높이 설정
function setRowHeights(sheet) {
    sheet['!rows'] = [];
    
    // 특별한 높이를 가진 행들
    sheet['!rows'][0] = { hpt: 16.9 };   // 1행
    sheet['!rows'][1] = { hpt: 49.9 };   // 2행
    sheet['!rows'][2] = { hpt: 16.9 };   // 3행
    sheet['!rows'][3] = { hpt: 16.9 };   // 4행
    sheet['!rows'][4] = { hpt: 190.15 }; // 5행 (이미지 영역)
    sheet['!rows'][5] = { hpt: 79.9 };   // 6행
    sheet['!rows'][8] = { hpt: 60.0 };   // 9행 (위치 정보)
    
    // 나머지 행들은 기본 높이 (16.9)
    for (let i = 6; i <= 50; i++) {
        if (i !== 8) { // 9행은 이미 설정됨
            sheet['!rows'][i] = { hpt: 16.9 };
        }
    }
}

// 병합 셀 설정
function setMergedCells(sheet) {
    sheet['!merges'] = [
        { s: { r: 2, c: 1 }, e: { r: 3, c: 2 } },    // B3:C4
        { s: { r: 4, c: 1 }, e: { r: 4, c: 2 } },    // B5:C5
        { s: { r: 5, c: 1 }, e: { r: 5, c: 2 } },    // B6:C6
        { s: { r: 6, c: 1 }, e: { r: 17, c: 1 } },   // B7:B18
        { s: { r: 18, c: 1 }, e: { r: 19, c: 1 } },  // B19:B20
        { s: { r: 20, c: 1 }, e: { r: 22, c: 1 } },  // B21:B23
        { s: { r: 24, c: 1 }, e: { r: 30, c: 1 } },  // B25:B31
        { s: { r: 31, c: 1 }, e: { r: 38, c: 1 } },  // B32:B39
        { s: { r: 39, c: 1 }, e: { r: 43, c: 1 } },  // B40:B44
        { s: { r: 45, c: 1 }, e: { r: 49, c: 1 } }   // B46:B50
    ];
}

// 스타일 적용 함수
function applyTemplateStyles(sheet) {
    // B3 셀 (PRESENT TO) 스타일
    if (sheet['B3']) {
        sheet['B3'].s = convertToSheetJSStyle(templateStyles.presentTo);
    }
    
    // 빌딩명 헤더 (D4~H4) 스타일
    ['D4', 'E4', 'F4', 'G4', 'H4'].forEach(cell => {
        if (sheet[cell]) {
            sheet[cell].s = convertToSheetJSStyle(templateStyles.buildingHeader);
        }
    });
    
    // B5 (고객사 로고 삽입)
    if (sheet['B5']) {
        sheet['B5'].s = convertToSheetJSStyle({
            font: { name: 'Noto Sans KR', size: 11, bold: true },
            fill: { fgColor: { rgb: 'FFFFFFFF' }, patternType: 'solid' },
            alignment: { horizontal: 'center', vertical: 'center' },
            border: { all: { style: 'thin', color: { rgb: 'FFB8B8B8' } } }
        });
    }
    
    // B열 카테고리 스타일 (색상별로 적용)
    const categoryStyles = {
        // 기본 흰색 배경
        white: [7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23],
        // 노란색 배경 (FFF9D6AE)
        yellow: [25, 26, 27, 28, 29, 30, 31],
        // 파란색 배경 (FFD9ECF2)
        blue: [32, 33, 34, 35, 36, 37, 38, 39],
        // 밝은 노란색 배경 (FFFBCF3A)
        brightYellow: [40, 41, 42, 43, 44, 46, 47, 48, 49, 50]
    };
    
    // 각 색상별로 스타일 적용
    categoryStyles.white.forEach(row => {
        const cell = `B${row}`;
        if (sheet[cell]) {
            sheet[cell].s = convertToSheetJSStyle(templateStyles.category);
        }
    });
    
    categoryStyles.yellow.forEach(row => {
        const cell = `B${row}`;
        if (sheet[cell]) {
            sheet[cell].s = convertToSheetJSStyle(templateStyles.categoryYellow);
        }
    });
    
    categoryStyles.blue.forEach(row => {
        const cell = `B${row}`;
        if (sheet[cell]) {
            sheet[cell].s = convertToSheetJSStyle(templateStyles.categoryBlue);
        }
    });
    
    categoryStyles.brightYellow.forEach(row => {
        const cell = `B${row}`;
        if (sheet[cell]) {
            sheet[cell].s = convertToSheetJSStyle(templateStyles.categoryBrightYellow);
        }
    });
    
    // C열 스타일 적용 (카테고리 설명)
    for (let row = 7; row <= 50; row++) {
        const cell = `C${row}`;
        if (sheet[cell]) {
            // 특정 행은 중앙 정렬
            const centerAlignRows = [12, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 40, 41, 42, 43, 44];
            const isCenter = centerAlignRows.includes(row);
            
            sheet[cell].s = convertToSheetJSStyle({
                font: { name: 'Noto Sans KR', size: 9, bold: false },
                alignment: { 
                    horizontal: isCenter ? 'center' : 'left', 
                    vertical: 'center' 
                },
                border: { all: { style: 'thin', color: { rgb: 'FFB8B8B8' } } }
            });
        }
    }
    
    // 데이터 셀 스타일 적용 (D~H열)
    for (let col of ['D', 'E', 'F', 'G', 'H']) {
        for (let row = 5; row <= 50; row++) {
            const cell = `${col}${row}`;
            if (sheet[cell] && !sheet[cell].s) {
                // 정렬 설정
                let alignment = 'center'; // 기본값
                const rightAlignRows = [32, 33, 34, 35, 36, 37, 38, 39, 40, 42, 43, 46, 47, 48];
                const centerAlignRows = [41, 44, 49, 50]; // 41, 44, 49, 50행은 중앙 정렬
                
                if (rightAlignRows.includes(row)) {
                    alignment = 'right';
                }
                
                sheet[cell].s = convertToSheetJSStyle({
                    font: { name: 'Noto Sans KR', size: 9, bold: false },
                    alignment: { 
                        horizontal: alignment, 
                        vertical: 'center',
                        wrapText: row >= 32 && row <= 50 // 32~50행은 텍스트 줄바꿈
                    },
                    border: { all: { style: 'thin', color: { rgb: 'FFB8B8B8' } } }
                });
            }
        }
    }
    
    // 용어 설명 부분 스타일 (52~55행)
    if (sheet['B52']) {
        sheet['B52'].s = convertToSheetJSStyle({
            font: { name: '맑은 고딕', size: 10, bold: true },
            alignment: { horizontal: 'left', vertical: 'center' }
        });
    }
    
    for (let row = 53; row <= 55; row++) {
        const cell = `B${row}`;
        if (sheet[cell]) {
            sheet[cell].s = convertToSheetJSStyle({
                font: { name: '맑은 고딕', size: 10, bold: false },
                alignment: { horizontal: 'left', vertical: 'center' }
            });
        }
    }
    
    // B24 셀 (빨간색 안내 문구)
    if (sheet['B24']) {
        sheet['B24'].s = convertToSheetJSStyle(templateStyles.warningText);
    }
    
    // 숫자 형식 적용
    applyNumberFormats(sheet);
}

// SheetJS 스타일 형식으로 변환
function convertToSheetJSStyle(style) {
    const sheetJSStyle = {};
    
    // 폰트 스타일
    if (style.font) {
        sheetJSStyle.font = {
            name: style.font.name,
            sz: style.font.size,
            bold: style.font.bold,
            color: style.font.color ? { rgb: style.font.color } : undefined
        };
    }
    
    // 채우기 스타일
    if (style.fill) {
        sheetJSStyle.fill = {
            patternType: style.fill.patternType || 'solid',
            fgColor: { rgb: style.fill.fgColor.rgb }
        };
    }
    
    // 정렬 스타일
    if (style.alignment) {
        sheetJSStyle.alignment = {
            horizontal: style.alignment.horizontal,
            vertical: style.alignment.vertical,
            wrapText: style.alignment.wrapText || false
        };
    }
    
    // 테두리
    if (style.border && style.border.all) {
        sheetJSStyle.border = {
            top: style.border.all,
            bottom: style.border.all,
            left: style.border.all,
            right: style.border.all
        };
    }
    
    return sheetJSStyle;
}

// 숫자 형식 적용
function applyNumberFormats(sheet) {
    // 전용률 (%) - 12행
    ['D12', 'E12', 'F12', 'G12', 'H12'].forEach(cell => {
        if (sheet[cell]) {
            sheet[cell].z = templateStyles.numberFormats.percentage;
        }
    });
    
    // m² 형식 - 13, 14, 28, 29행
    [13, 14, 28, 29].forEach(row => {
        ['D', 'E', 'F', 'G', 'H'].forEach(col => {
            const cell = `${col}${row}`;
            if (sheet[cell]) {
                sheet[cell].z = templateStyles.numberFormats.squareMeter;
            }
        });
    });
    
    // 평 형식 - 15, 16, 30, 31행
    [15, 16, 30, 31].forEach(row => {
        ['D', 'E', 'F', 'G', 'H'].forEach(col => {
            const cell = `${col}${row}`;
            if (sheet[cell]) {
                sheet[cell].z = templateStyles.numberFormats.pyeong;
            }
        });
    });
    
    // 원화 형식 - 32~50행 (금액 관련)
    for (let row = 32; row <= 50; row++) {
        ['D', 'E', 'F', 'G', 'H'].forEach(col => {
            const cell = `${col}${row}`;
            if (sheet[cell] && typeof sheet[cell].v === 'number') {
                sheet[cell].z = templateStyles.numberFormats.currency;
            }
        });
    }
}

// 빌딩 데이터 채우기 함수에 추가할 내용
function fillBuildingDataToTemplate(sheet, building, columnIndex) {
    // 기존 코드 유지...
    const col = String.fromCharCode(68 + columnIndex); // D, E, F, G, H
    
    // === 빌딩 기본 정보 (수식 없음) ===
    setCellValue(sheet, `${col}4`, building.name || '');
    setCellValue(sheet, `${col}7`, building.addressJibun || '');
    setCellValue(sheet, `${col}8`, building.address || '');
    setCellValue(sheet, `${col}9`, building.station || '');
    setCellValue(sheet, `${col}10`, building.floors || '');
    setCellValue(sheet, `${col}11`, building.completionYear || '');
    setCellValue(sheet, `${col}12`, building.dedicatedRate || 0);
    setCellValue(sheet, `${col}13`, building.baseFloorArea || 0);
    setCellValue(sheet, `${col}14`, building.baseFloorAreaPy || 0);
    setCellValue(sheet, `${col}15`, building.baseFloorAreaDedicated || 0);
    setCellValue(sheet, `${col}16`, building.baseFloorAreaDedicatedPy || 0);
    setCellValue(sheet, `${col}17`, building.parkingSpace || '');
    setCellValue(sheet, `${col}18`, building.hvac || '');
    setCellValue(sheet, `${col}19`, building.buildingUse || '');
    setCellValue(sheet, `${col}20`, building.structure || '');
    setCellValue(sheet, `${col}21`, building.elevator || '');
    setCellValue(sheet, `${col}22`, building.parkingOperation || '');
    setCellValue(sheet, `${col}23`, building.parkingSpace || '');
    setCellValue(sheet, `${col}24`, building.parkingFee || '');
    
    // === 임차 제안 (기본값 설정) ===
    setCellValue(sheet, `${col}26`, '-');
    setCellValue(sheet, `${col}27`, '-');
    setCellValue(sheet, `${col}28`, '-');
    setCellValue(sheet, `${col}29`, 0);
    setCellValue(sheet, `${col}30`, 0);
    setCellValue(sheet, `${col}31`, 0);
    
    // === 임대 기준 (기본값 0) ===
    for (let row = 32; row <= 44; row++) {
        if (row === 40) {
            // 40행은 수식 (=D32)
            sheet[`${col}40`] = { f: `=${col}32`, t: 'n' };
        } else if (row === 42) {
            // 42행은 수식 (평균 임대료 계산)
            sheet[`${col}42`] = { f: `=${col}33-((${col}33*${col}41)/12)`, t: 'n' };
        } else if (row === 43) {
            // 43행은 수식 (=D34)
            sheet[`${col}43`] = { f: `=${col}34`, t: 'n' };
        } else if (row === 44) {
            // 44행은 수식 (NOC 계산)
            sheet[`${col}44`] = { f: `=((${col}42+${col}43)*(${col}30/${col}31))`, t: 'n' };
        } else {
            setCellValue(sheet, `${col}${row}`, 0);
        }
    }
    
    // === 예상비용 (수식에 의한 자동 계산) ===
    // 46행: 보증금 (=D40*D30)
    sheet[`${col}46`] = { f: `=${col}40*${col}30`, t: 'n' };
    
    // 47행: 평균 월 임대료 (=D42*D30)
    sheet[`${col}47`] = { f: `=${col}42*${col}30`, t: 'n' };
    
    // 48행: 평균 월 관리비 (=D43*D30)
    sheet[`${col}48`] = { f: `=${col}43*${col}30`, t: 'n' };
    
    // 49행: 월 (임대료 + 관리비) (=D47+D48)
    sheet[`${col}49`] = { f: `=${col}47+${col}48`, t: 'n' };
    
    // 50행: 연 실제 부담 고정금액 (=D49*12)
    sheet[`${col}50`] = { f: `=${col}49*12`, t: 'n' };
    
    // === 임차 특이사항 ===
    if (building.description) {
        setCellValue(sheet, `${col}52`, building.description);
    }
}

// exportToExcel 함수 수정 (스타일 적용 부분 추가)
async function exportToExcel() {
    if (selectedBuildings.length === 0) {
        alert('선택된 빌딩이 없습니다.');
        return;
    }
    
    if (selectedBuildings.length > 5) {
        alert('최대 5개까지만 비교할 수 있습니다.');
        return;
    }
    
    try {
        // GitHub Pages 경로 처리
        const basePath = window.location.pathname.includes('/cre1/') ? '/cre1' : '';
        const templatePath = `${basePath}/templates/template.xlsx`;
        
        console.log('템플릿 경로:', templatePath);
        const response = await fetch(templatePath);
        
        if (!response.ok) {
            throw new Error('템플릿 파일을 찾을 수 없습니다.');
        }
        
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, {
            cellStyles: true,
            cellFormulas: true,
            cellDates: true,
            cellNF: true,
            sheetStubs: true
        });
        
        // '후보지' 시트 찾기
        let sheetName = '후보지';
        if (!workbook.Sheets[sheetName]) {
            sheetName = workbook.SheetNames[0];
        }
        const sheet = workbook.Sheets[sheetName];
        
        // 선택된 빌딩 데이터 가져오기
        const buildingsToExport = selectedBuildings.map(id => 
            allBuildings.find(b => b.id === id)
        ).filter(b => b);
        
        // 각 빌딩 데이터 입력
        buildingsToExport.forEach((building, index) => {
            fillBuildingDataToTemplate(sheet, building, index);
        });
        
        // 스타일 적용
        applyTemplateStyles(sheet);
        
        // 병합 셀 설정
        setMergedCells(sheet);
        
        // 열 너비 설정
        setColumnWidths(sheet);
        
        // 행 높이 설정
        setRowHeights(sheet);
        
        // 엑셀 파일 생성
        const wbout = XLSX.write(workbook, {
            bookType: 'xlsx',
            type: 'array',
            cellFormulas: true,
            cellStyles: true,
            cellDates: true
        });
        
        // 다운로드
        const blob = new Blob([wbout], { 
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
        });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `CompList_${getCurrentDate()}.xlsx`;
        a.click();
        URL.revokeObjectURL(url);
        
        alert('✅ Comp List가 생성되었습니다!\n\n' +
              '📋 적용된 스타일:\n' +
              '• 폰트: Noto Sans KR\n' +
              '• 색상: 카테고리별 배경색\n' +
              '• 테두리 및 병합 셀\n' +
              '• 숫자 형식 (%, m², 평, 원화)\n' +
              '• 열 너비 및 행 높이');
        
    } catch (error) {
        console.error('엑셀 생성 오류:', error);
        
        if (error.message.includes('템플릿 파일을 찾을 수 없습니다')) {
            const useBasic = confirm('템플릿 파일을 찾을 수 없습니다.\n기본 형식으로 내보내시겠습니까?');
            if (useBasic) {
                exportToExcelBasic();
            }
        } else {
            alert('엑셀 파일 생성 중 오류가 발생했습니다.');
        }
    }
}