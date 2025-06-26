// ===== LG Comp List 템플릿 구조 생성 =====

// 워크시트 기본 설정
function setupWorksheet(worksheet) {
    // 시트 이름은 이미 설정됨 (COMP)
    
    // 기본 행 높이 설정 (15)
    for (let i = 1; i <= LG_TEMPLATE_CONFIG.maxRow; i++) {
        worksheet.getRow(i).height = 15;
    }
    
    // 특수 행 높이 설정
    worksheet.getRow(19).height = 30;
    worksheet.getRow(20).height = 30;
    worksheet.getRow(26).height = 40.5;
    worksheet.getRow(53).height = 35.25;
    worksheet.getRow(54).height = 21.75;
    worksheet.getRow(56).height = 12;
    worksheet.getRow(62).height = 33;
    worksheet.getRow(83).height = 33.75;
}

// LG 템플릿 생성 메인 함수
function createLGTemplate(workbook, worksheet, buildings, companyName, reportTitle) {
    // 1. 워크시트 기본 설정
    setupWorksheet(worksheet);
    
    // 2. 헤더 영역 생성 (1-4행)
    createHeaderSection(worksheet, reportTitle);
    
    // 3. 섹션 타이틀 생성
    createSectionTitles(worksheet);
    
    // 4. B열 라벨 생성
    createRowLabels(worksheet);
    
    // 5. 빌딩 열 설정
    createBuildingColumns(worksheet, buildings);
    
    // 6. 병합 셀 적용
    applyMergedCells(worksheet);
    
    // 7. 테두리 적용
    applyBorders(worksheet);
    
    // 8. 하단 주석 추가
    addFooterNotes(worksheet);
}

// 헤더 섹션 생성
function createHeaderSection(worksheet, reportTitle) {
    // A1: 임차제안 제목
    const titleCell = worksheet.getCell('A1');
    titleCell.value = reportTitle || '[임차제안 제목을 입력하세요. ]';
    titleCell.font = { size: 14, bold: true };
    titleCell.alignment = { horizontal: 'left', vertical: 'top' };
    
    // A2-A4: 제안 정보
    worksheet.getCell('A2').value = '- 규모: 전용 0000PY 이상';
    worksheet.getCell('A3').value = '- 계약기간: 2025.00.00~2025.00.00 (00개월 간)';
    worksheet.getCell('A4').value = '- 위치: 0000역 인근';
    
    // 스타일 적용
    for (let i = 2; i <= 4; i++) {
        const cell = worksheet.getCell(`A${i}`);
        cell.font = { size: 10 };
        cell.alignment = { horizontal: 'left', vertical: 'middle' };
    }
}

// 섹션 타이틀 생성
function createSectionTitles(worksheet) {
    const sections = {
        'A6': '위치',
        'A7': '제안',
        'A9': '건물 외관',
        'A18': '기초\n정보',
        'A26': '채권\n분석',
        'A33': '현재 공실',
        'A40': '제안',
        'A45': '기준층\n임대기준',
        'A48': '실질\n임대기준',
        'A50': '비용검토',
        'A56': '공사기간\nFAVOR',
        'A59': '주차현황',
        'A63': '기타'
    };
    
    Object.entries(sections).forEach(([cellRef, title]) => {
        const cell = worksheet.getCell(cellRef);
        cell.value = title;
        cell.font = { size: 10, bold: true };
        cell.alignment = { 
            horizontal: 'center', 
            vertical: 'middle',
            wrapText: true
        };
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFE0E0E0' }
        };
    });
}

// B열 라벨 생성
function createRowLabels(worksheet) {
    Object.entries(LG_TEMPLATE_CONFIG.labels).forEach(([row, label]) => {
        const cell = worksheet.getCell(`B${row}`);
        cell.value = label;
        cell.font = { size: 10 };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        
        // 배경색 설정
        if (row >= 45 && row <= 47) {
            // 기준층 임대기준
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFFFF2CC' }
            };
        } else if (row >= 48 && row <= 49) {
            // 실질 임대기준
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFE2EFDA' }
            };
        }
    });
    
    // C열 특수 라벨
    worksheet.getCell('C30').value = '공시지가 대비 담보율';
    worksheet.getCell('C32').value = '토지가격 적용';
}

// 빌딩 열 설정
function createBuildingColumns(worksheet, buildings) {
    buildings.forEach((building, index) => {
        if (index >= 6) return; // 최대 6개
        
        const startCol = LG_TEMPLATE_CONFIG.buildingColumns[index];
        const colIndex = LG_UTILS.getColumnIndex(startCol);
        
        // 빌딩명 (6행)
        const nameCell = worksheet.getCell(`${startCol}6`);
        nameCell.value = building.name;
        nameCell.font = { size: 12, bold: true };
        nameCell.alignment = { horizontal: 'center', vertical: 'middle' };
        nameCell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF4472C4' }
        };
        nameCell.font.color = { argb: 'FFFFFFFF' };
        
        // 공실 테이블 헤더 (33행)
        const headers = ['층', '전용', '임대'];
        headers.forEach((header, offset) => {
            const col = LG_UTILS.getColumnLetter(colIndex + offset);
            const cell = worksheet.getCell(`${col}33`);
            cell.value = header;
            cell.font = { size: 10, bold: true };
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFD9D9D9' }
            };
        });
        
        // 소계 라벨 (39행)
        worksheet.getCell(`${startCol}39`).value = '소계';
        worksheet.getCell(`${startCol}39`).font = { bold: true };
    });
}

// 병합 셀 적용
function applyMergedCells(worksheet) {
    // 헤더 영역
    worksheet.mergeCells('A1:V1');
    worksheet.mergeCells('A2:V2');
    worksheet.mergeCells('A3:V3');
    worksheet.mergeCells('A4:V4');
    
    // 위치/제안 영역
    worksheet.mergeCells('A6:D6');
    worksheet.mergeCells('A7:D8');
    
    // 건물 외관
    worksheet.mergeCells('A9:D17');
    
    // 섹션 타이틀 병합
    worksheet.mergeCells('A18:A25'); // 기초정보
    worksheet.mergeCells('A26:A32'); // 채권분석
    worksheet.mergeCells('A33:D39'); // 현재 공실
    worksheet.mergeCells('A40:A44'); // 제안
    worksheet.mergeCells('A45:A47'); // 기준층 임대기준
    worksheet.mergeCells('A48:A49'); // 실질 임대기준
    worksheet.mergeCells('A50:A55'); // 비용검토
    worksheet.mergeCells('A56:A58'); // 공사기간
    worksheet.mergeCells('A59:A62'); // 주차현황
    worksheet.mergeCells('A63:A83'); // 기타
    
    // B열 라벨 병합
    const bMerges = [
        'B18:D18', 'B19:D19', 'B20:D20', 'B21:D21', 'B22:D22',
        'B23:D23', 'B24:D24', 'B25:D25', 'B26:D26', 'B27:D27',
        'B28:D28', 'B29:D29', 'B31:D31', 'C30:D30', 'C32:D32',
        'B40:D40', 'B41:D41', 'B42:D42', 'B43:D43', 'B44:D44',
        'B45:D45', 'B46:D46', 'B47:D47', 'B48:D48', 'B49:D49',
        'B50:D50', 'B51:D51', 'B52:D52', 'B53:D53', 'B54:D54',
        'B55:D55', 'B56:D56', 'B57:D58', 'B59:D59', 'B60:D60',
        'B61:D61', 'B62:D62', 'B63:D72', 'B73:D83'
    ];
    
    bMerges.forEach(range => worksheet.mergeCells(range));
    
    // 빌딩별 병합 (6개 빌딩)
    LG_TEMPLATE_CONFIG.buildingColumns.forEach((startCol, index) => {
        const colIndex = LG_UTILS.getColumnIndex(startCol);
        const endCol = LG_UTILS.getColumnLetter(colIndex + 2);
        
        // 빌딩명
        worksheet.mergeCells(`${startCol}6:${endCol}6`);
        worksheet.mergeCells(`${startCol}7:${endCol}7`);
        worksheet.mergeCells(`${startCol}8:${endCol}8`);
        
        // 이미지 영역
        worksheet.mergeCells(`${startCol}9:${endCol}17`);
        
        // 데이터 영역
        const mergeRanges = [
            '18:18', '19:19', '20:20', '21:21', '22:22', '23:23', 
            '24:24', '25:25', '26:26', '27:27', '29:29', '30:30',
            '31:31', '32:32', '40:40', '41:41', '42:42', '43:43',
            '44:44', '45:45', '46:46', '47:47', '48:48', '49:49',
            '50:50', '51:51', '52:52', '53:53', '54:54', '55:55',
            '56:56', '57:58', '59:59', '60:60', '61:61', '62:62'
        ];
        
        mergeRanges.forEach(range => {
            const [startRow, endRow] = range.split(':');
            worksheet.mergeCells(`${startCol}${startRow}:${endCol}${endRow}`);
        });
        
        // 평면도 & 특이사항
        worksheet.mergeCells(`${startCol}63:${endCol}72`);
        worksheet.mergeCells(`${startCol}73:${endCol}83`);
    });
    
    // 하단 주석
    worksheet.mergeCells('A84:V84');
    worksheet.mergeCells('A85:V85');
}

// 테두리 적용
function applyBorders(worksheet) {
    // 전체 영역에 테두리 적용
    const borderStyle = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };
    
    // 주요 영역 테두리
    for (let row = 1; row <= 83; row++) {
        for (let col = 1; col <= 22; col++) { // V열까지
            const cell = worksheet.getCell(row, col);
            if (cell.value !== undefined || cell.isMerged) {
                cell.border = borderStyle;
            }
        }
    }
}

// 하단 주석 추가
function addFooterNotes(worksheet) {
    const notes = [
        '1) 실질임대료(Rent Free 반영한 임대가)  / 2) 월 납부액 = 월 실질임대료 + 월관리비 (초기년도 기준으로 인상률 미반영)',
        '3) 연간납부비용 = 연임대료 + 연관리비 (초기년도 기준으로 인상률 미반영, 보증금 미반영)  4) RF : Rent Free (임대료 무상, 관리비 부과)  5) FO : Fit-out (인테리어공사기간(무상 임대료 무상, 관리비 부과)'
    ];
    
    notes.forEach((note, index) => {
        const cell = worksheet.getCell(`A${84 + index}`);
        cell.value = note;
        cell.font = { size: 9 };
        cell.alignment = { horizontal: 'left', vertical: 'middle' };
    });
}

// 전역 함수로 등록
window.createLGTemplate = createLGTemplate;
window.validateTemplate = function(worksheet) {
    // 템플릿 검증
    return worksheet.getCell('A1').value !== undefined;
};
