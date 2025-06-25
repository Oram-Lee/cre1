// ===== LG용 Comp List 생성 함수 =====

// 현재 날짜 반환 (LG 양식용)
function getCurrentDateLG() {
    const now = new Date();
    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const day = String(now.getDate()).padStart(2, '0');
    return `${year}.${month}.${day}`;
}

// LG 양식으로 엑셀 생성
async function generateExcelLG() {
    if (selectedBuildings.length === 0) {
        alert('빌딩을 선택해주세요.');
        return;
    }
    
    if (selectedBuildings.length > 6) {
        alert('LG 양식은 최대 6개까지만 비교할 수 있습니다.');
        return;
    }
    
    try {
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('후보지');
        
        // 회사명과 제목 가져오기
        const companyName = document.getElementById('company-name').value || 'LG CNS';
        const reportTitle = document.getElementById('report-title').value || '단기임차 가능 공간';
        
        // 1. 열 너비 설정 (이미지 참고하여 조정)
        const columnWidths = [
            2.5,   // A열
            13,    // B열
            20,    // C열
        ];
        // 빌딩 수에 따라 D열부터 추가
        for (let i = 0; i < selectedBuildings.length; i++) {
            columnWidths.push(26); // 각 빌딩별 열 너비
        }
        
        worksheet.columns = columnWidths.map(width => ({ width }));
        
        // 2. 행 높이 설정
        worksheet.getRow(1).height = 40;   // 제목 행
        worksheet.getRow(2).height = 35;   // 부제목 행
        worksheet.getRow(3).height = 20;   // 날짜 행
        worksheet.getRow(4).height = 25;   // 빌딩명 헤더
        worksheet.getRow(5).height = 180;  // 빌딩 이미지
        worksheet.getRow(6).height = 80;   // 빌딩 개요
        
        // 나머지 행들 기본 높이
        for (let i = 7; i <= 80; i++) {
            worksheet.getRow(i).height = 18;
        }
        worksheet.getRow(9).height = 50;  // 위치 정보 (더 높게)
        worksheet.getRow(64).height = 180; // 평면도
        worksheet.getRow(75).height = 60;  // 특이사항
        
        // 3. 상단 헤더 영역 설정
        const titleEndCol = String.fromCharCode(67 + selectedBuildings.length);
        
        // 회사 로고 및 제목 (1행)
        worksheet.mergeCells(`A1:${titleEndCol}1`);
        const titleCell = worksheet.getCell('A1');
        titleCell.value = `[${companyName} ${reportTitle}]`;
        titleCell.font = { name: 'Noto Sans KR', size: 20, bold: true };
        titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
        titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFA50034' } }; // LG 레드
        titleCell.font.color = { argb: 'FFFFFFFF' };
        
        // 부제목 (2행)
        worksheet.mergeCells(`A2:C2`);
        const subtitleCell = worksheet.getCell('A2');
        subtitleCell.value = '- 후보 건물: 총 ' + selectedBuildings.length + '개 곳 -';
        subtitleCell.font = { name: 'Noto Sans KR', size: 12 };
        subtitleCell.alignment = { horizontal: 'center', vertical: 'middle' };
        
        // 작성일자 (3행)
        worksheet.mergeCells(`D2:${titleEndCol}2`);
        const dateCell = worksheet.getCell('D2');
        dateCell.value = `작성기간: ${getCurrentDateLG()} (12개월 간)`;
        dateCell.font = { name: 'Noto Sans KR', size: 10 };
        dateCell.alignment = { horizontal: 'right', vertical: 'middle' };
        
        // 4. 카테고리 설정
        setupCategoriesLG(worksheet);
        
        // 5. 빌딩별 데이터 입력
        selectedBuildings.forEach((building, index) => {
            const col = String.fromCharCode(68 + index); // D, E, F, G...
            
            // 빌딩명 헤더 (4행)
            const nameCell = worksheet.getCell(`${col}4`);
            nameCell.value = building.name;
            nameCell.font = { name: 'Noto Sans KR', size: 11, bold: true };
            nameCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9D9D9' } };
            nameCell.alignment = { horizontal: 'center', vertical: 'middle' };
            setBordersLG(nameCell);
            
            // 빌딩 이미지 영역 (5행)
            const imgCell = worksheet.getCell(`${col}5`);
            imgCell.value = '빌딩 외관 이미지';
            imgCell.font = { name: 'Noto Sans KR', size: 10, italic: true, color: { argb: 'FF999999' } };
            imgCell.alignment = { horizontal: 'center', vertical: 'middle' };
            imgCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF5F5F5' } };
            setBordersLG(imgCell);
            
            // 빌딩 상세 정보 입력
            fillBuildingDetailsLG(worksheet, building, col);
        });
        
        // 6. 테두리 설정
        applyBordersLG(worksheet, selectedBuildings.length);
        
        // 7. 파일 저장
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        saveAs(blob, `CompList_LG_${getCurrentDateLG().replace(/\./g, '')}.xlsx`);
        
        alert(`✅ LG용 Comp List 생성 완료!\n\n` +
              `📊 빌딩 ${selectedBuildings.length}개의 정보가 입력되었습니다.\n\n` +
              `📝 추가 입력 필요 항목:\n` +
              `• 빌딩 외관 이미지\n` +
              `• 평면도 이미지\n` +
              `• 임차 제안 상세 정보\n` +
              `• 특이사항\n\n` +
              `💡 입력한 정보에 따라 비용이 자동 계산됩니다.`);
        
    } catch (error) {
        console.error('엑셀 생성 오류:', error);
        alert('엑셀 파일 생성 중 오류가 발생했습니다.\n' + error.message);
    }
}

// LG 양식 카테고리 설정
function setupCategoriesLG(worksheet) {
    // 카테고리 병합
    worksheet.mergeCells('B4:C4');   // 항목 헤더
    worksheet.mergeCells('B5:C5');   // 건물 이미지
    worksheet.mergeCells('B6:C6');   // 빌딩개요/일반
    worksheet.mergeCells('B7:B18');  // 건물 현황
    worksheet.mergeCells('B19:B24'); // 현황/세부
    worksheet.mergeCells('B25:B31'); // 임차 제안
    worksheet.mergeCells('B32:B39'); // 임대 기준
    worksheet.mergeCells('B40:B44'); // 임대기준 조정
    worksheet.mergeCells('B46:B50'); // 예상비용
    worksheet.mergeCells('B52:B56'); // 예상비용2
    worksheet.mergeCells('B58:B62'); // 공실가감
    worksheet.mergeCells('B64:C64'); // 평면도
    worksheet.mergeCells('B75:C75'); // 특이사항
    
    // 카테고리 텍스트와 스타일
    const categories = {
        'B4': { text: '항목', bg: 'FF808080', color: 'FFFFFFFF' },
        'B6': { text: '빌딩개요/일반', bg: 'FFE7E6E6' },
        'B7': { text: '건물 현황', bg: 'FFE7E6E6' },
        'B19': { text: '현황/세부', bg: 'FFE7E6E6' },
        'B25': { text: '임차 제안', bg: 'FFF9D6AE' },
        'B32': { text: '임대 기준', bg: 'FFD9ECF2' },
        'B40': { text: '임대기준 조정', bg: 'FFD9ECF2' },
        'B46': { text: '예상비용', bg: 'FFFBCF3A' },
        'B52': { text: '예상비용', bg: 'FFFBCF3A' },
        'B58': { text: '공실가감', bg: 'FFCCE5FF' },
        'B64': { text: '평면도', bg: 'FFE7E6E6' },
        'B75': { text: '특이사항', bg: 'FFE7E6E6' }
    };
    
    Object.entries(categories).forEach(([cell, config]) => {
        const categoryCell = worksheet.getCell(cell);
        categoryCell.value = config.text;
        categoryCell.font = { 
            name: 'Noto Sans KR', 
            size: 10, 
            bold: true,
            color: config.color ? { argb: config.color } : { argb: 'FF000000' }
        };
        categoryCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: config.bg } };
        categoryCell.alignment = { horizontal: 'center', vertical: 'middle' };
        setBordersLG(categoryCell);
    });
    
    // C열 항목명 설정
    const items = {
        7: '주소 지번',
        8: '도로명 주소',
        9: '위치',
        10: '빌딩 규모',
        11: '준공연도',
        12: '전용률 (%)',
        13: '기준층 임대면적 (m²)',
        14: '기준층 임대면적 (평)',
        15: '기준층 전용면적 (m²)',
        16: '기준층 전용면적 (평)',
        17: '엘레베이터',
        18: '냉난방 방식',
        19: '건물용도',
        20: '구조',
        21: '주차 대수 정보',
        22: '주차비',
        23: '주차 대수',
        24: '대지면적', // 새로 추가
        25: '최적 임차 층수',
        26: '입주 가능 시기',
        27: '거래유형',
        28: '임대면적 (m²)',
        29: '전용면적 (m²)',
        30: '임대면적 (평)',
        31: '전용면적 (평)',
        32: '월 평당 보증금',
        33: '월 평당 임대료',
        34: '월 평당 관리비',
        35: '월 평당 지출비용',
        36: '총 보증금',
        37: '월 임대료 총액',
        38: '월 관리비 총액',
        39: '월 지출비용',
        40: '렌트프리 적용 시',
        41: '렌트프리 (개월/년)',
        42: '평균임대료',
        43: '관리비',
        44: 'NOC',
        46: '보증금',
        47: '평균 월 임대료',
        48: '평균 월 관리비',
        49: '월 (임대료+관리비)',
        50: '연 실제 부담 고정금액',
        52: '비용감면',
        53: '렌트프리(1.5개월)',
        54: '연내부수익률(IRR)',
        55: '순현재가치(NPV)',
        56: '투자수익률(ROI)',
        58: '순임대면적 차이',
        59: '주차장 대비 임직원',
        60: '접근성 차이',
        61: '건물 노후도',
        62: '종합 평가'
    };
    
    Object.entries(items).forEach(([row, text]) => {
        const cell = worksheet.getCell(`C${row}`);
        cell.value = text;
        cell.font = { name: 'Noto Sans KR', size: 9 };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        setBordersLG(cell);
    });
}

// LG 양식 빌딩 상세 정보 입력
function fillBuildingDetailsLG(worksheet, building, col) {
    // 빌딩개요/일반 (6행)
    setCellLG(worksheet, `${col}6`, building.description || '', true);
    
    // 건물 현황
    setCellLG(worksheet, `${col}7`, building.addressJibun || '');
    setCellLG(worksheet, `${col}8`, building.address || '');
    setCellLG(worksheet, `${col}9`, building.station || '', true); // 위치는 줄바꿈
    setCellLG(worksheet, `${col}10`, building.floors || '');
    setCellLG(worksheet, `${col}11`, building.completionYear || '');
    
    // 전용률 (퍼센트 표시)
    const dedicatedRateCell = worksheet.getCell(`${col}12`);
    dedicatedRateCell.value = building.dedicatedRate ? (building.dedicatedRate / 100) : 0;
    dedicatedRateCell.numFmt = '0.00%';
    applyDataCellStyleLG(dedicatedRateCell);
    
    // 면적 정보
    setNumericCellLG(worksheet, `${col}13`, building.baseFloorArea || 0, '#,##0.000');
    setNumericCellLG(worksheet, `${col}14`, building.baseFloorAreaPy || 0, '#,##0.000');
    setNumericCellLG(worksheet, `${col}15`, building.baseFloorAreaDedicated || 0, '#,##0.000');
    setNumericCellLG(worksheet, `${col}16`, building.baseFloorAreaDedicatedPy || 0, '#,##0.000');
    
    // 시설 정보
    setCellLG(worksheet, `${col}17`, building.elevator || '');
    setCellLG(worksheet, `${col}18`, building.hvac || '');
    setCellLG(worksheet, `${col}19`, building.buildingUse || '');
    setCellLG(worksheet, `${col}20`, building.structure || '');
    setCellLG(worksheet, `${col}21`, building.parkingSpace || '');
    setCellLG(worksheet, `${col}22`, building.parkingFee || '');
    setCellLG(worksheet, `${col}23`, building.parkingSpace || '');
    setCellLG(worksheet, `${col}24`, building.landAreaPy ? `${building.landAreaPy}평` : '');
    
    // 임차 제안 (예시 데이터)
    setCellLG(worksheet, `${col}25`, '전층');
    setCellLG(worksheet, `${col}26`, '즉시');
    setCellLG(worksheet, `${col}27`, '임대');
    
    // 임대면적/전용면적 - 평 기준 입력, m²는 수식으로 자동 계산
    worksheet.getCell(`${col}28`).value = { formula: `ROUNDDOWN(${col}30*3.305785,3)` };
    worksheet.getCell(`${col}29`).value = { formula: `ROUNDDOWN(${col}31*3.305785,3)` };
    worksheet.getCell(`${col}28`).numFmt = '#,##0.000';
    worksheet.getCell(`${col}29`).numFmt = '#,##0.000';
    
    // 평 단위 (사용자 입력 가능하도록)
    setNumericCellLG(worksheet, `${col}30`, 217, '#,##0');
    setNumericCellLG(worksheet, `${col}31`, 130, '#,##0');
    
    // 평당 가격에서 숫자만 추출
    const rentPrice = parseFloat(building.rentPricePy?.replace(/[^0-9.]/g, '')) * 10000 || 0;
    const mgmtFee = parseFloat(building.managementFeePy?.replace(/[^0-9.]/g, '')) * 10000 || 0;
    const deposit = parseFloat(building.depositPy?.replace(/[^0-9.]/g, '')) * 10000 || 0;
    
    // 임대 기준
    setNumericCellLG(worksheet, `${col}32`, deposit, '₩#,##0');
    setNumericCellLG(worksheet, `${col}33`, rentPrice, '₩#,##0');
    setNumericCellLG(worksheet, `${col}34`, mgmtFee, '₩#,##0');
    
    // 합계 계산
    worksheet.getCell(`${col}35`).value = { formula: `${col}33+${col}34` };
    worksheet.getCell(`${col}35`).numFmt = '₩#,##0';
    
    // 총액 계산
    worksheet.getCell(`${col}36`).value = { formula: `${col}32*${col}30` };
    worksheet.getCell(`${col}37`).value = { formula: `${col}33*${col}30` };
    worksheet.getCell(`${col}38`).value = { formula: `${col}34*${col}30` };
    worksheet.getCell(`${col}39`).value = { formula: `${col}37+${col}38` };
    
    [36, 37, 38, 39].forEach(row => {
        worksheet.getCell(`${col}${row}`).numFmt = '₩#,##0';
        applyDataCellStyleLG(worksheet.getCell(`${col}${row}`));
    });
    
    // 렌트프리 적용
    setCellLG(worksheet, `${col}40`, '적용');
    setNumericCellLG(worksheet, `${col}41`, 2, '0.0'); // 2개월
    
    // 실질 임대료 계산
    worksheet.getCell(`${col}42`).value = { formula: `${col}33-((${col}33*${col}41)/12)` };
    worksheet.getCell(`${col}43`).value = { formula: `${col}34` };
    worksheet.getCell(`${col}44`).value = { formula: `${col}42+${col}43` };
    
    [42, 43, 44].forEach(row => {
        worksheet.getCell(`${col}${row}`).numFmt = '₩#,##0';
        applyDataCellStyleLG(worksheet.getCell(`${col}${row}`));
    });
    
    // 예상비용
    worksheet.getCell(`${col}46`).value = { formula: `${col}36` };
    worksheet.getCell(`${col}47`).value = { formula: `${col}42*${col}30` };
    worksheet.getCell(`${col}48`).value = { formula: `${col}43*${col}30` };
    worksheet.getCell(`${col}49`).value = { formula: `${col}47+${col}48` };
    worksheet.getCell(`${col}50`).value = { formula: `${col}49*12` };
    
    [46, 47, 48, 49, 50].forEach(row => {
        worksheet.getCell(`${col}${row}`).numFmt = '₩#,##0';
        applyDataCellStyleLG(worksheet.getCell(`${col}${row}`));
    });
    
    // 평면도 영역 (64행)
    const floorPlanCell = worksheet.getCell(`${col}64`);
    floorPlanCell.value = '평면도 이미지';
    floorPlanCell.font = { name: 'Noto Sans KR', size: 10, italic: true, color: { argb: 'FF999999' } };
    floorPlanCell.alignment = { horizontal: 'center', vertical: 'middle' };
    floorPlanCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF5F5F5' } };
    setBordersLG(floorPlanCell);
    
    // 특이사항 (75행)
    const remarkCell = worksheet.getCell(`${col}75`);
    remarkCell.value = building.remarks || '';
    remarkCell.font = { name: 'Noto Sans KR', size: 9 };
    remarkCell.alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
    setBordersLG(remarkCell);
}

// LG 양식 셀 설정 헬퍼 함수
function setCellLG(worksheet, address, value, wrap = false) {
    const cell = worksheet.getCell(address);
    cell.value = value;
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: wrap };
    applyDataCellStyleLG(cell);
}

function setNumericCellLG(worksheet, address, value, format = '#,##0') {
    const cell = worksheet.getCell(address);
    cell.value = value;
    cell.numFmt = format;
    applyDataCellStyleLG(cell);
}

function applyDataCellStyleLG(cell) {
    cell.font = { name: 'Noto Sans KR', size: 9 };
    if (!cell.alignment) {
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
    }
    setBordersLG(cell);
}

function setBordersLG(cell) {
    cell.border = {
        top: { style: 'thin', color: { argb: 'FF000000' } },
        left: { style: 'thin', color: { argb: 'FF000000' } },
        bottom: { style: 'thin', color: { argb: 'FF000000' } },
        right: { style: 'thin', color: { argb: 'FF000000' } }
    };
}

// LG 양식 전체 테두리 적용
function applyBordersLG(worksheet, buildingCount) {
    const endCol = String.fromCharCode(67 + buildingCount);
    
    // 모든 데이터 영역에 테두리 적용
    for (let row = 1; row <= 80; row++) {
        for (let col = 65; col <= 67 + buildingCount; col++) { // A부터 끝 열까지
            const cell = worksheet.getCell(row, col - 64);
            if (!cell.border) {
                setBordersLG(cell);
            }
        }
    }
}