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
        const reportTitle = document.getElementById('report-title').value || '구로&가산디지털단지/반포역 인근';
        
        // 1. 열 너비 설정
        const columnWidths = [
            3,      // A열
            14,     // B열
            25,     // C열
        ];
        // 빌딩 수에 따라 D열부터 추가
        for (let i = 0; i < selectedBuildings.length; i++) {
            columnWidths.push(26.5); // 각 빌딩별 열 너비
        }
        
        worksheet.columns = columnWidths.map(width => ({ width }));
        
        // 2. 행 높이 설정
        worksheet.getRow(1).height = 30;   // 제목
        worksheet.getRow(2).height = 20;   // 규모
        worksheet.getRow(3).height = 20;   // 계약기간
        worksheet.getRow(4).height = 20;   // 위치 설명
        worksheet.getRow(5).height = 25;   // 위치 헤더
        worksheet.getRow(9).height = 190;  // 사진
        
        // 나머지 행들 기본 높이
        for (let i = 6; i <= 84; i++) {
            if (i !== 9) worksheet.getRow(i).height = 18;
        }
        
        // 특별한 행 높이
        for (let i = 10; i <= 18; i++) {
            worksheet.getRow(i).height = 0; // 건물 외관 영역 숨김
        }
        worksheet.getRow(80).height = 60;  // 특이사항
        
        // 3. 상단 헤더 영역 설정
        const endCol = String.fromCharCode(67 + selectedBuildings.length);
        
        // 제목 (1행 전체)
        worksheet.mergeCells(`A1:${endCol}1`);
        const titleCell = worksheet.getCell('A1');
        titleCell.value = `[${companyName} ${reportTitle}]`;
        titleCell.font = { name: 'Arial', size: 14, bold: true };
        titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
        setBordersLG(titleCell);
        
        // 규모 (2행)
        worksheet.mergeCells(`A2:${endCol}2`);
        const scaleCell = worksheet.getCell('A2');
        scaleCell.value = `규모: 건물 ${selectedBuildings.length}개 곳간`;
        scaleCell.font = { name: 'Arial', size: 10 };
        scaleCell.alignment = { horizontal: 'center', vertical: 'middle' };
        setBordersLG(scaleCell);
        
        // 계약기간 (3행)
        worksheet.mergeCells(`A3:${endCol}3`);
        const periodCell = worksheet.getCell('A3');
        const nextYear = new Date();
        nextYear.setFullYear(nextYear.getFullYear() + 1);
        periodCell.value = `계약기간: ${getCurrentDateLG()}~${nextYear.getFullYear()}.${String(nextYear.getMonth() + 1).padStart(2, '0')}.${String(nextYear.getDate()).padStart(2, '0')} (12개월 간)`;
        periodCell.font = { name: 'Arial', size: 10 };
        periodCell.alignment = { horizontal: 'center', vertical: 'middle' };
        setBordersLG(periodCell);
        
        // 위치 설명 (4행)
        worksheet.mergeCells(`A4:${endCol}4`);
        const locationDescCell = worksheet.getCell('A4');
        locationDescCell.value = '위치: 구로&가산디지털단지역 인근 반포역 인근';
        locationDescCell.font = { name: 'Arial', size: 10 };
        locationDescCell.alignment = { horizontal: 'center', vertical: 'middle' };
        setBordersLG(locationDescCell);
        
        // 4. 카테고리 설정
        setupCategoriesLG(worksheet);
        
        // 5. 빌딩별 데이터 입력
        selectedBuildings.forEach((building, index) => {
            const col = String.fromCharCode(68 + index); // D, E, F, G...
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
              `• 재권분석 정보\n` +
              `• 임차 층수 및 면적 정보\n` +
              `• 평면도 이미지\n` +
              `• 특이사항\n\n` +
              `💡 입력한 정보에 따라 비용이 자동 계산됩니다.`);
        
    } catch (error) {
        console.error('엑셀 생성 오류:', error);
        alert('엑셀 파일 생성 중 오류가 발생했습니다.\n' + error.message);
    }
}

// LG 양식 카테고리 설정
function setupCategoriesLG(worksheet) {
    // B열 카테고리 병합
    worksheet.mergeCells('B5:C5');   // 위치
    worksheet.mergeCells('B6:C7');   // 제안
    worksheet.mergeCells('B9:C9');   // 사진
    worksheet.mergeCells('B10:C18'); // 건물 외관
    worksheet.mergeCells('B20:B21'); // 기본정보
    worksheet.mergeCells('B27:B28'); // 재권분석
    worksheet.mergeCells('C30:C32'); // 계약공시지가
    worksheet.mergeCells('B41:B42'); // 대수가능 시기
    worksheet.mergeCells('B43:B48'); // 채권
    worksheet.mergeCells('B49:B50'); // 실질임대료(보증금변환시)
    worksheet.mergeCells('B74:B82'); // 거리
    
    // 카테고리 텍스트 설정
    setCategoryCell(worksheet, 'B5', '위치', 'FF808080', true);
    setCategoryCell(worksheet, 'B6', '제안', 'FFFFFFFF');
    setCategoryCell(worksheet, 'B9', '사진', 'FFFFFFFF');
    setCategoryCell(worksheet, 'B10', '건물 외관', 'FFFFFFFF');
    setCategoryCell(worksheet, 'B20', '기본\n정보', 'FFFFFFFF');
    setCategoryCell(worksheet, 'B26', '소유자 (임대인)', 'FFFFFFFF');
    setCategoryCell(worksheet, 'B27', '재권\n분석', 'FFFFFFFF');
    setCategoryCell(worksheet, 'B41', '대수가능 시기', 'FFFFFFFF');
    setCategoryCell(worksheet, 'B43', '채권', 'FFD9ECF2');
    setCategoryCell(worksheet, 'B49', '실질\n임대료\n(보증금\n변환시)', 'FFF0F8FF');
    setCategoryCell(worksheet, 'B54', '비용감면', 'FFFBCF3A');
    setCategoryCell(worksheet, 'B56', '공사거리', 'FFFBCF3A');
    setCategoryCell(worksheet, 'B58', '주차현황', 'FFFFFFFF');
    setCategoryCell(worksheet, 'B61', '유료주차', 'FFFFFFFF');
    setCategoryCell(worksheet, 'B74', '거리', 'FFFFFFFF');
    
    // C열 항목명 설정
    const items = {
        19: '주소',
        20: '준공일',
        21: '규모',
        22: '연면적',
        23: '기준층 전용면적',
        24: '전용률',
        25: '대지면적',
        27: '재권담보 설정여부',
        28: '선순위 담보 총액',
        29: '공시지가 대비 담보율',
        30: '계약공시지가(2024.1월 기준)',
        33: '통지가격 적용',
        35: '현재 공실',
        40: '평당가격',
        41: '대수가능 시기',
        42: '제안 층',
        43: '평형정보',
        44: '임대면적',
        45: '전용률',
        46: '임대료',
        47: '관리비',
        48: '경비비',
        49: '실질 보증률(월평환 변환)',
        50: '연간 부상임대료 (Y.F)',
        51: '보증금',
        52: '평 임대료',
        53: '평 관리비',
        54: '관리비 내역',
        56: '렌트프리',
        57: '(21개월 기준) 순 년째 비용',
        58: '임대인이 지급 가능',
        59: '총 추가대손',
        60: '무료주차 등(임대면적)',
        61: '무료주차 제공/협의',
        62: '유료주차(VAT별도)',
        68: '평면도',
        80: '특이사항',
        84: '[X] 산업위원회(Rent Free 협의필 임재신): 1-2) 픽'
    };
    
    Object.entries(items).forEach(([row, text]) => {
        const cell = worksheet.getCell(`C${row}`);
        cell.value = text;
        cell.font = { name: 'Arial', size: 9 };
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        setBordersLG(cell);
    });
}

// 카테고리 셀 설정 헬퍼 함수
function setCategoryCell(worksheet, cellAddress, value, bgColor, isWhiteText = false) {
    const cell = worksheet.getCell(cellAddress);
    cell.value = value;
    cell.font = { 
        name: 'Arial', 
        size: 9, 
        bold: true,
        color: isWhiteText ? { argb: 'FFFFFFFF' } : { argb: 'FF000000' }
    };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgColor } };
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    setBordersLG(cell);
}

// LG 양식 빌딩 상세 정보 입력
function fillBuildingDetailsLG(worksheet, building, col) {
    // 위치 (5행)
    setCellLG(worksheet, `${col}5`, '가산디지털단지역', false, 'FFFF9900');
    
    // 제안 (6-7행)
    worksheet.mergeCells(`${col}6:${col}7`);
    const proposalCell = worksheet.getCell(`${col}6`);
    proposalCell.value = building.name || '';
    proposalCell.font = { name: 'Arial', size: 11, bold: true };
    proposalCell.alignment = { horizontal: 'center', vertical: 'middle' };
    proposalCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE7E6E6' } };
    setBordersLG(proposalCell);
    
    // 사진 (9행)
    setCellLG(worksheet, `${col}9`, '', false, 'FFF5F5F5');
    
    // 건물 외관 (10-18행)
    worksheet.mergeCells(`${col}10:${col}18`);
    setCellLG(worksheet, `${col}10`, '', false, 'FFF5F5F5');
    
    // 주소 (19행)
    setCellLG(worksheet, `${col}19`, `${building.addressJibun || ''}\n${building.address || ''}`, true);
    
    // 기본정보
    setCellLG(worksheet, `${col}20`, building.completionYear || '');
    setCellLG(worksheet, `${col}21`, building.floors || '');
    setCellLG(worksheet, `${col}22`, building.grossFloorAreaPy ? `${building.grossFloorAreaPy}평` : '');
    setCellLG(worksheet, `${col}23`, building.baseFloorAreaDedicatedPy ? `${building.baseFloorAreaDedicatedPy}평` : '');
    setCellLG(worksheet, `${col}24`, building.dedicatedRate ? `${building.dedicatedRate}%` : '');
    setCellLG(worksheet, `${col}25`, building.landAreaPy ? `${building.landAreaPy}평` : '');
    
    // 소유자 (26행)
    setCellLG(worksheet, `${col}26`, '');
    
    // 재권분석
    setCellLG(worksheet, `${col}27`, '0.00%');
    setNumericCellLG(worksheet, `${col}28`, 0, '₩#,##0');
    setCellLG(worksheet, `${col}29`, '');
    
    // 계약공시지가 (30-32행)
    worksheet.mergeCells(`${col}30:${col}32`);
    setNumericCellLG(worksheet, `${col}30`, 104363358000, '₩#,##0');
    
    // 통지가격 적용 (33행)
    setCellLG(worksheet, `${col}33`, '');
    
    // 현재 공실 (35행)
    setCellLG(worksheet, `${col}35`, '');
    
    // 평당가격 (40행)
    setCellLG(worksheet, `${col}40`, '');
    
    // 대수가능 시기
    setCellLG(worksheet, `${col}41`, '');
    setCellLG(worksheet, `${col}42`, '4층');
    
    // 채권 정보 - 평당 가격에서 숫자 추출
    const rentPrice = parseFloat(building.rentPricePy?.replace(/[^0-9.]/g, '')) || 45;
    const mgmtFee = parseFloat(building.managementFeePy?.replace(/[^0-9.]/g, '')) || 25;
    
    setCellLG(worksheet, `${col}43`, '217평');
    setCellLG(worksheet, `${col}44`, '467평');
    setCellLG(worksheet, `${col}45`, '46.39%');
    setCellLG(worksheet, `${col}46`, '');
    setCellLG(worksheet, `${col}47`, '');
    setCellLG(worksheet, `${col}48`, '');
    
    // 실질임대료
    setCellLG(worksheet, `${col}49`, `@${rentPrice*10000}+${mgmtFee*10000}`);
    setCellLG(worksheet, `${col}50`, '@96,135');
    setCellLG(worksheet, `${col}51`, '1.0개월');
    setCellLG(worksheet, `${col}52`, '0.0개월');
    setCellLG(worksheet, `${col}53`, '0.0개월');
    
    // 비용감면
    setCellLG(worksheet, `${col}54`, '');
    
    // 공사거리
    setCellLG(worksheet, `${col}56`, '없음');
    setCellLG(worksheet, `${col}57`, '');
    
    // 주차현황
    setCellLG(worksheet, `${col}58`, '');
    setCellLG(worksheet, `${col}59`, '');
    setCellLG(worksheet, `${col}60`, '');
    setCellLG(worksheet, `${col}61`, '');
    setCellLG(worksheet, `${col}62`, '');
    
    // 평면도 (68행)
    setCellLG(worksheet, `${col}68`, '', false, 'FFF5F5F5');
    
    // 거리 (74-82행)
    worksheet.mergeCells(`${col}74:${col}82`);
    const distanceCell = worksheet.getCell(`${col}74`);
    distanceCell.value = '';
    distanceCell.font = { name: 'Arial', size: 8 };
    distanceCell.alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
    setBordersLG(distanceCell);
    
    // 특이사항 (80행)
    setCellLG(worksheet, `${col}80`, '', true);
    
    // 마지막 행 (84행)
    setCellLG(worksheet, `${col}84`, '');
}

// LG 양식 셀 설정 헬퍼 함수
function setCellLG(worksheet, address, value, wrap = false, bgColor = null) {
    const cell = worksheet.getCell(address);
    cell.value = value;
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: wrap };
    if (bgColor) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgColor } };
    }
    applyDataCellStyleLG(cell);
}

function setNumericCellLG(worksheet, address, value, format = '#,##0') {
    const cell = worksheet.getCell(address);
    cell.value = value;
    cell.numFmt = format;
    applyDataCellStyleLG(cell);
}

function applyDataCellStyleLG(cell) {
    cell.font = { name: 'Arial', size: 9 };
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
    const endCol = 67 + buildingCount; // D열부터 시작
    
    // 주요 데이터 영역에만 테두리 적용 (1-84행)
    for (let row = 1; row <= 84; row++) {
        // A열부터 C열
        ['A', 'B', 'C'].forEach(col => {
            const cell = worksheet.getCell(`${col}${row}`);
            if (!cell.border) {
                setBordersLG(cell);
            }
        });
        
        // 빌딩 데이터 열들
        for (let col = 68; col <= endCol; col++) {
            const colLetter = String.fromCharCode(col);
            const cell = worksheet.getCell(`${colLetter}${row}`);
            if (!cell.border) {
                setBordersLG(cell);
            }
        }
    }
}