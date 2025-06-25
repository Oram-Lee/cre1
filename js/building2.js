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
    
    if (selectedBuildings.length > 7) {
        alert('LG 양식은 최대 7개까지만 비교할 수 있습니다.');
        return;
    }
    
    try {
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('후보지');
        
        // 회사명과 제목 가져오기
        const companyName = document.getElementById('company-name').value || 'LG CNS';
        const reportTitle = document.getElementById('report-title').value || '구로&가산디지털단지/반포역 인근 단기임차 가능 공간';
        
        // 1. 열 너비 설정
        const columnWidths = [
            2,      // A열
            10,     // B열 (위치)
            20,     // C열 (제안)
        ];
        // 빌딩 수에 따라 D열부터 추가
        for (let i = 0; i < selectedBuildings.length; i++) {
            columnWidths.push(22); // 각 빌딩별 열 너비
        }
        
        worksheet.columns = columnWidths.map(width => ({ width }));
        
        // 2. 행 높이 설정
        worksheet.getRow(1).height = 30;   // 제목
        worksheet.getRow(2).height = 20;   // 계약기간
        worksheet.getRow(3).height = 20;   // 위치
        worksheet.getRow(4).height = 20;   // 빈 행
        worksheet.getRow(5).height = 25;   // 헤더
        worksheet.getRow(6).height = 180;  // 건물 외관
        
        // 나머지 행들 기본 높이
        for (let i = 7; i <= 40; i++) {
            worksheet.getRow(i).height = 18;
        }
        
        // 특별한 행 높이
        worksheet.getRow(32).height = 120; // 평면도
        
        // 3. 상단 헤더 영역 설정
        const endCol = String.fromCharCode(67 + selectedBuildings.length);
        
        // 제목 (1행)
        worksheet.mergeCells(`A1:${endCol}1`);
        const titleCell = worksheet.getCell('A1');
        titleCell.value = `[${companyName} ${reportTitle}]`;
        titleCell.font = { name: 'Arial', size: 16, bold: true };
        titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
        setBordersLG(titleCell);
        
        // 계약기간 (2행)
        worksheet.mergeCells(`A2:${endCol}2`);
        const periodCell = worksheet.getCell('A2');
        periodCell.value = `- 계약기간: ${getCurrentDateLG()}~${getCurrentDateLG()} (12개월 간) -`;
        periodCell.font = { name: 'Arial', size: 10 };
        periodCell.alignment = { horizontal: 'center', vertical: 'middle' };
        
        // 위치 (3행)
        worksheet.mergeCells(`A3:${endCol}3`);
        const locationCell = worksheet.getCell('A3');
        locationCell.value = '- 위치: 구로&가산디지털단지역 인근 반포역 인근 -';
        locationCell.font = { name: 'Arial', size: 10 };
        locationCell.alignment = { horizontal: 'center', vertical: 'middle' };
        
        // 회사 로고 영역 (우측 상단)
        const logoCol = String.fromCharCode(67 + selectedBuildings.length - 1);
        worksheet.getCell(`${logoCol}1`).value = 'S&I Corp.';
        worksheet.getCell(`${logoCol}1`).font = { name: 'Arial', size: 12, bold: true, color: { argb: 'FFFF0000' } };
        worksheet.getCell(`${logoCol}1`).alignment = { horizontal: 'right', vertical: 'top' };
        
        // 4. 카테고리 설정
        setupCategoriesLG(worksheet);
        
        // 5. 빌딩별 데이터 입력
        selectedBuildings.forEach((building, index) => {
            const col = String.fromCharCode(68 + index); // D, E, F, G...
            fillBuildingDataLG(worksheet, building, col);
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
              `• 재권분석 세부 정보\n` +
              `• 현재 공실 상세\n` +
              `• 특이사항\n\n` +
              `💡 입력한 정보에 따라 비용이 자동 계산됩니다.`);
        
    } catch (error) {
        console.error('엑셀 생성 오류:', error);
        alert('엑셀 파일 생성 중 오류가 발생했습니다.\n' + error.message);
    }
}

// LG 양식 카테고리 설정
function setupCategoriesLG(worksheet) {
    // 5행 - 헤더
    worksheet.mergeCells('B5:C5');
    setCategoryCell(worksheet, 'B5', '위치', 'FF808080', true);
    setCategoryCell(worksheet, 'C5', '반포역', 'FFCCCCCC');
    
    // 6행 - 건물 외관
    worksheet.mergeCells('B6:C6');
    setCategoryCell(worksheet, 'B6', '건물 외관', 'FFFFFFFF');
    
    // 7-8행 - 주소/위치
    setCellLG(worksheet, 'B7', '주 소', false, 'FFF2F2F2');
    setCellLG(worksheet, 'C7', '', false, 'FFF2F2F2');
    setCellLG(worksheet, 'B8', '위 치', false, 'FFF2F2F2');
    setCellLG(worksheet, 'C8', '', false, 'FFF2F2F2');
    
    // 9-14행 - 기본정보
    worksheet.mergeCells('B9:B14');
    setCategoryCell(worksheet, 'B9', '기본\n정보', 'FFFFFFFF');
    
    setCellLG(worksheet, 'C9', '준공일', false, 'FFF2F2F2');
    setCellLG(worksheet, 'C10', '규 모', false, 'FFF2F2F2');
    setCellLG(worksheet, 'C11', '연면적', false, 'FFF2F2F2');
    setCellLG(worksheet, 'C12', '기준층 전용면적', false, 'FFF2F2F2');
    setCellLG(worksheet, 'C13', '전용률', false, 'FFF2F2F2');
    setCellLG(worksheet, 'C14', '대지면적', false, 'FFF2F2F2');
    
    // 15행 - 소유자
    setCellLG(worksheet, 'B15', '소유자 (임대인)', false, 'FFF2F2F2');
    setCellLG(worksheet, 'C15', '', false, 'FFF2F2F2');
    
    // 16-20행 - 재권분석
    worksheet.mergeCells('B16:B20');
    setCategoryCell(worksheet, 'B16', '재권\n분석', 'FFFFFFFF');
    
    setCellLG(worksheet, 'C16', '재권담보 설정여부', false, 'FFF2F2F2');
    setCellLG(worksheet, 'C17', '선순위 담보 총액', false, 'FFF2F2F2');
    setCellLG(worksheet, 'C18', '공시지가 대비 담보율', false, 'FFF2F2F2');
    setCellLG(worksheet, 'C19', '계약공시지가(2024.1월 기준)', false, 'FFF2F2F2');
    setCellLG(worksheet, 'C20', '통지가격 적용', false, 'FFF2F2F2');
    
    // 21행 - 현재 공실
    setCellLG(worksheet, 'B21', '현재 공실', false, 'FFF9D6AE');
    setCellLG(worksheet, 'C21', '', false, 'FFF9D6AE');
    
    // 22-27행 - 채권
    worksheet.mergeCells('B22:B27');
    setCategoryCell(worksheet, 'B22', '채권', 'FFD9ECF2');
    
    setCellLG(worksheet, 'C22', '수요자', false, 'FFF2F2F2');
    setCellLG(worksheet, 'C23', '계약기간', false, 'FFF2F2F2');
    setCellLG(worksheet, 'C24', '임중가능 시기', false, 'FFF2F2F2');
    setCellLG(worksheet, 'C25', '제안 층', false, 'FFF2F2F2');
    setCellLG(worksheet, 'C26', '전용면적', false, 'FFF2F2F2');
    setCellLG(worksheet, 'C27', '임대면적', false, 'FFF2F2F2');
    
    // 28-29행 - 비용감면
    worksheet.mergeCells('B28:B29');
    setCategoryCell(worksheet, 'B28', '비용감면', 'FFFBCF3A');
    
    setCellLG(worksheet, 'C28', '관리비 내역', false, 'FFF2F2F2');
    setCellLG(worksheet, 'C29', '렌트프리(개월)', false, 'FFF2F2F2');
    
    // 30-31행 - 주차현황
    worksheet.mergeCells('B30:B31');
    setCategoryCell(worksheet, 'B30', '주차현황', 'FFFFFFFF');
    
    setCellLG(worksheet, 'C30', '무료주차 제공대수', false, 'FFF2F2F2');
    setCellLG(worksheet, 'C31', '유료주차(VAT별도)', false, 'FFF2F2F2');
    
    // 32행 - 평면도
    worksheet.mergeCells('B32:C32');
    setCategoryCell(worksheet, 'B32', '평면도', 'FFFFFFFF');
    
    // 33행 - 특이사항
    worksheet.mergeCells('B33:C33');
    setCategoryCell(worksheet, 'B33', '특이사항', 'FFFFFFFF');
}

// 빌딩 데이터 입력
function fillBuildingDataLG(worksheet, building, col) {
    // 5행 - 빌딩별 위치
    setCellLG(worksheet, `${col}5`, '가산디지털단지역', false, 'FFFF9900');
    
    // 6행 - 건물 외관 이미지
    setCellLG(worksheet, `${col}6`, '', false, 'FFF5F5F5');
    
    // 7행 - 주소
    const addressCell = worksheet.getCell(`${col}7`);
    addressCell.value = `${building.addressJibun || ''}\n${building.address || ''}`;
    addressCell.font = { name: 'Arial', size: 8 };
    addressCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    setBordersLG(addressCell);
    
    // 8행 - 위치
    setCellLG(worksheet, `${col}8`, building.station || '');
    
    // 9-14행 - 기본정보
    setCellLG(worksheet, `${col}9`, building.completionYear || '');
    setCellLG(worksheet, `${col}10`, building.floors || '');
    setCellLG(worksheet, `${col}11`, building.grossFloorAreaPy ? `${building.grossFloorAreaPy} 평` : '');
    setCellLG(worksheet, `${col}12`, building.baseFloorAreaDedicatedPy ? `${building.baseFloorAreaDedicatedPy} 평` : '');
    setCellLG(worksheet, `${col}13`, building.dedicatedRate ? `${building.dedicatedRate}%` : '');
    setCellLG(worksheet, `${col}14`, building.landAreaPy ? 
        `${building.landAreaPy} 평\n(${building.landArea || 0} m²)` : '', true);
    
    // 15행 - 소유자
    setCellLG(worksheet, `${col}15`, '에크자산개발주식회사');
    
    // 16-20행 - 재권분석
    setCellLG(worksheet, `${col}16`, '전세권 설정 가능', false, 'FFFFCC00');
    setCellLG(worksheet, `${col}17`, '-');
    setCellLG(worksheet, `${col}18`, '0.00%', false, 'FFFF0000');
    setNumericCellLG(worksheet, `${col}19`, 5995000, '₩#,##0/m²');
    setCellLG(worksheet, `${col}20`, '104,363,358,000');
    
    // 21행 - 현재 공실
    setCellLG(worksheet, `${col}21`, '4층        217평        467평', false, 'FFF9D6AE');
    
    // 22-27행 - 채권
    setCellLG(worksheet, `${col}22`, 'LG CNS');
    setCellLG(worksheet, `${col}23`, '2025.7~2027.6 (12개월)');
    setCellLG(worksheet, `${col}24`, '즉시');
    setCellLG(worksheet, `${col}25`, '4층 일부');
    setCellLG(worksheet, `${col}26`, '217 평');
    setCellLG(worksheet, `${col}27`, '467 평');
    
    // 28-29행 - 비용감면
    setCellLG(worksheet, `${col}28`, '전기료, 수도료 포함 / 청소,시큐리티 별도');
    setCellLG(worksheet, `${col}29`, '2개월');
    
    // 30-31행 - 주차현황
    setCellLG(worksheet, `${col}30`, building.parkingSpace || '');
    setCellLG(worksheet, `${col}31`, building.parkingFee || '');
    
    // 32행 - 평면도
    setCellLG(worksheet, `${col}32`, '', false, 'FFF5F5F5');
    
    // 33행 - 특이사항
    const remarkCell = worksheet.getCell(`${col}33`);
    remarkCell.value = '렌트프리 2개월 (보증금 12개월 적용 조건)\n' +
                      'Rent Free : 1층 제외\n' +
                      '공실 영역 대형 호실만 사무실 사용';
    remarkCell.font = { name: 'Arial', size: 8 };
    remarkCell.alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
    setBordersLG(remarkCell);
    
    // 빌딩명 표시 (5행과 같은 행에 작은 글씨로)
    const nameCell = worksheet.getCell(`${col}4`);
    nameCell.value = building.name;
    nameCell.font = { name: 'Arial', size: 10, bold: true };
    nameCell.alignment = { horizontal: 'center', vertical: 'middle' };
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

function setNumericCellLG(worksheet, address, value, format = '#,##0', bgColor = null) {
    const cell = worksheet.getCell(address);
    cell.value = value;
    cell.numFmt = format;
    if (bgColor) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgColor } };
    }
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
    
    // 데이터 영역 (1-33행)
    for (let row = 1; row <= 33; row++) {
        // A, B, C열
        ['A', 'B', 'C'].forEach(col => {
            const cell = worksheet.getCell(`${col}${row}`);
            if (!cell.border && cell.value !== undefined) {
                setBordersLG(cell);
            }
        });
        
        // 빌딩 데이터 열들
        for (let col = 68; col <= endCol; col++) {
            const colLetter = String.fromCharCode(col);
            const cell = worksheet.getCell(`${colLetter}${row}`);
            if (!cell.border && cell.value !== undefined) {
                setBordersLG(cell);
            }
        }
    }
    
    // 하단 주석
    worksheet.getCell('A35').value = '1) 조세공과금(재산세,화재보험료,관리대행수수료 등)는 별도이며 2층 1/2평당 경우 입맞추어 발꿈 - 렌트프리 적용시 재정의';
    worksheet.getCell('A35').font = { name: 'Arial', size: 8 };
    worksheet.getCell('A36').value = '2) 무료주차는 - 매임대인 - 매임대인(입주기간: 매일 08:00-18:00 평일별 보유 및 매임대인: 33.0591 / Rent Free(임대료 관리비 면제 보증금 있음), 프리렌트 포리그 12-13개월 기준(렌트프리기간차)이트리출은 원칙적 허지 않겠습니다 보리)';
    worksheet.getCell('A36').font = { name: 'Arial', size: 8 };
}
