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
    
    if (selectedBuildings.length > 1) {
        alert('LG 양식은 한 번에 1개 빌딩만 생성 가능합니다.\n여러 빌딩을 선택하셨다면 각각 생성해주세요.');
        return;
    }
    
    const building = selectedBuildings[0];
    
    try {
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('후보지');
        
        // 회사명과 제목 가져오기
        const companyName = document.getElementById('company-name').value || 'LG CNS';
        const reportTitle = document.getElementById('report-title').value || '구로&가산디지털단지/반포역 인근 단기임차 가능 공간';
        
        // 1. 열 너비 설정
        worksheet.columns = [
            { width: 3 },      // A열
            { width: 10 },     // B열
            { width: 20 },     // C열
            { width: 3 },      // D열
            { width: 30 },     // E열
            { width: 3 },      // F열
            { width: 30 }      // G열
        ];
        
        // 2. 행 높이 설정
        worksheet.getRow(1).height = 30;   // 제목
        worksheet.getRow(2).height = 20;   // 규모
        worksheet.getRow(3).height = 20;   // 계약기간
        worksheet.getRow(4).height = 20;   // 위치
        worksheet.getRow(5).height = 15;   // 빈 행
        worksheet.getRow(6).height = 25;   // 위치 헤더
        worksheet.getRow(7).height = 25;   // 제안 헤더
        worksheet.getRow(8).height = 15;   // 빈 행
        
        // 건물 외관 영역
        for (let i = 9; i <= 17; i++) {
            worksheet.getRow(i).height = 20;
        }
        
        // 나머지 행들 기본 높이
        for (let i = 18; i <= 85; i++) {
            worksheet.getRow(i).height = 18;
        }
        
        // 특별한 행 높이
        worksheet.getRow(68).height = 120; // 평면도
        worksheet.getRow(73).height = 60;  // 특이사항
        
        // 3. 상단 헤더 영역 설정
        // 제목 (1행)
        worksheet.mergeCells('A1:G1');
        const titleCell = worksheet.getCell('A1');
        titleCell.value = `[${companyName} ${reportTitle}]`;
        titleCell.font = { name: 'Arial', size: 14, bold: true };
        titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
        titleCell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            right: { style: 'thin' },
            bottom: { style: 'thin' }
        };
        
        // 규모 (2행)
        worksheet.mergeCells('A2:G2');
        const scaleCell = worksheet.getCell('A2');
        scaleCell.value = `- 규모: 전용 ${building.baseFloorAreaDedicatedPy || 200}PY 이상`;
        scaleCell.font = { name: 'Arial', size: 10 };
        scaleCell.alignment = { horizontal: 'center', vertical: 'middle' };
        
        // 계약기간 (3행)
        worksheet.mergeCells('A3:G3');
        const periodCell = worksheet.getCell('A3');
        const endDate = new Date();
        endDate.setFullYear(endDate.getFullYear() + 1);
        periodCell.value = `- 계약기간: ${getCurrentDateLG()}~${endDate.getFullYear()}.${String(endDate.getMonth() + 1).padStart(2, '0')}.${String(endDate.getDate()).padStart(2, '0')} (12개월 간) -`;
        periodCell.font = { name: 'Arial', size: 10 };
        periodCell.alignment = { horizontal: 'center', vertical: 'middle' };
        
        // 위치 (4행)
        worksheet.mergeCells('A4:G4');
        const locationDescCell = worksheet.getCell('A4');
        locationDescCell.value = '- 위치: 구로&가산디지털단지역 인근, 반포역 인근 -';
        locationDescCell.font = { name: 'Arial', size: 10 };
        locationDescCell.alignment = { horizontal: 'center', vertical: 'middle' };
        
        // 4. 카테고리 및 데이터 설정
        setupStructureAndDataLG(worksheet, building);
        
        // 5. 테두리 설정
        applyBordersLG(worksheet);
        
        // 6. 파일 저장
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        saveAs(blob, `CompList_LG_${building.name}_${getCurrentDateLG().replace(/\./g, '')}.xlsx`);
        
        alert(`✅ LG용 Comp List 생성 완료!\n\n` +
              `📊 ${building.name}의 상세 정보가 입력되었습니다.\n\n` +
              `📝 추가 입력 필요 항목:\n` +
              `• 건물 외관 이미지\n` +
              `• 평면도 이미지\n` +
              `• 재권분석 세부 정보\n` +
              `• 실질 임대가준 정보\n` +
              `• 특이사항\n\n` +
              `💡 입력한 정보에 따라 비용이 자동 계산됩니다.`);
        
    } catch (error) {
        console.error('엑셀 생성 오류:', error);
        alert('엑셀 파일 생성 중 오류가 발생했습니다.\n' + error.message);
    }
}

// LG 양식 구조 및 데이터 설정
function setupStructureAndDataLG(worksheet, building) {
    // 6행 - 위치
    worksheet.mergeCells('B6:C6');
    setCategoryCell(worksheet, 'B6', '위치', 'FF808080', true);
    setCategoryCell(worksheet, 'E6', '반포역', 'FFCCCCCC');
    
    // 7행 - 제안
    worksheet.mergeCells('B7:C7');
    setCategoryCell(worksheet, 'B7', '제안', 'FF333333', true);
    setCellLG(worksheet, 'E7', building.name || '유스페이스1-A동', false);
    
    // 9-17행 - 건물 외관
    worksheet.mergeCells('B9:C17');
    setCategoryCell(worksheet, 'B9', '건물 외관', 'FFFFFFFF');
    worksheet.mergeCells('E9:E17');
    setCellLG(worksheet, 'E9', '', false, 'FFF5F5F5');
    
    // 18행 - 주소
    setCellLG(worksheet, 'B18', '주 소', false, 'FFF2F2F2');
    setCellLG(worksheet, 'C18', '', false, 'FFF2F2F2');
    setCellLG(worksheet, 'E18', `${building.addressJibun || '성남시 분당구 대왕판교로 660'}`);
    
    // 19행 - 위치
    setCellLG(worksheet, 'B19', '위 치', false, 'FFF2F2F2');
    setCellLG(worksheet, 'C19', '', false, 'FFF2F2F2');
    setCellLG(worksheet, 'E19', building.station || '신분당선 경강선 판교역 버스 10분');
    
    // 20-25행 - 기본정보
    worksheet.mergeCells('B20:B25');
    setCategoryCell(worksheet, 'B20', '기본\n정보', 'FFFFFFFF');
    
    setCellLG(worksheet, 'C20', '준공일', false, 'FFF2F2F2');
    setCellLG(worksheet, 'E20', building.completionYear || '2012년');
    
    setCellLG(worksheet, 'C21', '규 모', false, 'FFF2F2F2');
    setCellLG(worksheet, 'E21', building.floors || '12F / B5');
    
    setCellLG(worksheet, 'C22', '연면적', false, 'FFF2F2F2');
    setCellLG(worksheet, 'E22', building.grossFloorAreaPy ? `${building.grossFloorAreaPy} 평` : '41,281 평');
    
    setCellLG(worksheet, 'C23', '기준층 전용면적', false, 'FFF2F2F2');
    setCellLG(worksheet, 'E23', building.baseFloorAreaDedicatedPy ? `${building.baseFloorAreaDedicatedPy} 평` : '1,004 평');
    
    setCellLG(worksheet, 'C24', '전용률', false, 'FFF2F2F2');
    setCellLG(worksheet, 'E24', building.dedicatedRate ? `${building.dedicatedRate}%` : '46.39%');
    
    setCellLG(worksheet, 'C25', '대지면적', false, 'FFF2F2F2');
    const landAreaText = building.landAreaPy ? 
        `${building.landAreaPy} 평        (${building.landArea || 17408.4} m²)` : 
        '5,266 평        (17,408.4 m²)';
    setCellLG(worksheet, 'E25', landAreaText);
    
    // 26행 - 소유자
    setCellLG(worksheet, 'B26', '소유자 (임대인)', false, 'FFF2F2F2');
    setCellLG(worksheet, 'C26', '', false, 'FFF2F2F2');
    setCellLG(worksheet, 'E26', '에크자산개발주식회사');
    
    // 27-31행 - 재권분석
    worksheet.mergeCells('B27:B31');
    setCategoryCell(worksheet, 'B27', '재권\n분석', 'FFFFFFFF');
    
    setCellLG(worksheet, 'C27', '재권담보 설정여부', false, 'FFF2F2F2');
    setCellLG(worksheet, 'E27', '전세권 설정 가능', false, 'FFFFCC00');
    
    setCellLG(worksheet, 'C28', '선순위 담보 총액', false, 'FFF2F2F2');
    setCellLG(worksheet, 'E28', '-');
    
    setCellLG(worksheet, 'C29', '공시지가 대비 담보율', false, 'FFF2F2F2');
    setCellLG(worksheet, 'E29', '0.00%', false, 'FFFF0000');
    
    setCellLG(worksheet, 'C30', '계약공시지가(23년 1월 기준)', false, 'FFF2F2F2');
    setNumericCellLG(worksheet, 'E30', 5995000, '₩#,##0/㎡');
    
    setCellLG(worksheet, 'C31', '통지가격 적용', false, 'FFF2F2F2');
    setCellLG(worksheet, 'E31', '104,363,358,000');
    
    // 32-35행 - 현재 공실
    worksheet.mergeCells('B32:B35');
    setCategoryCell(worksheet, 'B32', '', 'FFF9D6AE');
    setCellLG(worksheet, 'C32', '현재 공실', false, 'FFF9D6AE');
    
    worksheet.mergeCells('E32:G32');
    setCellLG(worksheet, 'E32', '', false, 'FFF9D6AE');
    
    setCellLG(worksheet, 'C33', '', false, 'FFF9D6AE');
    setCellLG(worksheet, 'E33', '층', false, 'FFF9D6AE');
    setCellLG(worksheet, 'F33', '전용', false, 'FFF9D6AE');
    setCellLG(worksheet, 'G33', '임대', false, 'FFF9D6AE');
    
    setCellLG(worksheet, 'C34', '', false, 'FFF9D6AE');
    setCellLG(worksheet, 'E34', '4층', false, 'FFF9D6AE');
    setCellLG(worksheet, 'F34', '217평', false, 'FFF9D6AE');
    setCellLG(worksheet, 'G34', '467평', false, 'FFF9D6AE');
    
    setCellLG(worksheet, 'C35', '', false, 'FFF9D6AE');
    setCellLG(worksheet, 'E35', '', false, 'FFF9D6AE');
    setCellLG(worksheet, 'F35', '', false, 'FFF9D6AE');
    setCellLG(worksheet, 'G35', '', false, 'FFF9D6AE');
    
    // 36-39행 - 빈 공실 영역
    for (let row = 36; row <= 39; row++) {
        setCellLG(worksheet, `B${row}`, '', false);
        setCellLG(worksheet, `C${row}`, '현재 공실', false);
        setCellLG(worksheet, `E${row}`, '', false);
    }
    
    setCellLG(worksheet, 'C39', '', false);
    setCellLG(worksheet, 'E39', '소계', false);
    setCellLG(worksheet, 'F39', '217평', false);
    setCellLG(worksheet, 'G39', '467평', false);
    
    // 40-44행 - 제안
    setCellLG(worksheet, 'C40', '계약기간', false, 'FFF2F2F2');
    setCellLG(worksheet, 'E40', '2025.7~2027.6 (12개월)');
    
    setCellLG(worksheet, 'C41', '임중가능 시기', false, 'FFF2F2F2');
    setCellLG(worksheet, 'E41', '즉시');
    
    worksheet.mergeCells('B42:B44');
    setCategoryCell(worksheet, 'B42', '제안', 'FFFFFFFF');
    
    setCellLG(worksheet, 'C42', '제안 층', false, 'FFF2F2F2');
    setCellLG(worksheet, 'E42', '4층 일부');
    
    setCellLG(worksheet, 'C43', '전용면적', false, 'FFF2F2F2');
    setCellLG(worksheet, 'E43', '217 평', false, 'FFFF0000');
    
    setCellLG(worksheet, 'C44', '임대면적', false, 'FFF2F2F2');
    setCellLG(worksheet, 'E44', '467 평');
    
    // 45-47행 - 기준층 임대가준
    worksheet.mergeCells('B45:B47');
    setCategoryCell(worksheet, 'B45', '기준층\n임대가준', 'FFFFFFFF');
    
    setCellLG(worksheet, 'C45', '보증금', false, 'FFF2F2F2');
    setCellLG(worksheet, 'E45', '@1,048,752');
    
    setCellLG(worksheet, 'C46', '임대료', false, 'FFF2F2F2');
    setCellLG(worksheet, 'E46', '@104,875');
    
    setCellLG(worksheet, 'C47', '관리비', false, 'FFF2F2F2');
    setCellLG(worksheet, 'E47', '@6,000+실비별도');
    
    // 48-52행 - 실질 임대가준
    worksheet.mergeCells('B48:B52');
    setCategoryCell(worksheet, 'B48', '실질\n임대가준', 'FFFFFFFF');
    
    setCellLG(worksheet, 'C48', '실질 임대료(RF면 반영)¹⁾', false, 'FFF2F2F2');
    setCellLG(worksheet, 'E48', '@96,135');
    
    setCellLG(worksheet, 'C49', '연간 무상임대 (R.F)', false, 'FFF2F2F2');
    setCellLG(worksheet, 'E49', '1.0개월');
    
    setCellLG(worksheet, 'C50', '보증금', false, 'FFF2F2F2');
    setNumericCellLG(worksheet, 'E50', 490207660, '₩#,##0 원');
    
    setCellLG(worksheet, 'C51', '월 임대료', false, 'FFF2F2F2');
    setNumericCellLG(worksheet, 'E51', 49020673, '₩#,##0 원');
    
    setCellLG(worksheet, 'C52', '월 관리비', false, 'FFF2F2F2');
    setNumericCellLG(worksheet, 'E52', 2804520, '₩#,##0 원');
    
    // 53-54행 - 비용감면
    worksheet.mergeCells('B53:B54');
    setCategoryCell(worksheet, 'B53', '비용감면', 'FFFBCF3A');
    
    setCellLG(worksheet, 'C53', '관리비 내역', false, 'FFF2F2F2');
    setCellLG(worksheet, 'E53', '실비 관리비: 전기세, 수도세 별도 부과\n(예상 수광비 약 4천원대)', true, 'FFFFCC66');
    
    setCellLG(worksheet, 'C54', '렌트프리', false, 'FFF2F2F2');
    setCellLG(worksheet, 'E54', '49,020,673 원', false, 'FFFF0000');
    
    // 55-56행 - 공사거리
    worksheet.mergeCells('B55:B56');
    setCategoryCell(worksheet, 'B55', '공사거리', 'FFFFFFFF');
    
    setCellLG(worksheet, 'C55', '(21개월 기준) 총 년째 비용³⁾', false, 'FFF2F2F2');
    setNumericCellLG(worksheet, 'E55', 1029434123, '₩#,##0 원');
    
    setCellLG(worksheet, 'C56', '인테리어 기간 (F.O)', false, 'FFF2F2F2');
    setCellLG(worksheet, 'E56', '협의');
    
    // 57-62행 - 주차현황
    setCellLG(worksheet, 'C59', '총 주차대수', false, 'FFF2F2F2');
    setCellLG(worksheet, 'E59', building.parkingSpace || '1023 대');
    
    setCellLG(worksheet, 'C60', '무료주차 조건(임대면적)', false, 'FFF2F2F2');
    setCellLG(worksheet, 'E60', '임대면적 80평당 1대');
    
    worksheet.mergeCells('B61:B62');
    setCategoryCell(worksheet, 'B61', '주차현황', 'FFFFFFFF');
    
    setCellLG(worksheet, 'C61', '무료주차 제공대수', false, 'FFF2F2F2');
    setCellLG(worksheet, 'E61', '5.8 대');
    
    setCellLG(worksheet, 'C62', '유료주차(VAT별도)', false, 'FFF2F2F2');
    setCellLG(worksheet, 'E62', building.parkingFee || '협의');
    
    // 68행 - 평면도
    worksheet.mergeCells('B63:C72');
    setCategoryCell(worksheet, 'B68', '평면도', 'FFFFFFFF');
    worksheet.mergeCells('E63:G72');
    setCellLG(worksheet, 'E68', '', false, 'FFF5F5F5');
    
    // 73-79행 - 기타/특이사항
    worksheet.mergeCells('B73:B79');
    setCategoryCell(worksheet, 'B73', '기타', 'FFFFFFFF');
    
    worksheet.mergeCells('E73:G76');
    const remarkCell = worksheet.getCell('E73');
    remarkCell.value = '- 현재 4층(410~412호) 일부 층시 가능\n' +
                      '- Rent Free 1개월 제공\n' +
                      '- 공시기간 협의 필요\n' +
                      '- 실비 별도 (예상 수광비 실비: 4천원)';
    remarkCell.font = { name: 'Arial', size: 9 };
    remarkCell.alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
    setBordersLG(remarkCell);
    
    setCellLG(worksheet, 'C79', '특이사항', false, 'FFF2F2F2');
    setCellLG(worksheet, 'E79', '', false);
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
function applyBordersLG(worksheet) {
    // 데이터 영역 (1-83행)
    for (let row = 1; row <= 83; row++) {
        // 모든 열에 테두리 적용
        ['A', 'B', 'C', 'D', 'E', 'F', 'G'].forEach(col => {
            const cell = worksheet.getCell(`${col}${row}`);
            if (!cell.border && (cell.value !== undefined || cell.value !== null)) {
                setBordersLG(cell);
            }
        });
    }
}
