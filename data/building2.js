// ===== 전역 변수 =====
let buildingsData = [];
let selectedBuildings = [];

// ===== 빌딩 데이터 로드 =====
async function loadBuildingData() {
    try {
        // GitHub에서 buildings.json 파일 로드
        const response = await fetch('./data/buildings.json');
        if (!response.ok) {
            throw new Error('데이터 파일을 찾을 수 없습니다');
        }
        
        const data = await response.json();
        buildingsData = data.buildings || [];
        
        // 테스트용으로 처음 10개만 표시
        displayBuildingList(buildingsData.slice(0, 10));
        
    } catch (error) {
        console.error('데이터 로드 실패:', error);
        // 테스트용 샘플 데이터
        buildingsData = getSampleData();
        displayBuildingList(buildingsData);
    }
}

// ===== 샘플 데이터 (테스트용) =====
function getSampleData() {
    return [
        {
            id: 1,
            name: "유스페이스1-A동",
            address: "서울시 금천구 가산디지털1로 186",
            addressJibun: "서울시 금천구 가산동 60-18",
            station: "1,7호선 가산디지털단지역 도보 10분",
            floors: "지하 2층 ~ 지상 12층",
            completionYear: "2012년",
            baseFloorAreaPy: 467,
            grossFloorAreaPy: 5266,
            landAreaPy: 1023,
            dedicatedRate: 59.49,
            parkingSpace: "총 217대",
            rentPricePy: "4.5만원",
            managementFeePy: "2층: 4.5만원, 11층: 1.6만원",
            depositPy: "45만원",
            vacancy: "2층, 11층",
            description: "에코지오빌딩추천하여 진행",
            url: "#"
        },
        {
            id: 2,
            name: "페이플러스(1층+11층)",
            address: "서울시 금천구 가산디지털1로 186",
            station: "1,7호선 가산디지털단지역 도보 10분",
            floors: "지하 2층 ~ 지상 15층",
            completionYear: "2007년",
            baseFloorAreaPy: 162,
            grossFloorAreaPy: 3929,
            landAreaPy: 776,
            dedicatedRate: 58.10,
            parkingSpace: "전세권, 근저당권 설정 가능",
            rentPricePy: "2.5만원 (평균)",
            managementFeePy: "4.5만원",
            depositPy: "207만원",
            vacancy: "2층",
            description: "저능출입스, 전세권 근저당권 설정 가능",
            url: "#"
        },
        {
            id: 3,
            name: "하이엠빌딩(입암건물)",
            address: "서울시 금천구 디지털로10길 9",
            station: "1,7호선 가산디지털단지역 도보 10분",
            floors: "지하 2층 ~ 지상 20층",
            completionYear: "2013년",
            baseFloorAreaPy: 227,
            grossFloorAreaPy: 3813,
            landAreaPy: 269,
            dedicatedRate: 47.94,
            parkingSpace: "근처임권 설정가능, 공동 11층",
            rentPricePy: "9층: 2.3만원",
            managementFeePy: "11층: 4.5만원, 9층: 2만원",
            depositPy: "227만원",
            vacancy: "9층",
            description: "신한생명부동산식회사",
            url: "#"
        }
    ];
}

// ===== 빌딩 리스트 표시 =====
function displayBuildingList(buildings) {
    const container = document.getElementById('building-list');
    container.innerHTML = '';
    
    buildings.forEach(building => {
        const item = document.createElement('div');
        item.className = 'building-item';
        
        item.innerHTML = `
            <input type="checkbox" id="building-${building.id}" value="${building.id}" onchange="toggleBuilding(${building.id})">
            <div class="building-info">
                <div class="building-name">${building.name}</div>
                <div class="building-address">${building.address}</div>
            </div>
        `;
        
        container.appendChild(item);
    });
}

// ===== 빌딩 선택/해제 =====
function toggleBuilding(buildingId) {
    const building = buildingsData.find(b => b.id === buildingId);
    if (!building) return;
    
    const index = selectedBuildings.findIndex(b => b.id === buildingId);
    
    if (index > -1) {
        selectedBuildings.splice(index, 1);
    } else {
        if (selectedBuildings.length >= 6) {
            alert('최대 6개까지만 선택할 수 있습니다.');
            document.getElementById(`building-${buildingId}`).checked = false;
            return;
        }
        selectedBuildings.push(building);
    }
    
    updateSelectedCount();
}

// ===== 선택 개수 업데이트 =====
function updateSelectedCount() {
    document.getElementById('selected-count').textContent = selectedBuildings.length;
    
    const generateBtn = document.getElementById('generate-btn');
    const generateOriginalBtn = document.getElementById('generate-original-btn');
    
    if (selectedBuildings.length > 0) {
        generateBtn.disabled = false;
        generateOriginalBtn.disabled = false;
    } else {
        generateBtn.disabled = true;
        generateOriginalBtn.disabled = true;
    }
}

// ===== 현재 날짜 반환 =====
function getCurrentDate() {
    const now = new Date();
    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const day = String(now.getDate()).padStart(2, '0');
    return `${year}.${month}.${day}`;
}

// ===== 새로운 양식으로 엑셀 생성 =====
async function generateExcel() {
    if (selectedBuildings.length === 0) {
        alert('빌딩을 선택해주세요.');
        return;
    }
    
    try {
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('후보지');
        
        // 회사명과 제목 가져오기
        const companyName = document.getElementById('company-name').value || '회사명';
        const reportTitle = document.getElementById('report-title').value || '단기임차 가능 공간';
        
        // 1. 열 너비 설정 (A~H열까지, 빌딩 수에 따라 조정)
        const columnWidths = [3, 14, 23]; // A, B, C열
        for (let i = 0; i < selectedBuildings.length; i++) {
            columnWidths.push(25); // 각 빌딩별 열 너비
        }
        
        worksheet.columns = columnWidths.map(width => ({ width }));
        
        // 2. 행 높이 설정
        const rowHeights = {
            1: 20,   // 빈 행
            2: 30,   // 회사 로고
            3: 25,   // 제목
            4: 20,   // 날짜
            5: 150,  // 빌딩 이미지
            6: 20,   // 제안
            7: 20,   // 빈 행
            8: 25,   // 위치
            9: 40,   // 주소
            // ... 나머지 행들
        };
        
        for (let row = 1; row <= 80; row++) {
            worksheet.getRow(row).height = rowHeights[row] || 18;
        }
        
        // 3. 상단 헤더 영역
        // 회사 로고 영역 (B2:C2 병합)
        worksheet.mergeCells('B2:C2');
        const logoCell = worksheet.getCell('B2');
        logoCell.value = companyName;
        logoCell.font = { name: 'Noto Sans KR', size: 16, bold: true };
        logoCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE7E6E6' } };
        logoCell.alignment = { horizontal: 'center', vertical: 'middle' };
        
        // 제목 (D2부터 빌딩 수만큼 병합)
        const titleEndCol = String.fromCharCode(67 + selectedBuildings.length);
        worksheet.mergeCells(`D2:${titleEndCol}2`);
        const titleCell = worksheet.getCell('D2');
        titleCell.value = `[${companyName} ${reportTitle}]`;
        titleCell.font = { name: 'Noto Sans KR', size: 14, bold: true };
        titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
        
        // 작성일자
        worksheet.mergeCells(`D3:${titleEndCol}3`);
        const dateCell = worksheet.getCell('D3');
        dateCell.value = `- 작성기간: ${getCurrentDate()} (작성 기준) -`;
        dateCell.font = { name: 'Noto Sans KR', size: 10 };
        dateCell.alignment = { horizontal: 'center', vertical: 'middle' };
        
        // 4. 빌딩별 데이터 입력
        selectedBuildings.forEach((building, index) => {
            const col = String.fromCharCode(68 + index); // D, E, F, G...
            
            // 빌딩명 (4행)
            const nameCell = worksheet.getCell(`${col}4`);
            nameCell.value = building.name;
            nameCell.font = { name: 'Noto Sans KR', size: 11, bold: true };
            nameCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE7E6E6' } };
            nameCell.alignment = { horizontal: 'center', vertical: 'middle' };
            setBorders(nameCell);
            
            // 빌딩 이미지 영역 (5행)
            const imgCell = worksheet.getCell(`${col}5`);
            imgCell.value = '빌딩 이미지';
            imgCell.font = { name: 'Noto Sans KR', size: 10, italic: true };
            imgCell.alignment = { horizontal: 'center', vertical: 'middle' };
            imgCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF5F5F5' } };
            setBorders(imgCell);
            
            // 빌딩 상세 정보 입력
            fillBuildingDetails(worksheet, building, col);
        });
        
        // 5. 좌측 카테고리 설정
        setupCategories(worksheet);
        
        // 6. 테두리 설정
        applyBorders(worksheet, selectedBuildings.length);
        
        // 7. 파일 저장
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        saveAs(blob, `CompList_v2_${getCurrentDate().replace(/\./g, '')}.xlsx`);
        
        alert('새로운 양식의 Comp List가 생성되었습니다!');
        
    } catch (error) {
        console.error('엑셀 생성 오류:', error);
        alert('엑셀 파일 생성 중 오류가 발생했습니다.\n' + error.message);
    }
}

// ===== 카테고리 설정 =====
function setupCategories(worksheet) {
    // 카테고리 병합
    worksheet.mergeCells('B6:C6');   // 제안
    worksheet.mergeCells('B8:B9');   // 위치
    worksheet.mergeCells('B10:B18'); // 건물 정보
    worksheet.mergeCells('B19:B20'); // 세부 정보
    worksheet.mergeCells('B21:B23'); // 주차 정보
    worksheet.mergeCells('B25:B31'); // 임대 조건
    worksheet.mergeCells('B32:B39'); // 임대 기준
    worksheet.mergeCells('B41:B44'); // 임대료 계산
    worksheet.mergeCells('B46:B50'); // 비용 정보
    
    // 카테고리 텍스트와 스타일
    const categories = {
        'B6': { text: '제안', bg: 'FFF0F0F0' },
        'B8': { text: '위치', bg: 'FFF0F0F0' },
        'B10': { text: '건물 정보', bg: 'FFF0F0F0' },
        'B19': { text: '세부 정보', bg: 'FFF0F0F0' },
        'B21': { text: '주차 관련', bg: 'FFF0F0F0' },
        'B25': { text: '임대 조건', bg: 'FFF9D6AE' },
        'B32': { text: '임대 기준', bg: 'FFD9ECF2' },
        'B41': { text: '렌트프리 적용', bg: 'FFD9ECF2' },
        'B46': { text: '비용 계산', bg: 'FFFBCF3A' }
    };
    
    Object.entries(categories).forEach(([cell, config]) => {
        const categoryCell = worksheet.getCell(cell);
        categoryCell.value = config.text;
        categoryCell.font = { name: 'Noto Sans KR', size: 10, bold: true };
        categoryCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: config.bg } };
        categoryCell.alignment = { horizontal: 'center', vertical: 'middle' };
        setBorders(categoryCell);
    });
    
    // C열 항목명 설정
    const items = {
        8: '위치',
        9: '주소',
        10: '준공연도',
        11: '규모',
        12: '대지면적',
        13: '연면적',
        14: '기준층 임대면적',
        15: '기준층 전용면적',
        16: '전용률',
        17: '엘레베이터',
        18: '냉난방',
        19: '주차대수',
        20: '주차비',
        21: '보증금담보',
        22: '계약기간',
        23: '현황',
        25: '층',
        26: '전용 (평)',
        27: '임대 (평)',
        28: '보증금',
        29: '임대료',
        30: '관리비',
        31: '평당단가',
        32: '보증금 (평당)',
        33: '임대료 (평당)',
        34: '관리비 (평당)',
        35: '합계 (평당)',
        36: '총 보증금',
        37: '월 임대료',
        38: '월 관리비',
        39: '월 합계',
        41: '렌트프리',
        42: '실질임대료',
        43: '실질관리비',
        44: '실질합계'
    };
    
    Object.entries(items).forEach(([row, text]) => {
        const cell = worksheet.getCell(`C${row}`);
        cell.value = text;
        cell.font = { name: 'Noto Sans KR', size: 9 };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        setBorders(cell);
    });
}

// ===== 빌딩 상세 정보 입력 =====
function fillBuildingDetails(worksheet, building, col) {
    // 제안 (6행)
    setCell(worksheet, `${col}6`, building.description || '');
    
    // 위치 정보
    setCell(worksheet, `${col}8`, building.station || '');
    setCell(worksheet, `${col}9`, building.address || '');
    
    // 건물 정보
    setCell(worksheet, `${col}10`, building.completionYear || '');
    setCell(worksheet, `${col}11`, building.floors || '');
    setCell(worksheet, `${col}12`, building.landAreaPy ? `${building.landAreaPy}평` : '');
    setCell(worksheet, `${col}13`, building.grossFloorAreaPy ? `${building.grossFloorAreaPy}평` : '');
    setCell(worksheet, `${col}14`, building.baseFloorAreaPy ? `${building.baseFloorAreaPy}평` : '');
    setCell(worksheet, `${col}15`, building.baseFloorAreaDedicatedPy ? `${building.baseFloorAreaDedicatedPy}평` : '');
    setCell(worksheet, `${col}16`, building.dedicatedRate ? `${building.dedicatedRate}%` : '');
    setCell(worksheet, `${col}17`, building.elevator || '');
    setCell(worksheet, `${col}18`, building.hvac || '');
    
    // 주차 정보
    setCell(worksheet, `${col}19`, building.parkingSpace || '');
    setCell(worksheet, `${col}20`, building.parkingFee || '');
    setCell(worksheet, `${col}21`, '가능');
    setCell(worksheet, `${col}22`, '5년');
    setCell(worksheet, `${col}23`, building.vacancy || '공실 확인 필요');
    
    // 임대 조건 (예시 데이터)
    setCell(worksheet, `${col}25`, '전층');
    setNumericCell(worksheet, `${col}26`, 100);
    setNumericCell(worksheet, `${col}27`, 150);
    
    // 평당 가격에서 숫자만 추출
    const rentPrice = parseFloat(building.rentPricePy?.replace(/[^0-9.]/g, '')) || 0;
    const mgmtFee = parseFloat(building.managementFeePy?.replace(/[^0-9.]/g, '')) || 0;
    const deposit = parseFloat(building.depositPy?.replace(/[^0-9.]/g, '')) || 0;
    
    // 임대 기준
    setNumericCell(worksheet, `${col}32`, deposit, '₩#,##0');
    setNumericCell(worksheet, `${col}33`, rentPrice, '₩#,##0');
    setNumericCell(worksheet, `${col}34`, mgmtFee, '₩#,##0');
    
    // 합계 계산 (수식)
    worksheet.getCell(`${col}35`).value = { formula: `${col}33+${col}34` };
    worksheet.getCell(`${col}35`).numFmt = '₩#,##0';
    
    // 총액 계산 (임대면적 기준)
    worksheet.getCell(`${col}36`).value = { formula: `${col}32*${col}27` };
    worksheet.getCell(`${col}37`).value = { formula: `${col}33*${col}27` };
    worksheet.getCell(`${col}38`).value = { formula: `${col}34*${col}27` };
    worksheet.getCell(`${col}39`).value = { formula: `${col}37+${col}38` };
    
    // 숫자 포맷 적용
    [36, 37, 38, 39].forEach(row => {
        worksheet.getCell(`${col}${row}`).numFmt = '₩#,##0';
        applyDataCellStyle(worksheet.getCell(`${col}${row}`));
    });
    
    // 렌트프리 적용 (예시: 1개월)
    setNumericCell(worksheet, `${col}41`, 1);
    
    // 실질 임대료 계산
    worksheet.getCell(`${col}42`).value = { formula: `${col}33-((${col}33*${col}41)/12)` };
    worksheet.getCell(`${col}43`).value = { formula: `${col}34` };
    worksheet.getCell(`${col}44`).value = { formula: `${col}42+${col}43` };
    
    [42, 43, 44].forEach(row => {
        worksheet.getCell(`${col}${row}`).numFmt = '₩#,##0';
        applyDataCellStyle(worksheet.getCell(`${col}${row}`));
    });
    
    // 평면도 영역 (60행)
    const floorPlanCell = worksheet.getCell(`${col}60`);
    floorPlanCell.value = '평면도';
    floorPlanCell.font = { name: 'Noto Sans KR', size: 10, italic: true };
    floorPlanCell.alignment = { horizontal: 'center', vertical: 'middle' };
    floorPlanCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF5F5F5' } };
    worksheet.getRow(60).height = 150;
    setBorders(floorPlanCell);
    
    // 특이사항 (70행)
    worksheet.mergeCells('B70:C70');
    const remarkLabel = worksheet.getCell('B70');
    remarkLabel.value = '특이사항';
    remarkLabel.font = { name: 'Noto Sans KR', size: 10, bold: true };
    remarkLabel.alignment = { horizontal: 'center', vertical: 'middle' };
    setBorders(remarkLabel);
    
    const remarkCell = worksheet.getCell(`${col}70`);
    remarkCell.value = building.remarks || '';
    remarkCell.font = { name: 'Noto Sans KR', size: 9 };
    remarkCell.alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
    worksheet.getRow(70).height = 40;
    setBorders(remarkCell);
}

// ===== 셀 설정 헬퍼 함수 =====
function setCell(worksheet, address, value) {
    const cell = worksheet.getCell(address);
    cell.value = value;
    applyDataCellStyle(cell);
}

function setNumericCell(worksheet, address, value, format = '#,##0') {
    const cell = worksheet.getCell(address);
    cell.value = value;
    cell.numFmt = format;
    applyDataCellStyle(cell);
}

function applyDataCellStyle(cell) {
    cell.font = { name: 'Noto Sans KR', size: 9 };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    setBorders(cell);
}

function setBorders(cell) {
    cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };
}

// ===== 전체 테두리 적용 =====
function applyBorders(worksheet, buildingCount) {
    const endCol = String.fromCharCode(67 + buildingCount);
    
    // 모든 데이터 영역에 테두리 적용
    for (let row = 2; row <= 70; row++) {
        for (let col = 66; col <= 67 + buildingCount; col++) { // B부터 끝 열까지
            const cell = worksheet.getCell(row, col - 65);
            if (!cell.border) {
                setBorders(cell);
            }
        }
    }
}

// ===== 기존 양식으로 생성 (비교용) =====
function generateExcelOriginal() {
    alert('기존 양식으로 생성하려면 기존 index.html을 사용하세요.');
}