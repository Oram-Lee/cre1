// ===== 빌딩 관련 함수들 =====

// 빌딩 하이라이트
function highlightBuilding(buildingId) {
    // 이전 하이라이트 제거
    highlightedBuildingId = buildingId;
    
    // 지도에 표시되어 있을 때만 하이라이트
    if (currentZoomLevel < CLUSTER_MIN_LEVEL) {
        showBuildingLabels();
    }
}

// 하이라이트 제거
function removeHighlight() {
    highlightedBuildingId = null;
    
    if (currentZoomLevel < CLUSTER_MIN_LEVEL) {
        showBuildingLabels();
    }
}

// 빌딩 리스트 표시
function displayBuildingList(buildings) {
    const listContainer = document.getElementById('building-list');
    listContainer.innerHTML = '';
    
    // 현재 표시된 빌딩들 저장
    currentDisplayedBuildings = buildings;
    
    buildings.forEach(building => {
        const item = document.createElement('div');
        item.className = 'building-item';
        
        // 선택된 빌딩인지 확인
        const isSelected = selectedBuildings.some(b => b.id === building.id);
        if (isSelected) {
            item.classList.add('selected');
        }
        
        // 하이라이트된 빌딩인지 확인
        if (building.id === highlightedBuildingId) {
            item.classList.add('highlighted');
        }
        
        item.innerHTML = `
            <div class="building-name">${building.name}</div>
            <div class="building-address">${building.address}</div>
            ${building.buildingType ? `<span class="building-type">${building.buildingType}</span>` : ''}
        `;
        
        // 툴팁 추가
        item.title = '클릭하여 상세정보 보기';
        
        // 클릭 시 하이라이트 및 지도 이동
        item.onclick = function() {
            // 하이라이트
            highlightBuilding(building.id);
            
            // 모든 building-item에서 highlighted 클래스 제거
            document.querySelectorAll('.building-item').forEach(el => {
                el.classList.remove('highlighted');
            });
            
            // 현재 아이템에 highlighted 추가
            item.classList.add('highlighted');
            
            // 해당 빌딩으로 지도 이동
            if (building.lat && building.lng) {
                const position = new kakao.maps.LatLng(building.lat, building.lng);
                map.setCenter(position);
                
                // 현재 화면에 빌딩이 보이지 않으면 줌 레벨 조정
                const bounds = map.getBounds();
                if (!bounds.contain(position)) {
                    map.setLevel(3);
                }
            }
            
            // 팝업 표시
            showBuildingInfo(building);
        };
        
        // 마우스 벗어날 때 하이라이트 제거
        item.onmouseleave = function() {
            removeHighlight();
            item.classList.remove('highlighted');
        };
        
        listContainer.appendChild(item);
    });
    
    // 통계 업데이트
    updateStats();
}

// 빌딩 정보 팝업 표시
function showBuildingInfo(building) {
    const modal = document.getElementById('buildingModal');
    const modalTitle = document.getElementById('modalTitle');
    const modalBody = document.getElementById('modalBody');
    const selectBtn = document.getElementById('selectBuildingBtn');
    
    // 제목 설정
    modalTitle.textContent = building.name;
    
    // 선택 버튼 상태 업데이트는 삭제 (하단으로 이동했으므로)
    const isSelected = selectedBuildings.some(b => b.id === building.id);
    
    // 공실정보 확인
    let vacancySection = '';
    if (buildingMatches) {
        const match = buildingMatches.matches.find(m => m.buildingSystemId === building.id);
        
        if (match && match.vacancyMatches.length > 0) {
            const hasVacancy = match.vacancyMatches.some(vm => vm.hasVacancy);
            
            if (hasVacancy) {
                // 회사별 PDF 옵션 생성 (중복 제거)
                const companies = new Set();
                const pdfOptions = match.vacancyMatches
                    .filter(vm => {
                        if (vm.hasVacancy && !companies.has(vm.company)) {
                            companies.add(vm.company);
                            return true;
                        }
                        return false;
                    })
                    .map(vm => {
                        return `<option value="${vm.pdfFile}|${vm.buildingName}">${vm.company}</option>`;
                    })
                    .join('');
                
                vacancySection = `
                    <div class="info-row">
                        <span class="info-label">관련임대안내문</span>
                        <span class="info-value">
                            <span class="badge bg-success">관련안내문있음</span>
                            <div style="margin-top: 10px;">
                                <select id="pdfSelect" class="form-select form-select-sm" 
                                        style="width: 100%; margin-bottom: 10px;">
                                    <option value="">임대안내문 선택</option>
                                    ${pdfOptions}
                                </select>
                                <button class="btn btn-sm btn-primary" 
                                        onclick="openPdfViewer('${building.name}')"
                                        style="width: 100%;">
                                    임대안내문 열기
                                </button>
                            </div>
                        </span>
                    </div>
                `;
            } else {
                vacancySection = `
                    <div class="info-row">
                        <span class="info-label">관련임대안내문</span>
                        <span class="info-value">
                            <span class="badge bg-secondary">관련안내문없음</span>
                        </span>
                    </div>
                `;
            }
        }
    }
    
    // 모달 내용 설정 - 2열 레이아웃
    modalBody.innerHTML = `
        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 20px;">
            <!-- 왼쪽 열 -->
            <div>
                <div class="info-row">
                    <span class="info-label">주소</span>
                    <span class="info-value">${building.address}</span>
                </div>
                ${building.addressJibun ? `
                <div class="info-row">
                    <span class="info-label">지번주소</span>
                    <span class="info-value">${building.addressJibun}</span>
                </div>
                ` : ''}
                <div class="info-row">
                    <span class="info-label">인근역</span>
                    <span class="info-value">${building.station || '-'}</span>
                </div>
                <div class="info-row">
                    <span class="info-label">규모</span>
                    <span class="info-value">${building.floors || '-'}</span>
                </div>
                <div class="info-row">
                    <span class="info-label">빌딩유형</span>
                    <span class="info-value">${building.buildingType || '-'}</span>
                </div>
                <div class="info-row">
                    <span class="info-label">준공연도</span>
                    <span class="info-value">${building.completionYear || '-'}</span>
                </div>
            </div>
            
            <!-- 오른쪽 열 -->
            <div>
                <div class="info-row">
                    <span class="info-label">기준층전용면적(평)</span>
                    <span class="info-value">${building.baseFloorAreaDedicatedPy || '-'}</span>
                </div>
                <div class="info-row">
                    <span class="info-label">기준층임대면적(평)</span>
                    <span class="info-value">${building.baseFloorAreaPy || '-'}</span>
                </div>
                <div class="info-row">
                    <span class="info-label">전용율(%)</span>
                    <span class="info-value">${building.dedicatedRate || '-'}</span>
                </div>
                <div class="info-row">
                    <span class="info-label">임대료(평당)</span>
                    <span class="info-value">${building.rentPricePy || '-'}</span>
                </div>
                <div class="info-row">
                    <span class="info-label">관리비(평당)</span>
                    <span class="info-value">${building.managementFeePy || '-'}</span>
                </div>
                <div class="info-row">
                    <span class="info-label">보증금(평당)</span>
                    <span class="info-value">${building.depositPy || '-'}</span>
                </div>
            </div>
        </div>
        
        <!-- 공실정보와 상세정보 링크를 5:5 비율로 -->
        <div style="margin-top: 20px; display: grid; grid-template-columns: 1fr 1fr; gap: 20px;">
            <div>
                ${vacancySection || '<div class="info-row"><span class="info-label">관련임대안내문</span><span class="info-value"><span class="badge bg-secondary">정보 없음</span></span></div>'}
            </div>
            <div>
                ${building.url ? `
                <div class="info-row">
                    <span class="info-label">상세정보</span>
                    <span class="info-value">
                        <button onclick="window.open('${building.url}', '_blank')" class="btn btn-sm btn-primary" style="width: 100%;">
                            원문 보기
                        </button>
                    </span>
                </div>
                ` : '<div class="info-row"><span class="info-label">상세정보</span><span class="info-value">-</span></div>'}
            </div>
        </div>
        
        <!-- 선택 버튼을 하단으로 이동 -->
        <div style="margin-top: 20px;">
            <button id="selectBuildingBtn" class="${isSelected ? 'deselect-btn' : 'select-btn'}" 
                    onclick="toggleBuildingSelection(buildingsData.find(b => b.id === ${building.id})); showBuildingInfo(buildingsData.find(b => b.id === ${building.id}));">
                ${isSelected ? '선택 해제' : '선택'}
            </button>
        </div>
    `;
    
    modal.style.display = 'block';
}

// 빌딩 선택/해제
function toggleBuildingSelection(building) {
    const index = selectedBuildings.findIndex(b => b.id === building.id);
    
    if (index > -1) {
        selectedBuildings.splice(index, 1);
    } else {
        selectedBuildings.push(building);
    }
    
    updateCartCount();
    displayBuildingList(currentDisplayedBuildings.length > 0 ? currentDisplayedBuildings : buildingsData);
    
    // 줌 레벨에 따라 적절한 표시 업데이트
    if (currentZoomLevel < CLUSTER_MIN_LEVEL) {
        displayBuildingLabels(currentDisplayedBuildings.length > 0 ? currentDisplayedBuildings : buildingsData);
    }
    
    displaySelectedBuildingsList();
}

// 선택된 빌딩 리스트 표시
function displaySelectedBuildingsList() {
    const container = document.getElementById('selected-buildings-list');
    container.innerHTML = '';
    
    // 선택된 빌딩이 있을 때만 표시
    if (selectedBuildings.length > 0) {
        container.classList.add('has-items');
        
        selectedBuildings.forEach((building, index) => {
            const item = document.createElement('div');
            item.className = 'selected-building-item';
            item.draggable = true;
            item.dataset.index = index;
            
            item.innerHTML = `
                <div class="order-number">${index + 1}</div>
                <div class="building-info">${building.name}</div>
                <div class="order-controls">
                    <button onclick="moveBuilding(${index}, -1)">▲</button>
                    <button onclick="moveBuilding(${index}, 1)">▼</button>
                    <button onclick="removeSelectedBuilding(${index})">X</button>
                </div>
            `;
            
            // 드래그 이벤트
            item.addEventListener('dragstart', handleDragStart);
            item.addEventListener('dragover', handleDragOver);
            item.addEventListener('drop', handleDrop);
            item.addEventListener('dragend', handleDragEnd);
            
            container.appendChild(item);
        });
    } else {
        container.classList.remove('has-items');
    }
}

// 빌딩 순서 이동
function moveBuilding(index, direction) {
    const newIndex = index + direction;
    
    if (newIndex >= 0 && newIndex < selectedBuildings.length) {
        const temp = selectedBuildings[index];
        selectedBuildings[index] = selectedBuildings[newIndex];
        selectedBuildings[newIndex] = temp;
        
        displaySelectedBuildingsList();
    }
}

// 선택된 빌딩 제거
function removeSelectedBuilding(index) {
    selectedBuildings.splice(index, 1);
    updateCartCount();
    displayBuildingList(currentDisplayedBuildings.length > 0 ? currentDisplayedBuildings : buildingsData);
    
    // 줌 레벨에 따라 적절한 표시 업데이트
    if (currentZoomLevel < CLUSTER_MIN_LEVEL) {
        displayBuildingLabels(currentDisplayedBuildings.length > 0 ? currentDisplayedBuildings : buildingsData);
    }
    
    displaySelectedBuildingsList();
}

// 드래그 앤 드롭 핸들러
let draggedElement = null;

function handleDragStart(e) {
    draggedElement = this;
    this.classList.add('dragging');
    e.dataTransfer.effectAllowed = 'move';
    e.dataTransfer.setData('text/html', this.innerHTML);
}

function handleDragOver(e) {
    if (e.preventDefault) {
        e.preventDefault();
    }
    e.dataTransfer.dropEffect = 'move';
    return false;
}

function handleDrop(e) {
    if (e.stopPropagation) {
        e.stopPropagation();
    }
    
    if (draggedElement !== this) {
        const draggedIndex = parseInt(draggedElement.dataset.index);
        const targetIndex = parseInt(this.dataset.index);
        
        // 배열에서 위치 교환
        const draggedBuilding = selectedBuildings[draggedIndex];
        selectedBuildings.splice(draggedIndex, 1);
        selectedBuildings.splice(targetIndex, 0, draggedBuilding);
        
        displaySelectedBuildingsList();
    }
    
    return false;
}

function handleDragEnd(e) {
    this.classList.remove('dragging');
}

// ===== 템플릿 스타일 정의 =====
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
    
    // B24 셀 (빨간색 안내 문구)
    if (sheet['B24']) {
        sheet['B24'].s = convertToSheetJSStyle({
            font: { name: 'Noto Sans KR', size: 10, bold: true, color: 'FFFF0000' },
            alignment: { horizontal: 'left', vertical: 'center' }
        });
    }
    
    // 숫자 형식 적용
    applyNumberFormats(sheet);
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

// 빌딩 데이터를 템플릿에 입력하는 함수
function fillBuildingDataToTemplate(sheet, building, columnIndex) {
    // D열부터 시작 (D=3, E=4, F=5, G=6, H=7)
    const col = String.fromCharCode(68 + columnIndex); // D, E, F, G, H
    
    // === 빌딩 기본 정보 (수식 없음) ===
    
    // 빌딩명 (행 4)
    setCellValue(sheet, `${col}4`, building.name || '');
    
    // 주소 지번 (행 7)
    setCellValue(sheet, `${col}7`, building.addressJibun || '');
    
    // 도로명 주소 (행 8)
    setCellValue(sheet, `${col}8`, building.address || '');
    
    // 위치 - 지하철역 (행 9)
    setCellValue(sheet, `${col}9`, building.station || '');
    
    // 빌딩 규모 (행 10)
    setCellValue(sheet, `${col}10`, building.floors || '');
    
    // 사용승인일 (행 11)
    setCellValue(sheet, `${col}11`, building.completionYear || '');
    
    // 전용률 (행 12)
    setCellValue(sheet, `${col}12`, building.dedicatedRate || 0, 'n');
    
    // === 면적 정보 ===
    
    // 기준층 임대면적 (m²) (행 13)
    setCellValue(sheet, `${col}13`, building.baseFloorArea || 0, 'n');
    
    // 기준층 임대면적 (평) (행 14)
    setCellValue(sheet, `${col}14`, building.baseFloorAreaPy || 0, 'n');
    
    // 기준층 전용면적 (m²) (행 15)
    setCellValue(sheet, `${col}15`, building.baseFloorAreaDedicated || 0, 'n');
    
    // 기준층 전용면적 (평) (행 16)
    setCellValue(sheet, `${col}16`, building.baseFloorAreaDedicatedPy || 0, 'n');
    
    // === 빌딩 세부현황 ===
    
    // 주차 대수 정보 (행 17)
    setCellValue(sheet, `${col}17`, building.parkingSpace || '');
    
    // 냉난방 방식 (행 18)
    setCellValue(sheet, `${col}18`, building.hvac || '');
    
    // 건물종류 (행 19)
    setCellValue(sheet, `${col}19`, building.buildingUse || '');
    
    // 구조 (행 20)
    setCellValue(sheet, `${col}20`, building.structure || '');
    
    // 엘리베이터 (행 21)
    setCellValue(sheet, `${col}21`, building.elevator || '');
    
    // === 주차 관련 ===
    
    // 주차 운영 (행 22)
    setCellValue(sheet, `${col}22`, building.parkingOperation || '');
    
    // 주차 대수 (행 23) - 17행과 동일
    setCellValue(sheet, `${col}23`, building.parkingSpace || '');
    
    // 주차비 (행 24)
    setCellValue(sheet, `${col}24`, building.parkingFee || '');
    
    // === 임차 제안 (기본값 설정) ===
    
    // 최적 임차 층수 (행 26) - 기본값
    setCellValue(sheet, `${col}26`, '-');
    
    // 입주 가능 시기 (행 27) - 기본값
    setCellValue(sheet, `${col}27`, '-');
    
    // 거래유형 (행 28) - 기본값  
    setCellValue(sheet, `${col}28`, '-');
    
    // 임대면적 (m²) (행 29) - 기본값 0
    setCellValue(sheet, `${col}29`, 0, 'n');
    
    // 임대면적 (평) (행 30) - 기본값 0
    setCellValue(sheet, `${col}30`, 0, 'n');
    
    // 전용면적 (평) (행 31) - 기본값 0
    setCellValue(sheet, `${col}31`, 0, 'n');
    
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
            setCellValue(sheet, `${col}${row}`, 0, 'n');
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

// 셀에 값을 설정하는 헬퍼 함수
function setCellValue(sheet, cellAddress, value, type = 's') {
    if (!sheet[cellAddress]) {
        sheet[cellAddress] = {};
    }
    
    sheet[cellAddress].v = value;
    sheet[cellAddress].t = type; // 's' = string, 'n' = number
    
    // 기존 스타일 유지
    if (sheet[cellAddress].s) {
        // 스타일은 그대로 유지
    }
}

// 현재 날짜 반환
function getCurrentDate() {
    return new Date().toISOString().split('T')[0];
}

// 엑셀 내보내기 - 템플릿 기반으로 수정
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
        const buildingsToExport = selectedBuildings.filter(b => b);
        
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

// 기존의 간단한 엑셀 내보내기 (fallback용)
function exportToExcelBasic() {
    if (selectedBuildings.length === 0) {
        alert('선택된 빌딩이 없습니다.');
        return;
    }
    
    // 선택된 빌딩 데이터를 순서대로 가져오기
    const selectedData = selectedBuildings;
    
    // 워크시트 데이터 생성
    const wsData = [
        ['순번', '빌딩명', '주소', '지하철역', '층수', '건물유형', '기준층 면적', '임대료', '관리비', '주차비', '준공년도', '상세URL']
    ];
    
    selectedData.forEach((building, index) => {
        wsData.push([
            index + 1,
            building.name,
            building.address,
            building.station || '',
            building.floors || '',
            building.buildingType || '',
            building.baseFloorAreaPy ? building.baseFloorAreaPy + '평' : '',
            building.rentPricePy || '',
            building.managementFeePy || '',
            building.parkingFee || '',
            building.completionYear || '',
            building.url || ''
        ]);
    });
    
    // 워크북 생성
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    
    // 컬럼 너비 설정
    const wscols = [
        {wch: 6},   // 순번
        {wch: 25},  // 빌딩명
        {wch: 40},  // 주소
        {wch: 30},  // 지하철역
        {wch: 15},  // 층수
        {wch: 10},  // 건물유형
        {wch: 15},  // 기준층 면적
        {wch: 15},  // 임대료
        {wch: 15},  // 관리비
        {wch: 20},  // 주차비
        {wch: 10},  // 준공년도
        {wch: 40}   // URL
    ];
    ws['!cols'] = wscols;
    
    XLSX.utils.book_append_sheet(wb, ws, "Comp List");
    
    // 파일 다운로드
    XLSX.writeFile(wb, `CompList_${new Date().toISOString().slice(0,10)}.xlsx`);
}

// 구버전 호환용 함수들
function showBuildingPopup(building) {
    showBuildingInfo(building);
}

function toggleBuildingSelectionFromPopup(buildingId) {
    const building = buildingsData.find(b => b.id === buildingId);
    if (building) {
        toggleBuildingSelection(building);
        showBuildingInfo(building);
    }
}