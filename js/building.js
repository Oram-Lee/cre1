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
        const templatePath = `${basePath}/templates/template.xlsx`;  // templates로 수정
        
        console.log('템플릿 경로:', templatePath);
        
        // 템플릿 파일 로드
        const response = await fetch(templatePath);
        if (!response.ok) {
            throw new Error('템플릿 파일을 찾을 수 없습니다.');
        }
        
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, {
            cellStyles: true,
            cellFormulas: true,
            cellDates: true
        });
        
        // '후보지' 시트 찾기 (또는 첫 번째 시트)
        let sheetName = '후보지';
        if (!workbook.Sheets[sheetName]) {
            sheetName = workbook.SheetNames[0];
        }
        const sheet = workbook.Sheets[sheetName];
        
        // 선택된 빌딩 데이터 입력
        selectedBuildings.forEach((building, index) => {
            fillBuildingDataToTemplate(sheet, building, index);
        });
        
        // 엑셀 파일 생성
        const wbout = XLSX.write(workbook, {
            bookType: 'xlsx',
            type: 'array',
            cellFormulas: true,
            cellStyles: true
        });
        
        // 다운로드
        const blob = new Blob([wbout], {
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `CompList_${new Date().toISOString().split('T')[0]}.xlsx`;
        a.click();
        URL.revokeObjectURL(url);
        
        console.log('Comp List 생성 완료');
        
    } catch (error) {
        console.error('엑셀 생성 오류:', error);
        
        // 템플릿이 없을 경우 기존 방식으로 fallback
        if (error.message.includes('템플릿')) {
            if (confirm('템플릿 파일을 찾을 수 없습니다. 기본 형식으로 내보내시겠습니까?')) {
                exportToExcelBasic();
            }
        } else {
            alert('엑셀 파일 생성 중 오류가 발생했습니다.\n' + error.message);
        }
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
    
    // 임대면적 (m²) (행 30) - 기본값 0
    setCellValue(sheet, `${col}30`, 0, 'n');
    
    // 전용면적 (m²) (행 31) - 기본값 0
    setCellValue(sheet, `${col}31`, 0, 'n');
    
    // === 임대 기준 (기본값 설정) ===
    
    // 보증금 평당 (행 32) - 기본값 0
    setCellValue(sheet, `${col}32`, 0, 'n');
    
    // 임대료 평당 (행 33) - 기본값 0
    setCellValue(sheet, `${col}33`, 0, 'n');
    
    // 관리비 평당 (행 34) - 기본값 0
    setCellValue(sheet, `${col}34`, 0, 'n');
    
    // 렌트프리 (개월/년) (행 41) - 기본값 0
    setCellValue(sheet, `${col}41`, 0, 'n');
    
    // === 임차 특이사항 ===
    
    // description을 임차 특이사항으로 사용 (행 52 또는 적절한 위치)
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