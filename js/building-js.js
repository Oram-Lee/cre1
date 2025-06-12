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
    
    // 선택 버튼 상태 업데이트
    const isSelected = selectedBuildings.some(b => b.id === building.id);
    selectBtn.textContent = isSelected ? '선택 해제' : '선택';
    selectBtn.className = isSelected ? 'deselect-btn' : 'select-btn';
    selectBtn.onclick = function() {
        toggleBuildingSelection(building);
        showBuildingInfo(building); // 팝업 내용 갱신
    };
    
    // 공실정보 확인
    let vacancySection = '';
    if (buildingMatches) {
        const match = buildingMatches.matches.find(m => m.buildingSystemId === building.id);
        
        if (match && match.vacancyMatches.length > 0) {
            const hasVacancy = match.vacancyMatches.some(vm => vm.hasVacancy);
            
            if (hasVacancy) {
                // 회사별 PDF 옵션 생성
                const pdfOptions = match.vacancyMatches
                    .filter(vm => vm.hasVacancy)
                    .map(vm => {
                        const floors = vm.vacancyFloors.join(', ');
                        return `<option value="${vm.pdfFile}|${vm.buildingName}">${vm.company} - ${floors}</option>`;
                    })
                    .join('');
                
                vacancySection = `
                    <div class="info-row">
                        <span class="info-label">공실정보</span>
                        <span class="info-value">
                            <span class="badge bg-success">공실 있음</span>
                            <div style="margin-top: 10px;">
                                <select id="pdfSelect" class="form-select form-select-sm" 
                                        style="width: 100%; margin-bottom: 10px;">
                                    <option value="">임대안내문 선택</option>
                                    ${pdfOptions}
                                </select>
                                <button class="btn btn-sm btn-primary" 
                                        onclick="openPdfViewer('${building.name}')"
                                        style="width: 100%;">
                                    임대안내문 보기
                                </button>
                            </div>
                        </span>
                    </div>
                `;
            } else {
                vacancySection = `
                    <div class="info-row">
                        <span class="info-label">공실정보</span>
                        <span class="info-value">
                            <span class="badge bg-secondary">공실 없음</span>
                        </span>
                    </div>
                `;
            }
        }
    }
    
    // 모달 내용 설정
    modalBody.innerHTML = `
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
            <span class="info-label">기준층면적</span>
            <span class="info-value">${building.baseFloorAreaPy || '-'}</span>
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
            <span class="info-label">주차비</span>
            <span class="info-value">${building.parkingFee || '-'}</span>
        </div>
        <div class="info-row">
            <span class="info-label">준공연도</span>
            <span class="info-value">${building.completionYear || '-'}</span>
        </div>
        ${vacancySection}
        ${building.url ? `
        <div class="info-row">
            <span class="info-label">상세정보</span>
            <span class="info-value">
                <a href="${building.url}" target="_blank" class="btn btn-sm btn-outline-primary">
                    오피스파인드에서 보기
                </a>
            </span>
        </div>
        ` : ''}
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

// 엑셀 내보내기
function exportToExcel() {
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