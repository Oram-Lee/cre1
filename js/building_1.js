// displayBuildingList 함수 수정 - 컴팩트 디자인 적용
function displayBuildingList(buildings) {
    const listContainer = document.getElementById('building-list');
    
    // 빈 상태 처리
    if (buildings.length === 0) {
        listContainer.innerHTML = `
            <div class="empty-state">
                <div class="empty-state-icon">🏢</div>
                <div class="empty-state-text">검색 결과가 없습니다</div>
            </div>
        `;
        updateTabCounts();
        return;
    }
    
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
        
        // 컴팩트한 레이아웃
        item.innerHTML = `
            <div class="building-content">
                <div class="building-name">${building.name}</div>
                <div class="building-address">${building.address}</div>
            </div>
            <div class="building-badges">
                ${building.buildingType ? `<span class="building-type">${building.buildingType}</span>` : ''}
                ${isSelected ? '<span class="building-selected-badge">선택됨</span>' : ''}
            </div>
        `;
        
        // 툴팁 추가
        item.title = '클릭하여 상세정보 보기';
        
        // 클릭 이벤트
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
    updateTabCounts();
}

// displaySelectedBuildingsList 함수 수정
function displaySelectedBuildingsList() {
    const container = document.getElementById('selected-buildings-list');
    
    // 빈 상태 처리
    if (selectedBuildings.length === 0) {
        container.innerHTML = `
            <div class="empty-state">
                <div class="empty-state-icon">📋</div>
                <div class="empty-state-text">선택된 빌딩이 없습니다</div>
            </div>
        `;
        updateTabCounts();
        return;
    }
    
    container.innerHTML = '';
    
    selectedBuildings.forEach((building, index) => {
        const item = document.createElement('div');
        item.className = 'selected-building-item';
        item.draggable = true;
        item.dataset.index = index;
        
        item.innerHTML = `
            <div class="order-number">${index + 1}</div>
            <div class="building-info">${building.name}</div>
            <div class="order-controls">
                <button onclick="moveBuilding(${index}, -1)" title="위로">▲</button>
                <button onclick="moveBuilding(${index}, 1)" title="아래로">▼</button>
                <button onclick="removeSelectedBuilding(${index})" title="제거">✕</button>
            </div>
        `;
        
        // 드래그 이벤트
        item.addEventListener('dragstart', handleDragStart);
        item.addEventListener('dragover', handleDragOver);
        item.addEventListener('drop', handleDrop);
        item.addEventListener('dragend', handleDragEnd);
        
        container.appendChild(item);
    });
    
    updateTabCounts();
}

// 탭 카운트 업데이트 함수 추가
function updateTabCounts() {
    // 빌딩 목록 카운트
    const listCount = currentDisplayedBuildings.length;
    document.getElementById('list-count').textContent = listCount;
    
    // 선택 목록 카운트
    const selectedCount = selectedBuildings.length;
    document.getElementById('selected-count').textContent = selectedCount;
}

// toggleBuildingSelection 함수 수정
function toggleBuildingSelection(building) {
    const index = selectedBuildings.findIndex(b => b.id === building.id);
    
    if (index > -1) {
        selectedBuildings.splice(index, 1);
    } else {
        selectedBuildings.push(building);
    }
    
    // 탭 카운트 업데이트 추가
    updateTabCounts();
    displayBuildingList(currentDisplayedBuildings.length > 0 ? currentDisplayedBuildings : buildingsData);
    
    // 줌 레벨에 따라 적절한 표시 업데이트
    if (currentZoomLevel < CLUSTER_MIN_LEVEL) {
        displayBuildingLabels(currentDisplayedBuildings.length > 0 ? currentDisplayedBuildings : buildingsData);
    }
    
    // 선택 목록 탭이 활성화되어 있으면 업데이트
    if (document.querySelector('[data-tab="selected"]').classList.contains('active')) {
        displaySelectedBuildingsList();
    }
}

// removeSelectedBuilding 함수 수정
function removeSelectedBuilding(index) {
    selectedBuildings.splice(index, 1);
    updateTabCounts();
    displayBuildingList(currentDisplayedBuildings.length > 0 ? currentDisplayedBuildings : buildingsData);
    
    // 줌 레벨에 따라 적절한 표시 업데이트
    if (currentZoomLevel < CLUSTER_MIN_LEVEL) {
        displayBuildingLabels(currentDisplayedBuildings.length > 0 ? currentDisplayedBuildings : buildingsData);
    }
    
    displaySelectedBuildingsList();
}