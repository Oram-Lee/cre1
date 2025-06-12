// ===== 지도 관련 함수들 =====

// 지도 초기화
function initMap() {
    const container = document.getElementById('map');
    const options = {
        center: new kakao.maps.LatLng(37.5065, 127.0550),
        level: 5
    };
    
    map = new kakao.maps.Map(container, options);
    
    // 지도 타입 설정
    const mapTypeControl = new kakao.maps.MapTypeControl();
    map.addControl(mapTypeControl, kakao.maps.ControlPosition.TOPRIGHT);
    
    // 줌 컨트롤 추가
    const zoomControl = new kakao.maps.ZoomControl();
    map.addControl(zoomControl, kakao.maps.ControlPosition.RIGHT);
    
    // 줌 레벨 변경 이벤트 리스너
    kakao.maps.event.addListener(map, 'zoom_changed', function() {
        currentZoomLevel = map.getLevel();
        console.log('현재 줌 레벨:', currentZoomLevel);
        
        // 줌 레벨에 따라 표시 방식 변경
        if (currentZoomLevel >= CLUSTER_MIN_LEVEL) {
            // 클러스터링 표시
            showClustering();
        } else {
            // 빌딩명 라벨 표시
            showBuildingLabels();
        }
        
        updateLabelSizes();
    });
    
    // 클러스터러 생성
    clusterer = new kakao.maps.MarkerClusterer({
        map: null, // 초기에는 지도에 표시하지 않음
        averageCenter: true,
        minLevel: CLUSTER_MIN_LEVEL,
        disableClickZoom: false,
        calculator: [10, 30, 50, 100], // 클러스터 크기별 스타일
        styles: [{
            width : '30px', height : '30px',
            background: 'rgba(51, 153, 255, .8)',
            borderRadius: '15px',
            color: '#fff',
            textAlign: 'center',
            fontWeight: 'bold',
            lineHeight: '31px'
        }, {
            width : '40px', height : '40px',
            background: 'rgba(255, 153, 51, .8)',
            borderRadius: '20px',
            color: '#fff',
            textAlign: 'center',
            fontWeight: 'bold',
            lineHeight: '41px'
        }, {
            width : '50px', height : '50px',
            background: 'rgba(255, 51, 51, .8)',
            borderRadius: '25px',
            color: '#fff',
            textAlign: 'center',
            fontWeight: 'bold',
            lineHeight: '51px'
        }, {
            width : '60px', height : '60px',
            background: 'rgba(204, 51, 255, .8)',
            borderRadius: '30px',
            color: '#fff',
            textAlign: 'center',
            fontWeight: 'bold',
            lineHeight: '61px'
        }]
    });
    
    // 그리기 관리자 생성
    const drawingOptions = {
        map: map,
        drawingMode: [
            kakao.maps.drawing.OverlayType.RECTANGLE,
            kakao.maps.drawing.OverlayType.CIRCLE,
            kakao.maps.drawing.OverlayType.POLYGON
        ],
        guideTooltip: ['draw', 'drag', 'edit'],
        rectangleOptions: {
            draggable: true,
            removable: true,
            editable: true,
            strokeWeight: 2,
            strokeColor: '#39f',
            fillColor: '#39f',
            fillOpacity: 0.3
        },
        circleOptions: {
            draggable: true,
            removable: true,
            editable: true,
            strokeWeight: 2,
            strokeColor: '#39f',
            fillColor: '#39f',
            fillOpacity: 0.3
        },
        polygonOptions: {
            draggable: true,
            removable: true,
            editable: true,
            strokeWeight: 2,
            strokeColor: '#39f',
            fillColor: '#39f',
            fillOpacity: 0.3
        }
    };
    
    manager = new kakao.maps.drawing.DrawingManager(drawingOptions);
    drawEndListener = null; // 리스너 초기화
    
    // 실제 데이터 로드
    loadBuildingData();
}

// 줌 레벨에 따른 라벨 크기 조정
function updateLabelSizes() {
    const labels = document.querySelectorAll('.building-label, .building-spread-item');
    labels.forEach(label => {
        label.classList.remove('zoom-far', 'zoom-medium', 'zoom-near');
        
        if (currentZoomLevel >= 7) {
            label.classList.add('zoom-far');
        } else if (currentZoomLevel >= 4) {
            label.classList.add('zoom-medium');
        } else {
            label.classList.add('zoom-near');
        }
    });
}

// 마커 생성
function createMarkers(buildings) {
    // 기존 마커 제거
    markers.forEach(marker => marker.setMap(null));
    markers = [];
    
    buildings.forEach(building => {
        // 좌표가 있는 빌딩만 마커 생성
        if (!building.lat || !building.lng) return;
        
        const position = new kakao.maps.LatLng(building.lat, building.lng);
        
        const marker = new kakao.maps.Marker({
            position: position,
            title: building.name
        });
        
        // 마커 클릭 이벤트
        kakao.maps.event.addListener(marker, 'click', function() {
            showBuildingInfo(building);
        });
        
        markers.push(marker);
    });
}

// 클러스터링 표시
function showClustering() {
    console.log('클러스터링 모드로 전환');
    
    // 빌딩명 라벨 숨기기
    overlays.forEach(overlay => overlay.setMap(null));
    
    // 펼쳐진 그룹 초기화
    spreadGroups.clear();
    
    // 클러스터러에 마커 추가
    if (clusterer) {
        clusterer.clear();
        clusterer.addMarkers(markers);
        clusterer.setMap(map);
    }
}

// 빌딩명 라벨 표시
function showBuildingLabels() {
    console.log('빌딩명 표시 모드로 전환');
    
    // 클러스터러 숨기기
    if (clusterer) {
        clusterer.setMap(null);
    }
    
    // 마커도 숨기기
    markers.forEach(marker => marker.setMap(null));
    
    // 빌딩명 라벨 표시
    displayBuildingLabels(currentDisplayedBuildings);
}

// 빌딩 라벨(커스텀 오버레이) 표시
function displayBuildingLabels(buildings) {
    // 기존 오버레이 제거
    overlays.forEach(overlay => overlay.setMap(null));
    overlays = [];
    overlaysMap = {};
    
    // 그룹별로 처리
    const processedGroups = new Set();
    
    buildings.forEach(building => {
        // 좌표가 있는 빌딩만 표시
        if (!building.lat || !building.lng) return;
        
        const key = `${building.lat.toFixed(6)},${building.lng.toFixed(6)}`;
        
        // 이미 처리한 그룹이면 스킵
        if (processedGroups.has(key)) return;
        
        const group = buildingGroups.get(key);
        
        if (group && group.length > 1) {
            // 여러 빌딩이 같은 위치에 있는 경우
            processedGroups.add(key);
            
            // 현재 표시된 빌딩들 중에서 그룹 필터링
            const displayedGroup = group.filter(b => 
                buildings.some(displayedB => displayedB.id === b.id)
            );
            
            if (displayedGroup.length > 1) {
                // 펼쳐진 상태인지 확인
                if (spreadGroups.has(key)) {
                    // 펼쳐진 상태로 표시
                    displaySpreadBuildings(displayedGroup, key);
                } else {
                    // 그룹으로 표시
                    displayBuildingGroup(displayedGroup, key);
                }
            } else if (displayedGroup.length === 1) {
                // 필터링 후 하나만 남은 경우
                displaySingleBuilding(displayedGroup[0]);
            }
        } else {
            // 단일 빌딩
            displaySingleBuilding(building);
        }
    });
    
    // 클릭 이벤트 위임
    setTimeout(() => {
        updateLabelSizes();
    }, 100);
    
    // 통계 업데이트
    updateStats();
}

// 단일 빌딩 표시
function displaySingleBuilding(building) {
    const position = new kakao.maps.LatLng(building.lat, building.lng);
    
    let additionalClass = '';
    if (selectedBuildings.some(b => b.id === building.id)) {
        additionalClass = ' selected';
    }
    if (building.id === highlightedBuildingId) {
        additionalClass = ' highlighted';
    }
    
    const content = `<div class="building-label${additionalClass}" data-building-id="${building.id}">${building.name}</div>`;
    
    const customOverlay = new kakao.maps.CustomOverlay({
        position: position,
        content: content,
        yAnchor: 0.5,
        clickable: true,
        zIndex: building.id === highlightedBuildingId ? 10001 : (selectedBuildings.some(b => b.id === building.id) ? 10 : 1)
    });
    
    customOverlay.setMap(map);
    overlays.push(customOverlay);
    overlaysMap[building.id] = customOverlay;
    
    // 클릭 이벤트
    setTimeout(() => {
        const label = document.querySelector(`[data-building-id="${building.id}"]`);
        if (label) {
            label.addEventListener('click', function() {
                showBuildingInfo(building);
            });
        }
    }, 100);
}

// 빌딩 그룹 표시
function displayBuildingGroup(buildings, locationKey) {
    const firstBuilding = buildings[0];
    const position = new kakao.maps.LatLng(firstBuilding.lat, firstBuilding.lng);
    
    const content = `<div class="building-group-label" data-location-key="${locationKey}">
        ${buildings[0].name} 외 ${buildings.length - 1}개
    </div>`;
    
    const customOverlay = new kakao.maps.CustomOverlay({
        position: position,
        content: content,
        yAnchor: 0.5,
        clickable: true,
        zIndex: 100
    });
    
    customOverlay.setMap(map);
    overlays.push(customOverlay);
    
    // 클릭 이벤트
    setTimeout(() => {
        const label = document.querySelector(`[data-location-key="${locationKey}"]`);
        if (label) {
            label.addEventListener('click', function() {
                // 그룹 펼치기
                spreadGroups.set(locationKey, true);
                showBuildingLabels(); // 다시 그리기
            });
        }
    }, 100);
}

// 펼쳐진 빌딩들 표시
function displaySpreadBuildings(buildings, locationKey) {
    const centerLat = buildings[0].lat;
    const centerLng = buildings[0].lng;
    const radius = 0.0003; // 약 30m
    
    buildings.forEach((building, index) => {
        const angle = (index * 360) / buildings.length;
        const radians = (angle * Math.PI) / 180;
        
        const lat = centerLat + radius * Math.sin(radians);
        const lng = centerLng + radius * Math.cos(radians);
        
        const position = new kakao.maps.LatLng(lat, lng);
        
        let additionalClass = '';
        if (selectedBuildings.some(b => b.id === building.id)) {
            additionalClass = ' selected';
        }
        if (building.id === highlightedBuildingId) {
            additionalClass = ' highlighted';
        }
        
        const content = `
            <div class="building-spread-container">
                <div class="building-spread-item${additionalClass}" 
                     data-building-id="${building.id}"
                     data-location-key="${locationKey}">
                    ${building.name}
                </div>
            </div>
        `;
        
        const customOverlay = new kakao.maps.CustomOverlay({
            position: position,
            content: content,
            yAnchor: 0.5,
            clickable: true,
            zIndex: building.id === highlightedBuildingId ? 10001 : (selectedBuildings.some(b => b.id === building.id) ? 10 : 1)
        });
        
        customOverlay.setMap(map);
        overlays.push(customOverlay);
        overlaysMap[building.id] = customOverlay;
    });
    
    // 클릭 이벤트
    setTimeout(() => {
        document.querySelectorAll('.building-spread-item').forEach(label => {
            label.addEventListener('click', function() {
                const buildingId = parseInt(this.dataset.buildingId);
                const building = buildingsData.find(b => b.id === buildingId);
                if (building) {
                    showBuildingInfo(building);
                }
            });
        });
        
        updateLabelSizes();
    }, 100);
    
    // 지도 외부 클릭 시 펼쳐진 그룹 닫기
    kakao.maps.event.addListener(map, 'click', function() {
        if (spreadGroups.has(locationKey)) {
            spreadGroups.delete(locationKey);
            showBuildingLabels();
        }
    });
}
