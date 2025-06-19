// ===== 전역 변수 =====
let map;
let markers = [];
let overlays = [];
let overlaysMap = {};
let clusterer = null;
let selectedBuildings = [];
let buildingsData = [];
let manager;
let currentOverlay = null;
let polygonPoints = [];
let currentDisplayedBuildings = [];
let drawEndListener = null;
let currentZoomLevel = 5;
const CLUSTER_MIN_LEVEL = 5;
const LABEL_MAX_LEVEL = 4;
let highlightedBuildingId = null;
let spreadGroups = new Map();
let buildingGroups = new Map();
let buildingMatches = null; // 공실정보 매칭 데이터

// ===== 페이지 로드 시 초기화 =====
window.onload = function() {
    initMap();
    initDrawingToolsDrag();
};

// ===== 윈도우 리사이즈 처리 =====
window.addEventListener('resize', function() {
    if (map) {
        map.relayout();
        
        // 도형 그리기 도구 위치 확인 및 조정
        const drawingTools = document.getElementById('drawing-tools');
        const mapRect = document.getElementById('map').getBoundingClientRect();
        const toolsRect = drawingTools.getBoundingClientRect();
        
        const currentLeft = parseInt(drawingTools.style.left) || 0;
        const currentTop = parseInt(drawingTools.style.top) || 0;
        
        let needsUpdate = false;
        let newLeft = currentLeft;
        let newTop = currentTop;
        
        if (currentLeft + toolsRect.width > mapRect.width) {
            newLeft = mapRect.width - toolsRect.width - 10;
            needsUpdate = true;
        }
        
        if (currentTop + toolsRect.height > mapRect.height) {
            newTop = mapRect.height - toolsRect.height - 10;
            needsUpdate = true;
        }
        
        if (needsUpdate) {
            drawingTools.style.left = newLeft + 'px';
            drawingTools.style.top = newTop + 'px';
        }
    }
});

// ===== 토글 버튼 이벤트 =====
document.getElementById('toggle-btn').addEventListener('click', function() {
    const sidebar = document.getElementById('sidebar');
    sidebar.classList.toggle('collapsed');
    this.classList.toggle('collapsed');
    this.textContent = sidebar.classList.contains('collapsed') ? '▶' : '◀';
    
    if (!map) {
        return;
    }
    
    // 지도 크기 재조정
    setTimeout(function() {
        window.dispatchEvent(new Event('resize'));
        map.relayout();
        
        setTimeout(function() {
            map.relayout();
            const center = map.getCenter();
            map.setCenter(center);
            
            // 도형 그리기 도구 위치 재조정
            const drawingTools = document.getElementById('drawing-tools');
            const mapRect = document.getElementById('map').getBoundingClientRect();
            const toolsRect = drawingTools.getBoundingClientRect();
            
            const currentLeft = parseInt(drawingTools.style.left) || 0;
            const currentTop = parseInt(drawingTools.style.top) || 0;
            
            if (currentLeft + toolsRect.width > mapRect.width) {
                drawingTools.style.left = (mapRect.width - toolsRect.width - 10) + 'px';
            }
        }, 100);
    }, 350);
});

// ===== 로딩 표시 함수 =====
function showLoading(show) {
    document.getElementById('loading-overlay').style.display = show ? 'flex' : 'none';
}

function updateLoadingProgress(text) {
    document.getElementById('loading-progress').textContent = text;
}

// ===== 통계 업데이트 =====
function updateStats() {
    const validBuildings = buildingsData.filter(b => b.lat && b.lng);
    const displayedValidBuildings = currentDisplayedBuildings.filter(b => b.lat && b.lng);
    
    document.getElementById('stat-total').textContent = buildingsData.length;
    document.getElementById('stat-displayed').textContent = displayedValidBuildings.length;
    document.getElementById('stat-geocoded').textContent = validBuildings.length;
    
    document.getElementById('building-count').textContent = 
        `전체 빌딩: ${currentDisplayedBuildings.length}개`;
}

// ===== 장바구니 카운트 업데이트 =====
function updateCartCount() {
    document.getElementById('cart-count').textContent = selectedBuildings.length;
}

// ===== 팝업 닫기 =====
function closePopup(event) {
    if (!event || event.target.classList.contains('popup-overlay')) {
        document.getElementById('buildingModal').style.display = 'none';
        // 구버전 팝업도 닫기
        document.getElementById('popup').style.display = 'none';
    }
}

// ===== 빌딩 데이터 로드 =====
async function loadBuildingData() {
    try {
        showLoading(true);
        updateLoadingProgress('빌딩 데이터를 불러오는 중...');
        
        const response = await fetch('./data/buildings.json');
        if (!response.ok) {
            throw new Error('데이터 파일을 찾을 수 없습니다');
        }
        
        const data = await response.json();
        buildingsData = data.buildings || [];
        
        // 좌표가 있는 빌딩 필터링
        const validBuildings = buildingsData.filter(b => b.lat && b.lng);
        const invalidCount = buildingsData.length - validBuildings.length;
        
        updateLoadingProgress(`${buildingsData.length}개 중 ${validBuildings.length}개 빌딩 로드 완료`);
        
        if (invalidCount > 0) {
            console.warn(`좌표가 없는 빌딩: ${invalidCount}개`);
        }
        
        // 같은 좌표의 빌딩들 그룹화
        groupBuildingsByLocation();
        
        // 통계 업데이트
        updateStats();
        
        currentDisplayedBuildings = buildingsData;
        
        // 초기 표시 설정 (줌 레벨에 따라)
        if (map.getLevel() >= CLUSTER_MIN_LEVEL) {
            createMarkers(buildingsData);
            showClustering();
        } else {
            showBuildingLabels();
        }
        
        displayBuildingList(buildingsData);
        displaySelectedBuildingsList();
        
        // 지도 범위 조정
        if (validBuildings.length > 0) {
            const bounds = new kakao.maps.LatLngBounds();
            validBuildings.forEach(building => {
                bounds.extend(new kakao.maps.LatLng(building.lat, building.lng));
            });
            map.setBounds(bounds);
        }
        
        showLoading(false);
        
        // 공실정보 매칭 데이터 로드
        await loadMatchingData();
        
    } catch (error) {
        console.error('데이터 로드 실패:', error);
        showLoading(false);
        alert('데이터를 불러오는데 실패했습니다.');
    }
}

// ===== 같은 좌표의 빌딩들 그룹화 =====
function groupBuildingsByLocation() {
    buildingGroups.clear();
    
    buildingsData.forEach(building => {
        if (!building.lat || !building.lng) return;
        
        const key = `${building.lat.toFixed(6)},${building.lng.toFixed(6)}`;
        
        if (!buildingGroups.has(key)) {
            buildingGroups.set(key, []);
        }
        buildingGroups.get(key).push(building);
    });
}
