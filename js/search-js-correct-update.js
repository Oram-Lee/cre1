// ===== 검색 관련 함수들 =====

// 검색 엔터키 처리
function handleSearchKeyPress(event) {
    if (event.key === 'Enter') {
        searchBuildings();
    }
}

// 검색 결과 메시지 표시 함수
function showSearchResultMessage(message, type = 'success') {
    const searchSection = document.querySelector('.search-section');
    
    // 모든 기존 메시지 제거
    const existingMsgs = searchSection.querySelectorAll('.result-msg');
    existingMsgs.forEach(msg => msg.remove());
    
    // 새 메시지 생성
    const resultMsg = document.createElement('div');
    resultMsg.className = 'result-msg';
    
    // 타입에 따른 스타일 설정
    let bgColor, textColor;
    switch(type) {
        case 'success':
            bgColor = '#d4edda';
            textColor = '#155724';
            break;
        case 'error':
            bgColor = '#f8d7da';
            textColor = '#721c24';
            break;
        case 'info':
            bgColor = '#d1ecf1';
            textColor = '#0c5460';
            break;
        default:
            bgColor = '#d4edda';
            textColor = '#155724';
    }
    
    resultMsg.style.cssText = `background: ${bgColor}; color: ${textColor}; padding: 10px; margin-top: 10px; border-radius: 4px;`;
    resultMsg.textContent = message;
    searchSection.appendChild(resultMsg);
    
    // 일정 시간 후 메시지 제거
    setTimeout(() => {
        if (resultMsg && resultMsg.parentNode) {
            resultMsg.remove();
        }
    }, type === 'error' ? 5000 : 3000);
}

// 빌딩 검색
function searchBuildings() {
    const name = document.getElementById('search-name').value.toLowerCase();
    const address = document.getElementById('search-address').value.toLowerCase();
    const station = document.getElementById('search-station').value.toLowerCase();
    
    // 면적 검색 조건 추가
    const areaFrom = parseFloat(document.getElementById('search-area-from').value) || 0;
    const areaTo = parseFloat(document.getElementById('search-area-to').value) || Infinity;
    
    const filtered = buildingsData.filter(building => {
        const nameMatch = !name || building.name.toLowerCase().includes(name);
        const addressMatch = !address || building.address.toLowerCase().includes(address);
        const stationMatch = !station || building.station.toLowerCase().includes(station);
        
        // 기준층 전용면적 필터링 추가
        let areaMatch = true;
        if (building.typicalFloorArea || building.기준층전용면적) {
            // 여러 형태의 필드명 처리
            const areaField = building.typicalFloorArea || building.기준층전용면적 || '';
            // 문자열에서 숫자만 추출 (예: "450평", "450 평", "450")
            const areaValue = parseFloat(areaField.toString().replace(/[^0-9.]/g, ''));
            
            if (!isNaN(areaValue)) {
                areaMatch = areaValue >= areaFrom && areaValue <= areaTo;
            } else if (areaFrom > 0 || areaTo < Infinity) {
                // 면적 검색 조건이 있는데 유효한 면적 값이 없으면 제외
                areaMatch = false;
            }
        } else if (areaFrom > 0 || areaTo < Infinity) {
            // 면적 검색 조건이 있는데 건물에 면적 데이터가 없으면 제외
            areaMatch = false;
        }
        
        return nameMatch && addressMatch && stationMatch && areaMatch;
    });
    
    displayBuildingList(filtered);
    
    // 줌 레벨에 따라 적절한 표시
    createMarkers(filtered);
    if (currentZoomLevel >= CLUSTER_MIN_LEVEL) {
        showClustering();
    } else {
        showBuildingLabels();
    }
    
    // 검색 결과 메시지 표시
    if (filtered.length === 0) {
        showSearchResultMessage('검색 결과가 없습니다.', 'error');
    } else {
        showSearchResultMessage(`검색 결과: ${filtered.length}개의 빌딩을 찾았습니다.`, 'info');
    }
    
    // 검색 결과가 하나일 경우 해당 위치로 이동
    if (filtered.length === 1 && filtered[0].lat && filtered[0].lng) {
        const position = new kakao.maps.LatLng(filtered[0].lat, filtered[0].lng);
        map.setCenter(position);
        map.setLevel(3);
    } else if (filtered.length > 1) {
        // 여러 개일 경우 모든 마커가 보이도록 지도 범위 조정
        const bounds = new kakao.maps.LatLngBounds();
        let hasValidCoords = false;
        
        filtered.forEach(building => {
            if (building.lat && building.lng) {
                bounds.extend(new kakao.maps.LatLng(building.lat, building.lng));
                hasValidCoords = true;
            }
        });
        
        if (hasValidCoords) {
            map.setBounds(bounds);
        }
    }
}

// 검색 초기화
function resetSearch() {
    document.getElementById('search-name').value = '';
    document.getElementById('search-address').value = '';
    document.getElementById('search-station').value = '';
    document.getElementById('search-area-from').value = '';
    document.getElementById('search-area-to').value = '';
    
    // 하이라이트 제거
    removeHighlight();
    
    displayBuildingList(buildingsData);
    
    // 줌 레벨에 따라 적절한 표시
    createMarkers(buildingsData);
    if (currentZoomLevel >= CLUSTER_MIN_LEVEL) {
        showClustering();
    } else {
        showBuildingLabels();
    }
    
    // 도형도 제거
    if (currentOverlay) {
        manager.remove(currentOverlay);
        currentOverlay = null;
    }
    
    // 모든 기존 메시지 제거 (메시지를 표시하지 않고 조용히 제거)
    const searchSection = document.querySelector('.search-section');
    const existingMsgs = searchSection.querySelectorAll('.result-msg');
    existingMsgs.forEach(msg => msg.remove());
}

// 영역 내 빌딩 찾기
function findBuildingsInArea(overlay) {
    return buildingsData.filter(building => {
        // 좌표가 없는 빌딩은 제외
        if (!building.lat || !building.lng) return false;
        
        const position = new kakao.maps.LatLng(building.lat, building.lng);
        
        if (overlay instanceof kakao.maps.Rectangle) {
            const bounds = overlay.getBounds();
            return bounds.contain(position);
        } else if (overlay instanceof kakao.maps.Circle) {
            const center = overlay.getPosition();
            const radius = overlay.getRadius();
            const poly = new kakao.maps.Polyline({
                path: [center, position]
            });
            const distance = poly.getLength();
            return distance <= radius;
        } else if (overlay.overlayType === 'polygon' || overlay instanceof kakao.maps.Polygon) {
            try {
                // manager.getData()로 다각형 데이터 가져오기
                const managerData = manager.getData();
                
                let coords = [];
                
                // getData에서 polygon 배열 확인
                if (managerData && managerData.polygon && managerData.polygon.length > 0) {
                    // 가장 최근에 그린 다각형 (마지막 요소)
                    const lastPolygon = managerData.polygon[managerData.polygon.length - 1];
                    
                    if (lastPolygon.points && Array.isArray(lastPolygon.points)) {
                        // points 배열에서 좌표 추출
                        coords = lastPolygon.points.map(point => {
                            return new kakao.maps.LatLng(point.y, point.x);
                        });
                    }
                }
                
                // coords가 없으면 polygonPoints 사용
                if (coords.length === 0 && polygonPoints.length > 0) {
                    coords = polygonPoints;
                }
                
                // 여전히 좌표가 없으면 false 반환
                if (coords.length < 3) {
                    return false;
                }
                
                // Ray Casting Algorithm으로 내부 판별
                let inside = false;
                const x = position.getLng();
                const y = position.getLat();
                
                for (let i = 0, j = coords.length - 1; i < coords.length; j = i++) {
                    const xi = coords[i].getLng();
                    const yi = coords[i].getLat();
                    const xj = coords[j].getLng();
                    const yj = coords[j].getLat();
                    
                    const intersect = ((yi > y) !== (yj > y))
                        && (x < (xj - xi) * (y - yi) / (yj - yi) + xi);
                    
                    if (intersect) inside = !inside;
                }
                
                return inside;
                
            } catch (error) {
                console.error('다각형 내부 판별 오류:', error);
                return false;
            }
        }
        
        return false;
    });
}