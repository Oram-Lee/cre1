// ===== 도형 그리기 관련 함수들 =====

// 그리기 모드 설정
function setDrawingMode(type) {
    // 기존 도형 제거
    if (currentOverlay) {
        manager.remove(currentOverlay);
        currentOverlay = null;
    }
    
    // 모든 버튼 비활성화
    document.querySelectorAll('.tool-btn').forEach(btn => {
        btn.classList.remove('active');
    });
    
    // 선택된 버튼 활성화
    event.target.classList.add('active');
    
    // 다각형 점들 초기화
    polygonPoints = [];
    
    // 기존 drawend 리스너 제거
    if (drawEndListener) {
        manager.removeListener('drawend', drawEndListener);
        drawEndListener = null;
    }
    
    // Drawing Manager로 그리기 모드 선택
    if (type === 'rectangle') {
        manager.select(kakao.maps.drawing.OverlayType.RECTANGLE);
    } else if (type === 'circle') {
        manager.select(kakao.maps.drawing.OverlayType.CIRCLE);
    } else if (type === 'polygon') {
        manager.select(kakao.maps.drawing.OverlayType.POLYGON);
    }
    
    // 그리기 완료 이벤트 리스너 추가
    drawEndListener = function(data) {
        currentOverlay = data.target;
        
        // 도형에 이벤트 리스너 추가 (드래그 종료 시 재검색)
        if (currentOverlay) {
            // 각 도형 타입별로 이벤트 추가
            if (data.overlayType === 'rectangle' || data.overlayType === 'circle' || data.overlayType === 'polygon') {
                // 마우스 이벤트로 드래그 감지
                let isDragging = false;
                
                kakao.maps.event.addListener(currentOverlay, 'mousedown', function() {
                    isDragging = true;
                });
                
                kakao.maps.event.addListener(currentOverlay, 'mouseup', function() {
                    if (isDragging) {
                        isDragging = false;
                        // 드래그 완료 후 재검색
                        setTimeout(() => {
                            const buildingsInArea = findBuildingsInArea(currentOverlay);
                            displayBuildingList(buildingsInArea);
                            
                            // 줌 레벨에 따라 적절한 표시
                            createMarkers(buildingsInArea);
                            if (currentZoomLevel >= CLUSTER_MIN_LEVEL) {
                                showClustering();
                            } else {
                                showBuildingLabels();
                            }
                            
                            // 검색 결과 알림
                            showSearchResultMessage(`이동된 영역에서 ${buildingsInArea.length}개의 빌딩을 찾았습니다.`, 'success');
                        }, 100);
                    }
                });
            }
        }
        
        // 영역 내 빌딩 검색
        const buildingsInArea = findBuildingsInArea(data.target);
        displayBuildingList(buildingsInArea);
        
        // 줌 레벨에 따라 적절한 표시
        createMarkers(buildingsInArea);
        if (currentZoomLevel >= CLUSTER_MIN_LEVEL) {
            showClustering();
        } else {
            showBuildingLabels();
        }
        
        // 그리기 모드 해제
        manager.cancel();
        
        // 버튼 비활성화
        document.querySelectorAll('.tool-btn').forEach(btn => {
            btn.classList.remove('active');
        });
        
        // 검색 결과 알림
        showSearchResultMessage(`선택한 영역에서 ${buildingsInArea.length}개의 빌딩을 찾았습니다.`, 'success');
    };
    
    manager.addListener('drawend', drawEndListener);
}

// 그리기 지우기
function clearDrawing() {
    if (currentOverlay) {
        manager.remove(currentOverlay);
        currentOverlay = null;
    }
    
    manager.cancel();
    
    // 모든 버튼 비활성화
    document.querySelectorAll('.tool-btn').forEach(btn => {
        btn.classList.remove('active');
    });
    
    // 전체 빌딩 표시
    displayBuildingList(buildingsData);
    
    // 줌 레벨에 따라 적절한 표시
    createMarkers(buildingsData);
    if (currentZoomLevel >= CLUSTER_MIN_LEVEL) {
        showClustering();
    } else {
        showBuildingLabels();
    }
    
    // 모든 기존 메시지 제거 (메시지를 표시하지 않고 조용히 제거)
    const searchSection = document.querySelector('.search-section');
    const existingMsgs = searchSection.querySelectorAll('.result-msg');
    existingMsgs.forEach(msg => msg.remove());
}

// 도형 그리기 도구 드래그 기능 초기화
function initDrawingToolsDrag() {
    const drawingTools = document.getElementById('drawing-tools');
    let isDragging = false;
    let startX, startY, initialLeft, initialTop;
    
    // 로컬 스토리지에서 저장된 위치 불러오기
    const savedPosition = localStorage.getItem('drawingToolsPosition');
    if (savedPosition) {
        const position = JSON.parse(savedPosition);
        const mapRect = document.getElementById('map').getBoundingClientRect();
        const toolsRect = drawingTools.getBoundingClientRect();
        
        // 저장된 위치가 현재 맵 영역을 벗어나지 않는지 확인
        let left = position.left;
        let top = position.top;
        
        if (left < 0) left = 0;
        if (left + toolsRect.width > mapRect.width) {
            left = mapRect.width - toolsRect.width;
        }
        if (top < 0) top = 0;
        if (top + toolsRect.height > mapRect.height) {
            top = mapRect.height - toolsRect.height;
        }
        
        drawingTools.style.left = left + 'px';
        drawingTools.style.top = top + 'px';
        drawingTools.style.right = 'auto';
    }
    
    // 드래그 시작
    drawingTools.addEventListener('mousedown', function(e) {
        // 버튼 클릭은 제외
        if (e.target.tagName === 'BUTTON') return;
        
        isDragging = true;
        startX = e.clientX;
        startY = e.clientY;
        
        const rect = drawingTools.getBoundingClientRect();
        initialLeft = rect.left;
        initialTop = rect.top;
        
        drawingTools.classList.add('dragging');
        e.preventDefault();
    });
    
    // 드래그 중
    document.addEventListener('mousemove', function(e) {
        if (!isDragging) return;
        
        const deltaX = e.clientX - startX;
        const deltaY = e.clientY - startY;
        
        let newLeft = initialLeft + deltaX;
        let newTop = initialTop + deltaY;
        
        // 화면 밖으로 나가지 않도록 제한
        const toolsRect = drawingTools.getBoundingClientRect();
        const toolsWidth = toolsRect.width;
        const toolsHeight = toolsRect.height;
        const mapRect = document.getElementById('map').getBoundingClientRect();
        const mapWidth = mapRect.width;
        const mapHeight = mapRect.height;
        
        // 좌우 경계 체크 (맵 영역 기준)
        if (newLeft < mapRect.left) newLeft = mapRect.left;
        if (newLeft + toolsWidth > mapRect.left + mapWidth) {
            newLeft = mapRect.left + mapWidth - toolsWidth;
        }
        
        // 상하 경계 체크 (맵 영역 기준)
        if (newTop < mapRect.top) newTop = mapRect.top;
        if (newTop + toolsHeight > mapRect.top + mapHeight) {
            newTop = mapRect.top + mapHeight - toolsHeight;
        }
        
        // 맵 영역 내의 상대 위치로 변환
        const relativeLeft = newLeft - mapRect.left;
        const relativeTop = newTop - mapRect.top;
        
        drawingTools.style.left = relativeLeft + 'px';
        drawingTools.style.top = relativeTop + 'px';
        drawingTools.style.right = 'auto';
        
        e.preventDefault();
    });
    
    // 드래그 종료
    document.addEventListener('mouseup', function() {
        if (isDragging) {
            isDragging = false;
            drawingTools.classList.remove('dragging');
            
            // 위치 저장 (맵 영역 기준 상대 위치)
            const rect = drawingTools.getBoundingClientRect();
            const mapRect = document.getElementById('map').getBoundingClientRect();
            localStorage.setItem('drawingToolsPosition', JSON.stringify({
                left: rect.left - mapRect.left,
                top: rect.top - mapRect.top
            }));
        }
    });
    
    // 터치 이벤트 지원 (모바일)
    drawingTools.addEventListener('touchstart', function(e) {
        if (e.target.tagName === 'BUTTON') return;
        
        isDragging = true;
        const touch = e.touches[0];
        startX = touch.clientX;
        startY = touch.clientY;
        
        const rect = drawingTools.getBoundingClientRect();
        initialLeft = rect.left;
        initialTop = rect.top;
        
        drawingTools.classList.add('dragging');
        e.preventDefault();
    });
    
    document.addEventListener('touchmove', function(e) {
        if (!isDragging) return;
        
        const touch = e.touches[0];
        const deltaX = touch.clientX - startX;
        const deltaY = touch.clientY - startY;
        
        let newLeft = initialLeft + deltaX;
        let newTop = initialTop + deltaY;
        
        // 화면 밖으로 나가지 않도록 제한
        const toolsRect = drawingTools.getBoundingClientRect();
        const toolsWidth = toolsRect.width;
        const toolsHeight = toolsRect.height;
        const mapRect = document.getElementById('map').getBoundingClientRect();
        const mapWidth = mapRect.width;
        const mapHeight = mapRect.height;
        
        // 좌우 경계 체크 (맵 영역 기준)
        if (newLeft < mapRect.left) newLeft = mapRect.left;
        if (newLeft + toolsWidth > mapRect.left + mapWidth) {
            newLeft = mapRect.left + mapWidth - toolsWidth;
        }
        
        // 상하 경계 체크 (맵 영역 기준)
        if (newTop < mapRect.top) newTop = mapRect.top;
        if (newTop + toolsHeight > mapRect.top + mapHeight) {
            newTop = mapRect.top + mapHeight - toolsHeight;
        }
        
        // 맵 영역 내의 상대 위치로 변환
        const relativeLeft = newLeft - mapRect.left;
        const relativeTop = newTop - mapRect.top;
        
        drawingTools.style.left = relativeLeft + 'px';
        drawingTools.style.top = relativeTop + 'px';
        drawingTools.style.right = 'auto';
        
        e.preventDefault();
    });
    
    document.addEventListener('touchend', function() {
        if (isDragging) {
            isDragging = false;
            drawingTools.classList.remove('dragging');
            
            // 위치 저장 (맵 영역 기준 상대 위치)
            const rect = drawingTools.getBoundingClientRect();
            const mapRect = document.getElementById('map').getBoundingClientRect();
            localStorage.setItem('drawingToolsPosition', JSON.stringify({
                left: rect.left - mapRect.left,
                top: rect.top - mapRect.top
            }));
        }
    });
}

// 도형 그리기 도구 위치 초기화
function resetDrawingToolsPosition() {
    const drawingTools = document.getElementById('drawing-tools');
    drawingTools.style.top = '10px';
    drawingTools.style.right = '10px';
    drawingTools.style.left = 'auto';
    
    // 저장된 위치 삭제
    localStorage.removeItem('drawingToolsPosition');
    
    // 애니메이션 효과
    drawingTools.style.transition = 'all 0.3s ease';
    setTimeout(() => {
        drawingTools.style.transition = '';
    }, 300);
}

// 도형 그리기 도구 최소화/최대화
function toggleMinimize() {
    const drawingTools = document.getElementById('drawing-tools');
    const minimizeBtn = document.getElementById('minimize-btn');
    
    drawingTools.classList.toggle('minimized');
    
    if (drawingTools.classList.contains('minimized')) {
        minimizeBtn.textContent = '+';
        minimizeBtn.title = '최대화';
    } else {
        minimizeBtn.textContent = '−';
        minimizeBtn.title = '최소화';
    }
}
