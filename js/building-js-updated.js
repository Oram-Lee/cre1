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
                    <span class="info-label">대지면적</span>
                    <span class="info-value">${building.landAreaPy ? building.landAreaPy.toLocaleString() + '평' : '-'}</span>
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

// 현재 날짜 반환
function getCurrentDate() {
    return new Date().toISOString().split('T')[0];
}

// ===== ExcelJS를 사용한 엑셀 내보내기 =====
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
        // ExcelJS 워크북 생성
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('후보지');
        
        // 1. 열 너비 설정
        worksheet.columns = [
            { width: 2.6640625 },   // A열
            { width: 13.21875 },    // B열
            { width: 24.5546875 },  // C열
            { width: 26.33203125 }, // D열
            { width: 26.33203125 }, // E열
            { width: 26.33203125 }, // F열
            { width: 26.33203125 }, // G열
            { width: 26.33203125 }  // H열
        ];
        
        // 2. 행 높이 설정
        worksheet.getRow(1).height = 16.9;
        worksheet.getRow(2).height = 49.9;
        worksheet.getRow(3).height = 16.9;
        worksheet.getRow(4).height = 16.9;
        worksheet.getRow(5).height = 190.15; // 이미지 영역
        worksheet.getRow(6).height = 79.9;
        worksheet.getRow(9).height = 60.0; // 위치 정보
        
        // 나머지 행들은 기본 높이
        for (let i = 7; i <= 55; i++) {
            if (i !== 9) {
                worksheet.getRow(i).height = 16.9;
            }
        }
        
        // 3. 셀 병합 - 전체 구조 재정리
        worksheet.mergeCells('B3:C4');   // PRESENT TO (3-4행)
        worksheet.mergeCells('B5:C5');   // 로고 영역 (5행)
        worksheet.mergeCells('B6:C6');   // 빌딩개요/일반 (6행)
        worksheet.mergeCells('B7:B18');  // 빌딩 현황 (7-18행)
        worksheet.mergeCells('B19:B20'); // 빌딩 세부현황 (19-20행)
        worksheet.mergeCells('B21:B23'); // 주차 관련 (21-23행)
        // B24는 단독 셀 (안내문구)
        worksheet.mergeCells('B25:B31'); // 임차 제안 (25-31행)
        worksheet.mergeCells('B32:B39'); // 임대 기준 (32-39행)
        worksheet.mergeCells('B40:B44'); // 임대 기준 조정 (40-44행)
        // B45는 빈 행
        worksheet.mergeCells('B46:B50'); // 예상비용 (46-50행)
        
        // 4. 템플릿 기본 텍스트 설정
        const b3 = worksheet.getCell('B3');
        b3.value = 'PRESENT TO :';
        b3.font = { name: 'Noto Sans KR', size: 9, bold: true, color: { argb: 'FFFFFFFF' } };
        b3.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2C2A2A' } };
        b3.alignment = { horizontal: 'center', vertical: 'middle' };
        b3.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };
        
        // B5 로고 영역
        const b5 = worksheet.getCell('B5');
        b5.value = '고객사 로고 삽입';
        b5.font = { name: 'Noto Sans KR', size: 11, bold: true };
        b5.alignment = { horizontal: 'center', vertical: 'middle' };
        b5.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };
        
        // 카테고리 설정 - 템플릿과 동일한 색상 적용
        setCategoryCell(worksheet, 'B6', '빌딩개요/일반', 'FFFFFFFF');
        setCategoryCell(worksheet, 'B7', '빌딩 현황', 'FFFFFFFF');
        setCategoryCell(worksheet, 'B19', '빌딩 세부현황', 'FFFFFFFF');
        setCategoryCell(worksheet, 'B21', '주차 관련', 'FFFFFFFF');
        // B24는 빈 셀로 유지
        setCategoryCell(worksheet, 'B25', '임차 제안', 'FFF9D6AE');
        setCategoryCell(worksheet, 'B32', '임대 기준', 'FFD9ECF2');
        setCategoryCell(worksheet, 'B40', '임대기준 조정', 'FFD9ECF2');
        setCategoryCell(worksheet, 'B46', '예상비용', 'FFFBCF3A');
        
        // C열 항목명 설정
        const cColumnData = {
            7: '주소 지번',
            8: '도로명 주소',
            9: '위치',
            10: '빌딩 규모',
            11: '준공연도',
            12: '전용률 (%)',
            13: '기준층 임대면적 (m²)',
            14: '기준층 임대면적 (평)',
            15: '기준층 전용면적 (m²)',
            16: '기준층 전용면적 (평)',
            17: '엘레베이터',
            18: '냉난방 방식',
            19: '건물용도',
            20: '구조',
            21: '주차 대수 정보',
            22: '주차비',
            23: '주차 대수',
            // 24행은 빈칸
            25: '최적 임차 층수',
            26: '입주 가능 시기',
            27: '거래유형',
            28: '임대면적 (m²)',
            29: '전용면적 (m²)',
            30: '임대면적 (평)',
            31: '전용면적 (평)',
            32: '월 평당 보증금',
            33: '월 평당 임대료',
            34: '월 평당 관리비',
            35: '월 평당 지출비용',
            36: '총 보증금',
            37: '월 임대료 총액',
            38: '월 관리비 총액',
            39: '월 전용면적당 지출비용',
            40: '보증금',
            41: '렌트프리 (개월/년)',
            42: '평균 임대료',
            43: '관리비',
            44: 'NOC',
            46: '보증금',
            47: '평균 월 임대료',
            48: '평균 월 관리비',
            49: '월 (임대료 + 관리비)',
            50: '연 실제 부담 고정금액'
        };
        
        // C열 데이터 입력 및 스타일 적용 - 모든 항목 가운데 정렬
        Object.entries(cColumnData).forEach(([row, value]) => {
            const cell = worksheet.getCell(`C${row}`);
            cell.value = value;
            cell.font = { name: 'Noto Sans KR', size: 9 };
            
            // 모든 항목명 가운데 정렬
            cell.alignment = { 
                horizontal: 'center', 
                vertical: 'middle' 
            };
            cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
        });
        
        // 5. 빌딩 데이터 입력 (D열부터)
        selectedBuildings.forEach((building, index) => {
            const col = String.fromCharCode(68 + index); // D, E, F, G, H
            
            // 빌딩명 헤더
            const headerCell = worksheet.getCell(`${col}4`);
            headerCell.value = building.name;
            headerCell.font = { name: 'Noto Sans KR', size: 9, bold: true };
            headerCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFCCCCCC' } };
            headerCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
            headerCell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
            
            // 빌딩개요/일반 (D6~H6에 description 입력)
            if (building.description) {
                const descCell = worksheet.getCell(`${col}6`);
                descCell.value = building.description;
                descCell.font = { name: 'Noto Sans KR', size: 9 };
                descCell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
                descCell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
            }
            
            // 빌딩 데이터 입력
            fillBuildingData(worksheet, building, col);
        });
        
        // 6. 용어 설명 추가
        worksheet.getCell('B52').value = '용어 설명';
        worksheet.getCell('B52').font = { name: 'Noto Sans KR', size: 10, bold: true };
        
        worksheet.getCell('B53').value = 'NOC : Net Operating Cost의 약자로 임대료와 관리비를 합친 부동산 순 운영 비용';
        worksheet.getCell('B54').value = '렌트프리 : 임대료만 면제 (관리비, 보증금 有)';
        worksheet.getCell('B55').value = '프리렌트 : 임대료 + 관리비 면제 (보증금 有)';
        
        [53, 54, 55].forEach(row => {
            worksheet.getCell(`B${row}`).font = { name: 'Noto Sans KR', size: 10 };
            worksheet.getCell(`B${row}`).alignment = { horizontal: 'left', vertical: 'middle' };
        });
        
        // 7. 파일 저장
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        saveAs(blob, `CompList_${getCurrentDate()}.xlsx`);
        
        alert(`✅ Comp List 생성 완료!\n\n` +
              `📊 빌딩 ${selectedBuildings.length}개의 기본 정보가 입력되었습니다.\n\n` +
              `📝 추가 입력 필요 항목:\n` +
              `• 로고 및 빌딩 외관 이미지\n` +
              `• 임차 제안 (최적 층수, 입주 시기, 거래유형, 면적)\n` +
              `• 임대 기준 (보증금, 임대료, 관리비)\n` +
              `• 렌트프리 개월 수\n\n` +
              `💡 입력한 정보에 따라 예상비용이 자동 계산됩니다.`);
        
    } catch (error) {
        console.error('엑셀 생성 오류:', error);
        alert('엑셀 파일 생성 중 오류가 발생했습니다.\n' + error.message);
    }
}

// 카테고리 셀 설정 헬퍼 함수
function setCategoryCell(worksheet, cellAddress, value, bgColor, isRed = false) {
    const cell = worksheet.getCell(cellAddress);
    cell.value = value;
    cell.font = { 
        name: 'Noto Sans KR', 
        size: 9, 
        bold: true, 
        color: isRed ? { argb: 'FFFF0000' } : { argb: 'FF000000' } 
    };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgColor } };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };
}

// 빌딩 데이터 채우기 함수
function fillBuildingData(worksheet, building, col) {
    // 기본 정보
    setDataCell(worksheet, `${col}7`, building.addressJibun || '');
    setDataCell(worksheet, `${col}8`, building.address || '');
    setDataCell(worksheet, `${col}9`, building.station || '');
    setDataCell(worksheet, `${col}10`, building.floors || '');
    setDataCell(worksheet, `${col}11`, building.completionYear || '');
    
    // 전용률
    const dedicatedRateCell = worksheet.getCell(`${col}12`);
    dedicatedRateCell.value = (building.dedicatedRate || 0) / 100;
    dedicatedRateCell.numFmt = '0.00%';
    applyDataCellStyle(dedicatedRateCell);
    
    // 면적 정보 (m²)
    setNumericCell(worksheet, `${col}13`, building.baseFloorArea || 0, '#,##0.000 "m²"');
    setNumericCell(worksheet, `${col}14`, building.baseFloorAreaPy || 0, '#,##0.000 "평"');
    setNumericCell(worksheet, `${col}15`, building.baseFloorAreaDedicated || 0, '#,##0.000 "m²"');
    setNumericCell(worksheet, `${col}16`, building.baseFloorAreaDedicatedPy || 0, '#,##0.000 "평"');
    
    // 빌딩 세부현황
    setDataCell(worksheet, `${col}17`, building.elevator || '');  // 엘레베이터
    setDataCell(worksheet, `${col}18`, building.hvac || '');
    setDataCell(worksheet, `${col}19`, building.buildingUse || '');
    setDataCell(worksheet, `${col}20`, building.structure || '');
    setDataCell(worksheet, `${col}21`, building.parkingSpace || '');  // 주차 대수 정보
    
    // 주차 관련
    setDataCell(worksheet, `${col}22`, building.parkingFee || '');  // 주차비
    setDataCell(worksheet, `${col}23`, building.parkingSpace || '');
    // 24행은 빈칸
    
    // 임차 제안 (기본값)
    setDataCell(worksheet, `${col}25`, '-');
    setDataCell(worksheet, `${col}26`, '-');
    setDataCell(worksheet, `${col}27`, '-');
    
    // 임대면적/전용면적 - 평 기준 입력, m²는 수식으로 자동 계산
    // 28행: 임대면적 (m²) = ROUNDDOWN(D30*3.305785, 3)
    worksheet.getCell(`${col}28`).value = { formula: `ROUNDDOWN(${col}30*3.305785,3)` };
    worksheet.getCell(`${col}28`).numFmt = '#,##0.000 "m²"';
    applyDataCellStyle(worksheet.getCell(`${col}28`));
    
    // 29행: 전용면적 (m²) = ROUNDDOWN(D31*3.305785, 3)
    worksheet.getCell(`${col}29`).value = { formula: `ROUNDDOWN(${col}31*3.305785,3)` };
    worksheet.getCell(`${col}29`).numFmt = '#,##0.000 "m²"';
    applyDataCellStyle(worksheet.getCell(`${col}29`));
    
    // 30행, 31행: 평 단위 (사용자 입력, 기본값 100으로 설정하여 0으로 나누기 방지)
    setNumericCell(worksheet, `${col}30`, 100, '#,##0.000 "평"');
    setNumericCell(worksheet, `${col}31`, 50, '#,##0.000 "평"');
    
    // 임대 기준 (32-35행: 사용자 입력)
    setNumericCell(worksheet, `${col}32`, 0, '₩#,##0', 'right');  // 월 평당 보증금
    setNumericCell(worksheet, `${col}33`, 0, '₩#,##0', 'right');  // 월 평당 임대료
    setNumericCell(worksheet, `${col}34`, 0, '₩#,##0', 'right');  // 월 평당 관리비
    
    // 35행: 월 평당 지출비용 = D33+D34
    worksheet.getCell(`${col}35`).value = { formula: `${col}33+${col}34` };
    worksheet.getCell(`${col}35`).numFmt = '₩#,##0';
    applyDataCellStyle(worksheet.getCell(`${col}35`), 'right');
    
    // 36행: 총 보증금 = D32*D30
    worksheet.getCell(`${col}36`).value = { formula: `${col}32*${col}30` };
    worksheet.getCell(`${col}36`).numFmt = '₩#,##0';
    applyDataCellStyle(worksheet.getCell(`${col}36`), 'right');
    
    // 37행: 월 임대료 총액 = D33*D30
    worksheet.getCell(`${col}37`).value = { formula: `${col}33*${col}30` };
    worksheet.getCell(`${col}37`).numFmt = '₩#,##0';
    applyDataCellStyle(worksheet.getCell(`${col}37`), 'right');
    
    // 38행: 월 관리비 총액 = D34*D30
    worksheet.getCell(`${col}38`).value = { formula: `${col}34*${col}30` };
    worksheet.getCell(`${col}38`).numFmt = '₩#,##0';
    applyDataCellStyle(worksheet.getCell(`${col}38`), 'right');
    
    // 39행: 월 전용면적당 지출비용 = (D37+D38)/D31
    worksheet.getCell(`${col}39`).value = { formula: `IFERROR((${col}37+${col}38)/${col}31,0)` };
    worksheet.getCell(`${col}39`).numFmt = '₩#,##0';
    applyDataCellStyle(worksheet.getCell(`${col}39`), 'right');
    
    // 임대기준 조정
    // 40행: 보증금 = D32
    worksheet.getCell(`${col}40`).value = { formula: `${col}32` };
    worksheet.getCell(`${col}40`).numFmt = '₩#,##0';
    applyDataCellStyle(worksheet.getCell(`${col}40`), 'right');
    
    // 41행: 렌트프리 (사용자 입력)
    setNumericCell(worksheet, `${col}41`, 0, '0', 'center');
    
    // 42행: 평균 임대료 = D33-((D33*D41)/12)
    worksheet.getCell(`${col}42`).value = { formula: `${col}33-((${col}33*${col}41)/12)` };
    worksheet.getCell(`${col}42`).numFmt = '₩#,##0';
    applyDataCellStyle(worksheet.getCell(`${col}42`), 'right');
    
    // 43행: 관리비 = D34
    worksheet.getCell(`${col}43`).value = { formula: `${col}34` };
    worksheet.getCell(`${col}43`).numFmt = '₩#,##0';
    applyDataCellStyle(worksheet.getCell(`${col}43`), 'right');
    
    // 44행: NOC = ((D42+D43)*(D30/D31))
    worksheet.getCell(`${col}44`).value = { formula: `IFERROR(((${col}42+${col}43)*(${col}30/${col}31)),0)` };
    worksheet.getCell(`${col}44`).numFmt = '₩#,##0';
    applyDataCellStyle(worksheet.getCell(`${col}44`), 'center');
    
    // 예상비용
    // 46행: 보증금 = D40*D30
    worksheet.getCell(`${col}46`).value = { formula: `${col}40*${col}30` };
    worksheet.getCell(`${col}46`).numFmt = '₩#,##0';
    applyDataCellStyle(worksheet.getCell(`${col}46`), 'right');
    
    // 47행: 평균 월 임대료 = D42*D30
    worksheet.getCell(`${col}47`).value = { formula: `${col}42*${col}30` };
    worksheet.getCell(`${col}47`).numFmt = '₩#,##0';
    applyDataCellStyle(worksheet.getCell(`${col}47`), 'right');
    
    // 48행: 평균 월 관리비 = D43*D30
    worksheet.getCell(`${col}48`).value = { formula: `${col}43*${col}30` };
    worksheet.getCell(`${col}48`).numFmt = '₩#,##0';
    applyDataCellStyle(worksheet.getCell(`${col}48`), 'right');
    
    // 49행: 월 (임대료 + 관리비) = D47+D48
    worksheet.getCell(`${col}49`).value = { formula: `${col}47+${col}48` };
    worksheet.getCell(`${col}49`).numFmt = '₩#,##0';
    applyDataCellStyle(worksheet.getCell(`${col}49`), 'center');
    
    // 50행: 연 실제 부담 고정금액 = D49*12
    worksheet.getCell(`${col}50`).value = { formula: `${col}49*12` };
    worksheet.getCell(`${col}50`).numFmt = '₩#,##0';
    applyDataCellStyle(worksheet.getCell(`${col}50`), 'center');
}

// 데이터 셀 설정 헬퍼 함수
function setDataCell(worksheet, cellAddress, value) {
    const cell = worksheet.getCell(cellAddress);
    cell.value = value;
    applyDataCellStyle(cell);
}

// 숫자 셀 설정 헬퍼 함수
function setNumericCell(worksheet, cellAddress, value, format, alignment = 'center') {
    const cell = worksheet.getCell(cellAddress);
    cell.value = value;
    cell.numFmt = format;
    applyDataCellStyle(cell, alignment);
}

// 데이터 셀 스타일 적용
function applyDataCellStyle(cell, alignment = 'center') {
    cell.font = { name: 'Noto Sans KR', size: 9 };
    cell.alignment = { horizontal: alignment, vertical: 'middle' };
    cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };
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