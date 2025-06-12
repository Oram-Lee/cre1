// ===== 공실정보 연동 관련 함수들 =====

// 매칭 데이터 로드
async function loadMatchingData() {
    try {
        const response = await fetch('data/building-matching.json');
        buildingMatches = await response.json();
        console.log('매칭 데이터 로드 완료:', buildingMatches.metadata);
        console.log(`전체 ${buildingMatches.metadata.buildingSystemCount}개 중 ${buildingMatches.metadata.matchedCount}개 매칭`);
    } catch (error) {
        console.error('매칭 데이터 로드 실패:', error);
        // 매칭 데이터가 없어도 시스템은 정상 작동
    }
}

// PDF 뷰어 열기
function openPdfViewer(buildingName) {
    const selectElement = document.getElementById('pdfSelect');
    const selectedValue = selectElement.value;
    
    if (!selectedValue) {
        alert('임대안내문을 선택해주세요.');
        return;
    }
    
    const [pdfFile, companyBuildingName] = selectedValue.split('|');
    
    // PDF 뷰어 URL 구성
    // 빌딩명으로 검색하되, 회사별 표기가 다를 수 있으므로 fallback 검색어 추가
    const searchTerm = encodeURIComponent(companyBuildingName);
    const fallbackTerm = encodeURIComponent(buildingName);
    
    // 상대 경로로 PDF 뷰어 열기 (crema 폴더가 같은 레벨에 있다고 가정)
    const pdfViewerUrl = `../crema/pdf-viewer.html?file=${pdfFile}&search=${searchTerm}&fallback=${fallbackTerm}`;
    
    // 새 창에서 열기 (또는 팝업)
    const width = 1200;
    const height = 800;
    const left = (window.screen.width - width) / 2;
    const top = (window.screen.height - height) / 2;
    
    window.open(
        pdfViewerUrl,
        'pdfViewer',
        `width=${width},height=${height},left=${left},top=${top},scrollbars=yes,resizable=yes`
    );
}

// 공실정보 섹션 생성 (building.js의 showBuildingInfo 함수에서 사용)
function createVacancySection(building) {
    if (!buildingMatches) return '';
    
    const match = buildingMatches.matches.find(m => m.buildingSystemId === building.id);
    
    if (!match || match.vacancyMatches.length === 0) {
        return '';
    }
    
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
        
        return `
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
        return `
            <div class="info-row">
                <span class="info-label">공실정보</span>
                <span class="info-value">
                    <span class="badge bg-secondary">공실 없음</span>
                </span>
            </div>
        `;
    }
}