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
    
    // 검색어 변형 생성 함수
    function generateSearchVariations(text) {
        const variations = [];
        
        // 원본
        variations.push(text);
        
        // 띄어쓰기 처리
        if (text.includes(' ')) {
            // 띄어쓰기 제거
            variations.push(text.replace(/\s+/g, ''));
        } else {
            // 띄어쓰기 추가 - 더 많은 패턴
            // 한글+영문 (렉서스빌딩 → 렉서스 빌딩)
            let spaced = text.replace(/([가-힣]+)([A-Z])/g, '$1 $2');
            if (spaced !== text) variations.push(spaced);
            
            // 영문+한글 (LG광화문 → LG 광화문)
            spaced = text.replace(/([A-Z]+)([가-힣])/g, '$1 $2');
            if (spaced !== text) variations.push(spaced);
            
            // 한글 단어 사이 (그랜드센트럴 → 그랜드 센트럴)
            spaced = text.replace(/([가-힣])([A-Z][a-z])/g, '$1 $2');
            if (spaced !== text) variations.push(spaced);
            
            // 영문+숫자 (L7홍대 → L 7홍대)
            spaced = text.replace(/([A-Za-z]+)(\d+)/g, '$1 $2');
            if (spaced !== text) variations.push(spaced);
            
            // "빌딩", "타워" 앞에 띄어쓰기 추가
            if (text.includes('빌딩') && !text.includes(' 빌딩')) {
                variations.push(text.replace('빌딩', ' 빌딩'));
            }
            if (text.includes('타워') && !text.includes(' 타워')) {
                variations.push(text.replace('타워', ' 타워'));
            }
        }
        
        // 특수문자 처리
        if (text.includes('-') || text.includes('_') || text.includes('.')) {
            variations.push(text.replace(/[-_.]/g, ' ').trim());
            variations.push(text.replace(/[-_.]/g, '').trim());
        }
        
        // 괄호 제거
        if (text.includes('(') || text.includes(')')) {
            variations.push(text.replace(/\s*\([^)]*\)/g, '').trim());
        }
        
        // 빌딩/타워 제거한 버전도 추가
        const buildingKeywords = ['빌딩', '타워', 'Tower', 'Building'];
        buildingKeywords.forEach(keyword => {
            if (text.includes(keyword)) {
                const removed = text.replace(keyword, '').trim();
                if (removed) variations.push(removed);
            }
        });
        
        return [...new Set(variations)]; // 중복 제거
    }
    
    // 검색어 변형 생성
    const companyVariations = generateSearchVariations(companyBuildingName);
    const systemVariations = generateSearchVariations(buildingName);
    
    // 모든 변형을 결합 (중복 제거)
    const allVariations = [...new Set([...companyVariations, ...systemVariations])];
    
    // 첫 번째는 회사 표기를 우선
    const primarySearch = companyVariations[0];
    // 나머지는 fallback으로
    const fallbackVariations = allVariations.filter(v => v !== primarySearch);
    
    console.log('PDF 검색 변형:', {
        회사표기: companyBuildingName,
        시스템표기: buildingName,
        회사변형: companyVariations,
        시스템변형: systemVariations,
        전체변형: allVariations,
        주검색어: primarySearch,
        대체검색어: fallbackVariations
    });
    
    // URL 파라미터 구성
    const searchParam = encodeURIComponent(primarySearch);
    const fallbackParam = fallbackVariations.length > 0 
        ? '&fallback=' + encodeURIComponent(fallbackVariations.join('|'))
        : '';
    
    // PDF 뷰어 URL 구성
    const pdfViewerUrl = `https://oram-lee.github.io/cremap/pdf-viewer.html?file=${pdfFile}&search=${searchParam}${fallbackParam}`;
    
    console.log('PDF 뷰어 URL:', pdfViewerUrl);
    
    // 새 창에서 열기
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
        
        return `
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
        return `
            <div class="info-row">
                <span class="info-label">관련임대안내문</span>
                <span class="info-value">
                    <span class="badge bg-secondary">관련안내문없음</span>
                </span>
            </div>
        `;
    }
}
