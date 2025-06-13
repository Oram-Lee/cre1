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
    
    // 현재 선택된 빌딩 정보 가져오기
    const currentBuilding = buildingsData.find(b => b.name === buildingName);
    
    // 검색어 변형 생성 함수
    function generateSearchVariations(text) {
        if (!text) return [];
        
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
        
        // 빌딩/타워/사옥 제거한 버전도 추가
        const buildingKeywords = ['빌딩', '타워', 'Tower', 'Building', '사옥', '센터', 'Center'];
        buildingKeywords.forEach(keyword => {
            if (text.includes(keyword)) {
                const removed = text.replace(keyword, '').trim();
                if (removed && removed.length > 2) variations.push(removed);
            }
        });
        
        // "동" 처리 (신공덕동 → 신공덕)
        if (text.includes('동 ')) {
            variations.push(text.replace(/동 /g, ' '));
        }
        if (text.endsWith('동')) {
            variations.push(text.substring(0, text.length - 1));
        }
        
        // PDF에서 자주 나타나는 패턴들
        // "신공덕 아이파크타워" 같은 경우
        if (text.includes('동') && !text.includes(' ')) {
            // 신공덕동아이파크타워 → 신공덕 아이파크타워
            const dongRemoved = text.replace(/([가-힣]+)동([가-힣A-Z])/g, '$1 $2');
            if (dongRemoved !== text) variations.push(dongRemoved);
        }
        
        return [...new Set(variations)]; // 중복 제거
    }
    
    // 주소에서 유용한 정보 추출
    function extractAddressInfo(building) {
        const addressVariations = [];
        
        if (building) {
            // 🆕 지번주소 전체 및 변형 추가
            if (building.addressJibun) {
                // 전체 지번주소
                addressVariations.push(building.addressJibun);
                
                // 시/구 제거한 간략 버전 (예: "성수동2가 279-9")
                const simplifiedMatch = building.addressJibun.match(/([가-힣0-9]+동[0-9가-힣]*\s*\d+(-\d+)?)/);
                if (simplifiedMatch) {
                    addressVariations.push(simplifiedMatch[1]);
                    
                    // 띄어쓰기 변형 (성수동2가 → 성수동 2가)
                    const spaced = simplifiedMatch[1].replace(/동(\d+가)/, '동 $1');
                    if (spaced !== simplifiedMatch[1]) {
                        addressVariations.push(spaced);
                    }
                }
                
                // 동+번지만 추출 (예: "성수동2가 279")
                const dongBunjiMatch = building.addressJibun.match(/(\S+동[0-9가-힣]*)\s*(\d+-?\d*)/);
                if (dongBunjiMatch) {
                    addressVariations.push(`${dongBunjiMatch[1]} ${dongBunjiMatch[2]}`);
                    addressVariations.push(`${dongBunjiMatch[1]}${dongBunjiMatch[2]}`); // 띄어쓰기 없는 버전
                }
            }
            
            // 기존 도로명 주소 처리
            if (building.address) {
                // 도로명 주소에서 번지 추출 - 수정!
                const roadNumMatch = building.address.match(/\d+로\s*(\d+)/);
                if (roadNumMatch) {
                    // 🆕 3자리 이상 숫자만 추가 (단일 숫자 제외)
                    if (roadNumMatch[1].length >= 3) {
                        addressVariations.push(roadNumMatch[1] + '번지');
                        addressVariations.push(roadNumMatch[1]);
                    }
                }
                
                // 도로명 추출 (예: "노해로 464")
                const roadMatch = building.address.match(/([가-힣]+로\s*\d+)/);
                if (roadMatch) {
                    addressVariations.push(roadMatch[1]);
                }
                
                // 구 정보 추출 (예: "노원구")
                const guMatch = building.address.match(/(\S+구)/);
                if (guMatch) addressVariations.push(guMatch[1]);
                
                // 동 정보 추출 (예: "상계동")
                const dongMatch = building.address.match(/(\S+동)(?=\s|$)/);
                if (dongMatch) {
                    addressVariations.push(dongMatch[1]);
                    // "동" 제거한 버전도 추가
                    addressVariations.push(dongMatch[1].replace('동', ''));
                }
            }
            
            // 역 정보 추출
            if (building.station) {
                const stationMatch = building.station.match(/(\S+역)/g);
                if (stationMatch) {
                    stationMatch.forEach(station => {
                        addressVariations.push(station);
                        // "역" 제거한 버전도 추가
                        addressVariations.push(station.replace('역', ''));
                    });
                }
            }
        }
        
        // 🆕 너무 짧은 검색어 제거 (3자 이하)
        return [...new Set(addressVariations)].filter(v => v && v.length > 3);
    }
    
    // 검색어 변형 생성
    const companyVariations = generateSearchVariations(companyBuildingName);
    const systemVariations = generateSearchVariations(buildingName);
    
    // 주소 정보 추가
    const addressInfo = extractAddressInfo(currentBuilding);
    
    // 🆕 특별한 경우 처리 - 회사별 특수 표기
    if (pdfFile === 'GM.pdf' && buildingName.includes('한화')) {
        // GM에서는 한화 빌딩을 다르게 표기할 수 있음
        companyVariations.push('한화 노원');
        companyVariations.push('한화노원');
        companyVariations.push('노원 한화');
    }
    
    // 모든 변형을 결합 (중복 제거, 빈 값과 괄호 제거)
    const allVariations = [...new Set([...companyVariations, ...systemVariations, ...addressInfo])]
        .filter(v => v && !v.startsWith('(') && !v.endsWith(')') && v.length > 3);
    
    // 첫 번째는 회사 표기를 우선
    const primarySearch = companyVariations[0];
    // 나머지는 fallback으로 (최대 15개까지)
    const fallbackVariations = allVariations
        .filter(v => v !== primarySearch)
        .slice(0, 15); // URL 길이 제한 때문에 15개까지
    
    console.log('PDF 검색 변형:', {
        회사표기: companyBuildingName,
        시스템표기: buildingName,
        회사변형: companyVariations,
        시스템변형: systemVariations,
        주소정보: addressInfo,
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
    console.log('URL 길이:', pdfViewerUrl.length);
    
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
