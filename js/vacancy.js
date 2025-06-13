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
    
    // 🆕 핵심 숫자 추출 함수 (주소의 특징적인 숫자)
    function extractKeyNumbers(text) {
        if (!text) return [];
        const numbers = [];
        
        // 번지수 패턴 (279, 279-9, 731-1 등)
        const bunjiMatches = text.match(/\b(\d{2,4}(-\d{1,3})?)\b/g);
        if (bunjiMatches) numbers.push(...bunjiMatches);
        
        // 도로명 번지 (노해로 464에서 464)
        const roadNumMatch = text.match(/[가-힣]+로\s*(\d+)/);
        if (roadNumMatch) numbers.push(roadNumMatch[1]);
        
        // "동" 뒤의 숫자 (2가, 3가 등)
        const gaMatch = text.match(/(\d)가/);
        if (gaMatch) numbers.push(gaMatch[1] + '가');
        
        return [...new Set(numbers)];
    }
    
    // 🆕 주소 정규화 함수 (비교를 위한 표준화)
    function normalizeAddress(address) {
        if (!address) return '';
        
        return address
            .replace(/서울특별시|서울시/g, '')
            .replace(/\s+/g, ' ')
            .replace(/^\s+|\s+$/g, '')
            .toLowerCase();
    }
    
    // 🆕 스마트 매칭 스코어 계산
    function calculateMatchScore(text1, text2) {
        const norm1 = normalizeAddress(text1);
        const norm2 = normalizeAddress(text2);
        
        // 완전 일치
        if (norm1 === norm2) return 100;
        
        // 포함 관계
        if (norm1.includes(norm2) || norm2.includes(norm1)) return 80;
        
        // 숫자 일치도
        const numbers1 = extractKeyNumbers(text1);
        const numbers2 = extractKeyNumbers(text2);
        const commonNumbers = numbers1.filter(n => numbers2.includes(n));
        
        if (commonNumbers.length > 0) {
            return 50 + (commonNumbers.length * 10);
        }
        
        return 0;
    }
    
    // 검색어 변형 생성 함수 (개선된 버전)
    function generateSearchVariations(text) {
        if (!text) return [];
        
        const variations = [];
        
        // 원본
        variations.push(text);
        
        // 🆕 주소 형식 감지 및 처리
        const isAddress = text.match(/[시구동로]|번지|\d{2,}/);
        
        if (isAddress) {
            // 전체 주소 그대로
            variations.push(text);
            
            // 시/도 제거
            variations.push(text.replace(/서울특별시\s*|서울시\s*/g, ''));
            
            // 구 단위 압축 (노원구 노원동 → 노원구노원동)
            if (text.includes('구') && text.includes('동')) {
                variations.push(text.replace(/구\s+/g, '구'));
            }
            
            // 동+번지만 추출
            const dongBunjiMatch = text.match(/([가-힣]+동\d*가?)\s*(\d+-?\d*)/);
            if (dongBunjiMatch) {
                variations.push(`${dongBunjiMatch[1]} ${dongBunjiMatch[2]}`);
                variations.push(`${dongBunjiMatch[1]}${dongBunjiMatch[2]}`);
            }
            
            // 도로명만 추출
            const roadMatch = text.match(/([가-힣]+로)\s*(\d+)/);
            if (roadMatch) {
                variations.push(`${roadMatch[1]} ${roadMatch[2]}`);
                variations.push(`${roadMatch[1]}${roadMatch[2]}`);
            }
        }
        
        // 띄어쓰기 변형
        if (text.includes(' ')) {
            variations.push(text.replace(/\s+/g, ''));
        } else {
            // 패턴별 띄어쓰기 추가
            const patterns = [
                /([가-힣]+)([A-Z])/g,          // 한글+영문
                /([A-Z]+)([가-힣])/g,          // 영문+한글
                /([가-힣])([0-9])/g,           // 한글+숫자
                /([0-9])([가-힣])/g,           // 숫자+한글
                /([A-Za-z]+)(\d+)/g,           // 영문+숫자
                /(빌딩|타워|센터|사옥)$/g      // 건물 유형 앞
            ];
            
            patterns.forEach(pattern => {
                const spaced = text.replace(pattern, '$1 $2');
                if (spaced !== text) variations.push(spaced);
            });
        }
        
        // 특수문자 처리
        if (text.match(/[-_.()/]/)) {
            variations.push(text.replace(/[-_.()]/g, ' ').replace(/\s+/g, ' ').trim());
            variations.push(text.replace(/[-_.()]/g, ''));
        }
        
        // 건물 키워드 제거
        const buildingKeywords = ['빌딩', '타워', 'Tower', 'Building', '사옥', '센터', 'Center', '오피스'];
        buildingKeywords.forEach(keyword => {
            if (text.includes(keyword)) {
                const removed = text.replace(new RegExp(keyword + '$'), '').trim();
                if (removed && removed.length > 2) variations.push(removed);
            }
        });
        
        // "동" 처리
        if (text.includes('동')) {
            // 동 제거 (신공덕동 → 신공덕)
            if (text.endsWith('동')) {
                variations.push(text.slice(0, -1));
            }
            // 동 뒤 띄어쓰기 (신공덕동아이파크 → 신공덕 아이파크)
            const dongSpaced = text.replace(/동([가-힣A-Z])/g, '동 $1');
            if (dongSpaced !== text) variations.push(dongSpaced);
        }
        
        return [...new Set(variations)];
    }
    
    // 주소에서 유용한 정보 추출 (완전히 개선된 버전)
    function extractAddressInfo(building) {
        const addressVariations = [];
        
        if (!building) return [];
        
        // 🆕 모든 주소 형태를 수집
        const allAddresses = [];
        
        if (building.address) allAddresses.push(building.address);
        if (building.addressJibun) allAddresses.push(building.addressJibun);
        
        allAddresses.forEach(addr => {
            // 전체 주소
            addressVariations.push(addr);
            
            // 도로명 주소 처리
            if (addr.includes('로')) {
                // 전체 도로명 주소
                addressVariations.push(addr);
                
                // 시/구 제거 버전
                const simplified = addr.replace(/서울특별시\s*|서울시\s*/g, '').trim();
                addressVariations.push(simplified);
                
                // 도로명+번지만
                const roadMatch = addr.match(/([가-힣]+로\s*\d+)/);
                if (roadMatch) addressVariations.push(roadMatch[1]);
            }
            
            // 지번 주소 처리
            if (addr.includes('동')) {
                // 전체 지번 주소
                addressVariations.push(addr);
                
                // 구 압축 형태 (노원구 노원동 → 노원구노원동)
                if (addr.includes('구') && addr.includes('동')) {
                    addressVariations.push(addr.replace(/구\s+/g, '구'));
                }
                
                // 동+번지만
                const dongMatch = addr.match(/([가-힣]+동\d*가?)\s*(\d+-?\d*)/);
                if (dongMatch) {
                    addressVariations.push(`${dongMatch[1]} ${dongMatch[2]}`);
                    addressVariations.push(`${dongMatch[1]}${dongMatch[2]}`);
                    
                    // 동만
                    addressVariations.push(dongMatch[1]);
                    
                    // 번지만 (3자리 이상)
                    if (dongMatch[2].length >= 3) {
                        addressVariations.push(dongMatch[2]);
                    }
                }
            }
            
            // 구 정보
            const guMatch = addr.match(/([가-힣]+구)(?=\s|$)/);
            if (guMatch) addressVariations.push(guMatch[1]);
            
            // 🆕 핵심 숫자들
            const keyNumbers = extractKeyNumbers(addr);
            keyNumbers.forEach(num => {
                if (num.length >= 3) addressVariations.push(num);
            });
        });
        
        // 역 정보
        if (building.station) {
            const stations = building.station.match(/([가-힣]+역)/g);
            if (stations) {
                stations.forEach(station => {
                    addressVariations.push(station);
                    addressVariations.push(station.replace('역', ''));
                });
            }
        }
        
        // 중복 제거 및 길이 필터링
        return [...new Set(addressVariations)].filter(v => v && v.length > 3);
    }
    
    // 🆕 PDF별 특수 처리
    function applyPdfSpecificLogic(pdfFile, companyBuildingName, currentBuilding, variations) {
        // GM PDF: 도로명 주소를 빌딩명으로 사용
        if (pdfFile === 'GM.pdf' && currentBuilding) {
            if (currentBuilding.address) {
                variations.unshift(currentBuilding.address);
                
                // "서울특별시" 제거 버전도 추가
                const simpleAddr = currentBuilding.address.replace(/서울특별시\s*/g, '').trim();
                variations.unshift(simpleAddr);
            }
        }
        
        // SVS PDF: 지번 주소를 빌딩명으로 사용
        if (pdfFile === 'SVS.pdf' && currentBuilding) {
            if (currentBuilding.addressJibun) {
                // 동+번지 형태 우선
                const dongBunjiMatch = currentBuilding.addressJibun.match(/([가-힣]+동\d*가?)\s*(\d+-?\d*)/);
                if (dongBunjiMatch) {
                    variations.unshift(`${dongBunjiMatch[1]} ${dongBunjiMatch[2]}`);
                }
            }
        }
        
        return variations;
    }
    
    // 메인 로직
    let companyVariations = generateSearchVariations(companyBuildingName);
    const systemVariations = generateSearchVariations(buildingName);
    const addressInfo = extractAddressInfo(currentBuilding);
    
    // PDF별 특수 처리 적용
    companyVariations = applyPdfSpecificLogic(pdfFile, companyBuildingName, currentBuilding, companyVariations);
    
    // 🆕 스마트 정렬: 매칭 점수가 높은 순으로 정렬
    const scoredVariations = [];
    const allCandidates = [...new Set([...companyVariations, ...systemVariations, ...addressInfo])];
    
    allCandidates.forEach(variation => {
        let maxScore = 0;
        
        // 회사 빌딩명과의 점수
        maxScore = Math.max(maxScore, calculateMatchScore(variation, companyBuildingName));
        
        // 시스템 빌딩명과의 점수
        maxScore = Math.max(maxScore, calculateMatchScore(variation, buildingName));
        
        // 주소 정보와의 점수
        if (currentBuilding) {
            if (currentBuilding.address) {
                maxScore = Math.max(maxScore, calculateMatchScore(variation, currentBuilding.address));
            }
            if (currentBuilding.addressJibun) {
                maxScore = Math.max(maxScore, calculateMatchScore(variation, currentBuilding.addressJibun));
            }
        }
        
        scoredVariations.push({ variation, score: maxScore });
    });
    
    // 점수 순으로 정렬
    scoredVariations.sort((a, b) => b.score - a.score);
    
    // 상위 변형들 선택
    const topVariations = scoredVariations
        .filter(item => item.variation && item.variation.length > 3)
        .map(item => item.variation);
    
    // 첫 번째 검색어 (가장 높은 점수)
    const primarySearch = topVariations[0] || companyBuildingName || buildingName;
    
    // fallback 검색어들 (나머지 상위 20개)
    const fallbackVariations = topVariations.slice(1, 21);
    
    // 디버깅 정보
    console.log('PDF 검색 변형:', {
        PDF파일: pdfFile,
        회사표기: companyBuildingName,
        시스템표기: buildingName,
        현재빌딩주소: currentBuilding ? {
            도로명: currentBuilding.address,
            지번: currentBuilding.addressJibun
        } : null,
        점수순변형: topVariations.slice(0, 10),
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
