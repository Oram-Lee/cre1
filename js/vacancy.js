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
       
       // 띄어쓰기 제거
       if (text.includes(' ')) {
           variations.push(text.replace(/\s+/g, ''));
       }
       
       // 띄어쓰기 추가 (한글+영문 사이)
       if (!text.includes(' ')) {
           // 그랜드센트럴 → 그랜드 센트럴
           let spaced = text.replace(/([가-힣])([A-Z])/g, '$1 $2');
           // 더미SQARE → 더미 SQARE
           spaced = spaced.replace(/([가-힣])([A-Z]{2,})/g, '$1 $2');
           // LG광화문 → LG 광화문
           spaced = spaced.replace(/([A-Z]+)([가-힣])/g, '$1 $2');
           
           if (spaced !== text) {
               variations.push(spaced);
           }
       }
       
       // 특수문자 처리
       if (text.includes('-') || text.includes('_')) {
           variations.push(text.replace(/[-_]/g, ' '));
           variations.push(text.replace(/[-_]/g, ''));
       }
       
       return [...new Set(variations)]; // 중복 제거
   }
   
   // 검색어 변형 생성
   const searchVariations = generateSearchVariations(companyBuildingName);
   const fallbackVariations = generateSearchVariations(buildingName);
   
   // 모든 변형을 fallback으로 결합 (중복 제거)
   const allVariations = [...new Set([...searchVariations, ...fallbackVariations])];
   
   // URL 파라미터 구성
   const searchTerm = encodeURIComponent(searchVariations[0]);
   const fallbackTerms = allVariations.slice(1).map(v => encodeURIComponent(v)).join('|');
   
   console.log('PDF 검색 변형:', {
       회사표기: companyBuildingName,
       시스템표기: buildingName,
       검색변형: searchVariations,
       전체변형: allVariations
   });
   
   // PDF 뷰어 URL 구성
   const pdfViewerUrl = `https://oram-lee.github.io/cremap/pdf-viewer.html?file=${pdfFile}&search=${searchTerm}&fallback=${fallbackTerms}`;
   
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
