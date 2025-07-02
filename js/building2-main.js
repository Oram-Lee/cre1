// ===== LG Comp List 메인 함수 =====

// 메인 export 함수 (최대 10개 빌딩 지원)
async function generateExcelLG() {
    console.log('=== generateExcelLG 함수 시작 ===');
    console.log('선택된 빌딩 수:', selectedBuildings.length);
    console.log('선택된 빌딩:', selectedBuildings);
    
    try {
        // 1. 빌딩 개수 검증 (최대 10개)
        if (!LG_UTILS.validateBuildingCount(selectedBuildings.length)) {
            return;
        }
        
        // 2. 기본값 설정
        const companyName = 'LG CNS';
        const defaultTitle = `임차제안 제목을 입력하세요.`;  // 요구사항3
        
        // 3. 로딩 표시
        showLoadingMessage('LG Comp List를 생성하는 중...');
        
        // 4. ExcelJS 워크북 생성
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('COMP'); // LG 템플릿은 'COMP' 시트
        
        // 5. 템플릿 생성
        console.log('템플릿 생성 시작...');
        console.log('createLGTemplate 존재 여부:', typeof window.createLGTemplate);
        
        // 제목 설정 - 빌딩명들을 포함한 제목
        const buildingNames = selectedBuildings.map(b => b.name).join(', ');
        const reportTitle = buildingNames.length > 50 ? 
            `임차제안: ${buildingNames.substring(0, 47)}...` : 
            `임차제안: ${buildingNames}`;
        
        window.createLGTemplate(workbook, worksheet, selectedBuildings, companyName, reportTitle);
        console.log('템플릿 생성 완료');
        
        // 6. 빌딩 데이터 입력 (최대 10개)
        console.log('빌딩 데이터 입력 시작...');
        console.log('fillBuildingDataLG 존재 여부:', typeof window.fillBuildingDataLG);
        
        selectedBuildings.forEach((building, index) => {
            if (index < 10) { // 최대 10개
                console.log(`빌딩 ${index + 1} 데이터 입력 중:`, building.name);
                window.fillBuildingDataLG(worksheet, building, index);
            }
        });
        console.log('빌딩 데이터 입력 완료');
        
        // 7. 스타일 적용 (요구사항2: LG스마트체 Regular)
        console.log('스타일 적용 시작...');
        window.applyLGStyles(worksheet);
        console.log('스타일 적용 완료');
        
        // 8. 인쇄 설정
        window.applyPrintSettings(worksheet);
        
        // 9. 검증
        const validation = validateWorksheet(worksheet);
        if (!validation.isValid) {
            console.warn('워크시트 검증 경고:', validation.warnings);
        }
        
        // 10. 파일 저장
        console.log('파일 저장 시작...');
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { 
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
        });
        
        const fileName = `LG_CompList_${selectedBuildings.length}개빌딩_${LG_UTILS.getCurrentDate()}.xlsx`;
        saveAs(blob, fileName);
        console.log('파일 저장 완료:', fileName);
        
        // 11. 완료 메시지
        hideLoadingMessage();
        showCompletionMessage(selectedBuildings.length);
        
    } catch (error) {
        console.error('LG Comp List 생성 오류:', error);
        hideLoadingMessage();
        alert('LG Comp List 생성 중 오류가 발생했습니다.\n' + error.message);
    }
}

// 워크시트 전체 검증
function validateWorksheet(worksheet) {
    const warnings = [];
    let isValid = true;
    
    // 템플릿 검증
    if (window.validateTemplate && !window.validateTemplate(worksheet)) {
        warnings.push('템플릿 구조가 올바르지 않습니다.');
        isValid = false;
    }
    
    // 스타일 검증 (요구사항2: LG스마트체)
    if (window.validateStyles) {
        const styleValidation = window.validateStyles(worksheet);
        if (!styleValidation.isValid) {
            warnings.push(...styleValidation.errors);
        }
    }
    
    // 빌딩 데이터 검증
    selectedBuildings.forEach((building, index) => {
        if (index < 10 && window.validateBuildingDataLG) {
            const dataValidation = window.validateBuildingDataLG(worksheet, index);
            if (!dataValidation.isValid) {
                warnings.push(...dataValidation.errors);
            }
        }
    });
    
    // 색상 요구사항 검증 (요구사항 6-9)
    const colorValidation = validateColorRequirements(worksheet);
    if (!colorValidation.isValid) {
        warnings.push(...colorValidation.errors);
    }
    
    return {
        isValid: isValid && warnings.length === 0,
        warnings: warnings
    };
}

// 색상 요구사항 검증 (요구사항 6-9)
function validateColorRequirements(worksheet) {
    const errors = [];
    let isValid = true;
    
    try {
        // 요구사항6: A6 주황 80% 밝게
        const a6 = worksheet.getCell('A6');
        if (a6.fill && a6.fill.fgColor && a6.fill.fgColor.argb !== LG_TEMPLATE_CONFIG.colors.a6) {
            errors.push('A6: 주황 80% 밝게 색상이 적용되지 않았습니다.');
            isValid = false;
        }
        
        // 요구사항7, 8, 9: 빌딩별 색상 검증
        selectedBuildings.forEach((building, index) => {
            if (index < 10) {
                const col = LG_TEMPLATE_CONFIG.buildingColumns[index];
                
                // 요구사항7: 빌딩명 녹색 80% 밝게
                const nameCell = worksheet.getCell(`${col}6`);
                if (nameCell.fill && nameCell.fill.fgColor && 
                    nameCell.fill.fgColor.argb !== LG_TEMPLATE_CONFIG.colors.buildingName) {
                    errors.push(`${col}6: 녹색 80% 밝게 색상이 적용되지 않았습니다.`);
                    isValid = false;
                }
                
                // 요구사항8: 8행 파랑 80% 밝게
                const cell8 = worksheet.getCell(`${col}8`);
                if (cell8.fill && cell8.fill.fgColor && 
                    cell8.fill.fgColor.argb !== LG_TEMPLATE_CONFIG.colors.proposal) {
                    errors.push(`${col}8: 파랑 80% 밝게 색상이 적용되지 않았습니다.`);
                    isValid = false;
                }
                
                // 요구사항9: 7행 검정 35% 밝게
                const cell7 = worksheet.getCell(`${col}7`);
                if (cell7.fill && cell7.fill.fgColor && 
                    cell7.fill.fgColor.argb !== LG_TEMPLATE_CONFIG.colors.location) {
                    errors.push(`${col}7: 검정 35% 밝게 색상이 적용되지 않았습니다.`);
                    isValid = false;
                }
            }
        });
        
    } catch (error) {
        errors.push('색상 검증 중 오류 발생: ' + error.message);
        isValid = false;
    }
    
    return {
        isValid: isValid,
        errors: errors
    };
}

// 로딩 메시지 표시
function showLoadingMessage(message) {
    // 기존 로딩 오버레이가 있으면 사용
    const loadingOverlay = document.getElementById('loading-overlay');
    if (loadingOverlay) {
        const loadingText = loadingOverlay.querySelector('.loading-text');
        if (loadingText) {
            loadingText.textContent = message;
        }
        loadingOverlay.style.display = 'flex';
    }
}

// 로딩 메시지 숨기기
function hideLoadingMessage() {
    const loadingOverlay = document.getElementById('loading-overlay');
    if (loadingOverlay) {
        loadingOverlay.style.display = 'none';
    }
}

// 완료 메시지 표시 (요구사항 반영)
function showCompletionMessage(buildingCount) {
    const message = `✅ LG Comp List 생성 완료!\n\n` +
        `📊 빌딩 ${buildingCount}개의 정보가 입력되었습니다.\n\n` +
        `🎯 적용된 요구사항:\n` +
        `• 최대 10개 빌딩 지원\n` +
        `• 모든 폰트: LG Smart Regular\n` +
        `• A1: "임차제안 제목을 입력하세요."\n` +
        `• A6: 주황 80% 밝게\n` +
        `• 빌딩명(6행): 녹색 80% 밝게\n` +
        `• 8행: 파랑 80% 밝게\n` +
        `• 7행: 검정 35% 밝게\n\n` +
        `📝 추가 입력 필요 항목:\n` +
        `• 로고 이미지 (B5 셀)\n` +
        `• 빌딩 외관 이미지\n` +
        `• 임차 제안 정보 (층수, 입주시기, 거래유형)\n` +
        `• 임대 기준 (보증금, 임대료, 관리비)\n` +
        `• 렌트프리 개월 수\n\n` +
        `💡 입력한 정보에 따라 예상비용이 자동 계산됩니다.`;
    
    alert(message);
}

// 빠른 실행을 위한 단축 함수
function quickGenerateLG() {
    // 기본값으로 빠르게 생성
    generateExcelLG();
}

// 디버그 모드 실행
async function generateExcelLGDebug() {
    console.log('=== LG Comp List 생성 시작 (디버그 모드) ===');
    console.log('선택된 빌딩:', selectedBuildings);
    console.log('회사명: LG CNS');
    console.log('최대 지원 빌딩 수: 10개');
    console.log('요구사항 적용:');
    console.log('- 폰트: LG Smart Regular');
    console.log('- A1: 임차제안 제목을 입력하세요.');
    console.log('- 색상 요구사항 6-9 적용');
    
    try {
        await generateExcelLG();
        console.log('=== LG Comp List 생성 완료 ===');
    } catch (error) {
        console.error('=== LG Comp List 생성 실패 ===', error);
        throw error;
    }
}

// 템플릿 미리보기 (개발용)
function previewLGTemplate() {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('COMP');
    
    // 샘플 빌딩 데이터 (최대 10개)
    const sampleBuildings = Array.from({length: 10}, (_, i) => ({
        name: `샘플빌딩${i + 1}`,
        address: `서울시 강남구 테헤란로 ${100 + i}`,
        addressJibun: `서울시 강남구 역삼동 ${100 + i}-${i + 1}`,
        station: '강남역',
        floors: 'B5~20F',
        completionYear: '2020',
        dedicatedRate: 72.5,
        baseFloorArea: 1650.5,
        baseFloorAreaPy: 500,
        baseFloorAreaDedicated: 1200,
        baseFloorAreaDedicatedPy: 363
    }));
    
    // 처음 3개만 사용하여 테스트
    const testBuildings = sampleBuildings.slice(0, 3);
    
    // 템플릿 생성
    window.createLGTemplate(workbook, worksheet, testBuildings, 'LG CNS', '테스트 보고서');
    
    // 데이터 입력
    testBuildings.forEach((building, index) => {
        window.fillBuildingDataLG(worksheet, building, index);
    });
    
    // 스타일 적용
    window.applyLGStyles(worksheet);
    
    console.log('템플릿 미리보기 생성됨:', worksheet);
    return worksheet;
}

// 성능 측정 함수
async function measureLGGenerationPerformance(buildingCount = 5) {
    console.log(`=== LG Comp List 성능 측정 시작 (${buildingCount}개 빌딩) ===`);
    
    const startTime = performance.now();
    
    try {
        // 샘플 데이터로 성능 측정
        const originalSelected = [...selectedBuildings];
        
        // 임시로 샘플 데이터 설정
        selectedBuildings.length = 0;
        for (let i = 0; i < buildingCount; i++) {
            selectedBuildings.push({
                id: i,
                name: `성능측정빌딩${i + 1}`,
                address: `서울시 강남구 테헤란로 ${100 + i}`,
                station: '강남역'
            });
        }
        
        await generateExcelLG();
        
        // 원본 데이터 복원
        selectedBuildings.length = 0;
        selectedBuildings.push(...originalSelected);
        
        const endTime = performance.now();
        const duration = endTime - startTime;
        
        console.log(`=== 성능 측정 완료 ===`);
        console.log(`처리 시간: ${duration.toFixed(2)}ms`);
        console.log(`빌딩당 평균 시간: ${(duration / buildingCount).toFixed(2)}ms`);
        
        return duration;
        
    } catch (error) {
        console.error('성능 측정 실패:', error);
        throw error;
    }
}

// 배치 처리 함수 (대량 데이터용)
async function batchGenerateLG(buildingGroups) {
    console.log(`배치 처리 시작: ${buildingGroups.length}개 그룹`);
    
    const results = [];
    
    for (let i = 0; i < buildingGroups.length; i++) {
        const group = buildingGroups[i];
        console.log(`그룹 ${i + 1}/${buildingGroups.length} 처리 중...`);
        
        try {
            // 선택된 빌딩 임시 교체
            const originalSelected = [...selectedBuildings];
            selectedBuildings.length = 0;
            selectedBuildings.push(...group);
            
            await generateExcelLG();
            
            results.push({
                groupIndex: i,
                success: true,
                buildingCount: group.length
            });
            
            // 원본 복원
            selectedBuildings.length = 0;
            selectedBuildings.push(...originalSelected);
            
        } catch (error) {
            console.error(`그룹 ${i + 1} 처리 실패:`, error);
            results.push({
                groupIndex: i,
                success: false,
                error: error.message
            });
        }
        
        // 다음 처리 전 잠시 대기 (브라우저 응답성 유지)
        await new Promise(resolve => setTimeout(resolve, 100));
    }
    
    console.log('배치 처리 완료:', results);
    return results;
}

// 전역 함수로 등록 (index.html에서 호출 가능)
window.generateExcelLG = generateExcelLG;
window.generateExcelLGDebug = generateExcelLGDebug;
window.previewLGTemplate = previewLGTemplate;
window.quickGenerateLG = quickGenerateLG;
window.measureLGGenerationPerformance = measureLGGenerationPerformance;
window.batchGenerateLG = batchGenerateLG;
window.validateColorRequirements = validateColorRequirements;

// 초기화 확인
console.log('building2-main.js 로드 완료 (최대 10개 빌딩 지원)');
console.log('사용 가능한 함수:');
console.log('- generateExcelLG() : LG Comp List 생성');
console.log('- generateExcelLGDebug() : 디버그 모드');
console.log('- previewLGTemplate() : 템플릿 미리보기');
console.log('- measureLGGenerationPerformance() : 성능 측정');
console.log('- batchGenerateLG() : 배치 처리');
console.log('적용된 요구사항: 최대 10개, LG스마트체, 색상 설정 등');