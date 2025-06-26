// ===== LG Comp List 메인 함수 =====

// 메인 export 함수
async function generateExcelLG() {
    console.log('=== generateExcelLG 함수 시작 ===');
    console.log('선택된 빌딩 수:', selectedBuildings.length);
    console.log('선택된 빌딩:', selectedBuildings);
    
    try {
        // 1. 빌딩 개수 검증
        if (!LG_UTILS.validateBuildingCount(selectedBuildings.length)) {
            return;
        }
        
        // 2. 기본값 설정
        const companyName = 'LG CNS';
        const defaultTitle = '단기임차 가능 공간';
        
        // 3. 로딩 표시 (옵션)
        showLoadingMessage('LG Comp List를 생성하는 중...');
        
        // 4. ExcelJS 워크북 생성
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('COMP'); // LG 템플릿은 'COMP' 시트
        
        // 5. 템플릿 생성 (window 함수 사용)
        console.log('템플릿 생성 시작...');
        console.log('createLGTemplate 존재 여부:', typeof window.createLGTemplate);
        
        // 제목 설정
        const reportTitle = `[임차제안] ${selectedBuildings.map(b => b.name).join(', ')}`;
        
        window.createLGTemplate(workbook, worksheet, selectedBuildings, companyName, reportTitle);
        console.log('템플릿 생성 완료');
        
        // 6. 빌딩 데이터 입력 (window 함수 사용)
        console.log('빌딩 데이터 입력 시작...');
        console.log('fillBuildingDataLG 존재 여부:', typeof window.fillBuildingDataLG);
        selectedBuildings.forEach((building, index) => {
            if (index < 6) { // 최대 6개
                console.log(`빌딩 ${index + 1} 데이터 입력 중:`, building.name);
                window.fillBuildingDataLG(worksheet, building, index);
            }
        });
        console.log('빌딩 데이터 입력 완료');
        
        // 7. 스타일 적용 (window 함수 사용)
        window.applyLGStyles(worksheet);
        
        // 9. 인쇄 설정 (window 함수 사용)
        window.applyPrintSettings(worksheet);
        
        // 10. 검증 (옵션)
        const validation = validateWorksheet(worksheet);
        if (!validation.isValid) {
            console.warn('워크시트 검증 경고:', validation.warnings);
        }
        
        // 11. 파일 저장
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { 
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
        });
        
        const fileName = `LG_CompList_${LG_UTILS.getCurrentDate()}.xlsx`;
        saveAs(blob, fileName);
        
        // 12. 완료 메시지
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
    
    // 스타일 검증
    if (window.validateStyles) {
        const styleValidation = window.validateStyles(worksheet);
        if (!styleValidation.isValid) {
            warnings.push(...styleValidation.errors);
        }
    }
    
    // 수식 검증
    selectedBuildings.forEach((building, index) => {
        if (index < 5 && window.validateFormulas) {
            const col = String.fromCharCode(68 + index);
            const formulaValidation = window.validateFormulas(worksheet, col);
            if (!formulaValidation.isValid) {
                warnings.push(...formulaValidation.errors);
            }
        }
    });
    
    return {
        isValid: isValid && warnings.length === 0,
        warnings: warnings
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

// 완료 메시지 표시
function showCompletionMessage(buildingCount) {
    const message = `✅ LG Comp List 생성 완료!\n\n` +
        `📊 빌딩 ${buildingCount}개의 정보가 입력되었습니다.\n\n` +
        `📝 추가 입력 필요 항목:\n` +
        `• 로고 이미지 (B5 셀)\n` +
        `• 빌딩 외관 이미지 (D5:H5)\n` +
        `• 임차 제안 정보 (층수, 입주시기, 거래유형)\n` +
        `• 임대 기준 (보증금, 임대료, 관리비)\n` +
        `• 렌트프리 개월 수\n\n` +
        `💡 입력한 정보에 따라 예상비용이 자동 계산됩니다.\n` +
        `📌 모든 텍스트는 LG Smart Regular 폰트로 설정되었습니다.`;
    
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
    console.log('보고서 제목: 단기임차 가능 공간');
    
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
    
    // 샘플 빌딩 데이터
    const sampleBuildings = [{
        name: '샘플빌딩1',
        address: '서울시 강남구 테헤란로 123',
        addressJibun: '서울시 강남구 역삼동 123-45',
        station: '강남역',
        floors: 'B5~20F',
        completionYear: '2020',
        dedicatedRate: 72.5,
        baseFloorArea: 1650.5,
        baseFloorAreaPy: 500,
        baseFloorAreaDedicated: 1200,
        baseFloorAreaDedicatedPy: 363
    }];
    
    // 템플릿 생성
    window.createLGTemplate(workbook, worksheet, sampleBuildings, 'LG CNS', '테스트 보고서');
    
    // 데이터 입력
    window.fillBuildingDataLG(worksheet, sampleBuildings[0], 4); // D열
    
    // 수식 적용
    window.applyLGFormulas(worksheet, 'D');
    
    // 스타일 적용
    window.applyLGStyles(worksheet);
    
    console.log('템플릿 미리보기 생성됨:', worksheet);
    return worksheet;
}

// 전역 함수로 등록 (index2.html에서 호출 가능)
window.generateExcelLG = generateExcelLG;
window.generateExcelLGDebug = generateExcelLGDebug;
window.previewLGTemplate = previewLGTemplate;

// 초기화 확인
console.log('building2-main.js 로드 완료');
console.log('사용 가능한 함수: generateExcelLG(), generateExcelLGDebug(), previewLGTemplate()');
