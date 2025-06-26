// ===== LG Comp List 템플릿 설정 (실제 LG 템플릿 기반) =====
const LG_TEMPLATE_CONFIG = {
    // 시트명
    sheetName: 'COMP',
    
    // 최대 행/열
    maxRow: 109,
    maxColumn: 22,  // V열까지
    
    // 빌딩 시작 열 (6개 빌딩)
    buildingColumns: ['E', 'H', 'K', 'N', 'Q', 'T'],
    
    // 각 빌딩당 열 구성
    columnsPerBuilding: 3,  // 각 빌딩당 3열 사용
    
    // 헤더 정보 (1-4행)
    headerInfo: {
        1: '[임차제안 제목을 입력하세요. ]',
        2: '- 규모: 전용 0000PY 이상',
        3: '- 계약기간: 2025.00.00~2025.00.00 (00개월 간)',
        4: '- 위치: 0000역 인근'
    },
    
    // 섹션별 행 정보
    sections: {
        header: { start: 1, end: 4 },
        location: { start: 6, end: 8 },
        buildingImage: { start: 9, end: 17 },
        basicInfo: { start: 18, end: 25 },
        bondAnalysis: { start: 26, end: 32 },
        vacancy: { start: 33, end: 39 },
        proposal: { start: 40, end: 55 },
        construction: { start: 56, end: 58 },
        parking: { start: 59, end: 62 },
        floorPlan: { start: 63, end: 72 },
        remarks: { start: 73, end: 83 },
        notes: { start: 84, end: 85 }
    },
    
    // B열 라벨 정보
    labels: {
        // 기초정보
        18: '주   소',
        19: '위   치',
        20: '준공일',
        21: '규  모',
        22: '연면적',
        23: '기준층 전용면적',
        24: '전용률',
        25: '대지면적',
        // 채권분석
        26: '소유자 (임대인)',
        27: '채권담보 설정여부',
        28: '공동담보 총 대지지분',
        29: '선순위 담보 총액',
        30: '공시지가 대비 담보율',  // C열
        31: '개별공시지가(25년 1월 기준)',
        32: '토지가격 적용',  // C열
        // 제안
        40: '계약기간',
        41: '입주가능 시기',
        42: '제안 층',
        43: '전용면적',
        44: '임대면적',
        45: '보증금',
        46: '임대료',
        47: '관리비',
        48: '실질 임대료(RF만 반영)',
        49: '연간 무상임대 (R.F)',
        50: '보증금',
        51: '월 임대료',
        52: '월 관리비',
        53: '관리비 내역',
        54: '월납부액',
        55: '(21개월 기준) 총 납부 비용',
        56: '인테리어 기간 (F.O)',
        57: '인테리어지원금 (T.I)',
        // 주차현황
        59: '총 주차대수',
        60: '무료주차 조건(임대면적)',
        61: '무료주차 제공대수',
        62: '유료주차(VAT별도)',
        // 기타
        63: '평면도',
        73: '특이사항'
    },
    
    // 공실 현황 테이블 구조 (33-39행)
    vacancyTable: {
        header: 33,  // 층/전용/임대
        dataStart: 34,
        dataEnd: 38,
        total: 39    // 소계
    },
    
    // 수식 정보
    formulas: {
        // 담보율 계산
        담보율: (col) => `=${col}29/${col}32`,
        토지가격: (col) => `=${col}31*${String.fromCharCode(col.charCodeAt(0)+2)}25`,
        
        // 면적 소계
        전용소계: (col) => `=SUM(${col}34:${col}38)`,
        임대소계: (col) => `=SUM(${String.fromCharCode(col.charCodeAt(0)+1)}34:${String.fromCharCode(col.charCodeAt(0)+1)}38)`,
        
        // 실질 임대료 (RF 반영)
        실질임대료: (col) => `=${col}46*(12-${col}49)/12`,
        
        // 비용 계산
        보증금총액: (col) => `=${col}45*${col}44`,
        월임대료총액: (col) => `=${col}46*${col}44`,
        월관리비총액: (col) => `=${col}47*${col}44`,
        월납부액: (col) => `=${col}51+${col}52`,
        총납부비용21개월: (col) => `=${col}54*21`,
        
        // 주차대수 계산
        주차대수비율: (col) => `=${col}44/${col}60`
    },
    
    // 병합 셀 주요 영역
    majorMergedCells: {
        // 헤더
        title: 'A1:V4',
        location: 'A6:D6',
        proposal: 'A7:D8',
        
        // 이미지 영역
        buildingExterior: 'A9:D17',
        
        // 빌딩별 이미지 영역
        buildingImages: [
            'E9:G17',
            'H9:J17',
            'K9:M17',
            'N9:P17',
            'Q9:S17',
            'T9:V17'
        ],
        
        // 섹션 타이틀
        basicInfo: 'A18:A25',
        bondAnalysis: 'A26:A32',
        vacancy: 'A33:D39',
        proposalTitle: 'A40:A44',
        baseRent: 'A45:A47',
        realRent: 'A48:A49',
        costReview: 'A50:A55',
        construction: 'A56:A58',
        parking: 'A59:A62',
        others: 'A63:A83',
        
        // 평면도 영역
        floorPlans: [
            'B63:D72',
            'E63:G72',
            'H63:J72',
            'K63:M72',
            'N63:P72',
            'Q63:S72',
            'T63:V72'
        ],
        
        // 특이사항 영역
        remarks: [
            'B73:D83',
            'E73:G83',
            'H73:J83',
            'K73:M83',
            'N73:P83',
            'Q73:S83',
            'T73:V83'
        ]
    },
    
    // 데이터 매핑 (buildings.json 필드 → 엑셀 행)
    dataMapping: {
        18: 'address',           // 주소
        19: 'station',           // 위치
        20: 'completionYear',    // 준공일
        21: 'floors',            // 규모
        22: 'grossFloorAreaPy',  // 연면적
        23: 'baseFloorAreaDedicatedPy', // 기준층 전용면적
        24: 'dedicatedRate',     // 전용률
        25: 'landAreaPy',        // 대지면적
        26: '',                  // 소유자 (수동입력)
        27: '',                  // 채권담보 설정여부 (수동입력)
        31: '',                  // 개별공시지가 (수동입력)
        
        // 제안 관련
        40: '',                  // 계약기간 (수동입력)
        41: '',                  // 입주가능 시기 (수동입력)
        42: '',                  // 제안 층 (수동입력)
        43: '',                  // 전용면적 (공실 테이블에서 계산)
        44: '',                  // 임대면적 (공실 테이블에서 계산)
        45: 'depositPy',         // 보증금 (평당)
        46: 'rentPricePy',       // 임대료 (평당)
        47: 'managementFeePy',   // 관리비 (평당)
        49: '',                  // R.F 개월수 (수동입력, 기본값 0)
        
        // 주차
        59: 'parkingSpace',      // 총 주차대수
        60: '',                  // 무료주차 조건 (수동입력)
        62: 'parkingFee'         // 유료주차비
    },
    
    // 스타일 설정
    styles: {
        // 전체 기본 폰트
        defaultFont: {
            name: '맑은 고딕',
            size: 10
        },
        
        // 타이틀 스타일
        titleStyle: {
            font: { size: 14, bold: true },
            alignment: { horizontal: 'left', vertical: 'top' }
        },
        
        // 섹션 헤더 스타일
        sectionHeader: {
            fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0E0E0' } },
            font: { bold: true },
            alignment: { horizontal: 'center', vertical: 'middle' }
        },
        
        // 데이터 셀 스타일
        dataCell: {
            alignment: { horizontal: 'center', vertical: 'middle' },
            border: {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            }
        },
        
        // 수식 결과 스타일
        formulaCell: {
            font: { color: { argb: 'FF0000FF' } },
            alignment: { horizontal: 'center', vertical: 'middle' }
        }
    },
    
    // 기본값
    defaultValues: {
        49: 0,  // R.F 개월수 기본값
        56: '미제공',  // 인테리어 기간
        57: '미제공'   // 인테리어 지원금
    }
};

// 유틸리티 함수
const LG_UTILS = {
    // 빌딩 인덱스로 시작 열 가져오기
    getBuildingStartColumn: (buildingIndex) => {
        return LG_TEMPLATE_CONFIG.buildingColumns[buildingIndex];
    },
    
    // 빌딩 개수 검증
    validateBuildingCount: (count) => {
        if (count === 0) {
            alert('선택된 빌딩이 없습니다.');
            return false;
        }
        if (count > 6) {
            alert('최대 6개까지만 비교할 수 있습니다.');
            return false;
        }
        return true;
    },
    
    // 열 문자를 인덱스로 변환
    getColumnIndex: (letter) => {
        let index = 0;
        for (let i = 0; i < letter.length; i++) {
            index = index * 26 + (letter.charCodeAt(i) - 64);
        }
        return index;
    },
    
    // 인덱스를 열 문자로 변환
    getColumnLetter: (index) => {
        let letter = '';
        while (index > 0) {
            index--;
            letter = String.fromCharCode(65 + (index % 26)) + letter;
            index = Math.floor(index / 26);
        }
        return letter;
    },
    
    // 날짜 포맷
    getCurrentDate: () => {
        return new Date().toISOString().split('T')[0];
    }
};
