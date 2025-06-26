// ===== LG Comp List 템플릿 설정 =====
const LG_TEMPLATE_CONFIG = {
    // 열 너비 설정
    columnWidths: {
        'A': 2.6640625,
        'B': 13.21875,
        'C': 24.5546875,
        'D': 26.33203125,
        'E': 26.33203125,
        'F': 26.33203125,
        'G': 26.33203125,
        'H': 26.33203125
    },

    // 행 높이 설정
    rowHeights: {
        1: 16.9,
        2: 49.9,
        3: 16.9,
        4: 16.9,
        5: 190.15,  // 이미지 영역
        6: 79.9,
        7: 16.9,
        8: 16.9,
        9: 60.0,    // 위치 정보
        10: 16.9,
        11: 16.9,
        12: 16.9,
        13: 16.9,
        14: 16.9,
        15: 16.9,
        16: 16.9,
        17: 16.9,
        18: 16.9,
        19: 16.9,
        20: 16.9,
        21: 16.9,
        22: 16.9,
        23: 16.9,
        24: 16.9,
        25: 16.9,
        26: 16.9,
        27: 16.9,
        28: 16.9,
        29: 16.9,
        30: 16.9,
        31: 16.9,
        32: 16.9,
        33: 16.9,
        34: 16.9,
        35: 16.9,
        36: 16.9,
        37: 16.9,
        38: 16.9,
        39: 16.9,
        40: 16.9,
        41: 16.9,
        42: 16.9,
        43: 16.9,
        44: 16.9,
        45: 16.9,
        46: 16.9,
        47: 16.9,
        48: 16.9,
        49: 16.9,
        50: 16.9
    },

    // 병합 셀 정보
    mergedCells: [
        'B3:C4',    // PRESENT TO
        'B5:C5',    // 로고 영역
        'B6:C6',    // 빌딩개요/일반
        'B7:B18',   // 빌딩 현황
        'B19:B20',  // 빌딩 세부현황
        'B21:B23',  // 주차 관련
        'B25:B31',  // 임차 제안
        'B32:B39',  // 임대 기준
        'B40:B44',  // 임대 기준 조정
        'B46:B50'   // 예상비용
    ],

    // 카테고리 스타일 설정
    categories: {
        'B3': {
            text: 'PRESENT TO :',
            bgColor: 'FF2C2A2A',
            fontColor: 'FFFFFFFF',
            fontSize: 9,
            bold: true
        },
        'B6': {
            text: '빌딩개요/일반',
            bgColor: 'FFE0E0E0',  // 연한 그레이
            fontColor: 'FF000000',
            fontSize: 9,
            bold: true
        },
        'B7': {
            text: '빌딩 현황',
            bgColor: 'FFE0E0E0',  // 연한 그레이
            fontColor: 'FF000000',
            fontSize: 9,
            bold: true
        },
        'B19': {
            text: '빌딩 세부현황',
            bgColor: 'FFE0E0E0',  // 연한 그레이
            fontColor: 'FF000000',
            fontSize: 9,
            bold: true
        },
        'B21': {
            text: '주차 관련',
            bgColor: 'FFE0E0E0',  // 연한 그레이
            fontColor: 'FF000000',
            fontSize: 9,
            bold: true
        },
        'B25': {
            text: '임차 제안',
            bgColor: 'FFC0C0C0',  // 중간 그레이
            fontColor: 'FF000000',
            fontSize: 9,
            bold: true
        },
        'B32': {
            text: '임대 기준',
            bgColor: 'FFA0A0A0',  // 진한 그레이
            fontColor: 'FF000000',
            fontSize: 9,
            bold: true
        },
        'B40': {
            text: '임대기준 조정',
            bgColor: 'FFA0A0A0',  // 진한 그레이
            fontColor: 'FF000000',
            fontSize: 9,
            bold: true
        },
        'B46': {
            text: '예상비용',
            bgColor: 'FF808080',  // 더 진한 그레이
            fontColor: 'FFFFFFFF',
            fontSize: 9,
            bold: true
        }
    },

    // C열 항목명 설정
    rowLabels: {
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
    },

    // 스타일 설정
    styles: {
        // 기본 폰트
        defaultFont: {
            name: 'LG Smart Regular',
            size: 9
        },
        
        // 헤더 스타일
        headerStyle: {
            font: {
                name: 'LG Smart Regular',
                size: 9,
                bold: true
            },
            fill: {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFCCCCCC' }
            },
            alignment: {
                horizontal: 'center',
                vertical: 'middle',
                wrapText: true
            },
            border: {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            }
        },
        
        // 데이터 셀 스타일
        dataCellStyle: {
            font: {
                name: 'LG Smart Regular',
                size: 9
            },
            alignment: {
                horizontal: 'center',
                vertical: 'middle'
            },
            border: {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            }
        },
        
        // 라벨 셀 스타일 (C열)
        labelCellStyle: {
            font: {
                name: 'LG Smart Regular',
                size: 9
            },
            alignment: {
                horizontal: 'center',
                vertical: 'middle'
            },
            border: {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            }
        }
    },

    // 수식 템플릿
    formulas: {
        // 면적 변환 (평 → m²)
        pyToM2: (col, row) => `ROUNDDOWN(${col}${row}*3.305785,3)`,
        
        // 임대 기준 계산
        monthlyExpense: (col) => `${col}33+${col}34`,
        totalDeposit: (col) => `${col}32*${col}30`,
        totalRent: (col) => `${col}33*${col}30`,
        totalManagement: (col) => `${col}34*${col}30`,
        expensePerDedicated: (col) => `IFERROR((${col}37+${col}38)/${col}31,0)`,
        
        // 렌트프리 적용
        avgRent: (col) => `${col}33-((${col}33*${col}41)/12)`,
        noc: (col) => `IFERROR(((${col}42+${col}43)*(${col}30/${col}31)),0)`,
        
        // 예상비용
        expectedDeposit: (col) => `${col}40*${col}30`,
        expectedMonthlyRent: (col) => `${col}42*${col}30`,
        expectedMonthlyManagement: (col) => `${col}43*${col}30`,
        monthlyTotal: (col) => `${col}47+${col}48`,
        yearlyTotal: (col) => `${col}49*12`
    },

    // 숫자 포맷
    numberFormats: {
        currency: '₩#,##0',
        percentage: '0.00%',
        area_m2: '#,##0.000 "m²"',
        area_py: '#,##0.000 "평"',
        number: '#,##0'
    },

    // 용어 설명
    terminology: {
        52: '용어 설명',
        53: 'NOC : Net Operating Cost의 약자로 임대료와 관리비를 합친 부동산 순 운영 비용',
        54: '렌트프리 : 임대료만 면제 (관리비, 보증금 有)',
        55: '프리렌트 : 임대료 + 관리비 면제 (보증금 有)'
    },

    // 데이터 매핑 (building 객체의 필드명과 엑셀 행 매핑)
    dataMapping: {
        7: 'addressJibun',
        8: 'address',
        9: 'station',
        10: 'floors',
        11: 'completionYear',
        12: 'dedicatedRate',  // 퍼센트로 변환 필요
        13: 'baseFloorArea',
        14: 'baseFloorAreaPy',
        15: 'baseFloorAreaDedicated',
        16: 'baseFloorAreaDedicatedPy',
        17: 'elevator',
        18: 'hvac',
        19: 'buildingUse',
        20: 'structure',
        21: 'parkingSpace',
        22: 'parkingFee',
        23: 'parkingSpace'
    },

    // 기본값 설정
    defaultValues: {
        30: 100,  // 임대면적 (평) 기본값
        31: 50,   // 전용면적 (평) 기본값
        32: 0,    // 월 평당 보증금
        33: 0,    // 월 평당 임대료
        34: 0,    // 월 평당 관리비
        41: 0     // 렌트프리 개월수
    }
};

// 유틸리티 함수들
const LG_UTILS = {
    // 열 인덱스를 문자로 변환 (1 → A, 2 → B, ...)
    getColumnLetter: (index) => {
        let letter = '';
        while (index > 0) {
            index--;
            letter = String.fromCharCode(65 + (index % 26)) + letter;
            index = Math.floor(index / 26);
        }
        return letter;
    },

    // 열 문자를 인덱스로 변환 (A → 1, B → 2, ...)
    getColumnIndex: (letter) => {
        let index = 0;
        for (let i = 0; i < letter.length; i++) {
            index = index * 26 + (letter.charCodeAt(i) - 64);
        }
        return index;
    },

    // 날짜 포맷
    getCurrentDate: () => {
        return new Date().toISOString().split('T')[0];
    },

    // 빌딩 개수 제한 체크
    validateBuildingCount: (count) => {
        if (count === 0) {
            alert('선택된 빌딩이 없습니다.');
            return false;
        }
        if (count > 5) {
            alert('최대 5개까지만 비교할 수 있습니다.');
            return false;
        }
        return true;
    }
};
