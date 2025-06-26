// ===== LG CNS용 Comp List 생성 함수 (최종 버전) =====

// 현재 날짜 반환
function getCurrentDateLG() {
    const now = new Date();
    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const day = String(now.getDate()).padStart(2, '0');
    return `${year}.${month}.${day}`;
}

// LG 양식으로 엑셀 생성
async function generateExcelLG() {
    if (!selectedBuildings || selectedBuildings.length === 0) {
        alert('빌딩을 선택해주세요.');
        return;
    }
    
    const building = selectedBuildings[0];
    console.log('선택된 빌딩:', building);
    
    try {
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('COMP');
        
        // 1. 열 너비 설정
        worksheet.columns = [
            { width: 9.375 },  // A열
            { width: 4.5 },  // B열
            { width: 9.375 },  // C열
            { width: 9.375 },  // D열
            { width: 10.625 },  // E열
            { width: 10 },  // F열
            { width: 10 },  // G열
            { width: 10 },  // H열
            { width: 10 },  // I열
            { width: 10 },  // J열
            { width: 10 },  // K열
            { width: 10 },  // L열
            { width: 10 },  // M열
            { width: 10 },  // N열
            { width: 10 },  // O열
            { width: 10 },  // P열
            { width: 10 },  // Q열
            { width: 10 },  // R열
            { width: 10 },  // S열
            { width: 10 },  // T열
            { width: 10 },  // U열
            { width: 10 },  // V열
            { width: 10 },  // W열
            { width: 10 },  // X열
            { width: 10 },  // Y열
            { width: 9.375 },  // Z열
            { width: 10 },  // AA열
            { width: 10 },  // AB열
            { width: 10 },  // AC열
            { width: 10 },  // AD열
        ];
        
        // 2. 행 높이 설정 (86, 87, 88행 제외)

        // 86-88행 숨기기
        [86, 87, 88].forEach(row => {
            worksheet.getRow(row).height = 0;
        });

        // 3. 병합 셀 설정
        const mergedCells = [
            'B60:D60',
            'N42:P42',
            'B53:D53',
            'K56:M56',
            'N44:P44',
            'N19:P19',
            'H22:J22',
            'H19:J19',
            'E9:G17',
            'O99:P99',
            'H73:J83',
            'E18:G18',
            'T30:V30',
            'N63:P72',
            'T29:V29',
            'N32:P32',
            'K29:M29',
            'A63:A83',
            'T54:V54',
            'C30:D30',
            'K44:M44',
            'K31:M31',
            'H7:J7',
            'E6:G6',
            'Q59:S59',
            'Q22:S22',
            'Q46:S46',
            'T60:V60',
            'A50:A55',
            'Q61:S61',
            'Q73:S83',
            'Q48:S48',
            'N45:P45',
            'K49:M49',
            'N20:P20',
            'H20:J20',
            'T49:V49',
            'E24:G24',
            'N47:P47',
            'N22:P22',
            'T42:V42',
            'B21:D21',
            'K57:M58',
            'T50:V50',
            'K25:L25',
            'T44:V44',
            'E19:G19',
            'B23:D23',
            'Q49:S49',
            'N40:P40',
            'Q30:S30',
            'B51:D51',
            'O95:P95',
            'H8:J8',
            'O89:P89',
            'F92:G92',
            'X100:Y100',
            'Q62:S62',
            'K21:M21',
            'O90:P90',
            'Q54:S54',
            'N53:P53',
            'X102:Y102',
            'N46:P46',
            'K50:M50',
            'E57:G58',
            'N61:P61',
            'A7:D8',
            'B29:D29',
            'N48:P48',
            'K52:M52',
            'E27:G27',
            'B24:D24',
            'K53:M53',
            'B63:D72',
            'A6:D6',
            'H23:J23',
            'B26:D26',
            'N9:P17',
            'N54:P54',
            'H54:J54',
            'N41:P41',
            'E63:G72',
            'H41:J41',
            'Q52:S52',
            'B27:D27',
            'N51:P51',
            'N56:P56',
            'K55:M55',
            'A59:A62',
            'N43:P43',
            'H56:J56',
            'B42:D42',
            'O96:P96',
            'K46:M46',
            'K40:M40',
            'E40:G40',
            'B44:D44',
            'O98:P98',
            'E73:G83',
            'T31:V31',
            'O93:P93',
            'B20:D20',
            'T6:V6',
            'A9:D17',
            'H30:J30',
            'H29:J29',
            'H44:J44',
            'H31:J31',
            'X96:Y96',
            'A48:A49',
            'H60:J60',
            'N18:P18',
            'X98:Y98',
            'Q60:S60',
            'T26:V26',
            'E30:G30',
            'E59:G59',
            'E46:G46',
            'K8:M8',
            'T25:U25',
            'T8:V8',
            'E61:G61',
            'K23:M23',
            'E54:G54',
            'H42:J42',
            'E41:G41',
            'B45:D45',
            'K62:M62',
            'E56:G56',
            'K18:M18',
            'Q7:S7',
            'E43:G43',
            'B47:D47',
            'B59:D59',
            'B46:D46',
            'T32:V32',
            'T47:V47',
            'Q8:S8',
            'B48:D48',
            'O101:P101',
            'H52:J52',
            'T21:V21',
            'H45:J45',
            'H32:J32',
            'T23:V23',
            'T57:V58',
            'T9:V17',
            'H47:J47',
            'A45:A47',
            'F91:G91',
            'T52:V52',
            'B41:D41',
            'T27:V27',
            'T18:V18',
            'N27:P27',
            'Q21:S21',
            'E62:G62',
            'K24:M24',
            'H50:J50',
            'K51:M51',
            'K26:M26',
            'E51:G51',
            'Q25:R25',
            'B61:D61',
            'K42:M42',
            'K19:M19',
            'K6:M6',
            'E50:G50',
            'A18:A25',
            'T24:V24',
            'T63:V72',
            'E52:G52',
            'B18:D18',
            'T53:V53',
            'Q29:S29',
            'Q23:S23',
            'K32:M32',
            'T55:V55',
            'K9:M17',
            'Q31:S31',
            'Q6:S6',
            'K27:M27',
            'E7:G7',
            'N30:P30',
            'A40:A44',
            'T46:V46',
            'B62:D62',
            'T48:V48',
            'N73:P83',
            'T20:V20',
            'K63:M72',
            'Q42:S42',
            'K45:M45',
            'K20:M20',
            'H21:J21',
            'E20:G20',
            'K47:M47',
            'Q19:S19',
            'E60:G60',
            'K22:M22',
            'H9:J17',
            'T40:V40',
            'H25:I25',
            'Q44:S44',
            'B19:D19',
            'H27:J27',
            'N59:P59',
            'Q45:S45',
            'N49:P49',
            'Q32:S32',
            'Q47:S47',
            'O91:P91',
            'E25:F25',
            'T45:V45',
            'E8:G8',
            'H6:J6',
            'T62:V62',
            'T59:V59',
            'N62:P62',
            'Q50:S50',
            'Q57:S58',
            'K59:M59',
            'T51:V51',
            'Q27:S27',
            'K61:M61',
            'K48:M48',
            'H24:J24',
            'H18:J18',
            'O104:P104',
            'E21:G21',
            'H26:J26',
            'H53:J53',
            'B25:D25',
            'B22:D22',
            'E23:G23',
            'H55:J55',
            'Q55:S55',
            'X99:Y99',
            'K41:M41',
            'N52:P52',
            'Q40:S40',
            'X101:Y101',
            'K43:M43',
            'E49:G49',
            'T61:V61',
            'B40:D40',
            'H48:J48',
            'B73:D83',
            'X104:Y104',
            'Q53:S53',
            'O94:P94',
            'H63:J72',
            'H43:J43',
            'X106:Y106',
            'N7:P7',
            'N50:P50',
            'N57:P58',
            'K54:M54',
            'H57:J58',
            'O106:P106',
            'E29:G29',
            'O105:P105',
            'E31:G31',
            'B28:D28',
            'A33:D39',
            'N60:P60',
            'T22:V22',
            'E26:G26',
            'T73:V83',
            'Q41:S41',
            'Q56:S56',
            'N55:P55',
            'Q43:S43',
            'B54:D54',
            'E42:G42',
            'O100:P100',
            'B56:D56',
            'H40:J40',
            'B43:D43',
            'E44:G44',
            'K60:M60',
            'O102:P102',
            'X105:Y105',
            'H51:J51',
            'N25:O25',
            'N8:P8',
            'O97:P97',
            'K7:M7',
            'E32:G32',
            'T19:V19',
            'Q9:S17',
            'E47:G47',
            'B31:D31',
            'Q24:S24',
            'Q18:S18',
            'H59:J59',
            'Q51:S51',
            'N21:P21',
            'H46:J46',
            'X95:Y95',
            'H61:J61',
            'N23:P23',
            'F89:G89',
            'X97:Y97',
            'O92:P92',
            'E45:G45',
            'T41:V41',
            'B49:D49',
            'O103:P103',
            'T56:V56',
            'T7:V7',
            'T43:V43',
            'E22:G22',
            'B57:D58',
            'B50:D50',
            'K73:M83',
            'C32:D32',
            'A56:A58',
            'H49:J49',
            'Q63:S72',
            'Q26:S26',
            'N29:P29',
            'A26:A32',
            'X103:Y103',
            'N31:P31',
            'B55:D55',
            'N6:P6',
            'H62:J62',
            'N24:P24',
            'F90:G90',
            'E53:G53',
            'N26:P26',
            'K30:M30',
            'E55:G55',
            'E48:G48',
            'B52:D52',
            'Q20:S20',
        ];
        
        mergedCells.forEach(range => {
            try {
                worksheet.mergeCells(range);
            } catch (e) {
                console.warn(`병합 실패: ${range}`, e);
            }
        });
        
        // 4. 셀 데이터 및 스타일 설정
        setupCellsLG(worksheet, building);
        
        // 5. 파일 저장
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { 
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
        });
        
        const fileName = `CompList_LG_${building.name}_${getCurrentDateLG().replace(/\./g, '')}.xlsx`;
        saveAs(blob, fileName);
        
        alert('LG용 Comp List가 생성되었습니다!');
        
    } catch (error) {
        console.error('엑셀 생성 오류:', error);
        alert('엑셀 파일 생성 중 오류가 발생했습니다.\n' + error.message);
    }
}

// 셀 설정 함수
function setupCellsLG(worksheet, building) {
    console.log('셀 설정 시작');
    
    // A1-A4 샘플 문구
    const sampleTexts = {
        'A1': '[임차제안건 제목을 입력하세요.]',
        'A2': '규모를 입력하세요.',
        'A3': '계약기간을 입력하세요.',
        'A4': '위치를 입력하세요.'
    };
    
    // 모든 수식 정보 (원본에서 복사)
    const formulas = {
        'N22': '=98995.31*0.3025',
        'T22': '=46549.73*0.3025',
        'Q24': '=279/469',
        'T24': '=268.88/522.05',
        'E25': '=G25*0.3025',
        'H25': '=J25*0.3025',
        'K25': '=M25*0.3025',
        'N25': '=P25*0.3025',
        'E30': '=E29/E32',
        'H30': '=H29/H32',
        'K30': '=K29/K32',
        'N30': '=N29/N32',
        'Q30': '=Q29/Q32',
        'T30': '=T29/T32',
        'E32': '=E31*G25',
        'H32': '=H31*J25',
        'K32': '=K31*M25',
        'N32': '=N31*P25',
        'Q32': '=Q31*S25',
        'T32': '=T31*V25',
        'F39': '=SUM(F34:F38)',
        'G39': '=SUM(G34:G38)',
        'I39': '=SUM(I34:I38)',
        'J39': '=SUM(J34:J38)',
        'L39': '=SUM(L34:L38)',
        'M39': '=SUM(M34:M38)',
        'O39': '=SUM(O34:O38)',
        'P39': '=SUM(P34:P38)',
        'R39': '=SUM(R34:R38)',
        'S39': '=SUM(S34:S38)',
        'U39': '=SUM(U34:U38)',
        'V39': '=SUM(V34:V38)',
        'E43': '=SUM(F34:F35)',
        'H43': '=I34',
        'K43': '=L35+L34',
        'N43': '=O35',
        'Q43': '=R34',
        'T43': '=U34',
        'E44': '=SUM(G34:G35)',
        'H44': '=J34',
        'K44': '=M35+M34',
        'N44': '=P35',
        'Q44': '=S34',
        'T44': '=V34',
        'E48': '=E46*(12-E49)/12',
        'H48': '=H46*(12-H49)/12',
        'K48': '=K46*(12-K49)/12',
        'N48': '=N46*(12-N49)/12',
        'Q48': '=Q46*(12-Q49)/12',
        'T48': '=T46*(12-T49)/12',
        'E50': '=E45*E44',
        'H50': '=H45*H44',
        'K50': '=K45*K44',
        'N50': '=N45*N44',
        'Q50': '=Q45*Q44',
        'T50': '=T45*T44',
        'E51': '=E46*E44',
        'H51': '=H46*H44',
        'K51': '=K46*K44',
        'N51': '=N46*N44',
        'Q51': '=Q46*Q44',
        'T51': '=T46*T44',
        'E52': '=E47*E44',
        'H52': '=H47*H44',
        'K52': '=K47*K44',
        'Q52': '=Q47*Q44',
        'T52': '=T47*T44',
        'E54': '=E51',
        'H54': '=H51+H52',
        'K54': '=K51+K52',
        'N54': '=N51',
        'Q54': '=Q51+Q52',
        'T54': '=T51+T52',
        'E55': '=E54*21',
        'H55': '=H54*21',
        'K55': '=K54*21',
        'N55': '=N54*21',
        'Q55': '=Q54*21',
        'T55': '=T54*21',
        'H59': '=732+44',
        'K59': '=732+44',
        'E61': '=E44/E60',
        'H61': '=H44/H60',
        'K61': '=K44/K60',
        'N61': '=N44/N60',
        'Q61': '=Q44/Q60',
        'T61': '=T44/T60',
    };
    
    // A1 셀
    try {
        const cell_A1 = worksheet.getCell('A1');
        cell_A1.value = sampleTexts['A1'];
        cell_A1.font = { name: 'LG스마트체 Regular', size: 14.0, color: { argb: 'FF000000' } };
        cell_A1.alignment = { horizontal: 'center', vertical: 'center' };
        cell_A1.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 A1 설정 실패:', e);
    }

    // A10 셀
    try {
        const cell_A10 = worksheet.getCell('A10');
        cell_A10.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A10.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A10.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A10 설정 실패:', e);
    }

    // A11 셀
    try {
        const cell_A11 = worksheet.getCell('A11');
        cell_A11.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A11.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A11.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A11 설정 실패:', e);
    }

    // A12 셀
    try {
        const cell_A12 = worksheet.getCell('A12');
        cell_A12.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A12.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A12.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A12 설정 실패:', e);
    }

    // A13 셀
    try {
        const cell_A13 = worksheet.getCell('A13');
        cell_A13.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A13.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A13.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A13 설정 실패:', e);
    }

    // A14 셀
    try {
        const cell_A14 = worksheet.getCell('A14');
        cell_A14.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A14.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A14.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A14 설정 실패:', e);
    }

    // A15 셀
    try {
        const cell_A15 = worksheet.getCell('A15');
        cell_A15.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A15.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A15.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A15 설정 실패:', e);
    }

    // A16 셀
    try {
        const cell_A16 = worksheet.getCell('A16');
        cell_A16.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A16.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A16.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A16 설정 실패:', e);
    }

    // A17 셀
    try {
        const cell_A17 = worksheet.getCell('A17');
        cell_A17.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A17.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A17.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A17 설정 실패:', e);
    }

    // A18 셀
    try {
        const cell_A18 = worksheet.getCell('A18');
        cell_A18.value = '기초\n정보';
        cell_A18.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_A18.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        cell_A18.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A18 설정 실패:', e);
    }

    // A19 셀
    try {
        const cell_A19 = worksheet.getCell('A19');
        cell_A19.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A19.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A19.border = { left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A19 설정 실패:', e);
    }

    // A2 셀
    try {
        const cell_A2 = worksheet.getCell('A2');
        cell_A2.value = sampleTexts['A2'];
        cell_A2.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A2.alignment = { horizontal: 'center', vertical: 'center' };
        cell_A2.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 A2 설정 실패:', e);
    }

    // A20 셀
    try {
        const cell_A20 = worksheet.getCell('A20');
        cell_A20.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A20.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A20.border = { left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A20 설정 실패:', e);
    }

    // A21 셀
    try {
        const cell_A21 = worksheet.getCell('A21');
        cell_A21.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A21.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A21.border = { left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A21 설정 실패:', e);
    }

    // A22 셀
    try {
        const cell_A22 = worksheet.getCell('A22');
        cell_A22.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A22.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A22.border = { left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A22 설정 실패:', e);
    }

    // A23 셀
    try {
        const cell_A23 = worksheet.getCell('A23');
        cell_A23.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A23.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A23.border = { left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A23 설정 실패:', e);
    }

    // A24 셀
    try {
        const cell_A24 = worksheet.getCell('A24');
        cell_A24.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A24.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A24.border = { left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A24 설정 실패:', e);
    }

    // A25 셀
    try {
        const cell_A25 = worksheet.getCell('A25');
        cell_A25.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A25.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A25.border = { left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A25 설정 실패:', e);
    }

    // A26 셀
    try {
        const cell_A26 = worksheet.getCell('A26');
        cell_A26.value = '채권\n분석';
        cell_A26.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_A26.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        cell_A26.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A26 설정 실패:', e);
    }

    // A27 셀
    try {
        const cell_A27 = worksheet.getCell('A27');
        cell_A27.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A27.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A27.border = { left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A27 설정 실패:', e);
    }

    // A28 셀
    try {
        const cell_A28 = worksheet.getCell('A28');
        cell_A28.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A28.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A28.border = { left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A28 설정 실패:', e);
    }

    // A29 셀
    try {
        const cell_A29 = worksheet.getCell('A29');
        cell_A29.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A29.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A29.border = { left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A29 설정 실패:', e);
    }

    // A3 셀
    try {
        const cell_A3 = worksheet.getCell('A3');
        cell_A3.value = sampleTexts['A3'];
        cell_A3.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A3.alignment = { horizontal: 'center', vertical: 'center' };
        cell_A3.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 A3 설정 실패:', e);
    }

    // A30 셀
    try {
        const cell_A30 = worksheet.getCell('A30');
        cell_A30.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A30.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A30.border = { left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A30 설정 실패:', e);
    }

    // A31 셀
    try {
        const cell_A31 = worksheet.getCell('A31');
        cell_A31.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A31.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A31.border = { left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A31 설정 실패:', e);
    }

    // A32 셀
    try {
        const cell_A32 = worksheet.getCell('A32');
        cell_A32.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A32.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A32.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A32 설정 실패:', e);
    }

    // A33 셀
    try {
        const cell_A33 = worksheet.getCell('A33');
        cell_A33.value = '현재 공실';
        cell_A33.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_A33.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        cell_A33.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_A33.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 A33 설정 실패:', e);
    }

    // A34 셀
    try {
        const cell_A34 = worksheet.getCell('A34');
        cell_A34.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A34.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A34.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A34 설정 실패:', e);
    }

    // A35 셀
    try {
        const cell_A35 = worksheet.getCell('A35');
        cell_A35.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A35.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A35.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A35 설정 실패:', e);
    }

    // A36 셀
    try {
        const cell_A36 = worksheet.getCell('A36');
        cell_A36.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A36.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A36.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A36 설정 실패:', e);
    }

    // A37 셀
    try {
        const cell_A37 = worksheet.getCell('A37');
        cell_A37.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A37.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A37.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A37 설정 실패:', e);
    }

    // A38 셀
    try {
        const cell_A38 = worksheet.getCell('A38');
        cell_A38.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A38.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A38.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A38 설정 실패:', e);
    }

    // A39 셀
    try {
        const cell_A39 = worksheet.getCell('A39');
        cell_A39.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A39.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A39.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A39 설정 실패:', e);
    }

    // A4 셀
    try {
        const cell_A4 = worksheet.getCell('A4');
        cell_A4.value = sampleTexts['A4'];
        cell_A4.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A4.alignment = { horizontal: 'center', vertical: 'center' };
        cell_A4.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 A4 설정 실패:', e);
    }

    // A40 셀
    try {
        const cell_A40 = worksheet.getCell('A40');
        cell_A40.value = '제안';
        cell_A40.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_A40.alignment = { horizontal: 'center', vertical: 'center' };
        cell_A40.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_A40.numFmt = '@';
    } catch (e) {
        console.warn('셀 A40 설정 실패:', e);
    }

    // A41 셀
    try {
        const cell_A41 = worksheet.getCell('A41');
        cell_A41.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A41.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A41.border = { left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A41 설정 실패:', e);
    }

    // A42 셀
    try {
        const cell_A42 = worksheet.getCell('A42');
        cell_A42.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A42.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A42.border = { left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A42 설정 실패:', e);
    }

    // A43 셀
    try {
        const cell_A43 = worksheet.getCell('A43');
        cell_A43.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A43.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A43.border = { left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A43 설정 실패:', e);
    }

    // A44 셀
    try {
        const cell_A44 = worksheet.getCell('A44');
        cell_A44.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A44.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A44.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A44 설정 실패:', e);
    }

    // A45 셀
    try {
        const cell_A45 = worksheet.getCell('A45');
        cell_A45.value = '기준층\n임대기준';
        cell_A45.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_A45.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        cell_A45.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_A45.numFmt = '#,##0_ ';
    } catch (e) {
        console.warn('셀 A45 설정 실패:', e);
    }

    // A46 셀
    try {
        const cell_A46 = worksheet.getCell('A46');
        cell_A46.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A46.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A46.border = { left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A46 설정 실패:', e);
    }

    // A47 셀
    try {
        const cell_A47 = worksheet.getCell('A47');
        cell_A47.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A47.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A47.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A47 설정 실패:', e);
    }

    // A48 셀
    try {
        const cell_A48 = worksheet.getCell('A48');
        cell_A48.value = '실질\n임대기준';
        cell_A48.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_A48.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        cell_A48.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_A48.numFmt = '#,##0_ ';
    } catch (e) {
        console.warn('셀 A48 설정 실패:', e);
    }

    // A49 셀
    try {
        const cell_A49 = worksheet.getCell('A49');
        cell_A49.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A49.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A49.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A49 설정 실패:', e);
    }

    // A50 셀
    try {
        const cell_A50 = worksheet.getCell('A50');
        cell_A50.value = '비용검토';
        cell_A50.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_A50.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        cell_A50.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_A50.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 A50 설정 실패:', e);
    }

    // A51 셀
    try {
        const cell_A51 = worksheet.getCell('A51');
        cell_A51.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A51.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A51.border = { left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A51 설정 실패:', e);
    }

    // A52 셀
    try {
        const cell_A52 = worksheet.getCell('A52');
        cell_A52.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A52.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A52.border = { left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A52 설정 실패:', e);
    }

    // A53 셀
    try {
        const cell_A53 = worksheet.getCell('A53');
        cell_A53.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A53.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A53.border = { left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A53 설정 실패:', e);
    }

    // A54 셀
    try {
        const cell_A54 = worksheet.getCell('A54');
        cell_A54.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A54.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A54.border = { left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A54 설정 실패:', e);
    }

    // A55 셀
    try {
        const cell_A55 = worksheet.getCell('A55');
        cell_A55.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A55.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A55.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A55 설정 실패:', e);
    }

    // A56 셀
    try {
        const cell_A56 = worksheet.getCell('A56');
        cell_A56.value = '공사기간\nFAVOR';
        cell_A56.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_A56.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        cell_A56.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_A56.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 A56 설정 실패:', e);
    }

    // A57 셀
    try {
        const cell_A57 = worksheet.getCell('A57');
        cell_A57.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A57.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A57.border = { left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A57 설정 실패:', e);
    }

    // A58 셀
    try {
        const cell_A58 = worksheet.getCell('A58');
        cell_A58.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A58.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A58.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A58 설정 실패:', e);
    }

    // A59 셀
    try {
        const cell_A59 = worksheet.getCell('A59');
        cell_A59.value = '주차현황';
        cell_A59.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_A59.alignment = { horizontal: 'center', vertical: 'center' };
        cell_A59.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_A59.numFmt = '"월"\ ##,##0"만원 / 대"';
    } catch (e) {
        console.warn('셀 A59 설정 실패:', e);
    }

    // A6 셀
    try {
        const cell_A6 = worksheet.getCell('A6');
        cell_A6.value = '위치';
        cell_A6.font = { name: 'LG스마트체 Regular', size: 10.0, color: { argb: 'FF000000' } };
        cell_A6.alignment = { horizontal: 'center', vertical: 'center' };
        cell_A6.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_A6.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 A6 설정 실패:', e);
    }

    // A60 셀
    try {
        const cell_A60 = worksheet.getCell('A60');
        cell_A60.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A60.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A60.border = { left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A60 설정 실패:', e);
    }

    // A61 셀
    try {
        const cell_A61 = worksheet.getCell('A61');
        cell_A61.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A61.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A61.border = { left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A61 설정 실패:', e);
    }

    // A62 셀
    try {
        const cell_A62 = worksheet.getCell('A62');
        cell_A62.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A62.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A62.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A62 설정 실패:', e);
    }

    // A63 셀
    try {
        const cell_A63 = worksheet.getCell('A63');
        cell_A63.value = '기타';
        cell_A63.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_A63.alignment = { horizontal: 'center', vertical: 'center' };
        cell_A63.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_A63.numFmt = '"월"\ ##,##0"만원 / 대"';
    } catch (e) {
        console.warn('셀 A63 설정 실패:', e);
    }

    // A64 셀
    try {
        const cell_A64 = worksheet.getCell('A64');
        cell_A64.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A64.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A64.border = { left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A64 설정 실패:', e);
    }

    // A65 셀
    try {
        const cell_A65 = worksheet.getCell('A65');
        cell_A65.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A65.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A65.border = { left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A65 설정 실패:', e);
    }

    // A66 셀
    try {
        const cell_A66 = worksheet.getCell('A66');
        cell_A66.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A66.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A66.border = { left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A66 설정 실패:', e);
    }

    // A67 셀
    try {
        const cell_A67 = worksheet.getCell('A67');
        cell_A67.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A67.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A67.border = { left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A67 설정 실패:', e);
    }

    // A68 셀
    try {
        const cell_A68 = worksheet.getCell('A68');
        cell_A68.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A68.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A68.border = { left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A68 설정 실패:', e);
    }

    // A69 셀
    try {
        const cell_A69 = worksheet.getCell('A69');
        cell_A69.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A69.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A69.border = { left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A69 설정 실패:', e);
    }

    // A7 셀
    try {
        const cell_A7 = worksheet.getCell('A7');
        cell_A7.value = '제안';
        cell_A7.font = { name: 'LG스마트체 Regular', size: 10.0, color: { argb: 'FF000000' } };
        cell_A7.alignment = { horizontal: 'center', vertical: 'center' };
        cell_A7.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_A7.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 A7 설정 실패:', e);
    }

    // A70 셀
    try {
        const cell_A70 = worksheet.getCell('A70');
        cell_A70.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A70.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A70.border = { left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A70 설정 실패:', e);
    }

    // A71 셀
    try {
        const cell_A71 = worksheet.getCell('A71');
        cell_A71.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A71.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A71.border = { left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A71 설정 실패:', e);
    }

    // A72 셀
    try {
        const cell_A72 = worksheet.getCell('A72');
        cell_A72.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A72.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A72.border = { left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A72 설정 실패:', e);
    }

    // A73 셀
    try {
        const cell_A73 = worksheet.getCell('A73');
        cell_A73.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A73.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A73.border = { left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A73 설정 실패:', e);
    }

    // A74 셀
    try {
        const cell_A74 = worksheet.getCell('A74');
        cell_A74.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A74.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A74.border = { left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A74 설정 실패:', e);
    }

    // A75 셀
    try {
        const cell_A75 = worksheet.getCell('A75');
        cell_A75.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A75.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A75.border = { left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A75 설정 실패:', e);
    }

    // A76 셀
    try {
        const cell_A76 = worksheet.getCell('A76');
        cell_A76.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A76.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A76.border = { left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A76 설정 실패:', e);
    }

    // A77 셀
    try {
        const cell_A77 = worksheet.getCell('A77');
        cell_A77.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A77.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A77.border = { left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A77 설정 실패:', e);
    }

    // A78 셀
    try {
        const cell_A78 = worksheet.getCell('A78');
        cell_A78.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A78.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A78.border = { left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A78 설정 실패:', e);
    }

    // A79 셀
    try {
        const cell_A79 = worksheet.getCell('A79');
        cell_A79.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A79.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A79.border = { left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A79 설정 실패:', e);
    }

    // A8 셀
    try {
        const cell_A8 = worksheet.getCell('A8');
        cell_A8.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A8.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A8.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A8 설정 실패:', e);
    }

    // A80 셀
    try {
        const cell_A80 = worksheet.getCell('A80');
        cell_A80.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A80.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A80.border = { left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A80 설정 실패:', e);
    }

    // A81 셀
    try {
        const cell_A81 = worksheet.getCell('A81');
        cell_A81.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A81.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A81.border = { left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A81 설정 실패:', e);
    }

    // A82 셀
    try {
        const cell_A82 = worksheet.getCell('A82');
        cell_A82.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A82.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A82.border = { left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A82 설정 실패:', e);
    }

    // A83 셀
    try {
        const cell_A83 = worksheet.getCell('A83');
        cell_A83.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_A83.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_A83.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A83 설정 실패:', e);
    }

    // A84 셀
    try {
        const cell_A84 = worksheet.getCell('A84');
        cell_A84.value = '1) 실질임대료(Rent Free 반영한 임대가)  / 2) 월 납부액 = 월 실질임대료 + 월관리비 (초기년도 기준으로 인상률 미반영)';
        cell_A84.font = { name: 'LG스마트체 Regular', size: 10.0, color: { argb: 'FF000000' } };
        cell_A84.alignment = { horizontal: 'center', vertical: 'center' };
        cell_A84.border = { top: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A84 설정 실패:', e);
    }

    // A85 셀
    try {
        const cell_A85 = worksheet.getCell('A85');
        cell_A85.value = '3) 연간납부비용 = 연임대료 + 연관리비 (초기년도 기준으로 인상률 미반영, 보증금 미반영)  4) RF : Rent Free (임대료 무상, 관리비 부과)  5) FO : Fit-out (인테리어공사기간동안 임대료 무상, 관리비 부과)';
        cell_A85.font = { name: 'LG스마트체 Regular', size: 10.0, color: { argb: 'FF000000' } };
        cell_A85.alignment = { horizontal: 'center', vertical: 'center' };
    } catch (e) {
        console.warn('셀 A85 설정 실패:', e);
    }

    // A9 셀
    try {
        const cell_A9 = worksheet.getCell('A9');
        cell_A9.value = '건물 외관';
        cell_A9.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_A9.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        cell_A9.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 A9 설정 실패:', e);
    }

    // B17 셀
    try {
        const cell_B17 = worksheet.getCell('B17');
        cell_B17.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_B17.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_B17.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 B17 설정 실패:', e);
    }

    // B18 셀
    try {
        const cell_B18 = worksheet.getCell('B18');
        cell_B18.value = '주   소';
        cell_B18.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_B18.alignment = { horizontal: 'center', vertical: 'center' };
        cell_B18.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_B18.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 B18 설정 실패:', e);
    }

    // B19 셀
    try {
        const cell_B19 = worksheet.getCell('B19');
        cell_B19.value = '위   치';
        cell_B19.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_B19.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        cell_B19.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_B19.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 B19 설정 실패:', e);
    }

    // B20 셀
    try {
        const cell_B20 = worksheet.getCell('B20');
        cell_B20.value = building.completionYear ? `${building.completionYear}년` : '준공일';
        cell_B20.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_B20.alignment = { horizontal: 'center', vertical: 'center' };
        cell_B20.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_B20.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 B20 설정 실패:', e);
    }

    // B21 셀
    try {
        const cell_B21 = worksheet.getCell('B21');
        cell_B21.value = '규  모';
        cell_B21.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_B21.alignment = { horizontal: 'center', vertical: 'center' };
        cell_B21.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_B21.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 B21 설정 실패:', e);
    }

    // B22 셀
    try {
        const cell_B22 = worksheet.getCell('B22');
        cell_B22.value = building.totalFloorArea ? `${building.totalFloorArea} ㎡` : '연면적';
        cell_B22.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_B22.alignment = { horizontal: 'center', vertical: 'center' };
        cell_B22.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_B22.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 B22 설정 실패:', e);
    }

    // B23 셀
    try {
        const cell_B23 = worksheet.getCell('B23');
        cell_B23.value = building.baseFloorAreaDedicatedPy ? `${building.baseFloorAreaDedicatedPy} 평` : '기준층 전용면적';
        cell_B23.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_B23.alignment = { horizontal: 'center', vertical: 'center' };
        cell_B23.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_B23.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 B23 설정 실패:', e);
    }

    // B24 셀
    try {
        const cell_B24 = worksheet.getCell('B24');
        cell_B24.value = '전용률';
        cell_B24.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_B24.alignment = { horizontal: 'center', vertical: 'center' };
        cell_B24.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_B24.numFmt = '0%';
    } catch (e) {
        console.warn('셀 B24 설정 실패:', e);
    }

    // B25 셀
    try {
        const cell_B25 = worksheet.getCell('B25');
        cell_B25.value = '대지면적';
        cell_B25.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_B25.alignment = { horizontal: 'center', vertical: 'center' };
        cell_B25.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_B25.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 B25 설정 실패:', e);
    }

    // B26 셀
    try {
        const cell_B26 = worksheet.getCell('B26');
        cell_B26.value = '소유자 (임대인)';
        cell_B26.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_B26.alignment = { horizontal: 'center', vertical: 'center' };
        cell_B26.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_B26.numFmt = '@';
    } catch (e) {
        console.warn('셀 B26 설정 실패:', e);
    }

    // B27 셀
    try {
        const cell_B27 = worksheet.getCell('B27');
        cell_B27.value = '채권담보 설정여부';
        cell_B27.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_B27.alignment = { horizontal: 'center', vertical: 'center' };
        cell_B27.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_B27.numFmt = '@';
    } catch (e) {
        console.warn('셀 B27 설정 실패:', e);
    }

    // B28 셀
    try {
        const cell_B28 = worksheet.getCell('B28');
        cell_B28.value = '공동담보 총 대지지분';
        cell_B28.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_B28.alignment = { horizontal: 'center', vertical: 'center' };
        cell_B28.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_B28.numFmt = '@';
    } catch (e) {
        console.warn('셀 B28 설정 실패:', e);
    }

    // B29 셀
    try {
        const cell_B29 = worksheet.getCell('B29');
        cell_B29.value = '선순위 담보 총액';
        cell_B29.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_B29.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        cell_B29.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_B29.numFmt = '0%';
    } catch (e) {
        console.warn('셀 B29 설정 실패:', e);
    }

    // B30 셀
    try {
        const cell_B30 = worksheet.getCell('B30');
        cell_B30.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_B30.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        cell_B30.border = { bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'hair', color: { argb: 'FF000000' } } };
        cell_B30.numFmt = '0%';
    } catch (e) {
        console.warn('셀 B30 설정 실패:', e);
    }

    // B31 셀
    try {
        const cell_B31 = worksheet.getCell('B31');
        cell_B31.value = '개별공시지가(25년 1월 기준)';
        cell_B31.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_B31.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        cell_B31.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_B31.numFmt = '0%';
    } catch (e) {
        console.warn('셀 B31 설정 실패:', e);
    }

    // B32 셀
    try {
        const cell_B32 = worksheet.getCell('B32');
        cell_B32.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_B32.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        cell_B32.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'hair', color: { argb: 'FF000000' } } };
        cell_B32.numFmt = '0%';
    } catch (e) {
        console.warn('셀 B32 설정 실패:', e);
    }

    // B33 셀
    try {
        const cell_B33 = worksheet.getCell('B33');
        cell_B33.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_B33.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_B33.border = { top: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 B33 설정 실패:', e);
    }

    // B40 셀
    try {
        const cell_B40 = worksheet.getCell('B40');
        cell_B40.value = '계약기간';
        cell_B40.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_B40.alignment = { horizontal: 'center', vertical: 'center' };
        cell_B40.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_B40.numFmt = '@';
    } catch (e) {
        console.warn('셀 B40 설정 실패:', e);
    }

    // B41 셀
    try {
        const cell_B41 = worksheet.getCell('B41');
        cell_B41.value = '입주가능 시기';
        cell_B41.font = { name: 'LG스마트체 Regular', size: 9.0, bold: true, color: { argb: 'FF000000' } };
        cell_B41.alignment = { horizontal: 'center', vertical: 'center' };
        cell_B41.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_B41.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 B41 설정 실패:', e);
    }

    // B42 셀
    try {
        const cell_B42 = worksheet.getCell('B42');
        cell_B42.value = '제안 층';
        cell_B42.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_B42.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        cell_B42.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_B42.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 B42 설정 실패:', e);
    }

    // B43 셀
    try {
        const cell_B43 = worksheet.getCell('B43');
        cell_B43.value = '전용면적'; // 병합된 항목명
        cell_B43.font = { name: 'LG스마트체 Regular', size: 9.0, bold: true, color: { argb: 'FFC00000' } };
        cell_B43.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        cell_B43.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_B43.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 B43 설정 실패:', e);
    }

    // B44 셀
    try {
        const cell_B44 = worksheet.getCell('B44');
        cell_B44.value = '임대면적';
        cell_B44.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_B44.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        cell_B44.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_B44.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 B44 설정 실패:', e);
    }

    // B45 셀
    try {
        const cell_B45 = worksheet.getCell('B45');
        cell_B45.value = '보증금';
        cell_B45.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_B45.alignment = { horizontal: 'center', vertical: 'center' };
        cell_B45.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_B45.numFmt = '#,##0_ ';
    } catch (e) {
        console.warn('셀 B45 설정 실패:', e);
    }

    // B46 셀
    try {
        const cell_B46 = worksheet.getCell('B46');
        cell_B46.value = '임대료';
        cell_B46.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_B46.alignment = { horizontal: 'center', vertical: 'center' };
        cell_B46.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_B46.numFmt = '#,##0_ ';
    } catch (e) {
        console.warn('셀 B46 설정 실패:', e);
    }

    // B47 셀
    try {
        const cell_B47 = worksheet.getCell('B47');
        cell_B47.value = '관리비';
        cell_B47.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_B47.alignment = { horizontal: 'center', vertical: 'center' };
        cell_B47.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_B47.numFmt = '#,##0_ ';
    } catch (e) {
        console.warn('셀 B47 설정 실패:', e);
    }

    // B48 셀
    try {
        const cell_B48 = worksheet.getCell('B48');
        cell_B48.value = '실질 임대료(RF만 반영)1)';
        cell_B48.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_B48.alignment = { horizontal: 'center', vertical: 'center' };
        cell_B48.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_B48.numFmt = '#,##0_ ';
    } catch (e) {
        console.warn('셀 B48 설정 실패:', e);
    }

    // B49 셀
    try {
        const cell_B49 = worksheet.getCell('B49');
        cell_B49.value = '연간 무상임대 (R.F)';
        cell_B49.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_B49.alignment = { horizontal: 'center', vertical: 'center' };
        cell_B49.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_B49.numFmt = '#,##0_ ';
    } catch (e) {
        console.warn('셀 B49 설정 실패:', e);
    }

    // B50 셀
    try {
        const cell_B50 = worksheet.getCell('B50');
        cell_B50.value = '보증금';
        cell_B50.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_B50.alignment = { horizontal: 'center', vertical: 'center' };
        cell_B50.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_B50.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 B50 설정 실패:', e);
    }

    // B51 셀
    try {
        const cell_B51 = worksheet.getCell('B51');
        cell_B51.value = '월 임대료';
        cell_B51.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_B51.alignment = { horizontal: 'center', vertical: 'center' };
        cell_B51.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_B51.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 B51 설정 실패:', e);
    }

    // B52 셀
    try {
        const cell_B52 = worksheet.getCell('B52');
        cell_B52.value = '월 관리비';
        cell_B52.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_B52.alignment = { horizontal: 'center', vertical: 'center' };
        cell_B52.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_B52.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 B52 설정 실패:', e);
    }

    // B53 셀
    try {
        const cell_B53 = worksheet.getCell('B53');
        cell_B53.value = '관리비 내역';
        cell_B53.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FFC00000' } };
        cell_B53.alignment = { horizontal: 'center', vertical: 'center' };
        cell_B53.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_B53.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 B53 설정 실패:', e);
    }

    // B54 셀
    try {
        const cell_B54 = worksheet.getCell('B54');
        cell_B54.value = '월납부액';
        cell_B54.font = { name: 'LG스마트체 Regular', size: 9.0, bold: true, color: { argb: 'FFC00000' } };
        cell_B54.alignment = { horizontal: 'center', vertical: 'center' };
        cell_B54.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_B54.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 B54 설정 실패:', e);
    }

    // B55 셀
    try {
        const cell_B55 = worksheet.getCell('B55');
        cell_B55.value = '(21개월 기준) 총 납부 비용3)';
        cell_B55.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_B55.alignment = { horizontal: 'center', vertical: 'center' };
        cell_B55.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_B55.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 B55 설정 실패:', e);
    }

    // B56 셀
    try {
        const cell_B56 = worksheet.getCell('B56');
        cell_B56.value = '인테리어 기간 (F.O)';
        cell_B56.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_B56.alignment = { horizontal: 'center', vertical: 'center' };
        cell_B56.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_B56.numFmt = '#,##0_ ';
    } catch (e) {
        console.warn('셀 B56 설정 실패:', e);
    }

    // B57 셀
    try {
        const cell_B57 = worksheet.getCell('B57');
        cell_B57.value = '인테리어지원금 (T.I)';
        cell_B57.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_B57.alignment = { horizontal: 'center', vertical: 'center' };
        cell_B57.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_B57.numFmt = '#,##0_ ';
    } catch (e) {
        console.warn('셀 B57 설정 실패:', e);
    }

    // B58 셀
    try {
        const cell_B58 = worksheet.getCell('B58');
        cell_B58.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_B58.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_B58.border = { bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 B58 설정 실패:', e);
    }

    // B59 셀
    try {
        const cell_B59 = worksheet.getCell('B59');
        cell_B59.value = '총 주차대수';
        cell_B59.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_B59.alignment = { horizontal: 'center', vertical: 'center' };
        cell_B59.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_B59.numFmt = '#,##0_ ';
    } catch (e) {
        console.warn('셀 B59 설정 실패:', e);
    }

    // B6 셀
    try {
        const cell_B6 = worksheet.getCell('B6');
        cell_B6.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_B6.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_B6.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 B6 설정 실패:', e);
    }

    // B60 셀
    try {
        const cell_B60 = worksheet.getCell('B60');
        cell_B60.value = '무료주차 조건(임대면적)';
        cell_B60.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_B60.alignment = { horizontal: 'center', vertical: 'center' };
        cell_B60.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_B60.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 B60 설정 실패:', e);
    }

    // B61 셀
    try {
        const cell_B61 = worksheet.getCell('B61');
        cell_B61.value = '무료주차 제공대수';
        cell_B61.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_B61.alignment = { horizontal: 'center', vertical: 'center' };
        cell_B61.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_B61.numFmt = '#,##0\ "대"';
    } catch (e) {
        console.warn('셀 B61 설정 실패:', e);
    }

    // B62 셀
    try {
        const cell_B62 = worksheet.getCell('B62');
        cell_B62.value = '유료주차(VAT별도)';
        cell_B62.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_B62.alignment = { horizontal: 'center', vertical: 'center' };
        cell_B62.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_B62.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 B62 설정 실패:', e);
    }

    // B63 셀
    try {
        const cell_B63 = worksheet.getCell('B63');
        cell_B63.value = '평면도';
        cell_B63.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_B63.alignment = { horizontal: 'center', vertical: 'center' };
        cell_B63.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_B63.numFmt = '#,##0_ ';
    } catch (e) {
        console.warn('셀 B63 설정 실패:', e);
    }

    // B64 셀
    try {
        const cell_B64 = worksheet.getCell('B64');
        cell_B64.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_B64.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_B64.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 B64 설정 실패:', e);
    }

    // B65 셀
    try {
        const cell_B65 = worksheet.getCell('B65');
        cell_B65.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_B65.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_B65.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 B65 설정 실패:', e);
    }

    // B66 셀
    try {
        const cell_B66 = worksheet.getCell('B66');
        cell_B66.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_B66.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_B66.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 B66 설정 실패:', e);
    }

    // B67 셀
    try {
        const cell_B67 = worksheet.getCell('B67');
        cell_B67.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_B67.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_B67.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 B67 설정 실패:', e);
    }

    // B68 셀
    try {
        const cell_B68 = worksheet.getCell('B68');
        cell_B68.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_B68.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_B68.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 B68 설정 실패:', e);
    }

    // B69 셀
    try {
        const cell_B69 = worksheet.getCell('B69');
        cell_B69.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_B69.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_B69.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 B69 설정 실패:', e);
    }

    // B7 셀
    try {
        const cell_B7 = worksheet.getCell('B7');
        cell_B7.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_B7.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_B7.border = { top: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 B7 설정 실패:', e);
    }

    // B70 셀
    try {
        const cell_B70 = worksheet.getCell('B70');
        cell_B70.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_B70.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_B70.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 B70 설정 실패:', e);
    }

    // B71 셀
    try {
        const cell_B71 = worksheet.getCell('B71');
        cell_B71.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_B71.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_B71.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 B71 설정 실패:', e);
    }

    // B72 셀
    try {
        const cell_B72 = worksheet.getCell('B72');
        cell_B72.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_B72.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_B72.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 B72 설정 실패:', e);
    }

    // B73 셀
    try {
        const cell_B73 = worksheet.getCell('B73');
        cell_B73.value = '특이사항';
        cell_B73.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_B73.alignment = { horizontal: 'center', vertical: 'center' };
        cell_B73.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_B73.numFmt = '#,##0\ "대"';
    } catch (e) {
        console.warn('셀 B73 설정 실패:', e);
    }

    // B74 셀
    try {
        const cell_B74 = worksheet.getCell('B74');
        cell_B74.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_B74.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_B74.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 B74 설정 실패:', e);
    }

    // B75 셀
    try {
        const cell_B75 = worksheet.getCell('B75');
        cell_B75.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_B75.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_B75.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 B75 설정 실패:', e);
    }

    // B76 셀
    try {
        const cell_B76 = worksheet.getCell('B76');
        cell_B76.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_B76.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_B76.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 B76 설정 실패:', e);
    }

    // B77 셀
    try {
        const cell_B77 = worksheet.getCell('B77');
        cell_B77.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_B77.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_B77.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 B77 설정 실패:', e);
    }

    // B78 셀
    try {
        const cell_B78 = worksheet.getCell('B78');
        cell_B78.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_B78.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_B78.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 B78 설정 실패:', e);
    }

    // B79 셀
    try {
        const cell_B79 = worksheet.getCell('B79');
        cell_B79.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_B79.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_B79.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 B79 설정 실패:', e);
    }

    // B8 셀
    try {
        const cell_B8 = worksheet.getCell('B8');
        cell_B8.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_B8.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_B8.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 B8 설정 실패:', e);
    }

    // B80 셀
    try {
        const cell_B80 = worksheet.getCell('B80');
        cell_B80.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_B80.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_B80.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 B80 설정 실패:', e);
    }

    // B81 셀
    try {
        const cell_B81 = worksheet.getCell('B81');
        cell_B81.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_B81.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_B81.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 B81 설정 실패:', e);
    }

    // B82 셀
    try {
        const cell_B82 = worksheet.getCell('B82');
        cell_B82.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_B82.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_B82.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 B82 설정 실패:', e);
    }

    // B83 셀
    try {
        const cell_B83 = worksheet.getCell('B83');
        cell_B83.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_B83.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_B83.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 B83 설정 실패:', e);
    }

    // B84 셀
    try {
        const cell_B84 = worksheet.getCell('B84');
        cell_B84.font = { name: 'LG스마트체 Regular', size: 8.0, color: { argb: 'FF000000' } };
        cell_B84.alignment = { horizontal: 'center', vertical: 'center' };
        cell_B84.border = { top: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 B84 설정 실패:', e);
    }

    // B9 셀
    try {
        const cell_B9 = worksheet.getCell('B9');
        cell_B9.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_B9.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_B9.border = { top: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 B9 설정 실패:', e);
    }

    // C17 셀
    try {
        const cell_C17 = worksheet.getCell('C17');
        cell_C17.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_C17.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_C17.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 C17 설정 실패:', e);
    }

    // C18 셀
    try {
        const cell_C18 = worksheet.getCell('C18');
        cell_C18.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_C18.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_C18.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 C18 설정 실패:', e);
    }

    // C19 셀
    try {
        const cell_C19 = worksheet.getCell('C19');
        cell_C19.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_C19.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_C19.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 C19 설정 실패:', e);
    }

    // C20 셀
    try {
        const cell_C20 = worksheet.getCell('C20');
        cell_C20.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_C20.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_C20.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 C20 설정 실패:', e);
    }

    // C21 셀
    try {
        const cell_C21 = worksheet.getCell('C21');
        cell_C21.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_C21.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_C21.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 C21 설정 실패:', e);
    }

    // C22 셀
    try {
        const cell_C22 = worksheet.getCell('C22');
        cell_C22.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_C22.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_C22.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 C22 설정 실패:', e);
    }

    // C23 셀
    try {
        const cell_C23 = worksheet.getCell('C23');
        cell_C23.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_C23.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_C23.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 C23 설정 실패:', e);
    }

    // C24 셀
    try {
        const cell_C24 = worksheet.getCell('C24');
        cell_C24.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_C24.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_C24.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 C24 설정 실패:', e);
    }

    // C25 셀
    try {
        const cell_C25 = worksheet.getCell('C25');
        cell_C25.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_C25.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_C25.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 C25 설정 실패:', e);
    }

    // C26 셀
    try {
        const cell_C26 = worksheet.getCell('C26');
        cell_C26.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_C26.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_C26.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 C26 설정 실패:', e);
    }

    // C27 셀
    try {
        const cell_C27 = worksheet.getCell('C27');
        cell_C27.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_C27.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_C27.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 C27 설정 실패:', e);
    }

    // C28 셀
    try {
        const cell_C28 = worksheet.getCell('C28');
        cell_C28.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_C28.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_C28.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 C28 설정 실패:', e);
    }

    // C29 셀
    try {
        const cell_C29 = worksheet.getCell('C29');
        cell_C29.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_C29.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_C29.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 C29 설정 실패:', e);
    }

    // C30 셀
    try {
        const cell_C30 = worksheet.getCell('C30');
        cell_C30.value = '공시지가 대비 담보율';
        cell_C30.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FFC00000' } };
        cell_C30.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        cell_C30.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_C30.numFmt = '0%';
    } catch (e) {
        console.warn('셀 C30 설정 실패:', e);
    }

    // C31 셀
    try {
        const cell_C31 = worksheet.getCell('C31');
        cell_C31.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_C31.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_C31.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 C31 설정 실패:', e);
    }

    // C32 셀
    try {
        const cell_C32 = worksheet.getCell('C32');
        cell_C32.value = '토지가격 적용';
        cell_C32.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_C32.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        cell_C32.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_C32.numFmt = '0%';
    } catch (e) {
        console.warn('셀 C32 설정 실패:', e);
    }

    // C33 셀
    try {
        const cell_C33 = worksheet.getCell('C33');
        cell_C33.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_C33.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_C33.border = { top: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 C33 설정 실패:', e);
    }

    // C40 셀
    try {
        const cell_C40 = worksheet.getCell('C40');
        cell_C40.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_C40.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_C40.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 C40 설정 실패:', e);
    }

    // C41 셀
    try {
        const cell_C41 = worksheet.getCell('C41');
        cell_C41.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_C41.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_C41.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 C41 설정 실패:', e);
    }

    // C42 셀
    try {
        const cell_C42 = worksheet.getCell('C42');
        cell_C42.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_C42.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_C42.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 C42 설정 실패:', e);
    }

    // C43 셀
    try {
        const cell_C43 = worksheet.getCell('C43');
        cell_C43.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_C43.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_C43.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 C43 설정 실패:', e);
    }

    // C44 셀
    try {
        const cell_C44 = worksheet.getCell('C44');
        cell_C44.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_C44.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_C44.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 C44 설정 실패:', e);
    }

    // C45 셀
    try {
        const cell_C45 = worksheet.getCell('C45');
        cell_C45.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_C45.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_C45.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 C45 설정 실패:', e);
    }

    // C46 셀
    try {
        const cell_C46 = worksheet.getCell('C46');
        cell_C46.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_C46.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_C46.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 C46 설정 실패:', e);
    }

    // C47 셀
    try {
        const cell_C47 = worksheet.getCell('C47');
        cell_C47.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_C47.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_C47.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 C47 설정 실패:', e);
    }

    // C48 셀
    try {
        const cell_C48 = worksheet.getCell('C48');
        cell_C48.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_C48.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_C48.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 C48 설정 실패:', e);
    }

    // C49 셀
    try {
        const cell_C49 = worksheet.getCell('C49');
        cell_C49.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_C49.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_C49.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 C49 설정 실패:', e);
    }

    // C50 셀
    try {
        const cell_C50 = worksheet.getCell('C50');
        cell_C50.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_C50.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_C50.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 C50 설정 실패:', e);
    }

    // C51 셀
    try {
        const cell_C51 = worksheet.getCell('C51');
        cell_C51.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_C51.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_C51.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 C51 설정 실패:', e);
    }

    // C52 셀
    try {
        const cell_C52 = worksheet.getCell('C52');
        cell_C52.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_C52.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_C52.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 C52 설정 실패:', e);
    }

    // C53 셀
    try {
        const cell_C53 = worksheet.getCell('C53');
        cell_C53.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_C53.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_C53.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 C53 설정 실패:', e);
    }

    // C54 셀
    try {
        const cell_C54 = worksheet.getCell('C54');
        cell_C54.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_C54.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_C54.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 C54 설정 실패:', e);
    }

    // C55 셀
    try {
        const cell_C55 = worksheet.getCell('C55');
        cell_C55.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_C55.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_C55.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 C55 설정 실패:', e);
    }

    // C56 셀
    try {
        const cell_C56 = worksheet.getCell('C56');
        cell_C56.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_C56.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_C56.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 C56 설정 실패:', e);
    }

    // C57 셀
    try {
        const cell_C57 = worksheet.getCell('C57');
        cell_C57.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_C57.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_C57.border = { top: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 C57 설정 실패:', e);
    }

    // C58 셀
    try {
        const cell_C58 = worksheet.getCell('C58');
        cell_C58.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_C58.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_C58.border = { bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 C58 설정 실패:', e);
    }

    // C59 셀
    try {
        const cell_C59 = worksheet.getCell('C59');
        cell_C59.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_C59.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_C59.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 C59 설정 실패:', e);
    }

    // C6 셀
    try {
        const cell_C6 = worksheet.getCell('C6');
        cell_C6.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_C6.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_C6.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 C6 설정 실패:', e);
    }

    // C60 셀
    try {
        const cell_C60 = worksheet.getCell('C60');
        cell_C60.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_C60.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_C60.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 C60 설정 실패:', e);
    }

    // C61 셀
    try {
        const cell_C61 = worksheet.getCell('C61');
        cell_C61.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_C61.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_C61.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 C61 설정 실패:', e);
    }

    // C62 셀
    try {
        const cell_C62 = worksheet.getCell('C62');
        cell_C62.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_C62.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_C62.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 C62 설정 실패:', e);
    }

    // C7 셀
    try {
        const cell_C7 = worksheet.getCell('C7');
        cell_C7.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_C7.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_C7.border = { top: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 C7 설정 실패:', e);
    }

    // C72 셀
    try {
        const cell_C72 = worksheet.getCell('C72');
        cell_C72.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_C72.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_C72.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 C72 설정 실패:', e);
    }

    // C73 셀
    try {
        const cell_C73 = worksheet.getCell('C73');
        cell_C73.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_C73.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_C73.border = { top: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 C73 설정 실패:', e);
    }

    // C8 셀
    try {
        const cell_C8 = worksheet.getCell('C8');
        cell_C8.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_C8.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_C8.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 C8 설정 실패:', e);
    }

    // C83 셀
    try {
        const cell_C83 = worksheet.getCell('C83');
        cell_C83.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_C83.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_C83.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 C83 설정 실패:', e);
    }

    // C84 셀
    try {
        const cell_C84 = worksheet.getCell('C84');
        cell_C84.font = { name: 'LG스마트체 Regular', size: 8.0, color: { argb: 'FF000000' } };
        cell_C84.alignment = { horizontal: 'center', vertical: 'center' };
        cell_C84.border = { top: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 C84 설정 실패:', e);
    }

    // C9 셀
    try {
        const cell_C9 = worksheet.getCell('C9');
        cell_C9.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_C9.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_C9.border = { top: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 C9 설정 실패:', e);
    }

    // D10 셀
    try {
        const cell_D10 = worksheet.getCell('D10');
        cell_D10.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D10.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D10.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D10 설정 실패:', e);
    }

    // D11 셀
    try {
        const cell_D11 = worksheet.getCell('D11');
        cell_D11.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D11.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D11.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D11 설정 실패:', e);
    }

    // D12 셀
    try {
        const cell_D12 = worksheet.getCell('D12');
        cell_D12.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D12.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D12.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D12 설정 실패:', e);
    }

    // D13 셀
    try {
        const cell_D13 = worksheet.getCell('D13');
        cell_D13.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D13.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D13.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D13 설정 실패:', e);
    }

    // D14 셀
    try {
        const cell_D14 = worksheet.getCell('D14');
        cell_D14.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D14.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D14.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D14 설정 실패:', e);
    }

    // D15 셀
    try {
        const cell_D15 = worksheet.getCell('D15');
        cell_D15.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D15.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D15.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D15 설정 실패:', e);
    }

    // D16 셀
    try {
        const cell_D16 = worksheet.getCell('D16');
        cell_D16.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D16.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D16.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D16 설정 실패:', e);
    }

    // D17 셀
    try {
        const cell_D17 = worksheet.getCell('D17');
        cell_D17.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D17.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D17.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D17 설정 실패:', e);
    }

    // D18 셀
    try {
        const cell_D18 = worksheet.getCell('D18');
        cell_D18.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D18.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D18.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D18 설정 실패:', e);
    }

    // D19 셀
    try {
        const cell_D19 = worksheet.getCell('D19');
        cell_D19.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D19.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D19.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D19 설정 실패:', e);
    }

    // D20 셀
    try {
        const cell_D20 = worksheet.getCell('D20');
        cell_D20.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D20.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D20.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D20 설정 실패:', e);
    }

    // D21 셀
    try {
        const cell_D21 = worksheet.getCell('D21');
        cell_D21.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D21.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D21.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D21 설정 실패:', e);
    }

    // D22 셀
    try {
        const cell_D22 = worksheet.getCell('D22');
        cell_D22.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D22.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D22.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D22 설정 실패:', e);
    }

    // D23 셀
    try {
        const cell_D23 = worksheet.getCell('D23');
        cell_D23.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D23.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D23.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D23 설정 실패:', e);
    }

    // D24 셀
    try {
        const cell_D24 = worksheet.getCell('D24');
        cell_D24.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D24.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D24.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D24 설정 실패:', e);
    }

    // D25 셀
    try {
        const cell_D25 = worksheet.getCell('D25');
        cell_D25.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D25.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D25.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D25 설정 실패:', e);
    }

    // D26 셀
    try {
        const cell_D26 = worksheet.getCell('D26');
        cell_D26.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D26.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D26.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D26 설정 실패:', e);
    }

    // D27 셀
    try {
        const cell_D27 = worksheet.getCell('D27');
        cell_D27.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D27.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D27.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D27 설정 실패:', e);
    }

    // D28 셀
    try {
        const cell_D28 = worksheet.getCell('D28');
        cell_D28.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D28.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D28.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D28 설정 실패:', e);
    }

    // D29 셀
    try {
        const cell_D29 = worksheet.getCell('D29');
        cell_D29.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D29.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D29.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D29 설정 실패:', e);
    }

    // D30 셀
    try {
        const cell_D30 = worksheet.getCell('D30');
        cell_D30.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D30.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D30.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D30 설정 실패:', e);
    }

    // D31 셀
    try {
        const cell_D31 = worksheet.getCell('D31');
        cell_D31.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D31.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D31.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D31 설정 실패:', e);
    }

    // D32 셀
    try {
        const cell_D32 = worksheet.getCell('D32');
        cell_D32.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D32.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D32.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D32 설정 실패:', e);
    }

    // D33 셀
    try {
        const cell_D33 = worksheet.getCell('D33');
        cell_D33.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D33.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D33.border = { top: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D33 설정 실패:', e);
    }

    // D40 셀
    try {
        const cell_D40 = worksheet.getCell('D40');
        cell_D40.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D40.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D40.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D40 설정 실패:', e);
    }

    // D41 셀
    try {
        const cell_D41 = worksheet.getCell('D41');
        cell_D41.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D41.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D41.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D41 설정 실패:', e);
    }

    // D42 셀
    try {
        const cell_D42 = worksheet.getCell('D42');
        cell_D42.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D42.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D42.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D42 설정 실패:', e);
    }

    // D43 셀
    try {
        const cell_D43 = worksheet.getCell('D43');
        cell_D43.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D43.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D43.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D43 설정 실패:', e);
    }

    // D44 셀
    try {
        const cell_D44 = worksheet.getCell('D44');
        cell_D44.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D44.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D44.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D44 설정 실패:', e);
    }

    // D45 셀
    try {
        const cell_D45 = worksheet.getCell('D45');
        cell_D45.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D45.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D45.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D45 설정 실패:', e);
    }

    // D46 셀
    try {
        const cell_D46 = worksheet.getCell('D46');
        cell_D46.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D46.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D46.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D46 설정 실패:', e);
    }

    // D47 셀
    try {
        const cell_D47 = worksheet.getCell('D47');
        cell_D47.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D47.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D47.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D47 설정 실패:', e);
    }

    // D48 셀
    try {
        const cell_D48 = worksheet.getCell('D48');
        cell_D48.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D48.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D48.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D48 설정 실패:', e);
    }

    // D49 셀
    try {
        const cell_D49 = worksheet.getCell('D49');
        cell_D49.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D49.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D49.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D49 설정 실패:', e);
    }

    // D50 셀
    try {
        const cell_D50 = worksheet.getCell('D50');
        cell_D50.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D50.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D50.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D50 설정 실패:', e);
    }

    // D51 셀
    try {
        const cell_D51 = worksheet.getCell('D51');
        cell_D51.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D51.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D51.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D51 설정 실패:', e);
    }

    // D52 셀
    try {
        const cell_D52 = worksheet.getCell('D52');
        cell_D52.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D52.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D52.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D52 설정 실패:', e);
    }

    // D53 셀
    try {
        const cell_D53 = worksheet.getCell('D53');
        cell_D53.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D53.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D53.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D53 설정 실패:', e);
    }

    // D54 셀
    try {
        const cell_D54 = worksheet.getCell('D54');
        cell_D54.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D54.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D54.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D54 설정 실패:', e);
    }

    // D55 셀
    try {
        const cell_D55 = worksheet.getCell('D55');
        cell_D55.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D55.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D55.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D55 설정 실패:', e);
    }

    // D56 셀
    try {
        const cell_D56 = worksheet.getCell('D56');
        cell_D56.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D56.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D56.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D56 설정 실패:', e);
    }

    // D57 셀
    try {
        const cell_D57 = worksheet.getCell('D57');
        cell_D57.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D57.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D57.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D57 설정 실패:', e);
    }

    // D58 셀
    try {
        const cell_D58 = worksheet.getCell('D58');
        cell_D58.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D58.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D58.border = { bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D58 설정 실패:', e);
    }

    // D59 셀
    try {
        const cell_D59 = worksheet.getCell('D59');
        cell_D59.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D59.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D59.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D59 설정 실패:', e);
    }

    // D6 셀
    try {
        const cell_D6 = worksheet.getCell('D6');
        cell_D6.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D6.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D6.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D6 설정 실패:', e);
    }

    // D60 셀
    try {
        const cell_D60 = worksheet.getCell('D60');
        cell_D60.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D60.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D60.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D60 설정 실패:', e);
    }

    // D61 셀
    try {
        const cell_D61 = worksheet.getCell('D61');
        cell_D61.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D61.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D61.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D61 설정 실패:', e);
    }

    // D62 셀
    try {
        const cell_D62 = worksheet.getCell('D62');
        cell_D62.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D62.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D62.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D62 설정 실패:', e);
    }

    // D63 셀
    try {
        const cell_D63 = worksheet.getCell('D63');
        cell_D63.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D63.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D63.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D63 설정 실패:', e);
    }

    // D64 셀
    try {
        const cell_D64 = worksheet.getCell('D64');
        cell_D64.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D64.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D64.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D64 설정 실패:', e);
    }

    // D65 셀
    try {
        const cell_D65 = worksheet.getCell('D65');
        cell_D65.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D65.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D65.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D65 설정 실패:', e);
    }

    // D66 셀
    try {
        const cell_D66 = worksheet.getCell('D66');
        cell_D66.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D66.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D66.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D66 설정 실패:', e);
    }

    // D67 셀
    try {
        const cell_D67 = worksheet.getCell('D67');
        cell_D67.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D67.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D67.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D67 설정 실패:', e);
    }

    // D68 셀
    try {
        const cell_D68 = worksheet.getCell('D68');
        cell_D68.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D68.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D68.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D68 설정 실패:', e);
    }

    // D69 셀
    try {
        const cell_D69 = worksheet.getCell('D69');
        cell_D69.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D69.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D69.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D69 설정 실패:', e);
    }

    // D7 셀
    try {
        const cell_D7 = worksheet.getCell('D7');
        cell_D7.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D7.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D7.border = { top: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D7 설정 실패:', e);
    }

    // D70 셀
    try {
        const cell_D70 = worksheet.getCell('D70');
        cell_D70.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D70.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D70.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D70 설정 실패:', e);
    }

    // D71 셀
    try {
        const cell_D71 = worksheet.getCell('D71');
        cell_D71.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D71.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D71.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D71 설정 실패:', e);
    }

    // D72 셀
    try {
        const cell_D72 = worksheet.getCell('D72');
        cell_D72.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D72.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D72.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D72 설정 실패:', e);
    }

    // D73 셀
    try {
        const cell_D73 = worksheet.getCell('D73');
        cell_D73.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D73.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D73.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D73 설정 실패:', e);
    }

    // D74 셀
    try {
        const cell_D74 = worksheet.getCell('D74');
        cell_D74.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D74.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D74.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D74 설정 실패:', e);
    }

    // D75 셀
    try {
        const cell_D75 = worksheet.getCell('D75');
        cell_D75.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D75.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D75.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D75 설정 실패:', e);
    }

    // D76 셀
    try {
        const cell_D76 = worksheet.getCell('D76');
        cell_D76.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D76.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D76.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D76 설정 실패:', e);
    }

    // D77 셀
    try {
        const cell_D77 = worksheet.getCell('D77');
        cell_D77.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D77.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D77.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D77 설정 실패:', e);
    }

    // D78 셀
    try {
        const cell_D78 = worksheet.getCell('D78');
        cell_D78.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D78.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D78.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D78 설정 실패:', e);
    }

    // D79 셀
    try {
        const cell_D79 = worksheet.getCell('D79');
        cell_D79.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D79.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D79.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D79 설정 실패:', e);
    }

    // D8 셀
    try {
        const cell_D8 = worksheet.getCell('D8');
        cell_D8.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D8.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D8.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D8 설정 실패:', e);
    }

    // D80 셀
    try {
        const cell_D80 = worksheet.getCell('D80');
        cell_D80.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D80.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D80.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D80 설정 실패:', e);
    }

    // D81 셀
    try {
        const cell_D81 = worksheet.getCell('D81');
        cell_D81.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D81.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D81.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D81 설정 실패:', e);
    }

    // D82 셀
    try {
        const cell_D82 = worksheet.getCell('D82');
        cell_D82.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D82.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D82.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D82 설정 실패:', e);
    }

    // D83 셀
    try {
        const cell_D83 = worksheet.getCell('D83');
        cell_D83.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D83.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D83.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D83 설정 실패:', e);
    }

    // D84 셀
    try {
        const cell_D84 = worksheet.getCell('D84');
        cell_D84.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_D84.alignment = { horizontal: 'center', vertical: 'center' };
        cell_D84.border = { top: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D84 설정 실패:', e);
    }

    // D9 셀
    try {
        const cell_D9 = worksheet.getCell('D9');
        cell_D9.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_D9.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_D9.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 D9 설정 실패:', e);
    }

    // E10 셀
    try {
        const cell_E10 = worksheet.getCell('E10');
        cell_E10.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_E10.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_E10.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 E10 설정 실패:', e);
    }

    // E106 셀
    try {
        const cell_E106 = worksheet.getCell('E106');
        cell_E106.font = { name: 'LG스마트체 Regular', size: 10.0, bold: true, color: { argb: 'FF000000' } };
        cell_E106.alignment = { horizontal: 'center', vertical: 'center' };
        cell_E106.numFmt = '_-* #,##0_-;\-* #,##0_-;_-* "-"_-;_-@_-';
    } catch (e) {
        console.warn('셀 E106 설정 실패:', e);
    }

    // E11 셀
    try {
        const cell_E11 = worksheet.getCell('E11');
        cell_E11.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_E11.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_E11.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 E11 설정 실패:', e);
    }

    // E12 셀
    try {
        const cell_E12 = worksheet.getCell('E12');
        cell_E12.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_E12.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_E12.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 E12 설정 실패:', e);
    }

    // E13 셀
    try {
        const cell_E13 = worksheet.getCell('E13');
        cell_E13.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_E13.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_E13.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 E13 설정 실패:', e);
    }

    // E14 셀
    try {
        const cell_E14 = worksheet.getCell('E14');
        cell_E14.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_E14.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_E14.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 E14 설정 실패:', e);
    }

    // E15 셀
    try {
        const cell_E15 = worksheet.getCell('E15');
        cell_E15.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_E15.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_E15.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 E15 설정 실패:', e);
    }

    // E16 셀
    try {
        const cell_E16 = worksheet.getCell('E16');
        cell_E16.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_E16.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_E16.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 E16 설정 실패:', e);
    }

    // E17 셀
    try {
        const cell_E17 = worksheet.getCell('E17');
        cell_E17.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_E17.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_E17.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 E17 설정 실패:', e);
    }

    // E18 셀
    try {
        const cell_E18 = worksheet.getCell('E18');
        cell_E18.value = '성남시 분당구 대왕판교로 660';
        cell_E18.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_E18.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        cell_E18.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_E18.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 E18 설정 실패:', e);
    }

    // E19 셀
    try {
        const cell_E19 = worksheet.getCell('E19');
        cell_E19.value = '신분당선, 경강선 판교역 버스 10분';
        cell_E19.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_E19.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        cell_E19.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_E19.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 E19 설정 실패:', e);
    }

    // E20 셀
    try {
        const cell_E20 = worksheet.getCell('E20');
        cell_E20.value = '2012';
        cell_E20.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_E20.alignment = { horizontal: 'center', vertical: 'center' };
        cell_E20.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_E20.numFmt = '###0"년"';
    } catch (e) {
        console.warn('셀 E20 설정 실패:', e);
    }

    // E21 셀
    try {
        const cell_E21 = worksheet.getCell('E21');
        cell_E21.value = '125';
        cell_E21.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_E21.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        cell_E21.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_E21.numFmt = '##"F / B"#';
    } catch (e) {
        console.warn('셀 E21 설정 실패:', e);
    }

    // E22 셀
    try {
        const cell_E22 = worksheet.getCell('E22');
        cell_E22.value = '41280.82';
        cell_E22.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_E22.alignment = { horizontal: 'center', vertical: 'center' };
        cell_E22.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_E22.numFmt = '#,##0\ "평"';
    } catch (e) {
        console.warn('셀 E22 설정 실패:', e);
    }

    // E23 셀
    try {
        const cell_E23 = worksheet.getCell('E23');
        cell_E23.value = '1004';
        cell_E23.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_E23.alignment = { horizontal: 'center', vertical: 'center' };
        cell_E23.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_E23.numFmt = '#,##0\ "평"';
    } catch (e) {
        console.warn('셀 E23 설정 실패:', e);
    }

    // E24 셀
    try {
        const cell_E24 = worksheet.getCell('E24');
        cell_E24.value = '0.4639';
        cell_E24.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_E24.alignment = { horizontal: 'center', vertical: 'center' };
        cell_E24.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_E24.numFmt = '0.00%';
    } catch (e) {
        console.warn('셀 E24 설정 실패:', e);
    }

    // E25 셀
    try {
        const cell_E25 = worksheet.getCell('E25');
        cell_E25.value = { formula: '=G25*0.3025' };
        cell_E25.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_E25.alignment = { horizontal: 'center', vertical: 'center' };
        cell_E25.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_E25.numFmt = '#,##0\ "평"';
    } catch (e) {
        console.warn('셀 E25 설정 실패:', e);
    }

    // E26 셀
    try {
        const cell_E26 = worksheet.getCell('E26');
        cell_E26.value = '에코자산개발주식회사';
        cell_E26.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_E26.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        cell_E26.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_E26.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 E26 설정 실패:', e);
    }

    // E27 셀
    try {
        const cell_E27 = worksheet.getCell('E27');
        cell_E27.value = '전세권 설정 가능';
        cell_E27.font = { name: 'LG스마트체 Regular', size: 9.0, bold: true, color: { argb: 'FFC00000' } };
        cell_E27.alignment = { horizontal: 'center', vertical: 'center' };
        cell_E27.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_E27.numFmt = '0.00%';
    } catch (e) {
        console.warn('셀 E27 설정 실패:', e);
    }

    // E28 셀
    try {
        const cell_E28 = worksheet.getCell('E28');
        cell_E28.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_E28.alignment = { horizontal: 'center', vertical: 'center' };
        cell_E28.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_E28.numFmt = '#,##0.000\ "평"';
    } catch (e) {
        console.warn('셀 E28 설정 실패:', e);
    }

    // E29 셀
    try {
        const cell_E29 = worksheet.getCell('E29');
        cell_E29.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_E29.alignment = { horizontal: 'center', vertical: 'center' };
        cell_E29.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_E29.numFmt = '_-* #,##0_-;\-* #,##0_-;_-* "-"_-;_-@_-';
    } catch (e) {
        console.warn('셀 E29 설정 실패:', e);
    }

    // E30 셀
    try {
        const cell_E30 = worksheet.getCell('E30');
        cell_E30.value = { formula: '=E29/E32' };
        cell_E30.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FFC00000' } };
        cell_E30.alignment = { horizontal: 'center', vertical: 'center' };
        cell_E30.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_E30.numFmt = '0.00%';
    } catch (e) {
        console.warn('셀 E30 설정 실패:', e);
    }

    // E31 셀
    try {
        const cell_E31 = worksheet.getCell('E31');
        cell_E31.value = '5995000';
        cell_E31.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_E31.alignment = { horizontal: 'center', vertical: 'center' };
        cell_E31.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_E31.numFmt = '#,###"원/㎡"';
    } catch (e) {
        console.warn('셀 E31 설정 실패:', e);
    }

    // E32 셀
    try {
        const cell_E32 = worksheet.getCell('E32');
        cell_E32.value = { formula: '=E31*G25' };
        cell_E32.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_E32.alignment = { horizontal: 'center', vertical: 'center' };
        cell_E32.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_E32.numFmt = '_-* #,##0_-;\-* #,##0_-;_-* "-"_-;_-@_-';
    } catch (e) {
        console.warn('셀 E32 설정 실패:', e);
    }

    // E33 셀
    try {
        const cell_E33 = worksheet.getCell('E33');
        cell_E33.value = '층';
        cell_E33.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_E33.alignment = { horizontal: 'center', vertical: 'center' };
        cell_E33.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_E33.numFmt = '@';
    } catch (e) {
        console.warn('셀 E33 설정 실패:', e);
    }

    // E34 셀
    try {
        const cell_E34 = worksheet.getCell('E34');
        cell_E34.value = '4';
        cell_E34.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF0000FF' } };
        cell_E34.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEECEC' } };
        cell_E34.alignment = { horizontal: 'center', vertical: 'center' };
        cell_E34.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_E34.numFmt = '#"층"';
    } catch (e) {
        console.warn('셀 E34 설정 실패:', e);
    }

    // E35 셀
    try {
        const cell_E35 = worksheet.getCell('E35');
        cell_E35.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF0000FF' } };
        cell_E35.alignment = { horizontal: 'center', vertical: 'center' };
        cell_E35.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_E35.numFmt = '#"층"';
    } catch (e) {
        console.warn('셀 E35 설정 실패:', e);
    }

    // E36 셀
    try {
        const cell_E36 = worksheet.getCell('E36');
        cell_E36.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_E36.alignment = { horizontal: 'center', vertical: 'center' };
        cell_E36.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_E36.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 E36 설정 실패:', e);
    }

    // E37 셀
    try {
        const cell_E37 = worksheet.getCell('E37');
        cell_E37.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_E37.alignment = { horizontal: 'center', vertical: 'center' };
        cell_E37.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_E37.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 E37 설정 실패:', e);
    }

    // E38 셀
    try {
        const cell_E38 = worksheet.getCell('E38');
        cell_E38.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_E38.alignment = { horizontal: 'center', vertical: 'center' };
        cell_E38.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_E38.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 E38 설정 실패:', e);
    }

    // E39 셀
    try {
        const cell_E39 = worksheet.getCell('E39');
        cell_E39.value = '소계';
        cell_E39.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_E39.alignment = { horizontal: 'center', vertical: 'center' };
        cell_E39.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_E39.numFmt = '@';
    } catch (e) {
        console.warn('셀 E39 설정 실패:', e);
    }

    // E40 셀
    try {
        const cell_E40 = worksheet.getCell('E40');
        cell_E40.value = '2025.7~2027.6 (12개월)';
        cell_E40.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_E40.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        cell_E40.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_E40.numFmt = '#"개월 계약 가능"';
    } catch (e) {
        console.warn('셀 E40 설정 실패:', e);
    }

    // E41 셀
    try {
        const cell_E41 = worksheet.getCell('E41');
        cell_E41.value = '즉시';
        cell_E41.font = { name: 'LG스마트체 Regular', size: 9.0, bold: true, color: { argb: 'FF000000' } };
        cell_E41.alignment = { horizontal: 'center', vertical: 'center' };
        cell_E41.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_E41.numFmt = '#,##0\ "원"';
    } catch (e) {
        console.warn('셀 E41 설정 실패:', e);
    }

    // E42 셀
    try {
        const cell_E42 = worksheet.getCell('E42');
        cell_E42.value = '4층 일부';
        cell_E42.font = { name: 'LG스마트체 Regular', size: 9.0, bold: true, color: { argb: 'FF000000' } };
        cell_E42.alignment = { horizontal: 'center', vertical: 'center' };
        cell_E42.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_E42.numFmt = '#,##0\ "층"';
    } catch (e) {
        console.warn('셀 E42 설정 실패:', e);
    }

    // E43 셀
    try {
        const cell_E43 = worksheet.getCell('E43');
        cell_E43.value = { formula: '=SUM(F34:F35)' };
        cell_E43.font = { name: 'LG스마트체 Regular', size: 9.0, bold: true, color: { argb: 'FFC00000' } };
        cell_E43.alignment = { horizontal: 'center', vertical: 'center' };
        cell_E43.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_E43.numFmt = '#,##0\ "평"';
    } catch (e) {
        console.warn('셀 E43 설정 실패:', e);
    }

    // E44 셀
    try {
        const cell_E44 = worksheet.getCell('E44');
        cell_E44.value = { formula: '=SUM(G34:G35)' };
        cell_E44.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_E44.alignment = { horizontal: 'center', vertical: 'center' };
        cell_E44.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_E44.numFmt = '#,##0\ "평"';
    } catch (e) {
        console.warn('셀 E44 설정 실패:', e);
    }

    // E45 셀
    try {
        const cell_E45 = worksheet.getCell('E45');
        cell_E45.value = '1048752';
        cell_E45.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_E45.alignment = { horizontal: 'center', vertical: 'center' };
        cell_E45.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_E45.numFmt = '"@"#,###';
    } catch (e) {
        console.warn('셀 E45 설정 실패:', e);
    }

    // E46 셀
    try {
        const cell_E46 = worksheet.getCell('E46');
        cell_E46.value = '104875';
        cell_E46.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_E46.alignment = { horizontal: 'center', vertical: 'center' };
        cell_E46.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_E46.numFmt = '"@"#,###';
    } catch (e) {
        console.warn('셀 E46 설정 실패:', e);
    }

    // E47 셀
    try {
        const cell_E47 = worksheet.getCell('E47');
        cell_E47.value = '6000';
        cell_E47.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_E47.alignment = { horizontal: 'center', vertical: 'center' };
        cell_E47.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_E47.numFmt = '"@"#,###\+"실""비""별""도"';
    } catch (e) {
        console.warn('셀 E47 설정 실패:', e);
    }

    // E48 셀
    try {
        const cell_E48 = worksheet.getCell('E48');
        cell_E48.value = { formula: '=E46*(12-E49)/12' };
        cell_E48.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_E48.alignment = { horizontal: 'center', vertical: 'center' };
        cell_E48.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_E48.numFmt = '"@"#,###';
    } catch (e) {
        console.warn('셀 E48 설정 실패:', e);
    }

    // E49 셀
    try {
        const cell_E49 = worksheet.getCell('E49');
        cell_E49.value = '1';
        cell_E49.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_E49.alignment = { horizontal: 'center', vertical: 'center' };
        cell_E49.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_E49.numFmt = '#0.0"개월"';
    } catch (e) {
        console.warn('셀 E49 설정 실패:', e);
    }

    // E50 셀
    try {
        const cell_E50 = worksheet.getCell('E50');
        cell_E50.value = { formula: '=E45*E44' };
        cell_E50.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_E50.alignment = { horizontal: 'center', vertical: 'center' };
        cell_E50.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_E50.numFmt = '#,##0\ "원"';
    } catch (e) {
        console.warn('셀 E50 설정 실패:', e);
    }

    // E51 셀
    try {
        const cell_E51 = worksheet.getCell('E51');
        cell_E51.value = { formula: '=E46*E44' };
        cell_E51.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_E51.alignment = { horizontal: 'center', vertical: 'center' };
        cell_E51.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_E51.numFmt = '#,##0\ "원"';
    } catch (e) {
        console.warn('셀 E51 설정 실패:', e);
    }

    // E52 셀
    try {
        const cell_E52 = worksheet.getCell('E52');
        cell_E52.value = { formula: '=E47*E44' };
        cell_E52.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_E52.alignment = { horizontal: 'center', vertical: 'center' };
        cell_E52.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_E52.numFmt = '#,##0\ "원"';
    } catch (e) {
        console.warn('셀 E52 설정 실패:', e);
    }

    // E53 셀
    try {
        const cell_E53 = worksheet.getCell('E53');
        cell_E53.value = '실비 관리비: 전기세, 수도세  별도 부과\n(예상 수광비 약 4천원대)';
        cell_E53.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FFC00000' } };
        cell_E53.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        cell_E53.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_E53.numFmt = '#,##0\ "원"';
    } catch (e) {
        console.warn('셀 E53 설정 실패:', e);
    }

    // E54 셀
    try {
        const cell_E54 = worksheet.getCell('E54');
        cell_E54.value = { formula: '=E51' };
        cell_E54.font = { name: 'LG스마트체 Regular', size: 9.0, bold: true, color: { argb: 'FFC00000' } };
        cell_E54.alignment = { horizontal: 'center', vertical: 'center' };
        cell_E54.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_E54.numFmt = '#,##0\ "원"';
    } catch (e) {
        console.warn('셀 E54 설정 실패:', e);
    }

    // E55 셀
    try {
        const cell_E55 = worksheet.getCell('E55');
        cell_E55.value = { formula: '=E54*21' };
        cell_E55.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_E55.alignment = { horizontal: 'center', vertical: 'center' };
        cell_E55.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_E55.numFmt = '#,##0\ "원"';
    } catch (e) {
        console.warn('셀 E55 설정 실패:', e);
    }

    // E56 셀
    try {
        const cell_E56 = worksheet.getCell('E56');
        cell_E56.value = '협의';
        cell_E56.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_E56.alignment = { horizontal: 'center', vertical: 'center' };
        cell_E56.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_E56.numFmt = '0.#"개월"';
    } catch (e) {
        console.warn('셀 E56 설정 실패:', e);
    }

    // E57 셀
    try {
        const cell_E57 = worksheet.getCell('E57');
        cell_E57.value = '미제공';
        cell_E57.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_E57.alignment = { horizontal: 'center', vertical: 'center' };
        cell_E57.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_E57.numFmt = '"총액 "##,##0"원"';
    } catch (e) {
        console.warn('셀 E57 설정 실패:', e);
    }

    // E58 셀
    try {
        const cell_E58 = worksheet.getCell('E58');
        cell_E58.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_E58.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_E58.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 E58 설정 실패:', e);
    }

    // E59 셀
    try {
        const cell_E59 = worksheet.getCell('E59');
        cell_E59.value = '1023';
        cell_E59.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_E59.alignment = { horizontal: 'center', vertical: 'center' };
        cell_E59.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_E59.numFmt = '#\ "대"';
    } catch (e) {
        console.warn('셀 E59 설정 실패:', e);
    }

    // E6 셀
    try {
        const cell_E6 = worksheet.getCell('E6');
        cell_E6.value = '판교역';
        cell_E6.font = { name: 'LG스마트체 Regular', size: 10.0, color: { argb: 'FF000000' } };
        cell_E6.alignment = { horizontal: 'center', vertical: 'center' };
        cell_E6.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_E6.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 E6 설정 실패:', e);
    }

    // E60 셀
    try {
        const cell_E60 = worksheet.getCell('E60');
        cell_E60.value = '80';
        cell_E60.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_E60.alignment = { horizontal: 'center', vertical: 'center' };
        cell_E60.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_E60.numFmt = '"임대면적"\ #"평당 1대"';
    } catch (e) {
        console.warn('셀 E60 설정 실패:', e);
    }

    // E61 셀
    try {
        const cell_E61 = worksheet.getCell('E61');
        cell_E61.value = { formula: '=E44/E60' };
        cell_E61.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_E61.alignment = { horizontal: 'center', vertical: 'center' };
        cell_E61.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_E61.numFmt = '#,##0.0\ "대"';
    } catch (e) {
        console.warn('셀 E61 설정 실패:', e);
    }

    // E62 셀
    try {
        const cell_E62 = worksheet.getCell('E62');
        cell_E62.value = '협의';
        cell_E62.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_E62.alignment = { horizontal: 'center', vertical: 'center' };
        cell_E62.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_E62.numFmt = '"월"#"만원/대"';
    } catch (e) {
        console.warn('셀 E62 설정 실패:', e);
    }

    // E63 셀
    try {
        const cell_E63 = worksheet.getCell('E63');
        cell_E63.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_E63.alignment = { horizontal: 'left', vertical: 'center', wrapText: true };
        cell_E63.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_E63.numFmt = '#,##0_ ';
    } catch (e) {
        console.warn('셀 E63 설정 실패:', e);
    }

    // E64 셀
    try {
        const cell_E64 = worksheet.getCell('E64');
        cell_E64.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_E64.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_E64.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 E64 설정 실패:', e);
    }

    // E65 셀
    try {
        const cell_E65 = worksheet.getCell('E65');
        cell_E65.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_E65.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_E65.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 E65 설정 실패:', e);
    }

    // E66 셀
    try {
        const cell_E66 = worksheet.getCell('E66');
        cell_E66.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_E66.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_E66.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 E66 설정 실패:', e);
    }

    // E67 셀
    try {
        const cell_E67 = worksheet.getCell('E67');
        cell_E67.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_E67.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_E67.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 E67 설정 실패:', e);
    }

    // E68 셀
    try {
        const cell_E68 = worksheet.getCell('E68');
        cell_E68.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_E68.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_E68.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 E68 설정 실패:', e);
    }

    // E69 셀
    try {
        const cell_E69 = worksheet.getCell('E69');
        cell_E69.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_E69.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_E69.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 E69 설정 실패:', e);
    }

    // E7 셀
    try {
        const cell_E7 = worksheet.getCell('E7');
        cell_E7.font = { name: 'LG스마트체 Regular', size: 10.0, color: { argb: 'FF000000' } };
        cell_E7.alignment = { horizontal: 'center', vertical: 'center' };
        cell_E7.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_E7.numFmt = '0_);[Red]\(0\)';
    } catch (e) {
        console.warn('셀 E7 설정 실패:', e);
    }

    // E70 셀
    try {
        const cell_E70 = worksheet.getCell('E70');
        cell_E70.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_E70.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_E70.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 E70 설정 실패:', e);
    }

    // E71 셀
    try {
        const cell_E71 = worksheet.getCell('E71');
        cell_E71.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_E71.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_E71.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 E71 설정 실패:', e);
    }

    // E72 셀
    try {
        const cell_E72 = worksheet.getCell('E72');
        cell_E72.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_E72.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_E72.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 E72 설정 실패:', e);
    }

    // E73 셀
    try {
        const cell_E73 = worksheet.getCell('E73');
        cell_E73.value = ' - 현재 4층(410~412호) 일부 즉시 가능\n - Rent free 1개월 제공\n - 공사기간 협의 필요\n - 실비 별도 (예상 수광비 실비: 4천원)';
        cell_E73.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_E73.alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
        cell_E73.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_E73.numFmt = '#,##0\ "대"';
    } catch (e) {
        console.warn('셀 E73 설정 실패:', e);
    }

    // E74 셀
    try {
        const cell_E74 = worksheet.getCell('E74');
        cell_E74.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_E74.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_E74.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 E74 설정 실패:', e);
    }

    // E75 셀
    try {
        const cell_E75 = worksheet.getCell('E75');
        cell_E75.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_E75.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_E75.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 E75 설정 실패:', e);
    }

    // E76 셀
    try {
        const cell_E76 = worksheet.getCell('E76');
        cell_E76.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_E76.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_E76.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 E76 설정 실패:', e);
    }

    // E77 셀
    try {
        const cell_E77 = worksheet.getCell('E77');
        cell_E77.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_E77.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_E77.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 E77 설정 실패:', e);
    }

    // E78 셀
    try {
        const cell_E78 = worksheet.getCell('E78');
        cell_E78.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_E78.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_E78.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 E78 설정 실패:', e);
    }

    // E79 셀
    try {
        const cell_E79 = worksheet.getCell('E79');
        cell_E79.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_E79.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_E79.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 E79 설정 실패:', e);
    }

    // E8 셀
    try {
        const cell_E8 = worksheet.getCell('E8');
        cell_E8.value = '유스페이스1-A동';
        cell_E8.font = { name: 'LG스마트체 Regular', size: 10.0, color: { argb: 'FF000000' } };
        cell_E8.alignment = { horizontal: 'center', vertical: 'center' };
        cell_E8.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_E8.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 E8 설정 실패:', e);
    }

    // E80 셀
    try {
        const cell_E80 = worksheet.getCell('E80');
        cell_E80.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_E80.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_E80.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 E80 설정 실패:', e);
    }

    // E81 셀
    try {
        const cell_E81 = worksheet.getCell('E81');
        cell_E81.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_E81.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_E81.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 E81 설정 실패:', e);
    }

    // E82 셀
    try {
        const cell_E82 = worksheet.getCell('E82');
        cell_E82.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_E82.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_E82.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 E82 설정 실패:', e);
    }

    // E83 셀
    try {
        const cell_E83 = worksheet.getCell('E83');
        cell_E83.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_E83.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_E83.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 E83 설정 실패:', e);
    }

    // E9 셀
    try {
        const cell_E9 = worksheet.getCell('E9');
        cell_E9.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_E9.alignment = { horizontal: 'center', vertical: 'center' };
        cell_E9.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_E9.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 E9 설정 실패:', e);
    }

    // E90 셀
    try {
        const cell_E90 = worksheet.getCell('E90');
        cell_E90.font = { name: 'LG스마트체 Regular', size: 6.0, bold: true, color: { argb: 'FF000000' } };
        cell_E90.alignment = { horizontal: 'center', vertical: 'center' };
        cell_E90.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 E90 설정 실패:', e);
    }

    // E91 셀
    try {
        const cell_E91 = worksheet.getCell('E91');
        cell_E91.font = { name: 'LG스마트체 Regular', size: 6.0, bold: true, color: { argb: 'FF000000' } };
        cell_E91.alignment = { horizontal: 'center', vertical: 'center' };
        cell_E91.numFmt = '_-* #,##0_-;\-* #,##0_-;_-* "-"_-;_-@_-';
    } catch (e) {
        console.warn('셀 E91 설정 실패:', e);
    }

    // F106 셀
    try {
        const cell_F106 = worksheet.getCell('F106');
        cell_F106.font = { name: 'LG스마트체 Regular', size: 10.0, bold: true, color: { argb: 'FF000000' } };
        cell_F106.alignment = { horizontal: 'center', vertical: 'center' };
        cell_F106.numFmt = '_-* #,##0_-;\-* #,##0_-;_-* "-"_-;_-@_-';
    } catch (e) {
        console.warn('셀 F106 설정 실패:', e);
    }

    // F17 셀
    try {
        const cell_F17 = worksheet.getCell('F17');
        cell_F17.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_F17.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_F17.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 F17 설정 실패:', e);
    }

    // F18 셀
    try {
        const cell_F18 = worksheet.getCell('F18');
        cell_F18.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_F18.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_F18.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 F18 설정 실패:', e);
    }

    // F19 셀
    try {
        const cell_F19 = worksheet.getCell('F19');
        cell_F19.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_F19.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_F19.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 F19 설정 실패:', e);
    }

    // F20 셀
    try {
        const cell_F20 = worksheet.getCell('F20');
        cell_F20.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_F20.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_F20.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 F20 설정 실패:', e);
    }

    // F21 셀
    try {
        const cell_F21 = worksheet.getCell('F21');
        cell_F21.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_F21.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_F21.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 F21 설정 실패:', e);
    }

    // F22 셀
    try {
        const cell_F22 = worksheet.getCell('F22');
        cell_F22.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_F22.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_F22.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 F22 설정 실패:', e);
    }

    // F23 셀
    try {
        const cell_F23 = worksheet.getCell('F23');
        cell_F23.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_F23.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_F23.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 F23 설정 실패:', e);
    }

    // F24 셀
    try {
        const cell_F24 = worksheet.getCell('F24');
        cell_F24.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_F24.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_F24.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 F24 설정 실패:', e);
    }

    // F25 셀
    try {
        const cell_F25 = worksheet.getCell('F25');
        cell_F25.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_F25.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_F25.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 F25 설정 실패:', e);
    }

    // F26 셀
    try {
        const cell_F26 = worksheet.getCell('F26');
        cell_F26.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_F26.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_F26.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 F26 설정 실패:', e);
    }

    // F27 셀
    try {
        const cell_F27 = worksheet.getCell('F27');
        cell_F27.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_F27.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_F27.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 F27 설정 실패:', e);
    }

    // F28 셀
    try {
        const cell_F28 = worksheet.getCell('F28');
        cell_F28.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_F28.alignment = { horizontal: 'center', vertical: 'center' };
        cell_F28.border = { top: { style: 'hair', color: { argb: 'FF000000' } } };
        cell_F28.numFmt = '#,##0.000\ "평"';
    } catch (e) {
        console.warn('셀 F28 설정 실패:', e);
    }

    // F29 셀
    try {
        const cell_F29 = worksheet.getCell('F29');
        cell_F29.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_F29.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_F29.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 F29 설정 실패:', e);
    }

    // F30 셀
    try {
        const cell_F30 = worksheet.getCell('F30');
        cell_F30.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_F30.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_F30.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 F30 설정 실패:', e);
    }

    // F31 셀
    try {
        const cell_F31 = worksheet.getCell('F31');
        cell_F31.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_F31.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_F31.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 F31 설정 실패:', e);
    }

    // F32 셀
    try {
        const cell_F32 = worksheet.getCell('F32');
        cell_F32.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_F32.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_F32.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 F32 설정 실패:', e);
    }

    // F33 셀
    try {
        const cell_F33 = worksheet.getCell('F33');
        cell_F33.value = '전용';
        cell_F33.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_F33.alignment = { horizontal: 'center', vertical: 'center' };
        cell_F33.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_F33.numFmt = '@';
    } catch (e) {
        console.warn('셀 F33 설정 실패:', e);
    }

    // F34 셀
    try {
        const cell_F34 = worksheet.getCell('F34');
        cell_F34.value = '216.82';
        cell_F34.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF0000FF' } };
        cell_F34.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEECEC' } };
        cell_F34.alignment = { horizontal: 'center', vertical: 'center' };
        cell_F34.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_F34.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 F34 설정 실패:', e);
    }

    // F35 셀
    try {
        const cell_F35 = worksheet.getCell('F35');
        cell_F35.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF0000FF' } };
        cell_F35.alignment = { horizontal: 'center', vertical: 'center' };
        cell_F35.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_F35.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 F35 설정 실패:', e);
    }

    // F36 셀
    try {
        const cell_F36 = worksheet.getCell('F36');
        cell_F36.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_F36.alignment = { horizontal: 'center', vertical: 'center' };
        cell_F36.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_F36.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 F36 설정 실패:', e);
    }

    // F37 셀
    try {
        const cell_F37 = worksheet.getCell('F37');
        cell_F37.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_F37.alignment = { horizontal: 'center', vertical: 'center' };
        cell_F37.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_F37.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 F37 설정 실패:', e);
    }

    // F38 셀
    try {
        const cell_F38 = worksheet.getCell('F38');
        cell_F38.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_F38.alignment = { horizontal: 'center', vertical: 'center' };
        cell_F38.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_F38.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 F38 설정 실패:', e);
    }

    // F39 셀
    try {
        const cell_F39 = worksheet.getCell('F39');
        cell_F39.value = { formula: '=SUM(F34:F38)' };
        cell_F39.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_F39.alignment = { horizontal: 'center', vertical: 'center' };
        cell_F39.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_F39.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 F39 설정 실패:', e);
    }

    // F40 셀
    try {
        const cell_F40 = worksheet.getCell('F40');
        cell_F40.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_F40.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_F40.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 F40 설정 실패:', e);
    }

    // F41 셀
    try {
        const cell_F41 = worksheet.getCell('F41');
        cell_F41.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_F41.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_F41.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 F41 설정 실패:', e);
    }

    // F42 셀
    try {
        const cell_F42 = worksheet.getCell('F42');
        cell_F42.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_F42.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_F42.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 F42 설정 실패:', e);
    }

    // F43 셀
    try {
        const cell_F43 = worksheet.getCell('F43');
        cell_F43.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_F43.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_F43.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 F43 설정 실패:', e);
    }

    // F44 셀
    try {
        const cell_F44 = worksheet.getCell('F44');
        cell_F44.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_F44.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_F44.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 F44 설정 실패:', e);
    }

    // F45 셀
    try {
        const cell_F45 = worksheet.getCell('F45');
        cell_F45.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_F45.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_F45.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 F45 설정 실패:', e);
    }

    // F46 셀
    try {
        const cell_F46 = worksheet.getCell('F46');
        cell_F46.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_F46.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_F46.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 F46 설정 실패:', e);
    }

    // F47 셀
    try {
        const cell_F47 = worksheet.getCell('F47');
        cell_F47.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_F47.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_F47.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 F47 설정 실패:', e);
    }

    // F48 셀
    try {
        const cell_F48 = worksheet.getCell('F48');
        cell_F48.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_F48.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_F48.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 F48 설정 실패:', e);
    }

    // F49 셀
    try {
        const cell_F49 = worksheet.getCell('F49');
        cell_F49.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_F49.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_F49.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 F49 설정 실패:', e);
    }

    // F50 셀
    try {
        const cell_F50 = worksheet.getCell('F50');
        cell_F50.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_F50.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_F50.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 F50 설정 실패:', e);
    }

    // F51 셀
    try {
        const cell_F51 = worksheet.getCell('F51');
        cell_F51.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_F51.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_F51.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 F51 설정 실패:', e);
    }

    // F52 셀
    try {
        const cell_F52 = worksheet.getCell('F52');
        cell_F52.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_F52.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_F52.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 F52 설정 실패:', e);
    }

    // F53 셀
    try {
        const cell_F53 = worksheet.getCell('F53');
        cell_F53.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_F53.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_F53.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 F53 설정 실패:', e);
    }

    // F54 셀
    try {
        const cell_F54 = worksheet.getCell('F54');
        cell_F54.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_F54.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_F54.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 F54 설정 실패:', e);
    }

    // F55 셀
    try {
        const cell_F55 = worksheet.getCell('F55');
        cell_F55.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_F55.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_F55.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 F55 설정 실패:', e);
    }

    // F56 셀
    try {
        const cell_F56 = worksheet.getCell('F56');
        cell_F56.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_F56.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_F56.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 F56 설정 실패:', e);
    }

    // F57 셀
    try {
        const cell_F57 = worksheet.getCell('F57');
        cell_F57.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_F57.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_F57.border = { top: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 F57 설정 실패:', e);
    }

    // F58 셀
    try {
        const cell_F58 = worksheet.getCell('F58');
        cell_F58.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_F58.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_F58.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 F58 설정 실패:', e);
    }

    // F59 셀
    try {
        const cell_F59 = worksheet.getCell('F59');
        cell_F59.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_F59.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_F59.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 F59 설정 실패:', e);
    }

    // F6 셀
    try {
        const cell_F6 = worksheet.getCell('F6');
        cell_F6.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_F6.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_F6.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 F6 설정 실패:', e);
    }

    // F60 셀
    try {
        const cell_F60 = worksheet.getCell('F60');
        cell_F60.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_F60.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_F60.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 F60 설정 실패:', e);
    }

    // F61 셀
    try {
        const cell_F61 = worksheet.getCell('F61');
        cell_F61.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_F61.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_F61.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 F61 설정 실패:', e);
    }

    // F62 셀
    try {
        const cell_F62 = worksheet.getCell('F62');
        cell_F62.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_F62.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_F62.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 F62 설정 실패:', e);
    }

    // F63 셀
    try {
        const cell_F63 = worksheet.getCell('F63');
        cell_F63.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_F63.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_F63.border = { top: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 F63 설정 실패:', e);
    }

    // F7 셀
    try {
        const cell_F7 = worksheet.getCell('F7');
        cell_F7.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_F7.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_F7.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 F7 설정 실패:', e);
    }

    // F72 셀
    try {
        const cell_F72 = worksheet.getCell('F72');
        cell_F72.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_F72.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_F72.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 F72 설정 실패:', e);
    }

    // F73 셀
    try {
        const cell_F73 = worksheet.getCell('F73');
        cell_F73.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_F73.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_F73.border = { top: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 F73 설정 실패:', e);
    }

    // F8 셀
    try {
        const cell_F8 = worksheet.getCell('F8');
        cell_F8.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_F8.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_F8.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 F8 설정 실패:', e);
    }

    // F83 셀
    try {
        const cell_F83 = worksheet.getCell('F83');
        cell_F83.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_F83.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_F83.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 F83 설정 실패:', e);
    }

    // F9 셀
    try {
        const cell_F9 = worksheet.getCell('F9');
        cell_F9.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_F9.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_F9.border = { top: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 F9 설정 실패:', e);
    }

    // G10 셀
    try {
        const cell_G10 = worksheet.getCell('G10');
        cell_G10.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G10.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G10.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G10 설정 실패:', e);
    }

    // G106 셀
    try {
        const cell_G106 = worksheet.getCell('G106');
        cell_G106.font = { name: 'LG스마트체 Regular', size: 10.0, bold: true, color: { argb: 'FF000000' } };
        cell_G106.alignment = { horizontal: 'center', vertical: 'center' };
        cell_G106.numFmt = '_-* #,##0_-;\-* #,##0_-;_-* "-"_-;_-@_-';
    } catch (e) {
        console.warn('셀 G106 설정 실패:', e);
    }

    // G11 셀
    try {
        const cell_G11 = worksheet.getCell('G11');
        cell_G11.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G11.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G11.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G11 설정 실패:', e);
    }

    // G12 셀
    try {
        const cell_G12 = worksheet.getCell('G12');
        cell_G12.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G12.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G12.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G12 설정 실패:', e);
    }

    // G13 셀
    try {
        const cell_G13 = worksheet.getCell('G13');
        cell_G13.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G13.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G13.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G13 설정 실패:', e);
    }

    // G14 셀
    try {
        const cell_G14 = worksheet.getCell('G14');
        cell_G14.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G14.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G14.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G14 설정 실패:', e);
    }

    // G15 셀
    try {
        const cell_G15 = worksheet.getCell('G15');
        cell_G15.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G15.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G15.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G15 설정 실패:', e);
    }

    // G16 셀
    try {
        const cell_G16 = worksheet.getCell('G16');
        cell_G16.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G16.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G16.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G16 설정 실패:', e);
    }

    // G17 셀
    try {
        const cell_G17 = worksheet.getCell('G17');
        cell_G17.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G17.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G17.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G17 설정 실패:', e);
    }

    // G18 셀
    try {
        const cell_G18 = worksheet.getCell('G18');
        cell_G18.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G18.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G18.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G18 설정 실패:', e);
    }

    // G19 셀
    try {
        const cell_G19 = worksheet.getCell('G19');
        cell_G19.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G19.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G19.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G19 설정 실패:', e);
    }

    // G20 셀
    try {
        const cell_G20 = worksheet.getCell('G20');
        cell_G20.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G20.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G20.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G20 설정 실패:', e);
    }

    // G21 셀
    try {
        const cell_G21 = worksheet.getCell('G21');
        cell_G21.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G21.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G21.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G21 설정 실패:', e);
    }

    // G22 셀
    try {
        const cell_G22 = worksheet.getCell('G22');
        cell_G22.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G22.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G22.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G22 설정 실패:', e);
    }

    // G23 셀
    try {
        const cell_G23 = worksheet.getCell('G23');
        cell_G23.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G23.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G23.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G23 설정 실패:', e);
    }

    // G24 셀
    try {
        const cell_G24 = worksheet.getCell('G24');
        cell_G24.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G24.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G24.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G24 설정 실패:', e);
    }

    // G25 셀
    try {
        const cell_G25 = worksheet.getCell('G25');
        cell_G25.value = '17408.4';
        cell_G25.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_G25.alignment = { horizontal: 'center', vertical: 'center' };
        cell_G25.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_G25.numFmt = '"("#,##0.0\ "㎡)"';
    } catch (e) {
        console.warn('셀 G25 설정 실패:', e);
    }

    // G26 셀
    try {
        const cell_G26 = worksheet.getCell('G26');
        cell_G26.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G26.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G26.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G26 설정 실패:', e);
    }

    // G27 셀
    try {
        const cell_G27 = worksheet.getCell('G27');
        cell_G27.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G27.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G27.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G27 설정 실패:', e);
    }

    // G28 셀
    try {
        const cell_G28 = worksheet.getCell('G28');
        cell_G28.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_G28.alignment = { horizontal: 'center', vertical: 'center' };
        cell_G28.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_G28.numFmt = '#,##0.000\ "평"';
    } catch (e) {
        console.warn('셀 G28 설정 실패:', e);
    }

    // G29 셀
    try {
        const cell_G29 = worksheet.getCell('G29');
        cell_G29.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G29.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G29.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G29 설정 실패:', e);
    }

    // G30 셀
    try {
        const cell_G30 = worksheet.getCell('G30');
        cell_G30.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G30.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G30.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G30 설정 실패:', e);
    }

    // G31 셀
    try {
        const cell_G31 = worksheet.getCell('G31');
        cell_G31.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G31.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G31.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G31 설정 실패:', e);
    }

    // G32 셀
    try {
        const cell_G32 = worksheet.getCell('G32');
        cell_G32.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G32.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G32.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G32 설정 실패:', e);
    }

    // G33 셀
    try {
        const cell_G33 = worksheet.getCell('G33');
        cell_G33.value = '임대';
        cell_G33.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_G33.alignment = { horizontal: 'center', vertical: 'center' };
        cell_G33.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_G33.numFmt = '@';
    } catch (e) {
        console.warn('셀 G33 설정 실패:', e);
    }

    // G34 셀
    try {
        const cell_G34 = worksheet.getCell('G34');
        cell_G34.value = '467.42';
        cell_G34.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_G34.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEECEC' } };
        cell_G34.alignment = { horizontal: 'center', vertical: 'center' };
        cell_G34.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_G34.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 G34 설정 실패:', e);
    }

    // G35 셀
    try {
        const cell_G35 = worksheet.getCell('G35');
        cell_G35.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_G35.alignment = { horizontal: 'center', vertical: 'center' };
        cell_G35.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_G35.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 G35 설정 실패:', e);
    }

    // G36 셀
    try {
        const cell_G36 = worksheet.getCell('G36');
        cell_G36.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_G36.alignment = { horizontal: 'center', vertical: 'center' };
        cell_G36.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_G36.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 G36 설정 실패:', e);
    }

    // G37 셀
    try {
        const cell_G37 = worksheet.getCell('G37');
        cell_G37.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_G37.alignment = { horizontal: 'center', vertical: 'center' };
        cell_G37.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_G37.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 G37 설정 실패:', e);
    }

    // G38 셀
    try {
        const cell_G38 = worksheet.getCell('G38');
        cell_G38.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_G38.alignment = { horizontal: 'center', vertical: 'center' };
        cell_G38.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_G38.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 G38 설정 실패:', e);
    }

    // G39 셀
    try {
        const cell_G39 = worksheet.getCell('G39');
        cell_G39.value = { formula: '=SUM(G34:G38)' };
        cell_G39.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_G39.alignment = { horizontal: 'center', vertical: 'center' };
        cell_G39.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_G39.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 G39 설정 실패:', e);
    }

    // G40 셀
    try {
        const cell_G40 = worksheet.getCell('G40');
        cell_G40.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G40.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G40.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G40 설정 실패:', e);
    }

    // G41 셀
    try {
        const cell_G41 = worksheet.getCell('G41');
        cell_G41.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G41.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G41.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G41 설정 실패:', e);
    }

    // G42 셀
    try {
        const cell_G42 = worksheet.getCell('G42');
        cell_G42.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G42.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G42.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G42 설정 실패:', e);
    }

    // G43 셀
    try {
        const cell_G43 = worksheet.getCell('G43');
        cell_G43.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G43.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G43.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G43 설정 실패:', e);
    }

    // G44 셀
    try {
        const cell_G44 = worksheet.getCell('G44');
        cell_G44.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G44.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G44.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G44 설정 실패:', e);
    }

    // G45 셀
    try {
        const cell_G45 = worksheet.getCell('G45');
        cell_G45.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G45.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G45.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G45 설정 실패:', e);
    }

    // G46 셀
    try {
        const cell_G46 = worksheet.getCell('G46');
        cell_G46.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G46.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G46.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G46 설정 실패:', e);
    }

    // G47 셀
    try {
        const cell_G47 = worksheet.getCell('G47');
        cell_G47.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G47.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G47.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G47 설정 실패:', e);
    }

    // G48 셀
    try {
        const cell_G48 = worksheet.getCell('G48');
        cell_G48.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G48.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G48.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G48 설정 실패:', e);
    }

    // G49 셀
    try {
        const cell_G49 = worksheet.getCell('G49');
        cell_G49.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G49.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G49.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G49 설정 실패:', e);
    }

    // G50 셀
    try {
        const cell_G50 = worksheet.getCell('G50');
        cell_G50.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G50.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G50.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G50 설정 실패:', e);
    }

    // G51 셀
    try {
        const cell_G51 = worksheet.getCell('G51');
        cell_G51.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G51.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G51.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G51 설정 실패:', e);
    }

    // G52 셀
    try {
        const cell_G52 = worksheet.getCell('G52');
        cell_G52.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G52.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G52.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G52 설정 실패:', e);
    }

    // G53 셀
    try {
        const cell_G53 = worksheet.getCell('G53');
        cell_G53.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G53.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G53.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G53 설정 실패:', e);
    }

    // G54 셀
    try {
        const cell_G54 = worksheet.getCell('G54');
        cell_G54.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G54.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G54.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G54 설정 실패:', e);
    }

    // G55 셀
    try {
        const cell_G55 = worksheet.getCell('G55');
        cell_G55.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G55.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G55.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G55 설정 실패:', e);
    }

    // G56 셀
    try {
        const cell_G56 = worksheet.getCell('G56');
        cell_G56.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G56.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G56.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G56 설정 실패:', e);
    }

    // G57 셀
    try {
        const cell_G57 = worksheet.getCell('G57');
        cell_G57.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G57.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G57.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G57 설정 실패:', e);
    }

    // G58 셀
    try {
        const cell_G58 = worksheet.getCell('G58');
        cell_G58.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G58.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G58.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G58 설정 실패:', e);
    }

    // G59 셀
    try {
        const cell_G59 = worksheet.getCell('G59');
        cell_G59.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G59.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G59.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G59 설정 실패:', e);
    }

    // G6 셀
    try {
        const cell_G6 = worksheet.getCell('G6');
        cell_G6.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G6.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G6.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G6 설정 실패:', e);
    }

    // G60 셀
    try {
        const cell_G60 = worksheet.getCell('G60');
        cell_G60.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G60.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G60.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G60 설정 실패:', e);
    }

    // G61 셀
    try {
        const cell_G61 = worksheet.getCell('G61');
        cell_G61.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G61.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G61.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G61 설정 실패:', e);
    }

    // G62 셀
    try {
        const cell_G62 = worksheet.getCell('G62');
        cell_G62.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G62.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G62.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G62 설정 실패:', e);
    }

    // G63 셀
    try {
        const cell_G63 = worksheet.getCell('G63');
        cell_G63.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G63.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G63.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G63 설정 실패:', e);
    }

    // G64 셀
    try {
        const cell_G64 = worksheet.getCell('G64');
        cell_G64.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G64.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G64.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G64 설정 실패:', e);
    }

    // G65 셀
    try {
        const cell_G65 = worksheet.getCell('G65');
        cell_G65.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G65.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G65.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G65 설정 실패:', e);
    }

    // G66 셀
    try {
        const cell_G66 = worksheet.getCell('G66');
        cell_G66.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G66.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G66.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G66 설정 실패:', e);
    }

    // G67 셀
    try {
        const cell_G67 = worksheet.getCell('G67');
        cell_G67.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G67.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G67.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G67 설정 실패:', e);
    }

    // G68 셀
    try {
        const cell_G68 = worksheet.getCell('G68');
        cell_G68.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G68.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G68.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G68 설정 실패:', e);
    }

    // G69 셀
    try {
        const cell_G69 = worksheet.getCell('G69');
        cell_G69.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G69.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G69.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G69 설정 실패:', e);
    }

    // G7 셀
    try {
        const cell_G7 = worksheet.getCell('G7');
        cell_G7.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G7.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G7.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G7 설정 실패:', e);
    }

    // G70 셀
    try {
        const cell_G70 = worksheet.getCell('G70');
        cell_G70.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G70.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G70.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G70 설정 실패:', e);
    }

    // G71 셀
    try {
        const cell_G71 = worksheet.getCell('G71');
        cell_G71.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G71.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G71.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G71 설정 실패:', e);
    }

    // G72 셀
    try {
        const cell_G72 = worksheet.getCell('G72');
        cell_G72.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G72.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G72.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G72 설정 실패:', e);
    }

    // G73 셀
    try {
        const cell_G73 = worksheet.getCell('G73');
        cell_G73.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G73.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G73.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G73 설정 실패:', e);
    }

    // G74 셀
    try {
        const cell_G74 = worksheet.getCell('G74');
        cell_G74.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G74.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G74.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G74 설정 실패:', e);
    }

    // G75 셀
    try {
        const cell_G75 = worksheet.getCell('G75');
        cell_G75.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G75.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G75.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G75 설정 실패:', e);
    }

    // G76 셀
    try {
        const cell_G76 = worksheet.getCell('G76');
        cell_G76.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G76.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G76.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G76 설정 실패:', e);
    }

    // G77 셀
    try {
        const cell_G77 = worksheet.getCell('G77');
        cell_G77.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G77.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G77.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G77 설정 실패:', e);
    }

    // G78 셀
    try {
        const cell_G78 = worksheet.getCell('G78');
        cell_G78.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G78.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G78.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G78 설정 실패:', e);
    }

    // G79 셀
    try {
        const cell_G79 = worksheet.getCell('G79');
        cell_G79.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G79.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G79.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G79 설정 실패:', e);
    }

    // G8 셀
    try {
        const cell_G8 = worksheet.getCell('G8');
        cell_G8.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G8.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G8.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G8 설정 실패:', e);
    }

    // G80 셀
    try {
        const cell_G80 = worksheet.getCell('G80');
        cell_G80.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G80.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G80.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G80 설정 실패:', e);
    }

    // G81 셀
    try {
        const cell_G81 = worksheet.getCell('G81');
        cell_G81.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G81.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G81.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G81 설정 실패:', e);
    }

    // G82 셀
    try {
        const cell_G82 = worksheet.getCell('G82');
        cell_G82.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G82.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G82.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G82 설정 실패:', e);
    }

    // G83 셀
    try {
        const cell_G83 = worksheet.getCell('G83');
        cell_G83.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G83.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G83.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G83 설정 실패:', e);
    }

    // G9 셀
    try {
        const cell_G9 = worksheet.getCell('G9');
        cell_G9.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_G9.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_G9.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 G9 설정 실패:', e);
    }

    // H106 셀
    try {
        const cell_H106 = worksheet.getCell('H106');
        cell_H106.font = { name: 'LG스마트체 Regular', size: 10.0, bold: true, color: { argb: 'FF000000' } };
        cell_H106.alignment = { horizontal: 'center', vertical: 'center' };
        cell_H106.numFmt = '_-* #,##0_-;\-* #,##0_-;_-* "-"_-;_-@_-';
    } catch (e) {
        console.warn('셀 H106 설정 실패:', e);
    }

    // H17 셀
    try {
        const cell_H17 = worksheet.getCell('H17');
        cell_H17.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_H17.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_H17.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 H17 설정 실패:', e);
    }

    // H18 셀
    try {
        const cell_H18 = worksheet.getCell('H18');
        cell_H18.value = '서울시 금천구 가산디지털1로 186';
        cell_H18.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_H18.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        cell_H18.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_H18.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 H18 설정 실패:', e);
    }

    // H19 셀
    try {
        const cell_H19 = worksheet.getCell('H19');
        cell_H19.value = '1,7호선 가산디지털단지역 도보 1분';
        cell_H19.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_H19.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        cell_H19.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_H19.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 H19 설정 실패:', e);
    }

    // H20 셀
    try {
        const cell_H20 = worksheet.getCell('H20');
        cell_H20.value = '2007';
        cell_H20.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_H20.alignment = { horizontal: 'center', vertical: 'center' };
        cell_H20.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_H20.numFmt = '###0"년"';
    } catch (e) {
        console.warn('셀 H20 설정 실패:', e);
    }

    // H21 셀
    try {
        const cell_H21 = worksheet.getCell('H21');
        cell_H21.value = '154';
        cell_H21.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_H21.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        cell_H21.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_H21.numFmt = '##"F / B"#';
    } catch (e) {
        console.warn('셀 H21 설정 실패:', e);
    }

    // H22 셀
    try {
        const cell_H22 = worksheet.getCell('H22');
        cell_H22.value = '29745';
        cell_H22.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_H22.alignment = { horizontal: 'center', vertical: 'center' };
        cell_H22.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_H22.numFmt = '#,##0\ "평"';
    } catch (e) {
        console.warn('셀 H22 설정 실패:', e);
    }

    // H23 셀
    try {
        const cell_H23 = worksheet.getCell('H23');
        cell_H23.value = '878.52';
        cell_H23.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_H23.alignment = { horizontal: 'center', vertical: 'center' };
        cell_H23.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_H23.numFmt = '#,##0\ "평"';
    } catch (e) {
        console.warn('셀 H23 설정 실패:', e);
    }

    // H24 셀
    try {
        const cell_H24 = worksheet.getCell('H24');
        cell_H24.value = '0.581';
        cell_H24.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_H24.alignment = { horizontal: 'center', vertical: 'center' };
        cell_H24.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_H24.numFmt = '0.00%';
    } catch (e) {
        console.warn('셀 H24 설정 실패:', e);
    }

    // H25 셀
    try {
        const cell_H25 = worksheet.getCell('H25');
        cell_H25.value = { formula: '=J25*0.3025' };
        cell_H25.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_H25.alignment = { horizontal: 'center', vertical: 'center' };
        cell_H25.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_H25.numFmt = '#,##0\ "평"';
    } catch (e) {
        console.warn('셀 H25 설정 실패:', e);
    }

    // H26 셀
    try {
        const cell_H26 = worksheet.getCell('H26');
        cell_H26.value = '재능홀딩스';
        cell_H26.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_H26.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        cell_H26.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_H26.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 H26 설정 실패:', e);
    }

    // H27 셀
    try {
        const cell_H27 = worksheet.getCell('H27');
        cell_H27.value = '전세권, 근저당권 설정 가능';
        cell_H27.font = { name: 'LG스마트체 Regular', size: 9.0, bold: true, color: { argb: 'FFC00000' } };
        cell_H27.alignment = { horizontal: 'center', vertical: 'center' };
        cell_H27.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_H27.numFmt = '0.00%';
    } catch (e) {
        console.warn('셀 H27 설정 실패:', e);
    }

    // H28 셀
    try {
        const cell_H28 = worksheet.getCell('H28');
        cell_H28.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_H28.alignment = { horizontal: 'center', vertical: 'center' };
        cell_H28.border = { top: { style: 'hair', color: { argb: 'FF000000' } } };
        cell_H28.numFmt = '#,##0.000\ "평"';
    } catch (e) {
        console.warn('셀 H28 설정 실패:', e);
    }

    // H29 셀
    try {
        const cell_H29 = worksheet.getCell('H29');
        cell_H29.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_H29.alignment = { horizontal: 'center', vertical: 'center' };
        cell_H29.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_H29.numFmt = '_-* #,##0_-;\-* #,##0_-;_-* "-"_-;_-@_-';
    } catch (e) {
        console.warn('셀 H29 설정 실패:', e);
    }

    // H30 셀
    try {
        const cell_H30 = worksheet.getCell('H30');
        cell_H30.value = { formula: '=H29/H32' };
        cell_H30.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FFC00000' } };
        cell_H30.alignment = { horizontal: 'center', vertical: 'center' };
        cell_H30.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_H30.numFmt = '0.00%';
    } catch (e) {
        console.warn('셀 H30 설정 실패:', e);
    }

    // H31 셀
    try {
        const cell_H31 = worksheet.getCell('H31');
        cell_H31.value = '4768000';
        cell_H31.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_H31.alignment = { horizontal: 'center', vertical: 'center' };
        cell_H31.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_H31.numFmt = '#,###"원/㎡"';
    } catch (e) {
        console.warn('셀 H31 설정 실패:', e);
    }

    // H32 셀
    try {
        const cell_H32 = worksheet.getCell('H32');
        cell_H32.value = { formula: '=H31*J25' };
        cell_H32.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_H32.alignment = { horizontal: 'center', vertical: 'center' };
        cell_H32.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_H32.numFmt = '_-* #,##0_-;\-* #,##0_-;_-* "-"_-;_-@_-';
    } catch (e) {
        console.warn('셀 H32 설정 실패:', e);
    }

    // H33 셀
    try {
        const cell_H33 = worksheet.getCell('H33');
        cell_H33.value = '층';
        cell_H33.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_H33.alignment = { horizontal: 'center', vertical: 'center' };
        cell_H33.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_H33.numFmt = '@';
    } catch (e) {
        console.warn('셀 H33 설정 실패:', e);
    }

    // H34 셀
    try {
        const cell_H34 = worksheet.getCell('H34');
        cell_H34.value = '11';
        cell_H34.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF0000FF' } };
        cell_H34.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEECEC' } };
        cell_H34.alignment = { horizontal: 'center', vertical: 'center' };
        cell_H34.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_H34.numFmt = '#"층"';
    } catch (e) {
        console.warn('셀 H34 설정 실패:', e);
    }

    // H35 셀
    try {
        const cell_H35 = worksheet.getCell('H35');
        cell_H35.value = '2';
        cell_H35.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_H35.alignment = { horizontal: 'center', vertical: 'center' };
        cell_H35.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_H35.numFmt = '#"층"';
    } catch (e) {
        console.warn('셀 H35 설정 실패:', e);
    }

    // H36 셀
    try {
        const cell_H36 = worksheet.getCell('H36');
        cell_H36.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_H36.alignment = { horizontal: 'center', vertical: 'center' };
        cell_H36.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_H36.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 H36 설정 실패:', e);
    }

    // H37 셀
    try {
        const cell_H37 = worksheet.getCell('H37');
        cell_H37.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_H37.alignment = { horizontal: 'center', vertical: 'center' };
        cell_H37.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_H37.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 H37 설정 실패:', e);
    }

    // H38 셀
    try {
        const cell_H38 = worksheet.getCell('H38');
        cell_H38.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_H38.alignment = { horizontal: 'center', vertical: 'center' };
        cell_H38.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_H38.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 H38 설정 실패:', e);
    }

    // H39 셀
    try {
        const cell_H39 = worksheet.getCell('H39');
        cell_H39.value = '소계';
        cell_H39.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_H39.alignment = { horizontal: 'center', vertical: 'center' };
        cell_H39.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_H39.numFmt = '@';
    } catch (e) {
        console.warn('셀 H39 설정 실패:', e);
    }

    // H40 셀
    try {
        const cell_H40 = worksheet.getCell('H40');
        cell_H40.value = '2025.7~2027.6 (12개월)';
        cell_H40.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_H40.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        cell_H40.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_H40.numFmt = '#"개월 계약 가능"';
    } catch (e) {
        console.warn('셀 H40 설정 실패:', e);
    }

    // H41 셀
    try {
        const cell_H41 = worksheet.getCell('H41');
        cell_H41.value = '즉시';
        cell_H41.font = { name: 'LG스마트체 Regular', size: 9.0, bold: true, color: { argb: 'FF000000' } };
        cell_H41.alignment = { horizontal: 'center', vertical: 'center' };
        cell_H41.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_H41.numFmt = '#,##0\ "원"';
    } catch (e) {
        console.warn('셀 H41 설정 실패:', e);
    }

    // H42 셀
    try {
        const cell_H42 = worksheet.getCell('H42');
        cell_H42.value = '11층 일부';
        cell_H42.font = { name: 'LG스마트체 Regular', size: 9.0, bold: true, color: { argb: 'FF000000' } };
        cell_H42.alignment = { horizontal: 'center', vertical: 'center' };
        cell_H42.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_H42.numFmt = '#,##0\ "층"';
    } catch (e) {
        console.warn('셀 H42 설정 실패:', e);
    }

    // H43 셀
    try {
        const cell_H43 = worksheet.getCell('H43');
        cell_H43.value = { formula: '=I34' };
        cell_H43.font = { name: 'LG스마트체 Regular', size: 9.0, bold: true, color: { argb: 'FFC00000' } };
        cell_H43.alignment = { horizontal: 'center', vertical: 'center' };
        cell_H43.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_H43.numFmt = '#,##0\ "평"';
    } catch (e) {
        console.warn('셀 H43 설정 실패:', e);
    }

    // H44 셀
    try {
        const cell_H44 = worksheet.getCell('H44');
        cell_H44.value = { formula: '=J34' };
        cell_H44.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_H44.alignment = { horizontal: 'center', vertical: 'center' };
        cell_H44.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_H44.numFmt = '#,##0\ "평"';
    } catch (e) {
        console.warn('셀 H44 설정 실패:', e);
    }

    // H45 셀
    try {
        const cell_H45 = worksheet.getCell('H45');
        cell_H45.value = '430000';
        cell_H45.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_H45.alignment = { horizontal: 'center', vertical: 'center' };
        cell_H45.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_H45.numFmt = '"@"#,###';
    } catch (e) {
        console.warn('셀 H45 설정 실패:', e);
    }

    // H46 셀
    try {
        const cell_H46 = worksheet.getCell('H46');
        cell_H46.value = '43000';
        cell_H46.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_H46.alignment = { horizontal: 'center', vertical: 'center' };
        cell_H46.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_H46.numFmt = '"@"#,###';
    } catch (e) {
        console.warn('셀 H46 설정 실패:', e);
    }

    // H47 셀
    try {
        const cell_H47 = worksheet.getCell('H47');
        cell_H47.value = '8000';
        cell_H47.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_H47.alignment = { horizontal: 'center', vertical: 'center' };
        cell_H47.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_H47.numFmt = '"@"#,###\+"실""비""별""도"';
    } catch (e) {
        console.warn('셀 H47 설정 실패:', e);
    }

    // H48 셀
    try {
        const cell_H48 = worksheet.getCell('H48');
        cell_H48.value = { formula: '=H46*(12-H49)/12' };
        cell_H48.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_H48.alignment = { horizontal: 'center', vertical: 'center' };
        cell_H48.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_H48.numFmt = '"@"#,###';
    } catch (e) {
        console.warn('셀 H48 설정 실패:', e);
    }

    // H49 셀
    try {
        const cell_H49 = worksheet.getCell('H49');
        cell_H49.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_H49.alignment = { horizontal: 'center', vertical: 'center' };
        cell_H49.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_H49.numFmt = '#0.0"개월"';
    } catch (e) {
        console.warn('셀 H49 설정 실패:', e);
    }

    // H50 셀
    try {
        const cell_H50 = worksheet.getCell('H50');
        cell_H50.value = { formula: '=H45*H44' };
        cell_H50.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_H50.alignment = { horizontal: 'center', vertical: 'center' };
        cell_H50.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_H50.numFmt = '#,##0\ "원"';
    } catch (e) {
        console.warn('셀 H50 설정 실패:', e);
    }

    // H51 셀
    try {
        const cell_H51 = worksheet.getCell('H51');
        cell_H51.value = { formula: '=H46*H44' };
        cell_H51.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_H51.alignment = { horizontal: 'center', vertical: 'center' };
        cell_H51.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_H51.numFmt = '#,##0\ "원"';
    } catch (e) {
        console.warn('셀 H51 설정 실패:', e);
    }

    // H52 셀
    try {
        const cell_H52 = worksheet.getCell('H52');
        cell_H52.value = { formula: '=H47*H44' };
        cell_H52.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_H52.alignment = { horizontal: 'center', vertical: 'center' };
        cell_H52.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_H52.numFmt = '#,##0\ "원"';
    } catch (e) {
        console.warn('셀 H52 설정 실패:', e);
    }

    // H53 셀
    try {
        const cell_H53 = worksheet.getCell('H53');
        cell_H53.value = '실비 관리비 : 전기세, 수도세  별도 부과';
        cell_H53.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FFC00000' } };
        cell_H53.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        cell_H53.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_H53.numFmt = '#,##0\ "원"';
    } catch (e) {
        console.warn('셀 H53 설정 실패:', e);
    }

    // H54 셀
    try {
        const cell_H54 = worksheet.getCell('H54');
        cell_H54.value = { formula: '=H51+H52' };
        cell_H54.font = { name: 'LG스마트체 Regular', size: 9.0, bold: true, color: { argb: 'FFC00000' } };
        cell_H54.alignment = { horizontal: 'center', vertical: 'center' };
        cell_H54.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_H54.numFmt = '#,##0\ "원"';
    } catch (e) {
        console.warn('셀 H54 설정 실패:', e);
    }

    // H55 셀
    try {
        const cell_H55 = worksheet.getCell('H55');
        cell_H55.value = { formula: '=H54*21' };
        cell_H55.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_H55.alignment = { horizontal: 'center', vertical: 'center' };
        cell_H55.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_H55.numFmt = '#,##0\ "원"';
    } catch (e) {
        console.warn('셀 H55 설정 실패:', e);
    }

    // H56 셀
    try {
        const cell_H56 = worksheet.getCell('H56');
        cell_H56.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_H56.alignment = { horizontal: 'center', vertical: 'center' };
        cell_H56.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_H56.numFmt = '0.#"개월"';
    } catch (e) {
        console.warn('셀 H56 설정 실패:', e);
    }

    // H57 셀
    try {
        const cell_H57 = worksheet.getCell('H57');
        cell_H57.value = '미제공';
        cell_H57.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_H57.alignment = { horizontal: 'center', vertical: 'center' };
        cell_H57.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_H57.numFmt = '"총액 "##,##0"원"';
    } catch (e) {
        console.warn('셀 H57 설정 실패:', e);
    }

    // H58 셀
    try {
        const cell_H58 = worksheet.getCell('H58');
        cell_H58.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_H58.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_H58.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 H58 설정 실패:', e);
    }

    // H59 셀
    try {
        const cell_H59 = worksheet.getCell('H59');
        cell_H59.value = { formula: '=732+44' };
        cell_H59.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_H59.alignment = { horizontal: 'center', vertical: 'center' };
        cell_H59.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_H59.numFmt = '#\ "대"';
    } catch (e) {
        console.warn('셀 H59 설정 실패:', e);
    }

    // H6 셀
    try {
        const cell_H6 = worksheet.getCell('H6');
        cell_H6.value = '가산디지털단지역';
        cell_H6.font = { name: 'LG스마트체 Regular', size: 10.0, color: { argb: 'FF000000' } };
        cell_H6.alignment = { horizontal: 'center', vertical: 'center' };
        cell_H6.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_H6.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 H6 설정 실패:', e);
    }

    // H60 셀
    try {
        const cell_H60 = worksheet.getCell('H60');
        cell_H60.value = '35';
        cell_H60.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_H60.alignment = { horizontal: 'center', vertical: 'center' };
        cell_H60.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_H60.numFmt = '"임대면적"\ #"평당 1대"';
    } catch (e) {
        console.warn('셀 H60 설정 실패:', e);
    }

    // H61 셀
    try {
        const cell_H61 = worksheet.getCell('H61');
        cell_H61.value = { formula: '=H44/H60' };
        cell_H61.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_H61.alignment = { horizontal: 'center', vertical: 'center' };
        cell_H61.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_H61.numFmt = '#,##0.0\ "대"';
    } catch (e) {
        console.warn('셀 H61 설정 실패:', e);
    }

    // H62 셀
    try {
        const cell_H62 = worksheet.getCell('H62');
        cell_H62.value = '10';
        cell_H62.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_H62.alignment = { horizontal: 'center', vertical: 'center' };
        cell_H62.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_H62.numFmt = '"월"#"만원/대"';
    } catch (e) {
        console.warn('셀 H62 설정 실패:', e);
    }

    // H63 셀
    try {
        const cell_H63 = worksheet.getCell('H63');
        cell_H63.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_H63.alignment = { horizontal: 'left', vertical: 'center', wrapText: true };
        cell_H63.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_H63.numFmt = '#,##0_ ';
    } catch (e) {
        console.warn('셀 H63 설정 실패:', e);
    }

    // H7 셀
    try {
        const cell_H7 = worksheet.getCell('H7');
        cell_H7.font = { name: 'LG스마트체 Regular', size: 10.0, color: { argb: 'FF000000' } };
        cell_H7.alignment = { horizontal: 'center', vertical: 'center' };
        cell_H7.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_H7.numFmt = '0_);[Red]\(0\)';
    } catch (e) {
        console.warn('셀 H7 설정 실패:', e);
    }

    // H72 셀
    try {
        const cell_H72 = worksheet.getCell('H72');
        cell_H72.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_H72.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_H72.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 H72 설정 실패:', e);
    }

    // H73 셀
    try {
        const cell_H73 = worksheet.getCell('H73');
        cell_H73.value = ' - 현재 2층 & 11층 일부 공실\n - 2층(204호) 지원시설, 11층(1103호) 업무시설\n - 11층 월 관리비 평균 170만원, 2층 실비 관리비\n   41만원\n - 실내청소 무료 서비스\n - 주차 : 자주식\n   ';
        cell_H73.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_H73.alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
        cell_H73.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_H73.numFmt = '#,##0\ "대"';
    } catch (e) {
        console.warn('셀 H73 설정 실패:', e);
    }

    // H8 셀
    try {
        const cell_H8 = worksheet.getCell('H8');
        cell_H8.value = '제이플라츠(11층)';
        cell_H8.font = { name: 'LG스마트체 Regular', size: 10.0, color: { argb: 'FF000000' } };
        cell_H8.alignment = { horizontal: 'center', vertical: 'center' };
        cell_H8.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_H8.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 H8 설정 실패:', e);
    }

    // H83 셀
    try {
        const cell_H83 = worksheet.getCell('H83');
        cell_H83.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_H83.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_H83.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 H83 설정 실패:', e);
    }

    // H9 셀
    try {
        const cell_H9 = worksheet.getCell('H9');
        cell_H9.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_H9.alignment = { horizontal: 'center', vertical: 'center' };
        cell_H9.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_H9.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 H9 설정 실패:', e);
    }

    // H91 셀
    try {
        const cell_H91 = worksheet.getCell('H91');
        cell_H91.font = { name: 'LG스마트체 Regular', size: 6.0, bold: true, color: { argb: 'FF000000' } };
        cell_H91.alignment = { horizontal: 'center', vertical: 'center' };
        cell_H91.numFmt = '_-* #,##0_-;\-* #,##0_-;_-* "-"_-;_-@_-';
    } catch (e) {
        console.warn('셀 H91 설정 실패:', e);
    }

    // I106 셀
    try {
        const cell_I106 = worksheet.getCell('I106');
        cell_I106.font = { name: 'LG스마트체 Regular', size: 10.0, bold: true, color: { argb: 'FF000000' } };
        cell_I106.alignment = { horizontal: 'center', vertical: 'center' };
        cell_I106.numFmt = '_-* #,##0_-;\-* #,##0_-;_-* "-"_-;_-@_-';
    } catch (e) {
        console.warn('셀 I106 설정 실패:', e);
    }

    // I17 셀
    try {
        const cell_I17 = worksheet.getCell('I17');
        cell_I17.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_I17.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_I17.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 I17 설정 실패:', e);
    }

    // I18 셀
    try {
        const cell_I18 = worksheet.getCell('I18');
        cell_I18.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_I18.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_I18.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 I18 설정 실패:', e);
    }

    // I19 셀
    try {
        const cell_I19 = worksheet.getCell('I19');
        cell_I19.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_I19.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_I19.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 I19 설정 실패:', e);
    }

    // I20 셀
    try {
        const cell_I20 = worksheet.getCell('I20');
        cell_I20.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_I20.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_I20.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 I20 설정 실패:', e);
    }

    // I21 셀
    try {
        const cell_I21 = worksheet.getCell('I21');
        cell_I21.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_I21.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_I21.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 I21 설정 실패:', e);
    }

    // I22 셀
    try {
        const cell_I22 = worksheet.getCell('I22');
        cell_I22.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_I22.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_I22.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 I22 설정 실패:', e);
    }

    // I23 셀
    try {
        const cell_I23 = worksheet.getCell('I23');
        cell_I23.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_I23.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_I23.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 I23 설정 실패:', e);
    }

    // I24 셀
    try {
        const cell_I24 = worksheet.getCell('I24');
        cell_I24.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_I24.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_I24.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 I24 설정 실패:', e);
    }

    // I25 셀
    try {
        const cell_I25 = worksheet.getCell('I25');
        cell_I25.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_I25.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_I25.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 I25 설정 실패:', e);
    }

    // I26 셀
    try {
        const cell_I26 = worksheet.getCell('I26');
        cell_I26.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_I26.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_I26.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 I26 설정 실패:', e);
    }

    // I27 셀
    try {
        const cell_I27 = worksheet.getCell('I27');
        cell_I27.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_I27.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_I27.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 I27 설정 실패:', e);
    }

    // I28 셀
    try {
        const cell_I28 = worksheet.getCell('I28');
        cell_I28.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_I28.alignment = { horizontal: 'center', vertical: 'center' };
        cell_I28.border = { top: { style: 'hair', color: { argb: 'FF000000' } } };
        cell_I28.numFmt = '#,##0.000\ "평"';
    } catch (e) {
        console.warn('셀 I28 설정 실패:', e);
    }

    // I29 셀
    try {
        const cell_I29 = worksheet.getCell('I29');
        cell_I29.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_I29.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_I29.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 I29 설정 실패:', e);
    }

    // I30 셀
    try {
        const cell_I30 = worksheet.getCell('I30');
        cell_I30.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_I30.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_I30.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 I30 설정 실패:', e);
    }

    // I31 셀
    try {
        const cell_I31 = worksheet.getCell('I31');
        cell_I31.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_I31.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_I31.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 I31 설정 실패:', e);
    }

    // I32 셀
    try {
        const cell_I32 = worksheet.getCell('I32');
        cell_I32.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_I32.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_I32.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 I32 설정 실패:', e);
    }

    // I33 셀
    try {
        const cell_I33 = worksheet.getCell('I33');
        cell_I33.value = '전용';
        cell_I33.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_I33.alignment = { horizontal: 'center', vertical: 'center' };
        cell_I33.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_I33.numFmt = '@';
    } catch (e) {
        console.warn('셀 I33 설정 실패:', e);
    }

    // I34 셀
    try {
        const cell_I34 = worksheet.getCell('I34');
        cell_I34.value = '162.07';
        cell_I34.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF0000FF' } };
        cell_I34.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEECEC' } };
        cell_I34.alignment = { horizontal: 'center', vertical: 'center' };
        cell_I34.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_I34.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 I34 설정 실패:', e);
    }

    // I35 셀
    try {
        const cell_I35 = worksheet.getCell('I35');
        cell_I35.value = '45.05';
        cell_I35.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_I35.alignment = { horizontal: 'center', vertical: 'center' };
        cell_I35.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_I35.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 I35 설정 실패:', e);
    }

    // I36 셀
    try {
        const cell_I36 = worksheet.getCell('I36');
        cell_I36.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_I36.alignment = { horizontal: 'center', vertical: 'center' };
        cell_I36.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_I36.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 I36 설정 실패:', e);
    }

    // I37 셀
    try {
        const cell_I37 = worksheet.getCell('I37');
        cell_I37.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_I37.alignment = { horizontal: 'center', vertical: 'center' };
        cell_I37.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_I37.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 I37 설정 실패:', e);
    }

    // I38 셀
    try {
        const cell_I38 = worksheet.getCell('I38');
        cell_I38.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_I38.alignment = { horizontal: 'center', vertical: 'center' };
        cell_I38.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_I38.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 I38 설정 실패:', e);
    }

    // I39 셀
    try {
        const cell_I39 = worksheet.getCell('I39');
        cell_I39.value = { formula: '=SUM(I34:I38)' };
        cell_I39.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_I39.alignment = { horizontal: 'center', vertical: 'center' };
        cell_I39.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_I39.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 I39 설정 실패:', e);
    }

    // I40 셀
    try {
        const cell_I40 = worksheet.getCell('I40');
        cell_I40.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_I40.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_I40.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 I40 설정 실패:', e);
    }

    // I41 셀
    try {
        const cell_I41 = worksheet.getCell('I41');
        cell_I41.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_I41.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_I41.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 I41 설정 실패:', e);
    }

    // I42 셀
    try {
        const cell_I42 = worksheet.getCell('I42');
        cell_I42.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_I42.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_I42.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 I42 설정 실패:', e);
    }

    // I43 셀
    try {
        const cell_I43 = worksheet.getCell('I43');
        cell_I43.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_I43.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_I43.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 I43 설정 실패:', e);
    }

    // I44 셀
    try {
        const cell_I44 = worksheet.getCell('I44');
        cell_I44.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_I44.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_I44.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 I44 설정 실패:', e);
    }

    // I45 셀
    try {
        const cell_I45 = worksheet.getCell('I45');
        cell_I45.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_I45.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_I45.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 I45 설정 실패:', e);
    }

    // I46 셀
    try {
        const cell_I46 = worksheet.getCell('I46');
        cell_I46.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_I46.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_I46.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 I46 설정 실패:', e);
    }

    // I47 셀
    try {
        const cell_I47 = worksheet.getCell('I47');
        cell_I47.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_I47.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_I47.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 I47 설정 실패:', e);
    }

    // I48 셀
    try {
        const cell_I48 = worksheet.getCell('I48');
        cell_I48.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_I48.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_I48.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 I48 설정 실패:', e);
    }

    // I49 셀
    try {
        const cell_I49 = worksheet.getCell('I49');
        cell_I49.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_I49.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_I49.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 I49 설정 실패:', e);
    }

    // I50 셀
    try {
        const cell_I50 = worksheet.getCell('I50');
        cell_I50.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_I50.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_I50.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 I50 설정 실패:', e);
    }

    // I51 셀
    try {
        const cell_I51 = worksheet.getCell('I51');
        cell_I51.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_I51.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_I51.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 I51 설정 실패:', e);
    }

    // I52 셀
    try {
        const cell_I52 = worksheet.getCell('I52');
        cell_I52.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_I52.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_I52.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 I52 설정 실패:', e);
    }

    // I53 셀
    try {
        const cell_I53 = worksheet.getCell('I53');
        cell_I53.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_I53.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_I53.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 I53 설정 실패:', e);
    }

    // I54 셀
    try {
        const cell_I54 = worksheet.getCell('I54');
        cell_I54.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_I54.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_I54.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 I54 설정 실패:', e);
    }

    // I55 셀
    try {
        const cell_I55 = worksheet.getCell('I55');
        cell_I55.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_I55.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_I55.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 I55 설정 실패:', e);
    }

    // I56 셀
    try {
        const cell_I56 = worksheet.getCell('I56');
        cell_I56.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_I56.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_I56.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 I56 설정 실패:', e);
    }

    // I57 셀
    try {
        const cell_I57 = worksheet.getCell('I57');
        cell_I57.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_I57.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_I57.border = { top: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 I57 설정 실패:', e);
    }

    // I58 셀
    try {
        const cell_I58 = worksheet.getCell('I58');
        cell_I58.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_I58.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_I58.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 I58 설정 실패:', e);
    }

    // I59 셀
    try {
        const cell_I59 = worksheet.getCell('I59');
        cell_I59.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_I59.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_I59.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 I59 설정 실패:', e);
    }

    // I6 셀
    try {
        const cell_I6 = worksheet.getCell('I6');
        cell_I6.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_I6.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_I6.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 I6 설정 실패:', e);
    }

    // I60 셀
    try {
        const cell_I60 = worksheet.getCell('I60');
        cell_I60.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_I60.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_I60.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 I60 설정 실패:', e);
    }

    // I61 셀
    try {
        const cell_I61 = worksheet.getCell('I61');
        cell_I61.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_I61.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_I61.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 I61 설정 실패:', e);
    }

    // I62 셀
    try {
        const cell_I62 = worksheet.getCell('I62');
        cell_I62.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_I62.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_I62.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 I62 설정 실패:', e);
    }

    // I63 셀
    try {
        const cell_I63 = worksheet.getCell('I63');
        cell_I63.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_I63.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_I63.border = { top: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 I63 설정 실패:', e);
    }

    // I7 셀
    try {
        const cell_I7 = worksheet.getCell('I7');
        cell_I7.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_I7.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_I7.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 I7 설정 실패:', e);
    }

    // I72 셀
    try {
        const cell_I72 = worksheet.getCell('I72');
        cell_I72.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_I72.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_I72.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 I72 설정 실패:', e);
    }

    // I73 셀
    try {
        const cell_I73 = worksheet.getCell('I73');
        cell_I73.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_I73.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_I73.border = { top: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 I73 설정 실패:', e);
    }

    // I8 셀
    try {
        const cell_I8 = worksheet.getCell('I8');
        cell_I8.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_I8.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_I8.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 I8 설정 실패:', e);
    }

    // I83 셀
    try {
        const cell_I83 = worksheet.getCell('I83');
        cell_I83.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_I83.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_I83.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 I83 설정 실패:', e);
    }

    // I9 셀
    try {
        const cell_I9 = worksheet.getCell('I9');
        cell_I9.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_I9.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_I9.border = { top: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 I9 설정 실패:', e);
    }

    // I91 셀
    try {
        const cell_I91 = worksheet.getCell('I91');
        cell_I91.font = { name: 'LG스마트체 Regular', size: 6.0, bold: true, color: { argb: 'FF000000' } };
        cell_I91.alignment = { horizontal: 'center', vertical: 'center' };
        cell_I91.numFmt = '_-* #,##0_-;\-* #,##0_-;_-* "-"_-;_-@_-';
    } catch (e) {
        console.warn('셀 I91 설정 실패:', e);
    }

    // J10 셀
    try {
        const cell_J10 = worksheet.getCell('J10');
        cell_J10.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J10.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J10.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J10 설정 실패:', e);
    }

    // J106 셀
    try {
        const cell_J106 = worksheet.getCell('J106');
        cell_J106.font = { name: 'LG스마트체 Regular', size: 10.0, bold: true, color: { argb: 'FF000000' } };
        cell_J106.alignment = { horizontal: 'center', vertical: 'center' };
        cell_J106.numFmt = '_-* #,##0_-;\-* #,##0_-;_-* "-"_-;_-@_-';
    } catch (e) {
        console.warn('셀 J106 설정 실패:', e);
    }

    // J11 셀
    try {
        const cell_J11 = worksheet.getCell('J11');
        cell_J11.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J11.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J11.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J11 설정 실패:', e);
    }

    // J12 셀
    try {
        const cell_J12 = worksheet.getCell('J12');
        cell_J12.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J12.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J12.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J12 설정 실패:', e);
    }

    // J13 셀
    try {
        const cell_J13 = worksheet.getCell('J13');
        cell_J13.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J13.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J13.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J13 설정 실패:', e);
    }

    // J14 셀
    try {
        const cell_J14 = worksheet.getCell('J14');
        cell_J14.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J14.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J14.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J14 설정 실패:', e);
    }

    // J15 셀
    try {
        const cell_J15 = worksheet.getCell('J15');
        cell_J15.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J15.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J15.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J15 설정 실패:', e);
    }

    // J16 셀
    try {
        const cell_J16 = worksheet.getCell('J16');
        cell_J16.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J16.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J16.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J16 설정 실패:', e);
    }

    // J17 셀
    try {
        const cell_J17 = worksheet.getCell('J17');
        cell_J17.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J17.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J17.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J17 설정 실패:', e);
    }

    // J18 셀
    try {
        const cell_J18 = worksheet.getCell('J18');
        cell_J18.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J18.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J18.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J18 설정 실패:', e);
    }

    // J19 셀
    try {
        const cell_J19 = worksheet.getCell('J19');
        cell_J19.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J19.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J19.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J19 설정 실패:', e);
    }

    // J20 셀
    try {
        const cell_J20 = worksheet.getCell('J20');
        cell_J20.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J20.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J20.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J20 설정 실패:', e);
    }

    // J21 셀
    try {
        const cell_J21 = worksheet.getCell('J21');
        cell_J21.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J21.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J21.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J21 설정 실패:', e);
    }

    // J22 셀
    try {
        const cell_J22 = worksheet.getCell('J22');
        cell_J22.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J22.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J22.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J22 설정 실패:', e);
    }

    // J23 셀
    try {
        const cell_J23 = worksheet.getCell('J23');
        cell_J23.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J23.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J23.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J23 설정 실패:', e);
    }

    // J24 셀
    try {
        const cell_J24 = worksheet.getCell('J24');
        cell_J24.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J24.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J24.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J24 설정 실패:', e);
    }

    // J25 셀
    try {
        const cell_J25 = worksheet.getCell('J25');
        cell_J25.value = '12987';
        cell_J25.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_J25.alignment = { horizontal: 'center', vertical: 'center' };
        cell_J25.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_J25.numFmt = '"("#,##0.0\ "㎡)"';
    } catch (e) {
        console.warn('셀 J25 설정 실패:', e);
    }

    // J26 셀
    try {
        const cell_J26 = worksheet.getCell('J26');
        cell_J26.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J26.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J26.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J26 설정 실패:', e);
    }

    // J27 셀
    try {
        const cell_J27 = worksheet.getCell('J27');
        cell_J27.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J27.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J27.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J27 설정 실패:', e);
    }

    // J28 셀
    try {
        const cell_J28 = worksheet.getCell('J28');
        cell_J28.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_J28.alignment = { horizontal: 'center', vertical: 'center' };
        cell_J28.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_J28.numFmt = '#,##0.000\ "평"';
    } catch (e) {
        console.warn('셀 J28 설정 실패:', e);
    }

    // J29 셀
    try {
        const cell_J29 = worksheet.getCell('J29');
        cell_J29.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J29.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J29.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J29 설정 실패:', e);
    }

    // J30 셀
    try {
        const cell_J30 = worksheet.getCell('J30');
        cell_J30.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J30.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J30.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J30 설정 실패:', e);
    }

    // J31 셀
    try {
        const cell_J31 = worksheet.getCell('J31');
        cell_J31.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J31.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J31.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J31 설정 실패:', e);
    }

    // J32 셀
    try {
        const cell_J32 = worksheet.getCell('J32');
        cell_J32.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J32.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J32.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J32 설정 실패:', e);
    }

    // J33 셀
    try {
        const cell_J33 = worksheet.getCell('J33');
        cell_J33.value = '임대';
        cell_J33.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_J33.alignment = { horizontal: 'center', vertical: 'center' };
        cell_J33.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_J33.numFmt = '@';
    } catch (e) {
        console.warn('셀 J33 설정 실패:', e);
    }

    // J34 셀
    try {
        const cell_J34 = worksheet.getCell('J34');
        cell_J34.value = '278.74';
        cell_J34.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_J34.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEECEC' } };
        cell_J34.alignment = { horizontal: 'center', vertical: 'center' };
        cell_J34.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_J34.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 J34 설정 실패:', e);
    }

    // J35 셀
    try {
        const cell_J35 = worksheet.getCell('J35');
        cell_J35.value = '82.47';
        cell_J35.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_J35.alignment = { horizontal: 'center', vertical: 'center' };
        cell_J35.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_J35.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 J35 설정 실패:', e);
    }

    // J36 셀
    try {
        const cell_J36 = worksheet.getCell('J36');
        cell_J36.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_J36.alignment = { horizontal: 'center', vertical: 'center' };
        cell_J36.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_J36.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 J36 설정 실패:', e);
    }

    // J37 셀
    try {
        const cell_J37 = worksheet.getCell('J37');
        cell_J37.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_J37.alignment = { horizontal: 'center', vertical: 'center' };
        cell_J37.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_J37.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 J37 설정 실패:', e);
    }

    // J38 셀
    try {
        const cell_J38 = worksheet.getCell('J38');
        cell_J38.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_J38.alignment = { horizontal: 'center', vertical: 'center' };
        cell_J38.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_J38.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 J38 설정 실패:', e);
    }

    // J39 셀
    try {
        const cell_J39 = worksheet.getCell('J39');
        cell_J39.value = { formula: '=SUM(J34:J38)' };
        cell_J39.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_J39.alignment = { horizontal: 'center', vertical: 'center' };
        cell_J39.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_J39.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 J39 설정 실패:', e);
    }

    // J40 셀
    try {
        const cell_J40 = worksheet.getCell('J40');
        cell_J40.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J40.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J40.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J40 설정 실패:', e);
    }

    // J41 셀
    try {
        const cell_J41 = worksheet.getCell('J41');
        cell_J41.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J41.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J41.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J41 설정 실패:', e);
    }

    // J42 셀
    try {
        const cell_J42 = worksheet.getCell('J42');
        cell_J42.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J42.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J42.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J42 설정 실패:', e);
    }

    // J43 셀
    try {
        const cell_J43 = worksheet.getCell('J43');
        cell_J43.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J43.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J43.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J43 설정 실패:', e);
    }

    // J44 셀
    try {
        const cell_J44 = worksheet.getCell('J44');
        cell_J44.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J44.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J44.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J44 설정 실패:', e);
    }

    // J45 셀
    try {
        const cell_J45 = worksheet.getCell('J45');
        cell_J45.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J45.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J45.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J45 설정 실패:', e);
    }

    // J46 셀
    try {
        const cell_J46 = worksheet.getCell('J46');
        cell_J46.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J46.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J46.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J46 설정 실패:', e);
    }

    // J47 셀
    try {
        const cell_J47 = worksheet.getCell('J47');
        cell_J47.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J47.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J47.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J47 설정 실패:', e);
    }

    // J48 셀
    try {
        const cell_J48 = worksheet.getCell('J48');
        cell_J48.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J48.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J48.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J48 설정 실패:', e);
    }

    // J49 셀
    try {
        const cell_J49 = worksheet.getCell('J49');
        cell_J49.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J49.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J49.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J49 설정 실패:', e);
    }

    // J50 셀
    try {
        const cell_J50 = worksheet.getCell('J50');
        cell_J50.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J50.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J50.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J50 설정 실패:', e);
    }

    // J51 셀
    try {
        const cell_J51 = worksheet.getCell('J51');
        cell_J51.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J51.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J51.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J51 설정 실패:', e);
    }

    // J52 셀
    try {
        const cell_J52 = worksheet.getCell('J52');
        cell_J52.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J52.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J52.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J52 설정 실패:', e);
    }

    // J53 셀
    try {
        const cell_J53 = worksheet.getCell('J53');
        cell_J53.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J53.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J53.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J53 설정 실패:', e);
    }

    // J54 셀
    try {
        const cell_J54 = worksheet.getCell('J54');
        cell_J54.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J54.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J54.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J54 설정 실패:', e);
    }

    // J55 셀
    try {
        const cell_J55 = worksheet.getCell('J55');
        cell_J55.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J55.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J55.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J55 설정 실패:', e);
    }

    // J56 셀
    try {
        const cell_J56 = worksheet.getCell('J56');
        cell_J56.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J56.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J56.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J56 설정 실패:', e);
    }

    // J57 셀
    try {
        const cell_J57 = worksheet.getCell('J57');
        cell_J57.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J57.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J57.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J57 설정 실패:', e);
    }

    // J58 셀
    try {
        const cell_J58 = worksheet.getCell('J58');
        cell_J58.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J58.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J58.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J58 설정 실패:', e);
    }

    // J59 셀
    try {
        const cell_J59 = worksheet.getCell('J59');
        cell_J59.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J59.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J59.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J59 설정 실패:', e);
    }

    // J6 셀
    try {
        const cell_J6 = worksheet.getCell('J6');
        cell_J6.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J6.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J6.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J6 설정 실패:', e);
    }

    // J60 셀
    try {
        const cell_J60 = worksheet.getCell('J60');
        cell_J60.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J60.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J60.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J60 설정 실패:', e);
    }

    // J61 셀
    try {
        const cell_J61 = worksheet.getCell('J61');
        cell_J61.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J61.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J61.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J61 설정 실패:', e);
    }

    // J62 셀
    try {
        const cell_J62 = worksheet.getCell('J62');
        cell_J62.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J62.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J62.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J62 설정 실패:', e);
    }

    // J63 셀
    try {
        const cell_J63 = worksheet.getCell('J63');
        cell_J63.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J63.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J63.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J63 설정 실패:', e);
    }

    // J64 셀
    try {
        const cell_J64 = worksheet.getCell('J64');
        cell_J64.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J64.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J64.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J64 설정 실패:', e);
    }

    // J65 셀
    try {
        const cell_J65 = worksheet.getCell('J65');
        cell_J65.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J65.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J65.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J65 설정 실패:', e);
    }

    // J66 셀
    try {
        const cell_J66 = worksheet.getCell('J66');
        cell_J66.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J66.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J66.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J66 설정 실패:', e);
    }

    // J67 셀
    try {
        const cell_J67 = worksheet.getCell('J67');
        cell_J67.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J67.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J67.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J67 설정 실패:', e);
    }

    // J68 셀
    try {
        const cell_J68 = worksheet.getCell('J68');
        cell_J68.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J68.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J68.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J68 설정 실패:', e);
    }

    // J69 셀
    try {
        const cell_J69 = worksheet.getCell('J69');
        cell_J69.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J69.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J69.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J69 설정 실패:', e);
    }

    // J7 셀
    try {
        const cell_J7 = worksheet.getCell('J7');
        cell_J7.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J7.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J7.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J7 설정 실패:', e);
    }

    // J70 셀
    try {
        const cell_J70 = worksheet.getCell('J70');
        cell_J70.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J70.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J70.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J70 설정 실패:', e);
    }

    // J71 셀
    try {
        const cell_J71 = worksheet.getCell('J71');
        cell_J71.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J71.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J71.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J71 설정 실패:', e);
    }

    // J72 셀
    try {
        const cell_J72 = worksheet.getCell('J72');
        cell_J72.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J72.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J72.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J72 설정 실패:', e);
    }

    // J73 셀
    try {
        const cell_J73 = worksheet.getCell('J73');
        cell_J73.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J73.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J73.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J73 설정 실패:', e);
    }

    // J74 셀
    try {
        const cell_J74 = worksheet.getCell('J74');
        cell_J74.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J74.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J74.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J74 설정 실패:', e);
    }

    // J75 셀
    try {
        const cell_J75 = worksheet.getCell('J75');
        cell_J75.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J75.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J75.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J75 설정 실패:', e);
    }

    // J76 셀
    try {
        const cell_J76 = worksheet.getCell('J76');
        cell_J76.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J76.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J76.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J76 설정 실패:', e);
    }

    // J77 셀
    try {
        const cell_J77 = worksheet.getCell('J77');
        cell_J77.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J77.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J77.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J77 설정 실패:', e);
    }

    // J78 셀
    try {
        const cell_J78 = worksheet.getCell('J78');
        cell_J78.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J78.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J78.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J78 설정 실패:', e);
    }

    // J79 셀
    try {
        const cell_J79 = worksheet.getCell('J79');
        cell_J79.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J79.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J79.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J79 설정 실패:', e);
    }

    // J8 셀
    try {
        const cell_J8 = worksheet.getCell('J8');
        cell_J8.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J8.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J8.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J8 설정 실패:', e);
    }

    // J80 셀
    try {
        const cell_J80 = worksheet.getCell('J80');
        cell_J80.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J80.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J80.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J80 설정 실패:', e);
    }

    // J81 셀
    try {
        const cell_J81 = worksheet.getCell('J81');
        cell_J81.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J81.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J81.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J81 설정 실패:', e);
    }

    // J82 셀
    try {
        const cell_J82 = worksheet.getCell('J82');
        cell_J82.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J82.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J82.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J82 설정 실패:', e);
    }

    // J83 셀
    try {
        const cell_J83 = worksheet.getCell('J83');
        cell_J83.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J83.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J83.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J83 설정 실패:', e);
    }

    // J9 셀
    try {
        const cell_J9 = worksheet.getCell('J9');
        cell_J9.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_J9.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_J9.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 J9 설정 실패:', e);
    }

    // J91 셀
    try {
        const cell_J91 = worksheet.getCell('J91');
        cell_J91.font = { name: 'LG스마트체 Regular', size: 6.0, bold: true, color: { argb: 'FF000000' } };
        cell_J91.alignment = { horizontal: 'center', vertical: 'center' };
        cell_J91.numFmt = '_-* #,##0_-;\-* #,##0_-;_-* "-"_-;_-@_-';
    } catch (e) {
        console.warn('셀 J91 설정 실패:', e);
    }

    // K106 셀
    try {
        const cell_K106 = worksheet.getCell('K106');
        cell_K106.font = { name: 'LG스마트체 Regular', size: 10.0, bold: true, color: { argb: 'FF000000' } };
        cell_K106.alignment = { horizontal: 'center', vertical: 'center' };
        cell_K106.numFmt = '_-* #,##0_-;\-* #,##0_-;_-* "-"_-;_-@_-';
    } catch (e) {
        console.warn('셀 K106 설정 실패:', e);
    }

    // K17 셀
    try {
        const cell_K17 = worksheet.getCell('K17');
        cell_K17.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_K17.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_K17.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 K17 설정 실패:', e);
    }

    // K18 셀
    try {
        const cell_K18 = worksheet.getCell('K18');
        cell_K18.value = '서울시 금천구 가산디지털1로 186';
        cell_K18.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_K18.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        cell_K18.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_K18.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 K18 설정 실패:', e);
    }

    // K19 셀
    try {
        const cell_K19 = worksheet.getCell('K19');
        cell_K19.value = '1,7호선 가산디지털단지역 도보 1분';
        cell_K19.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_K19.alignment = { horizontal: 'center', vertical: 'center' };
        cell_K19.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_K19.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 K19 설정 실패:', e);
    }

    // K20 셀
    try {
        const cell_K20 = worksheet.getCell('K20');
        cell_K20.value = '2007';
        cell_K20.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_K20.alignment = { horizontal: 'center', vertical: 'center' };
        cell_K20.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_K20.numFmt = '###0"년"';
    } catch (e) {
        console.warn('셀 K20 설정 실패:', e);
    }

    // K21 셀
    try {
        const cell_K21 = worksheet.getCell('K21');
        cell_K21.value = '154';
        cell_K21.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_K21.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        cell_K21.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_K21.numFmt = '##"F / B"#';
    } catch (e) {
        console.warn('셀 K21 설정 실패:', e);
    }

    // K22 셀
    try {
        const cell_K22 = worksheet.getCell('K22');
        cell_K22.value = '29745';
        cell_K22.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_K22.alignment = { horizontal: 'center', vertical: 'center' };
        cell_K22.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_K22.numFmt = '#,##0\ "평"';
    } catch (e) {
        console.warn('셀 K22 설정 실패:', e);
    }

    // K23 셀
    try {
        const cell_K23 = worksheet.getCell('K23');
        cell_K23.value = '878.52';
        cell_K23.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_K23.alignment = { horizontal: 'center', vertical: 'center' };
        cell_K23.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_K23.numFmt = '#,##0\ "평"';
    } catch (e) {
        console.warn('셀 K23 설정 실패:', e);
    }

    // K24 셀
    try {
        const cell_K24 = worksheet.getCell('K24');
        cell_K24.value = '0.581';
        cell_K24.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_K24.alignment = { horizontal: 'center', vertical: 'center' };
        cell_K24.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_K24.numFmt = '0.00%';
    } catch (e) {
        console.warn('셀 K24 설정 실패:', e);
    }

    // K25 셀
    try {
        const cell_K25 = worksheet.getCell('K25');
        cell_K25.value = { formula: '=M25*0.3025' };
        cell_K25.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_K25.alignment = { horizontal: 'center', vertical: 'center' };
        cell_K25.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_K25.numFmt = '#,##0\ "평"';
    } catch (e) {
        console.warn('셀 K25 설정 실패:', e);
    }

    // K26 셀
    try {
        const cell_K26 = worksheet.getCell('K26');
        cell_K26.value = '재능홀딩스';
        cell_K26.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_K26.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        cell_K26.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_K26.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 K26 설정 실패:', e);
    }

    // K27 셀
    try {
        const cell_K27 = worksheet.getCell('K27');
        cell_K27.value = '전세권, 근저당권 설정 가능';
        cell_K27.font = { name: 'LG스마트체 Regular', size: 9.0, bold: true, color: { argb: 'FFC00000' } };
        cell_K27.alignment = { horizontal: 'center', vertical: 'center' };
        cell_K27.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_K27.numFmt = '0.00%';
    } catch (e) {
        console.warn('셀 K27 설정 실패:', e);
    }

    // K28 셀
    try {
        const cell_K28 = worksheet.getCell('K28');
        cell_K28.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_K28.alignment = { horizontal: 'center', vertical: 'center' };
        cell_K28.border = { top: { style: 'hair', color: { argb: 'FF000000' } } };
        cell_K28.numFmt = '#,##0.000\ "평"';
    } catch (e) {
        console.warn('셀 K28 설정 실패:', e);
    }

    // K29 셀
    try {
        const cell_K29 = worksheet.getCell('K29');
        cell_K29.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_K29.alignment = { horizontal: 'center', vertical: 'center' };
        cell_K29.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_K29.numFmt = '_-* #,##0_-;\-* #,##0_-;_-* "-"_-;_-@_-';
    } catch (e) {
        console.warn('셀 K29 설정 실패:', e);
    }

    // K30 셀
    try {
        const cell_K30 = worksheet.getCell('K30');
        cell_K30.value = { formula: '=K29/K32' };
        cell_K30.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FFC00000' } };
        cell_K30.alignment = { horizontal: 'center', vertical: 'center' };
        cell_K30.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_K30.numFmt = '0.00%';
    } catch (e) {
        console.warn('셀 K30 설정 실패:', e);
    }

    // K31 셀
    try {
        const cell_K31 = worksheet.getCell('K31');
        cell_K31.value = '4768000';
        cell_K31.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_K31.alignment = { horizontal: 'center', vertical: 'center' };
        cell_K31.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_K31.numFmt = '#,###"원/㎡"';
    } catch (e) {
        console.warn('셀 K31 설정 실패:', e);
    }

    // K32 셀
    try {
        const cell_K32 = worksheet.getCell('K32');
        cell_K32.value = { formula: '=K31*M25' };
        cell_K32.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_K32.alignment = { horizontal: 'center', vertical: 'center' };
        cell_K32.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_K32.numFmt = '_-* #,##0_-;\-* #,##0_-;_-* "-"_-;_-@_-';
    } catch (e) {
        console.warn('셀 K32 설정 실패:', e);
    }

    // K33 셀
    try {
        const cell_K33 = worksheet.getCell('K33');
        cell_K33.value = '층';
        cell_K33.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_K33.alignment = { horizontal: 'center', vertical: 'center' };
        cell_K33.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_K33.numFmt = '@';
    } catch (e) {
        console.warn('셀 K33 설정 실패:', e);
    }

    // K34 셀
    try {
        const cell_K34 = worksheet.getCell('K34');
        cell_K34.value = '11';
        cell_K34.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF0000FF' } };
        cell_K34.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEECEC' } };
        cell_K34.alignment = { horizontal: 'center', vertical: 'center' };
        cell_K34.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_K34.numFmt = '#"층"';
    } catch (e) {
        console.warn('셀 K34 설정 실패:', e);
    }

    // K35 셀
    try {
        const cell_K35 = worksheet.getCell('K35');
        cell_K35.value = '2';
        cell_K35.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF0000FF' } };
        cell_K35.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEECEC' } };
        cell_K35.alignment = { horizontal: 'center', vertical: 'center' };
        cell_K35.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_K35.numFmt = '#"층"';
    } catch (e) {
        console.warn('셀 K35 설정 실패:', e);
    }

    // K36 셀
    try {
        const cell_K36 = worksheet.getCell('K36');
        cell_K36.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_K36.alignment = { horizontal: 'center', vertical: 'center' };
        cell_K36.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_K36.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 K36 설정 실패:', e);
    }

    // K37 셀
    try {
        const cell_K37 = worksheet.getCell('K37');
        cell_K37.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_K37.alignment = { horizontal: 'center', vertical: 'center' };
        cell_K37.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_K37.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 K37 설정 실패:', e);
    }

    // K38 셀
    try {
        const cell_K38 = worksheet.getCell('K38');
        cell_K38.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_K38.alignment = { horizontal: 'center', vertical: 'center' };
        cell_K38.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_K38.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 K38 설정 실패:', e);
    }

    // K39 셀
    try {
        const cell_K39 = worksheet.getCell('K39');
        cell_K39.value = '소계';
        cell_K39.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_K39.alignment = { horizontal: 'center', vertical: 'center' };
        cell_K39.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_K39.numFmt = '@';
    } catch (e) {
        console.warn('셀 K39 설정 실패:', e);
    }

    // K40 셀
    try {
        const cell_K40 = worksheet.getCell('K40');
        cell_K40.value = '2025.7~2027.6 (12개월)';
        cell_K40.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_K40.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        cell_K40.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_K40.numFmt = '#"개월 계약 가능"';
    } catch (e) {
        console.warn('셀 K40 설정 실패:', e);
    }

    // K41 셀
    try {
        const cell_K41 = worksheet.getCell('K41');
        cell_K41.value = '즉시';
        cell_K41.font = { name: 'LG스마트체 Regular', size: 9.0, bold: true, color: { argb: 'FF000000' } };
        cell_K41.alignment = { horizontal: 'center', vertical: 'center' };
        cell_K41.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_K41.numFmt = '#,##0\ "원"';
    } catch (e) {
        console.warn('셀 K41 설정 실패:', e);
    }

    // K42 셀
    try {
        const cell_K42 = worksheet.getCell('K42');
        cell_K42.value = '2층+11층 (일부)';
        cell_K42.font = { name: 'LG스마트체 Regular', size: 9.0, bold: true, color: { argb: 'FF000000' } };
        cell_K42.alignment = { horizontal: 'center', vertical: 'center' };
        cell_K42.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_K42.numFmt = '#,##0\ "층"';
    } catch (e) {
        console.warn('셀 K42 설정 실패:', e);
    }

    // K43 셀
    try {
        const cell_K43 = worksheet.getCell('K43');
        cell_K43.value = { formula: '=L35+L34' };
        cell_K43.font = { name: 'LG스마트체 Regular', size: 9.0, bold: true, color: { argb: 'FFC00000' } };
        cell_K43.alignment = { horizontal: 'center', vertical: 'center' };
        cell_K43.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_K43.numFmt = '#,##0\ "평"';
    } catch (e) {
        console.warn('셀 K43 설정 실패:', e);
    }

    // K44 셀
    try {
        const cell_K44 = worksheet.getCell('K44');
        cell_K44.value = { formula: '=M35+M34' };
        cell_K44.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_K44.alignment = { horizontal: 'center', vertical: 'center' };
        cell_K44.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_K44.numFmt = '#,##0\ "평"';
    } catch (e) {
        console.warn('셀 K44 설정 실패:', e);
    }

    // K45 셀
    try {
        const cell_K45 = worksheet.getCell('K45');
        cell_K45.value = '430000';
        cell_K45.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_K45.alignment = { horizontal: 'center', vertical: 'center' };
        cell_K45.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_K45.numFmt = '"@"#,###';
    } catch (e) {
        console.warn('셀 K45 설정 실패:', e);
    }

    // K46 셀
    try {
        const cell_K46 = worksheet.getCell('K46');
        cell_K46.value = '43000';
        cell_K46.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_K46.alignment = { horizontal: 'center', vertical: 'center' };
        cell_K46.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_K46.numFmt = '"@"#,###';
    } catch (e) {
        console.warn('셀 K46 설정 실패:', e);
    }

    // K47 셀
    try {
        const cell_K47 = worksheet.getCell('K47');
        cell_K47.value = '8000';
        cell_K47.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_K47.alignment = { horizontal: 'center', vertical: 'center' };
        cell_K47.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_K47.numFmt = '"@"#,###\+"실""비""별""도"';
    } catch (e) {
        console.warn('셀 K47 설정 실패:', e);
    }

    // K48 셀
    try {
        const cell_K48 = worksheet.getCell('K48');
        cell_K48.value = { formula: '=K46*(12-K49)/12' };
        cell_K48.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_K48.alignment = { horizontal: 'center', vertical: 'center' };
        cell_K48.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_K48.numFmt = '"@"#,###';
    } catch (e) {
        console.warn('셀 K48 설정 실패:', e);
    }

    // K49 셀
    try {
        const cell_K49 = worksheet.getCell('K49');
        cell_K49.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_K49.alignment = { horizontal: 'center', vertical: 'center' };
        cell_K49.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_K49.numFmt = '#0.0"개월"';
    } catch (e) {
        console.warn('셀 K49 설정 실패:', e);
    }

    // K50 셀
    try {
        const cell_K50 = worksheet.getCell('K50');
        cell_K50.value = { formula: '=K45*K44' };
        cell_K50.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_K50.alignment = { horizontal: 'center', vertical: 'center' };
        cell_K50.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_K50.numFmt = '#,##0\ "원"';
    } catch (e) {
        console.warn('셀 K50 설정 실패:', e);
    }

    // K51 셀
    try {
        const cell_K51 = worksheet.getCell('K51');
        cell_K51.value = { formula: '=K46*K44' };
        cell_K51.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_K51.alignment = { horizontal: 'center', vertical: 'center' };
        cell_K51.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_K51.numFmt = '#,##0\ "원"';
    } catch (e) {
        console.warn('셀 K51 설정 실패:', e);
    }

    // K52 셀
    try {
        const cell_K52 = worksheet.getCell('K52');
        cell_K52.value = { formula: '=K47*K44' };
        cell_K52.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_K52.alignment = { horizontal: 'center', vertical: 'center' };
        cell_K52.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_K52.numFmt = '#,##0\ "원"';
    } catch (e) {
        console.warn('셀 K52 설정 실패:', e);
    }

    // K53 셀
    try {
        const cell_K53 = worksheet.getCell('K53');
        cell_K53.value = '실비 관리비 : 전기세, 수도세  별도 부과';
        cell_K53.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FFC00000' } };
        cell_K53.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        cell_K53.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_K53.numFmt = '#,##0\ "원"';
    } catch (e) {
        console.warn('셀 K53 설정 실패:', e);
    }

    // K54 셀
    try {
        const cell_K54 = worksheet.getCell('K54');
        cell_K54.value = { formula: '=K51+K52' };
        cell_K54.font = { name: 'LG스마트체 Regular', size: 9.0, bold: true, color: { argb: 'FFC00000' } };
        cell_K54.alignment = { horizontal: 'center', vertical: 'center' };
        cell_K54.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_K54.numFmt = '#,##0\ "원"';
    } catch (e) {
        console.warn('셀 K54 설정 실패:', e);
    }

    // K55 셀
    try {
        const cell_K55 = worksheet.getCell('K55');
        cell_K55.value = { formula: '=K54*21' };
        cell_K55.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_K55.alignment = { horizontal: 'center', vertical: 'center' };
        cell_K55.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_K55.numFmt = '#,##0\ "원"';
    } catch (e) {
        console.warn('셀 K55 설정 실패:', e);
    }

    // K56 셀
    try {
        const cell_K56 = worksheet.getCell('K56');
        cell_K56.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_K56.alignment = { horizontal: 'center', vertical: 'center' };
        cell_K56.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_K56.numFmt = '0.#"개월"';
    } catch (e) {
        console.warn('셀 K56 설정 실패:', e);
    }

    // K57 셀
    try {
        const cell_K57 = worksheet.getCell('K57');
        cell_K57.value = '미제공';
        cell_K57.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_K57.alignment = { horizontal: 'center', vertical: 'center' };
        cell_K57.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_K57.numFmt = '"총액 "##,##0"원"';
    } catch (e) {
        console.warn('셀 K57 설정 실패:', e);
    }

    // K58 셀
    try {
        const cell_K58 = worksheet.getCell('K58');
        cell_K58.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_K58.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_K58.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 K58 설정 실패:', e);
    }

    // K59 셀
    try {
        const cell_K59 = worksheet.getCell('K59');
        cell_K59.value = { formula: '=732+44' };
        cell_K59.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_K59.alignment = { horizontal: 'center', vertical: 'center' };
        cell_K59.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_K59.numFmt = '#\ "대"';
    } catch (e) {
        console.warn('셀 K59 설정 실패:', e);
    }

    // K6 셀
    try {
        const cell_K6 = worksheet.getCell('K6');
        cell_K6.value = '가산디지털단지역';
        cell_K6.font = { name: 'LG스마트체 Regular', size: 10.0, color: { argb: 'FF000000' } };
        cell_K6.alignment = { horizontal: 'center', vertical: 'center' };
        cell_K6.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_K6.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 K6 설정 실패:', e);
    }

    // K60 셀
    try {
        const cell_K60 = worksheet.getCell('K60');
        cell_K60.value = '35';
        cell_K60.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_K60.alignment = { horizontal: 'center', vertical: 'center' };
        cell_K60.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_K60.numFmt = '"임대면적"\ #"평당 1대"';
    } catch (e) {
        console.warn('셀 K60 설정 실패:', e);
    }

    // K61 셀
    try {
        const cell_K61 = worksheet.getCell('K61');
        cell_K61.value = { formula: '=K44/K60' };
        cell_K61.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_K61.alignment = { horizontal: 'center', vertical: 'center' };
        cell_K61.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_K61.numFmt = '#,##0.0\ "대"';
    } catch (e) {
        console.warn('셀 K61 설정 실패:', e);
    }

    // K62 셀
    try {
        const cell_K62 = worksheet.getCell('K62');
        cell_K62.value = '10';
        cell_K62.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_K62.alignment = { horizontal: 'center', vertical: 'center' };
        cell_K62.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_K62.numFmt = '"월"#"만원/대"';
    } catch (e) {
        console.warn('셀 K62 설정 실패:', e);
    }

    // K63 셀
    try {
        const cell_K63 = worksheet.getCell('K63');
        cell_K63.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_K63.alignment = { horizontal: 'left', vertical: 'center', wrapText: true };
        cell_K63.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_K63.numFmt = '#,##0_ ';
    } catch (e) {
        console.warn('셀 K63 설정 실패:', e);
    }

    // K7 셀
    try {
        const cell_K7 = worksheet.getCell('K7');
        cell_K7.font = { name: 'LG스마트체 Regular', size: 10.0, color: { argb: 'FF000000' } };
        cell_K7.alignment = { horizontal: 'center', vertical: 'center' };
        cell_K7.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_K7.numFmt = '0_);[Red]\(0\)';
    } catch (e) {
        console.warn('셀 K7 설정 실패:', e);
    }

    // K72 셀
    try {
        const cell_K72 = worksheet.getCell('K72');
        cell_K72.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_K72.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_K72.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 K72 설정 실패:', e);
    }

    // K73 셀
    try {
        const cell_K73 = worksheet.getCell('K73');
        cell_K73.value = ' - 현재 2층 & 11층 일부 공실\n - 2층(204호) 지원시설, 11층(1103호) 업무시설\n - 11층 월 관리비 평균 170만원, 2층 실비 관리비\n   41만원\n - 실내청소 무료 서비스\n - 주차 : 자주식\n   ';
        cell_K73.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_K73.alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
        cell_K73.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_K73.numFmt = '#,##0\ "대"';
    } catch (e) {
        console.warn('셀 K73 설정 실패:', e);
    }

    // K8 셀
    try {
        const cell_K8 = worksheet.getCell('K8');
        cell_K8.value = '제이플라츠(2층+11층)';
        cell_K8.font = { name: 'LG스마트체 Regular', size: 10.0, color: { argb: 'FF000000' } };
        cell_K8.alignment = { horizontal: 'center', vertical: 'center' };
        cell_K8.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_K8.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 K8 설정 실패:', e);
    }

    // K83 셀
    try {
        const cell_K83 = worksheet.getCell('K83');
        cell_K83.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_K83.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_K83.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 K83 설정 실패:', e);
    }

    // K9 셀
    try {
        const cell_K9 = worksheet.getCell('K9');
        cell_K9.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_K9.alignment = { horizontal: 'center', vertical: 'center' };
        cell_K9.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_K9.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 K9 설정 실패:', e);
    }

    // K91 셀
    try {
        const cell_K91 = worksheet.getCell('K91');
        cell_K91.font = { name: 'LG스마트체 Regular', size: 6.0, bold: true, color: { argb: 'FF000000' } };
        cell_K91.alignment = { horizontal: 'center', vertical: 'center' };
        cell_K91.numFmt = '_-* #,##0_-;\-* #,##0_-;_-* "-"_-;_-@_-';
    } catch (e) {
        console.warn('셀 K91 설정 실패:', e);
    }

    // L106 셀
    try {
        const cell_L106 = worksheet.getCell('L106');
        cell_L106.font = { name: 'LG스마트체 Regular', size: 10.0, bold: true, color: { argb: 'FF000000' } };
        cell_L106.alignment = { horizontal: 'center', vertical: 'center' };
        cell_L106.numFmt = '_-* #,##0_-;\-* #,##0_-;_-* "-"_-;_-@_-';
    } catch (e) {
        console.warn('셀 L106 설정 실패:', e);
    }

    // L17 셀
    try {
        const cell_L17 = worksheet.getCell('L17');
        cell_L17.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_L17.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_L17.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 L17 설정 실패:', e);
    }

    // L18 셀
    try {
        const cell_L18 = worksheet.getCell('L18');
        cell_L18.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_L18.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_L18.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 L18 설정 실패:', e);
    }

    // L19 셀
    try {
        const cell_L19 = worksheet.getCell('L19');
        cell_L19.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_L19.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_L19.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 L19 설정 실패:', e);
    }

    // L20 셀
    try {
        const cell_L20 = worksheet.getCell('L20');
        cell_L20.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_L20.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_L20.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 L20 설정 실패:', e);
    }

    // L21 셀
    try {
        const cell_L21 = worksheet.getCell('L21');
        cell_L21.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_L21.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_L21.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 L21 설정 실패:', e);
    }

    // L22 셀
    try {
        const cell_L22 = worksheet.getCell('L22');
        cell_L22.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_L22.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_L22.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 L22 설정 실패:', e);
    }

    // L23 셀
    try {
        const cell_L23 = worksheet.getCell('L23');
        cell_L23.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_L23.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_L23.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 L23 설정 실패:', e);
    }

    // L24 셀
    try {
        const cell_L24 = worksheet.getCell('L24');
        cell_L24.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_L24.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_L24.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 L24 설정 실패:', e);
    }

    // L25 셀
    try {
        const cell_L25 = worksheet.getCell('L25');
        cell_L25.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_L25.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_L25.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 L25 설정 실패:', e);
    }

    // L26 셀
    try {
        const cell_L26 = worksheet.getCell('L26');
        cell_L26.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_L26.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_L26.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 L26 설정 실패:', e);
    }

    // L27 셀
    try {
        const cell_L27 = worksheet.getCell('L27');
        cell_L27.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_L27.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_L27.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 L27 설정 실패:', e);
    }

    // L28 셀
    try {
        const cell_L28 = worksheet.getCell('L28');
        cell_L28.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_L28.alignment = { horizontal: 'center', vertical: 'center' };
        cell_L28.border = { top: { style: 'hair', color: { argb: 'FF000000' } } };
        cell_L28.numFmt = '#,##0.000\ "평"';
    } catch (e) {
        console.warn('셀 L28 설정 실패:', e);
    }

    // L29 셀
    try {
        const cell_L29 = worksheet.getCell('L29');
        cell_L29.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_L29.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_L29.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 L29 설정 실패:', e);
    }

    // L30 셀
    try {
        const cell_L30 = worksheet.getCell('L30');
        cell_L30.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_L30.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_L30.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 L30 설정 실패:', e);
    }

    // L31 셀
    try {
        const cell_L31 = worksheet.getCell('L31');
        cell_L31.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_L31.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_L31.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 L31 설정 실패:', e);
    }

    // L32 셀
    try {
        const cell_L32 = worksheet.getCell('L32');
        cell_L32.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_L32.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_L32.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 L32 설정 실패:', e);
    }

    // L33 셀
    try {
        const cell_L33 = worksheet.getCell('L33');
        cell_L33.value = '전용';
        cell_L33.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_L33.alignment = { horizontal: 'center', vertical: 'center' };
        cell_L33.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_L33.numFmt = '@';
    } catch (e) {
        console.warn('셀 L33 설정 실패:', e);
    }

    // L34 셀
    try {
        const cell_L34 = worksheet.getCell('L34');
        cell_L34.value = '162.07';
        cell_L34.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF0000FF' } };
        cell_L34.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEECEC' } };
        cell_L34.alignment = { horizontal: 'center', vertical: 'center' };
        cell_L34.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_L34.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 L34 설정 실패:', e);
    }

    // L35 셀
    try {
        const cell_L35 = worksheet.getCell('L35');
        cell_L35.value = '45.05';
        cell_L35.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF0000FF' } };
        cell_L35.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEECEC' } };
        cell_L35.alignment = { horizontal: 'center', vertical: 'center' };
        cell_L35.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_L35.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 L35 설정 실패:', e);
    }

    // L36 셀
    try {
        const cell_L36 = worksheet.getCell('L36');
        cell_L36.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_L36.alignment = { horizontal: 'center', vertical: 'center' };
        cell_L36.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_L36.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 L36 설정 실패:', e);
    }

    // L37 셀
    try {
        const cell_L37 = worksheet.getCell('L37');
        cell_L37.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_L37.alignment = { horizontal: 'center', vertical: 'center' };
        cell_L37.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_L37.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 L37 설정 실패:', e);
    }

    // L38 셀
    try {
        const cell_L38 = worksheet.getCell('L38');
        cell_L38.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_L38.alignment = { horizontal: 'center', vertical: 'center' };
        cell_L38.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_L38.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 L38 설정 실패:', e);
    }

    // L39 셀
    try {
        const cell_L39 = worksheet.getCell('L39');
        cell_L39.value = { formula: '=SUM(L34:L38)' };
        cell_L39.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_L39.alignment = { horizontal: 'center', vertical: 'center' };
        cell_L39.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_L39.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 L39 설정 실패:', e);
    }

    // L40 셀
    try {
        const cell_L40 = worksheet.getCell('L40');
        cell_L40.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_L40.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_L40.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 L40 설정 실패:', e);
    }

    // L41 셀
    try {
        const cell_L41 = worksheet.getCell('L41');
        cell_L41.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_L41.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_L41.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 L41 설정 실패:', e);
    }

    // L42 셀
    try {
        const cell_L42 = worksheet.getCell('L42');
        cell_L42.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_L42.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_L42.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 L42 설정 실패:', e);
    }

    // L43 셀
    try {
        const cell_L43 = worksheet.getCell('L43');
        cell_L43.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_L43.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_L43.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 L43 설정 실패:', e);
    }

    // L44 셀
    try {
        const cell_L44 = worksheet.getCell('L44');
        cell_L44.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_L44.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_L44.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 L44 설정 실패:', e);
    }

    // L45 셀
    try {
        const cell_L45 = worksheet.getCell('L45');
        cell_L45.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_L45.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_L45.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 L45 설정 실패:', e);
    }

    // L46 셀
    try {
        const cell_L46 = worksheet.getCell('L46');
        cell_L46.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_L46.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_L46.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 L46 설정 실패:', e);
    }

    // L47 셀
    try {
        const cell_L47 = worksheet.getCell('L47');
        cell_L47.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_L47.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_L47.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 L47 설정 실패:', e);
    }

    // L48 셀
    try {
        const cell_L48 = worksheet.getCell('L48');
        cell_L48.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_L48.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_L48.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 L48 설정 실패:', e);
    }

    // L49 셀
    try {
        const cell_L49 = worksheet.getCell('L49');
        cell_L49.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_L49.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_L49.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 L49 설정 실패:', e);
    }

    // L50 셀
    try {
        const cell_L50 = worksheet.getCell('L50');
        cell_L50.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_L50.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_L50.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 L50 설정 실패:', e);
    }

    // L51 셀
    try {
        const cell_L51 = worksheet.getCell('L51');
        cell_L51.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_L51.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_L51.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 L51 설정 실패:', e);
    }

    // L52 셀
    try {
        const cell_L52 = worksheet.getCell('L52');
        cell_L52.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_L52.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_L52.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 L52 설정 실패:', e);
    }

    // L53 셀
    try {
        const cell_L53 = worksheet.getCell('L53');
        cell_L53.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_L53.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_L53.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 L53 설정 실패:', e);
    }

    // L54 셀
    try {
        const cell_L54 = worksheet.getCell('L54');
        cell_L54.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_L54.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_L54.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 L54 설정 실패:', e);
    }

    // L55 셀
    try {
        const cell_L55 = worksheet.getCell('L55');
        cell_L55.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_L55.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_L55.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 L55 설정 실패:', e);
    }

    // L56 셀
    try {
        const cell_L56 = worksheet.getCell('L56');
        cell_L56.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_L56.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_L56.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 L56 설정 실패:', e);
    }

    // L57 셀
    try {
        const cell_L57 = worksheet.getCell('L57');
        cell_L57.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_L57.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_L57.border = { top: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 L57 설정 실패:', e);
    }

    // L58 셀
    try {
        const cell_L58 = worksheet.getCell('L58');
        cell_L58.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_L58.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_L58.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 L58 설정 실패:', e);
    }

    // L59 셀
    try {
        const cell_L59 = worksheet.getCell('L59');
        cell_L59.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_L59.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_L59.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 L59 설정 실패:', e);
    }

    // L6 셀
    try {
        const cell_L6 = worksheet.getCell('L6');
        cell_L6.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_L6.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_L6.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 L6 설정 실패:', e);
    }

    // L60 셀
    try {
        const cell_L60 = worksheet.getCell('L60');
        cell_L60.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_L60.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_L60.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 L60 설정 실패:', e);
    }

    // L61 셀
    try {
        const cell_L61 = worksheet.getCell('L61');
        cell_L61.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_L61.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_L61.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 L61 설정 실패:', e);
    }

    // L62 셀
    try {
        const cell_L62 = worksheet.getCell('L62');
        cell_L62.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_L62.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_L62.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 L62 설정 실패:', e);
    }

    // L63 셀
    try {
        const cell_L63 = worksheet.getCell('L63');
        cell_L63.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_L63.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_L63.border = { top: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 L63 설정 실패:', e);
    }

    // L7 셀
    try {
        const cell_L7 = worksheet.getCell('L7');
        cell_L7.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_L7.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_L7.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 L7 설정 실패:', e);
    }

    // L72 셀
    try {
        const cell_L72 = worksheet.getCell('L72');
        cell_L72.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_L72.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_L72.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 L72 설정 실패:', e);
    }

    // L73 셀
    try {
        const cell_L73 = worksheet.getCell('L73');
        cell_L73.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_L73.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_L73.border = { top: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 L73 설정 실패:', e);
    }

    // L8 셀
    try {
        const cell_L8 = worksheet.getCell('L8');
        cell_L8.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_L8.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_L8.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 L8 설정 실패:', e);
    }

    // L83 셀
    try {
        const cell_L83 = worksheet.getCell('L83');
        cell_L83.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_L83.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_L83.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 L83 설정 실패:', e);
    }

    // L9 셀
    try {
        const cell_L9 = worksheet.getCell('L9');
        cell_L9.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_L9.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_L9.border = { top: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 L9 설정 실패:', e);
    }

    // L91 셀
    try {
        const cell_L91 = worksheet.getCell('L91');
        cell_L91.font = { name: 'LG스마트체 Regular', size: 6.0, bold: true, color: { argb: 'FF000000' } };
        cell_L91.alignment = { horizontal: 'center', vertical: 'center' };
        cell_L91.numFmt = '_-* #,##0_-;\-* #,##0_-;_-* "-"_-;_-@_-';
    } catch (e) {
        console.warn('셀 L91 설정 실패:', e);
    }

    // M10 셀
    try {
        const cell_M10 = worksheet.getCell('M10');
        cell_M10.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M10.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M10.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M10 설정 실패:', e);
    }

    // M106 셀
    try {
        const cell_M106 = worksheet.getCell('M106');
        cell_M106.font = { name: 'LG스마트체 Regular', size: 10.0, bold: true, color: { argb: 'FF000000' } };
        cell_M106.alignment = { horizontal: 'center', vertical: 'center' };
        cell_M106.numFmt = '_-* #,##0_-;\-* #,##0_-;_-* "-"_-;_-@_-';
    } catch (e) {
        console.warn('셀 M106 설정 실패:', e);
    }

    // M11 셀
    try {
        const cell_M11 = worksheet.getCell('M11');
        cell_M11.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M11.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M11.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M11 설정 실패:', e);
    }

    // M12 셀
    try {
        const cell_M12 = worksheet.getCell('M12');
        cell_M12.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M12.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M12.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M12 설정 실패:', e);
    }

    // M13 셀
    try {
        const cell_M13 = worksheet.getCell('M13');
        cell_M13.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M13.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M13.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M13 설정 실패:', e);
    }

    // M14 셀
    try {
        const cell_M14 = worksheet.getCell('M14');
        cell_M14.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M14.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M14.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M14 설정 실패:', e);
    }

    // M15 셀
    try {
        const cell_M15 = worksheet.getCell('M15');
        cell_M15.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M15.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M15.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M15 설정 실패:', e);
    }

    // M16 셀
    try {
        const cell_M16 = worksheet.getCell('M16');
        cell_M16.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M16.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M16.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M16 설정 실패:', e);
    }

    // M17 셀
    try {
        const cell_M17 = worksheet.getCell('M17');
        cell_M17.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M17.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M17.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M17 설정 실패:', e);
    }

    // M18 셀
    try {
        const cell_M18 = worksheet.getCell('M18');
        cell_M18.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M18.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M18.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M18 설정 실패:', e);
    }

    // M19 셀
    try {
        const cell_M19 = worksheet.getCell('M19');
        cell_M19.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M19.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M19.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M19 설정 실패:', e);
    }

    // M20 셀
    try {
        const cell_M20 = worksheet.getCell('M20');
        cell_M20.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M20.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M20.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M20 설정 실패:', e);
    }

    // M21 셀
    try {
        const cell_M21 = worksheet.getCell('M21');
        cell_M21.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M21.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M21.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M21 설정 실패:', e);
    }

    // M22 셀
    try {
        const cell_M22 = worksheet.getCell('M22');
        cell_M22.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M22.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M22.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M22 설정 실패:', e);
    }

    // M23 셀
    try {
        const cell_M23 = worksheet.getCell('M23');
        cell_M23.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M23.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M23.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M23 설정 실패:', e);
    }

    // M24 셀
    try {
        const cell_M24 = worksheet.getCell('M24');
        cell_M24.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M24.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M24.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M24 설정 실패:', e);
    }

    // M25 셀
    try {
        const cell_M25 = worksheet.getCell('M25');
        cell_M25.value = '12987';
        cell_M25.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_M25.alignment = { horizontal: 'center', vertical: 'center' };
        cell_M25.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_M25.numFmt = '"("#,##0.0\ "㎡)"';
    } catch (e) {
        console.warn('셀 M25 설정 실패:', e);
    }

    // M26 셀
    try {
        const cell_M26 = worksheet.getCell('M26');
        cell_M26.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M26.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M26.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M26 설정 실패:', e);
    }

    // M27 셀
    try {
        const cell_M27 = worksheet.getCell('M27');
        cell_M27.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M27.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M27.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M27 설정 실패:', e);
    }

    // M28 셀
    try {
        const cell_M28 = worksheet.getCell('M28');
        cell_M28.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_M28.alignment = { horizontal: 'center', vertical: 'center' };
        cell_M28.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_M28.numFmt = '#,##0.000\ "평"';
    } catch (e) {
        console.warn('셀 M28 설정 실패:', e);
    }

    // M29 셀
    try {
        const cell_M29 = worksheet.getCell('M29');
        cell_M29.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M29.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M29.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M29 설정 실패:', e);
    }

    // M30 셀
    try {
        const cell_M30 = worksheet.getCell('M30');
        cell_M30.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M30.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M30.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M30 설정 실패:', e);
    }

    // M31 셀
    try {
        const cell_M31 = worksheet.getCell('M31');
        cell_M31.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M31.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M31.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M31 설정 실패:', e);
    }

    // M32 셀
    try {
        const cell_M32 = worksheet.getCell('M32');
        cell_M32.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M32.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M32.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M32 설정 실패:', e);
    }

    // M33 셀
    try {
        const cell_M33 = worksheet.getCell('M33');
        cell_M33.value = '임대';
        cell_M33.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_M33.alignment = { horizontal: 'center', vertical: 'center' };
        cell_M33.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_M33.numFmt = '@';
    } catch (e) {
        console.warn('셀 M33 설정 실패:', e);
    }

    // M34 셀
    try {
        const cell_M34 = worksheet.getCell('M34');
        cell_M34.value = '278.74';
        cell_M34.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_M34.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEECEC' } };
        cell_M34.alignment = { horizontal: 'center', vertical: 'center' };
        cell_M34.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_M34.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 M34 설정 실패:', e);
    }

    // M35 셀
    try {
        const cell_M35 = worksheet.getCell('M35');
        cell_M35.value = '82.47';
        cell_M35.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_M35.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEECEC' } };
        cell_M35.alignment = { horizontal: 'center', vertical: 'center' };
        cell_M35.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_M35.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 M35 설정 실패:', e);
    }

    // M36 셀
    try {
        const cell_M36 = worksheet.getCell('M36');
        cell_M36.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_M36.alignment = { horizontal: 'center', vertical: 'center' };
        cell_M36.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_M36.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 M36 설정 실패:', e);
    }

    // M37 셀
    try {
        const cell_M37 = worksheet.getCell('M37');
        cell_M37.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_M37.alignment = { horizontal: 'center', vertical: 'center' };
        cell_M37.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_M37.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 M37 설정 실패:', e);
    }

    // M38 셀
    try {
        const cell_M38 = worksheet.getCell('M38');
        cell_M38.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_M38.alignment = { horizontal: 'center', vertical: 'center' };
        cell_M38.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_M38.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 M38 설정 실패:', e);
    }

    // M39 셀
    try {
        const cell_M39 = worksheet.getCell('M39');
        cell_M39.value = { formula: '=SUM(M34:M38)' };
        cell_M39.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_M39.alignment = { horizontal: 'center', vertical: 'center' };
        cell_M39.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_M39.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 M39 설정 실패:', e);
    }

    // M40 셀
    try {
        const cell_M40 = worksheet.getCell('M40');
        cell_M40.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M40.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M40.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M40 설정 실패:', e);
    }

    // M41 셀
    try {
        const cell_M41 = worksheet.getCell('M41');
        cell_M41.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M41.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M41.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M41 설정 실패:', e);
    }

    // M42 셀
    try {
        const cell_M42 = worksheet.getCell('M42');
        cell_M42.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M42.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M42.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M42 설정 실패:', e);
    }

    // M43 셀
    try {
        const cell_M43 = worksheet.getCell('M43');
        cell_M43.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M43.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M43.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M43 설정 실패:', e);
    }

    // M44 셀
    try {
        const cell_M44 = worksheet.getCell('M44');
        cell_M44.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M44.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M44.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M44 설정 실패:', e);
    }

    // M45 셀
    try {
        const cell_M45 = worksheet.getCell('M45');
        cell_M45.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M45.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M45.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M45 설정 실패:', e);
    }

    // M46 셀
    try {
        const cell_M46 = worksheet.getCell('M46');
        cell_M46.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M46.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M46.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M46 설정 실패:', e);
    }

    // M47 셀
    try {
        const cell_M47 = worksheet.getCell('M47');
        cell_M47.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M47.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M47.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M47 설정 실패:', e);
    }

    // M48 셀
    try {
        const cell_M48 = worksheet.getCell('M48');
        cell_M48.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M48.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M48.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M48 설정 실패:', e);
    }

    // M49 셀
    try {
        const cell_M49 = worksheet.getCell('M49');
        cell_M49.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M49.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M49.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M49 설정 실패:', e);
    }

    // M50 셀
    try {
        const cell_M50 = worksheet.getCell('M50');
        cell_M50.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M50.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M50.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M50 설정 실패:', e);
    }

    // M51 셀
    try {
        const cell_M51 = worksheet.getCell('M51');
        cell_M51.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M51.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M51.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M51 설정 실패:', e);
    }

    // M52 셀
    try {
        const cell_M52 = worksheet.getCell('M52');
        cell_M52.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M52.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M52.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M52 설정 실패:', e);
    }

    // M53 셀
    try {
        const cell_M53 = worksheet.getCell('M53');
        cell_M53.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M53.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M53.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M53 설정 실패:', e);
    }

    // M54 셀
    try {
        const cell_M54 = worksheet.getCell('M54');
        cell_M54.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M54.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M54.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M54 설정 실패:', e);
    }

    // M55 셀
    try {
        const cell_M55 = worksheet.getCell('M55');
        cell_M55.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M55.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M55.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M55 설정 실패:', e);
    }

    // M56 셀
    try {
        const cell_M56 = worksheet.getCell('M56');
        cell_M56.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M56.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M56.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M56 설정 실패:', e);
    }

    // M57 셀
    try {
        const cell_M57 = worksheet.getCell('M57');
        cell_M57.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M57.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M57.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M57 설정 실패:', e);
    }

    // M58 셀
    try {
        const cell_M58 = worksheet.getCell('M58');
        cell_M58.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M58.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M58.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M58 설정 실패:', e);
    }

    // M59 셀
    try {
        const cell_M59 = worksheet.getCell('M59');
        cell_M59.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M59.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M59.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M59 설정 실패:', e);
    }

    // M6 셀
    try {
        const cell_M6 = worksheet.getCell('M6');
        cell_M6.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M6.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M6.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M6 설정 실패:', e);
    }

    // M60 셀
    try {
        const cell_M60 = worksheet.getCell('M60');
        cell_M60.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M60.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M60.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M60 설정 실패:', e);
    }

    // M61 셀
    try {
        const cell_M61 = worksheet.getCell('M61');
        cell_M61.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M61.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M61.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M61 설정 실패:', e);
    }

    // M62 셀
    try {
        const cell_M62 = worksheet.getCell('M62');
        cell_M62.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M62.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M62.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M62 설정 실패:', e);
    }

    // M63 셀
    try {
        const cell_M63 = worksheet.getCell('M63');
        cell_M63.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M63.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M63.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M63 설정 실패:', e);
    }

    // M64 셀
    try {
        const cell_M64 = worksheet.getCell('M64');
        cell_M64.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M64.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M64.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M64 설정 실패:', e);
    }

    // M65 셀
    try {
        const cell_M65 = worksheet.getCell('M65');
        cell_M65.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M65.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M65.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M65 설정 실패:', e);
    }

    // M66 셀
    try {
        const cell_M66 = worksheet.getCell('M66');
        cell_M66.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M66.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M66.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M66 설정 실패:', e);
    }

    // M67 셀
    try {
        const cell_M67 = worksheet.getCell('M67');
        cell_M67.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M67.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M67.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M67 설정 실패:', e);
    }

    // M68 셀
    try {
        const cell_M68 = worksheet.getCell('M68');
        cell_M68.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M68.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M68.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M68 설정 실패:', e);
    }

    // M69 셀
    try {
        const cell_M69 = worksheet.getCell('M69');
        cell_M69.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M69.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M69.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M69 설정 실패:', e);
    }

    // M7 셀
    try {
        const cell_M7 = worksheet.getCell('M7');
        cell_M7.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M7.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M7.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M7 설정 실패:', e);
    }

    // M70 셀
    try {
        const cell_M70 = worksheet.getCell('M70');
        cell_M70.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M70.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M70.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M70 설정 실패:', e);
    }

    // M71 셀
    try {
        const cell_M71 = worksheet.getCell('M71');
        cell_M71.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M71.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M71.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M71 설정 실패:', e);
    }

    // M72 셀
    try {
        const cell_M72 = worksheet.getCell('M72');
        cell_M72.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M72.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M72.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M72 설정 실패:', e);
    }

    // M73 셀
    try {
        const cell_M73 = worksheet.getCell('M73');
        cell_M73.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M73.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M73.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M73 설정 실패:', e);
    }

    // M74 셀
    try {
        const cell_M74 = worksheet.getCell('M74');
        cell_M74.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M74.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M74.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M74 설정 실패:', e);
    }

    // M75 셀
    try {
        const cell_M75 = worksheet.getCell('M75');
        cell_M75.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M75.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M75.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M75 설정 실패:', e);
    }

    // M76 셀
    try {
        const cell_M76 = worksheet.getCell('M76');
        cell_M76.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M76.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M76.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M76 설정 실패:', e);
    }

    // M77 셀
    try {
        const cell_M77 = worksheet.getCell('M77');
        cell_M77.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M77.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M77.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M77 설정 실패:', e);
    }

    // M78 셀
    try {
        const cell_M78 = worksheet.getCell('M78');
        cell_M78.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M78.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M78.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M78 설정 실패:', e);
    }

    // M79 셀
    try {
        const cell_M79 = worksheet.getCell('M79');
        cell_M79.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M79.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M79.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M79 설정 실패:', e);
    }

    // M8 셀
    try {
        const cell_M8 = worksheet.getCell('M8');
        cell_M8.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M8.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M8.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M8 설정 실패:', e);
    }

    // M80 셀
    try {
        const cell_M80 = worksheet.getCell('M80');
        cell_M80.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M80.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M80.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M80 설정 실패:', e);
    }

    // M81 셀
    try {
        const cell_M81 = worksheet.getCell('M81');
        cell_M81.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M81.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M81.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M81 설정 실패:', e);
    }

    // M82 셀
    try {
        const cell_M82 = worksheet.getCell('M82');
        cell_M82.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M82.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M82.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M82 설정 실패:', e);
    }

    // M83 셀
    try {
        const cell_M83 = worksheet.getCell('M83');
        cell_M83.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M83.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M83.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M83 설정 실패:', e);
    }

    // M9 셀
    try {
        const cell_M9 = worksheet.getCell('M9');
        cell_M9.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_M9.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_M9.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 M9 설정 실패:', e);
    }

    // M91 셀
    try {
        const cell_M91 = worksheet.getCell('M91');
        cell_M91.font = { name: 'LG스마트체 Regular', size: 6.0, bold: true, color: { argb: 'FF000000' } };
        cell_M91.alignment = { horizontal: 'center', vertical: 'center' };
        cell_M91.numFmt = '_-* #,##0_-;\-* #,##0_-;_-* "-"_-;_-@_-';
    } catch (e) {
        console.warn('셀 M91 설정 실패:', e);
    }

    // N10 셀
    try {
        const cell_N10 = worksheet.getCell('N10');
        cell_N10.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_N10.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_N10.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 N10 설정 실패:', e);
    }

    // N106 셀
    try {
        const cell_N106 = worksheet.getCell('N106');
        cell_N106.font = { name: 'LG스마트체 Regular', size: 10.0, bold: true, color: { argb: 'FF000000' } };
        cell_N106.alignment = { horizontal: 'center', vertical: 'center' };
        cell_N106.numFmt = '_-* #,##0_-;\-* #,##0_-;_-* "-"_-;_-@_-';
    } catch (e) {
        console.warn('셀 N106 설정 실패:', e);
    }

    // N11 셀
    try {
        const cell_N11 = worksheet.getCell('N11');
        cell_N11.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_N11.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_N11.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 N11 설정 실패:', e);
    }

    // N12 셀
    try {
        const cell_N12 = worksheet.getCell('N12');
        cell_N12.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_N12.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_N12.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 N12 설정 실패:', e);
    }

    // N13 셀
    try {
        const cell_N13 = worksheet.getCell('N13');
        cell_N13.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_N13.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_N13.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 N13 설정 실패:', e);
    }

    // N14 셀
    try {
        const cell_N14 = worksheet.getCell('N14');
        cell_N14.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_N14.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_N14.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 N14 설정 실패:', e);
    }

    // N15 셀
    try {
        const cell_N15 = worksheet.getCell('N15');
        cell_N15.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_N15.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_N15.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 N15 설정 실패:', e);
    }

    // N16 셀
    try {
        const cell_N16 = worksheet.getCell('N16');
        cell_N16.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_N16.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_N16.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 N16 설정 실패:', e);
    }

    // N17 셀
    try {
        const cell_N17 = worksheet.getCell('N17');
        cell_N17.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_N17.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_N17.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 N17 설정 실패:', e);
    }

    // N18 셀
    try {
        const cell_N18 = worksheet.getCell('N18');
        cell_N18.value = '서울시 금천구 디지털로10길 9';
        cell_N18.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_N18.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        cell_N18.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_N18.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 N18 설정 실패:', e);
    }

    // N19 셀
    try {
        const cell_N19 = worksheet.getCell('N19');
        cell_N19.value = '1,7호선 가산디지털단지역 도보 10분';
        cell_N19.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_N19.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        cell_N19.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_N19.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 N19 설정 실패:', e);
    }

    // N20 셀
    try {
        const cell_N20 = worksheet.getCell('N20');
        cell_N20.value = '2013년';
        cell_N20.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_N20.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        cell_N20.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_N20.numFmt = '###0"년"';
    } catch (e) {
        console.warn('셀 N20 설정 실패:', e);
    }

    // N21 셀
    try {
        const cell_N21 = worksheet.getCell('N21');
        cell_N21.value = '205';
        cell_N21.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_N21.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        cell_N21.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_N21.numFmt = '##"F / B"#';
    } catch (e) {
        console.warn('셀 N21 설정 실패:', e);
    }

    // N22 셀
    try {
        const cell_N22 = worksheet.getCell('N22');
        cell_N22.value = { formula: '=98995.31*0.3025' };
        cell_N22.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_N22.alignment = { horizontal: 'center', vertical: 'center' };
        cell_N22.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_N22.numFmt = '#,##0\ "평"';
    } catch (e) {
        console.warn('셀 N22 설정 실패:', e);
    }

    // N23 셀
    try {
        const cell_N23 = worksheet.getCell('N23');
        cell_N23.value = '226.76';
        cell_N23.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_N23.alignment = { horizontal: 'center', vertical: 'center' };
        cell_N23.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_N23.numFmt = '#,##0\ "평"';
    } catch (e) {
        console.warn('셀 N23 설정 실패:', e);
    }

    // N24 셀
    try {
        const cell_N24 = worksheet.getCell('N24');
        cell_N24.value = '0.4794';
        cell_N24.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_N24.alignment = { horizontal: 'center', vertical: 'center' };
        cell_N24.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_N24.numFmt = '0.00%';
    } catch (e) {
        console.warn('셀 N24 설정 실패:', e);
    }

    // N25 셀
    try {
        const cell_N25 = worksheet.getCell('N25');
        cell_N25.value = { formula: '=P25*0.3025' };
        cell_N25.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_N25.alignment = { horizontal: 'center', vertical: 'center' };
        cell_N25.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_N25.numFmt = '#,##0\ "평"';
    } catch (e) {
        console.warn('셀 N25 설정 실패:', e);
    }

    // N26 셀
    try {
        const cell_N26 = worksheet.getCell('N26');
        cell_N26.value = '신한생명보험주식회사';
        cell_N26.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_N26.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        cell_N26.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_N26.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 N26 설정 실패:', e);
    }

    // N27 셀
    try {
        const cell_N27 = worksheet.getCell('N27');
        cell_N27.value = '근저당권 설정가능';
        cell_N27.font = { name: 'LG스마트체 Regular', size: 9.0, bold: true, color: { argb: 'FFC00000' } };
        cell_N27.alignment = { horizontal: 'center', vertical: 'center' };
        cell_N27.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_N27.numFmt = '0.00%';
    } catch (e) {
        console.warn('셀 N27 설정 실패:', e);
    }

    // N28 셀
    try {
        const cell_N28 = worksheet.getCell('N28');
        cell_N28.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_N28.alignment = { horizontal: 'center', vertical: 'center' };
        cell_N28.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_N28.numFmt = '#,##0.000\ "평"';
    } catch (e) {
        console.warn('셀 N28 설정 실패:', e);
    }

    // N29 셀
    try {
        const cell_N29 = worksheet.getCell('N29');
        cell_N29.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_N29.alignment = { horizontal: 'center', vertical: 'center' };
        cell_N29.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_N29.numFmt = '_-* #,##0_-;\-* #,##0_-;_-* "-"_-;_-@_-';
    } catch (e) {
        console.warn('셀 N29 설정 실패:', e);
    }

    // N30 셀
    try {
        const cell_N30 = worksheet.getCell('N30');
        cell_N30.value = { formula: '=N29/N32' };
        cell_N30.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FFC00000' } };
        cell_N30.alignment = { horizontal: 'center', vertical: 'center' };
        cell_N30.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_N30.numFmt = '0.00%';
    } catch (e) {
        console.warn('셀 N30 설정 실패:', e);
    }

    // N31 셀
    try {
        const cell_N31 = worksheet.getCell('N31');
        cell_N31.value = '6998000';
        cell_N31.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_N31.alignment = { horizontal: 'center', vertical: 'center' };
        cell_N31.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_N31.numFmt = '#,###"원/㎡"';
    } catch (e) {
        console.warn('셀 N31 설정 실패:', e);
    }

    // N32 셀
    try {
        const cell_N32 = worksheet.getCell('N32');
        cell_N32.value = { formula: '=N31*P25' };
        cell_N32.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_N32.alignment = { horizontal: 'center', vertical: 'center' };
        cell_N32.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_N32.numFmt = '_-* #,##0_-;\-* #,##0_-;_-* "-"_-;_-@_-';
    } catch (e) {
        console.warn('셀 N32 설정 실패:', e);
    }

    // N33 셀
    try {
        const cell_N33 = worksheet.getCell('N33');
        cell_N33.value = '층';
        cell_N33.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_N33.alignment = { horizontal: 'center', vertical: 'center' };
        cell_N33.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_N33.numFmt = '@';
    } catch (e) {
        console.warn('셀 N33 설정 실패:', e);
    }

    // N34 셀
    try {
        const cell_N34 = worksheet.getCell('N34');
        cell_N34.value = '11';
        cell_N34.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_N34.alignment = { horizontal: 'center', vertical: 'center' };
        cell_N34.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_N34.numFmt = '#"층"';
    } catch (e) {
        console.warn('셀 N34 설정 실패:', e);
    }

    // N35 셀
    try {
        const cell_N35 = worksheet.getCell('N35');
        cell_N35.value = '9';
        cell_N35.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF0000FF' } };
        cell_N35.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEECEC' } };
        cell_N35.alignment = { horizontal: 'center', vertical: 'center' };
        cell_N35.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_N35.numFmt = '#"층"';
    } catch (e) {
        console.warn('셀 N35 설정 실패:', e);
    }

    // N36 셀
    try {
        const cell_N36 = worksheet.getCell('N36');
        cell_N36.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF0000FF' } };
        cell_N36.alignment = { horizontal: 'center', vertical: 'center' };
        cell_N36.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_N36.numFmt = '##"층-"#';
    } catch (e) {
        console.warn('셀 N36 설정 실패:', e);
    }

    // N37 셀
    try {
        const cell_N37 = worksheet.getCell('N37');
        cell_N37.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF0000FF' } };
        cell_N37.alignment = { horizontal: 'center', vertical: 'center' };
        cell_N37.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_N37.numFmt = '##"층-"#';
    } catch (e) {
        console.warn('셀 N37 설정 실패:', e);
    }

    // N38 셀
    try {
        const cell_N38 = worksheet.getCell('N38');
        cell_N38.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF0000FF' } };
        cell_N38.alignment = { horizontal: 'center', vertical: 'center' };
        cell_N38.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_N38.numFmt = '##"층-"#';
    } catch (e) {
        console.warn('셀 N38 설정 실패:', e);
    }

    // N39 셀
    try {
        const cell_N39 = worksheet.getCell('N39');
        cell_N39.value = '소계';
        cell_N39.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_N39.alignment = { horizontal: 'center', vertical: 'center' };
        cell_N39.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_N39.numFmt = '@';
    } catch (e) {
        console.warn('셀 N39 설정 실패:', e);
    }

    // N40 셀
    try {
        const cell_N40 = worksheet.getCell('N40');
        cell_N40.value = '2025.7~2027.6 (12개월)';
        cell_N40.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_N40.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        cell_N40.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_N40.numFmt = '#"개월 계약 가능"';
    } catch (e) {
        console.warn('셀 N40 설정 실패:', e);
    }

    // N41 셀
    try {
        const cell_N41 = worksheet.getCell('N41');
        cell_N41.value = '즉시';
        cell_N41.font = { name: 'LG스마트체 Regular', size: 9.0, bold: true, color: { argb: 'FF000000' } };
        cell_N41.alignment = { horizontal: 'center', vertical: 'center' };
        cell_N41.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_N41.numFmt = '#,##0\ "원"';
    } catch (e) {
        console.warn('셀 N41 설정 실패:', e);
    }

    // N42 셀
    try {
        const cell_N42 = worksheet.getCell('N42');
        cell_N42.value = '9층 전체';
        cell_N42.font = { name: 'LG스마트체 Regular', size: 9.0, bold: true, color: { argb: 'FF000000' } };
        cell_N42.alignment = { horizontal: 'center', vertical: 'center' };
        cell_N42.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_N42.numFmt = '#,##0\ "층"';
    } catch (e) {
        console.warn('셀 N42 설정 실패:', e);
    }

    // N43 셀
    try {
        const cell_N43 = worksheet.getCell('N43');
        cell_N43.value = { formula: '=O35' };
        cell_N43.font = { name: 'LG스마트체 Regular', size: 9.0, bold: true, color: { argb: 'FFC00000' } };
        cell_N43.alignment = { horizontal: 'center', vertical: 'center' };
        cell_N43.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_N43.numFmt = '#,##0\ "평"';
    } catch (e) {
        console.warn('셀 N43 설정 실패:', e);
    }

    // N44 셀
    try {
        const cell_N44 = worksheet.getCell('N44');
        cell_N44.value = { formula: '=P35' };
        cell_N44.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_N44.alignment = { horizontal: 'center', vertical: 'center' };
        cell_N44.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_N44.numFmt = '#,##0\ "평"';
    } catch (e) {
        console.warn('셀 N44 설정 실패:', e);
    }

    // N45 셀
    try {
        const cell_N45 = worksheet.getCell('N45');
        cell_N45.value = '420000';
        cell_N45.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_N45.alignment = { horizontal: 'center', vertical: 'center' };
        cell_N45.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_N45.numFmt = '"@"#,###';
    } catch (e) {
        console.warn('셀 N45 설정 실패:', e);
    }

    // N46 셀
    try {
        const cell_N46 = worksheet.getCell('N46');
        cell_N46.value = '42000';
        cell_N46.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_N46.alignment = { horizontal: 'center', vertical: 'center' };
        cell_N46.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_N46.numFmt = '"@"#,###';
    } catch (e) {
        console.warn('셀 N46 설정 실패:', e);
    }

    // N47 셀
    try {
        const cell_N47 = worksheet.getCell('N47');
        cell_N47.value = '실비부과';
        cell_N47.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_N47.alignment = { horizontal: 'center', vertical: 'center' };
        cell_N47.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_N47.numFmt = '"@"#,###';
    } catch (e) {
        console.warn('셀 N47 설정 실패:', e);
    }

    // N48 셀
    try {
        const cell_N48 = worksheet.getCell('N48');
        cell_N48.value = { formula: '=N46*(12-N49)/12' };
        cell_N48.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_N48.alignment = { horizontal: 'center', vertical: 'center' };
        cell_N48.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_N48.numFmt = '"@"#,###';
    } catch (e) {
        console.warn('셀 N48 설정 실패:', e);
    }

    // N49 셀
    try {
        const cell_N49 = worksheet.getCell('N49');
        cell_N49.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_N49.alignment = { horizontal: 'center', vertical: 'center' };
        cell_N49.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_N49.numFmt = '#0.0"개월"';
    } catch (e) {
        console.warn('셀 N49 설정 실패:', e);
    }

    // N50 셀
    try {
        const cell_N50 = worksheet.getCell('N50');
        cell_N50.value = { formula: '=N45*N44' };
        cell_N50.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_N50.alignment = { horizontal: 'center', vertical: 'center' };
        cell_N50.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_N50.numFmt = '#,##0\ "원"';
    } catch (e) {
        console.warn('셀 N50 설정 실패:', e);
    }

    // N51 셀
    try {
        const cell_N51 = worksheet.getCell('N51');
        cell_N51.value = { formula: '=N46*N44' };
        cell_N51.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_N51.alignment = { horizontal: 'center', vertical: 'center' };
        cell_N51.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_N51.numFmt = '#,##0\ "원"';
    } catch (e) {
        console.warn('셀 N51 설정 실패:', e);
    }

    // N52 셀
    try {
        const cell_N52 = worksheet.getCell('N52');
        cell_N52.value = '-';
        cell_N52.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FFC00000' } };
        cell_N52.alignment = { horizontal: 'center', vertical: 'center' };
        cell_N52.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_N52.numFmt = '#,##0\ "원"';
    } catch (e) {
        console.warn('셀 N52 설정 실패:', e);
    }

    // N53 셀
    try {
        const cell_N53 = worksheet.getCell('N53');
        cell_N53.value = '관리비 전체 실비 부과';
        cell_N53.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FFC00000' } };
        cell_N53.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        cell_N53.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_N53.numFmt = '#,##0\ "원"';
    } catch (e) {
        console.warn('셀 N53 설정 실패:', e);
    }

    // N54 셀
    try {
        const cell_N54 = worksheet.getCell('N54');
        cell_N54.value = { formula: '=N51' };
        cell_N54.font = { name: 'LG스마트체 Regular', size: 9.0, bold: true, color: { argb: 'FFC00000' } };
        cell_N54.alignment = { horizontal: 'center', vertical: 'center' };
        cell_N54.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_N54.numFmt = '#,##0\ "원"';
    } catch (e) {
        console.warn('셀 N54 설정 실패:', e);
    }

    // N55 셀
    try {
        const cell_N55 = worksheet.getCell('N55');
        cell_N55.value = { formula: '=N54*21' };
        cell_N55.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_N55.alignment = { horizontal: 'center', vertical: 'center' };
        cell_N55.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_N55.numFmt = '#,##0\ "원"';
    } catch (e) {
        console.warn('셀 N55 설정 실패:', e);
    }

    // N56 셀
    try {
        const cell_N56 = worksheet.getCell('N56');
        cell_N56.value = '0.5';
        cell_N56.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_N56.alignment = { horizontal: 'center', vertical: 'center' };
        cell_N56.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_N56.numFmt = '0.#"개월"';
    } catch (e) {
        console.warn('셀 N56 설정 실패:', e);
    }

    // N57 셀
    try {
        const cell_N57 = worksheet.getCell('N57');
        cell_N57.value = '미제공';
        cell_N57.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_N57.alignment = { horizontal: 'center', vertical: 'center' };
        cell_N57.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_N57.numFmt = '"총액 "##,##0"원"';
    } catch (e) {
        console.warn('셀 N57 설정 실패:', e);
    }

    // N58 셀
    try {
        const cell_N58 = worksheet.getCell('N58');
        cell_N58.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_N58.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_N58.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 N58 설정 실패:', e);
    }

    // N59 셀
    try {
        const cell_N59 = worksheet.getCell('N59');
        cell_N59.value = '902';
        cell_N59.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_N59.alignment = { horizontal: 'center', vertical: 'center' };
        cell_N59.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_N59.numFmt = '#\ "대"';
    } catch (e) {
        console.warn('셀 N59 설정 실패:', e);
    }

    // N6 셀
    try {
        const cell_N6 = worksheet.getCell('N6');
        cell_N6.value = '가산디지털단지역';
        cell_N6.font = { name: 'LG스마트체 Regular', size: 10.0, color: { argb: 'FF000000' } };
        cell_N6.alignment = { horizontal: 'center', vertical: 'center' };
        cell_N6.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_N6.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 N6 설정 실패:', e);
    }

    // N60 셀
    try {
        const cell_N60 = worksheet.getCell('N60');
        cell_N60.value = '40';
        cell_N60.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_N60.alignment = { horizontal: 'center', vertical: 'center' };
        cell_N60.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_N60.numFmt = '"임대면적"\ #"평당 1대"';
    } catch (e) {
        console.warn('셀 N60 설정 실패:', e);
    }

    // N61 셀
    try {
        const cell_N61 = worksheet.getCell('N61');
        cell_N61.value = { formula: '=N44/N60' };
        cell_N61.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_N61.alignment = { horizontal: 'center', vertical: 'center' };
        cell_N61.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_N61.numFmt = '#,##0.0\ "대"';
    } catch (e) {
        console.warn('셀 N61 설정 실패:', e);
    }

    // N62 셀
    try {
        const cell_N62 = worksheet.getCell('N62');
        cell_N62.value = '7';
        cell_N62.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_N62.alignment = { horizontal: 'center', vertical: 'center' };
        cell_N62.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_N62.numFmt = '"월"#"만원/대"';
    } catch (e) {
        console.warn('셀 N62 설정 실패:', e);
    }

    // N63 셀
    try {
        const cell_N63 = worksheet.getCell('N63');
        cell_N63.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_N63.alignment = { horizontal: 'left', vertical: 'center', wrapText: true };
        cell_N63.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_N63.numFmt = '#,##0_ ';
    } catch (e) {
        console.warn('셀 N63 설정 실패:', e);
    }

    // N64 셀
    try {
        const cell_N64 = worksheet.getCell('N64');
        cell_N64.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_N64.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_N64.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 N64 설정 실패:', e);
    }

    // N65 셀
    try {
        const cell_N65 = worksheet.getCell('N65');
        cell_N65.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_N65.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_N65.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 N65 설정 실패:', e);
    }

    // N66 셀
    try {
        const cell_N66 = worksheet.getCell('N66');
        cell_N66.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_N66.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_N66.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 N66 설정 실패:', e);
    }

    // N67 셀
    try {
        const cell_N67 = worksheet.getCell('N67');
        cell_N67.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_N67.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_N67.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 N67 설정 실패:', e);
    }

    // N68 셀
    try {
        const cell_N68 = worksheet.getCell('N68');
        cell_N68.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_N68.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_N68.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 N68 설정 실패:', e);
    }

    // N69 셀
    try {
        const cell_N69 = worksheet.getCell('N69');
        cell_N69.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_N69.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_N69.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 N69 설정 실패:', e);
    }

    // N7 셀
    try {
        const cell_N7 = worksheet.getCell('N7');
        cell_N7.font = { name: 'LG스마트체 Regular', size: 10.0, color: { argb: 'FF000000' } };
        cell_N7.alignment = { horizontal: 'center', vertical: 'center' };
        cell_N7.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_N7.numFmt = '0_);[Red]\(0\)';
    } catch (e) {
        console.warn('셀 N7 설정 실패:', e);
    }

    // N70 셀
    try {
        const cell_N70 = worksheet.getCell('N70');
        cell_N70.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_N70.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_N70.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 N70 설정 실패:', e);
    }

    // N71 셀
    try {
        const cell_N71 = worksheet.getCell('N71');
        cell_N71.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_N71.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_N71.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 N71 설정 실패:', e);
    }

    // N72 셀
    try {
        const cell_N72 = worksheet.getCell('N72');
        cell_N72.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_N72.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_N72.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 N72 설정 실패:', e);
    }

    // N73 셀
    try {
        const cell_N73 = worksheet.getCell('N73');
        cell_N73.value = ' - 현재 9층 전체 즉시 가능\n - 11층(4-5호) 일부 즉시 가능\n - Rent free  제공 불가 (단기, 소형면적으로)\n - 인테리어공사기간 15일 제공 \n   (관리비 부과)\n - 화장실 : 남 - 소변기2개+양변기2개\n                 여 - 양변기2개\n \n\n\n';
        cell_N73.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_N73.alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
        cell_N73.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_N73.numFmt = '#,##0\ "대"';
    } catch (e) {
        console.warn('셀 N73 설정 실패:', e);
    }

    // N74 셀
    try {
        const cell_N74 = worksheet.getCell('N74');
        cell_N74.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_N74.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_N74.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 N74 설정 실패:', e);
    }

    // N75 셀
    try {
        const cell_N75 = worksheet.getCell('N75');
        cell_N75.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_N75.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_N75.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 N75 설정 실패:', e);
    }

    // N76 셀
    try {
        const cell_N76 = worksheet.getCell('N76');
        cell_N76.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_N76.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_N76.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 N76 설정 실패:', e);
    }

    // N77 셀
    try {
        const cell_N77 = worksheet.getCell('N77');
        cell_N77.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_N77.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_N77.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 N77 설정 실패:', e);
    }

    // N78 셀
    try {
        const cell_N78 = worksheet.getCell('N78');
        cell_N78.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_N78.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_N78.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 N78 설정 실패:', e);
    }

    // N79 셀
    try {
        const cell_N79 = worksheet.getCell('N79');
        cell_N79.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_N79.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_N79.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 N79 설정 실패:', e);
    }

    // N8 셀
    try {
        const cell_N8 = worksheet.getCell('N8');
        cell_N8.value = '하이힐빌딩(집합건물)';
        cell_N8.font = { name: 'LG스마트체 Regular', size: 10.0, color: { argb: 'FF000000' } };
        cell_N8.alignment = { horizontal: 'center', vertical: 'center' };
        cell_N8.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_N8.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 N8 설정 실패:', e);
    }

    // N80 셀
    try {
        const cell_N80 = worksheet.getCell('N80');
        cell_N80.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_N80.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_N80.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 N80 설정 실패:', e);
    }

    // N81 셀
    try {
        const cell_N81 = worksheet.getCell('N81');
        cell_N81.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_N81.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_N81.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 N81 설정 실패:', e);
    }

    // N82 셀
    try {
        const cell_N82 = worksheet.getCell('N82');
        cell_N82.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_N82.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_N82.border = { left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 N82 설정 실패:', e);
    }

    // N83 셀
    try {
        const cell_N83 = worksheet.getCell('N83');
        cell_N83.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_N83.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_N83.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 N83 설정 실패:', e);
    }

    // N9 셀
    try {
        const cell_N9 = worksheet.getCell('N9');
        cell_N9.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_N9.alignment = { horizontal: 'center', vertical: 'center' };
        cell_N9.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_N9.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 N9 설정 실패:', e);
    }

    // N91 셀
    try {
        const cell_N91 = worksheet.getCell('N91');
        cell_N91.font = { name: 'LG스마트체 Regular', size: 6.0, bold: true, color: { argb: 'FF000000' } };
        cell_N91.alignment = { horizontal: 'center', vertical: 'center' };
        cell_N91.numFmt = '_-* #,##0_-;\-* #,##0_-;_-* "-"_-;_-@_-';
    } catch (e) {
        console.warn('셀 N91 설정 실패:', e);
    }

    // O106 셀
    try {
        const cell_O106 = worksheet.getCell('O106');
        cell_O106.font = { name: 'LG스마트체 Regular', size: 10.0, bold: true, color: { argb: 'FF000000' } };
        cell_O106.alignment = { horizontal: 'center', vertical: 'center' };
        cell_O106.numFmt = '_-* #,##0_-;\-* #,##0_-;_-* "-"_-;_-@_-';
    } catch (e) {
        console.warn('셀 O106 설정 실패:', e);
    }

    // O17 셀
    try {
        const cell_O17 = worksheet.getCell('O17');
        cell_O17.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_O17.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_O17.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 O17 설정 실패:', e);
    }

    // O18 셀
    try {
        const cell_O18 = worksheet.getCell('O18');
        cell_O18.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_O18.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_O18.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 O18 설정 실패:', e);
    }

    // O19 셀
    try {
        const cell_O19 = worksheet.getCell('O19');
        cell_O19.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_O19.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_O19.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 O19 설정 실패:', e);
    }

    // O20 셀
    try {
        const cell_O20 = worksheet.getCell('O20');
        cell_O20.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_O20.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_O20.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 O20 설정 실패:', e);
    }

    // O21 셀
    try {
        const cell_O21 = worksheet.getCell('O21');
        cell_O21.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_O21.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_O21.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 O21 설정 실패:', e);
    }

    // O22 셀
    try {
        const cell_O22 = worksheet.getCell('O22');
        cell_O22.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_O22.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_O22.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 O22 설정 실패:', e);
    }

    // O23 셀
    try {
        const cell_O23 = worksheet.getCell('O23');
        cell_O23.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_O23.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_O23.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 O23 설정 실패:', e);
    }

    // O24 셀
    try {
        const cell_O24 = worksheet.getCell('O24');
        cell_O24.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_O24.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_O24.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 O24 설정 실패:', e);
    }

    // O25 셀
    try {
        const cell_O25 = worksheet.getCell('O25');
        cell_O25.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_O25.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_O25.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 O25 설정 실패:', e);
    }

    // O26 셀
    try {
        const cell_O26 = worksheet.getCell('O26');
        cell_O26.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_O26.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_O26.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 O26 설정 실패:', e);
    }

    // O27 셀
    try {
        const cell_O27 = worksheet.getCell('O27');
        cell_O27.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_O27.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_O27.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 O27 설정 실패:', e);
    }

    // O28 셀
    try {
        const cell_O28 = worksheet.getCell('O28');
        cell_O28.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_O28.alignment = { horizontal: 'center', vertical: 'center' };
        cell_O28.border = { top: { style: 'hair', color: { argb: 'FF000000' } } };
        cell_O28.numFmt = '#,##0.000\ "평"';
    } catch (e) {
        console.warn('셀 O28 설정 실패:', e);
    }

    // O29 셀
    try {
        const cell_O29 = worksheet.getCell('O29');
        cell_O29.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_O29.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_O29.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 O29 설정 실패:', e);
    }

    // O30 셀
    try {
        const cell_O30 = worksheet.getCell('O30');
        cell_O30.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_O30.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_O30.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 O30 설정 실패:', e);
    }

    // O31 셀
    try {
        const cell_O31 = worksheet.getCell('O31');
        cell_O31.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_O31.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_O31.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 O31 설정 실패:', e);
    }

    // O32 셀
    try {
        const cell_O32 = worksheet.getCell('O32');
        cell_O32.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_O32.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_O32.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 O32 설정 실패:', e);
    }

    // O33 셀
    try {
        const cell_O33 = worksheet.getCell('O33');
        cell_O33.value = '전용';
        cell_O33.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_O33.alignment = { horizontal: 'center', vertical: 'center' };
        cell_O33.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_O33.numFmt = '@';
    } catch (e) {
        console.warn('셀 O33 설정 실패:', e);
    }

    // O34 셀
    try {
        const cell_O34 = worksheet.getCell('O34');
        cell_O34.value = '110.2';
        cell_O34.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_O34.alignment = { horizontal: 'center', vertical: 'center' };
        cell_O34.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_O34.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 O34 설정 실패:', e);
    }

    // O35 셀
    try {
        const cell_O35 = worksheet.getCell('O35');
        cell_O35.value = '226.7';
        cell_O35.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF0000FF' } };
        cell_O35.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEECEC' } };
        cell_O35.alignment = { horizontal: 'center', vertical: 'center' };
        cell_O35.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_O35.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 O35 설정 실패:', e);
    }

    // O36 셀
    try {
        const cell_O36 = worksheet.getCell('O36');
        cell_O36.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF0000FF' } };
        cell_O36.alignment = { horizontal: 'center', vertical: 'center' };
        cell_O36.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_O36.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 O36 설정 실패:', e);
    }

    // O37 셀
    try {
        const cell_O37 = worksheet.getCell('O37');
        cell_O37.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF0000FF' } };
        cell_O37.alignment = { horizontal: 'center', vertical: 'center' };
        cell_O37.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_O37.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 O37 설정 실패:', e);
    }

    // O38 셀
    try {
        const cell_O38 = worksheet.getCell('O38');
        cell_O38.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF0000FF' } };
        cell_O38.alignment = { horizontal: 'center', vertical: 'center' };
        cell_O38.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_O38.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 O38 설정 실패:', e);
    }

    // O39 셀
    try {
        const cell_O39 = worksheet.getCell('O39');
        cell_O39.value = { formula: '=SUM(O34:O38)' };
        cell_O39.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_O39.alignment = { horizontal: 'center', vertical: 'center' };
        cell_O39.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_O39.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 O39 설정 실패:', e);
    }

    // O40 셀
    try {
        const cell_O40 = worksheet.getCell('O40');
        cell_O40.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_O40.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_O40.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 O40 설정 실패:', e);
    }

    // O41 셀
    try {
        const cell_O41 = worksheet.getCell('O41');
        cell_O41.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_O41.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_O41.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 O41 설정 실패:', e);
    }

    // O42 셀
    try {
        const cell_O42 = worksheet.getCell('O42');
        cell_O42.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_O42.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_O42.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 O42 설정 실패:', e);
    }

    // O43 셀
    try {
        const cell_O43 = worksheet.getCell('O43');
        cell_O43.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_O43.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_O43.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 O43 설정 실패:', e);
    }

    // O44 셀
    try {
        const cell_O44 = worksheet.getCell('O44');
        cell_O44.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_O44.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_O44.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 O44 설정 실패:', e);
    }

    // O45 셀
    try {
        const cell_O45 = worksheet.getCell('O45');
        cell_O45.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_O45.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_O45.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 O45 설정 실패:', e);
    }

    // O46 셀
    try {
        const cell_O46 = worksheet.getCell('O46');
        cell_O46.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_O46.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_O46.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 O46 설정 실패:', e);
    }

    // O47 셀
    try {
        const cell_O47 = worksheet.getCell('O47');
        cell_O47.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_O47.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_O47.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 O47 설정 실패:', e);
    }

    // O48 셀
    try {
        const cell_O48 = worksheet.getCell('O48');
        cell_O48.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_O48.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_O48.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 O48 설정 실패:', e);
    }

    // O49 셀
    try {
        const cell_O49 = worksheet.getCell('O49');
        cell_O49.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_O49.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_O49.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 O49 설정 실패:', e);
    }

    // O50 셀
    try {
        const cell_O50 = worksheet.getCell('O50');
        cell_O50.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_O50.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_O50.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 O50 설정 실패:', e);
    }

    // O51 셀
    try {
        const cell_O51 = worksheet.getCell('O51');
        cell_O51.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_O51.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_O51.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 O51 설정 실패:', e);
    }

    // O52 셀
    try {
        const cell_O52 = worksheet.getCell('O52');
        cell_O52.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_O52.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_O52.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 O52 설정 실패:', e);
    }

    // O53 셀
    try {
        const cell_O53 = worksheet.getCell('O53');
        cell_O53.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_O53.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_O53.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 O53 설정 실패:', e);
    }

    // O54 셀
    try {
        const cell_O54 = worksheet.getCell('O54');
        cell_O54.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_O54.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_O54.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 O54 설정 실패:', e);
    }

    // O55 셀
    try {
        const cell_O55 = worksheet.getCell('O55');
        cell_O55.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_O55.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_O55.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 O55 설정 실패:', e);
    }

    // O56 셀
    try {
        const cell_O56 = worksheet.getCell('O56');
        cell_O56.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_O56.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_O56.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 O56 설정 실패:', e);
    }

    // O57 셀
    try {
        const cell_O57 = worksheet.getCell('O57');
        cell_O57.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_O57.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_O57.border = { top: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 O57 설정 실패:', e);
    }

    // O58 셀
    try {
        const cell_O58 = worksheet.getCell('O58');
        cell_O58.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_O58.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_O58.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 O58 설정 실패:', e);
    }

    // O59 셀
    try {
        const cell_O59 = worksheet.getCell('O59');
        cell_O59.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_O59.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_O59.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 O59 설정 실패:', e);
    }

    // O6 셀
    try {
        const cell_O6 = worksheet.getCell('O6');
        cell_O6.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_O6.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_O6.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 O6 설정 실패:', e);
    }

    // O60 셀
    try {
        const cell_O60 = worksheet.getCell('O60');
        cell_O60.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_O60.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_O60.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 O60 설정 실패:', e);
    }

    // O61 셀
    try {
        const cell_O61 = worksheet.getCell('O61');
        cell_O61.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_O61.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_O61.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 O61 설정 실패:', e);
    }

    // O62 셀
    try {
        const cell_O62 = worksheet.getCell('O62');
        cell_O62.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_O62.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_O62.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 O62 설정 실패:', e);
    }

    // O7 셀
    try {
        const cell_O7 = worksheet.getCell('O7');
        cell_O7.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_O7.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_O7.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 O7 설정 실패:', e);
    }

    // O72 셀
    try {
        const cell_O72 = worksheet.getCell('O72');
        cell_O72.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_O72.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_O72.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 O72 설정 실패:', e);
    }

    // O73 셀
    try {
        const cell_O73 = worksheet.getCell('O73');
        cell_O73.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_O73.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_O73.border = { top: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 O73 설정 실패:', e);
    }

    // O8 셀
    try {
        const cell_O8 = worksheet.getCell('O8');
        cell_O8.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_O8.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_O8.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 O8 설정 실패:', e);
    }

    // O83 셀
    try {
        const cell_O83 = worksheet.getCell('O83');
        cell_O83.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_O83.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_O83.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 O83 설정 실패:', e);
    }

    // O9 셀
    try {
        const cell_O9 = worksheet.getCell('O9');
        cell_O9.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_O9.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_O9.border = { top: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 O9 설정 실패:', e);
    }

    // O91 셀
    try {
        const cell_O91 = worksheet.getCell('O91');
        cell_O91.font = { name: 'LG스마트체 Regular', size: 6.0, bold: true, color: { argb: 'FF000000' } };
        cell_O91.alignment = { horizontal: 'center', vertical: 'center' };
        cell_O91.numFmt = '_-* #,##0_-;\-* #,##0_-;_-* "-"_-;_-@_-';
    } catch (e) {
        console.warn('셀 O91 설정 실패:', e);
    }

    // P10 셀
    try {
        const cell_P10 = worksheet.getCell('P10');
        cell_P10.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P10.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P10.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P10 설정 실패:', e);
    }

    // P11 셀
    try {
        const cell_P11 = worksheet.getCell('P11');
        cell_P11.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P11.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P11.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P11 설정 실패:', e);
    }

    // P12 셀
    try {
        const cell_P12 = worksheet.getCell('P12');
        cell_P12.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P12.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P12.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P12 설정 실패:', e);
    }

    // P13 셀
    try {
        const cell_P13 = worksheet.getCell('P13');
        cell_P13.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P13.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P13.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P13 설정 실패:', e);
    }

    // P14 셀
    try {
        const cell_P14 = worksheet.getCell('P14');
        cell_P14.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P14.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P14.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P14 설정 실패:', e);
    }

    // P15 셀
    try {
        const cell_P15 = worksheet.getCell('P15');
        cell_P15.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P15.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P15.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P15 설정 실패:', e);
    }

    // P16 셀
    try {
        const cell_P16 = worksheet.getCell('P16');
        cell_P16.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P16.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P16.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P16 설정 실패:', e);
    }

    // P17 셀
    try {
        const cell_P17 = worksheet.getCell('P17');
        cell_P17.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P17.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P17.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P17 설정 실패:', e);
    }

    // P18 셀
    try {
        const cell_P18 = worksheet.getCell('P18');
        cell_P18.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P18.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P18.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P18 설정 실패:', e);
    }

    // P19 셀
    try {
        const cell_P19 = worksheet.getCell('P19');
        cell_P19.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P19.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P19.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P19 설정 실패:', e);
    }

    // P20 셀
    try {
        const cell_P20 = worksheet.getCell('P20');
        cell_P20.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P20.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P20.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P20 설정 실패:', e);
    }

    // P21 셀
    try {
        const cell_P21 = worksheet.getCell('P21');
        cell_P21.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P21.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P21.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P21 설정 실패:', e);
    }

    // P22 셀
    try {
        const cell_P22 = worksheet.getCell('P22');
        cell_P22.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P22.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P22.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P22 설정 실패:', e);
    }

    // P23 셀
    try {
        const cell_P23 = worksheet.getCell('P23');
        cell_P23.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P23.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P23.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P23 설정 실패:', e);
    }

    // P24 셀
    try {
        const cell_P24 = worksheet.getCell('P24');
        cell_P24.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P24.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P24.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P24 설정 실패:', e);
    }

    // P25 셀
    try {
        const cell_P25 = worksheet.getCell('P25');
        cell_P25.value = '12604';
        cell_P25.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_P25.alignment = { horizontal: 'center', vertical: 'center' };
        cell_P25.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_P25.numFmt = '"("#,##0.0\ "㎡)"';
    } catch (e) {
        console.warn('셀 P25 설정 실패:', e);
    }

    // P26 셀
    try {
        const cell_P26 = worksheet.getCell('P26');
        cell_P26.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P26.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P26.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P26 설정 실패:', e);
    }

    // P27 셀
    try {
        const cell_P27 = worksheet.getCell('P27');
        cell_P27.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P27.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P27.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P27 설정 실패:', e);
    }

    // P28 셀
    try {
        const cell_P28 = worksheet.getCell('P28');
        cell_P28.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_P28.alignment = { horizontal: 'center', vertical: 'center' };
        cell_P28.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_P28.numFmt = '#,##0.000\ "평"';
    } catch (e) {
        console.warn('셀 P28 설정 실패:', e);
    }

    // P29 셀
    try {
        const cell_P29 = worksheet.getCell('P29');
        cell_P29.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P29.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P29.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P29 설정 실패:', e);
    }

    // P30 셀
    try {
        const cell_P30 = worksheet.getCell('P30');
        cell_P30.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P30.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P30.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P30 설정 실패:', e);
    }

    // P31 셀
    try {
        const cell_P31 = worksheet.getCell('P31');
        cell_P31.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P31.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P31.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P31 설정 실패:', e);
    }

    // P32 셀
    try {
        const cell_P32 = worksheet.getCell('P32');
        cell_P32.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P32.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P32.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P32 설정 실패:', e);
    }

    // P33 셀
    try {
        const cell_P33 = worksheet.getCell('P33');
        cell_P33.value = '임대';
        cell_P33.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_P33.alignment = { horizontal: 'center', vertical: 'center' };
        cell_P33.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_P33.numFmt = '@';
    } catch (e) {
        console.warn('셀 P33 설정 실패:', e);
    }

    // P34 셀
    try {
        const cell_P34 = worksheet.getCell('P34');
        cell_P34.value = '229.8';
        cell_P34.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_P34.alignment = { horizontal: 'center', vertical: 'center' };
        cell_P34.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_P34.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 P34 설정 실패:', e);
    }

    // P35 셀
    try {
        const cell_P35 = worksheet.getCell('P35');
        cell_P35.value = '473';
        cell_P35.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_P35.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEECEC' } };
        cell_P35.alignment = { horizontal: 'center', vertical: 'center' };
        cell_P35.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_P35.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 P35 설정 실패:', e);
    }

    // P36 셀
    try {
        const cell_P36 = worksheet.getCell('P36');
        cell_P36.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_P36.alignment = { horizontal: 'center', vertical: 'center' };
        cell_P36.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_P36.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 P36 설정 실패:', e);
    }

    // P37 셀
    try {
        const cell_P37 = worksheet.getCell('P37');
        cell_P37.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_P37.alignment = { horizontal: 'center', vertical: 'center' };
        cell_P37.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_P37.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 P37 설정 실패:', e);
    }

    // P38 셀
    try {
        const cell_P38 = worksheet.getCell('P38');
        cell_P38.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_P38.alignment = { horizontal: 'center', vertical: 'center' };
        cell_P38.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_P38.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 P38 설정 실패:', e);
    }

    // P39 셀
    try {
        const cell_P39 = worksheet.getCell('P39');
        cell_P39.value = { formula: '=SUM(P34:P38)' };
        cell_P39.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_P39.alignment = { horizontal: 'center', vertical: 'center' };
        cell_P39.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_P39.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 P39 설정 실패:', e);
    }

    // P40 셀
    try {
        const cell_P40 = worksheet.getCell('P40');
        cell_P40.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P40.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P40.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P40 설정 실패:', e);
    }

    // P41 셀
    try {
        const cell_P41 = worksheet.getCell('P41');
        cell_P41.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P41.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P41.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P41 설정 실패:', e);
    }

    // P42 셀
    try {
        const cell_P42 = worksheet.getCell('P42');
        cell_P42.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P42.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P42.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P42 설정 실패:', e);
    }

    // P43 셀
    try {
        const cell_P43 = worksheet.getCell('P43');
        cell_P43.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P43.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P43.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P43 설정 실패:', e);
    }

    // P44 셀
    try {
        const cell_P44 = worksheet.getCell('P44');
        cell_P44.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P44.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P44.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P44 설정 실패:', e);
    }

    // P45 셀
    try {
        const cell_P45 = worksheet.getCell('P45');
        cell_P45.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P45.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P45.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P45 설정 실패:', e);
    }

    // P46 셀
    try {
        const cell_P46 = worksheet.getCell('P46');
        cell_P46.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P46.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P46.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P46 설정 실패:', e);
    }

    // P47 셀
    try {
        const cell_P47 = worksheet.getCell('P47');
        cell_P47.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P47.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P47.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P47 설정 실패:', e);
    }

    // P48 셀
    try {
        const cell_P48 = worksheet.getCell('P48');
        cell_P48.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P48.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P48.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P48 설정 실패:', e);
    }

    // P49 셀
    try {
        const cell_P49 = worksheet.getCell('P49');
        cell_P49.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P49.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P49.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P49 설정 실패:', e);
    }

    // P50 셀
    try {
        const cell_P50 = worksheet.getCell('P50');
        cell_P50.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P50.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P50.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P50 설정 실패:', e);
    }

    // P51 셀
    try {
        const cell_P51 = worksheet.getCell('P51');
        cell_P51.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P51.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P51.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P51 설정 실패:', e);
    }

    // P52 셀
    try {
        const cell_P52 = worksheet.getCell('P52');
        cell_P52.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P52.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P52.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P52 설정 실패:', e);
    }

    // P53 셀
    try {
        const cell_P53 = worksheet.getCell('P53');
        cell_P53.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P53.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P53.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P53 설정 실패:', e);
    }

    // P54 셀
    try {
        const cell_P54 = worksheet.getCell('P54');
        cell_P54.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P54.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P54.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P54 설정 실패:', e);
    }

    // P55 셀
    try {
        const cell_P55 = worksheet.getCell('P55');
        cell_P55.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P55.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P55.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P55 설정 실패:', e);
    }

    // P56 셀
    try {
        const cell_P56 = worksheet.getCell('P56');
        cell_P56.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P56.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P56.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P56 설정 실패:', e);
    }

    // P57 셀
    try {
        const cell_P57 = worksheet.getCell('P57');
        cell_P57.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P57.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P57.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P57 설정 실패:', e);
    }

    // P58 셀
    try {
        const cell_P58 = worksheet.getCell('P58');
        cell_P58.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P58.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P58.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P58 설정 실패:', e);
    }

    // P59 셀
    try {
        const cell_P59 = worksheet.getCell('P59');
        cell_P59.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P59.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P59.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P59 설정 실패:', e);
    }

    // P6 셀
    try {
        const cell_P6 = worksheet.getCell('P6');
        cell_P6.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P6.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P6.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P6 설정 실패:', e);
    }

    // P60 셀
    try {
        const cell_P60 = worksheet.getCell('P60');
        cell_P60.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P60.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P60.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P60 설정 실패:', e);
    }

    // P61 셀
    try {
        const cell_P61 = worksheet.getCell('P61');
        cell_P61.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P61.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P61.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P61 설정 실패:', e);
    }

    // P62 셀
    try {
        const cell_P62 = worksheet.getCell('P62');
        cell_P62.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P62.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P62.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P62 설정 실패:', e);
    }

    // P63 셀
    try {
        const cell_P63 = worksheet.getCell('P63');
        cell_P63.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P63.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P63.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P63 설정 실패:', e);
    }

    // P64 셀
    try {
        const cell_P64 = worksheet.getCell('P64');
        cell_P64.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P64.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P64.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P64 설정 실패:', e);
    }

    // P65 셀
    try {
        const cell_P65 = worksheet.getCell('P65');
        cell_P65.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P65.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P65.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P65 설정 실패:', e);
    }

    // P66 셀
    try {
        const cell_P66 = worksheet.getCell('P66');
        cell_P66.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P66.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P66.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P66 설정 실패:', e);
    }

    // P67 셀
    try {
        const cell_P67 = worksheet.getCell('P67');
        cell_P67.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P67.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P67.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P67 설정 실패:', e);
    }

    // P68 셀
    try {
        const cell_P68 = worksheet.getCell('P68');
        cell_P68.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P68.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P68.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P68 설정 실패:', e);
    }

    // P69 셀
    try {
        const cell_P69 = worksheet.getCell('P69');
        cell_P69.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P69.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P69.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P69 설정 실패:', e);
    }

    // P7 셀
    try {
        const cell_P7 = worksheet.getCell('P7');
        cell_P7.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P7.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P7.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P7 설정 실패:', e);
    }

    // P70 셀
    try {
        const cell_P70 = worksheet.getCell('P70');
        cell_P70.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P70.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P70.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P70 설정 실패:', e);
    }

    // P71 셀
    try {
        const cell_P71 = worksheet.getCell('P71');
        cell_P71.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P71.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P71.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P71 설정 실패:', e);
    }

    // P72 셀
    try {
        const cell_P72 = worksheet.getCell('P72');
        cell_P72.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P72.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P72.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P72 설정 실패:', e);
    }

    // P73 셀
    try {
        const cell_P73 = worksheet.getCell('P73');
        cell_P73.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P73.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P73.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P73 설정 실패:', e);
    }

    // P74 셀
    try {
        const cell_P74 = worksheet.getCell('P74');
        cell_P74.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P74.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P74.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P74 설정 실패:', e);
    }

    // P75 셀
    try {
        const cell_P75 = worksheet.getCell('P75');
        cell_P75.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P75.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P75.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P75 설정 실패:', e);
    }

    // P76 셀
    try {
        const cell_P76 = worksheet.getCell('P76');
        cell_P76.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P76.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P76.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P76 설정 실패:', e);
    }

    // P77 셀
    try {
        const cell_P77 = worksheet.getCell('P77');
        cell_P77.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P77.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P77.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P77 설정 실패:', e);
    }

    // P78 셀
    try {
        const cell_P78 = worksheet.getCell('P78');
        cell_P78.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P78.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P78.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P78 설정 실패:', e);
    }

    // P79 셀
    try {
        const cell_P79 = worksheet.getCell('P79');
        cell_P79.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P79.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P79.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P79 설정 실패:', e);
    }

    // P8 셀
    try {
        const cell_P8 = worksheet.getCell('P8');
        cell_P8.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P8.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P8.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P8 설정 실패:', e);
    }

    // P80 셀
    try {
        const cell_P80 = worksheet.getCell('P80');
        cell_P80.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P80.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P80.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P80 설정 실패:', e);
    }

    // P81 셀
    try {
        const cell_P81 = worksheet.getCell('P81');
        cell_P81.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P81.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P81.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P81 설정 실패:', e);
    }

    // P82 셀
    try {
        const cell_P82 = worksheet.getCell('P82');
        cell_P82.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P82.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P82.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P82 설정 실패:', e);
    }

    // P83 셀
    try {
        const cell_P83 = worksheet.getCell('P83');
        cell_P83.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P83.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P83.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P83 설정 실패:', e);
    }

    // P9 셀
    try {
        const cell_P9 = worksheet.getCell('P9');
        cell_P9.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_P9.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_P9.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 P9 설정 실패:', e);
    }

    // Q106 셀
    try {
        const cell_Q106 = worksheet.getCell('Q106');
        cell_Q106.font = { name: 'LG스마트체 Regular', size: 10.0, bold: true, color: { argb: 'FF000000' } };
        cell_Q106.alignment = { horizontal: 'center', vertical: 'center' };
        cell_Q106.numFmt = '_-* #,##0_-;\-* #,##0_-;_-* "-"_-;_-@_-';
    } catch (e) {
        console.warn('셀 Q106 설정 실패:', e);
    }

    // Q17 셀
    try {
        const cell_Q17 = worksheet.getCell('Q17');
        cell_Q17.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_Q17.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_Q17.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 Q17 설정 실패:', e);
    }

    // Q18 셀
    try {
        const cell_Q18 = worksheet.getCell('Q18');
        cell_Q18.value = '서울시 금천구 가산디지털2로 30';
        cell_Q18.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_Q18.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        cell_Q18.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_Q18.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 Q18 설정 실패:', e);
    }

    // Q19 셀
    try {
        const cell_Q19 = worksheet.getCell('Q19');
        cell_Q19.value = '1호선 독산역 도보 12분\n1,7호선 가산디지털단지역 도보 12분';
        cell_Q19.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_Q19.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        cell_Q19.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_Q19.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 Q19 설정 실패:', e);
    }

    // Q20 셀
    try {
        const cell_Q20 = worksheet.getCell('Q20');
        cell_Q20.value = '2012';
        cell_Q20.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_Q20.alignment = { horizontal: 'center', vertical: 'center' };
        cell_Q20.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_Q20.numFmt = '###0"년"';
    } catch (e) {
        console.warn('셀 Q20 설정 실패:', e);
    }

    // Q21 셀
    try {
        const cell_Q21 = worksheet.getCell('Q21');
        cell_Q21.value = '133';
        cell_Q21.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_Q21.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        cell_Q21.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_Q21.numFmt = '##"F / B"#';
    } catch (e) {
        console.warn('셀 Q21 설정 실패:', e);
    }

    // Q22 셀
    try {
        const cell_Q22 = worksheet.getCell('Q22');
        cell_Q22.value = '9845.52';
        cell_Q22.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_Q22.alignment = { horizontal: 'center', vertical: 'center' };
        cell_Q22.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_Q22.numFmt = '#,##0\ "평"';
    } catch (e) {
        console.warn('셀 Q22 설정 실패:', e);
    }

    // Q23 셀
    try {
        const cell_Q23 = worksheet.getCell('Q23');
        cell_Q23.value = '303';
        cell_Q23.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_Q23.alignment = { horizontal: 'center', vertical: 'center' };
        cell_Q23.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_Q23.numFmt = '#,##0\ "평"';
    } catch (e) {
        console.warn('셀 Q23 설정 실패:', e);
    }

    // Q24 셀
    try {
        const cell_Q24 = worksheet.getCell('Q24');
        cell_Q24.value = { formula: '=279/469' };
        cell_Q24.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_Q24.alignment = { horizontal: 'center', vertical: 'center' };
        cell_Q24.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_Q24.numFmt = '0.00%';
    } catch (e) {
        console.warn('셀 Q24 설정 실패:', e);
    }

    // Q25 셀
    try {
        const cell_Q25 = worksheet.getCell('Q25');
        cell_Q25.value = '268.88';
        cell_Q25.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_Q25.alignment = { horizontal: 'center', vertical: 'center' };
        cell_Q25.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_Q25.numFmt = '#,##0\ "평"';
    } catch (e) {
        console.warn('셀 Q25 설정 실패:', e);
    }

    // Q26 셀
    try {
        const cell_Q26 = worksheet.getCell('Q26');
        cell_Q26.value = '㈜동암씨티';
        cell_Q26.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_Q26.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        cell_Q26.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_Q26.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 Q26 설정 실패:', e);
    }

    // Q27 셀
    try {
        const cell_Q27 = worksheet.getCell('Q27');
        cell_Q27.value = '질권 설정 가능';
        cell_Q27.font = { name: 'LG스마트체 Regular', size: 9.0, bold: true, color: { argb: 'FFC00000' } };
        cell_Q27.alignment = { horizontal: 'center', vertical: 'center' };
        cell_Q27.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_Q27.numFmt = '0.00%';
    } catch (e) {
        console.warn('셀 Q27 설정 실패:', e);
    }

    // Q28 셀
    try {
        const cell_Q28 = worksheet.getCell('Q28');
        cell_Q28.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_Q28.alignment = { horizontal: 'center', vertical: 'center' };
        cell_Q28.border = { top: { style: 'hair', color: { argb: 'FF000000' } } };
        cell_Q28.numFmt = '#,##0.000\ "평"';
    } catch (e) {
        console.warn('셀 Q28 설정 실패:', e);
    }

    // Q29 셀
    try {
        const cell_Q29 = worksheet.getCell('Q29');
        cell_Q29.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_Q29.alignment = { horizontal: 'center', vertical: 'center' };
        cell_Q29.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_Q29.numFmt = '_-* #,##0_-;\-* #,##0_-;_-* "-"_-;_-@_-';
    } catch (e) {
        console.warn('셀 Q29 설정 실패:', e);
    }

    // Q30 셀
    try {
        const cell_Q30 = worksheet.getCell('Q30');
        cell_Q30.value = { formula: '=Q29/Q32' };
        cell_Q30.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FFC00000' } };
        cell_Q30.alignment = { horizontal: 'center', vertical: 'center' };
        cell_Q30.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_Q30.numFmt = '0.00%';
    } catch (e) {
        console.warn('셀 Q30 설정 실패:', e);
    }

    // Q31 셀
    try {
        const cell_Q31 = worksheet.getCell('Q31');
        cell_Q31.value = '4246000';
        cell_Q31.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_Q31.alignment = { horizontal: 'center', vertical: 'center' };
        cell_Q31.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_Q31.numFmt = '#,###"원/㎡"';
    } catch (e) {
        console.warn('셀 Q31 설정 실패:', e);
    }

    // Q32 셀
    try {
        const cell_Q32 = worksheet.getCell('Q32');
        cell_Q32.value = { formula: '=Q31*S25' };
        cell_Q32.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_Q32.alignment = { horizontal: 'center', vertical: 'center' };
        cell_Q32.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_Q32.numFmt = '_-* #,##0_-;\-* #,##0_-;_-* "-"_-;_-@_-';
    } catch (e) {
        console.warn('셀 Q32 설정 실패:', e);
    }

    // Q33 셀
    try {
        const cell_Q33 = worksheet.getCell('Q33');
        cell_Q33.value = '층';
        cell_Q33.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_Q33.alignment = { horizontal: 'center', vertical: 'center' };
        cell_Q33.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_Q33.numFmt = '@';
    } catch (e) {
        console.warn('셀 Q33 설정 실패:', e);
    }

    // Q34 셀
    try {
        const cell_Q34 = worksheet.getCell('Q34');
        cell_Q34.value = '8층';
        cell_Q34.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF0000FF' } };
        cell_Q34.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEECEC' } };
        cell_Q34.alignment = { horizontal: 'center', vertical: 'center' };
        cell_Q34.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_Q34.numFmt = '#"층"';
    } catch (e) {
        console.warn('셀 Q34 설정 실패:', e);
    }

    // Q35 셀
    try {
        const cell_Q35 = worksheet.getCell('Q35');
        cell_Q35.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF0000FF' } };
        cell_Q35.alignment = { horizontal: 'center', vertical: 'center' };
        cell_Q35.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_Q35.numFmt = '#"층"';
    } catch (e) {
        console.warn('셀 Q35 설정 실패:', e);
    }

    // Q36 셀
    try {
        const cell_Q36 = worksheet.getCell('Q36');
        cell_Q36.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_Q36.alignment = { horizontal: 'center', vertical: 'center' };
        cell_Q36.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_Q36.numFmt = '#"층"';
    } catch (e) {
        console.warn('셀 Q36 설정 실패:', e);
    }

    // Q37 셀
    try {
        const cell_Q37 = worksheet.getCell('Q37');
        cell_Q37.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_Q37.alignment = { horizontal: 'center', vertical: 'center' };
        cell_Q37.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_Q37.numFmt = '#"층"';
    } catch (e) {
        console.warn('셀 Q37 설정 실패:', e);
    }

    // Q38 셀
    try {
        const cell_Q38 = worksheet.getCell('Q38');
        cell_Q38.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_Q38.alignment = { horizontal: 'center', vertical: 'center' };
        cell_Q38.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_Q38.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 Q38 설정 실패:', e);
    }

    // Q39 셀
    try {
        const cell_Q39 = worksheet.getCell('Q39');
        cell_Q39.value = '소계';
        cell_Q39.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_Q39.alignment = { horizontal: 'center', vertical: 'center' };
        cell_Q39.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_Q39.numFmt = '@';
    } catch (e) {
        console.warn('셀 Q39 설정 실패:', e);
    }

    // Q40 셀
    try {
        const cell_Q40 = worksheet.getCell('Q40');
        cell_Q40.value = '2025.7~2027.6 (12개월)';
        cell_Q40.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_Q40.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        cell_Q40.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_Q40.numFmt = '#"개월 계약 가능"';
    } catch (e) {
        console.warn('셀 Q40 설정 실패:', e);
    }

    // Q41 셀
    try {
        const cell_Q41 = worksheet.getCell('Q41');
        cell_Q41.value = '즉시';
        cell_Q41.font = { name: 'LG스마트체 Regular', size: 9.0, bold: true, color: { argb: 'FF000000' } };
        cell_Q41.alignment = { horizontal: 'center', vertical: 'center' };
        cell_Q41.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_Q41.numFmt = '#,##0\ "원"';
    } catch (e) {
        console.warn('셀 Q41 설정 실패:', e);
    }

    // Q42 셀
    try {
        const cell_Q42 = worksheet.getCell('Q42');
        cell_Q42.value = '8층 전체';
        cell_Q42.font = { name: 'LG스마트체 Regular', size: 9.0, bold: true, color: { argb: 'FF000000' } };
        cell_Q42.alignment = { horizontal: 'center', vertical: 'center' };
        cell_Q42.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_Q42.numFmt = '#,##0\ "층"';
    } catch (e) {
        console.warn('셀 Q42 설정 실패:', e);
    }

    // Q43 셀
    try {
        const cell_Q43 = worksheet.getCell('Q43');
        cell_Q43.value = { formula: '=R34' };
        cell_Q43.font = { name: 'LG스마트체 Regular', size: 9.0, bold: true, color: { argb: 'FFC00000' } };
        cell_Q43.alignment = { horizontal: 'center', vertical: 'center' };
        cell_Q43.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_Q43.numFmt = '#,##0\ "평"';
    } catch (e) {
        console.warn('셀 Q43 설정 실패:', e);
    }

    // Q44 셀
    try {
        const cell_Q44 = worksheet.getCell('Q44');
        cell_Q44.value = { formula: '=S34' };
        cell_Q44.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_Q44.alignment = { horizontal: 'center', vertical: 'center' };
        cell_Q44.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_Q44.numFmt = '#,##0\ "평"';
    } catch (e) {
        console.warn('셀 Q44 설정 실패:', e);
    }

    // Q45 셀
    try {
        const cell_Q45 = worksheet.getCell('Q45');
        cell_Q45.value = '350000';
        cell_Q45.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_Q45.alignment = { horizontal: 'center', vertical: 'center' };
        cell_Q45.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_Q45.numFmt = '"@"#,###';
    } catch (e) {
        console.warn('셀 Q45 설정 실패:', e);
    }

    // Q46 셀
    try {
        const cell_Q46 = worksheet.getCell('Q46');
        cell_Q46.value = '35000';
        cell_Q46.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_Q46.alignment = { horizontal: 'center', vertical: 'center' };
        cell_Q46.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_Q46.numFmt = '"@"#,###';
    } catch (e) {
        console.warn('셀 Q46 설정 실패:', e);
    }

    // Q47 셀
    try {
        const cell_Q47 = worksheet.getCell('Q47');
        cell_Q47.value = '10000';
        cell_Q47.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_Q47.alignment = { horizontal: 'center', vertical: 'center' };
        cell_Q47.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_Q47.numFmt = '"@"#,###\+"실""비""별""도"';
    } catch (e) {
        console.warn('셀 Q47 설정 실패:', e);
    }

    // Q48 셀
    try {
        const cell_Q48 = worksheet.getCell('Q48');
        cell_Q48.value = { formula: '=Q46*(12-Q49)/12' };
        cell_Q48.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_Q48.alignment = { horizontal: 'center', vertical: 'center' };
        cell_Q48.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_Q48.numFmt = '"@"#,###';
    } catch (e) {
        console.warn('셀 Q48 설정 실패:', e);
    }

    // Q49 셀
    try {
        const cell_Q49 = worksheet.getCell('Q49');
        cell_Q49.value = '1';
        cell_Q49.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_Q49.alignment = { horizontal: 'center', vertical: 'center' };
        cell_Q49.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_Q49.numFmt = '#0.0"개월"';
    } catch (e) {
        console.warn('셀 Q49 설정 실패:', e);
    }

    // Q50 셀
    try {
        const cell_Q50 = worksheet.getCell('Q50');
        cell_Q50.value = { formula: '=Q45*Q44' };
        cell_Q50.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_Q50.alignment = { horizontal: 'center', vertical: 'center' };
        cell_Q50.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_Q50.numFmt = '#,##0\ "원"';
    } catch (e) {
        console.warn('셀 Q50 설정 실패:', e);
    }

    // Q51 셀
    try {
        const cell_Q51 = worksheet.getCell('Q51');
        cell_Q51.value = { formula: '=Q46*Q44' };
        cell_Q51.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_Q51.alignment = { horizontal: 'center', vertical: 'center' };
        cell_Q51.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_Q51.numFmt = '#,##0\ "원"';
    } catch (e) {
        console.warn('셀 Q51 설정 실패:', e);
    }

    // Q52 셀
    try {
        const cell_Q52 = worksheet.getCell('Q52');
        cell_Q52.value = { formula: '=Q47*Q44' };
        cell_Q52.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_Q52.alignment = { horizontal: 'center', vertical: 'center' };
        cell_Q52.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_Q52.numFmt = '#,##0\ "원"';
    } catch (e) {
        console.warn('셀 Q52 설정 실패:', e);
    }

    // Q53 셀
    try {
        const cell_Q53 = worksheet.getCell('Q53');
        cell_Q53.value = '실비 관리비: 전기세, 수도세  별도 부과';
        cell_Q53.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FFC00000' } };
        cell_Q53.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        cell_Q53.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_Q53.numFmt = '#,##0\ "원"';
    } catch (e) {
        console.warn('셀 Q53 설정 실패:', e);
    }

    // Q54 셀
    try {
        const cell_Q54 = worksheet.getCell('Q54');
        cell_Q54.value = { formula: '=Q51+Q52' };
        cell_Q54.font = { name: 'LG스마트체 Regular', size: 9.0, bold: true, color: { argb: 'FFC00000' } };
        cell_Q54.alignment = { horizontal: 'center', vertical: 'center' };
        cell_Q54.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_Q54.numFmt = '#,##0\ "원"';
    } catch (e) {
        console.warn('셀 Q54 설정 실패:', e);
    }

    // Q55 셀
    try {
        const cell_Q55 = worksheet.getCell('Q55');
        cell_Q55.value = { formula: '=Q54*21' };
        cell_Q55.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_Q55.alignment = { horizontal: 'center', vertical: 'center' };
        cell_Q55.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_Q55.numFmt = '#,##0\ "원"';
    } catch (e) {
        console.warn('셀 Q55 설정 실패:', e);
    }

    // Q56 셀
    try {
        const cell_Q56 = worksheet.getCell('Q56');
        cell_Q56.value = '1';
        cell_Q56.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_Q56.alignment = { horizontal: 'center', vertical: 'center' };
        cell_Q56.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_Q56.numFmt = '0.#"개월"';
    } catch (e) {
        console.warn('셀 Q56 설정 실패:', e);
    }

    // Q57 셀
    try {
        const cell_Q57 = worksheet.getCell('Q57');
        cell_Q57.value = '미제공';
        cell_Q57.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_Q57.alignment = { horizontal: 'center', vertical: 'center' };
        cell_Q57.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_Q57.numFmt = '"총액 "##,##0"원"';
    } catch (e) {
        console.warn('셀 Q57 설정 실패:', e);
    }

    // Q58 셀
    try {
        const cell_Q58 = worksheet.getCell('Q58');
        cell_Q58.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_Q58.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_Q58.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 Q58 설정 실패:', e);
    }

    // Q59 셀
    try {
        const cell_Q59 = worksheet.getCell('Q59');
        cell_Q59.value = '232';
        cell_Q59.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_Q59.alignment = { horizontal: 'center', vertical: 'center' };
        cell_Q59.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_Q59.numFmt = '#\ "대"';
    } catch (e) {
        console.warn('셀 Q59 설정 실패:', e);
    }

    // Q6 셀
    try {
        const cell_Q6 = worksheet.getCell('Q6');
        cell_Q6.value = '가산디지털단지역';
        cell_Q6.font = { name: 'LG스마트체 Regular', size: 10.0, color: { argb: 'FF000000' } };
        cell_Q6.alignment = { horizontal: 'center', vertical: 'center' };
        cell_Q6.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_Q6.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 Q6 설정 실패:', e);
    }

    // Q60 셀
    try {
        const cell_Q60 = worksheet.getCell('Q60');
        cell_Q60.value = '40';
        cell_Q60.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_Q60.alignment = { horizontal: 'center', vertical: 'center' };
        cell_Q60.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_Q60.numFmt = '"임대면적"\ #"평당 1대"';
    } catch (e) {
        console.warn('셀 Q60 설정 실패:', e);
    }

    // Q61 셀
    try {
        const cell_Q61 = worksheet.getCell('Q61');
        cell_Q61.value = { formula: '=Q44/Q60' };
        cell_Q61.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_Q61.alignment = { horizontal: 'center', vertical: 'center' };
        cell_Q61.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_Q61.numFmt = '#,##0.0\ "대"';
    } catch (e) {
        console.warn('셀 Q61 설정 실패:', e);
    }

    // Q62 셀
    try {
        const cell_Q62 = worksheet.getCell('Q62');
        cell_Q62.value = '협의';
        cell_Q62.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_Q62.alignment = { horizontal: 'center', vertical: 'center' };
        cell_Q62.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_Q62.numFmt = '"월"#"만원/대"';
    } catch (e) {
        console.warn('셀 Q62 설정 실패:', e);
    }

    // Q63 셀
    try {
        const cell_Q63 = worksheet.getCell('Q63');
        cell_Q63.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_Q63.alignment = { horizontal: 'left', vertical: 'center', wrapText: true };
        cell_Q63.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_Q63.numFmt = '#,##0_ ';
    } catch (e) {
        console.warn('셀 Q63 설정 실패:', e);
    }

    // Q7 셀
    try {
        const cell_Q7 = worksheet.getCell('Q7');
        cell_Q7.font = { name: 'LG스마트체 Regular', size: 10.0, color: { argb: 'FF000000' } };
        cell_Q7.alignment = { horizontal: 'center', vertical: 'center' };
        cell_Q7.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_Q7.numFmt = '0_);[Red]\(0\)';
    } catch (e) {
        console.warn('셀 Q7 설정 실패:', e);
    }

    // Q72 셀
    try {
        const cell_Q72 = worksheet.getCell('Q72');
        cell_Q72.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_Q72.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_Q72.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 Q72 설정 실패:', e);
    }

    // Q73 셀
    try {
        const cell_Q73 = worksheet.getCell('Q73');
        cell_Q73.value = ' - 현재 8층 즉시 가능 (임대인 직접 계약)\n - Rent free 1개월 제공\n - 인테리어공사기간 1개월 제공\n - 실비 별도\n - 8층 서버실 전용 전력선 인입되어 있음\n - 매일 청소서비스 제공\n - 사무실 내부 조명기구, 냉난방 등 설비 무상수리\n - 임대인 단독 소유 및 직접 관리/운영\n - 5층 구내식당 및 피트니스 센터 운영';
        cell_Q73.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_Q73.alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
        cell_Q73.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_Q73.numFmt = '#,##0\ "대"';
    } catch (e) {
        console.warn('셀 Q73 설정 실패:', e);
    }

    // Q8 셀
    try {
        const cell_Q8 = worksheet.getCell('Q8');
        cell_Q8.value = 'RSM타워';
        cell_Q8.font = { name: 'LG스마트체 Regular', size: 10.0, color: { argb: 'FF000000' } };
        cell_Q8.alignment = { horizontal: 'center', vertical: 'center' };
        cell_Q8.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_Q8.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 Q8 설정 실패:', e);
    }

    // Q83 셀
    try {
        const cell_Q83 = worksheet.getCell('Q83');
        cell_Q83.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_Q83.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_Q83.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 Q83 설정 실패:', e);
    }

    // Q9 셀
    try {
        const cell_Q9 = worksheet.getCell('Q9');
        cell_Q9.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_Q9.alignment = { horizontal: 'center', vertical: 'center' };
        cell_Q9.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_Q9.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 Q9 설정 실패:', e);
    }

    // R106 셀
    try {
        const cell_R106 = worksheet.getCell('R106');
        cell_R106.font = { name: 'LG스마트체 Regular', size: 10.0, bold: true, color: { argb: 'FF000000' } };
        cell_R106.alignment = { horizontal: 'center', vertical: 'center' };
        cell_R106.numFmt = '_-* #,##0_-;\-* #,##0_-;_-* "-"_-;_-@_-';
    } catch (e) {
        console.warn('셀 R106 설정 실패:', e);
    }

    // R17 셀
    try {
        const cell_R17 = worksheet.getCell('R17');
        cell_R17.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_R17.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_R17.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 R17 설정 실패:', e);
    }

    // R18 셀
    try {
        const cell_R18 = worksheet.getCell('R18');
        cell_R18.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_R18.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_R18.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 R18 설정 실패:', e);
    }

    // R19 셀
    try {
        const cell_R19 = worksheet.getCell('R19');
        cell_R19.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_R19.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_R19.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 R19 설정 실패:', e);
    }

    // R20 셀
    try {
        const cell_R20 = worksheet.getCell('R20');
        cell_R20.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_R20.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_R20.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 R20 설정 실패:', e);
    }

    // R21 셀
    try {
        const cell_R21 = worksheet.getCell('R21');
        cell_R21.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_R21.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_R21.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 R21 설정 실패:', e);
    }

    // R22 셀
    try {
        const cell_R22 = worksheet.getCell('R22');
        cell_R22.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_R22.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_R22.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 R22 설정 실패:', e);
    }

    // R23 셀
    try {
        const cell_R23 = worksheet.getCell('R23');
        cell_R23.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_R23.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_R23.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 R23 설정 실패:', e);
    }

    // R24 셀
    try {
        const cell_R24 = worksheet.getCell('R24');
        cell_R24.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_R24.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_R24.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 R24 설정 실패:', e);
    }

    // R25 셀
    try {
        const cell_R25 = worksheet.getCell('R25');
        cell_R25.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_R25.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_R25.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 R25 설정 실패:', e);
    }

    // R26 셀
    try {
        const cell_R26 = worksheet.getCell('R26');
        cell_R26.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_R26.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_R26.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 R26 설정 실패:', e);
    }

    // R27 셀
    try {
        const cell_R27 = worksheet.getCell('R27');
        cell_R27.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_R27.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_R27.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 R27 설정 실패:', e);
    }

    // R28 셀
    try {
        const cell_R28 = worksheet.getCell('R28');
        cell_R28.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_R28.alignment = { horizontal: 'center', vertical: 'center' };
        cell_R28.border = { top: { style: 'hair', color: { argb: 'FF000000' } } };
        cell_R28.numFmt = '#,##0.000\ "평"';
    } catch (e) {
        console.warn('셀 R28 설정 실패:', e);
    }

    // R29 셀
    try {
        const cell_R29 = worksheet.getCell('R29');
        cell_R29.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_R29.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_R29.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 R29 설정 실패:', e);
    }

    // R30 셀
    try {
        const cell_R30 = worksheet.getCell('R30');
        cell_R30.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_R30.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_R30.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 R30 설정 실패:', e);
    }

    // R31 셀
    try {
        const cell_R31 = worksheet.getCell('R31');
        cell_R31.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_R31.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_R31.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 R31 설정 실패:', e);
    }

    // R32 셀
    try {
        const cell_R32 = worksheet.getCell('R32');
        cell_R32.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_R32.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_R32.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 R32 설정 실패:', e);
    }

    // R33 셀
    try {
        const cell_R33 = worksheet.getCell('R33');
        cell_R33.value = '전용';
        cell_R33.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_R33.alignment = { horizontal: 'center', vertical: 'center' };
        cell_R33.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_R33.numFmt = '@';
    } catch (e) {
        console.warn('셀 R33 설정 실패:', e);
    }

    // R34 셀
    try {
        const cell_R34 = worksheet.getCell('R34');
        cell_R34.value = '279';
        cell_R34.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF0000FF' } };
        cell_R34.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEECEC' } };
        cell_R34.alignment = { horizontal: 'center', vertical: 'center' };
        cell_R34.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_R34.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 R34 설정 실패:', e);
    }

    // R35 셀
    try {
        const cell_R35 = worksheet.getCell('R35');
        cell_R35.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF0000FF' } };
        cell_R35.alignment = { horizontal: 'center', vertical: 'center' };
        cell_R35.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_R35.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 R35 설정 실패:', e);
    }

    // R36 셀
    try {
        const cell_R36 = worksheet.getCell('R36');
        cell_R36.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_R36.alignment = { horizontal: 'center', vertical: 'center' };
        cell_R36.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_R36.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 R36 설정 실패:', e);
    }

    // R37 셀
    try {
        const cell_R37 = worksheet.getCell('R37');
        cell_R37.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_R37.alignment = { horizontal: 'center', vertical: 'center' };
        cell_R37.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_R37.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 R37 설정 실패:', e);
    }

    // R38 셀
    try {
        const cell_R38 = worksheet.getCell('R38');
        cell_R38.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_R38.alignment = { horizontal: 'center', vertical: 'center' };
        cell_R38.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_R38.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 R38 설정 실패:', e);
    }

    // R39 셀
    try {
        const cell_R39 = worksheet.getCell('R39');
        cell_R39.value = { formula: '=SUM(R34:R38)' };
        cell_R39.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_R39.alignment = { horizontal: 'center', vertical: 'center' };
        cell_R39.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_R39.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 R39 설정 실패:', e);
    }

    // R40 셀
    try {
        const cell_R40 = worksheet.getCell('R40');
        cell_R40.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_R40.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_R40.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 R40 설정 실패:', e);
    }

    // R41 셀
    try {
        const cell_R41 = worksheet.getCell('R41');
        cell_R41.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_R41.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_R41.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 R41 설정 실패:', e);
    }

    // R42 셀
    try {
        const cell_R42 = worksheet.getCell('R42');
        cell_R42.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_R42.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_R42.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 R42 설정 실패:', e);
    }

    // R43 셀
    try {
        const cell_R43 = worksheet.getCell('R43');
        cell_R43.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_R43.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_R43.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 R43 설정 실패:', e);
    }

    // R44 셀
    try {
        const cell_R44 = worksheet.getCell('R44');
        cell_R44.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_R44.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_R44.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 R44 설정 실패:', e);
    }

    // R45 셀
    try {
        const cell_R45 = worksheet.getCell('R45');
        cell_R45.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_R45.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_R45.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 R45 설정 실패:', e);
    }

    // R46 셀
    try {
        const cell_R46 = worksheet.getCell('R46');
        cell_R46.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_R46.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_R46.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 R46 설정 실패:', e);
    }

    // R47 셀
    try {
        const cell_R47 = worksheet.getCell('R47');
        cell_R47.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_R47.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_R47.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 R47 설정 실패:', e);
    }

    // R48 셀
    try {
        const cell_R48 = worksheet.getCell('R48');
        cell_R48.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_R48.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_R48.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 R48 설정 실패:', e);
    }

    // R49 셀
    try {
        const cell_R49 = worksheet.getCell('R49');
        cell_R49.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_R49.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_R49.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 R49 설정 실패:', e);
    }

    // R50 셀
    try {
        const cell_R50 = worksheet.getCell('R50');
        cell_R50.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_R50.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_R50.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 R50 설정 실패:', e);
    }

    // R51 셀
    try {
        const cell_R51 = worksheet.getCell('R51');
        cell_R51.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_R51.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_R51.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 R51 설정 실패:', e);
    }

    // R52 셀
    try {
        const cell_R52 = worksheet.getCell('R52');
        cell_R52.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_R52.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_R52.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 R52 설정 실패:', e);
    }

    // R53 셀
    try {
        const cell_R53 = worksheet.getCell('R53');
        cell_R53.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_R53.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_R53.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 R53 설정 실패:', e);
    }

    // R54 셀
    try {
        const cell_R54 = worksheet.getCell('R54');
        cell_R54.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_R54.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_R54.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 R54 설정 실패:', e);
    }

    // R55 셀
    try {
        const cell_R55 = worksheet.getCell('R55');
        cell_R55.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_R55.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_R55.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 R55 설정 실패:', e);
    }

    // R56 셀
    try {
        const cell_R56 = worksheet.getCell('R56');
        cell_R56.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_R56.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_R56.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 R56 설정 실패:', e);
    }

    // R57 셀
    try {
        const cell_R57 = worksheet.getCell('R57');
        cell_R57.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_R57.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_R57.border = { top: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 R57 설정 실패:', e);
    }

    // R58 셀
    try {
        const cell_R58 = worksheet.getCell('R58');
        cell_R58.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_R58.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_R58.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 R58 설정 실패:', e);
    }

    // R59 셀
    try {
        const cell_R59 = worksheet.getCell('R59');
        cell_R59.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_R59.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_R59.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 R59 설정 실패:', e);
    }

    // R6 셀
    try {
        const cell_R6 = worksheet.getCell('R6');
        cell_R6.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_R6.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_R6.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 R6 설정 실패:', e);
    }

    // R60 셀
    try {
        const cell_R60 = worksheet.getCell('R60');
        cell_R60.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_R60.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_R60.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 R60 설정 실패:', e);
    }

    // R61 셀
    try {
        const cell_R61 = worksheet.getCell('R61');
        cell_R61.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_R61.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_R61.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 R61 설정 실패:', e);
    }

    // R62 셀
    try {
        const cell_R62 = worksheet.getCell('R62');
        cell_R62.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_R62.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_R62.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 R62 설정 실패:', e);
    }

    // R63 셀
    try {
        const cell_R63 = worksheet.getCell('R63');
        cell_R63.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_R63.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_R63.border = { top: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 R63 설정 실패:', e);
    }

    // R7 셀
    try {
        const cell_R7 = worksheet.getCell('R7');
        cell_R7.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_R7.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_R7.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 R7 설정 실패:', e);
    }

    // R72 셀
    try {
        const cell_R72 = worksheet.getCell('R72');
        cell_R72.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_R72.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_R72.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 R72 설정 실패:', e);
    }

    // R73 셀
    try {
        const cell_R73 = worksheet.getCell('R73');
        cell_R73.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_R73.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_R73.border = { top: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 R73 설정 실패:', e);
    }

    // R8 셀
    try {
        const cell_R8 = worksheet.getCell('R8');
        cell_R8.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_R8.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_R8.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 R8 설정 실패:', e);
    }

    // R83 셀
    try {
        const cell_R83 = worksheet.getCell('R83');
        cell_R83.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_R83.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_R83.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 R83 설정 실패:', e);
    }

    // R9 셀
    try {
        const cell_R9 = worksheet.getCell('R9');
        cell_R9.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_R9.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_R9.border = { top: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 R9 설정 실패:', e);
    }

    // S10 셀
    try {
        const cell_S10 = worksheet.getCell('S10');
        cell_S10.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S10.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S10.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S10 설정 실패:', e);
    }

    // S106 셀
    try {
        const cell_S106 = worksheet.getCell('S106');
        cell_S106.font = { name: 'LG스마트체 Regular', size: 10.0, bold: true, color: { argb: 'FF000000' } };
        cell_S106.alignment = { horizontal: 'center', vertical: 'center' };
        cell_S106.numFmt = '_-* #,##0_-;\-* #,##0_-;_-* "-"_-;_-@_-';
    } catch (e) {
        console.warn('셀 S106 설정 실패:', e);
    }

    // S11 셀
    try {
        const cell_S11 = worksheet.getCell('S11');
        cell_S11.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S11.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S11.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S11 설정 실패:', e);
    }

    // S12 셀
    try {
        const cell_S12 = worksheet.getCell('S12');
        cell_S12.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S12.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S12.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S12 설정 실패:', e);
    }

    // S13 셀
    try {
        const cell_S13 = worksheet.getCell('S13');
        cell_S13.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S13.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S13.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S13 설정 실패:', e);
    }

    // S14 셀
    try {
        const cell_S14 = worksheet.getCell('S14');
        cell_S14.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S14.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S14.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S14 설정 실패:', e);
    }

    // S15 셀
    try {
        const cell_S15 = worksheet.getCell('S15');
        cell_S15.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S15.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S15.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S15 설정 실패:', e);
    }

    // S16 셀
    try {
        const cell_S16 = worksheet.getCell('S16');
        cell_S16.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S16.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S16.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S16 설정 실패:', e);
    }

    // S17 셀
    try {
        const cell_S17 = worksheet.getCell('S17');
        cell_S17.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S17.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S17.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S17 설정 실패:', e);
    }

    // S18 셀
    try {
        const cell_S18 = worksheet.getCell('S18');
        cell_S18.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S18.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S18.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S18 설정 실패:', e);
    }

    // S19 셀
    try {
        const cell_S19 = worksheet.getCell('S19');
        cell_S19.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S19.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S19.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S19 설정 실패:', e);
    }

    // S20 셀
    try {
        const cell_S20 = worksheet.getCell('S20');
        cell_S20.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S20.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S20.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S20 설정 실패:', e);
    }

    // S21 셀
    try {
        const cell_S21 = worksheet.getCell('S21');
        cell_S21.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S21.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S21.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S21 설정 실패:', e);
    }

    // S22 셀
    try {
        const cell_S22 = worksheet.getCell('S22');
        cell_S22.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S22.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S22.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S22 설정 실패:', e);
    }

    // S23 셀
    try {
        const cell_S23 = worksheet.getCell('S23');
        cell_S23.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S23.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S23.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S23 설정 실패:', e);
    }

    // S24 셀
    try {
        const cell_S24 = worksheet.getCell('S24');
        cell_S24.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S24.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S24.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S24 설정 실패:', e);
    }

    // S25 셀
    try {
        const cell_S25 = worksheet.getCell('S25');
        cell_S25.value = '4877';
        cell_S25.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_S25.alignment = { horizontal: 'center', vertical: 'center' };
        cell_S25.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_S25.numFmt = '"("#,##0.0\ "㎡)"';
    } catch (e) {
        console.warn('셀 S25 설정 실패:', e);
    }

    // S26 셀
    try {
        const cell_S26 = worksheet.getCell('S26');
        cell_S26.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S26.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S26.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S26 설정 실패:', e);
    }

    // S27 셀
    try {
        const cell_S27 = worksheet.getCell('S27');
        cell_S27.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S27.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S27.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S27 설정 실패:', e);
    }

    // S28 셀
    try {
        const cell_S28 = worksheet.getCell('S28');
        cell_S28.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_S28.alignment = { horizontal: 'center', vertical: 'center' };
        cell_S28.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_S28.numFmt = '#,##0.000\ "평"';
    } catch (e) {
        console.warn('셀 S28 설정 실패:', e);
    }

    // S29 셀
    try {
        const cell_S29 = worksheet.getCell('S29');
        cell_S29.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S29.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S29.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S29 설정 실패:', e);
    }

    // S30 셀
    try {
        const cell_S30 = worksheet.getCell('S30');
        cell_S30.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S30.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S30.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S30 설정 실패:', e);
    }

    // S31 셀
    try {
        const cell_S31 = worksheet.getCell('S31');
        cell_S31.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S31.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S31.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S31 설정 실패:', e);
    }

    // S32 셀
    try {
        const cell_S32 = worksheet.getCell('S32');
        cell_S32.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S32.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S32.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S32 설정 실패:', e);
    }

    // S33 셀
    try {
        const cell_S33 = worksheet.getCell('S33');
        cell_S33.value = '임대';
        cell_S33.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_S33.alignment = { horizontal: 'center', vertical: 'center' };
        cell_S33.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_S33.numFmt = '@';
    } catch (e) {
        console.warn('셀 S33 설정 실패:', e);
    }

    // S34 셀
    try {
        const cell_S34 = worksheet.getCell('S34');
        cell_S34.value = '469';
        cell_S34.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_S34.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEECEC' } };
        cell_S34.alignment = { horizontal: 'center', vertical: 'center' };
        cell_S34.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_S34.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 S34 설정 실패:', e);
    }

    // S35 셀
    try {
        const cell_S35 = worksheet.getCell('S35');
        cell_S35.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_S35.alignment = { horizontal: 'center', vertical: 'center' };
        cell_S35.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_S35.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 S35 설정 실패:', e);
    }

    // S36 셀
    try {
        const cell_S36 = worksheet.getCell('S36');
        cell_S36.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_S36.alignment = { horizontal: 'center', vertical: 'center' };
        cell_S36.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_S36.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 S36 설정 실패:', e);
    }

    // S37 셀
    try {
        const cell_S37 = worksheet.getCell('S37');
        cell_S37.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_S37.alignment = { horizontal: 'center', vertical: 'center' };
        cell_S37.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_S37.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 S37 설정 실패:', e);
    }

    // S38 셀
    try {
        const cell_S38 = worksheet.getCell('S38');
        cell_S38.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_S38.alignment = { horizontal: 'center', vertical: 'center' };
        cell_S38.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_S38.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 S38 설정 실패:', e);
    }

    // S39 셀
    try {
        const cell_S39 = worksheet.getCell('S39');
        cell_S39.value = { formula: '=SUM(S34:S38)' };
        cell_S39.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_S39.alignment = { horizontal: 'center', vertical: 'center' };
        cell_S39.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_S39.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 S39 설정 실패:', e);
    }

    // S40 셀
    try {
        const cell_S40 = worksheet.getCell('S40');
        cell_S40.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S40.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S40.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S40 설정 실패:', e);
    }

    // S41 셀
    try {
        const cell_S41 = worksheet.getCell('S41');
        cell_S41.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S41.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S41.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S41 설정 실패:', e);
    }

    // S42 셀
    try {
        const cell_S42 = worksheet.getCell('S42');
        cell_S42.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S42.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S42.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S42 설정 실패:', e);
    }

    // S43 셀
    try {
        const cell_S43 = worksheet.getCell('S43');
        cell_S43.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S43.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S43.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S43 설정 실패:', e);
    }

    // S44 셀
    try {
        const cell_S44 = worksheet.getCell('S44');
        cell_S44.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S44.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S44.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S44 설정 실패:', e);
    }

    // S45 셀
    try {
        const cell_S45 = worksheet.getCell('S45');
        cell_S45.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S45.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S45.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S45 설정 실패:', e);
    }

    // S46 셀
    try {
        const cell_S46 = worksheet.getCell('S46');
        cell_S46.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S46.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S46.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S46 설정 실패:', e);
    }

    // S47 셀
    try {
        const cell_S47 = worksheet.getCell('S47');
        cell_S47.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S47.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S47.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S47 설정 실패:', e);
    }

    // S48 셀
    try {
        const cell_S48 = worksheet.getCell('S48');
        cell_S48.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S48.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S48.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S48 설정 실패:', e);
    }

    // S49 셀
    try {
        const cell_S49 = worksheet.getCell('S49');
        cell_S49.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S49.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S49.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S49 설정 실패:', e);
    }

    // S50 셀
    try {
        const cell_S50 = worksheet.getCell('S50');
        cell_S50.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S50.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S50.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S50 설정 실패:', e);
    }

    // S51 셀
    try {
        const cell_S51 = worksheet.getCell('S51');
        cell_S51.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S51.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S51.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S51 설정 실패:', e);
    }

    // S52 셀
    try {
        const cell_S52 = worksheet.getCell('S52');
        cell_S52.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S52.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S52.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S52 설정 실패:', e);
    }

    // S53 셀
    try {
        const cell_S53 = worksheet.getCell('S53');
        cell_S53.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S53.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S53.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S53 설정 실패:', e);
    }

    // S54 셀
    try {
        const cell_S54 = worksheet.getCell('S54');
        cell_S54.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S54.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S54.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S54 설정 실패:', e);
    }

    // S55 셀
    try {
        const cell_S55 = worksheet.getCell('S55');
        cell_S55.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S55.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S55.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S55 설정 실패:', e);
    }

    // S56 셀
    try {
        const cell_S56 = worksheet.getCell('S56');
        cell_S56.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S56.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S56.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S56 설정 실패:', e);
    }

    // S57 셀
    try {
        const cell_S57 = worksheet.getCell('S57');
        cell_S57.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S57.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S57.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S57 설정 실패:', e);
    }

    // S58 셀
    try {
        const cell_S58 = worksheet.getCell('S58');
        cell_S58.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S58.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S58.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S58 설정 실패:', e);
    }

    // S59 셀
    try {
        const cell_S59 = worksheet.getCell('S59');
        cell_S59.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S59.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S59.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S59 설정 실패:', e);
    }

    // S6 셀
    try {
        const cell_S6 = worksheet.getCell('S6');
        cell_S6.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S6.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S6.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S6 설정 실패:', e);
    }

    // S60 셀
    try {
        const cell_S60 = worksheet.getCell('S60');
        cell_S60.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S60.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S60.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S60 설정 실패:', e);
    }

    // S61 셀
    try {
        const cell_S61 = worksheet.getCell('S61');
        cell_S61.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S61.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S61.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S61 설정 실패:', e);
    }

    // S62 셀
    try {
        const cell_S62 = worksheet.getCell('S62');
        cell_S62.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S62.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S62.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S62 설정 실패:', e);
    }

    // S63 셀
    try {
        const cell_S63 = worksheet.getCell('S63');
        cell_S63.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S63.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S63.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S63 설정 실패:', e);
    }

    // S64 셀
    try {
        const cell_S64 = worksheet.getCell('S64');
        cell_S64.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S64.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S64.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S64 설정 실패:', e);
    }

    // S65 셀
    try {
        const cell_S65 = worksheet.getCell('S65');
        cell_S65.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S65.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S65.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S65 설정 실패:', e);
    }

    // S66 셀
    try {
        const cell_S66 = worksheet.getCell('S66');
        cell_S66.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S66.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S66.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S66 설정 실패:', e);
    }

    // S67 셀
    try {
        const cell_S67 = worksheet.getCell('S67');
        cell_S67.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S67.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S67.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S67 설정 실패:', e);
    }

    // S68 셀
    try {
        const cell_S68 = worksheet.getCell('S68');
        cell_S68.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S68.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S68.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S68 설정 실패:', e);
    }

    // S69 셀
    try {
        const cell_S69 = worksheet.getCell('S69');
        cell_S69.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S69.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S69.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S69 설정 실패:', e);
    }

    // S7 셀
    try {
        const cell_S7 = worksheet.getCell('S7');
        cell_S7.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S7.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S7.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S7 설정 실패:', e);
    }

    // S70 셀
    try {
        const cell_S70 = worksheet.getCell('S70');
        cell_S70.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S70.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S70.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S70 설정 실패:', e);
    }

    // S71 셀
    try {
        const cell_S71 = worksheet.getCell('S71');
        cell_S71.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S71.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S71.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S71 설정 실패:', e);
    }

    // S72 셀
    try {
        const cell_S72 = worksheet.getCell('S72');
        cell_S72.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S72.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S72.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S72 설정 실패:', e);
    }

    // S73 셀
    try {
        const cell_S73 = worksheet.getCell('S73');
        cell_S73.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S73.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S73.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S73 설정 실패:', e);
    }

    // S74 셀
    try {
        const cell_S74 = worksheet.getCell('S74');
        cell_S74.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S74.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S74.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S74 설정 실패:', e);
    }

    // S75 셀
    try {
        const cell_S75 = worksheet.getCell('S75');
        cell_S75.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S75.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S75.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S75 설정 실패:', e);
    }

    // S76 셀
    try {
        const cell_S76 = worksheet.getCell('S76');
        cell_S76.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S76.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S76.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S76 설정 실패:', e);
    }

    // S77 셀
    try {
        const cell_S77 = worksheet.getCell('S77');
        cell_S77.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S77.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S77.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S77 설정 실패:', e);
    }

    // S78 셀
    try {
        const cell_S78 = worksheet.getCell('S78');
        cell_S78.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S78.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S78.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S78 설정 실패:', e);
    }

    // S79 셀
    try {
        const cell_S79 = worksheet.getCell('S79');
        cell_S79.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S79.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S79.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S79 설정 실패:', e);
    }

    // S8 셀
    try {
        const cell_S8 = worksheet.getCell('S8');
        cell_S8.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S8.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S8.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S8 설정 실패:', e);
    }

    // S80 셀
    try {
        const cell_S80 = worksheet.getCell('S80');
        cell_S80.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S80.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S80.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S80 설정 실패:', e);
    }

    // S81 셀
    try {
        const cell_S81 = worksheet.getCell('S81');
        cell_S81.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S81.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S81.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S81 설정 실패:', e);
    }

    // S82 셀
    try {
        const cell_S82 = worksheet.getCell('S82');
        cell_S82.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S82.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S82.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S82 설정 실패:', e);
    }

    // S83 셀
    try {
        const cell_S83 = worksheet.getCell('S83');
        cell_S83.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S83.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S83.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S83 설정 실패:', e);
    }

    // S9 셀
    try {
        const cell_S9 = worksheet.getCell('S9');
        cell_S9.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_S9.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_S9.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 S9 설정 실패:', e);
    }

    // T106 셀
    try {
        const cell_T106 = worksheet.getCell('T106');
        cell_T106.font = { name: 'LG스마트체 Regular', size: 10.0, bold: true, color: { argb: 'FF000000' } };
        cell_T106.alignment = { horizontal: 'center', vertical: 'center' };
        cell_T106.numFmt = '_-* #,##0_-;\-* #,##0_-;_-* "-"_-;_-@_-';
    } catch (e) {
        console.warn('셀 T106 설정 실패:', e);
    }

    // T17 셀
    try {
        const cell_T17 = worksheet.getCell('T17');
        cell_T17.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_T17.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_T17.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 T17 설정 실패:', e);
    }

    // T18 셀
    try {
        const cell_T18 = worksheet.getCell('T18');
        cell_T18.value = '서울시 구로구 디지털로31길 12';
        cell_T18.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_T18.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        cell_T18.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_T18.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 T18 설정 실패:', e);
    }

    // T19 셀
    try {
        const cell_T19 = worksheet.getCell('T19');
        cell_T19.value = '2호선 구로디지털단지역 도보 7분\n7호선 남구로역 도보 7분';
        cell_T19.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_T19.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        cell_T19.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_T19.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 T19 설정 실패:', e);
    }

    // T20 셀
    try {
        const cell_T20 = worksheet.getCell('T20');
        cell_T20.value = '2010';
        cell_T20.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_T20.alignment = { horizontal: 'center', vertical: 'center' };
        cell_T20.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_T20.numFmt = '###0"년"';
    } catch (e) {
        console.warn('셀 T20 설정 실패:', e);
    }

    // T21 셀
    try {
        const cell_T21 = worksheet.getCell('T21');
        cell_T21.value = '194';
        cell_T21.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_T21.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        cell_T21.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_T21.numFmt = '##"F / B"#';
    } catch (e) {
        console.warn('셀 T21 설정 실패:', e);
    }

    // T22 셀
    try {
        const cell_T22 = worksheet.getCell('T22');
        cell_T22.value = { formula: '=46549.73*0.3025' };
        cell_T22.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_T22.alignment = { horizontal: 'center', vertical: 'center' };
        cell_T22.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_T22.numFmt = '#,##0\ "평"';
    } catch (e) {
        console.warn('셀 T22 설정 실패:', e);
    }

    // T23 셀
    try {
        const cell_T23 = worksheet.getCell('T23');
        cell_T23.value = '268.88';
        cell_T23.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_T23.alignment = { horizontal: 'center', vertical: 'center' };
        cell_T23.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_T23.numFmt = '#,##0\ "평"';
    } catch (e) {
        console.warn('셀 T23 설정 실패:', e);
    }

    // T24 셀
    try {
        const cell_T24 = worksheet.getCell('T24');
        cell_T24.value = { formula: '=268.88/522.05' };
        cell_T24.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_T24.alignment = { horizontal: 'center', vertical: 'center' };
        cell_T24.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_T24.numFmt = '0.00%';
    } catch (e) {
        console.warn('셀 T24 설정 실패:', e);
    }

    // T25 셀
    try {
        const cell_T25 = worksheet.getCell('T25');
        cell_T25.value = '268.88';
        cell_T25.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_T25.alignment = { horizontal: 'center', vertical: 'center' };
        cell_T25.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_T25.numFmt = '#,##0\ "평"';
    } catch (e) {
        console.warn('셀 T25 설정 실패:', e);
    }

    // T26 셀
    try {
        const cell_T26 = worksheet.getCell('T26');
        cell_T26.value = 'TP그룹';
        cell_T26.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_T26.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        cell_T26.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_T26.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 T26 설정 실패:', e);
    }

    // T27 셀
    try {
        const cell_T27 = worksheet.getCell('T27');
        cell_T27.value = '전세권, 예금질권 설정 가능';
        cell_T27.font = { name: 'LG스마트체 Regular', size: 9.0, bold: true, color: { argb: 'FFC00000' } };
        cell_T27.alignment = { horizontal: 'center', vertical: 'center' };
        cell_T27.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_T27.numFmt = '0.00%';
    } catch (e) {
        console.warn('셀 T27 설정 실패:', e);
    }

    // T28 셀
    try {
        const cell_T28 = worksheet.getCell('T28');
        cell_T28.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_T28.alignment = { horizontal: 'center', vertical: 'center' };
        cell_T28.border = { top: { style: 'hair', color: { argb: 'FF000000' } } };
        cell_T28.numFmt = '#,##0.000\ "평"';
    } catch (e) {
        console.warn('셀 T28 설정 실패:', e);
    }

    // T29 셀
    try {
        const cell_T29 = worksheet.getCell('T29');
        cell_T29.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_T29.alignment = { horizontal: 'center', vertical: 'center' };
        cell_T29.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_T29.numFmt = '_-* #,##0_-;\-* #,##0_-;_-* "-"_-;_-@_-';
    } catch (e) {
        console.warn('셀 T29 설정 실패:', e);
    }

    // T30 셀
    try {
        const cell_T30 = worksheet.getCell('T30');
        cell_T30.value = { formula: '=T29/T32' };
        cell_T30.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FFC00000' } };
        cell_T30.alignment = { horizontal: 'center', vertical: 'center' };
        cell_T30.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_T30.numFmt = '0.00%';
    } catch (e) {
        console.warn('셀 T30 설정 실패:', e);
    }

    // T31 셀
    try {
        const cell_T31 = worksheet.getCell('T31');
        cell_T31.value = '6602000';
        cell_T31.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_T31.alignment = { horizontal: 'center', vertical: 'center' };
        cell_T31.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_T31.numFmt = '#,###"원/㎡"';
    } catch (e) {
        console.warn('셀 T31 설정 실패:', e);
    }

    // T32 셀
    try {
        const cell_T32 = worksheet.getCell('T32');
        cell_T32.value = { formula: '=T31*V25' };
        cell_T32.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_T32.alignment = { horizontal: 'center', vertical: 'center' };
        cell_T32.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_T32.numFmt = '_-* #,##0_-;\-* #,##0_-;_-* "-"_-;_-@_-';
    } catch (e) {
        console.warn('셀 T32 설정 실패:', e);
    }

    // T33 셀
    try {
        const cell_T33 = worksheet.getCell('T33');
        cell_T33.value = '층';
        cell_T33.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_T33.alignment = { horizontal: 'center', vertical: 'center' };
        cell_T33.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_T33.numFmt = '@';
    } catch (e) {
        console.warn('셀 T33 설정 실패:', e);
    }

    // T34 셀
    try {
        const cell_T34 = worksheet.getCell('T34');
        cell_T34.value = '19';
        cell_T34.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF0000FF' } };
        cell_T34.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEECEC' } };
        cell_T34.alignment = { horizontal: 'center', vertical: 'center' };
        cell_T34.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_T34.numFmt = '#"층"';
    } catch (e) {
        console.warn('셀 T34 설정 실패:', e);
    }

    // T35 셀
    try {
        const cell_T35 = worksheet.getCell('T35');
        cell_T35.value = '18';
        cell_T35.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF0000FF' } };
        cell_T35.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEECEC' } };
        cell_T35.alignment = { horizontal: 'center', vertical: 'center' };
        cell_T35.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_T35.numFmt = '#"층"';
    } catch (e) {
        console.warn('셀 T35 설정 실패:', e);
    }

    // T36 셀
    try {
        const cell_T36 = worksheet.getCell('T36');
        cell_T36.value = '12';
        cell_T36.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_T36.alignment = { horizontal: 'center', vertical: 'center' };
        cell_T36.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_T36.numFmt = '#"층"';
    } catch (e) {
        console.warn('셀 T36 설정 실패:', e);
    }

    // T37 셀
    try {
        const cell_T37 = worksheet.getCell('T37');
        cell_T37.value = '10';
        cell_T37.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_T37.alignment = { horizontal: 'center', vertical: 'center' };
        cell_T37.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_T37.numFmt = '#"층"';
    } catch (e) {
        console.warn('셀 T37 설정 실패:', e);
    }

    // T38 셀
    try {
        const cell_T38 = worksheet.getCell('T38');
        cell_T38.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_T38.alignment = { horizontal: 'center', vertical: 'center' };
        cell_T38.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_T38.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 T38 설정 실패:', e);
    }

    // T39 셀
    try {
        const cell_T39 = worksheet.getCell('T39');
        cell_T39.value = '소계';
        cell_T39.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_T39.alignment = { horizontal: 'center', vertical: 'center' };
        cell_T39.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_T39.numFmt = '@';
    } catch (e) {
        console.warn('셀 T39 설정 실패:', e);
    }

    // T40 셀
    try {
        const cell_T40 = worksheet.getCell('T40');
        cell_T40.value = '2025.7~2027.6 (12개월)';
        cell_T40.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_T40.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        cell_T40.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_T40.numFmt = '#"개월 계약 가능"';
    } catch (e) {
        console.warn('셀 T40 설정 실패:', e);
    }

    // T41 셀
    try {
        const cell_T41 = worksheet.getCell('T41');
        cell_T41.value = '25년 7월';
        cell_T41.font = { name: 'LG스마트체 Regular', size: 9.0, bold: true, color: { argb: 'FF000000' } };
        cell_T41.alignment = { horizontal: 'center', vertical: 'center' };
        cell_T41.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_T41.numFmt = '#,##0\ "원"';
    } catch (e) {
        console.warn('셀 T41 설정 실패:', e);
    }

    // T42 셀
    try {
        const cell_T42 = worksheet.getCell('T42');
        cell_T42.value = '18층 또는 19층 전체';
        cell_T42.font = { name: 'LG스마트체 Regular', size: 9.0, bold: true, color: { argb: 'FF000000' } };
        cell_T42.alignment = { horizontal: 'center', vertical: 'center' };
        cell_T42.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_T42.numFmt = '#,##0\ "층"';
    } catch (e) {
        console.warn('셀 T42 설정 실패:', e);
    }

    // T43 셀
    try {
        const cell_T43 = worksheet.getCell('T43');
        cell_T43.value = { formula: '=U34' };
        cell_T43.font = { name: 'LG스마트체 Regular', size: 9.0, bold: true, color: { argb: 'FFC00000' } };
        cell_T43.alignment = { horizontal: 'center', vertical: 'center' };
        cell_T43.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_T43.numFmt = '#,##0\ "평"';
    } catch (e) {
        console.warn('셀 T43 설정 실패:', e);
    }

    // T44 셀
    try {
        const cell_T44 = worksheet.getCell('T44');
        cell_T44.value = { formula: '=V34' };
        cell_T44.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_T44.alignment = { horizontal: 'center', vertical: 'center' };
        cell_T44.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_T44.numFmt = '#,##0\ "평"';
    } catch (e) {
        console.warn('셀 T44 설정 실패:', e);
    }

    // T45 셀
    try {
        const cell_T45 = worksheet.getCell('T45');
        cell_T45.value = '600000';
        cell_T45.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_T45.alignment = { horizontal: 'center', vertical: 'center' };
        cell_T45.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_T45.numFmt = '"@"#,###';
    } catch (e) {
        console.warn('셀 T45 설정 실패:', e);
    }

    // T46 셀
    try {
        const cell_T46 = worksheet.getCell('T46');
        cell_T46.value = '60000';
        cell_T46.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_T46.alignment = { horizontal: 'center', vertical: 'center' };
        cell_T46.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_T46.numFmt = '"@"#,###';
    } catch (e) {
        console.warn('셀 T46 설정 실패:', e);
    }

    // T47 셀
    try {
        const cell_T47 = worksheet.getCell('T47');
        cell_T47.value = '20000';
        cell_T47.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_T47.alignment = { horizontal: 'center', vertical: 'center' };
        cell_T47.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_T47.numFmt = '"@"#,###\+"실""비""별""도"';
    } catch (e) {
        console.warn('셀 T47 설정 실패:', e);
    }

    // T48 셀
    try {
        const cell_T48 = worksheet.getCell('T48');
        cell_T48.value = { formula: '=T46*(12-T49)/12' };
        cell_T48.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_T48.alignment = { horizontal: 'center', vertical: 'center' };
        cell_T48.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_T48.numFmt = '"@"#,###';
    } catch (e) {
        console.warn('셀 T48 설정 실패:', e);
    }

    // T49 셀
    try {
        const cell_T49 = worksheet.getCell('T49');
        cell_T49.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_T49.alignment = { horizontal: 'center', vertical: 'center' };
        cell_T49.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_T49.numFmt = '#0.0"개월"';
    } catch (e) {
        console.warn('셀 T49 설정 실패:', e);
    }

    // T50 셀
    try {
        const cell_T50 = worksheet.getCell('T50');
        cell_T50.value = { formula: '=T45*T44' };
        cell_T50.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_T50.alignment = { horizontal: 'center', vertical: 'center' };
        cell_T50.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_T50.numFmt = '#,##0\ "원"';
    } catch (e) {
        console.warn('셀 T50 설정 실패:', e);
    }

    // T51 셀
    try {
        const cell_T51 = worksheet.getCell('T51');
        cell_T51.value = { formula: '=T46*T44' };
        cell_T51.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_T51.alignment = { horizontal: 'center', vertical: 'center' };
        cell_T51.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_T51.numFmt = '#,##0\ "원"';
    } catch (e) {
        console.warn('셀 T51 설정 실패:', e);
    }

    // T52 셀
    try {
        const cell_T52 = worksheet.getCell('T52');
        cell_T52.value = { formula: '=T47*T44' };
        cell_T52.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_T52.alignment = { horizontal: 'center', vertical: 'center' };
        cell_T52.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_T52.numFmt = '#,##0\ "원"';
    } catch (e) {
        console.warn('셀 T52 설정 실패:', e);
    }

    // T53 셀
    try {
        const cell_T53 = worksheet.getCell('T53');
        cell_T53.value = '실비 관리비: 전기세, 수도세  별도 부과';
        cell_T53.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FFC00000' } };
        cell_T53.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        cell_T53.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_T53.numFmt = '#,##0\ "원"';
    } catch (e) {
        console.warn('셀 T53 설정 실패:', e);
    }

    // T54 셀
    try {
        const cell_T54 = worksheet.getCell('T54');
        cell_T54.value = { formula: '=T51+T52' };
        cell_T54.font = { name: 'LG스마트체 Regular', size: 9.0, bold: true, color: { argb: 'FFC00000' } };
        cell_T54.alignment = { horizontal: 'center', vertical: 'center' };
        cell_T54.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_T54.numFmt = '#,##0\ "원"';
    } catch (e) {
        console.warn('셀 T54 설정 실패:', e);
    }

    // T55 셀
    try {
        const cell_T55 = worksheet.getCell('T55');
        cell_T55.value = { formula: '=T54*21' };
        cell_T55.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_T55.alignment = { horizontal: 'center', vertical: 'center' };
        cell_T55.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_T55.numFmt = '#,##0\ "원"';
    } catch (e) {
        console.warn('셀 T55 설정 실패:', e);
    }

    // T56 셀
    try {
        const cell_T56 = worksheet.getCell('T56');
        cell_T56.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_T56.alignment = { horizontal: 'center', vertical: 'center' };
        cell_T56.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_T56.numFmt = '0.#"개월"';
    } catch (e) {
        console.warn('셀 T56 설정 실패:', e);
    }

    // T57 셀
    try {
        const cell_T57 = worksheet.getCell('T57');
        cell_T57.value = '미제공';
        cell_T57.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_T57.alignment = { horizontal: 'center', vertical: 'center' };
        cell_T57.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_T57.numFmt = '"총액 "##,##0"원"';
    } catch (e) {
        console.warn('셀 T57 설정 실패:', e);
    }

    // T58 셀
    try {
        const cell_T58 = worksheet.getCell('T58');
        cell_T58.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_T58.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_T58.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 T58 설정 실패:', e);
    }

    // T59 셀
    try {
        const cell_T59 = worksheet.getCell('T59');
        cell_T59.value = '348';
        cell_T59.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_T59.alignment = { horizontal: 'center', vertical: 'center' };
        cell_T59.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_T59.numFmt = '#\ "대"';
    } catch (e) {
        console.warn('셀 T59 설정 실패:', e);
    }

    // T6 셀
    try {
        const cell_T6 = worksheet.getCell('T6');
        cell_T6.value = '구로디지털단지역';
        cell_T6.font = { name: 'LG스마트체 Regular', size: 10.0, color: { argb: 'FF000000' } };
        cell_T6.alignment = { horizontal: 'center', vertical: 'center' };
        cell_T6.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_T6.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 T6 설정 실패:', e);
    }

    // T60 셀
    try {
        const cell_T60 = worksheet.getCell('T60');
        cell_T60.value = '50';
        cell_T60.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_T60.alignment = { horizontal: 'center', vertical: 'center' };
        cell_T60.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_T60.numFmt = '"임대면적"\ #"평당 1대"';
    } catch (e) {
        console.warn('셀 T60 설정 실패:', e);
    }

    // T61 셀
    try {
        const cell_T61 = worksheet.getCell('T61');
        cell_T61.value = { formula: '=T44/T60' };
        cell_T61.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_T61.alignment = { horizontal: 'center', vertical: 'center' };
        cell_T61.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_T61.numFmt = '#,##0.0\ "대"';
    } catch (e) {
        console.warn('셀 T61 설정 실패:', e);
    }

    // T62 셀
    try {
        const cell_T62 = worksheet.getCell('T62');
        cell_T62.value = '12';
        cell_T62.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_T62.alignment = { horizontal: 'center', vertical: 'center' };
        cell_T62.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_T62.numFmt = '"월"#"만원/대"';
    } catch (e) {
        console.warn('셀 T62 설정 실패:', e);
    }

    // T63 셀
    try {
        const cell_T63 = worksheet.getCell('T63');
        cell_T63.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_T63.alignment = { horizontal: 'left', vertical: 'center', wrapText: true };
        cell_T63.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_T63.numFmt = '#,##0_ ';
    } catch (e) {
        console.warn('셀 T63 설정 실패:', e);
    }

    // T7 셀
    try {
        const cell_T7 = worksheet.getCell('T7');
        cell_T7.font = { name: 'LG스마트체 Regular', size: 10.0, color: { argb: 'FF000000' } };
        cell_T7.alignment = { horizontal: 'center', vertical: 'center' };
        cell_T7.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_T7.numFmt = '0_);[Red]\(0\)';
    } catch (e) {
        console.warn('셀 T7 설정 실패:', e);
    }

    // T72 셀
    try {
        const cell_T72 = worksheet.getCell('T72');
        cell_T72.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_T72.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_T72.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 T72 설정 실패:', e);
    }

    // T73 셀
    try {
        const cell_T73 = worksheet.getCell('T73');
        cell_T73.value = ' - 10층(1호): 즉시 가능 \n - 12층(3-4호): 25년 7~8월 입주\n - 18층: 25년 7월 입주\n - 19층: 25년 7월 입주\n - 실비(유틸리티비용) 별도\n - 임대인 직접 관리 \n - 베란다 서비스 면적 포함';
        cell_T73.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_T73.alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
        cell_T73.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_T73.numFmt = '#,##0\ "대"';
    } catch (e) {
        console.warn('셀 T73 설정 실패:', e);
    }

    // T8 셀
    try {
        const cell_T8 = worksheet.getCell('T8');
        cell_T8.value = '구로 TP타워';
        cell_T8.font = { name: 'LG스마트체 Regular', size: 10.0, color: { argb: 'FF000000' } };
        cell_T8.alignment = { horizontal: 'center', vertical: 'center' };
        cell_T8.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_T8.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 T8 설정 실패:', e);
    }

    // T83 셀
    try {
        const cell_T83 = worksheet.getCell('T83');
        cell_T83.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_T83.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_T83.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 T83 설정 실패:', e);
    }

    // T9 셀
    try {
        const cell_T9 = worksheet.getCell('T9');
        cell_T9.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_T9.alignment = { horizontal: 'center', vertical: 'center' };
        cell_T9.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_T9.numFmt = '#,##0';
    } catch (e) {
        console.warn('셀 T9 설정 실패:', e);
    }

    // U106 셀
    try {
        const cell_U106 = worksheet.getCell('U106');
        cell_U106.font = { name: 'LG스마트체 Regular', size: 10.0, bold: true, color: { argb: 'FF000000' } };
        cell_U106.alignment = { horizontal: 'center', vertical: 'center' };
        cell_U106.numFmt = '_-* #,##0_-;\-* #,##0_-;_-* "-"_-;_-@_-';
    } catch (e) {
        console.warn('셀 U106 설정 실패:', e);
    }

    // U17 셀
    try {
        const cell_U17 = worksheet.getCell('U17');
        cell_U17.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_U17.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_U17.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 U17 설정 실패:', e);
    }

    // U18 셀
    try {
        const cell_U18 = worksheet.getCell('U18');
        cell_U18.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_U18.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_U18.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 U18 설정 실패:', e);
    }

    // U19 셀
    try {
        const cell_U19 = worksheet.getCell('U19');
        cell_U19.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_U19.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_U19.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 U19 설정 실패:', e);
    }

    // U20 셀
    try {
        const cell_U20 = worksheet.getCell('U20');
        cell_U20.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_U20.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_U20.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 U20 설정 실패:', e);
    }

    // U21 셀
    try {
        const cell_U21 = worksheet.getCell('U21');
        cell_U21.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_U21.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_U21.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 U21 설정 실패:', e);
    }

    // U22 셀
    try {
        const cell_U22 = worksheet.getCell('U22');
        cell_U22.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_U22.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_U22.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 U22 설정 실패:', e);
    }

    // U23 셀
    try {
        const cell_U23 = worksheet.getCell('U23');
        cell_U23.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_U23.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_U23.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 U23 설정 실패:', e);
    }

    // U24 셀
    try {
        const cell_U24 = worksheet.getCell('U24');
        cell_U24.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_U24.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_U24.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 U24 설정 실패:', e);
    }

    // U25 셀
    try {
        const cell_U25 = worksheet.getCell('U25');
        cell_U25.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_U25.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_U25.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 U25 설정 실패:', e);
    }

    // U26 셀
    try {
        const cell_U26 = worksheet.getCell('U26');
        cell_U26.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_U26.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_U26.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 U26 설정 실패:', e);
    }

    // U27 셀
    try {
        const cell_U27 = worksheet.getCell('U27');
        cell_U27.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_U27.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_U27.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 U27 설정 실패:', e);
    }

    // U28 셀
    try {
        const cell_U28 = worksheet.getCell('U28');
        cell_U28.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_U28.alignment = { horizontal: 'center', vertical: 'center' };
        cell_U28.border = { top: { style: 'hair', color: { argb: 'FF000000' } } };
        cell_U28.numFmt = '#,##0.000\ "평"';
    } catch (e) {
        console.warn('셀 U28 설정 실패:', e);
    }

    // U29 셀
    try {
        const cell_U29 = worksheet.getCell('U29');
        cell_U29.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_U29.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_U29.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 U29 설정 실패:', e);
    }

    // U30 셀
    try {
        const cell_U30 = worksheet.getCell('U30');
        cell_U30.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_U30.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_U30.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 U30 설정 실패:', e);
    }

    // U31 셀
    try {
        const cell_U31 = worksheet.getCell('U31');
        cell_U31.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_U31.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_U31.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 U31 설정 실패:', e);
    }

    // U32 셀
    try {
        const cell_U32 = worksheet.getCell('U32');
        cell_U32.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_U32.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_U32.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 U32 설정 실패:', e);
    }

    // U33 셀
    try {
        const cell_U33 = worksheet.getCell('U33');
        cell_U33.value = '전용';
        cell_U33.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_U33.alignment = { horizontal: 'center', vertical: 'center' };
        cell_U33.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_U33.numFmt = '@';
    } catch (e) {
        console.warn('셀 U33 설정 실패:', e);
    }

    // U34 셀
    try {
        const cell_U34 = worksheet.getCell('U34');
        cell_U34.value = '268.88';
        cell_U34.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF0000FF' } };
        cell_U34.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEECEC' } };
        cell_U34.alignment = { horizontal: 'center', vertical: 'center' };
        cell_U34.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_U34.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 U34 설정 실패:', e);
    }

    // U35 셀
    try {
        const cell_U35 = worksheet.getCell('U35');
        cell_U35.value = '268.88';
        cell_U35.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF0000FF' } };
        cell_U35.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEECEC' } };
        cell_U35.alignment = { horizontal: 'center', vertical: 'center' };
        cell_U35.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_U35.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 U35 설정 실패:', e);
    }

    // U36 셀
    try {
        const cell_U36 = worksheet.getCell('U36');
        cell_U36.value = '129.27';
        cell_U36.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_U36.alignment = { horizontal: 'center', vertical: 'center' };
        cell_U36.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_U36.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 U36 설정 실패:', e);
    }

    // U37 셀
    try {
        const cell_U37 = worksheet.getCell('U37');
        cell_U37.value = '26.11';
        cell_U37.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_U37.alignment = { horizontal: 'center', vertical: 'center' };
        cell_U37.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_U37.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 U37 설정 실패:', e);
    }

    // U38 셀
    try {
        const cell_U38 = worksheet.getCell('U38');
        cell_U38.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_U38.alignment = { horizontal: 'center', vertical: 'center' };
        cell_U38.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_U38.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 U38 설정 실패:', e);
    }

    // U39 셀
    try {
        const cell_U39 = worksheet.getCell('U39');
        cell_U39.value = { formula: '=SUM(U34:U38)' };
        cell_U39.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_U39.alignment = { horizontal: 'center', vertical: 'center' };
        cell_U39.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_U39.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 U39 설정 실패:', e);
    }

    // U40 셀
    try {
        const cell_U40 = worksheet.getCell('U40');
        cell_U40.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_U40.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_U40.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 U40 설정 실패:', e);
    }

    // U41 셀
    try {
        const cell_U41 = worksheet.getCell('U41');
        cell_U41.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_U41.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_U41.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 U41 설정 실패:', e);
    }

    // U42 셀
    try {
        const cell_U42 = worksheet.getCell('U42');
        cell_U42.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_U42.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_U42.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 U42 설정 실패:', e);
    }

    // U43 셀
    try {
        const cell_U43 = worksheet.getCell('U43');
        cell_U43.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_U43.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_U43.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 U43 설정 실패:', e);
    }

    // U44 셀
    try {
        const cell_U44 = worksheet.getCell('U44');
        cell_U44.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_U44.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_U44.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 U44 설정 실패:', e);
    }

    // U45 셀
    try {
        const cell_U45 = worksheet.getCell('U45');
        cell_U45.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_U45.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_U45.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 U45 설정 실패:', e);
    }

    // U46 셀
    try {
        const cell_U46 = worksheet.getCell('U46');
        cell_U46.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_U46.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_U46.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 U46 설정 실패:', e);
    }

    // U47 셀
    try {
        const cell_U47 = worksheet.getCell('U47');
        cell_U47.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_U47.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_U47.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 U47 설정 실패:', e);
    }

    // U48 셀
    try {
        const cell_U48 = worksheet.getCell('U48');
        cell_U48.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_U48.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_U48.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 U48 설정 실패:', e);
    }

    // U49 셀
    try {
        const cell_U49 = worksheet.getCell('U49');
        cell_U49.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_U49.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_U49.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 U49 설정 실패:', e);
    }

    // U50 셀
    try {
        const cell_U50 = worksheet.getCell('U50');
        cell_U50.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_U50.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_U50.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 U50 설정 실패:', e);
    }

    // U51 셀
    try {
        const cell_U51 = worksheet.getCell('U51');
        cell_U51.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_U51.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_U51.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 U51 설정 실패:', e);
    }

    // U52 셀
    try {
        const cell_U52 = worksheet.getCell('U52');
        cell_U52.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_U52.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_U52.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 U52 설정 실패:', e);
    }

    // U53 셀
    try {
        const cell_U53 = worksheet.getCell('U53');
        cell_U53.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_U53.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_U53.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 U53 설정 실패:', e);
    }

    // U54 셀
    try {
        const cell_U54 = worksheet.getCell('U54');
        cell_U54.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_U54.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_U54.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 U54 설정 실패:', e);
    }

    // U55 셀
    try {
        const cell_U55 = worksheet.getCell('U55');
        cell_U55.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_U55.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_U55.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 U55 설정 실패:', e);
    }

    // U56 셀
    try {
        const cell_U56 = worksheet.getCell('U56');
        cell_U56.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_U56.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_U56.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 U56 설정 실패:', e);
    }

    // U57 셀
    try {
        const cell_U57 = worksheet.getCell('U57');
        cell_U57.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_U57.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_U57.border = { top: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 U57 설정 실패:', e);
    }

    // U58 셀
    try {
        const cell_U58 = worksheet.getCell('U58');
        cell_U58.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_U58.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_U58.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 U58 설정 실패:', e);
    }

    // U59 셀
    try {
        const cell_U59 = worksheet.getCell('U59');
        cell_U59.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_U59.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_U59.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 U59 설정 실패:', e);
    }

    // U6 셀
    try {
        const cell_U6 = worksheet.getCell('U6');
        cell_U6.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_U6.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_U6.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 U6 설정 실패:', e);
    }

    // U60 셀
    try {
        const cell_U60 = worksheet.getCell('U60');
        cell_U60.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_U60.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_U60.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 U60 설정 실패:', e);
    }

    // U61 셀
    try {
        const cell_U61 = worksheet.getCell('U61');
        cell_U61.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_U61.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_U61.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 U61 설정 실패:', e);
    }

    // U62 셀
    try {
        const cell_U62 = worksheet.getCell('U62');
        cell_U62.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_U62.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_U62.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 U62 설정 실패:', e);
    }

    // U63 셀
    try {
        const cell_U63 = worksheet.getCell('U63');
        cell_U63.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_U63.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_U63.border = { top: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 U63 설정 실패:', e);
    }

    // U7 셀
    try {
        const cell_U7 = worksheet.getCell('U7');
        cell_U7.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_U7.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_U7.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 U7 설정 실패:', e);
    }

    // U72 셀
    try {
        const cell_U72 = worksheet.getCell('U72');
        cell_U72.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_U72.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_U72.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 U72 설정 실패:', e);
    }

    // U73 셀
    try {
        const cell_U73 = worksheet.getCell('U73');
        cell_U73.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_U73.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_U73.border = { top: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 U73 설정 실패:', e);
    }

    // U8 셀
    try {
        const cell_U8 = worksheet.getCell('U8');
        cell_U8.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_U8.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_U8.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 U8 설정 실패:', e);
    }

    // U83 셀
    try {
        const cell_U83 = worksheet.getCell('U83');
        cell_U83.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_U83.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_U83.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 U83 설정 실패:', e);
    }

    // U9 셀
    try {
        const cell_U9 = worksheet.getCell('U9');
        cell_U9.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_U9.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_U9.border = { top: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 U9 설정 실패:', e);
    }

    // V10 셀
    try {
        const cell_V10 = worksheet.getCell('V10');
        cell_V10.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V10.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V10.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V10 설정 실패:', e);
    }

    // V106 셀
    try {
        const cell_V106 = worksheet.getCell('V106');
        cell_V106.font = { name: 'LG스마트체 Regular', size: 10.0, bold: true, color: { argb: 'FF000000' } };
        cell_V106.alignment = { horizontal: 'center', vertical: 'center' };
        cell_V106.numFmt = '_-* #,##0_-;\-* #,##0_-;_-* "-"_-;_-@_-';
    } catch (e) {
        console.warn('셀 V106 설정 실패:', e);
    }

    // V11 셀
    try {
        const cell_V11 = worksheet.getCell('V11');
        cell_V11.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V11.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V11.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V11 설정 실패:', e);
    }

    // V12 셀
    try {
        const cell_V12 = worksheet.getCell('V12');
        cell_V12.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V12.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V12.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V12 설정 실패:', e);
    }

    // V13 셀
    try {
        const cell_V13 = worksheet.getCell('V13');
        cell_V13.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V13.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V13.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V13 설정 실패:', e);
    }

    // V14 셀
    try {
        const cell_V14 = worksheet.getCell('V14');
        cell_V14.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V14.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V14.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V14 설정 실패:', e);
    }

    // V15 셀
    try {
        const cell_V15 = worksheet.getCell('V15');
        cell_V15.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V15.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V15.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V15 설정 실패:', e);
    }

    // V16 셀
    try {
        const cell_V16 = worksheet.getCell('V16');
        cell_V16.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V16.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V16.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V16 설정 실패:', e);
    }

    // V17 셀
    try {
        const cell_V17 = worksheet.getCell('V17');
        cell_V17.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V17.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V17.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V17 설정 실패:', e);
    }

    // V18 셀
    try {
        const cell_V18 = worksheet.getCell('V18');
        cell_V18.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V18.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V18.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V18 설정 실패:', e);
    }

    // V19 셀
    try {
        const cell_V19 = worksheet.getCell('V19');
        cell_V19.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V19.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V19.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V19 설정 실패:', e);
    }

    // V20 셀
    try {
        const cell_V20 = worksheet.getCell('V20');
        cell_V20.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V20.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V20.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V20 설정 실패:', e);
    }

    // V21 셀
    try {
        const cell_V21 = worksheet.getCell('V21');
        cell_V21.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V21.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V21.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V21 설정 실패:', e);
    }

    // V22 셀
    try {
        const cell_V22 = worksheet.getCell('V22');
        cell_V22.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V22.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V22.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V22 설정 실패:', e);
    }

    // V23 셀
    try {
        const cell_V23 = worksheet.getCell('V23');
        cell_V23.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V23.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V23.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V23 설정 실패:', e);
    }

    // V24 셀
    try {
        const cell_V24 = worksheet.getCell('V24');
        cell_V24.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V24.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V24.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V24 설정 실패:', e);
    }

    // V25 셀
    try {
        const cell_V25 = worksheet.getCell('V25');
        cell_V25.value = '6320';
        cell_V25.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_V25.alignment = { horizontal: 'center', vertical: 'center' };
        cell_V25.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_V25.numFmt = '"("#,##0.0\ "㎡)"';
    } catch (e) {
        console.warn('셀 V25 설정 실패:', e);
    }

    // V26 셀
    try {
        const cell_V26 = worksheet.getCell('V26');
        cell_V26.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V26.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V26.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V26 설정 실패:', e);
    }

    // V27 셀
    try {
        const cell_V27 = worksheet.getCell('V27');
        cell_V27.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V27.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V27.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V27 설정 실패:', e);
    }

    // V28 셀
    try {
        const cell_V28 = worksheet.getCell('V28');
        cell_V28.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_V28.alignment = { horizontal: 'center', vertical: 'center' };
        cell_V28.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_V28.numFmt = '#,##0.000\ "평"';
    } catch (e) {
        console.warn('셀 V28 설정 실패:', e);
    }

    // V29 셀
    try {
        const cell_V29 = worksheet.getCell('V29');
        cell_V29.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V29.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V29.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V29 설정 실패:', e);
    }

    // V30 셀
    try {
        const cell_V30 = worksheet.getCell('V30');
        cell_V30.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V30.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V30.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V30 설정 실패:', e);
    }

    // V31 셀
    try {
        const cell_V31 = worksheet.getCell('V31');
        cell_V31.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V31.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V31.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V31 설정 실패:', e);
    }

    // V32 셀
    try {
        const cell_V32 = worksheet.getCell('V32');
        cell_V32.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V32.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V32.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V32 설정 실패:', e);
    }

    // V33 셀
    try {
        const cell_V33 = worksheet.getCell('V33');
        cell_V33.value = '임대';
        cell_V33.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_V33.alignment = { horizontal: 'center', vertical: 'center' };
        cell_V33.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_V33.numFmt = '@';
    } catch (e) {
        console.warn('셀 V33 설정 실패:', e);
    }

    // V34 셀
    try {
        const cell_V34 = worksheet.getCell('V34');
        cell_V34.value = '522.05';
        cell_V34.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_V34.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEECEC' } };
        cell_V34.alignment = { horizontal: 'center', vertical: 'center' };
        cell_V34.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_V34.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 V34 설정 실패:', e);
    }

    // V35 셀
    try {
        const cell_V35 = worksheet.getCell('V35');
        cell_V35.value = '522.05';
        cell_V35.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_V35.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEECEC' } };
        cell_V35.alignment = { horizontal: 'center', vertical: 'center' };
        cell_V35.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_V35.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 V35 설정 실패:', e);
    }

    // V36 셀
    try {
        const cell_V36 = worksheet.getCell('V36');
        cell_V36.value = '249.45';
        cell_V36.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_V36.alignment = { horizontal: 'center', vertical: 'center' };
        cell_V36.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_V36.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 V36 설정 실패:', e);
    }

    // V37 셀
    try {
        const cell_V37 = worksheet.getCell('V37');
        cell_V37.value = '50.38';
        cell_V37.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_V37.alignment = { horizontal: 'center', vertical: 'center' };
        cell_V37.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_V37.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 V37 설정 실패:', e);
    }

    // V38 셀
    try {
        const cell_V38 = worksheet.getCell('V38');
        cell_V38.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_V38.alignment = { horizontal: 'center', vertical: 'center' };
        cell_V38.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_V38.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 V38 설정 실패:', e);
    }

    // V39 셀
    try {
        const cell_V39 = worksheet.getCell('V39');
        cell_V39.value = { formula: '=SUM(V34:V38)' };
        cell_V39.font = { name: 'LG스마트체 Regular', size: 9.0, color: { argb: 'FF000000' } };
        cell_V39.alignment = { horizontal: 'center', vertical: 'center' };
        cell_V39.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
        cell_V39.numFmt = '#,###"평"';
    } catch (e) {
        console.warn('셀 V39 설정 실패:', e);
    }

    // V40 셀
    try {
        const cell_V40 = worksheet.getCell('V40');
        cell_V40.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V40.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V40.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V40 설정 실패:', e);
    }

    // V41 셀
    try {
        const cell_V41 = worksheet.getCell('V41');
        cell_V41.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V41.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V41.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V41 설정 실패:', e);
    }

    // V42 셀
    try {
        const cell_V42 = worksheet.getCell('V42');
        cell_V42.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V42.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V42.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V42 설정 실패:', e);
    }

    // V43 셀
    try {
        const cell_V43 = worksheet.getCell('V43');
        cell_V43.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V43.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V43.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V43 설정 실패:', e);
    }

    // V44 셀
    try {
        const cell_V44 = worksheet.getCell('V44');
        cell_V44.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V44.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V44.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V44 설정 실패:', e);
    }

    // V45 셀
    try {
        const cell_V45 = worksheet.getCell('V45');
        cell_V45.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V45.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V45.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V45 설정 실패:', e);
    }

    // V46 셀
    try {
        const cell_V46 = worksheet.getCell('V46');
        cell_V46.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V46.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V46.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V46 설정 실패:', e);
    }

    // V47 셀
    try {
        const cell_V47 = worksheet.getCell('V47');
        cell_V47.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V47.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V47.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V47 설정 실패:', e);
    }

    // V48 셀
    try {
        const cell_V48 = worksheet.getCell('V48');
        cell_V48.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V48.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V48.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V48 설정 실패:', e);
    }

    // V49 셀
    try {
        const cell_V49 = worksheet.getCell('V49');
        cell_V49.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V49.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V49.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V49 설정 실패:', e);
    }

    // V50 셀
    try {
        const cell_V50 = worksheet.getCell('V50');
        cell_V50.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V50.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V50.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V50 설정 실패:', e);
    }

    // V51 셀
    try {
        const cell_V51 = worksheet.getCell('V51');
        cell_V51.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V51.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V51.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V51 설정 실패:', e);
    }

    // V52 셀
    try {
        const cell_V52 = worksheet.getCell('V52');
        cell_V52.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V52.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V52.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V52 설정 실패:', e);
    }

    // V53 셀
    try {
        const cell_V53 = worksheet.getCell('V53');
        cell_V53.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V53.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V53.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V53 설정 실패:', e);
    }

    // V54 셀
    try {
        const cell_V54 = worksheet.getCell('V54');
        cell_V54.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V54.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V54.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V54 설정 실패:', e);
    }

    // V55 셀
    try {
        const cell_V55 = worksheet.getCell('V55');
        cell_V55.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V55.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V55.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V55 설정 실패:', e);
    }

    // V56 셀
    try {
        const cell_V56 = worksheet.getCell('V56');
        cell_V56.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V56.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V56.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V56 설정 실패:', e);
    }

    // V57 셀
    try {
        const cell_V57 = worksheet.getCell('V57');
        cell_V57.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V57.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V57.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V57 설정 실패:', e);
    }

    // V58 셀
    try {
        const cell_V58 = worksheet.getCell('V58');
        cell_V58.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V58.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V58.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V58 설정 실패:', e);
    }

    // V59 셀
    try {
        const cell_V59 = worksheet.getCell('V59');
        cell_V59.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V59.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V59.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V59 설정 실패:', e);
    }

    // V6 셀
    try {
        const cell_V6 = worksheet.getCell('V6');
        cell_V6.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V6.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V6.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V6 설정 실패:', e);
    }

    // V60 셀
    try {
        const cell_V60 = worksheet.getCell('V60');
        cell_V60.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V60.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V60.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V60 설정 실패:', e);
    }

    // V61 셀
    try {
        const cell_V61 = worksheet.getCell('V61');
        cell_V61.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V61.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V61.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'hair', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V61 설정 실패:', e);
    }

    // V62 셀
    try {
        const cell_V62 = worksheet.getCell('V62');
        cell_V62.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V62.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V62.border = { top: { style: 'hair', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V62 설정 실패:', e);
    }

    // V63 셀
    try {
        const cell_V63 = worksheet.getCell('V63');
        cell_V63.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V63.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V63.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V63 설정 실패:', e);
    }

    // V64 셀
    try {
        const cell_V64 = worksheet.getCell('V64');
        cell_V64.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V64.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V64.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V64 설정 실패:', e);
    }

    // V65 셀
    try {
        const cell_V65 = worksheet.getCell('V65');
        cell_V65.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V65.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V65.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V65 설정 실패:', e);
    }

    // V66 셀
    try {
        const cell_V66 = worksheet.getCell('V66');
        cell_V66.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V66.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V66.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V66 설정 실패:', e);
    }

    // V67 셀
    try {
        const cell_V67 = worksheet.getCell('V67');
        cell_V67.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V67.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V67.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V67 설정 실패:', e);
    }

    // V68 셀
    try {
        const cell_V68 = worksheet.getCell('V68');
        cell_V68.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V68.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V68.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V68 설정 실패:', e);
    }

    // V69 셀
    try {
        const cell_V69 = worksheet.getCell('V69');
        cell_V69.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V69.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V69.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V69 설정 실패:', e);
    }

    // V7 셀
    try {
        const cell_V7 = worksheet.getCell('V7');
        cell_V7.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V7.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V7.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V7 설정 실패:', e);
    }

    // V70 셀
    try {
        const cell_V70 = worksheet.getCell('V70');
        cell_V70.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V70.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V70.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V70 설정 실패:', e);
    }

    // V71 셀
    try {
        const cell_V71 = worksheet.getCell('V71');
        cell_V71.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V71.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V71.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V71 설정 실패:', e);
    }

    // V72 셀
    try {
        const cell_V72 = worksheet.getCell('V72');
        cell_V72.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V72.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V72.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V72 설정 실패:', e);
    }

    // V73 셀
    try {
        const cell_V73 = worksheet.getCell('V73');
        cell_V73.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V73.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V73.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V73 설정 실패:', e);
    }

    // V74 셀
    try {
        const cell_V74 = worksheet.getCell('V74');
        cell_V74.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V74.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V74.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V74 설정 실패:', e);
    }

    // V75 셀
    try {
        const cell_V75 = worksheet.getCell('V75');
        cell_V75.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V75.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V75.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V75 설정 실패:', e);
    }

    // V76 셀
    try {
        const cell_V76 = worksheet.getCell('V76');
        cell_V76.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V76.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V76.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V76 설정 실패:', e);
    }

    // V77 셀
    try {
        const cell_V77 = worksheet.getCell('V77');
        cell_V77.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V77.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V77.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V77 설정 실패:', e);
    }

    // V78 셀
    try {
        const cell_V78 = worksheet.getCell('V78');
        cell_V78.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V78.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V78.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V78 설정 실패:', e);
    }

    // V79 셀
    try {
        const cell_V79 = worksheet.getCell('V79');
        cell_V79.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V79.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V79.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V79 설정 실패:', e);
    }

    // V8 셀
    try {
        const cell_V8 = worksheet.getCell('V8');
        cell_V8.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V8.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V8.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V8 설정 실패:', e);
    }

    // V80 셀
    try {
        const cell_V80 = worksheet.getCell('V80');
        cell_V80.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V80.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V80.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V80 설정 실패:', e);
    }

    // V81 셀
    try {
        const cell_V81 = worksheet.getCell('V81');
        cell_V81.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V81.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V81.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V81 설정 실패:', e);
    }

    // V82 셀
    try {
        const cell_V82 = worksheet.getCell('V82');
        cell_V82.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V82.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V82.border = { right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V82 설정 실패:', e);
    }

    // V83 셀
    try {
        const cell_V83 = worksheet.getCell('V83');
        cell_V83.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V83.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V83.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V83 설정 실패:', e);
    }

    // V9 셀
    try {
        const cell_V9 = worksheet.getCell('V9');
        cell_V9.font = { name: 'LG스마트체 Regular', size: 11.0, color: { argb: 'FF000000' } };
        cell_V9.alignment = { horizontal: 'center', vertical: 'middle' };
        cell_V9.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
    } catch (e) {
        console.warn('셀 V9 설정 실패:', e);
    }

    console.log('총 1461개의 셀 설정 완료');
    console.log('수식 87개 적용 완료');
}

// 빌딩 데이터 매핑 도우미 함수
function mapBuildingData(building, key) {
    const mapping = {
        'name': building.name || '',
        'address': building.address || '',
        'completionYear': building.completionYear || '',
        'floors': building.floors || '',
        'baseFloorAreaDedicatedPy': building.baseFloorAreaDedicatedPy || '',
        'totalFloorArea': building.totalFloorArea || '',
        'parkingCount': building.parkingCount || '',
        'elevatorCount': building.elevatorCount || ''
    };
    return mapping[key] || '';
}

// 디버깅 정보 출력 함수
function debugExcelGeneration() {
    console.log('=== 엑셀 생성 디버깅 정보 ===');
    console.log('1. A1-A4: 샘플 문구로 교체됨');
    console.log('2. B43: "전용면적" 항목명 유지');
    console.log('3. 86-88행: 제거됨 (높이 0)');
    console.log('4. 폰트: LG스마트체 Regular 적용');
    console.log('5. 정렬: 모든 셀 가운데 정렬');
    console.log('6. 수식: 총 87개 적용');
    console.log('7. 색상: 원본 배경색, 폰트색, 테두리 유지');
}
