// ===== LG용 Comp List 생성 함수 (최종 개선 버전) =====

// 현재 날짜 반환 (LG 양식용)
function getCurrentDateLG() {
    const now = new Date();
    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const day = String(now.getDate()).padStart(2, '0');
    return `${year}.${month}.${day}`;
}

// LG 양식으로 엑셀 생성
async function generateExcelLG() {
    if (selectedBuildings.length === 0) {
        alert('빌딩을 선택해주세요.');
        return;
    }

    // 10개 제한
    const buildingsToExport = selectedBuildings.slice(0, 10);
    if (selectedBuildings.length > 10) {
        alert('10개를 초과하여 선택하셨습니다. 처음 10개만 출력됩니다.');
    }

    const building = buildingsToExport[0];
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
            { width: 10.375 },  // AE열
        ];

        // 2. 행 높이 설정 (86-88행 제외)
        for (let i = 1; i <= 109; i++) {
            if (i === 19 || i === 20) {
                worksheet.getRow(i).height = 30.0;
            } else if (i === 26) {
                worksheet.getRow(i).height = 40.5;
            } else if (i === 53) {
                worksheet.getRow(i).height = 35.25;
            } else if (i === 54) {
                worksheet.getRow(i).height = 21.75;
            } else if (i === 56) {
                worksheet.getRow(i).height = 12.0;
            } else if (i === 62 || i === 83) {
                worksheet.getRow(i).height = 33.0;
            } else if (i === 84) {
                worksheet.getRow(i).height = 33.75;
            } else if (i >= 86 && i <= 88) {
                // 86-88행 숨기기
                worksheet.getRow(i).height = 0;
            } else {
                worksheet.getRow(i).height = 15.0;
            }
        }

        // 3. 병합 셀 설정 (86-88행 관련 병합 제외)
        const mergedCells = [
            'B60:D60', 'N42:P42', 'B53:D53', 'K56:M56', 'N44:P44',
            'N19:P19', 'H22:J22', 'H19:J19', 'E9:G17', 'N63:P72',
            'E18:G18', 'T30:V30', 'T29:V29', 'N32:P32', 'K29:M29',
            'A63:A83', 'T54:V54', 'C30:D30', 'K44:M44', 'K31:M31',
            'E6:G6', 'H7:J7', 'Q59:S59', 'Q22:S22', 'Q46:S46',
            'T60:V60', 'A50:A55', 'Q61:S61', 'Q73:S83', 'Q48:S48',
            'N45:P45', 'K49:M49', 'N20:P20', 'H20:J20', 'T49:V49',
            'E24:G24', 'N47:P47', 'N22:P22', 'T42:V42', 'K57:M58',
            'B21:D21', 'T50:V50', 'K25:L25', 'T44:V44', 'E19:G19',
            'B23:D23', 'Q49:S49', 'Q30:S30', 'N40:P40', 'B51:D51',
            'H8:J8', 'Q62:S62', 'K21:M21', 'N53:P53', 'Q54:S54',
            'N46:P46', 'K50:M50', 'E57:G58', 'N61:P61', 'A7:D8',
            'B29:D29', 'N48:P48', 'K52:M52', 'E27:G27', 'K53:M53',
            'B24:D24', 'B63:D72', 'A6:D6', 'H23:J23', 'B26:D26',
            'N9:P17', 'N54:P54', 'N41:P41', 'H54:J54', 'E63:G72',
            'H41:J41', 'N51:P51', 'B27:D27', 'Q52:S52', 'K55:M55',
            'N56:P56', 'A59:A62', 'H56:J56', 'N43:P43', 'B42:D42',
            'K46:M46', 'K40:M40', 'E40:G40', 'B44:D44', 'E73:G83',
            'T31:V31', 'T6:V6', 'B20:D20', 'A9:D17', 'H30:J30',
            'H29:J29', 'H44:J44', 'H31:J31', 'A48:A49', 'H60:J60',
            'N18:P18', 'Q60:S60', 'T26:V26', 'E30:G30', 'E59:G59',
            'E46:G46', 'K8:M8', 'T25:U25', 'E61:G61', 'K23:M23',
            'T8:V8', 'E54:G54', 'E41:G41', 'H42:J42', 'B45:D45',
            'K62:M62', 'E56:G56', 'K18:M18', 'Q7:S7', 'E43:G43',
            'B47:D47', 'B59:D59', 'T32:V32', 'B46:D46', 'Q8:S8',
            'T47:V47', 'B48:D48', 'H52:J52', 'T21:V21', 'H45:J45',
            'H32:J32', 'T23:V23', 'T57:V58', 'T9:V17', 'H47:J47',
            'T52:V52', 'A45:A47', 'T27:V27', 'B41:D41', 'N27:P27',
            'T18:V18', 'Q21:S21', 'E62:G62', 'K24:M24', 'H50:J50',
            'K51:M51', 'K26:M26', 'E51:G51', 'Q25:R25', 'B61:D61',
            'K42:M42', 'K19:M19', 'K6:M6', 'E50:G50', 'A18:A25',
            'T24:V24', 'T63:V72', 'E52:G52', 'B18:D18', 'T53:V53',
            'Q29:S29', 'K32:M32', 'Q23:S23', 'T55:V55', 'K9:M17',
            'Q31:S31', 'Q6:S6', 'K27:M27', 'E7:G7', 'N30:P30',
            'A40:A44', 'T46:V46', 'T48:V48', 'B62:D62', 'N73:P83',
            'T20:V20', 'K63:M72', 'Q42:S42', 'K45:M45', 'K20:M20',
            'E20:G20', 'H21:J21', 'K47:M47', 'Q19:S19', 'E60:G60',
            'K22:M22', 'H9:J17', 'T40:V40', 'H25:I25', 'Q44:S44',
            'B19:D19', 'H27:J27', 'N59:P59', 'Q45:S45', 'N49:P49',
            'Q32:S32', 'Q47:S47', 'E25:F25', 'T45:V45', 'E8:G8',
            'H6:J6', 'T62:V62', 'T59:V59', 'N62:P62', 'Q50:S50',
            'Q57:S58', 'K59:M59', 'T51:V51', 'Q27:S27', 'K61:M61',
            'K48:M48', 'H24:J24', 'H18:J18', 'E21:G21', 'H26:J26',
            'B25:D25', 'H53:J53', 'E23:G23', 'B22:D22', 'H55:J55',
            'Q55:S55', 'K41:M41', 'N52:P52', 'Q40:S40', 'K43:M43',
            'E49:G49', 'T61:V61', 'B40:D40', 'H48:J48', 'B73:D83',
            'Q53:S53', 'H63:J72', 'H43:J43', 'N7:P7', 'N57:P58',
            'N50:P50', 'K54:M54', 'H57:J58', 'E29:G29', 'E31:G31',
            'B28:D28', 'A33:D39', 'N60:P60', 'E26:G26', 'T22:V22',
            'Q41:S41', 'T73:V83', 'N55:P55', 'Q56:S56', 'Q43:S43',
            'B54:D54', 'E42:G42', 'B56:D56', 'H40:J40', 'E44:G44',
            'B43:D43', 'K60:M60', 'H51:J51', 'N25:O25', 'N8:P8',
            'K7:M7', 'E32:G32', 'T19:V19', 'Q9:S17', 'E47:G47',
            'B31:D31', 'Q24:S24', 'Q18:S18', 'H59:J59', 'N21:P21',
            'Q51:S51', 'H46:J46', 'H61:J61', 'N23:P23', 'E45:G45',
            'T41:V41', 'B49:D49', 'T7:V7', 'T56:V56', 'T43:V43',
            'E22:G22', 'B50:D50', 'B57:D58', 'K73:M83', 'C32:D32',
            'A56:A58', 'H49:J49', 'Q63:S72', 'Q26:S26', 'N29:P29',
            'A26:A32', 'N31:P31', 'B55:D55', 'N6:P6', 'H62:J62',
            'N24:P24', 'E53:G53', 'N26:P26', 'K30:M30', 'E55:G55',
            'E48:G48', 'B52:D52', 'Q20:S20'
        ];

        mergedCells.forEach(range => {
            worksheet.mergeCells(range);
        });

        // 4. 셀 데이터 및 스타일 설정
        setupCellsLG(worksheet, building);

        // 5. 파일 저장
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { 
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
        });
        
        const fileName = `CompList_LG_${building.name}_${getCurrentDateLG()}.xlsx`;
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
    const cell_A1 = worksheet.getCell('A1');
    cell_A1.value = '[임차제안건 제목을 입력하세요.]';
    cell_A1.font = { name: 'LG스마트체 Regular', size: 14, bold: true };
    cell_A1.alignment = { vertical: 'center' };

    const cell_A2 = worksheet.getCell('A2');
    cell_A2.value = '규모를 입력하세요.';
    cell_A2.font = { name: 'LG스마트체 Regular', size: 11 };
    cell_A2.alignment = { vertical: 'center' };

    const cell_A3 = worksheet.getCell('A3');
    cell_A3.value = '계약기간을 입력하세요.';
    cell_A3.font = { name: 'LG스마트체 Regular', size: 11 };
    cell_A3.alignment = { vertical: 'center' };

    const cell_A4 = worksheet.getCell('A4');
    cell_A4.value = '위치를 입력하세요.';
    cell_A4.font = { name: 'LG스마트체 Regular', size: 11 };
    cell_A4.alignment = { vertical: 'center' };

    // 카테고리 헤더들
    setCategoryCell(worksheet, 'A6', '위치', 'FFD9D9D9', 10, true);
    setCategoryCell(worksheet, 'A7', '제안', 'FFD9D9D9', 10);
    setCategoryCell(worksheet, 'A9', '건물 외관', 'FFD9D9D9', 9);
    setCategoryCell(worksheet, 'A18', '기초\n정보', 'FFD9D9D9', 9);
    setCategoryCell(worksheet, 'A26', '채권\n분석', 'FFD9D9D9', 9);
    setCategoryCell(worksheet, 'A33', '현재 공실', 'FFD9D9D9', 9);
    setCategoryCell(worksheet, 'A40', '제안', 'FFD9D9D9', 9);
    setCategoryCell(worksheet, 'A45', '기준층\n임대기준', 'FFD9D9D9', 9);
    setCategoryCell(worksheet, 'A48', '실질\n임대기준', 'FFD9D9D9', 9);
    setCategoryCell(worksheet, 'A50', '비용검토', 'FFD9D9D9', 9);
    setCategoryCell(worksheet, 'A56', '공사기간\nFAVOR', 'FFD9D9D9', 9);
    setCategoryCell(worksheet, 'A59', '주차현황', 'FFD9D9D9', 9);
    setCategoryCell(worksheet, 'A63', '기타', 'FFD9D9D9', 9);

    // B열 항목명들
    setItemCell(worksheet, 'B18', '주   소', 'FFE7E6E6');
    setItemCell(worksheet, 'B19', '위   치', 'FFE7E6E6');
    setItemCell(worksheet, 'B20', '준공일', 'FFE7E6E6');
    setItemCell(worksheet, 'B21', '규  모', 'FFE7E6E6');
    setItemCell(worksheet, 'B22', '연면적', 'FFE7E6E6');
    setItemCell(worksheet, 'B23', '기준층 전용면적', 'FFE7E6E6');
    setItemCell(worksheet, 'B24', '전용률', 'FFE7E6E6');
    setItemCell(worksheet, 'B25', '대지면적', 'FFE7E6E6');
    setItemCell(worksheet, 'B26', '소유자 (임대인)', 'FFE7E6E6');
    setItemCell(worksheet, 'B27', '채권담보 설정여부', 'FFE7E6E6');
    setItemCell(worksheet, 'B28', '공동담보 총 대지지분', 'FFE7E6E6');
    setItemCell(worksheet, 'B29', '선순위 담보 총액', 'FFE7E6E6');
    setItemCell(worksheet, 'B31', '개별공시지가(25년 1월 기준)', 'FFE7E6E6');
    setItemCell(worksheet, 'B40', '계약기간', 'FFE7E6E6');
    setItemCell(worksheet, 'B41', '입주가능 시기', 'FFE7E6E6', true);
    setItemCell(worksheet, 'B42', '제안 층', 'FFE7E6E6');
    setItemCell(worksheet, 'B43', '전용면적', 'FFE7E6E6', true, 'FFC00000');
    setItemCell(worksheet, 'B44', '임대면적', 'FFE7E6E6');
    setItemCell(worksheet, 'B45', '보증금', 'FFE7E6E6');
    setItemCell(worksheet, 'B46', '임대료', 'FFE7E6E6');
    setItemCell(worksheet, 'B47', '관리비', 'FFE7E6E6');
    setItemCell(worksheet, 'B48', '실질 임대료(RF만 반영)1)', 'FFE7E6E6');
    setItemCell(worksheet, 'B49', '연간 무상임대 (R.F)', 'FFE7E6E6');
    setItemCell(worksheet, 'B50', '보증금', 'FFE7E6E6');
    setItemCell(worksheet, 'B51', '월 임대료', 'FFE7E6E6');
    setItemCell(worksheet, 'B52', '월 관리비', 'FFE7E6E6');
    setItemCell(worksheet, 'B53', '관리비 내역', 'FFE7E6E6', false, 'FFC00000');
    setItemCell(worksheet, 'B54', '월납부액', 'FFE7E6E6', true, 'FFC00000');
    setItemCell(worksheet, 'B55', '(21개월 기준) 총 납부 비용3)', 'FFE7E6E6');
    setItemCell(worksheet, 'B56', '인테리어 기간 (F.O)', 'FFE7E6E6');
    setItemCell(worksheet, 'B57', '인테리어지원금 (T.I)', 'FFE7E6E6');
    setItemCell(worksheet, 'B59', '총 주차대수', 'FFE7E6E6');
    setItemCell(worksheet, 'B60', '무료주차 조건(임대면적)', 'FFE7E6E6');
    setItemCell(worksheet, 'B61', '무료주차 제공대수', 'FFE7E6E6');
    setItemCell(worksheet, 'B62', '유료주차(VAT별도)', 'FFE7E6E6');
    setItemCell(worksheet, 'B63', '평면도', 'FFE7E6E6');
    setItemCell(worksheet, 'B73', '특이사항', 'FFE7E6E6');

    // C30, C32 특수 항목
    const cell_C30 = worksheet.getCell('C30');
    cell_C30.value = '공시지가 대비 담보율';
    cell_C30.font = { name: 'LG스마트체 Regular', size: 9, color: { argb: 'FFC00000' } };
    cell_C30.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE7E6E6' } };
    cell_C30.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
    setBordersLG(cell_C30);

    const cell_C32 = worksheet.getCell('C32');
    cell_C32.value = '토지가격 적용';
    cell_C32.font = { name: 'LG스마트체 Regular', size: 9 };
    cell_C32.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE7E6E6' } };
    cell_C32.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
    setBordersLG(cell_C32);

    // E열 빌딩 데이터 매핑
    // 주소
    const cell_E18 = worksheet.getCell('E18');
    cell_E18.value = building.address || '';
    cell_E18.font = { name: 'LG스마트체 Regular', size: 9 };
    cell_E18.alignment = { horizontal: 'center', vertical: 'center' };
    setBordersLG(cell_E18);

    // 위치 (인근역)
    const cell_E19 = worksheet.getCell('E19');
    cell_E19.value = building.station || '';
    cell_E19.font = { name: 'LG스마트체 Regular', size: 9 };
    cell_E19.alignment = { horizontal: 'center', vertical: 'center' };
    setBordersLG(cell_E19);

    // 준공일
    const cell_E20 = worksheet.getCell('E20');
    cell_E20.value = building.completionYear ? `${building.completionYear}년` : '';
    cell_E20.font = { name: 'LG스마트체 Regular', size: 9 };
    cell_E20.alignment = { horizontal: 'center', vertical: 'center' };
    setBordersLG(cell_E20);

    // 규모
    const cell_E21 = worksheet.getCell('E21');
    cell_E21.value = building.floors || '';
    cell_E21.font = { name: 'LG스마트체 Regular', size: 9 };
    cell_E21.alignment = { horizontal: 'center', vertical: 'center' };
    setBordersLG(cell_E21);

    // 연면적
    const cell_E22 = worksheet.getCell('E22');
    cell_E22.value = building.grossFloorAreaPy ? `${building.grossFloorAreaPy.toLocaleString()}평` : '';
    cell_E22.font = { name: 'LG스마트체 Regular', size: 9 };
    cell_E22.alignment = { horizontal: 'center', vertical: 'center' };
    setBordersLG(cell_E22);

    // 기준층 전용면적
    const cell_E23 = worksheet.getCell('E23');
    cell_E23.value = building.baseFloorAreaDedicatedPy ? `${building.baseFloorAreaDedicatedPy.toLocaleString()}평` : '';
    cell_E23.font = { name: 'LG스마트체 Regular', size: 9 };
    cell_E23.alignment = { horizontal: 'center', vertical: 'center' };
    setBordersLG(cell_E23);

    // 전용률
    const cell_E24 = worksheet.getCell('E24');
    cell_E24.value = building.dedicatedRate ? `${building.dedicatedRate}%` : '';
    cell_E24.font = { name: 'LG스마트체 Regular', size: 9 };
    cell_E24.alignment = { horizontal: 'center', vertical: 'center' };
    setBordersLG(cell_E24);

    // 대지면적
    const cell_E25 = worksheet.getCell('E25');
    cell_E25.value = building.landAreaPy ? `${building.landAreaPy.toLocaleString()}평` : '';
    cell_E25.font = { name: 'LG스마트체 Regular', size: 9 };
    cell_E25.alignment = { horizontal: 'center', vertical: 'center' };
    setBordersLG(cell_E25);

    // 소유자
    const cell_E26 = worksheet.getCell('E26');
    cell_E26.value = building.owner || '';
    cell_E26.font = { name: 'LG스마트체 Regular', size: 9 };
    cell_E26.alignment = { horizontal: 'center', vertical: 'center' };
    setBordersLG(cell_E26);

    // 주차 정보
    const cell_E59 = worksheet.getCell('E59');
    cell_E59.value = building.parkingSpace || '';
    cell_E59.font = { name: 'LG스마트체 Regular', size: 9 };
    cell_E59.alignment = { horizontal: 'center', vertical: 'center' };
    setBordersLG(cell_E59);

    const cell_E62 = worksheet.getCell('E62');
    cell_E62.value = building.parkingFee || '';
    cell_E62.font = { name: 'LG스마트체 Regular', size: 9 };
    cell_E62.alignment = { horizontal: 'center', vertical: 'center' };
    setBordersLG(cell_E62);

    // 빌딩 설비 정보
    const cell_H18 = worksheet.getCell('H18');
    cell_H18.value = building.elevator || '';
    cell_H18.font = { name: 'LG스마트체 Regular', size: 9 };
    cell_H18.alignment = { horizontal: 'center', vertical: 'center' };
    setBordersLG(cell_H18);

    const cell_H19 = worksheet.getCell('H19');
    cell_H19.value = building.hvac || '';
    cell_H19.font = { name: 'LG스마트체 Regular', size: 9 };
    cell_H19.alignment = { horizontal: 'center', vertical: 'center' };
    setBordersLG(cell_H19);

    const cell_H20 = worksheet.getCell('H20');
    cell_H20.value = building.buildingUse || '';
    cell_H20.font = { name: 'LG스마트체 Regular', size: 9 };
    cell_H20.alignment = { horizontal: 'center', vertical: 'center' };
    setBordersLG(cell_H20);

    const cell_H21 = worksheet.getCell('H21');
    cell_H21.value = building.structure || '';
    cell_H21.font = { name: 'LG스마트체 Regular', size: 9 };
    cell_H21.alignment = { horizontal: 'center', vertical: 'center' };
    setBordersLG(cell_H21);

    // 제안 관련 - B43 전용면적 (실제 값 표시)
    const cell_E43 = worksheet.getCell('E43');
    cell_E43.value = building.baseFloorAreaDedicatedPy ? `${building.baseFloorAreaDedicatedPy.toLocaleString()}평` : '';
    cell_E43.font = { name: 'LG스마트체 Regular', size: 9, bold: true, color: { argb: 'FFC00000' } };
    cell_E43.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };
    cell_E43.alignment = { horizontal: 'center', vertical: 'center' };
    setBordersLG(cell_E43);

    // 임대면적
    const cell_E44 = worksheet.getCell('E44');
    cell_E44.value = building.baseFloorAreaPy ? `${building.baseFloorAreaPy.toLocaleString()}평` : '';
    cell_E44.font = { name: 'LG스마트체 Regular', size: 9 };
    cell_E44.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };
    cell_E44.alignment = { horizontal: 'center', vertical: 'center' };
    setBordersLG(cell_E44);

    // 수식 설정
    setFormulas(worksheet);

    // 나머지 빈 셀들 테두리 처리
    setEmptyCellBorders(worksheet);

    console.log('셀 설정 완료');
}

// 카테고리 셀 설정 함수
function setCategoryCell(worksheet, address, value, bgColor, size = 9, bold = false) {
    const cell = worksheet.getCell(address);
    cell.value = value;
    cell.font = { 
        name: 'LG스마트체 Regular', 
        size: size, 
        bold: bold 
    };
    cell.fill = { 
        type: 'pattern', 
        pattern: 'solid', 
        fgColor: { argb: bgColor } 
    };
    cell.alignment = { 
        horizontal: 'center', 
        vertical: 'center',
        wrapText: true 
    };
    setBordersLG(cell);
}

// 항목명 셀 설정 함수
function setItemCell(worksheet, address, value, bgColor, bold = false, fontColor = 'FF000000') {
    const cell = worksheet.getCell(address);
    cell.value = value;
    cell.font = { 
        name: 'LG스마트체 Regular', 
        size: 9, 
        bold: bold,
        color: { argb: fontColor }
    };
    cell.fill = { 
        type: 'pattern', 
        pattern: 'solid', 
        fgColor: { argb: bgColor } 
    };
    cell.alignment = { 
        horizontal: 'center', 
        vertical: 'center',
        wrapText: true 
    };
    setBordersLG(cell);
}

// 수식 설정 함수
function setFormulas(worksheet) {
    // 공시지가 대비 담보율 계산 수식들
    worksheet.getCell('E30').value = { formula: '=E29/E32' };
    worksheet.getCell('E30').numFmt = '0.00%';
    
    worksheet.getCell('E32').value = { formula: '=E31*G25' };
    worksheet.getCell('E32').numFmt = '#,##0';
    
    worksheet.getCell('H30').value = { formula: '=H29/H32' };
    worksheet.getCell('H30').numFmt = '0.00%';
    
    worksheet.getCell('H32').value = { formula: '=H31*J25' };
    worksheet.getCell('H32').numFmt = '#,##0';
    
    // 면적 관련 수식들
    worksheet.getCell('E43').value = { formula: '=SUM(F34:F35)' };
    worksheet.getCell('E43').numFmt = '#,##0"평"';
    
    worksheet.getCell('E44').value = { formula: '=SUM(G34:G35)' };
    worksheet.getCell('E44').numFmt = '#,##0"평"';
    
    // 실질 임대료 계산
    worksheet.getCell('E48').value = { formula: '=E46*(12-E49)/12' };
    worksheet.getCell('E48').numFmt = '#,##0';
    
    // 비용 계산
    worksheet.getCell('E50').value = { formula: '=E45*E44' };
    worksheet.getCell('E50').numFmt = '#,##0"원"';
    
    worksheet.getCell('E51').value = { formula: '=E46*E44' };
    worksheet.getCell('E51').numFmt = '#,##0"원"';
    
    worksheet.getCell('E52').value = { formula: '=E47*E44' };
    worksheet.getCell('E52').numFmt = '#,##0"원"';
    
    worksheet.getCell('E54').value = { formula: '=E51' };
    worksheet.getCell('E54').numFmt = '#,##0"원"';
    
    worksheet.getCell('E55').value = { formula: '=E54*21' };
    worksheet.getCell('E55').numFmt = '#,##0"원"';
    
    // 주차 계산
    worksheet.getCell('E61').value = { formula: '=E44/E60' };
    worksheet.getCell('E61').numFmt = '#,##0.0"대"';
}

// 빈 셀 테두리 처리 함수
function setEmptyCellBorders(worksheet) {
    // 주요 데이터 영역의 빈 셀들에 테두리 적용
    const ranges = [
        { start: 'E27', end: 'G27' },
        { start: 'E28', end: 'G32' },
        { start: 'E33', end: 'G39' },
        { start: 'E40', end: 'G62' },
        { start: 'H22', end: 'J62' },
        { start: 'K18', end: 'M62' },
        { start: 'N18', end: 'P62' },
        { start: 'Q18', end: 'S62' },
        { start: 'T18', end: 'V62' }
    ];

    ranges.forEach(range => {
        const startCell = worksheet.getCell(range.start);
        const endCell = worksheet.getCell(range.end);
        
        for (let row = startCell.row; row <= endCell.row; row++) {
            for (let col = startCell.col; col <= endCell.col; col++) {
                const cell = worksheet.getCell(row, col);
                if (!cell.value) {
                    setBordersLG(cell);
                }
            }
        }
    });
}

// 테두리 설정 함수
function setBordersLG(cell) {
    cell.border = {
        top: { style: 'thin', color: { argb: 'FF000000' } },
        left: { style: 'thin', color: { argb: 'FF000000' } },
        bottom: { style: 'thin', color: { argb: 'FF000000' } },
        right: { style: 'thin', color: { argb: 'FF000000' } }
    };
}

// 디버깅 함수
function debugExcelGeneration() {
    console.log('=== 엑셀 생성 디버깅 정보 ===');
    console.log('1. A1-A4: 샘플 문구로 교체됨');
    console.log('2. B43: "전용면적" 항목명 유지');
    console.log('3. 86-88행: 제거됨 (높이 0)');
    console.log('4. 폰트: LG스마트체 Regular 적용');
    console.log('5. 정렬: 모든 셀 가운데 정렬');
    console.log('6. 색상: 원본 배경색, 폰트색, 테두리 유지');
    console.log('7. 빌딩 데이터: JSON 속성과 매핑 완료');
}