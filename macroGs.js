const menuItems = ['main', 'drink', 'sweets'];

const formatDayTime = (value, format) => Utilities.formatDate(value,
    SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), format);

const setListDetail = () => {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    // const activeSheet = spreadsheet.getActiveSheet();
    const activeSheet = spreadsheet.getSheetByName('test');

    const allValues = activeSheet.getDataRange().getValues();

    const days = activeSheet.getRange('C3:K3').getValues()[0].filter(item => item !== '');
    const times = activeSheet.getRange('C4:K4').getValues()[0].filter(item => item !== '');

    // 商品一覧の開始行
    const startRowNumber = allValues[0][4];
    // E
    const startAlphabet = 4 + 65;
    // 商品の種類
    const menuList = { 'main': [], 'drink': [], 'sweets': [] };
    const headerMenuItems = [];
    let countAlphabet = 0;
    allValues.forEach((value, i) => {
        if (menuItems.includes(value[0])) {
            headerMenuItems.push(`=$B$${i + 1}`);
            // 例）=$B$22, 70, $C$22 【商品名, アルファベットに変換するための番号, 商品の値段】
            menuList[value[0]].push([`=$B$${i + 1}`, startAlphabet + countAlphabet, `$C$${i + 1}`]);
            countAlphabet++;
        };
    });
    const createInitialPrace = (items) => [...Array(items.length)].map(() => 0);
    const createInitialEmptyData = (items) => [...Array(items.length - 2)].map(item => '');
    const getInsertDate = (day, time) => formatDayTime(day, 'YYYY/MM/dd') + ' ' + formatDayTime(time, 'hh:mm')
    const getSumMainPrace = (rowNumber) => menuList.main.length > 0 ? menuList.main.map((item, i) =>
        `(${item[2]}*$${String.fromCodePoint(item[1])}$${rowNumber})`).join('+') : 0;

    const getSumDrinkPrace = (rowNumber) => menuList.drink.length > 0 ? menuList.drink.map((item, i) =>
        `(${item[2]}*$${String.fromCodePoint(item[1])}$${rowNumber})`).join('+') : 0;

    const getSumSweetsPrace = (rowNumber) => menuList.sweets.length > 0 ? menuList.sweets.map((item, i) =>
        `(${item[2]}*$${String.fromCodePoint(item[1])}$${rowNumber})`).join('+') : 0;

    const getONlySumSweetsPrace = (rowNumber) => menuList.sweets.length > 0 ? menuList.sweets.map((item, i) =>
        `(${item[2]}*$${String.fromCodePoint(item[1] - (menuList.main.length + menuList.drink.length + 1))}$${rowNumber})`).join('+') : 0;

    const getSumIndividualPrace = (rowNumber) => `SUM($${String.fromCodePoint(startAlphabet + headerMenuItems.length)}$${rowNumber}:
    $${String.fromCodePoint(startAlphabet + headerMenuItems.length + 2)}$${rowNumber})`;

    const headerMenuList = ['誰', 'リピート', '備考', ...headerMenuItems, 'メイン合計', 'ドリンク合計', 'お菓子合計', '個人合計', '部合計'];
    const totalPartsAlphabet = String.fromCodePoint(startAlphabet + headerMenuList.length - 5);
    // 日付が入るので初期値に1行加算
    let targetRowNumberForList = startRowNumber + 1;
    const totalParts = `=SUM($${totalPartsAlphabet}$${targetRowNumberForList + 1}:$${totalPartsAlphabet}$${targetRowNumberForList + 6})`;
    const insertDataForSixPeople = [];
    days.forEach((day, dayIndex) => {
        times.forEach((time, timeIndex) => {
            const insertDate = [getInsertDate(day, time), ...createInitialEmptyData(headerMenuList), totalParts];
            const initialList = [...Array(6)].map((item, i) => [
                '',
                `=IFERROR(VLOOKUP($B$${targetRowNumberForList + 1 + i},'名簿'!$B$2:$F$103,2,false),'該当者無')`,
                `=IFERROR(VLOOKUP($B$${targetRowNumberForList + 1 + i},'名簿'!$B$2:$F$103,2,false),'該当者無')`,
                ...createInitialPrace(headerMenuItems),
                '=' + getSumMainPrace(targetRowNumberForList + 1 + i),
                '=' + getSumDrinkPrace(targetRowNumberForList + 1 + i),
                '=' + getSumSweetsPrace(targetRowNumberForList + 1 + i),
                '=' + getSumIndividualPrace(targetRowNumberForList + 1 + i),
                '']);
            initialList.unshift(insertDate);
            insertDataForSixPeople.push(initialList);
            targetRowNumberForList += initialList.length;
        });
    });
    const insertMenuList = [headerMenuList, ...insertDataForSixPeople.flat()];
    activeSheet.getRange(startRowNumber, 2, insertMenuList.length, headerMenuList.length).setValues(insertMenuList);

    const headerSweetsList = ['誰', '備考', ...menuList.sweets.map(item => item[0]), '個人合計', '部合計'];
    // 2行空けるため
    targetRowNumberForList += 2
    const totalPartsSweetsAlphabet = String.fromCodePoint(startAlphabet + menuList.sweets.length - 1);

    const totalSweetsParts = `=SUM($${totalPartsSweetsAlphabet}$${targetRowNumberForList + 2}:$${totalPartsSweetsAlphabet}$${targetRowNumberForList + 6})`;
    const insertDataSweetsForSixPeople = [];
    days.forEach((day, dayIndex) => {
        times.forEach((time, timeIndex) => {
            const insertDateSweets = [getInsertDate(dat, time), ...createInitialEmptyData(headerSweetsList), totalSweetsParts];
            const initialListSweets = [...Array(5)].map((item, i) => [
                '',
                '',
                ...createInitialPrace(menuList.sweets),
                '=' + getONlySumSweetsPrace(targetRowNumberForList + 2 + i),
                '']);
            initialListSweets.unshift(insertDateSweets);
            insertDataSweetsForSixPeople.push(initialListSweets);
            targetRowNumberForList += initialListSweets.length;
        });
    });

    const insertSweetsData = [headerSweetsList, ...insertDataSweetsForSixPeople.flat()];
    activeSheet.getRange(startRowNumber + insertMenuList.length + 2, 2, insertSweetsData.length,
        headerSweetsList.length).setValues(insertSweetsData);
};