const menuItems = ['main', 'drink', 'sweets'];

const formatDayTime = (value, format) => Utilities.formatDate(value,
    SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), format);

const formatAlphabet = (number) => String.fromCodePoint(number);

const createInitialPrace = (items) => [...Array(items.length)].map(() => 0);

const createInitialEmptyData = (items) => [...Array(items.length - 2)].map(item => '');

const getInsertDate = (day, time) => formatDayTime(day, 'YYYY/MM/dd') + ' ' + formatDayTime(time, 'hh:mm');

const getSumPrace = (items, rowNumber) => items.length > 0 ? '=' + items.map(item =>
    `(${item[2]}*$${formatAlphabet(item[1])}$${rowNumber})`).join('+') : 0;

const writeDetailPeople = (rowNNumber) => `=IFERROR(VLOOKUP($B$${rowNNumber},'名簿'!$B$2:$F$103,2,false),'該当者無')`;

const setBorderStyle = (func) => func.setBorder(true, true, true, true, true, true, 'black', SpreadsheetApp.BorderStyle.SOLID);

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

    const getOnlySumSweetsPrace = (rowNumber) => menuList.sweets.length > 0 ? '=' + menuList.sweets.map((item, i) =>
        `(${item[2]}*$${formatAlphabet(item[1] - (menuList.main.length + menuList.drink.length + 1))}$${rowNumber})`).join('+') : 0;

    const getSumIndividualPrace = (rowNumber) => `=SUM($${formatAlphabet(startAlphabet + headerMenuItems.length)}$${rowNumber}:
      $${formatAlphabet(startAlphabet + headerMenuItems.length + 2)}$${rowNumber})`;

    const headerMenuList = ['誰', 'リピート', '備考', ...headerMenuItems, 'メイン合計', 'ドリンク合計', 'お菓子合計', '個人合計', '部合計'];
    const totalPartsAlphabet = formatAlphabet(startAlphabet + headerMenuList.length - 5);
    // 日付が入るので初期値に1行加算
    let targetRowNumberForList = startRowNumber + 1;
    const totalParts = `=SUM($${totalPartsAlphabet}$${targetRowNumberForList + 1}:$${totalPartsAlphabet}$${targetRowNumberForList + 6})`;
    const insertDataForSixPeople = [];
    days.forEach(day => {
        times.forEach(time => {
            const insertDate = [getInsertDate(day, time), ...createInitialEmptyData(headerMenuList), totalParts];
            const initialList = [...Array(6)].map((item, i) => [
                '',
                writeDetailPeople(targetRowNumberForList + 1 + i),
                writeDetailPeople(targetRowNumberForList + 1 + i),
                ...createInitialPrace(headerMenuItems),
                getSumPrace(menuList.main, targetRowNumberForList + 1 + i),
                getSumPrace(menuList.drink, targetRowNumberForList + 1 + i),
                getSumPrace(menuList.sweets, targetRowNumberForList + 1 + i),
                getSumIndividualPrace(targetRowNumberForList + 1 + i),
                '']);
            initialList.unshift(insertDate);
            insertDataForSixPeople.push(initialList);
            targetRowNumberForList += initialList.length;
        });
    });
    const insertMenuList = [headerMenuList, ...insertDataForSixPeople.flat()];

    setBorderStyle(activeSheet.getRange(startRowNumber, 2, insertMenuList.length, headerMenuList.length).setValues(insertMenuList))
    activeSheet.getRange(`$B$${startRowNumber}:$${formatAlphabet(65 + headerMenuList.length)}$${startRowNumber}`).setBackground('yellow');

    for (let i = 1; i < insertDataForSixPeople.flat().length; i += 7) {
        activeSheet.getRange(`$B$${startRowNumber + i}:$${formatAlphabet(65 + headerMenuList.length)}$${startRowNumber + i}`).setBackground('orange');
    };

    const headerSweetsList = ['誰', '備考', ...menuList.sweets.map(item => item[0]), '個人合計', '部合計'];
    // 2行空けるため
    targetRowNumberForList += 2
    const totalPartsSweetsStartRow = targetRowNumberForList;
    const totalPartsSweetsAlphabet = formatAlphabet(startAlphabet + menuList.sweets.length - 1);

    let insertDataSweetsForSixPeople = [];
    days.forEach(day => {
        times.forEach(time => {
            const totalSweetsParts = `=SUM($${totalPartsSweetsAlphabet}$${targetRowNumberForList + 2}:$${totalPartsSweetsAlphabet}$${targetRowNumberForList + 6})`;
            const insertDateSweets = [getInsertDate(day, time), ...createInitialEmptyData(headerSweetsList), totalSweetsParts];
            const initialListSweets = [...Array(5)].map((item, i) => [
                '',
                '',
                ...createInitialPrace(menuList.sweets),
                getOnlySumSweetsPrace(targetRowNumberForList + 2 + i),
                '']);
            initialListSweets.unshift(insertDateSweets);
            insertDataSweetsForSixPeople.push(initialListSweets);
            targetRowNumberForList += initialListSweets.length;
        });
    });

    insertDataSweetsForSixPeople = [...insertDataSweetsForSixPeople.flat()];

    const insertSweetsData = [headerSweetsList, ...insertDataSweetsForSixPeople];
    const startRowFoeSweets = startRowNumber + insertMenuList.length + 2;
    const lastAlphabetForSweets = formatAlphabet(65 + headerSweetsList.length);

    setBorderStyle(activeSheet.getRange(startRowFoeSweets, 2, insertSweetsData.length, headerSweetsList.length).setValues(insertSweetsData));
    activeSheet.getRange(`$B$${startRowFoeSweets}:$${lastAlphabetForSweets}$${startRowFoeSweets}`).setBackground('yellow');

    for (let i = 1; i < insertDataSweetsForSixPeople.length; i += 6) {
        activeSheet.getRange(`$B$${startRowFoeSweets + i}:$${lastAlphabetForSweets}$${startRowFoeSweets + i}`).setBackground('orange');
    };

    const sumMainPrace = startAlphabet + headerMenuItems.length;
    const sumDrinkPrace = startAlphabet + headerMenuItems.length + 1;
    const sumSweetsPrace = startAlphabet + headerMenuItems.length + 2;

    const writeTotalPrace = (targetPrace) =>
        `=SUM($${formatAlphabet(targetPrace)}$${startRowNumber}:$${formatAlphabet(targetPrace)}$${startRowNumber + insertMenuList.length - 1})`;

    const startAlphabetForOnlySweets = ormatAlphabet(startAlphabet + menuList.sweets.length);

    activeSheet.getRange(6, 3).setValue(writeTotalPrace(sumMainPrace));
    activeSheet.getRange(7, 3).setValue(writeTotalPrace(sumDrinkPrace));
    activeSheet.getRange(8, 3).setValue(writeTotalPrace(sumSweetsPrace)`
    + SUM($${startAlphabetForOnlySweets}$${totalPartsSweetsStartRow + 1}:$${startAlphabetForOnlySweets}$${totalPartsSweetsStartRow + insertDataSweetsForSixPeople.length})`);
};