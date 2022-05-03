// スプレッドシート・各シートを定義
const ss = SpreadsheetApp.getActiveSpreadsheet();
const homeSheet = ss.getSheetByName('ホーム');
const mockUpSheet = ss.getSheetByName('モックアップ');
const databaseSheet = ss.getSheetByName('データベース');

// 今年の情報(例：2022)・曜日の配列を定義
const nowDate = new Date();
const year = nowDate.getFullYear();
// const dayOfWeekArr = ['日', '月', '火', '水', '木', '金', '土'];

// シフト表を生成する関数
function createShift() {
  // 選択した月を定義
  let choiceMonth;
  const tmpMonth = homeSheet.getRange('B7').getValue();

  switch (tmpMonth) {
    case '1月':
      choiceMonth = 1;
      break;
    case '2月':
      choiceMonth = 2;
      break;
    case '3月':
      choiceMonth = 3;
      break;
    case '4月':
      choiceMonth = 4;
      break;
    case '5月':
      choiceMonth = 5;
      break;
    case '6月':
      choiceMonth = 6;
      break;
    case '7月':
      choiceMonth = 7;
      break;
    case '8月':
      choiceMonth = 8;
      break;
    case '9月':
      choiceMonth = 9;
      break;
    case '10月':
      choiceMonth = 10;
      break;
    case '11月':
      choiceMonth = 11;
      break;
    case '12月':
      choiceMonth = 12;
      break;
  }

  // 選択した月を定義
  const team = homeSheet.getRange('B9').getValue();

  // 選択した月の情報を生成
  const createDate = new Date(year, choiceMonth - 1, 1);
  
  // シートを複製して名前を動的に変更
  const sheetName = `${year}年${choiceMonth}月${team}`;

  const sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    Browser.msgBox('すでに同じ名前のシートが存在します。');
    return;
  } else {
    const copySheet = mockUpSheet.copyTo(ss);
    copySheet.setName(sheetName);
  }

  // 複製したシートのタイトル(B1)を動的に変更
  const createSheet = ss.getSheetByName(sheetName);
  createSheet.getRange('B1').setValue(sheetName);

  // 日付を入力
  let monthRanges;
  let n;
  // 各月の末日を処理
  switch (choiceMonth) {
    case 1:
    case 3:
    case 5:
    case 7:
    case 8:
    case 10:
    case 12:
      monthRanges = createSheet.getRange('F1:AJ1');
      n = 0;
      break;
    
    case 2:
      monthRanges = createSheet.getRange('F1:AG1');
      n = 3;
      // 閏年の処理
      if (year % 4 == 0) {
        createSheet.deleteColumns(35, 2);
      } else {
        createSheet.deleteColumns(34, 3);
      }
      break;

    case 4:
    case 6:
    case 9:
    case 11:
      monthRanges = createSheet.getRange('F1:AI1');
      n = 1;
      createSheet.deleteColumn(36);
      break;
  }
  
  const monthVals = monthRanges.getValues();

  for (let i = 0; i < 1; i++) {
    for (let j = 0; j < 31 - n; j++) {
      monthVals[i][j] = new Date(year, choiceMonth - 1, j + 1);
    }
  }
  monthRanges.setValues(monthVals);

  // 閏年のとき3/29→2/29に修正
  if (year % 4 == 0 && choiceMonth == 2) {
    const correctDay = new Date(year, choiceMonth - 1, 29);
    const errorDay = createSheet.getRange('AH1');
    errorDay.setValue(correctDay);
  }

  // 休日の書式をクリアにして背景色をグレーに変更
  turnOffHolidayShift(sheetName);

  // 作成したシフト表をデータベースのシート履歴に追加
  addSheetNameToDatabase(sheetName);

  // 固定セルを編集できないように保護
  const protectRange1 = createSheet.getRange('A:E');
  const protectRange2 = createSheet.getRange('F1:AJ3');
  const protections1 = protectRange1.protect();
  const protections2 = protectRange2.protect();
  const userList1 = protections1.getEditors();
  const userList2 = protections2.getEditors();
  protections1.removeEditors(userList1);
  protections2.removeEditors(userList2);

  createSheet.showSheet();
}

// 休日かどうか判定する関数
function isHoliday(date) {
  // 土日の判定
  const day = date.getDay();
  if (day === 0 || day === 6) return true;

  // 祝日の判定
  const id = 'ja.japanese#holiday@group.v.calendar.google.com';
  const calender = CalendarApp.getCalendarById(id);
  const events = calender.getEventsForDay(date);
  if (events.length) return true;
}

// 休日を使用不可にする関数
function turnOffHolidayShift(sheetName) {
  const creaateSheet = ss.getSheetByName(sheetName);
  const lastRowNum = creaateSheet.getLastRow() - 1;
  const lastColumNum = creaateSheet.getLastColumn() - 5;
  const daysRanges = creaateSheet.getRange(1, 6, 1, lastColumNum);
  const daysVals = daysRanges.getValues();

  for (let i = 0; i < 1; i++) {
    for (let j = 0; j < lastColumNum; j++) {
      if (isHoliday(daysVals[i][j]) === true) {
        const tmpRanges = creaateSheet.getRange(4, j + 6, lastRowNum, 1);
        tmpRanges.clearFormat();
        tmpRanges.setBackground('#666666');
        tmpRanges.setFontColor('#ffffff');
      }
    }
  }
}

// データベースにシート履歴を残す関数
function addSheetNameToDatabase(sheetName) {
  const sheetLog = databaseSheet.getRange('E:E').getValues();
  
  // シート履歴列の最終行数を取得
  let lastLow = 0;
  for (let i = 0; i < sheetLog.length; i++) {
    if (sheetLog[i][0]) {
      lastLow++;
    }
  }

  // 作成したシート名をシート履歴に追加
  databaseSheet.getRange(lastLow + 1, 5).setValue(sheetName);
  
  // 入力規則の範囲を指定
  const tmpRange = databaseSheet.getRange(2, 5, lastLow, 1);
  // 入力規則を作成
  const rule = SpreadsheetApp.newDataValidation().requireValueInRange(tmpRange).build();
  const cell = homeSheet.getRange('E7');
  cell.setDataValidation(rule);
}

// 指定したシートを閲覧する関数
function showSheet() {
  const sheetName = homeSheet.getRange('E7').getValue();
  const activeSheet = ss.getSheetByName(sheetName);

  if (activeSheet) {
    console.log('aiueo');
    activeSheet.activate();
  } else {
    Browser.msgBox('対象のシートが存在しません。');
  }
  
}

