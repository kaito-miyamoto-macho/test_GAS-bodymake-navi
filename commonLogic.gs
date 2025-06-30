/**
 * コーチ用シートを取得
 * 
 * return コーチ用シート
 */
function GET_COACH_SHEET() {
  const targetSpreadsheetId = SPREAD_SHEET_IDS.PJTManagement; // PJT管理スプレッドシートのID
  const targetSpreadsheet = SpreadsheetApp.openById(targetSpreadsheetId);
  const coachSheet = targetSpreadsheet.getSheetByName(SHEET_NAMES_MANAGEMENT.COACH_LIST);

  if (!coachSheet) {
    Logger.log("❌ コーチ名簿シートが見つかりません");
    return null;
  }
  return coachSheet;
}

/**
 * クライアント用シートを取得
 * 
 * return クライアント用シート
 */
function GET_CLIENT_SHEET() {
  const targetSpreadsheetId = SPREAD_SHEET_IDS.PJTManagement; // PJT管理スプレッドシートのID
  const targetSpreadsheet = SpreadsheetApp.openById(targetSpreadsheetId);
  const clientSheet = targetSpreadsheet.getSheetByName(SHEET_NAMES_MANAGEMENT.CLIENT_LIST);

  if (!clientSheet) {
    Logger.log("❌ クライアント名簿シートが見つかりません");
    return null;
  }
  return clientSheet;
}

/**
 * 認証情報シートを取得
 * 
 * return 認証情報シート
 */
function GET_AUTH_SHEET() {
  const targetSpreadsheetId = SPREAD_SHEET_IDS.PJTManagement; // PJT管理スプレッドシートのID
  const targetSpreadsheet = SpreadsheetApp.openById(targetSpreadsheetId);
  const authSheet = targetSpreadsheet.getSheetByName(SHEET_NAMES_MANAGEMENT.AUTH_LIST);

  if (!authSheet) {
    Logger.log("❌ クライアント名簿シートが見つかりません");
    return null;
  }
  return authSheet;
}

/**
 * ログ出力用のシートを取得
 * @param sheetName シート名 (LOG_MANAGEMENT.MAIL_AUTHなどを引数に指定)
 * 
 * return logSheet 出力対象のシート
 */
// function GET_LOG_SHEET (sheetName) {
//   const targetSpreadsheetId = SPREAD_SHEET_IDS.LOG_LIST; // ログ管理スプレッドシートのID
//   const targetSpreadsheet = SpreadsheetApp.openById(targetSpreadsheetId);
//   const logSheet = targetSpreadsheet.getSheetByName(sheetName);

//     if (!logSheet) {
//     Logger.log(`❌ ${sheetName}が見つかりません`);
//     return null;
//   }
//   return logSheet;
// }

/**
 * ログ出力用のシートを取得
 * 
 * return logSeet 出力用のシート
 */
function GET_LOG_SHEET () {
  const targetSpreadsheetId = SPREAD_SHEET_IDS.PJTManagement; // PJT管理スプレッドシートのID
  const targetSpreadsheet = SpreadsheetApp.openById(targetSpreadsheetId);
  const logSheet = targetSpreadsheet.getSheetByName(SHEET_NAMES_MANAGEMENT.LOG_LIST);

    if (!logSheet) {
    Logger.log(`❌ ${sheetName}が見つかりません`);
    return null;
  }
  return logSheet;
}

/**
 * シートのヘッダーを取得
 * @param sheet シート本体
 * @param headerRow ヘッダー行番号 (サポートノート内のヘッダーは1行目とは限らないので注意)
 * 
 * return ヘッダー行
 */
function GET_HEADER (sheet, headerRow) {
  return sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
}

/**
 * シートのデータ部分(1行目のヘッダーを除いたデータ行)を取得
 * @param sheet シート本体
 * 
 * return データ部分
 */
function GET_All_DATA (sheet) {
  if (sheet.getLastRow() <= 1) {
    return [];
  }
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();;
}

/**
 * 入力データを取得
 * @param sheet 入力されたデータが溜まるシート
 * @param lastrow 入力対象列
 * 
 * return 入力されたデータ
 */
function GET_INPUT_DATA (sheet, lastRow) {
  return sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()[0];
}
