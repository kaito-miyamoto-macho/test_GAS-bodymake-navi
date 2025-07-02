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


/**
 * Gmailアドレスのみに編集権限を付与し、それ以外の権限を削除
 * 
 * @param clientEmail クライアントのメールアドレス
 * @param coachEmail コーチのメールアドレス
 * @param file 編集権限改訂対象のクライアントノート
 */ 
function allowClientNoteAccess(clientEmail, coachEmail, file) {
  const fileId = file.getId(); // ← これが必要

  // Gmailアドレスかつ、nullチェック通過したアドレス + GAS実行者を許可対象に
  const allowedEmails = [clientEmail, coachEmail, GAS_OWNER_EMAIL]
    .filter(email => email && email.endsWith("@gmail.com") || email === GAS_OWNER_EMAIL);

  // 現在の編集者を取得
  const currentEditors = file.getEditors().map(user => user.getEmail());

  // コーチとクライアント以外の編集者を削除
  currentEditors.forEach(email => {
    if (!allowedEmails.includes(email)) {
      try {
        file.removeEditor(email);
        Logger.log("不要な編集者を削除: " + email);
      } catch (e) {
        Logger.log("削除エラー: " + email + " - " + e.toString());
      }
    }
  });

  // Gmailアドレスのみに編集権限を付与（通知なしで）
  allowedEmails.forEach(email => {
    try {
      Drive.Permissions.create(
        {
          role: "writer",
          type: "user",
          emailAddress: email
        },
        fileId,
        { sendNotificationEmail: false } // ← v3ではここが単数形
      );
      Logger.log("✅ 通知なしで編集権限を付与: " + email);
    } catch (e) {
      Logger.log("⚠️ 編集権限の付与エラー: " + email + " - " + e.toString());
      return;
    }
  });
  // リンク共有を無効化（重要）
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.NONE);
}

/**
 * 積み上げレポートのPDFリンクを生成する
 */
function createPDFLink(spreadsheetId, sheetName) {
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(`シート ${sheetName} が見つかりません`);
    return '';
  }

  const sheetId = sheet.getSheetId();
  const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=pdf&size=A4&portrait=true&fitw=true&sheetnames=false&printtitle=false&gridlines=false&gid=${sheetId}`;
  return url;
}

/********************
 * 回答者IDまたはLINE名を基準に更新対象の行を取得
 *********************/
function findTargetRowByIdOrLineName(targetClientSheet, sourceHeaders, sourceData, targetHeaders) {
  Logger.log("🔹 findTargetRowByIdOrLineName 関数開始");

  const existingData = targetClientSheet.getDataRange().getValues();
  Logger.log(`📌 取得済みのシートデータの行数: ${existingData.length}`);

  const lineNameColIndex = targetHeaders.indexOf('LINE名');
  const answerIdColIndex = targetHeaders.indexOf('回答者ID');

  Logger.log(`📌 LINE名の列インデックス: ${lineNameColIndex}`);
  Logger.log(`📌 回答者IDの列インデックス: ${answerIdColIndex}`);

  if (lineNameColIndex === -1 || answerIdColIndex === -1) {
    Logger.log("❌ LINE名または回答者IDの列が見つかりません");
    return null;
  }

  const targetLineName = (sourceData[sourceHeaders.indexOf('回答者名')] || "").toString().trim();
  const targetAnswerId = (sourceData[sourceHeaders.indexOf('回答者ID')] || "").toString().trim();

  Logger.log(`🔍 照合対象: LINE名=${targetLineName}, 回答者ID=${targetAnswerId}`);

  for (let i = 1; i < existingData.length; i++) {
    const rowLineName = (existingData[i][lineNameColIndex] || "").toString().trim();
    const rowAnswerId = (existingData[i][answerIdColIndex] || "").toString().trim();

    if (rowLineName === targetLineName && rowAnswerId === targetAnswerId) {
      Logger.log(`✅ 完全一致する行を発見: ${i + 1}`);
      return i + 1;
    }

    if (rowLineName === targetLineName && rowAnswerId === "") {
      Logger.log(`✅ LINE名は一致し、回答者IDが未記入: ${i + 1}`);
      return i + 1;
    }
  }

  Logger.log("⚠️ 該当する行が見つからず、新規行を作成");
  return null;
}

/**
 * 認証管理シートから回答者IDに基づいて情報を取得
 */
function findAuthInfoByAnswerId(answerId, userType) {
  Logger.log("🔹 findAuthInfoByAnswerId 関数開始");
  Logger.log(`回答者ID: ${answerId}を元に対象データを取得`)
  const authTBL = GET_AUTH_SHEET();
  const authHeader = GET_HEADER (authTBL, 1);
  const authAllData = GET_All_DATA(authTBL);

  const authAnswerIdColIndex = authHeader.indexOf(AUTH_LIST_TBL.ANSWER_ID);
  const authUserTypeColIndex = authHeader.indexOf(AUTH_LIST_TBL.USER_TYPE);
  const authDeleteFlgColIndex = authHeader.indexOf(AUTH_LIST_TBL.DELETE_FLG);

  for (let i = 1; i < authAllData.length; i++) {
    if (authAllData[i][authAnswerIdColIndex] === answerId
        && authAllData[i][authUserTypeColIndex] === userType
        && authAllData[i][authDeleteFlgColIndex] !== DELETE_FLG.ON ) {
      Logger.log(`回答者IDを元に取得したデータ`);
      Logger.log(`メールアドレス:${authAllData[i][authHeader.indexOf(AUTH_LIST_TBL.MAIL)]}`);
      Logger.log(`LINE ID:${authAllData[i][authHeader.indexOf(AUTH_LIST_TBL.LINE_ID)]}`);
      Logger.log(`LINE 登録名:${authAllData[i][authHeader.indexOf(AUTH_LIST_TBL.LINE_NAME)]}`);
      return {
        email: authAllData[i][authHeader.indexOf(AUTH_LIST_TBL.MAIL)],
        lineId: authAllData[i][authHeader.indexOf(AUTH_LIST_TBL.LINE_ID)],
        lineName: authAllData[i][authHeader.indexOf(AUTH_LIST_TBL.LINE_NAME)]
      };
    }
  }
  Logger.log("🔹 findAuthInfoByAnswerId 関数終了");
  return null;
}

/**
 * 日付を yyyy/MM/dd 形式に整形
 */
function formatDateToYYYYMMDD(rawDateValue) {
  if (!rawDateValue) return "";

  try {
    const dateValue = new Date(rawDateValue.replace(/^(?:'|")|(?:'|")$/g, '').replace(/(am|pm)$/i, " $1").toUpperCase());
    return Utilities.formatDate(dateValue, Session.getScriptTimeZone(), 'yyyy/MM/dd');
  } catch (e) {
    Logger.log(`日付フォーマットエラー: ${rawDateValue}`);
    return "";
  }
}
