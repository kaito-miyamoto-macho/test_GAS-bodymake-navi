
// /**
//  * 一時的なメッセージ送信関数
//  *  */ 
// function temporaryMsg() {
//   const ID = ""
//   const msg =``
//   sendLINEMessage(ID,msg);
//   sendLINEMessage("U368ae7fe64961dcb012d71d6576e6cb8",msg); // テスト用にユウジに送る
// }

/**
 * 指定した LINE ユーザー ID にメッセージを送信する
 * @param {string} userId - 送信先の LINE ID
 * @param {string} message - 送信するメッセージ内容
 * @return {object} APIのレスポンス
 */
function sendLINEMessage(userId, message) {
    const debugsheet = GET_LOG_SHEET();
    
    const url = "https://api.line.me/v2/bot/message/push";

    const payload = {
        to: userId,
        messages: [{ type: "text", text: message }],
    };

    Logger.log("送信対象userId: " + userId);


    const options = {
        method: "post",
        headers: {
            "Content-Type": "application/json",
            "Authorization": "Bearer " + LINE_ACCESS_TOKEN,
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true, // エラーレスポンスも取得できるようにする
    };

    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    Logger.log("LINE API Response Code: " + responseCode);
    Logger.log("LINE API Response: " + responseText);

    debugsheet.appendRow([new Date(), `LINE API Response Code: ${responseCode}`]);
    debugsheet.appendRow([new Date(), `LINE API Response: ${responseText}`]);

    return JSON.parse(responseText);
}

/**
 * スプレッドシートから LINE ID を取得し、メッセージを送信する
 */
function sendMessagesToUsers() {
    const sheet = SpreadsheetApp.openById("YOUR_SPREADSHEET_ID") // ⭐ スプレッドシートのIDを設定
        .getSheetByName("顧客名簿"); // ⭐ 顧客リストがあるシート名を設定

    const data = sheet.getDataRange().getValues();
    const header = data[0];

    // ⭐ 「LINE ID」と「名前」の列を取得
    const lineIdColumn = header.indexOf("LINE ID");
    const nameColumn = header.indexOf("LINE名");

    if (lineIdColumn === -1 || nameColumn === -1) {
        Logger.log("Error: 必要な列が見つかりません");
        return;
    }

    // ⭐ 2行目以降のデータを処理
    for (let i = 1; i < data.length; i++) {
        const lineId = data[i][lineIdColumn];
        const name = data[i][nameColumn];

        if (lineId) {
            const message = `こんにちは、${name}さん！\n公式LINEからのメッセージです。`;
            const result = sendLineMessage(lineId, message);
            Logger.log(`送信結果: ${JSON.stringify(result)}`);
        } else {
            Logger.log(`LINE ID が見つかりません: ${name}`);
        }
    }
}


/**
 * コーチNo.からコーチのLINE IDを取得する関数（コーチ名簿シートを固定）
 * @returns {String} コーチのLINE ID（見つからない場合は空文字）
 */
function getCoachLineId(coachNumber) {
    const targetSpreadsheetId = SPREAD_SHEET_IDS.PJTManagement; // PJT管理シートのID
    const targetSpreadsheet = SpreadsheetApp.openById(targetSpreadsheetId);
    const targetCoachSheet = targetSpreadsheet.getSheetByName(SHEET_NAMES_MANAGEMENT.COACH_LIST);

    if (!targetCoachSheet) {
        Logger.log("⚠️ コーチ名簿シートが見つかりません");
        return "";
    }

    // コーチ名簿のヘッダーを取得
    const coachHeaders = targetCoachSheet.getRange(1, 1, 1, targetCoachSheet.getLastColumn()).getValues()[0];

    // 必要な列のインデックスを取得
    const coachNoColIndex = coachHeaders.indexOf('コーチNo.');
    const coachLineIdColIndex = coachHeaders.indexOf('LINE ID');

    if (coachNoColIndex === -1 || coachLineIdColIndex === -1) {
        Logger.log("⚠️ コーチNo.またはLINE IDの列が見つかりません");
        return "";
    }

    // コーチ名簿のデータを取得
    const coachData = targetCoachSheet.getRange(2, 1, targetCoachSheet.getLastRow() - 1, targetCoachSheet.getLastColumn()).getValues();

    // コーチNo.が一致する行を検索
    for (let i = 0; i < coachData.length; i++) {
        if (coachData[i][coachNoColIndex] == coachNumber) {
            const coachLineId = coachData[i][coachLineIdColIndex];
            Logger.log(`✅ コーチNo.${coachNumber} に対応するLINE IDを取得: ${coachLineId}`);
            return coachLineId;
        }
    }

    Logger.log(`⚠️ コーチNo.${coachNumber} に対応するLINE IDが見つかりません`);
    return "";
}


