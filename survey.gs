function processSurveyResponse() {
    const responseSS = SpreadsheetApp.getActiveSpreadsheet(); // アンケート結果シートのスプレッドシート
    const responseSheet = responseSS.getSheetByName("シート1"); // アンケート結果シート
    const pjtSS = SpreadsheetApp.openById(SPREAD_SHEET_IDS.PJTManagement); // プロジェクト管理スプレッドシート
    const customerSheet = pjtSS.getSheetByName(SHEET_NAMES_MANAGEMENT.CLIENT_LIST);

    if (!responseSheet || !customerSheet) {
        Logger.log("⚠️ 指定されたシートが見つかりません");
        return;
    }

    // 最新の回答データを取得
    const lastRow = responseSheet.getLastRow();
    if (lastRow < 2) return; // 1行目（ヘッダー）は無視

    const headers = responseSheet.getRange(1, 1, 1, responseSheet.getLastColumn()).getValues()[0]; // 1行目の項目名
    const responseData = responseSheet.getRange(lastRow, 1, 1, responseSheet.getLastColumn()).getValues()[0]; // 最新の回答

    // 「回答者ID」の列を取得
    const responseIdIndex = headers.indexOf(ELME_HIDDEN_COL.ANSWER_ID);
    if (responseIdIndex === -1) {
        Logger.log("⚠️ 回答者IDの列が見つかりません");
        return;
    }

    const responseId = responseData[responseIdIndex];
    if (!responseId) {
        Logger.log("⚠️ 無効な回答ID");
        return;
    }

    // クライアント名簿から回答者IDに対応するクライアントノートのURLを取得
    const customerData = customerSheet.getDataRange().getValues();
    const customerHeaders = customerData[0];
    const customerIdIndex = customerHeaders.indexOf(CLIENT_LIST_TBL.ANSWER_ID);
    const noteIndex = customerHeaders.indexOf(CLIENT_LIST_TBL.CLIENT_NOTE_URL);

    if (customerIdIndex === -1 || noteIndex === -1) {
        Logger.log("⚠️ クライアント名簿に必要な列がありません");
        return;
    }

    let clientNoteURL = "";
    for (let i = 1; i < customerData.length; i++) {
        if (customerData[i][customerIdIndex] === responseId) {
            clientNoteURL = customerData[i][noteIndex];
            break;
        }
    }

    if (!clientNoteURL) {
        Logger.log(`⚠️ クライアント ${responseId} のクライアントノートのURLが見つかりません`);
        return;
    }

    // クライアントノートのスプレッドシートを開く
    const clientSheetId = extractSpreadsheetId(clientNoteURL);
    if (!clientSheetId) {
        Logger.log(`⚠️ クライアント ${responseId} のクライアントノートURLが無効です`);
        return;
    }

    const clientSpreadsheet = SpreadsheetApp.openById(clientSheetId);
    const habitSheet = clientSpreadsheet.getSheetByName(SHEET_NAMES_CLIENTSUPPORT.MONTHLY_HABIT_CHECK_LIST);

    if (!habitSheet) {
        Logger.log(`⚠️ クライアント ${responseId} の習慣化ﾁｪｯｸﾘｽﾄシートが見つかりません`);
        return;
    }

    // B列の評価項目を取得
    const habitItems = habitSheet.getRange(6, 2, habitSheet.getLastRow() - 5, 1).getValues().flat();

    // 最新の空いている列を探す
    const latestColumn = findNextAvailableColumn(habitSheet, 6);
    if (latestColumn === -1) {
        Logger.log(`⚠️ クライアント ${responseId} の習慣化ﾁｪｯｸﾘｽﾄに空きがありません`);
        return;
    }

    // 回答データをB列の項目と照合して記入
    let writeCount = 0;
    for (let i = 0; i < headers.length; i++) {
        if (["回答ID", "回答日時", "回答者名"].includes(headers[i])) continue; // 除外リスト
        const rowIndex = habitItems.indexOf(headers[i]);
        if (rowIndex !== -1) {
            habitSheet.getRange(rowIndex + 6, latestColumn).setValue(responseData[i]);
            writeCount++;
        }
    }

    if (writeCount === 0) {
        Logger.log(`⚠️ クライアント ${responseId} に対応する習慣化ﾁｪｯｸﾘｽﾄの項目が見つかりません`);
        return;
    }

    Logger.log(`✅ クライアント ${responseId} に ${writeCount} 件のデータを記入（列: ${latestColumn}）`);
}

/**
 * GoogleスプレッドシートのURLからスプレッドシートIDを抽出
 */
function extractSpreadsheetId(url) {
    const match = url.match(/[-\w]{25,}/);
    return match ? match[0] : null;
}

/**
 * 指定した列範囲（C列からI列）内で、最新の空いている列を見つける
 */
function findNextAvailableColumn(sheet, startRow) {
    const lastColumn = 9; // I列
    const startColumn = 3; // C列
    const rowData = sheet.getRange(startRow, startColumn, 1, lastColumn - startColumn + 1).getValues()[0];

    for (let i = 0; i < rowData.length; i++) {
        if (!rowData[i]) {
            return startColumn + i;
        }
    }
    return -1;
}
