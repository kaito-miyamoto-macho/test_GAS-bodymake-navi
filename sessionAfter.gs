function transferSessionAfterData() {
    const settingsSS = SpreadsheetApp.getActiveSpreadsheet(); // システム設定用スプレッドシート
    const outputSS = SpreadsheetApp.openById(SPREAD_SHEET_IDS.SESSION_AFTER); // スプレッドシートのIDをオープン
    const pjtSS = SpreadsheetApp.openById(SPREAD_SHEET_IDS.PJTManagement);

    const outputSheet = outputSS.getSheetByName("シート1");
    const customerSheet = pjtSS.getSheetByName(SHEET_NAMES_MANAGEMENT.CLIENT_LIST);

    if (!outputSheet || !customerSheet) {
        Logger.log("指定されたシートが見つかりません");
        return;
    }

    // 「定期セッション後」のデータを取得
    const outputData = outputSheet.getDataRange().getValues();
    if (outputData.length < 2) return; // データがない場合は終了

    const outputHeaders = outputData[0]; // ヘッダー（1行目）
    const lastRow = outputData[outputData.length - 1]; // 最新の回答

    // 「回答者ID」の列を取得
    const responderIdIndex = outputHeaders.indexOf("回答者ID");
    if (responderIdIndex === -1) {
        Logger.log("回答者IDの列が見つかりません");
        return;
    }
    const responderId = lastRow[responderIdIndex];

    // 「顧客名簿」から回答者IDに対応するクライアントノートのURLを取得
    const customerData = customerSheet.getDataRange().getValues();
    const customerHeaders = customerData[0]; // 顧客名簿のヘッダー（1行目）
    const customerIndex = customerHeaders.indexOf("回答者ID");
    const noteIndex = customerHeaders.indexOf("クライアントノート");

    if (customerIndex === -1 || noteIndex === -1) {
        Logger.log("顧客名簿に必要な列がありません");
        return;
    }

    let clientNoteURL = "";
    for (let i = 1; i < customerData.length; i++) {
        if (customerData[i][customerIndex] === responderId) {
            clientNoteURL = customerData[i][noteIndex];
            break;
        }
    }

    if (!clientNoteURL) {
        Logger.log("該当するクライアントノートのURLが見つかりません");
        return;
    }

    // クライアントノートのスプレッドシートを開く
    const archiveSS = SpreadsheetApp.openByUrl(clientNoteURL);
    const archiveSheet = archiveSS.getSheetByName("ｾｯｼｮﾝｱｰｶｲﾌﾞ");

    if (!archiveSheet) {
        Logger.log("ｾｯｼｮﾝｱｰｶｲﾌﾞシートが見つかりません");
        return;
    }

    // 「セッション日」を取得し、日付を標準化
    const sessionDateIndex = outputHeaders.indexOf("■セッション日");
    if (sessionDateIndex === -1) {
        Logger.log("セッション日の列が見つかりません");
        return;
    }
    const sessionDate = Utilities.formatDate(new Date(lastRow[sessionDateIndex]), Session.getScriptTimeZone(), "yyyy/MM/dd");

    // 「セッションアーカイブ」内のデータを取得
    const archiveData = archiveSheet.getDataRange().getValues();
    let targetRow = -1;

    // **セッションアーカイブの既存データをループして、同じ日付があるか確認**
    for (let i = 2; i < archiveData.length; i++) {  // 3行目以降を検索
        const archiveDate = Utilities.formatDate(new Date(archiveData[i][0]), Session.getScriptTimeZone(), "yyyy/MM/dd");

        if (archiveDate === sessionDate) { // **標準化した日付で比較**
            targetRow = i + 1; // **スプレッドシートの行番号**
            break;
        }
    }

    // **セッションアーカイブの項目名（2行目）を取得**
    const archiveHeaders = archiveSheet.getRange(2, 1, 1, archiveSheet.getLastColumn()).getValues()[0];

    // **既存行がなければ新規行を作成**
    if (targetRow === -1) {
        targetRow = archiveSheet.getLastRow() + 1;
        archiveSheet.getRange(targetRow, 1).setValue(sessionDate); // **1列目にセッション日を追加**
    }

    // **「定期セッション後」→「セッションアーカイブ」の対応付け**
    const mapping = {
        "■セッション日": "■日付",
        "■今回のセッションを通して得た気づきや学び": "■ セッションでの気づき・学び",
        "■今回決定したアクションプラン": "■ 決定したアクションプラン",
        "■この期間の意気込み": "■ この期間の意気込み"
    };

    // **データ転記（対象の項目のみ上書き）**
    for (const [sourceHeader, targetHeader] of Object.entries(mapping)) {
        const outputColIndex = outputHeaders.indexOf(sourceHeader);
        const archiveColIndex = archiveHeaders.indexOf(targetHeader);

        if (outputColIndex !== -1 && archiveColIndex !== -1) {
            const value = lastRow[outputColIndex];
            archiveSheet.getRange(targetRow, archiveColIndex + 1).setValue(value);
        }
    }

    // **枠線を追加**
    archiveSheet.getRange(targetRow, 1, 1, archiveSheet.getLastColumn()).setBorder(true, true, true, true, true, true);

    // **ここで担当コーチに通知を送信**
    sendCoachNotificationSessionAfter(responderId, sessionDate);
}


/**
 * 担当コーチにLINEで通知を送る
 */
function sendCoachNotificationSessionAfter(responderId, sessionDate) {
  const pjtSS = SpreadsheetApp.openById(SPREAD_SHEET_IDS.PJTManagement);
  const customerSheet = pjtSS.getSheetByName(SHEET_NAMES_MANAGEMENT.CLIENT_LIST);
  const coachSheet = pjtSS.getSheetByName(SHEET_NAMES_MANAGEMENT.COACH_LIST);

  // 顧客名簿のデータ取得
  const customerData = customerSheet.getDataRange().getValues();
  const customerHeaders = customerData[0];

  const customerIdIndex = customerHeaders.indexOf("回答者ID");
  const coachNoIndex = customerHeaders.indexOf("担当コーチNo.");
  const clientNameIndex = customerHeaders.indexOf("名前");
  const clientNoteIndex = customerHeaders.indexOf("クライアントノート");

  let coachNo = null;
  let clientName = null;
  let clientNoteURL = null;

  for (let i = 1; i < customerData.length; i++) {
    if (customerData[i][customerIdIndex] === responderId) {
      coachNo = customerData[i][coachNoIndex];
      clientName = customerData[i][clientNameIndex];
      clientNoteURL = customerData[i][clientNoteIndex];
      break;
    }
  }


// 厳密にチェック
if (
  clientName === "" || clientName == null ||
  coachNo === "" || coachNo == null ||
  clientNoteURL === "" || clientNoteURL == null
) {
  Logger.log("❌ 顧客情報の一部が空です（厳密チェック）");
  return;
}


  // コーチ名簿のデータ取得
  const coachData = coachSheet.getDataRange().getValues();
  const coachHeaders = coachData[0];

  const coachNoColIndex = coachHeaders.indexOf("コーチNo.");
  const coachLineIdIndex = coachHeaders.indexOf("LINE ID");

  let coachLineId = null;
  for (let i = 1; i < coachData.length; i++) {
    if (coachData[i][coachNoColIndex] === coachNo) {
      coachLineId = coachData[i][coachLineIdIndex];
      break;
    }
  }

  if (!coachLineId) {
    Logger.log("❌ コーチのLINE IDが見つかりません");
    return;
  }

  // LINE送信用メッセージ作成
  const message = `【セッション後アウトプット提出】
${clientName} さんのアウトプットが記録されました。

回答内容はクライアントノートの「セッションアーカイブ」から確認できます。

▼クライアントノート
${clientNoteURL}`;

  sendLINEMessage(coachLineId, message);
}



// /**
//  * 担当コーチに通知メールを送る
//  */
// function sendCoachNotification(responderId, sessionDate) {
//     const pjtSS = SpreadsheetApp.openById(SPREAD_SHEET_IDS.PJTManagement);
//     const customerSheet = pjtSS.getSheetByNameI(SHEET_NAMES_MANAGEMENT.CLIENT_LIST);
//     const coachSheet = pjtSS.getSheetByName(SHEET_NAMES_MANAGEMENT.COACH_LIST);

//     // 顧客名簿のデータ取得
//     const customerData = customerSheet.getDataRange().getValues();
//     const customerHeaders = customerData[0]; // ヘッダー（1行目）

//     // "回答者ID", "担当コーチNo", "名前" の列を取得
//     const customerIdIndex = customerHeaders.indexOf("回答者ID");
//     const coachNoIndex = customerHeaders.indexOf("担当コーチNo.");
//     const clientNameIndex = customerHeaders.indexOf("名前"); // クライアント名を取得


//     let coachNo = null;
//     let clientName = null;
//     for (let i = 1; i < customerData.length; i++) {
//         if (customerData[i][customerIdIndex] === responderId) {
//             coachNo = customerData[i][coachNoIndex];
//             clientName = customerData[i][clientNameIndex]; // クライアント名を取得
//             break;
//         }
//     }


//     // コーチ名簿のデータ取得
//     const coachData = coachSheet.getDataRange().getValues();
//     const coachHeaders = coachData[0]; // ヘッダー（1行目）

//     // "担当コーチNo" と "メールアドレス" の列を取得
//     const coachNoColIndex = coachHeaders.indexOf("コーチNo.");
//     const emailIndex = coachHeaders.indexOf("メールアドレス");

   
//     let coachEmail = null;
//     for (let i = 1; i < coachData.length; i++) {
//         if (coachData[i][coachNoColIndex] === coachNo) {
//             coachEmail = coachData[i][emailIndex];
//             break;
//         }
//     }

  
// // クライアントノートのURLを取得
// const clientNoteIndex = customerHeaders.indexOf("クライアントノート");

// if (clientNoteIndex === -1) {
//     Logger.log("クライアントノートのURLの列が見つかりません");
//     return;
// }

// let clientNoteURL = null;
// for (let i = 1; i < customerData.length; i++) {
//     if (customerData[i][customerIdIndex] === responderId) {
//         clientNoteURL = customerData[i][clientNoteIndex];
//         break;
//     }
// }

// // メール送信
// const subject = `【セッション後アウトプット提出】${clientName}さん`;
// const body = `${clientName} さんのアウトプットが記録されました。\n\n` +
//              `セッション日: ${sessionDate}\n\n` +
//              '回答内容はクライアントノートの「セッションアーカイブ」から確認できます。\n\n' +
//              `クライアントノート: ${clientNoteURL}`
//              '\n\n\n' +
//              MAIL_FOOTER;

// const mailOptions = {
//     name: MAIL_SENDER_NAME // 送信者名を指定
// };

// // メール送信（表示名付き）
// MailApp.sendEmail(coachEmail, subject, body, mailOptions);

// }
