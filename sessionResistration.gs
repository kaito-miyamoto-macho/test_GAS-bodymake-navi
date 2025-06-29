function handleSessionRegistration(sheet, lastRow) {
  const clientSheet = GET_CLIENT_SHEET(); // クライアント名簿を取得
  const coachSheet = GET_COACH_SHEET();   // コーチ名簿を取得

  if (!clientSheet || !coachSheet) {
    Logger.log("❌ クライアント名簿またはコーチ名簿が見つかりません");
    return;
  }

  const sourceHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const formData = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()[0];

  const clientNoIndex = sourceHeaders.indexOf('クライアントNo.');
  const sessionDateIndex = sourceHeaders.indexOf('セッション日時');
  const zoomLinkIndex = sourceHeaders.indexOf('セッション用zoomリンク');
  const coachNoIndex = sourceHeaders.indexOf('コーチ番号');
  const responderIDNoIndex = sourceHeaders.indexOf(ELME_HIDDEN_COL.ANSWER_ID);

  if (clientNoIndex === -1 || sessionDateIndex === -1 || zoomLinkIndex === -1 || coachNoIndex === -1) {
    Logger.log('❌ 必須項目が見つかりません');
    return;
  }

  const clientNo = formData[clientNoIndex];
  const sessionDate = formData[sessionDateIndex];
  const zoomLink = formData[zoomLinkIndex];
  const coachNo = formData[coachNoIndex];
  const responderID = formData[responderIDNoIndex];

  if (!clientNo || !sessionDate || !zoomLink || !coachNo) {
    Logger.log('❌ 必須データが不足しています');
    return;
  }

  // **クライアント情報の取得**
  const clientData = findClientByCustomerNo(clientNo, clientSheet);
  if (!clientData) {
    Logger.log(`❌ クライアントNo.${clientNo} に対応するクライアントが見つかりません`);
    return;
  }

  //  // **コーチ名簿から回答者IDが一致する行のコーチ番号を取得**
  // const coachDataByResponderID = findCoachByResponderID(responderID, coachSheet); // 関数を用いてコーチデータ取得

  // if (!coachDataByResponderID) {
  //   Logger.log(`❌ 回答者ID ${responderID} に対応するコーチが見つかりません`);
  //   return;
  // }

   const { clientName, clientLineId, clientCoachNo } = clientData; // クライアントの担当コーチNo.も取得


  // **コーチ情報の取得（コーチ番号で特定）**
  const coachData = findCoachByNo(coachNo, coachSheet);
  if (!coachData) {
    Logger.log(`❌ コーチ番号 ${coachNo} に対応するコーチが見つかりません`);
    return;
  }

  const { coachName, coachLineId } = coachData;

  /**
   * 取得したコーチ番号が異なる場合、LINEでエラーを返して処理終了
  **/
  if (Number(clientCoachNo) !== Number(coachNo)) {
    Logger.log(`${coachNo} ${clientCoachNo}`)
    const errMsg = `⚠️セッション登録エラー\n顧客No.${clientNo}は担当クライアントに該当しません。\n正しいクライアントNoを確認して再度登録してください。`;

    sendLINEMessage(coachLineId, errMsg)
    return;
  }

  const calenderTitle = `${clientName}さんセッション`;
  const calenderDetails = `Zoomリンク: ${zoomLink}`;
  const calenderLocation = "Zoom";
  const formatSessionDate = new Date(sessionDate)
  const calenderEndDate = new Date(formatSessionDate.getTime() + 60 * 60 * 1000); // 1時間後

  const calendarUrl = generateGoogleCalendarUrl(calenderTitle, formatSessionDate, calenderEndDate, calenderDetails, calenderLocation);

  // **クライアントへのLINE通知**
  const clientMessage = `【セッション登録通知】\n${clientName}さん\n\n以下の日程でセッションが登録されました。
  ーーーーーーーーーーーー
  ▼開始日時
  ${sessionDate}

  ▼Zoomリンク
  ${zoomLink}

  ▼事前アウトプットはこちら
  ${LINE_FROM_URLS.SESSION_BEFORE_OUTPUT}

  ▼予定をGoogleカレンダーに追加する
  ${calendarUrl}
  ーーーーーーーーーーーー`;

  if (clientLineId) sendLINEMessage(clientLineId, clientMessage);

  // **コーチへのLINE通知**
  const coachMessage = `【セッション登録完了】\n${clientName}さんとのセッションが登録されました。
  ーーーーーーーーーーーー
  ▼開始日時
  ${sessionDate}

  ▼Zoomリンク
  ${zoomLink}
  ーーーーーーーーーーーー`;

  if (coachLineId) sendLINEMessage(coachLineId, coachMessage);
  
 // 転記先のシート（セッション管理シート）
  const sessionSheet = SpreadsheetApp.openById(SPREAD_SHEET_IDS.COMPASSNaviSystem).getSheetByName("セッションスケジュール管理");

  if (!sessionSheet) {
    Logger.log("❌ 'セッション一覧' シートが見つかりません");
    return;
  }

  const headers = sessionSheet.getRange(1, 1, 1, sessionSheet.getLastColumn()).getValues()[0];

  const valueMap = {
    "クライアント名": clientName,
    "担当コーチNo.": coachNo,
    "担当コーチ名": coachName,
    "セッション開始日時": formatSessionDate,
    "zoomリンク": zoomLink,
    "リマインド1日前": "未送信",
    "リマインド1時間前": "未送信",
    "リマインド1時間後": "未送信",
    "全体ステータス": "有効",
    "クライアントLINE ID": clientLineId
  };

  const rowValues = headers.map(header => valueMap[header] ?? ""); // 未定義項目には空文字を入れる

  sessionSheet.appendRow(rowValues);

 // 🔲 直後に枠線を追加（最終行に対して）
  const lastSessionRow = sessionSheet.getLastRow();
  sessionSheet.getRange(lastSessionRow, 1, 1, sessionSheet.getLastColumn()).setBorder(
    true,  // 上
    true,  // 左
    true,  // 下
    true,  // 右
    true,  // 垂直線
    true   // 水平線
  );

  Logger.log(`✅ セッション一覧に登録完了：${clientName} - ${sessionDate}`);


}




function findClientByCustomerNo(clientNo, clientSheet) {
  const data = clientSheet.getDataRange().getValues();
  const headers = data[0];

  const customerNoIndex = headers.indexOf(CLIENT_LIST_TBL.CLIENT_NO);
  const clientNameIndex = headers.indexOf(CLIENT_LIST_TBL.NAME_KANJI);
  const clientLineIdIndex = headers.indexOf(CLIENT_LIST_TBL.LINE_ID);
  const clientCoachNoIndex = headers.indexOf(CLIENT_LIST_TBL.RES_COACH_NO);

  // ❗見つからなかったカラムをログ出力
  const missingColumns = [];
  if (customerNoIndex === -1) missingColumns.push(CLIENT_LIST_TBL.CLIENT_NO);
  if (clientNameIndex === -1) missingColumns.push(CLIENT_LIST_TBL.NAME_KANJI);
  if (clientLineIdIndex === -1) missingColumns.push(CLIENT_LIST_TBL.LINE_ID);

  if (missingColumns.length > 0) {
    Logger.log(`❌ クライアント名簿に必要なカラムが見つかりません: ${missingColumns.join(', ')}`);
    return null;
  }

  for (let i = 1; i < data.length; i++) {
    if (data[i][customerNoIndex] == clientNo) {
      return {
        clientName: data[i][clientNameIndex],
        clientLineId: data[i][clientLineIdIndex],
        clientCoachNo: data[i][clientCoachNoIndex]  // ここも必要なら -1 チェック可能
      };
    }
  }

  Logger.log(`⚠️ 顧客No. ${clientNo} に一致するクライアントが見つかりません`);
  return null;
}



function findCoachByNo(coachNo, coachSheet) {
  const data = coachSheet.getDataRange().getValues();
  const headers = data[0];
  const coachNoIndex = headers.indexOf(COACH_LIST_TBL.COACH_NO);
  const coachNameIndex = headers.indexOf(COACH_LIST_TBL.NAME_KANJI);
  const coachLineIdIndex = headers.indexOf(COACH_LIST_TBL.LINE_ID);

  if (coachNoIndex === -1 || coachNameIndex === -1 || coachLineIdIndex === -1) {
    Logger.log("❌ コーチ名簿に必要なカラムが見つかりません");
    return null;
  }

  for (let i = 1; i < data.length; i++) {
    if (data[i][coachNoIndex] == coachNo) {
      return {
        coachName: data[i][coachNameIndex],
        coachLineId: data[i][coachLineIdIndex]
      };
    }
  }

  Logger.log(`⚠️ コーチ番号 ${coachNo} に一致するコーチが見つかりません`);
  return null;
}

/**
 * 回答者IDを元にコーチ番号を取得
 */
function findCoachByResponderID(responderId, coachSheet) {
  const data = coachSheet.getDataRange().getValues(); // シートの全データを取得
  const headers = data[0]; // ヘッダー行を取得

  const responderIdIndex = headers.indexOf(COACH_LIST_TBL.ANSWER_ID); // 回答者IDの列
  const coachNoIndex = headers.indexOf(COACH_LIST_TBL.COACH_NO); // コーチ番号の列

  if (responderIdIndex === -1 || coachNoIndex === -1) {
    Logger.log("❌ コーチ名簿に必要なカラムがありません");
    return null;
  }

  for (let i = 1; i < data.length; i++) {
    if (data[i][responderIdIndex] === responderId) {
      return { coachNo: data[i][coachNoIndex] }; // コーチ番号を返す
    }
  }

  return null; // 該当データなし
}






/**
 * リマインダーを設定する
 */
// function scheduleSessionReminders(clientLineId, clientName, sessionDate, zoomLink) {
//   // sessionDateが文字列の場合、日付オブジェクトに変換
//   if (!(sessionDate instanceof Date)) {
//     sessionDate = new Date(sessionDate);
//     if (isNaN(sessionDate.getTime())) {
//       Logger.log(`❌ エラー: sessionDate が無効な日付形式です (${sessionDate})`);
//       return;
//     }
//   }

//   const reminders = [
//     { offsetHours: 24, type: "24hour", message: `【明日はセッションです！】
//     ▼開始日時
//     ${sessionDate}
    
//     ▼Zoomリンク
//     ${zoomLink}
    
//     ▼事前アウトプットはこちら
//     ${LINE_FROM_URLS.SESSION_BEFORE_OUTPUT}`},

//     { offsetHours: 1, type: "1hour", message: `【もうすぐセッションです！】
//     ▼開始日時
//     ${sessionDate}
    
//     ▼Zoomリンク
//     ${zoomLink}
    
//     ▼事前アウトプットフォーム
//     ${LINE_FROM_URLS.SESSION_BEFORE_OUTPUT}`}
//   ];

//   reminders.forEach(reminder => {
//     const reminderTime = new Date(sessionDate.getTime() - reminder.offsetHours * 60 * 60 * 1000);

//     if (reminderTime > new Date()) {
//       ScriptApp.newTrigger("sendSessionReminder")
//         .timeBased()
//         .at(reminderTime)
//         .create();

//       // ユーザー情報をプロパティに保存
//       const scriptProperties = PropertiesService.getScriptProperties();
//       scriptProperties.setProperty("clientLineId", clientLineId);
//       scriptProperties.setProperty("clientName", clientName);
//       scriptProperties.setProperty("sessionDate", sessionDate.toISOString());
//       scriptProperties.setProperty("zoomLink", zoomLink);
//       scriptProperties.setProperty("reminderType", reminder.type);

//       Logger.log(`✅ ${reminder.offsetHours}時間前のリマインダーを設定しました (${reminderTime})`);
//     } else {
//       Logger.log(`⚠️ ${reminder.offsetHours}時間前のリマインダーは現在時刻より前のためスキップ (${reminderTime})`);
//     }
//   });
// }


/**
 * リマインダー送信用関数
 */
function sendSessionReminder() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const clientLineId = scriptProperties.getProperty("clientLineId");
  const clientName = scriptProperties.getProperty("clientName");
  const sessionDate = new Date(scriptProperties.getProperty("sessionDate"));
  const zoomLink = scriptProperties.getProperty("zoomLink");
  const reminderType = scriptProperties.getProperty("reminderType");

  if (!clientLineId || !clientName || isNaN(sessionDate.getTime()) || !zoomLink) {
    Logger.log("❌ リマインダーのデータが不足しているため送信を中止");
    return;
  }

  const message = reminderType === "24hour"
    ? `【明日はセッションです！】
    ▼開始日時: ${sessionDate}
    
    ▼Zoomリンク
    ${zoomLink}
    
    ▼事前アウトプットフォーム
    ${LINE_FROM_URLS.SESSION_BEFORE_OUTPUT}`
    :
    `【もうすぐセッションです！】
    ▼開始日時
    ${sessionDate}
    
    ▼Zoomリンク
    ${zoomLink}
    
    ▼事前アウトプットフォーム
    ${LINE_FROM_URLS.SESSION_BEFORE_OUTPUT}`;

  sendLINEMessage(clientLineId, message);

  Logger.log(`📩 ${reminderType} リマインダーを送信しました`);
}



/**
 * セッション一覧をチェックして、リマインドを送る関数（毎時間実行トリガー用）
 */
function sendSessionReminders() {
  const sessionSheet = SpreadsheetApp.openById(SPREAD_SHEET_IDS.COMPASSNaviSystem)
    .getSheetByName("セッションスケジュール管理");

  if (!sessionSheet) {
    Logger.log("❌ 'セッションスケジュール管理' シートが見つかりません");
    return;
  }

  const data = sessionSheet.getDataRange().getValues();
  const headers = data[0];
  const now = new Date();

  // カラム位置取得
  const colIndex = (name) => headers.indexOf(name);
  const clientNameCol = colIndex("クライアント名");
  const lineIdCol = colIndex("クライアントLINE ID");
  const sessionDateCol = colIndex("セッション開始日時");
  const zoomCol = colIndex("zoomリンク");
  const remind24hCol = colIndex("リマインド1日前");
  const remind1hCol = colIndex("リマインド1時間前");
  const remindAfterCol = colIndex("リマインド1時間後");
  const statusCol = colIndex("全体ステータス");

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    const status = row[statusCol];
    if (status !== "有効") continue;

    const clientName = row[clientNameCol];
    const lineId = row[lineIdCol];
    const sessionDate = new Date(row[sessionDateCol]);
    const zoomLink = row[zoomCol];

    if (!lineId || isNaN(sessionDate.getTime()) || !zoomLink) continue;

    const diffHours = (sessionDate.getTime() - now.getTime()) / (1000 * 60 * 60);
    const elapsedHours = (now.getTime() - sessionDate.getTime()) / (1000 * 60 * 60);

    // ① 24時間前（19時台に送信、セッションは翌日の日付）
    if (row[remind24hCol] === "未送信") {
  const sessionDateOnly = new Date(sessionDate.getFullYear(), sessionDate.getMonth(), sessionDate.getDate());
  const nowDateOnly = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const diffDays = (sessionDateOnly - nowDateOnly) / (1000 * 60 * 60 * 24);

  if (diffDays === 1 && now.getHours() >= 19 && now.getHours() <= 21) {
    const message = `【明日はセッションです！】\n▼開始日時\n${formatDate(sessionDate)}\n\n▼Zoomリンク\n${zoomLink}\n\n▼事前アウトプットはこちら\n${LINE_FROM_URLS.SESSION_BEFORE_OUTPUT}`;

    sendLINEMessage(lineId, message);
    sessionSheet.getRange(i + 1, remind24hCol + 1).setValue("済み");
  }
}


    // ② 1時間前通知
    if (row[remind1hCol] === "未送信" && diffHours <= 1.25 && diffHours >= 0.5) {
      const message = `【まもなくセッションです】
      \n▼開始日時\n${formatDate(sessionDate)}\n\n▼Zoomリンク\n${zoomLink}\n\n▼事前アウトプットフォーム\n${LINE_FROM_URLS.SESSION_BEFORE_OUTPUT}`;

      sendLINEMessage(lineId, message);
      sessionSheet.getRange(i + 1, remind1hCol + 1).setValue("済み");
    }

    // ③ 終了後1時間（フォローアップ）
    if (row[remindAfterCol] === "未送信" && elapsedHours >= 1 && elapsedHours <= 2) {
      const message = `【セッション後アウトプット】\nセッション終了後、以下のフォームからセッション後のアウトプットをしましょう！\nセッションの効果を最大化するために、セッション後すぐに忘れないうちのアウトプットがおすすめです！\n↓ ↓ ↓\n${LINE_FROM_URLS.SESSION_AFTER_OUTPUT}`;

      sendLINEMessage(lineId, message);
      sessionSheet.getRange(i + 1, remindAfterCol + 1).setValue("済み");
    }
  }
}

/**
 * 日付を yyyy/MM/dd HH:mm 形式で整形
 */
function formatDate(date) {
  return Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
}


/**
 * セッション日時をカレンダーに登録するリンクを作成する
 */
function generateGoogleCalendarUrl(title, startDate, endDate, details, location) {
  const formatDate = date => Utilities.formatDate(date, 'UTC', 'yyyyMMdd\'T\'HHmmss\'Z\'');

  const baseUrl = "https://calendar.google.com/calendar/render?action=TEMPLATE";
  const params = [
    `text=${encodeURIComponent(title)}`,
    `dates=${formatDate(startDate)}/${formatDate(endDate)}`,
    `details=${encodeURIComponent(details)}`,
    `location=${encodeURIComponent(location)}`
  ];

  return `${baseUrl}&${params.join('&')}`;
}
