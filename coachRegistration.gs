function processCoachRegistration(sheet, lastRow) {
  const targetSheet = GET_COACH_SHEET();

  const columnMapping = {
    '回答者名': 'LINE登録名',
    '名前（フルネーム)': 'お名前（フルネーム）',
    '名前（フリガナ）': 'フリガナ',
    'Slack名': 'Slack名',
    '回答者ID': '回答者ID'
  };

  const sourceHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const targetHeaders = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
  const sourceData = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()[0];

  const answerId = sourceData[sourceHeaders.indexOf("回答者ID")];
  const authInfo = findAuthInfoByAnswerId(answerId, USER_TYPE.COACH);

  if (!authInfo) {
    const logMsg = `❌ 認証管理シートに回答者ID [${answerId}] が見つかりませんでした。処理を停止します。`;
    Logger.log(logMsg);
    return;
  }

  // 既存のコーチを検索（LINE登録名で特定）
  let targetRow = findCoachRowByLineName(targetSheet, sourceHeaders, sourceData, targetHeaders);

  if (targetRow) {
    Logger.log(`✅ 既存のコーチデータを更新: 行 ${targetRow}`);
    updateCoachRow(targetSheet, targetRow, sourceHeaders, sourceData, columnMapping, targetHeaders);
  } else {
    Logger.log("⚠️ 新規コーチデータを追加");
    targetRow = addNewCoachRow(targetSheet, sourceHeaders, sourceData, columnMapping, targetHeaders, authInfo);
  }

  // コーチNo.の設定
  setCoachNumber(targetSheet, targetRow);

  // 専用フォルダURLの設定
  assignCoachFolderIfNotExists(targetSheet, targetHeaders, targetRow)

  // LINE メッセージ送信
  sendCoachNotification(targetSheet, targetRow, targetHeaders);
}



/**
 * 専用フォルダURLが未設定ならフォルダを作成
 */
function assignCoachFolderIfNotExists(targetSheet, targetHeaders, targetRow) {
    const folderUrlColIndex = targetHeaders.indexOf(COACH_LIST_TBL.DEDICATED_FOLDER_URL) + 1;
    const fullNameColIndex = targetHeaders.indexOf(COACH_LIST_TBL.NAME_KANJI) + 1;
    const coachNoColIndex = targetHeaders.indexOf(COACH_LIST_TBL.COACH_NO) + 1;
    const coachMailColIndex = targetHeaders.indexOf(COACH_LIST_TBL.MAIL) + 1;

    if (folderUrlColIndex === -1 || fullNameColIndex === -1 || coachNoColIndex === -1) return;

    let folderUrl = targetSheet.getRange(targetRow, folderUrlColIndex).getValue();
    if (folderUrl) {
        Logger.log(`🔹 行 ${targetRow} の専用フォルダURLは既に設定済み: ${folderUrl}`);
        return; // 既に設定されている場合はスキップ
    }

    const fullName = targetSheet.getRange(targetRow, fullNameColIndex).getValue();
    const coachNo = targetSheet.getRange(targetRow, coachNoColIndex).getValue();
    const coachMail = targetSheet.getRange(targetRow, coachMailColIndex).getValue();

    const parentFolder = DriveApp.getFolderById(FOLDER_IDS.COACH_PARENT);
    const folderName = `コーチNo.${coachNo}_${fullName}`;
    const folder = parentFolder.createFolder(folderName);
    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    folderUrl = folder.getUrl();

    // ✅ コーチのアカウントのみに権限を付与
    const allowedEmails = [coachMail].filter(email => email && email.includes("@"));


    // ✅ 最新の編集者リストを取得
    const currentEditors = DriveApp.getFolderById(folder.getId()).getEditors().map(user => user.getEmail());

    const ownerEmail = folder.getOwner().getEmail();

// 1. 先にコーチの権限を付与（確実に追加する）
allowedEmails.forEach(email => {
    try {
        const folderId = folder.getId();
        Drive.Permissions.create(
  {
    role: "writer",
    type: "user",
    emailAddress: coachMail
  },
  folderId,
  { sendNotificationEmail: false }
);

        Logger.log("✅ 通知なしで編集権限を付与: " + email);
    } catch (e) {
        Logger.log("⚠️ 編集権限の付与エラー: " + email + " - " + e.toString());
    }
});


// 2. その後、不要な編集者を削除
currentEditors.forEach(email => {
    if (!allowedEmails.includes(email) && email !== ownerEmail) {
        try {
            folder.removeEditor(email);
            Logger.log("✅ 不要な編集者を削除: " + email);
        } catch (e) {
            Logger.log("⚠️ 削除エラー: " + email + " - " + e.toString());
        }
    }
});


    // ✅ **フォルダの権限変更が完了した後にリンク共有を無効化**
Utilities.sleep(1000); // **1秒待機して権限変更を確実に反映**
folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.NONE);
Logger.log("🔒 フォルダのリンク共有を無効化しました");


    // ✅ フォルダURLをシートに保存
    targetSheet.getRange(targetRow, folderUrlColIndex).setValue(folderUrl);
    Logger.log(`✅ 行 ${targetRow} に専用フォルダを作成: ${folderUrl}`);
}


  



/**
 * コーチ登録完了の通知を **LINE** で送信
 */
function sendCoachNotification(targetSheet, row, targetHeaders) {
  const fullName = targetSheet.getRange(row, targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0].indexOf(COACH_LIST_TBL.NAME_KANJI) + 1).getValue();
  const foldeURL = targetSheet.getRange(row, targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0].indexOf(COACH_LIST_TBL.DEDICATED_FOLDER_URL) + 1).getValue();
  const coahNo = targetSheet.getRange(row, targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0].indexOf(COACH_LIST_TBL.COACH_NO) + 1).getValue();
  const message = `【コーチ登録完了】\n\n${fullName} 様\n\nコーチNo：${coahNo}\nコーチ専用フォルダURL：${foldeURL}`;
  const message2 = `▼クライアント登録の手順▼
  以下の内容をクライアントに送信してください！
  ※なお、コーチ番号が間違っていないかの確認と
  「サポートフェーズ」の〇〇に当てはまる数字を記入した上で送信していただくようお願いいたします！`;

const message3 = ` ▼以下の文章をコピペしてクライアントに送信▼

サポート開始にあたりボディメイクコーチサポート専用LINEの友達追加をお願いします！

こちら↓
${INFLOW_ACTION_URL.CLIENT_INFLOW}

「初回登録フォーム」が送信されますので、ご入力ください！

※入力に必要な情報は下記をご確認ください。
コーチ番号：${coahNo}
サポートフェーズ：⚪︎

以上、ご確認お願いいたします！`

  const lineIdColIndex = targetHeaders.indexOf(COACH_LIST_TBL.LINE_ID) + 1;
  const lineId = targetSheet.getRange(row, lineIdColIndex).getValue();
  if (lineId) sendLINEMessage(lineId, message);
  sendLINEMessage(lineId, message2);
  sendLINEMessage(lineId, message3);

}



/**
 * LINE登録名が一致する既存のコーチデータを探す
 */
function findCoachRowByLineName(targetSheet, sourceHeaders, sourceData, targetHeaders) {
  Logger.log("🔹 コーチ名簿を検索");

  const existingData = targetSheet.getDataRange().getValues();
  const lineNameColIndex = targetHeaders.indexOf(COACH_LIST_TBL.LINE_NAME);

  if (lineNameColIndex === -1) {
    Logger.log("❌ LINE登録名の列が見つかりません");
    return null;
  }

  const targetLineName = (sourceData[sourceHeaders.indexOf('回答者名')] || "").toString().trim();

  for (let i = 1; i < existingData.length; i++) {
    if ((existingData[i][lineNameColIndex] || "").toString().trim() === targetLineName) {
      Logger.log(`✅ LINE登録名が一致: 行 ${i + 1}`);
      return i + 1;
    }
  }

  Logger.log("⚠️ 既存データなし、新規登録");
  return null;
}


/**
 * 既存のコーチ情報を上書き更新
 */
function updateExistingCoachRow(targetSheet, targetHeaders, targetRow, transferData) {
  for (let i = 0; i < targetHeaders.length; i++) {
    if (targetHeaders[i] !== COACH_LIST_TBL.COACH_NO && targetHeaders[i] !== COACH_LIST_TBL.DEDICATED_FOLDER_URL) { // コーチNoとフォルダURLは上書きしない
      targetSheet.getRange(targetRow, i + 1).setValue(transferData[i]);
    }
  }
  Logger.log(`✅ 既存行(${targetRow})の情報を更新しました`);
}

/**
 * 新しいコーチを追加 & コーチNo.とフォルダを設定
 */
function addNewCoachRow(targetSheet, sourceHeaders, sourceData, columnMapping, targetHeaders, authInfo) {
  const newRow = new Array(targetHeaders.length).fill('');

  // 通常のマッピング（Slack名や氏名など）
  for (const [sourceHeader, targetHeader] of Object.entries(columnMapping)) {
    const sourceIndex = sourceHeaders.indexOf(sourceHeader);
    const targetIndex = targetHeaders.indexOf(targetHeader);

    if (sourceIndex !== -1 && targetIndex !== -1) {
      newRow[targetIndex] = sourceData[sourceIndex];
    }
  }

  // 🔻 追加で認証管理由来の項目もセット
  if (targetHeaders.includes(COACH_LIST_TBL.LINE_NAME)) {
    newRow[targetHeaders.indexOf(COACH_LIST_TBL.LINE_NAME)] = authInfo.lineName;
  }

  if (targetHeaders.includes(COACH_LIST_TBL.MAIL)) {
    newRow[targetHeaders.indexOf(COACH_LIST_TBL.MAIL)] = authInfo.email;
  }

  if (targetHeaders.includes(COACH_LIST_TBL.LINE_ID)) {
    newRow[targetHeaders.indexOf(COACH_LIST_TBL.LINE_ID)] = authInfo.lineId;
  }

  const newRowPosition = targetSheet.getLastRow() + 1;
  targetSheet.appendRow(newRow);
  Logger.log(`✅ 新規コーチを行 ${newRowPosition} に追加しました`);

  const range = targetSheet.getRange(newRowPosition, 1, 1, targetSheet.getLastColumn());
  range.setBackground('#ffffff');
  range.setFontColor("#000000");  
  range.setBorder(true, true, true, true, true, true);

  return newRowPosition;
}


/**
 * コーチNo.の割り当て
 */
function setCoachNumber(targetSheet, row) {
  const coachNoColIndex = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0].indexOf(COACH_LIST_TBL.COACH_NO) + 1;
  if (coachNoColIndex === 0) return;

  const currentCoachNo = targetSheet.getRange(row, coachNoColIndex).getValue();
  if (!currentCoachNo) {
    const maxCoachNo = Math.max(0, ...targetSheet.getRange(2, coachNoColIndex, targetSheet.getLastRow() - 1, 1).getValues().flat().filter(num => !isNaN(num) && num !== ''));
    const newCoachNo = maxCoachNo + 1;
    targetSheet.getRange(row, coachNoColIndex).setValue(newCoachNo);
    Logger.log(`✅ コーチNo. ${newCoachNo} を設定`);
  }
}



function updateCoachRow(targetSheet, targetRow, sourceHeaders, sourceData, columnMapping, targetHeaders) {
  for (const [sourceHeader, targetHeader] of Object.entries(columnMapping)) {
    const sourceIndex = sourceHeaders.indexOf(sourceHeader);
    const targetIndex = targetHeaders.indexOf(targetHeader);

    if (sourceIndex !== -1 && targetIndex !== -1) {
      // LINE ID / メールアドレス は上書きしない
      if (targetHeader === COACH_LIST_TBL.LINE_ID || targetHeader === COACH_LIST_TBL.MAIL) continue;

      targetSheet.getRange(targetRow, targetIndex + 1).setValue(sourceData[sourceIndex]);
    }
  }
}
