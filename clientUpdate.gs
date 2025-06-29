/**
 * 250115_【PJT管理】ボディメイクナビ登録者名簿のクライアント名簿シートのデータを更新
　- 回答者IDまたはLINE名で既存行を判定
　- 必要項目をマッピングして転記（顧客Noも自動採番）

  クライアント専用ノートをテンプレートから自動生成
　- サポートフェーズに応じてシートの列を非表示にする
　- コーチ専用フォルダ内にクライアントフォルダを作成し、ノートを格納

  パーソナルデータシートへの情報転記・アクセス権設定
　- アンケート内容をクライアントノートのパーソナルデータに反映
 */

function clientDateUpdate(sourceSheet, lastRow) {
   Logger.log("アカウント情報更新 start")
  // 250115_【PJT管理】ボディメイクナビ登録者名簿を取得
  const targetSpreadsheetId = SPREAD_SHEET_IDS.PJTManagement;
  const targetSpreadsheet = SpreadsheetApp.openById(targetSpreadsheetId);

  // クライアント名簿
  const targetClientSheet = targetSpreadsheet.getSheetByName(SHEET_NAMES_MANAGEMENT.CLIENT_LIST);

  if (!sourceSheet || !targetClientSheet) {
    Logger.log("シートが見つかりません");
    return;
  }

  const columnMapping = {
    '名前（フルネーム）': '名前',
    '名前（フリガナ）': 'フリガナ',
    '性別': '性別',
    'サポートフェーズ選択': 'サポートフェーズ',
  };

  // アカウント情報更新フォームに紐づくシートのヘッダー
  const sourceHeaders = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
  // クライアント名簿のヘッダー
  const targetHeaders = targetClientSheet.getRange(1, 1, 1, targetClientSheet.getLastColumn()).getValues()[0];

  // アカウント情報更新フォームから送信されたデータ
  const sourceData = sourceSheet.getRange(lastRow, 1, 1, sourceSheet.getLastColumn()).getValues()[0];

  Logger.log(`🔍 取得したソースデータ: ${JSON.stringify(sourceData)}`);

  // LINEIDと回答者IDを元に更新対象行を取得
  let targetRow = findTargetRowByIdOrLineName(targetClientSheet, sourceHeaders, sourceData, targetHeaders);

  Logger.log(`🔍 確定したターゲット行: ${targetRow}`);

  if (targetRow) {
    Logger.log(`✅ 既存行(${targetRow})を更新します`);
  } else {
    Logger.log("更新対象のデータが見つかりませんでした");
    return;
  }

  // 上書き用の空配列を用意（クライアント名簿の列数と同じ長さ）
  let overwriteValues = new Array(targetHeaders.length).fill("");

  // 必要な項目だけ、フォームから取得して上書き用配列にセット
  Object.entries(columnMapping).forEach(([sourceHeader, targetHeader]) => {
    const sourceColIndex = sourceHeaders.indexOf(sourceHeader); // フォームの列番号
    const targetColIndex = targetHeaders.indexOf(targetHeader); // 名簿の列番号

    if (sourceColIndex !== -1 && targetColIndex !== -1) {
      const value = sourceData[sourceColIndex]; // フォームで送られてきた値
      overwriteValues[targetColIndex] = value;  // 上書き用配列に格納

      Logger.log(`✅ 転記準備: ${targetHeader} ← ${value}`);
    } else {
      Logger.log(`⚠️ 転記スキップ（列が見つからない）: ${sourceHeader} → ${targetHeader}`);
    }
  });

  // 上書き実行
  for (let i = 0; i < targetHeaders.length; i++) {
    const headerName = targetHeaders[i];
    const isMappedTarget = Object.values(columnMapping).includes(headerName);

    if (isMappedTarget) {
      targetClientSheet.getRange(targetRow, i + 1).setValue(overwriteValues[i]);
      Logger.log(`📝 上書き実行: ${headerName} ← ${overwriteValues[i]}`);
    } else {
      Logger.log(`⏩ 上書きスキップ: ${headerName}`);
    }
  }

  Logger.log(`✅ 既存行(${targetRow})を更新しました`);
  
  targetClientSheet.getRange(targetRow, 1, 1, targetHeaders.length).setBorder(true, true, true, true, true, true);
  Logger.log(`📌 データ転記完了: 行 ${targetRow}`);


  // 「サポートフェーズ選択」に応じてサポートノートの列を切り替える
  // クライアントノートのURLを取得
  const noteIndex = targetHeaders.indexOf("クライアントノート");
  const clientNoteURL = noteIndex !== -1 ? targetClientSheet.getRange(targetRow, noteIndex + 1).getValue() : "";

  if (!clientNoteURL) {
      Logger.log("該当するクライアントノートのURLが見つかりません");
      return;
  }

  // クライアントノートのスプレッドシートを開く
  const archiveSS = SpreadsheetApp.openByUrl(clientNoteURL);

  // サポートフェーズによってシートの列の表示・非表示を切り替える
  // 一旦全列表示状態にしたあと、サポートフェーズによって非表示列を切り替える
  // 【日次】記録シートと【日次】積み上げ日記シート本体
  const dailyOutputSheet = archiveSS.getSheetByName(SHEET_NAMES_CLIENTSUPPORT.DAILY_OUTPUT);
  const dailyPiledUpSheet = archiveSS.getSheetByName(SHEET_NAMES_CLIENTSUPPORT.DAILY_PILED_UP);

  const supportPhaseColIndex = sourceHeaders.indexOf('サポートフェーズ選択');
  if (supportPhaseColIndex !== -1) {
    // サポートフェーズ
    const supportPhaseValue = sourceData[supportPhaseColIndex];

      // 【日次】記録シートと【日次】積み上げ日記シートのヘッダー
    const DAILY_OUTPUT_HEADER_ROW = 12;
    const dailyOutputHeaders = dailyOutputSheet.getRange(DAILY_OUTPUT_HEADER_ROW, 1, 1, dailyOutputSheet.getLastColumn()).getValues()[0];
    const DAILY_PILED_UP_HEADER_ROW = 2;
    const dailyPiledUpHeaders = dailyPiledUpSheet.getRange(DAILY_PILED_UP_HEADER_ROW, 1, 1, dailyPiledUpSheet.getLastColumn()).getValues()[0];

    // データ列を取得 【日次】記録シート
    const bodyWeghtColIndex = dailyOutputHeaders.indexOf(DAILY_OUTPUT.BODY_WEGHT);
    const mealColIndex = dailyOutputHeaders.indexOf(DAILY_OUTPUT.MEAL);
    const kcalColIndex = dailyOutputHeaders.indexOf(DAILY_OUTPUT.KCAL);
    const proteinColIndex = dailyOutputHeaders.indexOf(DAILY_OUTPUT.PROTEIN);
    const fatColIndex = dailyOutputHeaders.indexOf(DAILY_OUTPUT.FAT);
    const carbColIndex = dailyOutputHeaders.indexOf(DAILY_OUTPUT.CARB);
    const fiberColIndex = dailyOutputHeaders.indexOf(DAILY_OUTPUT.FIBER);
    const waterColIndex = dailyOutputHeaders.indexOf(DAILY_OUTPUT.WATER);
    const tarainingColIndex = dailyOutputHeaders.indexOf(DAILY_OUTPUT.TRAINING);
    const poopColIndex = dailyOutputHeaders.indexOf(DAILY_OUTPUT.POOP);
    const walkColIndex = dailyOutputHeaders.indexOf(DAILY_OUTPUT.WALK);
    const sleepColIndex = dailyOutputHeaders.indexOf(DAILY_OUTPUT.SLEEP);
    const conditionColIndex = dailyOutputHeaders.indexOf(DAILY_OUTPUT.CONDITION);

    // データ列を取得 【日次】積み上げ日記
    const executionRateColIndex = dailyPiledUpHeaders.indexOf(DAILY_PILED_UP.EXECUTION_RATE);
    const commitColIndex = dailyPiledUpHeaders.indexOf(DAILY_PILED_UP.COMMIT);
    const encouragementColIndex = dailyPiledUpHeaders.indexOf(DAILY_PILED_UP.ENCOURAGEMENT);
    const consultationColIndex = dailyPiledUpHeaders.indexOf(DAILY_PILED_UP.CONSULTATION);

    const dailyOutputHideColmunsList = [];
    const dailyPiledUpHideColmunsList = [];

    // 最大列数（非表示列も含む）
    const dailyOutputSheetlastColumn = dailyOutputSheet.getMaxColumns();
    const dailyPiledUpSheetlastColumn = dailyPiledUpSheet.getMaxColumns();

    // 全列表示
    dailyOutputSheet.showColumns(1, dailyOutputSheetlastColumn); 
    dailyPiledUpSheet.showColumns(1, dailyPiledUpSheetlastColumn); 


    switch (supportPhaseValue) {
      case SUPPORT_PHASE.DIETER_ROOKIE:
        // 非表示にする列を設定
        dailyOutputHideColmunsList.push(bodyWeghtColIndex);
        dailyOutputHideColmunsList.push(kcalColIndex);
        dailyOutputHideColmunsList.push(proteinColIndex);
        dailyOutputHideColmunsList.push(fatColIndex);
        dailyOutputHideColmunsList.push(carbColIndex);
        dailyOutputHideColmunsList.push(fiberColIndex);
        dailyOutputHideColmunsList.push(waterColIndex);
        dailyOutputHideColmunsList.push(poopColIndex);
        dailyOutputHideColmunsList.push(walkColIndex);
        // 隣の列も非表示にすると表示が自然
        dailyOutputHideColmunsList.push(bodyWeghtColIndex + 1);
        dailyOutputHideColmunsList.push(kcalColIndex + 1);
        dailyOutputHideColmunsList.push(proteinColIndex + 1);
        dailyOutputHideColmunsList.push(fatColIndex + 1);
        dailyOutputHideColmunsList.push(carbColIndex + 1);
        dailyOutputHideColmunsList.push(fiberColIndex + 1);
        dailyOutputHideColmunsList.push(waterColIndex + 1);
        dailyOutputHideColmunsList.push(poopColIndex + 1);
        dailyOutputHideColmunsList.push(walkColIndex + 1);

        dailyPiledUpHideColmunsList.push(executionRateColIndex + 1);
        dailyPiledUpHideColmunsList.push(commitColIndex + 1);
        dailyPiledUpHideColmunsList.push(encouragementColIndex + 1);

        // 非表示処理
        dailyOutputHideColmunsList.forEach(col => {
            if (col > 0 && col <= dailyOutputSheet.getMaxColumns()) {
              dailyOutputSheet.hideColumns(col);
            } else {
              Logger.log(`⚠️ 【日次】記録:無効な列番号でhideColumnsをスキップ: ${col}`);
            }
        });
        dailyPiledUpHideColmunsList.forEach(col => {
            if (col > 0 && col <= dailyPiledUpSheet.getMaxColumns()) {
              dailyPiledUpSheet.hideColumns(col);
            } else {
              Logger.log(`⚠️ 【日次】積み上げ日記:無効な列番号でhideColumnsをスキップ: ${col}`);
            }
        });

        break;
      case SUPPORT_PHASE.DIETER:
        // 非表示にする列を設定
        dailyOutputHideColmunsList.push(kcalColIndex);
        dailyOutputHideColmunsList.push(proteinColIndex);
        dailyOutputHideColmunsList.push(fatColIndex);
        dailyOutputHideColmunsList.push(carbColIndex);
        dailyOutputHideColmunsList.push(fiberColIndex);
        dailyOutputHideColmunsList.push(waterColIndex);
        dailyOutputHideColmunsList.push(walkColIndex);
        // 隣の列も非表示にすると表示が自然
        dailyOutputHideColmunsList.push(kcalColIndex + 1);
        dailyOutputHideColmunsList.push(proteinColIndex + 1);
        dailyOutputHideColmunsList.push(fatColIndex + 1);
        dailyOutputHideColmunsList.push(carbColIndex + 1);
        dailyOutputHideColmunsList.push(fiberColIndex + 1);
        dailyOutputHideColmunsList.push(waterColIndex + 1);
        dailyOutputHideColmunsList.push(walkColIndex + 1);

        dailyPiledUpHideColmunsList.push(encouragementColIndex + 1);

        // 非表示処理
        dailyOutputHideColmunsList.forEach(col => {
            if (col > 0 && col <= dailyOutputSheet.getMaxColumns()) {
              dailyOutputSheet.hideColumns(col);
            } else {
              Logger.log(`⚠️ 【日次】記録:無効な列番号でhideColumnsをスキップ: ${col}`);
            }
        });
        dailyPiledUpHideColmunsList.forEach(col => {
            if (col > 0 && col <= dailyPiledUpSheet.getMaxColumns()) {
              dailyPiledUpSheet.hideColumns(col);
            } else {
              Logger.log(`⚠️ 【日次】積み上げ日記:無効な列番号でhideColumnsをスキップ: ${col}`);
            }
        });
        break;
      case SUPPORT_PHASE.BODY_MAKE_BEGINNER:
        // 非表示にする列を設定
        dailyOutputHideColmunsList.push(proteinColIndex);
        dailyOutputHideColmunsList.push(fatColIndex);
        dailyOutputHideColmunsList.push(carbColIndex);
        dailyOutputHideColmunsList.push(fiberColIndex);
        // 隣の列も非表示にすると表示が自然
        dailyOutputHideColmunsList.push(proteinColIndex + 1);
        dailyOutputHideColmunsList.push(fatColIndex + 1);
        dailyOutputHideColmunsList.push(carbColIndex + 1);
        dailyOutputHideColmunsList.push(fiberColIndex + 1);

        dailyPiledUpHideColmunsList.push(encouragementColIndex + 1);          

        // 非表示処理
        dailyOutputHideColmunsList.forEach(col => {
            if (col > 0 && col <= dailyOutputSheet.getMaxColumns()) {
              dailyOutputSheet.hideColumns(col);
            } else {
              Logger.log(`⚠️ 【日次】記録:無効な列番号でhideColumnsをスキップ: ${col}`);
            }
        });
        dailyPiledUpHideColmunsList.forEach(col => {
            if (col > 0 && col <= dailyPiledUpSheet.getMaxColumns()) {
              dailyPiledUpSheet.hideColumns(col);
            } else {
              Logger.log(`⚠️ 【日次】積み上げ日記:無効な列番号でhideColumnsをスキップ: ${col}`);
            }
        });
        break;
      case SUPPORT_PHASE.BODY_MAKE_ADVANCED:
        // 食事評価を非表示にする
        dailyOutputHideColmunsList.push(mealColIndex);
        dailyOutputHideColmunsList.push(mealColIndex + 1);

        // 非表示処理
        dailyOutputHideColmunsList.forEach(col => {
            if (col > 0 && col <= dailyOutputSheet.getMaxColumns()) {
              dailyOutputSheet.hideColumns(col);
            } else {
              Logger.log(`⚠️ 【日次】記録:無効な列番号でhideColumnsをスキップ: ${col}`);
            }
        });
          break;
      default:
        Logger.log(`⚠️ 未知のサポートフェーズ: ${supportPhaseValue}。非表示処理なし`);
        break;
    }
  }

  // クライアントノートの「パーソナルデータ」シートに転記 
  updatePersonalDataToClientNote(archiveSS, sourceHeaders, sourceData, columnMapping);

  // 更新が完了したらクライアント、コーチにメッセージを送る
  // クライアントのLINE IDを取得
  const clientLineIdColIndex = targetHeaders.indexOf('LINE ID');
  const byCoachNoColIndex = targetHeaders.indexOf('担当コーチNo.');
  const clientNoteColIndex = targetHeaders.indexOf('クライアントノート');
  const clientNameColIndex = targetHeaders.indexOf('LINE名');

  const clientLineId = clientLineIdColIndex !== -1 ? targetClientSheet.getRange(targetRow, clientLineIdColIndex + 1).getValue() : "";
  const coachNo = byCoachNoColIndex !== -1 ? targetClientSheet.getRange(targetRow, byCoachNoColIndex + 1).getValue() : "";
  const clientNoteUrl = clientNoteColIndex !== -1 ? targetClientSheet.getRange(targetRow, clientNoteColIndex + 1).getValue() : "";
  const clientName = clientNameColIndex !== -1 ? targetClientSheet.getRange(targetRow, clientNameColIndex + 1).getValue() : "";

  const targetCoachSheet = targetSpreadsheet.getSheetByName(SHEET_NAMES_MANAGEMENT.COACH_LIST);

  // コーチ名簿のデータ取得
  const coachData = targetCoachSheet.getDataRange().getValues();
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

  // メッセージ作成
  // クライアントに更新完了メッセージを送信
  const clientMsg = 
    `アカウント情報が更新されました！`;
  sendLINEMessage(clientLineId, clientMsg);

  // コーチにクライアントがアカウント情報を更新したメッセージを送信
  const coachMsg =`【クライアント情報更新通知】\n\n
${clientName} さんのアカウント情報が更新されました。\n
更新内容はクライアントノートの「パーソナルデータ」から確認できます。\n\n
▼クライアントノート\n
${clientNoteUrl}`;
sendLINEMessage(coachLineId, coachMsg);

  Logger.log("アカウント情報更新 end");
  return;
}

/**
 * GOAL更新処理
 * 
 * 1. GOAL更新フォームから送信されたデータを取得
 *    - 必要なデータ（人生のGOAL、ボディメイクの目的など）をフォームから取得
 *    - 回答者IDで特定のクライアントを検索し、そのデータを更新対象として選定
 * 
 * 2. クライアント名簿シートを検索
 *    - クライアント名簿の行を回答者IDで検索
 *    - 該当するクライアント情報（LINE ID、クライアントノート、担当コーチNo.）を取得
 * 
 * 3. クライアントノートシートを更新
 *    - クライアントノートシートを開き、指定されたセルにGOALのデータを反映
 *    - 更新内容として人生のGOAL、ボディメイクの目的、理想の身体、1年後の目標、3ヶ月後の目標を記入
 * 
 * 4. LINE通知
 *    - クライアントにGOAL更新完了のメッセージを送信
 *    - 担当コーチにクライアントのGOAL更新を通知
 */
function goalUpdate (sourceSheet, lastRow) {
  Logger.log("GOAL更新start");
   // 250115_【PJT管理】ボディメイクナビ登録者名簿を取得
  const targetSpreadsheetId = SPREAD_SHEET_IDS.PJTManagement;
  const targetSpreadsheet = SpreadsheetApp.openById(targetSpreadsheetId);

  // クライアント名簿
  const targetClientSheet = targetSpreadsheet.getSheetByName(SHEET_NAMES_MANAGEMENT.CLIENT_LIST);

  if (!sourceSheet || !targetClientSheet) {
    Logger.log("シートが見つかりません");
    return;
  }

  // GOAL更新フォームに紐づくシートのヘッダー
  const sourceHeaders = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
  // クライアント名簿のヘッダー
  const targetHeaders = targetClientSheet.getRange(1, 1, 1, targetClientSheet.getLastColumn()).getValues()[0];

  // GOAL更新フォームから送信されたデータ
  const sourceData = sourceSheet.getRange(lastRow, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
  // クライアント名簿
  const clientData = targetClientSheet.getDataRange().getValues();

  Logger.log(`🔍 取得したソースデータ: ${JSON.stringify(sourceData)}`);

  const answerIdIndex = sourceHeaders.indexOf("回答者ID");

  const lifeGoalIndex = sourceHeaders.indexOf("人生のGOAL");
  const bodyMakeTargetIndex = sourceHeaders.indexOf("ボディメイクの目的");
  const idealBodyIndex = sourceHeaders.indexOf("理想のカラダ");
  const targetAfter1YearIndex = sourceHeaders.indexOf("1年後の目標");
  const currentTargetIndex = sourceHeaders.indexOf("現在の目標");
  

  // GOAL更新フォームを入力したデータ
  const answerId = sourceData[answerIdIndex];
  const lifeGoal = sourceData[lifeGoalIndex];
  const bodyMakeTarget = sourceData[bodyMakeTargetIndex];
  const idealBody = sourceData[idealBodyIndex];
  const targetAfter1Year = sourceData[targetAfter1YearIndex];
  const currentTarget = sourceData[currentTargetIndex];

  const clientAnswerIdColIndex = targetHeaders.indexOf('回答者ID');
  const clientNoteColIndex = targetHeaders.indexOf('クライアントノート');
  const clientLineIdColIndex = targetHeaders.indexOf('LINE ID');
  const assigneeCoachNoColIndex = targetHeaders.indexOf('担当コーチNo.');
  const clientNameIndex = targetHeaders.indexOf("名前");


  if (clientAnswerIdColIndex === -1 || clientNoteColIndex === -1) {
    Logger.log("❌ クライアント名簿内の必要列が見つかりません");
    return;
  }

  // LINEIDと回答者IDを元にGOAL更新クライアントを特定
  let clientNoteURL = "";
  let clientLineId = "";
  let assigneeCoachNo = "";
  let clientName = "";

  for (let i = 1; i < clientData.length; i++) {
    const rowAnswerId = clientData[i][clientAnswerIdColIndex];
    Logger.log(`回答者ID称号:${rowAnswerId} == ${answerId}`);
    if (rowAnswerId == answerId) {
      Logger.log(`GOAL更新対象クライアント列: ${i + 1}`);
      clientNoteURL = clientData[i][clientNoteColIndex];
      clientLineId = clientData[i][clientLineIdColIndex];
      assigneeCoachNo = clientData[i][assigneeCoachNoColIndex];
      clientName = clientData[i][clientNameIndex];
      if (!clientNoteURL) {
        Logger.log("クライアントノート存在なし");
        return;
      }
      break;
    }
  }

  if (!clientNoteURL) {
    Logger.log("GOAL更新対象クライアントが見つかりませんでした");
    return;
  }

  // クライアントノートのスプレッドシートを開く
  const archiveSS = SpreadsheetApp.openByUrl(clientNoteURL);

  // 目標シートのGOALを更新する
  // 目標シート本体
  const goalSheet = archiveSS.getSheetByName(SHEET_NAMES_CLIENTSUPPORT.GOAL);

  Logger.log("GOAL更新書き込み処理開始");
  // 人生のゴール
  const lifeGoalArea = "J9"
  goalSheet.getRange(lifeGoalArea).setValue(lifeGoal);
  // ボディメイクの目的
  const bodyMakeTargetArea = "J10"
  goalSheet.getRange(bodyMakeTargetArea).setValue(bodyMakeTarget);
  // 理想の身体
  const idealBodyArea = "J11"
  goalSheet.getRange(idealBodyArea).setValue(idealBody);
  // 1年後の目標
  const targetAfter1YearArea = "J12"
  goalSheet.getRange(targetAfter1YearArea).setValue(targetAfter1Year);
  // 3ヶ月後の目標
  const targetAfter3monthArea = "J13"
  goalSheet.getRange(targetAfter3monthArea).setValue(currentTarget);
   Logger.log("GOAL更新書き込み処理終了");

  // GOAL更新完了通知をLINEで送信
  const coachSheet = targetSpreadsheet.getSheetByName(SHEET_NAMES_MANAGEMENT.COACH_LIST);
  // コーチ名簿のデータ取得
  const coachData = coachSheet.getDataRange().getValues();
  // コーチ名簿のヘッダー
  const coachHeaders = coachSheet.getRange(1, 1, 1, coachSheet.getLastColumn()).getValues()[0];

  const coachNoColIndex = coachHeaders.indexOf(COACH_LIST_TBL.COACH_NO);
  const coachLineIdIndex = coachHeaders.indexOf(COACH_LIST_TBL.LINE_ID);

  let coachLineId = "";
  for (let i = 1; i < coachData.length; i++) {
    if (coachData[i][coachNoColIndex] === assigneeCoachNo) {
      coachLineId = coachData[i][coachLineIdIndex];
      break;
    }
  }

  if (!coachLineId) {
    Logger.log("❌ コーチのLINE IDが見つかりません");
    return;
  }

  Logger.log("クライアントと担当コーチにLINEで通知");
  // 更新したユーザーに通知
  const msg = `GOAL更新が完了しました！`;
  sendLINEMessage(clientLineId, msg);


  // 担当コーチへ通知
  const forCoachMsg = `【GOAL更新】\n\n
${clientName} さんのGOALが更新されました。\n\n
更新内容はクライアントノートの「目標」から確認できます。\n\n
▼クライアントノート\n\
${clientNoteURL}`;

  sendLINEMessage(coachLineId, forCoachMsg);

  Logger.log("GOAL更新end");
}





/**
 * Gmailアドレスのみに編集権限を付与し、それ以外の権限を削除
 */ 
function allowClientNoteAccess(clientEmail, coachEmail, file) {
  const fileId = file.getId(); // ← これが必要
  const allowedEmails = [clientEmail, coachEmail].filter(email => email && email.endsWith("@gmail.com")); // Gmailアドレスのみ許可

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
    }
  });
  // リンク共有を無効化（重要）
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.NONE);
}


/********************
 * フォーム回答内容でパーソナルデータシートを更新する
 *********************/
function updatePersonalDataToClientNote(archiveSS, sourceHeaders, sourceData) {
    const personalDataSheet = archiveSS.getSheetByName(SHEET_NAMES_CLIENTSUPPORT.PERSONAL_DATA);
    if (!personalDataSheet) {
        Logger.log("パーソナルデータシートが見つかりません");
        return;
    }

    // **B列の項目名リストを取得**
    const personalDataHeaders = personalDataSheet.getRange(1, 2, personalDataSheet.getLastRow(), 1).getValues().flat();

    // 解答内容転記用マッピング
  const personalDataMapping = {
    '名前（フルネーム）': '名前',
    '名前（フリガナ）': '名前（フリガナ）',
    '性別': '性別',
    'サポートフェーズ選択': 'サポートフェーズ',
    '体重（数値のみ）': '体重',
    '体脂肪率（数字のみ）': '体脂肪率',
    '居住状況': '居住状況',
};


    // **フォームのデータを転記**
    for (const [formHeader, personalDataHeader] of Object.entries(personalDataMapping)) {
        const formColIndex = sourceHeaders.indexOf(formHeader);
        const personalDataRowIndex = personalDataHeaders.indexOf(personalDataHeader);
        
        if (formColIndex !== -1 && personalDataRowIndex !== -1) {
            let valueToSet = sourceData[formColIndex];

            // ★ 居住状況の特殊処理
            if (formHeader === '居住状況' && valueToSet === 'その他') {
                const otherIndex = sourceHeaders.indexOf(`その他を選んだ方は、
具体的な居住状況をご記入ください`);
                const otherValue = otherIndex !== -1 ? sourceData[otherIndex] : '';
                valueToSet = otherValue ? otherValue : 'その他(未入力)';
            }

            personalDataSheet.getRange(personalDataRowIndex + 1, 3).setValue(valueToSet);
        }
    }

    Logger.log("パーソナルデータの転記が完了しました");
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

  Logger.log("⚠️ 該当する行が見つからない");
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
