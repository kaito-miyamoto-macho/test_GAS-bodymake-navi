/**
 * 顧客情報シートにデータを登録
　- 回答者IDまたはLINE名で既存行を判定
　- 必要項目をマッピングして転記（顧客Noも自動採番）
　- メールアドレスとLINE ID、LINE利用名は認証管理から回答者IDをマッピングして取得、追記

クライアント専用ノートをテンプレートから自動生成
　- サポートフェーズに応じてシートの列を非表示にする
　- コーチ専用フォルダ内にクライアントフォルダを作成し、ノートを格納

パーソナルデータシートへの情報転記・アクセス権設定
　- アンケート内容をクライアントノートのパーソナルデータに反映
　- クライアントとコーチ（Gmail）にのみ編集権限を付与

LINEで通知を自動送信
　- クライアントにノートのURLと登録フォームを送付
　- コーチにはノートURLと顧客Noなどの情報を通知

記録用に日次シートへ登録日も記録
 */

function transferMappedDataToAnotherSpreadsheet(sourceSheet, lastRow) {
  Logger.log("クライアント登録処理 start");

  const targetClientSheet = GET_CLIENT_SHEET();
  const targetCoachSheet = GET_COACH_SHEET();

  if (!sourceSheet || !targetClientSheet || !targetCoachSheet) {
    Logger.log("シートが見つかりません");
    return;
  }

  const columnMapping = {
    '名前（フルネーム）': '名前',
    '名前（フリガナ）': 'フリガナ',
    '回答者名': 'LINE名',
    '回答者ID': '回答者ID',
    '性別': '性別',
    '生年月日': '生年月日',
    '担当コーチ番号': '担当コーチNo.',
    'サポートフェーズ選択': 'サポートフェーズ',
    '回答日時': 'アンケート回答日'
  };

  const sourceHeaders = GET_HEADER(sourceSheet, 1);
  const targetHeaders = GET_HEADER(targetClientSheet, 1);
  const sourceData = sourceSheet.getRange(lastRow, 1, 1, sourceSheet.getLastColumn()).getValues()[0];

  const answerId = sourceData[sourceHeaders.indexOf(ELME_HIDDEN_COL.ANSWER_ID)];
  const authInfo = findAuthInfoByAnswerId(answerId, USER_TYPE.CLIENT);

  if (!authInfo) {
    Logger.log(`❌ 認証管理シートに回答者ID [${answerId}] が見つかりませんでした。処理を停止します。`);
    return;
  }

  const clientLineId = authInfo.lineId;
  const clientEmail = authInfo.email;

  Logger.log(`🔍 取得したソースデータ: ${JSON.stringify(sourceData)}`);

  let targetRow = findTargetRowByIdOrLineName(targetClientSheet, sourceHeaders, sourceData, targetHeaders);


  if (targetRow) {
    // 既存行が見つかった＝登録済み
    Logger.log(`⚠️ 既に登録済みのアカウントです（行: ${targetRow}）`);

    if (clientLineId) {
      const message = `⚠️ アカウント情報は既に登録済みです。\n\n` +
                      `情報を変更したい場合は、\n` +
                      `「クライアント情報の更新フォーム」から手続きをお願いします。\n\n`+
                      `▼更新フォーム\n`+
                      `${LINE_FROM_URLS.CLIENT_UPDATE}\n\n`+
                      `また、MYメニュー＞アカウント情報を更新する\n`+
                      `でも手続き可能です。`;

      sendLINEMessage(clientLineId, message);
      Logger.log(`⚠️ 重複登録を検出し、LINE通知を送信しました: ${clientLineId}`);
    }

     return; // これ以上処理せず終了
    } else {
      Logger.log("⚠️ 新規行を追加します");
      targetRow = targetClientSheet.getLastRow() + 1;
    }

  let newRow = new Array(targetHeaders.length).fill("");
  let clientName = '';
  let coachEmail = '';
  let coachName = '';
  let coachFolderUrl = '';
  let formattedDate = "";

  sourceHeaders.forEach((sourceHeader, colIndex) => {
    const targetHeader = columnMapping[sourceHeader];
    if (targetHeader) {
      const targetColIndex = targetHeaders.indexOf(targetHeader);
      if (targetColIndex !== -1) {
        if (sourceHeader === '名前（フルネーム）') {
          clientName = sourceData[colIndex];
        }
        if (sourceHeader === '担当コーチ番号') {
          const coachNumber = sourceData[colIndex];
          const coachData = targetCoachSheet.getRange(2, 1, targetCoachSheet.getLastRow() - 1, targetCoachSheet.getLastColumn()).getValues();
          const coachHeaders = targetCoachSheet.getRange(1, 1, 1, targetCoachSheet.getLastColumn()).getValues()[0];
          const coachRow = coachData.find(row => row[coachHeaders.indexOf('コーチNo.')] == coachNumber);
          if (coachRow) {
            coachName = coachRow[coachHeaders.indexOf('お名前（フルネーム）')];
            coachEmail = coachRow[coachHeaders.indexOf('メールアドレス')];
            coachFolderUrl = coachRow[coachHeaders.indexOf('専用フォルダURL')];
          }
        }
        if (sourceHeader === '回答日時') {
          const rawDateValue = sourceData[colIndex];
          formattedDate = formatDateToYYYYMMDD(rawDateValue);
          newRow[targetColIndex] = formattedDate;
        } else {
          newRow[targetColIndex] = sourceData[colIndex];
        }
      }
    }
  });

  const coachNameColIndex = targetHeaders.indexOf(CLIENT_LIST_TBL.RES_COACH_NAME);
  if (coachNameColIndex !== -1) {
    newRow[coachNameColIndex] = coachName;
  }

  const clientEmailColIndex = targetHeaders.indexOf(CLIENT_LIST_TBL.MAIL);
  if (clientEmailColIndex !== -1) {
    newRow[clientEmailColIndex] = clientEmail;
  }

  const clientLineIdColIndex = targetHeaders.indexOf(CLIENT_LIST_TBL.LINE_ID);
  if (clientLineIdColIndex !== -1) {
    newRow[clientLineIdColIndex] = clientLineId;
  }

  targetClientSheet.appendRow(newRow);
  Logger.log("✅ 新規行を追加しました");

  targetClientSheet.getRange(targetRow, 1, 1, targetHeaders.length).setBorder(true, true, true, true, true, true);
  Logger.log(`📌 データ転記完了: 行 ${targetRow}`);



  const customerNoColIndex = targetHeaders.indexOf(CLIENT_LIST_TBL.CLIENT_NO);
  let currentCustomerNo = '';

  if (customerNoColIndex !== -1) {
    // 顧客No.は常に新規発行
    const lastRowInClientSheet = targetClientSheet.getLastRow();
    const existingNos = lastRowInClientSheet > 1
      ? targetClientSheet.getRange(2, customerNoColIndex + 1, lastRowInClientSheet - 1, 1).getValues()
          .flat()
          .filter(num => !isNaN(num) && num !== '')
      : [];

    const maxCustomerNo = existingNos.length > 0 ? Math.max(...existingNos.map(Number)) : 0;
    currentCustomerNo = maxCustomerNo + 1;
    targetClientSheet.getRange(targetRow, customerNoColIndex + 1).setValue(currentCustomerNo);
    Logger.log(`✅ 新規 顧客No. ${currentCustomerNo} を設定`);
  }

  // 「サポートフェーズ選択」に応じて表示する列を隠す
  const supportPhaseColIndex = sourceHeaders.indexOf('サポートフェーズ選択');

  if (supportPhaseColIndex === -1) {
    Logger.log(`サポートフェーズの列が存在しません`);
    return;
  }

  const supportPhaseValue = sourceData[supportPhaseColIndex];

  let selectedTemplateId = SPREAD_SHEET_IDS.CLIENT_NOTE_TEMPLATE;

  const templateFile = DriveApp.getFileById(selectedTemplateId); // クライアントノートを生成
  const copiedSpreadsheetFile = templateFile.makeCopy(`クライアントノート_${clientName}`);
  const copiedSpreadsheet = SpreadsheetApp.openById(copiedSpreadsheetFile.getId());
  // 【日次】記録シートと【日次】積み上げ日記シート本体
  const dailyOutputSheet = copiedSpreadsheet.getSheetByName(SHEET_NAMES_CLIENTSUPPORT.DAILY_OUTPUT);
  const dailyPiledUpSheet = copiedSpreadsheet.getSheetByName(SHEET_NAMES_CLIENTSUPPORT.DAILY_PILED_UP);

  // 【日次】記録シートの入力開始日を入力
  const INPUT_START_DATE_AREA = "D7"
  dailyOutputSheet.getRange(INPUT_START_DATE_AREA).setValue(formattedDate);
  Logger.log(`【日次】記録シートの入力開始日を入力完了`);

  // 【日次】記録シートと【日次】積み上げ日記シートのヘッダー
  const DAILY_OUTPUT_HEADER_ROW = 12;
  const dailyOutputHeaders = GET_HEADER(dailyOutputSheet, DAILY_OUTPUT_HEADER_ROW);
  const DAILY_PILED_UP_HEADER_ROW = 2;
  const dailyPiledUpHeaders = GET_HEADER(dailyPiledUpSheet, DAILY_PILED_UP_HEADER_ROW);

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
  
  

// **専用フォルダ作成**
if (coachFolderUrl) {
  const coachFolder = DriveApp.getFolderById(coachFolderUrl.match(/[-\w]{25,}/)[0]);
  const clientFolder = coachFolder.createFolder(`クライアントNo.${currentCustomerNo}_${clientName}`);
  clientFolder.addFile(copiedSpreadsheetFile);
  DriveApp.getRootFolder().removeFile(copiedSpreadsheetFile);
  Logger.log(`✅ フォルダ作成: クライアント${currentCustomerNo}_${clientName}`);
}


  copiedSpreadsheetFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const sharedLink = copiedSpreadsheetFile.getUrl();
  const linkColIndex = targetHeaders.indexOf('クライアントノート');
  if (linkColIndex !== -1) {
    targetClientSheet.getRange(targetRow, linkColIndex + 1).setValue(sharedLink);
  }

   // PDF表示用リンクを生成
  const pdfLink = createPDFLink(copiedSpreadsheet.getId(), SHEET_NAMES_CLIENTSUPPORT.WEEKLY_PILED_UP);



// クライアントへのLINEメッセージ内容
const clientMessage = `${clientName} 様\n\nボディメイクナビのご利用登録が完了しました。\n\n▼以下からサポートノートにアクセスできます\n${sharedLink}\n\n▼積み上げレポート\n${pdfLink}
  \n
  ※本LINEは送信専用のため、お問い合わせの際は、下記お問い合わせフォームをご利用ください。
${INQUIRY_FORM}`;

// 2通目にサポートノート登録フォームを送信
const clientMessage2 = `サポートノートをこちらから登録すると、LINEからいつでも確認できるようになります！
${LINE_FROM_URLS.REGIST_SUPPORT_NOTE}`;

// コーチへのLINEメッセージ内容
const coachMessage = `【クライアント登録完了通知】${clientName}\n担当コーチ ${coachName} \n\n▼${clientName}さんのクライアントノートURL\n${sharedLink}\n\n▼積み上げレポート
${pdfLink}\n\n初回フォームの回答内容はクライアントノートの「パーソナルデータ」から確認できます。\n\n
※本LINEは送信専用のため、お問い合わせの際は、下記のお問い合わせフォームをご利用ください。
${INQUIRY_FORM}`;

// 続けて送るコーチ宛のメッセージ
const coachMessage2 = `セッションの日程登録は以下からお願いします！
${LINE_FROM_URLS.SESSION_REGIST}

${clientName}さんの顧客No: ${currentCustomerNo}
※顧客Noはコーチ専用フォルダからも確認できます。`;


// コーチの情報を取得
const coachNumber = sourceData[sourceHeaders.indexOf('担当コーチ番号')]; // コーチNo.を取得
const coachLineId= getCoachLineId(coachNumber);

// クライアントにLINEメッセージを送信
if (clientLineId) {
    sendLINEMessage(clientLineId, clientMessage);
    sendLINEMessage(clientLineId, clientMessage2);
    Logger.log(`✅ クライアント ${clientName} にLINEメッセージ送信`);
} else {
    Logger.log(`⚠️ クライアント ${clientName} の LINE ID が未登録`);
}

// コーチにLINEメッセージを送信
if (coachLineId) {
    sendLINEMessage(coachLineId, coachMessage);
    sendLINEMessage(coachLineId, coachMessage2);
    Logger.log(`✅ コーチ ${coachName} にLINEメッセージ送信`);
} else {
    Logger.log(`⚠️ コーチ${coachNumber} ${coachName} の LINE ID が未登録`);
}



Utilities.sleep(3000); // 少し待つ



  // クライアントノートの「パーソナルデータ」シートに転記
  insertPersonalDataToClientNote(copiedSpreadsheet, sourceHeaders, sourceData, columnMapping);
  
  // Gmailアドレスのみに編集権限を付与
  allowClientNoteAccess(clientEmail, coachEmail, copiedSpreadsheetFile);

  copiedSpreadsheetFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  Logger.log("クライアント登録処理 end");
}


/********************
 * フォーム回答内容をパーソナルデータシートに転記する
 *********************/
function insertPersonalDataToClientNote(clientNoteSpreadsheet, sourceHeaders, sourceData) {
    const personalDataSheet = clientNoteSpreadsheet.getSheetByName(SHEET_NAMES_CLIENTSUPPORT.PERSONAL_DATA);
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
    '身長（数字のみ）': '身長',
    '体重（数値のみ）': '体重',
    '体脂肪率（数字のみ）': '体脂肪率',
    '生年月日': '生年月日',
    '居住状況': '居住状況',
    '体験セッションを通して得た学びや気づきは何ですか？': '体験セッションを通して得た学びや気づきは何ですか？',
    '体験セッションを受けて行動に移したいと思ったことは何ですか？': '体験セッションを受けて行動に移したいと思ったことは？'
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
