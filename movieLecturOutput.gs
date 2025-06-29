/**
 * 各動画講座のアウトプット内容を該当するクライアントノートに転記する
 * @param sheet: 回答内容が書かれたシート
 * @param lastRow: 回答内容が記入された最終行
 * @param lectureIndex: 動画講座識別用インデックス（0〜4）
 */
CONST_LECTURE_INDEX = {
  DM: 0,
  BM1: 1,
  BM2: 2,
  BM3: 3,
  BM4: 4,
};

function handleLectureOutput(sheet, lastRow, lectureIndex) {
  Logger.log("▶ handleLectureOutput START");
  Logger.log("対象 lectureIndex: " + lectureIndex);
  Logger.log("対象 sheet name: " + sheet.getName() + ", lastRow: " + lastRow);

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowData = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()[0];

  const answerIdIndex = headers.indexOf("回答者ID");
  if (answerIdIndex === -1) {
    Logger.log("❌ 回答者ID列が見つかりません");
    return;
  }

  const answerId = rowData[answerIdIndex];
  Logger.log("✅ 回答者ID取得: " + answerId);

  const clientSS = SpreadsheetApp.openById(SPREAD_SHEET_IDS.PJTManagement);
  Logger.log("✅ PJT管理スプレッドシート取得");

  const clientSheet = clientSS.getSheetByName(SHEET_NAMES_MANAGEMENT.CLIENT_LIST);
  if (!clientSheet) {
    Logger.log("❌ クライアント名簿シートが見つかりません");
    return;
  }

  const clientData = clientSheet.getDataRange().getValues();
  const clientHeaders = clientData[0];
  const clientIdIndex = clientHeaders.indexOf("回答者ID");
  const noteUrlIndex = clientHeaders.indexOf("クライアントノート");

  if (clientIdIndex === -1 || noteUrlIndex === -1) {
    Logger.log("❌ クライアント名簿内の必要列が見つかりません");
    return;
  }

  const clientRow = clientData.find(row => row[clientIdIndex] === answerId);
  if (!clientRow) {
    Logger.log("❌ 回答者IDに一致するクライアントが見つかりません");
    return;
  }

  const noteUrl = clientRow[noteUrlIndex];
  Logger.log("✅ クライアントノートURL取得: " + noteUrl);

  let noteSpreadsheet;
  try {
    noteSpreadsheet = SpreadsheetApp.openByUrl(noteUrl);
  } catch (e) {
    Logger.log("❌ クライアントノートのスプレッドシート取得に失敗: " + e);
    return;
  }

  const progressSheet = noteSpreadsheet.getSheetByName("ﾎﾞﾃﾞｨﾒｲｸ完全解説　進捗管理");
  if (!progressSheet) {
    Logger.log("❌ 進捗管理シートが見つかりません");
    return;
  }

  const progressHeaders = progressSheet.getRange(2, 1, 1, progressSheet.getLastColumn()).getValues()[0];
  const progressData = progressSheet.getDataRange().getValues();
  Logger.log("✅ 進捗管理シート読み込み完了");

  const lectureKeywords = [
    "【全章まとめ　アウトプットフォーム】",
    "【第1章まとめ　アウトプットフォーム】",
    "【第2章まとめ　アウトプットフォーム】",
    "【第3章まとめ　アウトプットフォーム】",
    "【第4章まとめ　アウトプットフォーム】",
  ];

  const targetKeyword = lectureKeywords[lectureIndex];
  if (!targetKeyword) {
    Logger.log("❌ 指定された講座インデックスに対応するキーワードが存在しません");
    return;
  }

  Logger.log("🔍 転記対象キーワード: " + targetKeyword);

  let targetRowIndex = -1;
  for (let i = 2; i < progressData.length; i++) {
    if (progressData[i].some(cell => typeof cell === 'string' && cell.includes(targetKeyword))) {
      targetRowIndex = i + 1;
      Logger.log("✅ 対象行見つかりました: row " + targetRowIndex);
      break;
    }
  }

  if (targetRowIndex === -1) {
    Logger.log("❌ 対象行が見つかりません（部分一致: " + targetKeyword + "）");
    return;
  }

  const transferItems = {
    "動画を視聴する目的": null,
    "学びになったこと上位3つ": null,
    "「学びになった！」と思った理由": null,
    "気づいたことや実感したこと": null,
    "最初の一歩としていつまでに何をするか？（具体的なアクションは何か？）": null
  };

  Logger.log("🔁 転記項目マッピング開始");

  for (const key in transferItems) {
    const sourceIndex = headers.indexOf(key);
    const destIndex = progressHeaders.indexOf(key);
    Logger.log(`🔍 ${key} | sourceIndex: ${sourceIndex}, destIndex: ${destIndex}`);

    if (sourceIndex !== -1 && destIndex !== -1) {
      transferItems[key] = { sourceIndex, destIndex };
    } else {
      Logger.log(`⚠️ 項目「${key}」が見つからない（source or destination）`);
    }
  }

  Logger.log("✅ データ転記開始");
  for (const key in transferItems) {
    const indexes = transferItems[key];
    if (indexes) {
      const value = rowData[indexes.sourceIndex];
      Logger.log(`📝 転記: ${key} → row: ${targetRowIndex}, col: ${indexes.destIndex + 1}, value: ${value}`);
      progressSheet.getRange(targetRowIndex, indexes.destIndex + 1).setValue(value);
    }
  }

  Logger.log("🎉 handleLectureOutput 完了");
}
