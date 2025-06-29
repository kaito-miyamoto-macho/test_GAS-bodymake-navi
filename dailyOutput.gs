function  noteDailyOutput() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getActiveSheet();
  if (!sourceSheet) return Logger.log("エラー: アクティブなシートが取得できません。");

  const lastRow = sourceSheet.getLastRow();
  if (lastRow < 2) return Logger.log("エラー: データが不足しています（ヘッダーのみ or 空データ）");

  const headers = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
  const lastRowData = sourceSheet.getRange(lastRow, 1, 1, sourceSheet.getLastColumn()).getValues()[0];

  // 回答結果オブジェクトを生成
  const rawResponse = {};
  headers.forEach((header, i) => {
    rawResponse[header] = lastRowData[i];
  });
  Logger.log(`回答オブジェクト:  ${rawResponse}`);

  // 回答者IDを取得
  const clientId = Number(rawResponse["回答者ID"]);
  if (isNaN(clientId)) {
  return Logger.log("エラー: 回答者IDが数値として不正です");
}


  // クライアント名簿からスプレッドシートURLを取得
  const clientSpreadsheet = SpreadsheetApp.openById(SPREAD_SHEET_IDS.PJTManagement);
  const clientSheet = clientSpreadsheet.getSheetByName(SHEET_NAMES_MANAGEMENT.CLIENT_LIST);
  const clientData = clientSheet.getDataRange().getValues();

  const clientHeaders = clientData[0];
  const clientIdIndex = clientHeaders.indexOf("回答者ID");
  const clientSheetUrlIndex = clientHeaders.indexOf("クライアントノート");
  const clientNameIndex = clientHeaders.indexOf("名前");

  if (clientIdIndex === -1 || clientSheetUrlIndex === -1) {
    return Logger.log("エラー: クライアント名簿のヘッダーが不正です");
  }

  const clientMap = new Map(
  clientData.slice(1).map(row => [
    Number(row[clientIdIndex]), // ←ここを数値に
    extractSpreadsheetId(row[clientSheetUrlIndex])
  ])
);



  if (!clientMap.has(clientId)) {
    return Logger.log(`エラー: クライアントIDが見つかりません → ${clientId}`);
  }

  const targetSpreadsheetId = clientMap.get(clientId);
  const targetSpreadsheet = SpreadsheetApp.openById(targetSpreadsheetId);

  // 転記処理
  writeDailyRecords(rawResponse, targetSpreadsheet);

    // 担当コーチNoを取得
  const coachNoIndex = clientHeaders.indexOf("担当コーチNo.");
  if (coachNoIndex === -1) return Logger.log("エラー: クライアント名簿にコーチNo.の列が見つかりません");

  const clientRow = clientData.find(row => Number(row[clientIdIndex]) === clientId);
  if (!clientRow) return Logger.log("エラー: 回答者IDに一致するクライアント行が見つかりません");

  const coachNo = clientRow[coachNoIndex];
  const coachLineId = getCoachLineId(coachNo);
  const clientName = clientRow[clientNameIndex] // クライアント名取得

  if (coachLineId) {
    const message = `【日次アウトプット通知】\nクライアント ${clientName} さんが日次アウトプットを回答しました。\n\n▼クライアントノート\n${clientRow[clientSheetUrlIndex]}`;
    sendLINEMessage(coachLineId, message);
  } else {
    Logger.log(`⚠️ 担当コーチのLINE IDが見つからないため、通知できません（コーチNo: ${coachNo}）`);
  }

}

/**
 * 回答結果をクライアントノートの2つのシートに転記
 */
function writeDailyRecords(formResponse, spreadsheet) {
  Logger.log("=== 転記対象の回答データ ===");
  Object.entries(formResponse).forEach(([key, value]) => {
    Logger.log(`${key} → ${value}`);
  });
  Logger.log("===========================");

  // Key: エルメの入力フォーム, label: サポートノートのヘッダー
  const mappings = [
    {
      sheetName: SHEET_NAMES_CLIENTSUPPORT.DAILY_OUTPUT,
      headerRow: 12,
      mapping: [
        // ①ダイエット初心者
        { key: "■日付", label: "DATE" },
        { key: "■食事", label: "食事" },
        { keys: ["■運動","■運動・トレーニング"], // ③
         label: "運動・トレーニング" },
        { key: "■睡眠", label: "睡眠" },
        { key: "■体調", label: "体調" },
        // 追加：②ダイエッター
        { key: "■体重(kg)", label: "体重(kg)" },
        { key: "■お通じ回数", label: "お通じ回数" },
        // ③ボディメイク初心者
        { key: "■摂取カロリー(kcal)", label: "摂取カロリー(kcal)" },
        { key: "■水分摂取量(L)", label: "水分補給(L)" },
        { key: "■歩数", label: "歩数(歩)" },
        // ④ボディメイク上級者
        { key: "■タンパク質量(g）", label: "タンパク質(g)" },
        { key: "■脂質量(g）", label: "脂質(g)" },
        { key: "■炭水化物量(g）", label: "炭水化物(g)" },
        { key: "■食物繊維量(g）", label: "食物繊維(g)" },
      ]
    },

    {
      sheetName: SHEET_NAMES_CLIENTSUPPORT.DAILY_PILED_UP,
      headerRow: 2,
      mapping: [
        // ダイエット初心者
        { key: "■日付", label: "日付" },
        { key: "■コーチに相談、シェアしたいことなどあれば＾＾", label: "コーチに相談、シェアしたいこと　※任意" },
        // 追加：ダイエッター, ボディメイク初心者
        { key: "■宣言したアクションを何％実行できましたか？", label: "宣言したアクション\nの実行率" },
        { key: "■明日（今日）のアクション宣言　※1つでOK！", label: "明日（今日）のアクション宣言" },
        // 追加：ボディメイク上級者
        { key: "■理想の自分から今日の自分への一言", label: "理想の自分から今日の自分への一言　※任意" }
      ]
    }
  ];

  mappings.forEach(({ sheetName, headerRow, mapping }) => {
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) return Logger.log(`エラー: シートが見つかりません → ${sheetName}`);

    let headers = [];
    try {
      headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0] || [];
    } catch (e) {
      Logger.log(`ヘッダーの取得に失敗しました: ${e}`);
    }
    if (headers.length === 0) {
      Logger.log(`ヘッダーが見つかりません → ${sheetName} の ${headerRow} 行目`);
      return;
    }

    const rowData = new Array(headers.length).fill("");

    const dateKey = "■日付";
    const dateLabel = mapping.find(m => m.key === dateKey)?.label;
    const dateColIndex = headers.indexOf(dateLabel);
    console.log(`dateColIndex: ${dateColIndex}, dalaLabel: ${dateLabel}`)
    const inputDateRaw = formResponse[dateKey];
    const inputDate = new Date(inputDateRaw);
    if (isNaN(inputDate)) return Logger.log(`エラー: 日付が不正 → ${inputDateRaw}`);
    const formattedDate = Utilities.formatDate(inputDate, Session.getScriptTimeZone(), "MM/dd");

    let targetRow;

    // 【日次】記録：日付が一致する行を探して上書き
    if (sheetName === SHEET_NAMES_CLIENTSUPPORT.DAILY_OUTPUT) {
      const numRows = Math.max(1, sheet.getLastRow() - headerRow);
      const range = sheet.getRange(headerRow + 1, dateColIndex + 1, numRows);
      const values = range.getValues();
      const matchIndex = values.findIndex(row => {
        const cellDate = new Date(row[0]);
        const formatted = Utilities.formatDate(cellDate, Session.getScriptTimeZone(), "MM/dd");
        return formatted === formattedDate;
      });

      if (matchIndex !== -1) {
        targetRow = headerRow + 1 + matchIndex;
      } else {
        Logger.log(`エラー: ${sheetName} に該当の日付が見つかりません → ${formattedDate}`);
        return;
      }
    }

    // 【日次】積み上げ日記：補完 or 上書き
    if (sheetName === SHEET_NAMES_CLIENTSUPPORT.DAILY_PILED_UP) {
      const numRows = Math.max(1, sheet.getLastRow() - headerRow);
      const range = sheet.getRange(headerRow + 1, dateColIndex + 1, numRows);
      const values = range.getValues();
      const formattedDates = values.map(row => {
        const d = new Date(row[0]);
        return isNaN(d) ? null : Utilities.formatDate(d, Session.getScriptTimeZone(), "MM/dd");
      });

      const existingIndex = formattedDates.indexOf(formattedDate);

      if (existingIndex !== -1) {
        targetRow = headerRow + 1 + existingIndex;
      } else {
        const validDates = formattedDates.filter(d => d).map(d => new Date(d));
        const lastDate = validDates.length ? validDates.sort((a, b) => b - a)[0] : null;

        if (!lastDate) {
          // 新品状態 → 回答日だけの空行を作成
          const emptyRow = new Array(headers.length).fill("");
          emptyRow[dateColIndex] = formattedDate;

          targetRow = sheet.getLastRow() + 1;
          sheet.appendRow(emptyRow);

          const range = sheet.getRange(targetRow, 1, 1, headers.length);
          range.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);


        } else {
          const gapDays = Math.floor((inputDate - lastDate) / (1000 * 60 * 60 * 24));
          for (let i = 1; i <= gapDays; i++) {
            const newDate = new Date(lastDate.getTime());
            newDate.setDate(newDate.getDate() + i);

            const formattedGap = Utilities.formatDate(newDate, Session.getScriptTimeZone(), "MM/dd");

            if (formattedGap === formattedDate) {
              targetRow = sheet.getLastRow() + 1;
              break; // ← 回答日まで来たらループ終了！
            }

            const emptyRow = new Array(headers.length).fill("");
            emptyRow[dateColIndex] = formattedGap;

            const rowIndex = sheet.getLastRow() + 1;
            sheet.appendRow(emptyRow);
            const range = sheet.getRange(rowIndex, 1, 1, headers.length);
            range.setBackgrounds([new Array(headers.length).fill("#fce4e4")]);
            range.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
          }

          targetRow = sheet.getLastRow() + 1;
        }
      }
    }

    // 回答内容の転記（label or labels に対応）
    mapping.forEach(({ key, keys, label, labels }) => {
      const labelList = labels || [label];
      const colIndex = headers.findIndex(h => labelList.includes(h));
      if (colIndex === -1) return;

      const keyList = keys || [key];
      const responseValue = keyList.map(k => formResponse[k]).find(v => v !== undefined);

      if (responseValue !== undefined) {
        rowData[colIndex] = keyList.includes(dateKey) ? formattedDate : responseValue;
      }
    });



    sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);

    if (sheetName === SHEET_NAMES_CLIENTSUPPORT.DAILY_PILED_UP) {
      const range = sheet.getRange(targetRow, 1, 1, rowData.length);
      range.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
    }
  });
}



/**
 * スプレッドシートのURLから ID を抽出する関数
 */
function extractSpreadsheetId(url) {
  const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
  return match ? match[1] : null;
}
