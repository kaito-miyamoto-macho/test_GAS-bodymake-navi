/**
 * 各クライアントの利用開始日から現在までの利用週数を計算し、
 * 対応する週（WEEK◯）の行にPDFリンクを記録する処理。
 *
 * 処理対象：
 * - クライアントノートのURLが存在し、利用日数が上限未満のもの
 * - クライアントノートに「【集計】全期間積み上げレポート」シートが存在するもの
 *
 * 動作概要：
 * - 利用開始日から利用日数を算出
 * - 利用日数から週数（1週目〜）を算出
 * - 該当するWEEK行にPDFリンクを自動で記録
 *
 * 実行タイミング：
 * - 毎週土曜午前1時にトリガー起動（対象期間：前週の土曜〜金曜）
 */
function allPeriodStackedReportCreation() {
  const debugsheet = GET_LOG_SHEET();

  debugsheet.appendRow([new Date(), "全期間積み上げレポート作成 開始" ]);

  // クライアント名簿取得
  const clientSheet = GET_CLIENT_SHEET();

  if (!clientSheet) {
    Logger.log("シートが見つかりません");
    return;
  }

  // クライアント名簿のヘッダー
  const clientSheetHeaders = clientSheet.getRange(1, 1, 1, clientSheet.getLastColumn()).getValues()[0];
  // クライアント名簿のデータ部分
  const clientData = clientSheet.getDataRange().getValues();
  // クライアントノートのURLが入力されてる列
  const noteIndex = clientSheetHeaders.indexOf("クライアントノート");
  const useStartDateIndex = clientSheetHeaders.indexOf("アンケート回答日");

  const today = new Date();
  today.setDate(today.getDate() - 1); 
  const ALL_WEEKLY_PILED_UP_HEADER_ROW = 3; // 全期間積み上げレポートのヘッダー

  // 全員分実行する
  for (let i = 0; i < clientData.length; i++) {
    const dateString = clientData[i][useStartDateIndex].split('/').reverse().join('-');
    const startDate = new Date(dateString);
    debugsheet.appendRow([new Date(), `本日日付:${today} と利用開始日付:${startDate} から利用日数を計算`]);

    // 差分をミリ秒単位で計算 → 日数に変換
    const diffMs = today - startDate;
    const useDays = Math.floor(diffMs / (1000 * 60 * 60 * 24));
    debugsheet.appendRow([new Date(), `利用日数: ${useDays}`]);

    const clientNoteURL = clientData[i][noteIndex];
    // クライアントノートが存在しない、または利用日数が上限を超えているならスキップ
    if (!clientNoteURL || useDays >= DAYS_AVAILABLE_LONGEST) {
      debugsheet.appendRow([new Date(), `クライアントノートが存在しない、または利用日数が上限を超えているためスキップ`]);
      continue;
    }
    
    // クライアントノートのスプレッドシートを開く
    const clientSS = SpreadsheetApp.openByUrl(clientNoteURL);
    // 【集計】全期間積み上げﾚﾎﾟｰﾄ
    const allWeeklyPiledUpSheet = clientSS.getSheetByName(SHEET_NAMES_CLIENTSUPPORT.ALL_WEEKLY_PILED_UP);
   
    // 【集計】全期間積み上げﾚﾎﾟｰﾄが存在しないクライアントノートは古いノートのためPDF生成&リンク書き込み処理を行わない
    if (!allWeeklyPiledUpSheet) {
      debugsheet.appendRow([new Date(), `クライアントノートが古いためスキップ`]);
      continue;
    }

    // PDF表示用リンクを生成
    const pdfLink = createPDFLink(clientSS.getId(), SHEET_NAMES_CLIENTSUPPORT.WEEKLY_PILED_UP);
    debugsheet.appendRow([new Date(), `PDF表示用リンクを生成完了: ${pdfLink}`]);
    
    // 【集計】全期間積み上げﾚﾎﾟｰﾄのヘッダー
    const allWeeklyPiledUpHeaders = allWeeklyPiledUpSheet.getRange(ALL_WEEKLY_PILED_UP_HEADER_ROW, 1, 1, allWeeklyPiledUpSheet.getLastColumn()).getValues()[0];

    const useWeek = Math.floor(useDays / 7) + 1;

    // 「積み上げレポートURL」列を特定
    const urlColIndex = allWeeklyPiledUpHeaders.indexOf("積み上げレポートURL");

    // 書き込み対象の行を算出（WEEK1が4行目に対応）
    const targetRow = ALL_WEEKLY_PILED_UP_HEADER_ROW + useWeek;
    debugsheet.appendRow([new Date(), `PDF表示用リンク書き込み対象行: ${targetRow}`]);

    // もし行が26週を超えていたらスキップ(上で利用期間を用いて弾いてはいる)
    if (useWeek > 26) {
      debugsheet.appendRow([new Date(), `上限週を超えているためスキップ`]);
      continue;
    }

    debugsheet.appendRow([new Date(), `PDF表示用リンク書き込み処理実行`]);
    allWeeklyPiledUpSheet.getRange(targetRow, urlColIndex + 1).setValue(pdfLink);
  
  }

  debugsheet.appendRow([new Date(), "全期間積み上げレポート作成 終了" ]);
}
