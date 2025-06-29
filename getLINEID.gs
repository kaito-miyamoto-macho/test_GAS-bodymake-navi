/**
 * メール認証処理(登録導線)
 * 
 * パラメータのトークンを元に、対象の認証情報を取得し
 * 合えば認証OKとする
 * 
 * 利用者種別によって登録フォームを切り替え、パラメータで取得したLINE IDに送信する
 */
function doPost(e) {
  const debugsheet = GET_LOG_SHEET();
  try {
      debugsheet.appendRow([new Date(), 'メール認証処理(登録導線) start']);
      debugsheet.appendRow([new Date(), "Received request: " + JSON.stringify(e)]);

      // LIFFアプリで取得してきたLINE IDとLINE登録名、トークン
      const userId = e.parameter.userId;
      const displayName = e.parameter.displayName;
      const token = e.parameter.token;


      if (!userId || !displayName || !token) {
          debugsheet.appendRow([new Date(), `LIFFアプリからPOSTされた値が欠けています。`]);
          debugsheet.appendRow([new Date(), `userId: ${userId}, displayName: ${displayName}: token: ${token}`]);
          return createCORSResponse("Error: Missing parameters");
      }
      
      // 認証情報テーブルを取得
      const authTBL = GET_AUTH_SHEET();
      const authHeader = GET_HEADER (authTBL, 1);
      const authAllData = GET_All_DATA(authTBL);

      // 認証情報テーブルのデータ列を取得
      const tokenColIndex = authHeader.indexOf(AUTH_LIST_TBL.TOKEN);
      const authStatusColIndex = authHeader.indexOf(AUTH_LIST_TBL.AUTH_STATUS);
      const authLineIdColIndex = authHeader.indexOf(AUTH_LIST_TBL.LINE_ID);
      const authLineNameColIndex = authHeader.indexOf(AUTH_LIST_TBL.LINE_NAME);
      const userTypeColIndex = authHeader.indexOf(AUTH_LIST_TBL.USER_TYPE);
      const authDeleteFlgColIndex = authHeader.indexOf(AUTH_LIST_TBL.DELETE_FLG);
      const authUpdateDateColIndex = authHeader.indexOf(AUTH_LIST_TBL.UPDATE_DATE);

      // 認証情報TBLから認証対象行をトークンで取得
      let authTargetRow = null;
      let authTargetData = null;
      for (let i = 0; i < authAllData.length; i++) {
        if (authAllData[i][tokenColIndex] === token) {
          authTargetData = authAllData[i];
          authTargetRow = i + 2;
          debugsheet.appendRow([new Date(), `認証情報TBLから認証対象行を取得 ${authTargetRow}`]);
          break;
        }
      }

      // 認証更新対象行が存在しない場合認証失敗をLINE通知
      if (!authTargetRow) {
        debugsheet.appendRow([new Date(), 'メール認証更新対象行が見つからないためLINE通知して終了']);
        sendLINEMessage(userId, `メール認証に失敗しました。\n` +
        `再度メール認証をやり直してください。`);
        debugsheet.appendRow([new Date(), 'メール認証処理(登録導線) end']);
        return;
      }

      // 認証更新対象行の利用者種別が想定外の値の場合は先にエラーで弾く
      if (authTargetData[userTypeColIndex] !== USER_TYPE.CLIENT
          && authTargetData[userTypeColIndex] !== USER_TYPE.COACH) {
          // 利用者種別がクライアントかコーチ以外の場合エラーメッセージを送信し
          // 認証対象行を論理削除
          debugsheet.appendRow([new Date(), '認証済みのデータの整合性が取れないため論理削除実行']);
          authTBL.getRange(authTargetRow, authDeleteFlgColIndex + 1).setValue(DELETE_FLG.ON);
          authTBL.getRange(authTargetRow, authUpdateDateColIndex + 1).setValue(new Date()).setNumberFormat("yyyy/MM/dd HH:mm");

          const errMsg = `メール認証に失敗しました。\n\n` +
            `お手数ですが以下のフォームよりお問い合わせ願います。\n\n` +
            `${INQUIRY_FORM}`;

          debugsheet.appendRow([new Date(), 'LINE通知']);
          sendLINEMessage(userId, errMsg);
          debugsheet.appendRow([new Date(), 'メール認証処理(登録導線) end']);
          return;
      }

      // 認証更新対象行の認証ステータスを評価
      if (authTargetData[authStatusColIndex] === AUTH_STATUS.COMPLETE) {
        debugsheet.appendRow([new Date(), '認証更新対象行が既に認証済み。']);

        // 既に認証済みメッセージを利用者種別ごとで送信
        let msg = "";
        if (authTargetData[userTypeColIndex] === USER_TYPE.CLIENT) {
          msg = `メール認証は既に完了しています。\n` +
          `以下のフォームからクライアント登録を進めてください。\n\n` +
          `${LINE_FROM_URLS.CLIENT_REGIST}`;

        } else if (authTargetData[userTypeColIndex] === USER_TYPE.COACH) {
          msg = `メール認証は既に完了しています。\n` +
          `以下のフォームからコーチ登録を進めてください。\n\n` +
          `${LINE_FROM_URLS.COACH_REGIST}`;
        
        } 
        debugsheet.appendRow([new Date(), 'LINE通知']);
        sendLINEMessage(userId, msg);
        debugsheet.appendRow([new Date(), 'メール認証処理(登録導線) end']);
        return;

      } else if (authTargetData[authStatusColIndex] === AUTH_STATUS.STILL) {
        debugsheet.appendRow([new Date(), '認証の整合性がとれたため、認証ステータスを認証済みにし、LINE情報を入力']);

        authTBL.getRange(authTargetRow, authStatusColIndex + 1).setValue(AUTH_STATUS.COMPLETE);
        authTBL.getRange(authTargetRow, authLineIdColIndex + 1).setValue(userId);
        authTBL.getRange(authTargetRow, authLineNameColIndex + 1).setValue(displayName);
        authTBL.getRange(authTargetRow, authUpdateDateColIndex + 1).setValue(new Date()).setNumberFormat("yyyy/MM/dd HH:mm");

        // 認証完了と登録フォームをメッセージを利用者種別ごとで送信
        let compMsg = "";
        if (authTargetData[userTypeColIndex] === USER_TYPE.CLIENT) {
          compMsg = `✅ 認証が完了しました！\n\n` +
           `▼こちらからクライアント情報の登録をお願いします\n ` +
            `${LINE_FROM_URLS.CLIENT_REGIST}`;

        } else if (authTargetData[userTypeColIndex] === USER_TYPE.COACH) {
          compMsg = `✅ 認証が完了しました！\n\n` +
            `▼こちらからコーチ情報の登録をお願いします\n` +
            `${LINE_FROM_URLS.COACH_REGIST}`;
        
        } 
        debugsheet.appendRow([new Date(), '認証完了と登録フォームをLINE通知']);
        sendLINEMessage(userId, compMsg);
        debugsheet.appendRow([new Date(), 'メール認証処理(登録導線) end']);
        return;
        
      } else {
        // 認証ステータスが想定外の場合エラーメッセージを送信し
        // 認証対象行を論理削除
        debugsheet.appendRow([new Date(), '認証済みのデータの整合性が取れないため論理削除実行']);
        authTBL.getRange(authTargetRow, authDeleteFlgColIndex + 1).setValue(DELETE_FLG.ON);
        authTBL.getRange(authTargetRow, authUpdateDateColIndex + 1).setValue(new Date()).setNumberFormat("yyyy/MM/dd HH:mm");

        const errMsg = `メール認証に失敗しました。\n\n` +
          `お手数ですが以下のフォームよりお問い合わせ願います。\n\n` +
          `${INQUIRY_FORM}`;

        debugsheet.appendRow([new Date(), 'LINE通知']);
        sendLINEMessage(userId, errMsg);
        debugsheet.appendRow([new Date(), 'メール認証処理(登録導線) end']);
        return;

      }

  } catch (error) {
      Logger.log("Exception: " + error);
      debugsheet.appendRow([new Date(), `Exception: ${error}`]);
      return createCORSResponse("Error: " + error);
  }
}

// ✅ CORS対応のレスポンスを生成する関数
function createCORSResponse(message) {
    Logger.log("Response: " + message);
    var output = ContentService.createTextOutput(JSON.stringify({ result: message }));
    output.setMimeType(ContentService.MimeType.JSON);
    return output;
}
