/**
 * 認証用メール送信
 * 
 * メール認証登録・更新フォームからの処理を受け付ける
 * 
 * 登録の場合認証情報TBLやクライアント・コーチ名簿からの情報を元に
 * ステータスを判定し、認証不要の場合は処理を止めてユーザーに通知する
 * 
 * トークンを発行し、認証メールのリンクを作成して、
 * メールを送信する。
 * 
 * @param sourceSheet メール認証登録or更新フォームで入力された内容が出力されたシート
 * @param lastRow sourceSheetの最終行(行番号)
 * @param registFlg 登録フラグ(ON:登録, OFF:更新)
 * @param userType 利用者種別(クライアント・コーチ) 更新の場合は名簿を元に判断するのでnullが入る
 * 
 */
function authSendMail(sourceSheet, lastRow, registFlg, userType) {
  Logger.log("メール認証送信処理 start");

  // メール認証フォームに紐づくシートのヘッダー
  const mailAuthHeaders = GET_HEADER (sourceSheet, 1);
  const mailAuthInputData = GET_INPUT_DATA (sourceSheet, lastRow);
  
  // データ列を取得
  const emailColIndex = mailAuthHeaders.indexOf("メールアドレス");
  const answerIdColIndex = mailAuthHeaders.indexOf("回答者ID");

  // 認証で使用するデータを取得
  const inputEmail = mailAuthInputData[emailColIndex];
  const inputAnswerId = String(mailAuthInputData[answerIdColIndex]);

  if (!inputEmail || !inputEmail.includes("@")) {
    Logger.log(`⚠️ 有効なメールアドレスが取得できません。入力されたメールアドレス: ${inputEmail}`);
    return;
  }

  // 認証情報テーブルを取得
  const authTBL = GET_AUTH_SHEET();
  const authHeader = GET_HEADER (authTBL, 1);
  const authAllData = GET_All_DATA(authTBL);

  // 認証情報テーブルのデータ列を取得
  const authAnswerIdColIndex = authHeader.indexOf(AUTH_LIST_TBL.ANSWER_ID);
  const authMailColIndex = authHeader.indexOf(AUTH_LIST_TBL.MAIL);
  const tokenColIndex = authHeader.indexOf(AUTH_LIST_TBL.TOKEN);
  const authStatusColIndex = authHeader.indexOf(AUTH_LIST_TBL.AUTH_STATUS);
  const authLineIdColIndex = authHeader.indexOf(AUTH_LIST_TBL.LINE_ID);
  const authLineNameColIndex = authHeader.indexOf(AUTH_LIST_TBL.LINE_NAME);
  const userTypeColIndex = authHeader.indexOf(AUTH_LIST_TBL.USER_TYPE);
  const authTypeColIndex = authHeader.indexOf(AUTH_LIST_TBL.AUTH_TYPE);
  const authDeleteFlgColIndex = authHeader.indexOf(AUTH_LIST_TBL.DELETE_FLG);
  const authCreateDateColIndex = authHeader.indexOf(AUTH_LIST_TBL.CREATE_DATE);
  const authUpdateDateColIndex = authHeader.indexOf(AUTH_LIST_TBL.UPDATE_DATE);

  // コーチ名簿のデータを取得
  const coachTBL = GET_COACH_SHEET();
  const coachHeader = GET_HEADER(coachTBL, 1);
  const coachAllData = GET_All_DATA(coachTBL);

  // コーチ名簿のデータ列を取得
  const coachAnswerIdColIndex = coachHeader.indexOf(COACH_LIST_TBL.ANSWER_ID);
  const coachLineIdColIndex = coachHeader.indexOf(COACH_LIST_TBL.LINE_ID);
  const coachLineNameColIndex = coachHeader.indexOf(COACH_LIST_TBL.LINE_NAME);
  const coachMailColIndex = coachHeader.indexOf(COACH_LIST_TBL.MAIL);

  // クライアント名簿のデータを取得
  const clientTBL = GET_CLIENT_SHEET();
  const clientHeader = GET_HEADER(clientTBL, 1);
  const clientAllData = GET_All_DATA(clientTBL);

  // クライアント名簿のデータ列を取得
  const clientAnswerIdColIndex = clientHeader.indexOf(CLIENT_LIST_TBL.ANSWER_ID);
  const clientLineIdColIndex = clientHeader.indexOf(CLIENT_LIST_TBL.LINE_ID);
  const clientLineNameColIndex = clientHeader.indexOf(CLIENT_LIST_TBL.LINE_NAME);
  const clientMailColIndex = clientHeader.indexOf(CLIENT_LIST_TBL.MAIL);

  // ==========認証TBLにinsertする項目==========
  let inputAuthLineId = "";
  let inputAuthLineName = "";
  let inputUserType = "";
  // =========================================


  if (registFlg) {
    // メール認証(登録)
    Logger.log(`メール認証(登録処理)開始 registFlg: ${registFlg}`);

    if (userType === USER_TYPE.CLIENT) {
      Logger.log(`メール認証対象 : ${userType}`);

      // 回答者IDに紐づくクライアントレコードを取得
      let clientExistFlg = false;
      let targetClientEmail = "";
      let clientLineId = "";
      for (let i = 0; i < clientAllData.length; i++) {
        if (clientAllData[i][clientAnswerIdColIndex] == inputAnswerId) {
          clientExistFlg = true;
          targetClientEmail = clientAllData[i][clientMailColIndex];
          clientLineId = clientAllData[i][clientLineIdColIndex];
          Logger.log(`回答者IDに紐付くレコードがクライアント名簿 ${i + 2}行目に存在`);
          break;
        }
      }

      if (clientExistFlg) {
        // 回答者IDに紐づくデータがクライアント名簿に存在

        if (!clientLineId) {
          // クライアント名簿のデータ破損(LINE IDが登録されていない)
          Logger.log(`回答者IDに紐付くレコードが破損 LINE IDなし メールにて通知`);
          try {
            GmailApp.sendEmail(inputEmail, "【ボディメイクナビ】メール認証失敗",
              `メール認証に失敗しました。\n\n` +
              `お手数ですが以下のフォームよりお問い合わせ願います。\n\n` +
              `${INQUIRY_FORM}`);
              Logger.log(`✅ 失敗メール送信完了: ${inputEmail}`);
            } catch (e) {
              Logger.log(`⚠️ メール送信失敗: ${inputEmail} - ${e.toString()}`);
            }
          return;
        }

        let alreadyMsg = ``;

        if (targetClientEmail === inputEmail) {
          // クライアント名簿に登録されてるメールアドレスと入力されたメールアドレスが一致
          // 認証・登録済み案内
          Logger.log(`認証済み・登録済みをLINEで通知`);
          alreadyMsg = `メール認証とクライアント登録が既に完了しています。`;

        } else {
          // クライアント名簿に登録されてるメールアドレスと入力されたメールアドレスが不一致
          // メール更新案内
          Logger.log(`メールアドレスが異なるため、メールアドレス更新フォームをLINEで通知`);
          alreadyMsg = `メールアドレスの変更はこちらから行ってください。\n\n` +
          `${LINE_FROM_URLS.MAIL_UPDATE}`;

        }

        sendLINEMessage(clientLineId, alreadyMsg);
        return;


      } else {
        Logger.log(`回答者IDに紐付くレコードがクライアント名簿に存在なし。認証済みか特定開始`);
        let authExistFlg = false;
        let authAnswerId = "";
        let authStatus = "";
        let token = "";
        let authLineId = "";
        let authLineName = "";
        let userType = "";

        let targetRow = null;
        let targetData = null;

        for (let j = 0; j < authAllData.length; j++) {
          // メール認証フォームから送信されたメールアドレスと同じメールアドレスがあるか
          targetData = authAllData[j];
          if (targetData[authMailColIndex] === inputEmail 
              && targetData[authTypeColIndex] === AUTH_TYPE.REGIST
              && targetData[authDeleteFlgColIndex] !== DELETE_FLG.ON) {
                targetRow = j + 2; //authAllDataがヘッダーの1行目がないためとindexが0からに合わせて+2
                Logger.log(`送信されたメールアドレスが既に認証TBL ${targetRow}行目に存在`);
                authAnswerId = targetData[authAnswerIdColIndex];
                authStatus = targetData[authStatusColIndex];

                // ====== データ破損判定用 =====
                token = targetData[tokenColIndex];
                authLineId = targetData[authLineIdColIndex];
                authLineName = targetData[authLineNameColIndex];
                userType = targetData[userTypeColIndex];
                // ==========================

                authExistFlg = true;
                break;
            }
        }

        if (authExistFlg) {
          // メール認証フォームより送信されたメールアドレスに紐付くデータが
          // 認証情報TBLに存在

          // 送信されたメールアドレスが本人のものか判定
          if (authAnswerId != inputAnswerId) {
            Logger.log(`認証TBLに存在するメールアドレスの回答者IDがフォーム入力者と異なる`);
            GmailApp.sendEmail(inputEmail, "【ボディメイクナビ】メール認証失敗",
            `入力されたメールアドレスは既に別のLINEアカウントで使用されています。\n\n` +
            `お手数ですが以下のフォームよりお問い合わせ願います。\n\n` +
            `${INQUIRY_FORM}`);
            return;
          }

          // 既に認証済みか判定
          if (authStatus === AUTH_STATUS.COMPLETE) {
            // 認証情報が破損していないか判定
            if (token && authLineId && authLineName && userType) {
              // 認証情報が破損していない場合は認証済みであることとクライアント登録を促す
              const msg = `既にメール認証は完了しています。\n\n` +
              `以下のフォームよりクライアント登録を進めてください。\n\n` +
              `${LINE_FROM_URLS.CLIENT_REGIST}`;
              sendLINEMessage(authLineId, msg);
              return;
            }
          } 
          // 認証済みではない場合 or 認証情報が破損している場合
          // 該当の認証情報を論理削除し以降処理を行う
          authTBL.getRange(targetRow, authDeleteFlgColIndex + 1).setValue(DELETE_FLG.ON);
          authTBL.getRange(targetRow, authUpdateDateColIndex + 1).setValue(new Date()).setNumberFormat("yyyy/MM/dd HH:mm:ss");

        }

      }

    } else if (userType === USER_TYPE.COACH) {  
      Logger.log(`メール認証対象 : ${userType}`);

      // 回答者IDに紐づくコーチレコードを取得
      let coachExistFlg = false;
      let targetCoachEmail = "";
      let coachLineId = "";
      for (let i = 0; i < coachAllData.length; i++) {
        if (coachAllData[i][coachAnswerIdColIndex] == inputAnswerId) {
          coachExistFlg = true;
          targetCoachEmail = coachAllData[i][coachMailColIndex];
          coachLineId = coachAllData[i][coachLineIdColIndex];
          Logger.log(`回答者IDに紐付くレコードがコーチ名簿 ${i + 2}行目に存在`);
          break;
        }
      }

      if (coachExistFlg) {
        // 回答者IDに紐づくデータがコーチ名簿に存在

        if (!coachLineId) {
          // コーチ名簿のデータ破損(LINE IDが登録されていない)
          Logger.log(`回答者IDに紐付くレコードが破損 LINE IDなし メールにて通知`);
          try {
            GmailApp.sendEmail(inputEmail, "【ボディメイクナビ】メール認証失敗",
              `メール認証に失敗しました。\n\n` +
              `お手数ですが以下のフォームよりお問い合わせ願います。\n\n` +
              `${INQUIRY_FORM}`);
              Logger.log(`✅ 失敗メール送信完了: ${inputEmail}`);
            } catch (e) {
              Logger.log(`⚠️ メール送信失敗: ${inputEmail} - ${e.toString()}`);
            }
          return;
        }

        let alreadyMsg = ``;

        if (targetCoachEmail === inputEmail) {
          // コーチ名簿に登録されてるメールアドレスと入力されたメールアドレスが一致
          // 認証・登録済み案内
          Logger.log(`認証済み・登録済みをLINEで通知`);
          alreadyMsg = `メール認証とコーチ登録が既に完了しています。`;

        } else {
          // コーチ名簿に登録されてるメールアドレスと入力されたメールアドレスが不一致
          // メール更新案内
          Logger.log(`メールアドレスが異なるため、メールアドレス更新フォームをLINEで通知`);
          alreadyMsg = `メールアドレスの変更はこちらから行ってください。\n\n` +
          `${LINE_FROM_URLS.MAIL_UPDATE}`;

        }

        sendLINEMessage(coachLineId, alreadyMsg);
        return;

      } else {
        Logger.log(`回答者IDに紐付くレコードがコーチ名簿に存在なし。認証済みか特定開始`);
        let authExistFlg = false;
        let authAnswerId = "";
        let authStatus = "";
        let token = "";
        let authLineId = "";
        let authLineName = "";
        let userType = "";

        let targetRow = null;
        let targetData = null;

        for (let j = 0; j < authAllData.length; j++) {
          // メール認証フォームから送信されたメールアドレスと同じメールアドレスがあるか
          targetData = authAllData[j];
          if (targetData[authMailColIndex] === inputEmail 
              && targetData[authTypeColIndex] === AUTH_TYPE.REGIST
              && targetData[authDeleteFlgColIndex] !== DELETE_FLG.ON) {
                targetRow = j + 2; //authAllDataがヘッダーの1行目がないためとindexが0からに合わせて+2
                Logger.log(`送信されたメールアドレスが既に認証TBL ${targetRow}行目に存在`);
                authAnswerId = targetData[authAnswerIdColIndex];
                authStatus = targetData[authStatusColIndex];

                // ====== データ破損判定用 =====
                token = targetData[tokenColIndex];
                authLineId = targetData[authLineIdColIndex];
                authLineName = targetData[authLineNameColIndex];
                userType = targetData[userTypeColIndex];
                // ==========================

                authExistFlg = true;
                break;
            }
        }
        if (authExistFlg) {
          Logger.log(`メールアドレスが認証情報TBLに存在あり。ステータスを特定`);
          // メール認証フォームより送信されたメールアドレスに紐付くデータが
          // 認証情報TBLに存在

          // 送信されたメールアドレスが本人のものか判定
          if (authAnswerId != inputAnswerId) {
            Logger.log(`認証TBLに存在するメールアドレスの回答者IDがフォーム入力者と異なる`);
            GmailApp.sendEmail(inputEmail, "【ボディメイクナビ】メール認証失敗",
            `入力されたメールアドレスは既に別のLINEアカウントで使用されています。\n\n` +
            `お手数ですが以下のフォームよりお問い合わせ願います。\n\n` +
            `${INQUIRY_FORM}`);
            return;
          }

          // 既に認証済みか判定
          if (authStatus === AUTH_STATUS.COMPLETE) {
            // 認証情報が破損していないか判定
            if (token && authLineId && authLineName && userType) {
              // 認証情報が破損していない場合は認証済みであることとクライアント登録を促す
              const msg = `既にメール認証は完了しています。\n\n` +
              `以下のフォームよりコーチ登録を進めてください。\n\n` +
              `${LINE_FROM_URLS.COACH_REGIST}`;
              sendLINEMessage(authLineId, msg);
              return;
            }
          } 
          // 認証済みではない場合 or 認証情報が破損している場合
          // 該当の認証情報を論理削除し以降処理を行う
          Logger.log(`既に登録されている認証情報を論理削除し、以降処理を行う`);
          authTBL.getRange(targetRow, authDeleteFlgColIndex + 1).setValue(DELETE_FLG.ON);
          authTBL.getRange(targetRow, authUpdateDateColIndex + 1).setValue(new Date()).setNumberFormat("yyyy/MM/dd HH:mm:ss");
        }
      }

    } else {
      Logger.log(`想定外のメール認証対象 : ${userType}`);
      try {
        GmailApp.sendEmail(inputEmail, "【ボディメイクナビ】メール認証失敗",
          `メール認証に失敗しました。\n\n` +
          `お手数ですが以下のフォームよりお問い合わせ願います。\n\n` +
          `${INQUIRY_FORM}`);
          Logger.log(`✅ 失敗メール送信完了: ${inputEmail}`);
        } catch (e) {
          Logger.log(`⚠️ メール送信失敗: ${inputEmail} - ${e.toString()}`);
        }
        return;
    }

    // 認証情報TBLにinsertするデータをセット
    inputAuthType = AUTH_TYPE.REGIST;
    inputUserType = userType;

  } else {
    // メール認証更新
    Logger.log(`メール認証(更新処理)開始 registFlg: ${registFlg}`);

    // 回答者IDに紐づくクライアントレコードを取得
    let clientExistFlg = false;
    for (let i = 0; i < clientAllData.length; i++) {
      if (clientAllData[i][clientAnswerIdColIndex] == inputAnswerId) {
        clientExistFlg = true;
        inputAuthLineId = clientAllData[i][clientLineIdColIndex];
        inputAuthLineName  = clientAllData[i][clientLineNameColIndex];
        Logger.log(`回答者IDに紐付くレコードがクライアント名簿 ${i + 2}行目に存在`);
        break;
      }
    }

    // 回答者IDに紐づくコーチレコードを取得
    let coachExistFlg = false;
    for (let i = 0; i < coachAllData.length; i++) {
      if (coachAllData[i][coachAnswerIdColIndex] == inputAnswerId) {
        coachExistFlg = true;
        inputAuthLineId = coachAllData[i][coachLineIdColIndex];
        inputAuthLineName = coachAllData[i][coachLineNameColIndex];
        Logger.log(`回答者IDに紐付くレコードがコーチ名簿 ${i + 2}行目に存在`);
        break;
      }
    }

    // クライアント・コーチ両名簿に存在しない場合未登録
    // メール更新に失敗・アカウントがボディメイクナビに登録されていないをメール通知
    if (!clientExistFlg && !coachExistFlg) {
      Logger.log(`回答者ID ${answerId}の名簿が存在しません。`)
      try {
        GmailApp.sendEmail(inputEmail, "【ボディメイクナビ】メール認証失敗",
          `メール認証に失敗しました。\n\n` +
          `アカウントがボディメイクナビに登録されていません。`);
          Logger.log(`✅ 失敗メール送信完了: ${inputEmail}`);
        } catch (e) {
          Logger.log(`⚠️ メール送信失敗: ${inputEmail} - ${e.toString()}`);
        }
      return;
    }

    // クライアント名簿に存在していて、コーチ名簿に存在しない場合はクライアント
    if (clientExistFlg && !coachExistFlg) {
      inputUserType = USER_TYPE.CLIENT;
    }
    // コーチ名簿に存在していて、クライアント名簿に存在しない場合はコーチ
    if (!clientExistFlg && coachExistFlg) {
      inputUserType = USER_TYPE.COACH;
    }
    // 両名簿に存在する場合は両種別
    if (clientExistFlg && coachExistFlg) {
      inputUserType = USER_TYPE.BOTH;
    }
    inputAuthType = AUTH_TYPE.UPDATE;
  }

  // トークン発行
  let token = "";
  let isDuplicate = true;

  while (isDuplicate) {
    token = Utilities.getUuid();
    isDuplicate = false;

    for (let i = 0; i < authAllData.length; i++) {
      if (authAllData[i][tokenColIndex] === token) {
        isDuplicate = true;
        break;
      }
    }
  }
  Logger.log(`トークン発行: ${token}`);

  // 認証情報TBLにinsert
  const lastRowBeforeInsert = authTBL.getLastRow();
  authTBL.insertRowAfter(lastRowBeforeInsert); // 最終行に1行追加
  const insertTargetRow = lastRowBeforeInsert + 1; // insert対象行

  authTBL.getRange(insertTargetRow, authAnswerIdColIndex + 1).setValue(inputAnswerId);
  authTBL.getRange(insertTargetRow, authMailColIndex + 1).setValue(inputEmail);
  authTBL.getRange(insertTargetRow, tokenColIndex + 1).setValue(token);
  authTBL.getRange(insertTargetRow, authStatusColIndex + 1).setValue(AUTH_STATUS.STILL);
  authTBL.getRange(insertTargetRow, authLineIdColIndex + 1).setValue(inputAuthLineId);
  authTBL.getRange(insertTargetRow, authLineNameColIndex + 1).setValue(inputAuthLineName);
  authTBL.getRange(insertTargetRow, userTypeColIndex + 1).setValue(inputUserType);
  authTBL.getRange(insertTargetRow, authTypeColIndex + 1).setValue(registFlg ?  AUTH_TYPE.REGIST : AUTH_TYPE.UPDATE);
  authTBL.getRange(insertTargetRow, authDeleteFlgColIndex + 1).setValue("");
  authTBL.getRange(insertTargetRow, authCreateDateColIndex + 1).setValue(new Date()).setNumberFormat("yyyy/MM/dd HH:mm");
  authTBL.getRange(insertTargetRow, authUpdateDateColIndex + 1).setValue(new Date()).setNumberFormat("yyyy/MM/dd HH:mm");


  // 追加した行の各セルを囲う
  const rowRange = authTBL.getRange(insertTargetRow, 1, 1, authTBL.getLastColumn());
  rowRange.setBackground("#ffffff");
  rowRange.setBorder(true, true, true, true, true, true);


  // メール認証リンク作成 登録はLIFF_APP_URL、更新はWEB_APP_URL
  const URL = registFlg ? LIFF_APP_URL : WEB_APP_URL;
  const authUrl = `${URL}?token=${token}`;

  // メール送信
  // 認証メール送信
  try {
    GmailApp.sendEmail(inputEmail, "【ボディメイクナビ】メール認証のお願い",
      `以下のリンクをクリックして認証を完了してください：\n\n${authUrl}\n\n` +
      "※このメールに心当たりがない場合は無視してください。");
    Logger.log(`✅ 認証メール送信完了: ${inputEmail}`);
  } catch (e) {
    Logger.log(`⚠️ メール送信失敗: ${inputEmail} - ${e.toString()}`);
  }
  Logger.log("メール認証送信処理 end");
  return;
}

/**
 * メール認証処理(更新導線)
 * 
 * パラメータのトークンを元に、対象の認証情報を取得し
 * 合えば認証OKとする
 */
function doGet(e) {
  const debugsheet = GET_LOG_SHEET();

  try {
    debugsheet.appendRow([new Date(), "メールアドレス認証処理 start"]);
  } catch (err) {
    return ContentService.createTextOutput("❌ appendRow失敗: " + err.message);
  }
  // 値の取り出し
  const token = e.parameter.token;
  // トークン取得失敗
  if (!token) {
    debugsheet.appendRow([new Date(), 'トークンが取得できませんでした。']);
    return HtmlService.createHtmlOutput(`
      <html>
        <body style="font-family: sans-serif; text-align: center; padding: 40px;">
          <h2>メールアドレス認証失敗</h2>
          <p style="font-size: 16px; margin-top: 20px;">
            LINEに戻り、再度メール認証フォームより認証を行なってください。
          </p>
        </body>
      </html>`);
  }

  // 認証情報テーブルを取得
  const authTBL = GET_AUTH_SHEET();
  const authHeader = GET_HEADER (authTBL, 1);
  const authAllData = GET_All_DATA(authTBL);

  // 認証情報テーブルのデータ列を取得
  const authAnswerIdColIndex = authHeader.indexOf(AUTH_LIST_TBL.ANSWER_ID);
  const authMailColIndex = authHeader.indexOf(AUTH_LIST_TBL.MAIL);
  const tokenColIndex = authHeader.indexOf(AUTH_LIST_TBL.TOKEN);
  const authStatusColIndex = authHeader.indexOf(AUTH_LIST_TBL.AUTH_STATUS);
  const authLineIdColIndex = authHeader.indexOf(AUTH_LIST_TBL.LINE_ID);
  const authLineNameColIndex = authHeader.indexOf(AUTH_LIST_TBL.LINE_NAME);
  const userTypeColIndex = authHeader.indexOf(AUTH_LIST_TBL.USER_TYPE);
  const authTypeColIndex = authHeader.indexOf(AUTH_LIST_TBL.AUTH_TYPE);
  const authDeleteFlgColIndex = authHeader.indexOf(AUTH_LIST_TBL.DELETE_FLG);
  const authCreateDateColIndex = authHeader.indexOf(AUTH_LIST_TBL.CREATE_DATE);
  const authUpdateDateColIndex = authHeader.indexOf(AUTH_LIST_TBL.UPDATE_DATE);

  // トークンに紐づく認証情報を取得する
  let updateTargetRow = null;
  let updateTargetData = null;

  for (let i = 0; i < authAllData.length; i++) {
    if (authAllData[i][tokenColIndex] === token
        && authAllData[i][authDeleteFlgColIndex] !== DELETE_FLG.ON) {
      updateTargetRow = i + 2;
      updateTargetData = authAllData[i];
      debugsheet.appendRow([new Date(), `トークンに紐づく認証情報を発見: ${updateTargetRow}行`]);
      break;
    }      
  }

  // トークンに紐づく認証情報が存在しない場合エラー画面表示
  if (!updateTargetData) {
    debugsheet.appendRow([new Date(), `トークンに紐づく認証情報が存在なし`]);
    return HtmlService.createHtmlOutput(`
      <html>
        <body style="font-family: sans-serif; text-align: center; padding: 40px;">
          <h2>メールアドレス認証失敗</h2>
          <p style="font-size: 16px; margin-top: 20px;">
            LINEに戻り、再度メール認証フォームより認証を行なってください。
          </p>
        </body>
      </html>`);
  }

  const authAnswerId = updateTargetData[authAnswerIdColIndex];
  const authMail = updateTargetData[authMailColIndex];
  const authStatus = updateTargetData[authStatusColIndex];
  const authLineId = updateTargetData[authLineIdColIndex];
  const authLineName = updateTargetData[authLineNameColIndex];
  const authUserType = updateTargetData[userTypeColIndex];
  const authType = updateTargetData[authTypeColIndex];
  const authCreateDate = updateTargetData[authCreateDateColIndex];
  const authUpdateDate = updateTargetData[authUpdateDateColIndex];

  // データが欠けていないか確認
  // 欠けている場合論理削除
  if (!authAnswerId || !authMail || !authStatus || !authLineId || !authLineName
      || !authUserType || !authType || !authCreateDate || !authUpdateDate) {
        debugsheet.appendRow([new Date(), `トークンに紐づく認証情報が不正`]);
        authTBL.getRange(updateTargetRow, authDeleteFlgColIndex + 1).setValue(DELETE_FLG.ON);
        authTBL.getRange(updateTargetRow, authUpdateDateColIndex + 1).setValue(new Date()).setNumberFormat("yyyy/MM/dd HH:mm:ss");
        return HtmlService.createHtmlOutput(`
          <html>
            <body style="font-family: sans-serif; text-align: center; padding: 40px;">
              <h2>メールアドレス認証失敗</h2>
              <p style="font-size: 16px; margin-top: 20px;">
                LINEに戻り、再度メール認証フォームより認証を行なってください。
              </p>
            </body>
          </html>`);
  }

  // 認証ステータスが認証完了していないか確認(ダブルタップされて以降処理をやらないため)
  if (authStatus === AUTH_STATUS.COMPLETE) {
    debugsheet.appendRow([new Date(), `トークンに紐づく認証情報が認証済み`]);
    sendLINEMessage(authLineId, `メールアドレス認証は既に完了しています。`);
    return HtmlService.createHtmlOutput(`
        <html>
          <body style="font-family: sans-serif; text-align: center; padding: 40px;">
            <h2>メールアドレス認証済み</h2>
            <p style="font-size: 16px; margin-top: 20px;">
              メールアドレス認証は既に完了しています。
            </p>
          </body>
        </html>`);
  }

  // エラーが出なかったため認証完了にする
  debugsheet.appendRow([new Date(), `認証情報を認証済みにする`]);
  authTBL.getRange(updateTargetRow, authStatusColIndex + 1).setValue(AUTH_STATUS.COMPLETE);
  authTBL.getRange(updateTargetRow, authUpdateDateColIndex + 1).setValue(new Date()).setNumberFormat("yyyy/MM/dd HH:mm:ss");


  // 名簿からメールアドレスの更新とサポートノートの権限変更を行う
  // コーチ名簿のデータを取得
  const coachTBL = GET_COACH_SHEET();
  const coachHeader = GET_HEADER(coachTBL, 1);
  const coachAllData = GET_All_DATA(coachTBL);

  // コーチ名簿のデータ列を取得
  const coachAnswerIdColIndex = coachHeader.indexOf(COACH_LIST_TBL.ANSWER_ID);
  const coachLineIdColIndex = coachHeader.indexOf(COACH_LIST_TBL.LINE_ID);
  const coachMailColIndex = coachHeader.indexOf(COACH_LIST_TBL.MAIL);
  const coachNoColIndex = coachHeader.indexOf(COACH_LIST_TBL.COACH_NO);

  // クライアント名簿のデータを取得
  const clientTBL = GET_CLIENT_SHEET();
  const clientHeader = GET_HEADER(clientTBL, 1);
  const clientAllData = GET_All_DATA(clientTBL);

  // クライアント名簿のデータ列を取得
  const clientAnswerIdColIndex = clientHeader.indexOf(CLIENT_LIST_TBL.ANSWER_ID);
  const clientLineIdColIndex = clientHeader.indexOf(CLIENT_LIST_TBL.LINE_ID);
  const clientMailColIndex = clientHeader.indexOf(CLIENT_LIST_TBL.MAIL);
  const clientResCoachNoColIndex = clientHeader.indexOf(CLIENT_LIST_TBL.RES_COACH_NO);
  const clientNoteUrlColIndex = clientHeader.indexOf(CLIENT_LIST_TBL.CLIENT_NOTE_URL);

  // 名簿のメールアドレスを更新
  // クライアントノートの権限を変更したメールアドレスへ変更
  let clientMail = "";
  let coachMail = "";

  let resCoachNo = "";
  let clientNoteUrl = "";

  let coachNo = "";

  let clientNoteList = [];

  let targetClientData = null;
  let targetClientRow = null;
  let targetCoachData = null;
  let targetCoachRow = null;

  if (authUserType === USER_TYPE.CLIENT) {

      clientMail = authMail;

      // 回答者IDを元にメール更新対象のクライアントを取得
      for (let i = 1; i < clientAllData.length; i++) {
        if (clientAllData[i][clientAnswerIdColIndex] == authAnswerId) {
          targetClientRow = i + 2;
          targetClientData = clientAllData[i];
          debugsheet.appendRow([new Date(), `回答者IDを元にメール更新対象のクライアントを取得: ${targetClientRow}行`]);
          break;
        }
      }

      // 対象が存在しない場合エラー
      if (!targetClientData) {
        debugsheet.appendRow([new Date(), `回答者IDに紐づくクライアントが存在しない。`]);
        sendLINEMessage(authLineId, `メール認証に失敗しました。\n\n` +
              `お手数ですが以下のフォームよりお問い合わせ願います。\n\n` +
              `${INQUIRY_FORM}`)
        return;
      }
      
      // クライアント名簿のメールアドレスを更新
      debugsheet.appendRow([new Date(), `メール更新対象であるクライアントのメールアドレスを更新`]);
      clientTBL.getRange(targetClientRow, clientMailColIndex + 1).setValue(clientMail);

      // 取得したコーチNoとクライアントノートを格納
      resCoachNo = targetClientData[clientResCoachNoColIndex];
      clientNoteUrl = targetClientData[clientNoteUrlColIndex];

      // メール更新したクライアントの、コーチのメールアドレスをコーチNoを元に取得
      for (let i = 0; i < coachAllData.length; i++) {
        if (coachAllData[i][coachNoColIndex] === resCoachNo) {
          targetCoachRow = i + 2;
          targetCoachData = coachAllData[i];
          debugsheet.appendRow([new Date(), `コーチNo.を元にメール更新対象のクライアントのコーチを取得: ${targetClientRow}行`]);
        }
      }

      // 対象が存在しない場合エラー
      if (!targetCoachData) {
        debugsheet.appendRow([new Date(), `コーチNoに紐づくコーチが存在しない。`]);
        sendLINEMessage(authLineId, `メール認証に失敗しました。\n\n` +
              `お手数ですが以下のフォームよりお問い合わせ願います。\n\n` +
              `${INQUIRY_FORM}`)
        return;
      }
      // 取得したコーチのメールを格納
      coachEmail = targetCoachData[coachMailColIndex];

      // クライアントノートのファイルIDを取得
      const match = clientNoteUrl.match(/[-\w]{25,}/);
      if (!match) throw new Error(`📛 clientNoteUrlの形式が不正: ${clientNoteUrl}`);
      const fileId = match[0];
      const clientNoteFile = DriveApp.getFileById(fileId);

      // 更新したメールアドレスとコーチのメールアドレスでサポートノートの権限を変更する
      debugsheet.appendRow([new Date(), `クライアントノート:${clientNoteUrl}へ権限付与。コーチメール:${coachEmail},クライアントメール${clientMail}`]);
      allowClientNoteAccess(clientMail, coachEmail, clientNoteFile);
      debugsheet.appendRow([new Date(), `クライアントノートへの権限付与完了`]);


  } else if (authUserType === USER_TYPE.COACH) {

      coachMail = authMail;

      // 回答者IDを元にメール更新対象のクライアントを取得
      for (let i = 0; i < coachAllData.length; i++) {
        if (coachAllData[i][coachAnswerIdColIndex] === authAnswerId) {
          targetCoachRow = i + 2;
          targetCoachData = coachAllData[i];
          debugsheet.appendRow([new Date(), `回答者IDを元にメール更新対象のコーチを取得: ${targetCoachRow}行`]);
        }
      }

      // 対象が存在しない場合エラー
      if (!targetCoachData) {
        debugsheet.appendRow([new Date(), `回答者IDに紐づくコーチが存在しない。`]);
        sendLINEMessage(authLineId, `メール認証に失敗しました。\n\n` +
              `お手数ですが以下のフォームよりお問い合わせ願います。\n\n` +
              `${INQUIRY_FORM}`);
        return;
      }
      // コーチ名簿のメールアドレスを更新 
      debugsheet.appendRow([new Date(), `メール更新対象であるコーチのメールアドレスを更新`]);
      coachTBL.getRange(targetCoachRow, coachMailColIndex + 1).setValue(coachMail);
      // 取得したコーチNoを格納
      coachNo = targetCoachData[coachNoColIndex];

      // 担当クライアントのメールアドレスとクライアントノートを取得
      for (let i = 0; i < clientAllData.length; i++) {
        if (clientAllData[i][clientResCoachNoColIndex] === coachNo) {
          const tmpEmail = clientAllData[i][clientMailColIndex];
          const tmpNoteUrl = clientAllData[i][clientNoteUrlColIndex];

          if (tmpEmail && tmpNoteUrl) {
            clientNoteList.push({ tmpEmail, tmpNoteUrl });
            debugsheet.appendRow([new Date(), `担当クライアントのメールアドレス: ${tmpEmail}とクライアントノートURL${tmpNoteUrl}を取得 `]);
          }
        }
      }

      // 一括で権限変更
      clientNoteList.forEach(({ tmpEmail, noteUrl }) => {
        try {
          const fileId = noteUrl.match(/[-\w]{25,}/)[0];
          const file = DriveApp.getFileById(fileId);
          debugsheet.appendRow([new Date(), `クライアントノート:${noteUrl}へ権限付与。コーチメール:${coachMail},クライアントメール${tmpEmail}`]);
          allowClientNoteAccess(tmpEmail, coachMail, file);
        } catch (e) {
          debugsheet.appendRow([new Date(), `⚠️ ノート権限変更失敗: ${tmpEmail} - ${noteUrl} - ${e}`]);
          sendLINEMessage(authLineId, `クライアントノートの権限変更に失敗しました。\n\n` +
              `お手数ですが以下のフォームよりお問い合わせ願います。\n\n` +
              `${INQUIRY_FORM}`);
          return HtmlService.createHtmlOutput(`
            <html>
              <body style="font-family: sans-serif; text-align: center; padding: 40px;">
                <h2>クライアントノートの権限付与失敗</h2>
                <p style="font-size: 16px; margin-top: 20px;">
                  この画面は閉じて、LINEに戻り、お手数ですがお問い合わせフォームから報告してください。
                </p>
              </body>
            </html>`
          );
        }
      });

  } else if (authUserType === USER_TYPE.BOTH) { // クライアント、コーチ両権限

      // クライアントとしての処理
      clientMail = authMail;

      // 回答者IDを元にメール更新対象のクライアントを取得
      for (let i = 1; i < clientAllData.length; i++) {
        if (clientAllData[i][clientAnswerIdColIndex] == authAnswerId) {
          targetClientRow = i + 2;
          targetClientData = clientAllData[i];
          debugsheet.appendRow([new Date(), `回答者IDを元にメール更新対象のクライアントを取得: ${targetClientRow}行`]);
          break;
        }
      }

      // 対象が存在しない場合エラー
      if (!targetClientData) {
        debugsheet.appendRow([new Date(), `回答者IDに紐づくクライアントが存在しない。`]);
        sendLINEMessage(authLineId, `メール認証に失敗しました。\n\n` +
              `お手数ですが以下のフォームよりお問い合わせ願います。\n\n` +
              `${INQUIRY_FORM}`)
        return;
      }
      
      // クライアント名簿のメールアドレスを更新
      debugsheet.appendRow([new Date(), `メール更新対象であるクライアントのメールアドレスを更新`]);
      clientTBL.getRange(targetClientRow, clientMailColIndex + 1).setValue(clientMail);

      // 取得したコーチNoとクライアントノートを格納
      resCoachNo = targetClientData[clientResCoachNoColIndex];
      clientNoteUrl = targetClientData[clientNoteUrlColIndex];

      // メール更新したクライアントの、コーチのメールアドレスをコーチNoを元に取得
      for (let i = 0; i < coachAllData.length; i++) {
        if (coachAllData[i][coachNoColIndex] === resCoachNo) {
          targetCoachRow = i + 2;
          targetCoachData = coachAllData[i];
          debugsheet.appendRow([new Date(), `コーチNo.を元にメール更新対象のクライアントのコーチを取得: ${targetClientRow}行`]);
        }
      }

      // 対象が存在しない場合エラー
      if (!targetCoachData) {
        debugsheet.appendRow([new Date(), `コーチNoに紐づくコーチが存在しない。`]);
        sendLINEMessage(authLineId, `メール認証に失敗しました。\n\n` +
              `お手数ですが以下のフォームよりお問い合わせ願います。\n\n` +
              `${INQUIRY_FORM}`)
        return;
      }
      // 取得したコーチのメールを格納
      coachMail = targetCoachData[coachMailColIndex];

      // クライアントノートのファイルIDを取得
      const match = clientNoteUrl.match(/[-\w]{25,}/);
      if (!match) throw new Error(`📛 clientNoteUrlの形式が不正: ${clientNoteUrl}`);
      const fileId = match[0];
      const clientNoteFile = DriveApp.getFileById(fileId);

      // 更新したメールアドレスとコーチのメールアドレスでサポートノートの権限を変更する
      debugsheet.appendRow([new Date(), `クライアントノート:${clientNoteUrl}へ権限付与。コーチメール:${coachMail},クライアントメール${clientMail}`]);
      allowClientNoteAccess(clientMail, coachMail, clientNoteFile);
      debugsheet.appendRow([new Date(), `クライアントノートへの権限付与完了`]);


      // コーチとしての処理
      coachMail = authMail;

      // 回答者IDを元にメール更新対象のクライアントを取得
      for (let i = 0; i < coachAllData.length; i++) {
        if (coachAllData[i][coachAnswerIdColIndex] === authAnswerId) {
          targetCoachRow = i + 2;
          targetCoachData = coachAllData[i];
          debugsheet.appendRow([new Date(), `回答者IDを元にメール更新対象のコーチを取得: ${targetCoachRow}行`]);
        }
      }

      // 対象が存在しない場合エラー
      if (!targetCoachData) {
        debugsheet.appendRow([new Date(), `回答者IDに紐づくコーチが存在しない。`]);
        sendLINEMessage(authLineId, `メール認証に失敗しました。\n\n` +
              `お手数ですが以下のフォームよりお問い合わせ願います。\n\n` +
              `${INQUIRY_FORM}`)
        return;
      }
      // コーチ名簿のメールアドレスを更新 
      debugsheet.appendRow([new Date(), `メール更新対象であるコーチのメールアドレスを更新`]);
      coachTBL.getRange(targetCoachRow, coachMailColIndex + 1).setValue(coachMail);
      // 取得したコーチNoを格納
      coachNo = targetCoachData[coachNoColIndex];

      // 担当クライアントのメールアドレスとクライアントノートを取得
      for (let i = 0; i < clientAllData.length; i++) {
        if (clientAllData[i][clientResCoachNoColIndex] === coachNo) {
          const tmpEmail = clientAllData[i][clientMailColIndex];
          const tmpNoteUrl = clientAllData[i][clientNoteUrlColIndex];

          if (tmpEmail && tmpNoteUrl) {
            clientNoteList.push({ tmpEmail, tmpNoteUrl });
            debugsheet.appendRow([new Date(), `担当クライアントのメールアドレス: ${tmpEmail}とクライアントノートURL${tmpNoteUrl}を取得 `]);
          }
        }
      }

      // 一括で権限変更
      clientNoteList.forEach(({ tmpEmail, noteUrl }) => {
        try {
          const fileId = noteUrl.match(/[-\w]{25,}/)[0];
          const file = DriveApp.getFileById(fileId);
          debugsheet.appendRow([new Date(), `クライアントノート:${noteUrl}へ権限付与。コーチメール:${coachMail},クライアントメール${tmpEmail}`]);
          allowClientNoteAccess(tmpEmail, coachMail, file);
        } catch (e) {
          debugsheet.appendRow([new Date(), `⚠️ ノート権限変更失敗: ${tmpEmail} - ${noteUrl} - ${e}`]);
          sendLINEMessage(authLineId, `クライアントノートの権限変更に失敗しました。\n\n` +
              `お手数ですが以下のフォームよりお問い合わせ願います。\n\n` +
              `${INQUIRY_FORM}`);
          return HtmlService.createHtmlOutput(`
            <html>
              <body style="font-family: sans-serif; text-align: center; padding: 40px;">
                <h2>クライアントノートの権限付与失敗</h2>
                <p style="font-size: 16px; margin-top: 20px;">
                  この画面は閉じて、LINEに戻り、お手数ですがお問い合わせフォームから報告してください。
                </p>
              </body>
            </html>`
          );
        }
      });
  }

  // 変更完了メッセージ
  const msg = `メールアドレスの更新が完了しました！`;
  debugsheet.appendRow([new Date(), `更新完了メッセージをLINE ID : ${authLineId}へ送信`]);
  sendLINEMessage(authLineId, msg);

  debugsheet.appendRow([new Date(), "メールアドレス認証処理 end"]);

  return HtmlService.createHtmlOutput(`
    <html>
      <head>
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <style>
          body {
            background-color: #f5f5f5;
            font-family: "Helvetica Neue", sans-serif;
            margin: 0;
            padding: 60px 20px;
            text-align: center;
          }
          .box {
            background: #ffffff;
            border-radius: 14px;
            padding: 40px 24px;
            max-width: 90%;
            margin: 0 auto;
            box-shadow: 0 4px 16px rgba(0,0,0,0.1);
          }
          .checkmark {
            font-size: 64px;
            color: #4CAF50;
            margin-bottom: 16px;
          }
          h2 {
            font-size: 22px;
            margin-bottom: 10px;
            color: #333;
          }
          p {
            font-size: 16px;
            color: #333;
            margin: 0;
          }
          .close-msg {
            font-size: 13px;
            margin-top: 12px;
            color: #aaa;
          }
        </style>
        <script>
        // 3秒後にウィンドウを閉じる
        setTimeout(() => {
          window.close();
        }, 3000);
      </script>
      </head>
      <body>
        <div class="box">
          <div class="checkmark">✅</div>
          <h2>認証が完了しました</h2>
          <p>この画面は3秒後に自動で閉じます。</p>
          <p>閉じない場合は手動で閉じてください。</p>
        </div>
      </body>
    </html>
  `);
}
