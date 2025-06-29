/**
 * 
 * 共通の定数を管理するファイル
 * 
 */

/**
 * テスト環境判定用フラグ
 * true: 本番環境
 * false: テスト環境
 * 【本番にリリースしたらtrueにすること！】
 * 【開発時は基本falseにすること！】
 */
const IS_PRODUCTION_FLG = false;


// 各スプレッドシートのID
const SPREAD_SHEET_IDS = IS_PRODUCTION_FLG ? {
  // 本番環境
  COMPASSNaviSystem: "1_CONc-m8Fl8R0wvOG1KLyIfPc2TPnAyi7PQypAJoFwM", // システム設定
  PJTManagement: "1EQCubV_Gp2VGjva6gsyhKBMtAFPmGN9q4PeDOM_J9sg", // 【PJT管理】公式LINEツール名簿（COMPASSナビ） 
  CLIENT_REGISTRATION: "" , // クライアント登録フォームの回答結果が転記されるシート
  SESSION_BEFORE: "1XroZA-In2ImUaMJMZCPtezbdqCIaVtFY8TrrCvwyfAo", // セッション前アウトプット回答が転記されるシート
  SESSION_AFTER: "1FwMCyuuvPj6tyiccRj6ibmHHLy92_VkFPMYrz_Vwf7U", // セッション後アウトプット回答が転記されるシート
  CLIENT_NOTE_TEMPLATE: '1aWPbDfYs3SbyCBoT4Izu6CfMmzq1KnwVp8cXW6Ar3A0', //クライアントノートテンプレ
  LOG_LIST: '1z49pH1AdfGGJvur_rGasIdtuEd52A9CXGqmY8tJV0Hg' // ログ管理シート
} : {
  // テスト環境
  COMPASSNaviSystem: "1fvOZOIFeWrvgbAoerzD1NVz4gPGW-x2UoK-E7HSwN1A", // システム設定
  PJTManagement: "1OfDu3MpBwrPE9oP5BiEjz4KHksN_Jkno_p4jSDce7ao", // 【PJT管理】公式LINEツール名簿（COMPASSナビ） 
  CLIENT_REGISTRATION: "" , // クライアント登録フォームの回答結果が転記されるシート
  SESSION_BEFORE: "1w7-QZCXNUN9aT70souERYFP-avsUsis06d9Yrusj3es", // セッション前アウトプット回答が転記されるシート
  SESSION_AFTER: "1mt0_I7Yw9xAiM3zlJiEQIpKwlDoA27sayD93RHclZx0", // セッション後アウトプット回答が転記されるシート
  CLIENT_NOTE_TEMPLATE: '1aWPbDfYs3SbyCBoT4Izu6CfMmzq1KnwVp8cXW6Ar3A0', //クライアントノートテンプレ
  LOG_LIST: '' // ログ管理シート
};

// 各LINEフォームのURL
const LINE_FROM_URLS = IS_PRODUCTION_FLG ? {
  // 本番環境
  REGIST_SUPPORT_NOTE: "https://liff.line.me/2006759470-OZ0a7wX8?unique_key=EVXhUS&ts=1741625437",
  SESSION_REGIST: "https://liff.line.me/2006759470-OZ0a7wX8?unique_key=baILqA&ts=1741625437",
  CLIENT_REGIST: "https://liff.line.me/2006759470-OZ0a7wX8?unique_key=Ve3HHH&ts=1748866886",
  COACH_REGIST: "https://liff.line.me/2006759470-OZ0a7wX8?unique_key=GOCZ7R&ts=1748866886",
  SESSION_BEFORE_OUTPUT: " https://liff.line.me/2006759470-OZ0a7wX8?unique_key=dLuUdf&ts=1740606542",
  SESSION_AFTER_OUTPUT: "https://liff.line.me/2006759470-OZ0a7wX8?unique_key=JfiALo&ts=1743692602",
  CLIENT_UPDATE: "",
  MAIL_UPDATE:""
} : {
  // テスト環境
  REGIST_SUPPORT_NOTE: "https://liff.line.me/2007474035-rBkeNA5R?unique_key=rPaqHT&ts=1748664230",
  SESSION_REGIST: "https://liff.line.me/2007474035-rBkeNA5R?unique_key=hHhvCg&ts=1748664230",
  CLIENT_REGIST: "https://liff.line.me/2007474035-rBkeNA5R?unique_key=BIoUdI&ts=1748819867",
  COACH_REGIST: "https://liff.line.me/2007474035-rBkeNA5R?unique_key=iURwVE&ts=1748819867",
  SESSION_BEFORE_OUTPUT: "https://liff.line.me/2007474035-rBkeNA5R?unique_key=hbvfAE&ts=1748664230",
  SESSION_AFTER_OUTPUT: "https://liff.line.me/2007474035-rBkeNA5R?unique_key=zruPXW&ts=1748664230",
  CLIENT_UPDATE: "https://liff.line.me/2007474035-rBkeNA5R?unique_key=3xvwTi&ts=1748866958",
  MAIL_UPDATE: "https://liff.line.me/2007474035-rBkeNA5R?unique_key=EMzYhM&ts=1750926138"
};


// 【注意！】new deployするとwebURL変わるためnewした場合は設定しなおす。
const WEB_APP_URL = IS_PRODUCTION_FLG ? 
// 本番環境　TODO 本番環境で確認
'' 
// テスト環境
:'https://script.google.com/macros/s/AKfycbyAZwbNzz7F-MKMxGaPkXFw_2lLIQfzjUGVGa63GZWJx1WWqx3MKtxcDI4exZtALtU19g/exec';


const LIFF_APP_URL = IS_PRODUCTION_FLG ?
// 本番環境
`https://liff.line.me/2006759470-npBm9Mxr`
// テスト環境
:`https://liff.line.me/2007474035-goRlynEz`;


// 問い合わせフォーム(テストと本番分けなくてもいいので分けない)
const INQUIRY_FORM = "https://docs.google.com/forms/d/1noixlQVk13mqG5lSlu6wxVGJW0upC0u2ye4YPMV7FHk/edit";


// PJT管理用スプシの各シート名
const SHEET_NAMES_MANAGEMENT = {
  COACH_LIST: "コーチ名簿",
  CLIENT_LIST: "クライアント名簿",
  AUTH_LIST: "認証管理",
  LOG_LIST: "debugログ"
}

// クライアントノートの各シート名
const SHEET_NAMES_CLIENTSUPPORT = {
  GOAL: "目標",
  DAILY_OUTPUT: "【日次】記録",
  DAILY_PILED_UP: "【日次】積み上げ日記",
  WEEKLY_PILED_UP: "【週次】積み上げﾚﾎﾟｰﾄ",
  ALL_WEEKLY_PILED_UP: "【集計】全期間積み上げﾚﾎﾟｰﾄ",
  MONTHLY_BODY_CHANGE: "【月次】身体の変化 ",
  MONTHLY_HABIT_CHECK_LIST: "【月次】習慣化ﾁｪｯｸﾘｽﾄ",
  COURSE_PROGRESS_MANAGEMENT: "ﾎﾞﾃﾞｨﾒｲｸ完全解説　進捗管理",
  SESSION_ARCHIVE: "ｾｯｼｮﾝｱｰｶｲﾌﾞ",
  PERSONAL_DATA: "ﾊﾟｰｿﾅﾙﾃﾞｰﾀ"
}

// デバッグログ管理の各シート名
const LOG_MANAGEMENT = {
  MAIL_AUTH: "メール認証",
  CLIENT_REGIST: "クライアント登録",
  COACH_REGIST: "コーチ登録",
  DAILY_OUTPUT: "日次アウトプット",
  SESSION_REGIST: "セッション登録",
  SESSION_BEFORE: "セッション前アウトプット",
  SESSION_AFTER: "セッション後アウトプット",
  GOAL_UPDATE : "ゴール更新",
  MOVIE_COURSE_OUTPUT: "動画講座アウトプット",
  ALL_PERIOD_STACKED_REPORT: "【バッチ】全期間積み上げレポート作成",
}

/**
 * エルメフォームからのhidden送信項目
 */
const ELME_HIDDEN_COL = {
  ANSWER_ID: "回答者ID",
  ANSWER_DATE: "回答日時"
}

/**
 *. 各TBLのカラム名
 */
// コーチ名簿
const COACH_LIST_TBL = {
  COACH_NO: "コーチNo.",
  NAME_KANJI: "お名前（フルネーム）",
  NAME_KANA: "フリガナ",
  LINE_NAME: "LINE登録名",
  SLACK_NAME: "Slack名",
  MAIL: "メールアドレス",
  LINE_ID: "LINE ID",
  ANSWER_ID: "回答者ID",
  DEDICATED_FOLDER_URL: "専用フォルダURL"
}

// クライアント名簿
const CLIENT_LIST_TBL = {
  CLIENT_NO: "顧客No.",
  NAME_KANJI: "名前",
  NAME_KANA: "フリガナ",
  LINE_NAME: "LINE名",
  LINE_ID: "LINE ID",
  ANSWER_ID: "回答者ID",
  SEX: "性別",
  MAIL: "メールアドレス",
  AGE: "年齢",
  BIRTH_DATE: "生年月日",
  RES_COACH_NO: "担当コーチNo.",
  RES_COACH_NAME: "担当コーチ名",
  SUPPORT_PHASE: "サポートフェーズ",
  ENQUETE_ANSWER_DATE: "アンケート回答日",
  USE_DAYS: "利用日数",
  CLIENT_NOTE_URL: "クライアントノート",
  // それ以外にもカラムはあるが基本、処理で使用していない
}

// 認証情報
const AUTH_LIST_TBL = {
  MAIL: "メールアドレス",
  TOKEN: "トークン",
  AUTH_STATUS: "認証ステータス",
  LINE_ID: "LINE ID",
  USER_TYPE: "利用者種別",
  LINE_NAME: "LINE登録名",
  ANSWER_ID: "回答者ID",
  CREATE_DATE: "認証作成日時",
  UPDATE_DATE: "認証更新日時",
  AUTH_TYPE: "認証種別",
  DELETE_FLG: "削除フラグ"
}

/**
 * TBLのカラムの固定値
 */
// サポートフェーズの固定値
const SUPPORT_PHASE = {
  DIETER_ROOKIE: "①ダイエットルーキー",
  DIETER: "②ダイエッター",
  BODY_MAKE_BEGINNER: "③ボディメイク初心者",
  BODY_MAKE_ADVANCED: "④ボディメイク上級者"
}

// 認証情報の固定値
const AUTH_STATUS = {
  COMPLETE: "認証済み",
  STILL: "未認証",
  ERR : "認証エラー"
}

// 削除フラグの固定値
const DELETE_FLG = {
  ON: "1"
}

// 利用者種別の固定値
const USER_TYPE = {
  CLIENT: "client",
  COACH: "coach",
  BOTH: "both",
  MASTER: "master" // 管理者。まだ使ってない
}

//.認証種別の固定値
const AUTH_TYPE = {
  REGIST: "reigst",
  UPDATE: "update"
}

/**
 * クライアントノートの固定値
 */
// 【日次】記録
const DAILY_OUTPUT = {
  DATE: "DATE",
  BODY_WEGHT: "体重(kg)",
  MEAL: "食事",
  KCAL: "摂取カロリー(kcal)",
  PROTEIN: "タンパク質(g)",
  FAT: "脂質(g)",
  CARB: "炭水化物(g)",
  FIBER: "食物繊維(g)",
  WATER: "水分補給(L)",
  TRAINING: "運動・トレーニング",
  POOP: "お通じ回数",
  WALK: "歩数(歩)",
  SLEEP: "睡眠",
  CONDITION: "体調"
}
// 【日次】積み上げ日記
const DAILY_PILED_UP = {
  EXECUTION_RATE: "宣言したアクション\nの実行率",
  COMMIT: "明日（今日）のアクション宣言",
  ENCOURAGEMENT: "理想の自分から今日の自分への一言　※任意",
  CONSULTATION: "コーチに相談、シェアしたいこと　※任意"
}
// 【週次】積み上げﾚﾎﾟｰﾄ
const WEEKLY_PILED_UP = {

}
// 【集計】全期間積み上げﾚﾎﾟｰﾄ
const ALL_WEEKLY_PILED_UP = {
  STACKED_REPORT_URL: "積み上げレポートURL"
}
// 【月次】身体の変化 
const MONTHLY_BODY_CHANGE = {

}
// 【月次】習慣化ﾁｪｯｸﾘｽﾄ
const MONTHLY_HABIT_CHECK_LIST = {

}
// ﾎﾞﾃﾞｨﾒｲｸ完全解説　進捗管理
const COURSE_PROGRESS_MANAGEMENT = {

}


// ボディメイクナビ最長利用日数
const DAYS_AVAILABLE_LONGEST = 184;

// コーチフォルダを格納する場所のID
const FOLDER_IDS = IS_PRODUCTION_FLG ? {
  // 本番環境
  COACH_PARENT: '1i3V1es79erlAgdg5YbtXNZbDV-VujU5N'
} : {
  // テスト環境
  COACH_PARENT: '1i5t75oN56yuorGrzsikRU-NrKudO8AEe' 
};

// チャンネルアクセストークン
const LINE_ACCESS_TOKEN = IS_PRODUCTION_FLG ? 
// 本番環境
"CAxBceZtlSiw1H4ia8tnknRqzIgsVpO82T0RBxAHSTiSugu1fP7KbeTGtC6a4hx9HbcQiZ0iT0F6iiWDuM47Iah06vm3Fr9i/yJtmHqoUY3pq8+ktBbLSxR6VT5HaIQGI4AWVgro4glTtMIWlNrt4AdB04t89/1O/w1cDnyilFU=":
// テスト環境
"ITirHv9YeRc8RYKfR0gDUvVJcjOri00/TMnxnWtt/voskkTnGWIkJwqYYQTRuAlktrjGUksZGiopfJ54vckkAyNq3CavnYXlqocQCWD87rnTcfGWM3FCbFjKcSuUvM86yHGSnk6WhvewjHEEXgRf9AdB04t89/1O/w1cDnyilFU=";

const BODY_MAKE_NAVI_LINE_ID = IS_PRODUCTION_FLG ?
// 本番環境
"@249ctzxl":
// テスト環境
"@673gaxvr"
