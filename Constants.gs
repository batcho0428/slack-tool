/* --------------------------------------------------------------------------
 * 設定 & 定数定義
 * -------------------------------------------------------------------------- */
const APP_NAME = '45th NUTFES 実行委員マスタ';
const APP_HEADER_COLOR = '#1a237e'; // 紺色

// SPREADSHEET_IDは関数実行時に取得（定数化によるタイミング問題を回避）
const SHEET_USERS = 'Users';
const SHEET_LOGS = 'Logs';
const SHEET_GRADE = 'grade';
const SHEET_ORG = 'org';
const SHEET_DEPT = 'dept';
const SHEET_ROLE = 'role';
const SHEET_FIELD = 'field';
const SHEET_FORMS = 'Forms';
const SHEET_TOKENS = 'Tokens'; // 新規テーブル

const SESSION_DURATION_DAYS = 3; // セッション有効期限(日)

// Usersシートの列定義 (0-based index)
const COL = {
  NAME_JP: 0,
  NAME_EN: 1,
  STUDENT_ID: 2,
  GRADE: 3,
  FIELD: 4,
  EMAIL: 5,
  PHONE: 6,
  BIRTHDAY: 7,
  ALMA_MATER: 8,
  RETIRED: 9,
  CONTINUE_NEXT: 10,
  ADMIN: 11,
  CAR_OWNER: 12,
  AFFILIATION_START: 13,
  ORG_START: 13
};

const AFFILIATION_SLOTS = 10;

// Tokensシートの列定義 (新規)
const COL_TOKENS = {
  SESSION_ID: 0,
  EMAIL: 1,
  SLACK_TOKEN: 2, // User Token (xoxp-...)
  CREATED_AT: 3
};

const HEADER_USERS = [
  '氏名', 'Name', '学籍番号', '学年', '分野', 'メールアドレス', '電話番号', '生年月日', '出身校',
  '退局', '次年度継続', 'Admin', '車所有',
  '所属部門1', '役職1',
  '所属部門2', '役職2',
  '所属部門3', '役職3',
  '所属部門4', '役職4',
  '所属部門5', '役職5',
  '所属部門6', '役職6',
  '所属部門7', '役職7',
  '所属部門8', '役職8',
  '所属部門9', '役職9',
  '所属部門10', '役職10'
];

const HEADER_TOKENS = ['Session ID', 'Email', 'Slack Token', 'Created At'];
const HEADER_GRADE = ['pid', 'grade', 'count'];
const HEADER_ORG = ['pid', 'org', 'gen', 'status', 'not_main_org'];
const HEADER_DEPT = ['pid', 'dept', 'org', 'status', 'not_main_dept'];
const HEADER_ROLE = ['pid', 'role', 'gen', 'not_main_role'];
const HEADER_FIELD = ['pid', 'field'];
const HEADER_FORMS = ['アンケートシート', 'フォームURL', 'フォームタイトル', '担当部門', '収集中', 'スコアの名称', 'スコアの単位', '締め切り日'];
const HEADER_LOGS = ['Time', 'Sender', 'Recipient', 'Status', 'Details'];





/**
 * Debug helper: probe Forms sheet rows and report whether target spreadsheet and form can be opened.
 * Returns array of { rowIndex, rawA, rawB, spreadsheetId, ssOpen:bool, ssError, formOpen:bool, formError }
 */

// Stores in Script Properties (JSON maps)
// sessions: sessionId -> { email, created }
// tokensByEmail: email -> { slackToken, created }
const SESSIONS_PROP_KEY = 'SESSIONS_STORE';
const TOKENS_BY_EMAIL_PROP_KEY = 'TOKENS_BY_EMAIL_STORE';
const TOKENS_PROP_MAX = 200000; // safe threshold (characters). 古いものから削除して収める












/* --------------------------------------------------------------------------
 * 0. 初期セットアップ (マイグレーション & トリガー設定)
 * -------------------------------------------------------------------------- */







/* --------------------------------------------------------------------------
 * Webアプリ & OAuthエンドポイント
 * -------------------------------------------------------------------------- */

/* --------------------------------------------------------------------------
 * 1. ログイン & セッション管理 (Tokensシート利用)
 * -------------------------------------------------------------------------- */









// 1-A. OTPリクエスト (BotからDM送信)

// 1-B. OTP検証 & セッション発行

/* --------------------------------------------------------------------------
 * 2. OAuth関連 (PCからのログイン用 - Tokensシートに対応)
 * -------------------------------------------------------------------------- */




/* --------------------------------------------------------------------------
 * 3. 共通機能 (Token取得ロジック - Tokensシートから取得)
 * -------------------------------------------------------------------------- */


/* --------------------------------------------------------------------------
 * 4. マイページ機能 (プロフィール取得・更新)
 * -------------------------------------------------------------------------- */


/* --------------------------------------------------------------------------
 * 5. DM送信 & チャンネル招待
 * -------------------------------------------------------------------------- */
/**
 * 管理者向け: 新規ユーザー作成
 * @param {string} sessionToken
 * @param {object} userObj
 */




/* --------------------------------------------------------------------------
 * 6. 検索 & ユーティリティ
 * -------------------------------------------------------------------------- */



/* --------------------------------------------------------------------------
 * 管理者列初期化バッチ
 * Users シートのヘッダに 'Admin' を追加し、Z列の空セルを FALSE に設定します
 * -------------------------------------------------------------------------- */

/* --------------------------------------------------------------------------
 * 7. 名簿出力機能
 * - 共有フォルダIDはスクリプトプロパティ 'EXPORT_SHARED_FOLDER_ID' に保存する想定
 * - フロントからはフォルダ一覧取得と、選択した項目でスプレッドシートを生成するAPIを提供
 * -------------------------------------------------------------------------- */




/**
 * CSV 出力版: Drive に保存せず CSV 文字列を返す
 * フロントでダウンロード処理を行う
 */

/* --------------------------------------------------------------------------
 * 8. アンケート（Google Form 回答）表示機能
 * - スプレッドシート内の各シートを走査し、以下のいずれかでアンケートシートと判定する
 *   1) A1 に URL が含まれる
 *   2) ヘッダにメールアドレス(メール|Email) とタイムスタンプ(Timestamp|回答日等)が含まれる
 * - listSurveys(): アンケート一覧（タイトル、最新スコア、最新回答日）を返す
 * - getSurveyDetails(sheetName): 指定シートの全回答・最新回答（メールで一意化）・スコア統計を返す
 * -------------------------------------------------------------------------- */




// Try to detect which row contains the header (search first few rows for common header patterns)












// Try to extract Form ID from a URL like https://docs.google.com/forms/d/FORM_ID/

/* -----------------------------
 * Collections (集金) 機能
 * -----------------------------*/
const SHEET_COLLECTIONS = 'Collections';
const SHEET_COLLECTIONS_LOG = 'Collections_log';

const HEADER_COLLECTIONS = ['id','タイトル','スプレッドシート','担当部門','作成日時','作成者'];
const HEADER_COLLECTIONS_LOG = ['Collections_id','タイムスタンプ','取引先メールアドレス','取引種別','取引金額','担当者メールアドレス'];





// Use the primary helpers defined earlier in the file.
// Provide a utility that returns all matching header indices (plural) without overriding the single-index helper.









