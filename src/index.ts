/**
 * 請求書自動作成スクリプト
 *
 * 作業記録スプレッドシートから当月の作業時間を自動集計し、
 * 請求書テンプレートに値を差し込んでPDF出力・保存します。
 */

interface WorkLogEntry {
  id: string;
  date: Date;
  hours: number;
  description: string;
  createdAt: Date;
  updatedAt: Date;
}

/**
 * スクリプトプロパティからID設定を読み込む
 */
function getPropertyValue(key: string, defaultValue: string): string {
  const properties = PropertiesService.getScriptProperties();
  return properties.getProperty(key) || defaultValue;
}

/**
 * スクリプトプロパティを設定する
 */
function setPropertyValue(key: string, value: string): void {
  const properties = PropertiesService.getScriptProperties();
  properties.setProperty(key, value);
}

/**
 * 実行タイミングを判定する
 * 基本は月末の1日前、ただし土日の場合は木曜に前倒し
 * @returns 実行すべき日であればtrue
 */
function shouldRunToday(): boolean {
  const today = new Date();
  const month = today.getMonth();
  const year = today.getFullYear();

  // 月末日を取得
  const lastDayOfMonth = new Date(year, month + 1, 0);
  // 月末の1日前
  const dayBeforeLastDay = new Date(year, month, lastDayOfMonth.getDate() - 1);

  // 曜日を取得（0:日曜、1:月曜、..., 6:土曜）
  const todayDay = today.getDay();
  const lastDayDay = lastDayOfMonth.getDay();
  const dayBeforeLastDayDay = dayBeforeLastDay.getDay();

  // 今日の日付
  const todayDate = today.getDate();

  // ケース1: 月末が土日で、今日が木曜の場合
  if ((lastDayDay === 0 || lastDayDay === 6) && todayDay === 4) {
    // 月末の前の木曜日を計算
    const thursdayBeforeLastDay = new Date(lastDayOfMonth);
    thursdayBeforeLastDay.setDate(
      lastDayOfMonth.getDate() - ((lastDayDay + 3) % 7)
    );

    // 今日がその木曜と同じなら実行
    return todayDate === thursdayBeforeLastDay.getDate();
  }

  // ケース2: 月末前日が土日で、今日が木曜の場合
  if (
    (dayBeforeLastDayDay === 0 || dayBeforeLastDayDay === 6) &&
    todayDay === 4
  ) {
    // 月末前日の前の木曜日を計算
    const thursdayBeforeDayBeforeLastDay = new Date(dayBeforeLastDay);
    thursdayBeforeDayBeforeLastDay.setDate(
      dayBeforeLastDay.getDate() - ((dayBeforeLastDayDay + 3) % 7)
    );

    // 今日がその木曜と同じなら実行
    return todayDate === thursdayBeforeDayBeforeLastDay.getDate();
  }

  // ケース3: 通常ケース - 月末の1日前
  return todayDate === dayBeforeLastDay.getDate();
}

// 定数定義
const CONFIG = {
  // 作業記録スプレッドシート
  WORK_LOG: {
    FILE_ID: getPropertyValue(
      'WORK_LOG_FILE_ID',
      '1vCw-kt1eGBlRmLqAu8urEoX-oNYJbbbgqGDfLW7h9JI'
    ),
    SHEET_NAME: '作業',
    COLUMNS: {
      ID: 0, // A列
      DATE: 1, // B列
      HOURS: 2, // C列
      DESC: 3, // D列
      CREATED: 4, // E列
      UPDATED: 5, // F列
    },
  },
  // 請求書テンプレート
  INVOICE: {
    TEMPLATE_ID: getPropertyValue(
      'INVOICE_TEMPLATE_ID',
      '1gv2H0YM5qxhWm_bJTPjmFBK3YHXfFZrdRHXrG6bmSgg'
    ),
    WORK_FOLDER_ID: getPropertyValue(
      'INVOICE_WORK_FOLDER_ID',
      '1Bc6i4RccvPdU0U4Q9bZS7ogMpA0iu_UE'
    ),
    OUTPUT_FOLDER_ID: getPropertyValue(
      'INVOICE_OUTPUT_FOLDER_ID',
      '1FKxvmPu4icl4XOWZC9JM1Xc1mpMXnn5A'
    ),
    CELLS: {
      INVOICE_DATE: 'E4', // 請求日
      INVOICE_NUMBER: 'E5', // 請求書番号
      WORK_HOURS: 'C33', // 作業時間
      TOTAL_AMOUNT: 'E30', // 請求金額
    },
    OUTPUT_SHEET_NAME: 'ダウンロード用',
  },
  // 支払先情報
  PAYEE: {
    NAME: getPropertyValue('PAYEE_NAME', '東顕正'),
  },
  // メール通知設定
  NOTIFICATION: {
    EMAIL: getPropertyValue('NOTIFICATION_EMAIL', 'codino.baggio10@gmail.com'),
    SUBJECT: '請求書発行',
    SUCCESS_MESSAGE: '請求書の発行を行いました。',
    ERROR_MESSAGE:
      '請求書の発行を実行しましたがエラーになりました。エラー内容：',
  },
  // アーカイブフォルダ
  ARCHIVE_FOLDER_ID: getPropertyValue(
    'ARCHIVE_FOLDER_ID',
    '1mot5nlo4dynp3iK9PfCGx6EStOHdvHsv'
  ),
};

/**
 * 請求書を作成する関数
 * @param date 日付（オプション）
 * @returns 生成されたURLまたはエラーメッセージ
 */
function createInvoice(date?: string): string {
  try {
    // 当月の作業データを取得して集計
    const currentMonthData = getCurrentMonthWorkData();
    const totalHours = calculateTotalHours(currentMonthData);

    // 請求書テンプレートを複製
    const invoiceFile = duplicateInvoiceTemplate();

    // 請求書に値を入力
    const invoiceSheet = setInvoiceValues(invoiceFile, totalHours);

    // PDF生成
    const invoiceAmount = getInvoiceAmount(invoiceSheet);
    const pdfBlob = generatePDF(invoiceFile);

    // 結果を保存して原本削除
    const savedFile = saveInvoicePDF(pdfBlob, invoiceAmount);
    moveOriginalFile(invoiceFile);
    Logger.log(`請求書PDFを作成しました: ${savedFile.getName()}`);

    // 成功通知メール送信
    sendNotificationEmail(true);

    return savedFile.getUrl();
  } catch (error) {
    Logger.log(`エラーが発生しました: ${error.message}`);
    return `Error: ${error.toString()}`;
  }
}

/**
 * 当月または指定した月の作業データを取得する
 * @param targetYear 対象年（未指定時は当年）
 * @param targetMonth 対象月（0～11、未指定時は当月）
 * @returns 対象月の作業エントリ配列
 */
function getCurrentMonthWorkData(
  targetYear?: number,
  targetMonth?: number
): WorkLogEntry[] {
  try {
    // スプレッドシートIDをログに記録（デバッグ用）
    Logger.log(`使用中のスプレッドシートID: ${CONFIG.WORK_LOG.FILE_ID}`);

    // スプレッドシートを開く
    let ss;
    try {
      ss = SpreadsheetApp.openById(CONFIG.WORK_LOG.FILE_ID);
      Logger.log(`スプレッドシート「${ss.getName()}」を開きました`);
    } catch (e) {
      throw new Error(
        `スプレッドシートを開けませんでした: ${e.message}。正しいIDか確認してください: ${CONFIG.WORK_LOG.FILE_ID}`
      );
    }

    // シート名をログに記録（デバッグ用）
    Logger.log(`検索中のシート名: ${CONFIG.WORK_LOG.SHEET_NAME}`);

    // シートの一覧を取得してログに記録
    const sheets = ss.getSheets();
    Logger.log(
      `利用可能なシート: ${sheets
        .map(function (s) {
          return s.getName();
        })
        .join(', ')}`
    );

    // 指定されたシートを取得
    const sheet = ss.getSheetByName(CONFIG.WORK_LOG.SHEET_NAME);
    if (!sheet) {
      throw new Error(
        `作業記録シートが見つかりません。シート名「${
          CONFIG.WORK_LOG.SHEET_NAME
        }」を確認してください。利用可能なシート: ${sheets
          .map(function (s) {
            return s.getName();
          })
          .join(', ')}`
      );
    }
    Logger.log(`シート「${sheet.getName()}」が見つかりました`);

    // シートからデータを取得
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    Logger.log(`データ行数: ${values.length} 行取得しました`);

    // ヘッダー行をスキップ
    const data = values.slice(1);

    // 対象年月を決定
    const now = new Date();
    const currentMonth =
      targetMonth !== undefined ? targetMonth : now.getMonth();
    const currentYear =
      targetYear !== undefined ? targetYear : now.getFullYear();
    Logger.log(
      `対象年月: ${currentYear}年${currentMonth + 1}月のデータを抽出します`
    );
    const result = data
      .map(function (row): WorkLogEntry {
        return {
          id: row[CONFIG.WORK_LOG.COLUMNS.ID],
          date: new Date(row[CONFIG.WORK_LOG.COLUMNS.DATE]),
          hours: Number(row[CONFIG.WORK_LOG.COLUMNS.HOURS]),
          description: row[CONFIG.WORK_LOG.COLUMNS.DESC],
          createdAt: new Date(row[CONFIG.WORK_LOG.COLUMNS.CREATED]),
          updatedAt: new Date(row[CONFIG.WORK_LOG.COLUMNS.UPDATED]),
        };
      })
      .filter(function (entry) {
        // 日付が無効な場合はスキップ
        if (isNaN(entry.date.getTime())) return false;

        const entryMonth = entry.date.getMonth();
        const entryYear = entry.date.getFullYear();
        return entryMonth === currentMonth && entryYear === currentYear;
      });

    Logger.log(
      `${result.length}件のデータが${currentYear}年${
        currentMonth + 1
      }月分として抽出されました`
    );
    return result;
  } catch (error) {
    // エラーメッセージをより詳細に
    Logger.log(`作業データ取得エラー: ${error.message}`);
    throw new Error(`作業データの取得に失敗しました: ${error.message}`);
  }
}

/**
 * 合計作業時間を計算する
 */
function calculateTotalHours(entries: WorkLogEntry[]): number {
  return entries.reduce(function (total, entry) {
    return total + entry.hours;
  }, 0);
}

/**
 * 請求書テンプレートを複製する
 */
function duplicateInvoiceTemplate(): GoogleAppsScript.Drive.File {
  const templateFile = DriveApp.getFileById(CONFIG.INVOICE.TEMPLATE_ID);

  // 日付文字列を生成（yyyymmdd形式）
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const lastDay = new Date(year, now.getMonth() + 1, 0).getDate();
  const dateStr = `${year}${month}${lastDay}`;
  // 新規ファイル名
  const newFileName = `請求書_${dateStr}`;

  // テンプレートを作業フォルダに複製
  const folder = DriveApp.getFolderById(CONFIG.INVOICE.WORK_FOLDER_ID);
  const newFile = templateFile.makeCopy(newFileName, folder);

  return newFile;
}

/**
 * 請求書に値を入力する
 */
function setInvoiceValues(
  file: GoogleAppsScript.Drive.File,
  totalHours: number
): GoogleAppsScript.Spreadsheet.Sheet {
  const ss = SpreadsheetApp.openById(file.getId());
  const sheet = ss.getSheetByName('シート1');

  if (!sheet) {
    throw new Error('シート1が見つかりません');
  }
  // 月末の日付を取得
  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth();
  const lastDay = new Date(year, month + 1, 0);

  // 日付の文字列（yyyy/mm/dd）
  const dateStr = Utilities.formatDate(lastDay, 'JST', 'yyyy/MM/dd');

  // 請求書番号（yyyymmdd）- 文字列連結で確実にフォーマットを保証
  const yyyy = lastDay.getFullYear().toString();
  const mm = String(lastDay.getMonth() + 1).padStart(2, '0');
  const dd = String(lastDay.getDate()).padStart(2, '0');
  const invoiceNumber = yyyy + mm + dd;
  // 値を設定
  sheet.getRange(CONFIG.INVOICE.CELLS.INVOICE_DATE).setValue(dateStr);
  sheet.getRange(CONFIG.INVOICE.CELLS.INVOICE_NUMBER).setValue(invoiceNumber);
  sheet
    .getRange(CONFIG.INVOICE.CELLS.WORK_HOURS)
    .setValue(`作業時間：${totalHours}h`); // 時間単位を付加

  return sheet;
}

/**
 * 請求書の金額を取得する
 */
function getInvoiceAmount(sheet: GoogleAppsScript.Spreadsheet.Sheet): number {
  const amountCell = sheet.getRange(CONFIG.INVOICE.CELLS.TOTAL_AMOUNT);
  return amountCell.getValue();
}

/**
 * PDFファイルを生成する
 */
function generatePDF(
  file: GoogleAppsScript.Drive.File
): GoogleAppsScript.Base.Blob {
  const ss = SpreadsheetApp.openById(file.getId());
  const sheet = ss.getSheetByName(CONFIG.INVOICE.OUTPUT_SHEET_NAME);

  if (!sheet) {
    throw new Error('ダウンロード用シートが見つかりません');
  }

  try {
    // OAuth認証でPDF取得（外部リクエスト権限が必要）
    const url =
      `https://docs.google.com/spreadsheets/d/${file.getId()}/export?` +
      'exportFormat=pdf&format=pdf' +
      '&size=A4' +
      '&portrait=true' +
      '&fitw=true' +
      '&sheetnames=false' +
      '&printtitle=false' +
      '&pagenumbers=false' +
      '&gridlines=false' +
      '&fzr=false' +
      `&gid=${sheet.getSheetId()}`;

    // OAuth認証でPDF取得
    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(url, {
      headers: {
        Authorization: `Bearer ${token}`,
      },
    });
    return response.getBlob();
  } catch (error) {
    Logger.log(`外部リクエストでのPDF取得に失敗しました: ${error.message}`);
    Logger.log('代替方法でPDFを生成します');

    // 代替方法: 直接スプレッドシートからPDFを取得
    // この方法は外部リクエスト権限を必要としない
    const blob = DriveApp.getFileById(file.getId()).getBlob();
    return blob;
  }
}

/**
 * 生成したPDFを保存する
 */
function saveInvoicePDF(
  pdfBlob: GoogleAppsScript.Base.Blob,
  amount: number
): GoogleAppsScript.Drive.File {
  // 月末の日付を取得
  const now = new Date();
  const lastDay = new Date(now.getFullYear(), now.getMonth() + 1, 0);

  // ファイル名用日付（yyyymmdd）
  const dateStr = Utilities.formatDate(lastDay, 'JST', 'yyyyMMdd');

  // PDF名を設定: yyyymmdd_金額_東顕正.pdf
  const pdfName = `${dateStr}_${amount}_${CONFIG.PAYEE.NAME}.pdf`;

  // 名前を設定して保存
  const pdfBlob2 = pdfBlob.setName(pdfName);
  const folder = DriveApp.getFolderById(CONFIG.INVOICE.OUTPUT_FOLDER_ID);
  return folder.createFile(pdfBlob2);
}

/**
 * 元のファイルを保存用フォルダに移動する
 */
function moveOriginalFile(file: GoogleAppsScript.Drive.File): void {
  const archiveFolder = DriveApp.getFolderById(CONFIG.ARCHIVE_FOLDER_ID);

  // 現在のファイルの親フォルダを取得
  const parentFolders = file.getParents();

  if (parentFolders.hasNext()) {
    const parentFolder = parentFolders.next();
    // 元のフォルダから削除して新しいフォルダに追加
    parentFolder.removeFile(file);
    archiveFolder.addFile(file);
    Logger.log(
      `ファイル「${file.getName()}」をアーカイブフォルダに移動しました`
    );
  }
}

/* initializePropertiesの重複実装を削除 */

/**
 * 通知メールを送信する
 * @param isSuccess 成功したかどうか
 * @param errorMessage エラーメッセージ（失敗時のみ）
 */
function sendNotificationEmail(
  isSuccess: boolean,
  errorMessage: string = ''
): void {
  const recipient = CONFIG.NOTIFICATION.EMAIL;
  const subject = CONFIG.NOTIFICATION.SUBJECT;

  // 本文を設定
  let body = '';
  if (isSuccess) {
    body = CONFIG.NOTIFICATION.SUCCESS_MESSAGE;
  } else {
    body = `${CONFIG.NOTIFICATION.ERROR_MESSAGE}${errorMessage}`;
  }

  // メール送信
  try {
    GmailApp.sendEmail(recipient, subject, body);
    Logger.log(`通知メールを送信しました: ${recipient}`);
  } catch (error) {
    Logger.log(`メール送信中にエラーが発生しました: ${error.message}`);
  }
}

/**
 * トリガーから毎日実行される関数
 * 実行タイミングを判定し、適切な日に請求書を生成する
 * @returns 処理結果のメッセージ
 */
function dailyTrigger(): string {
  if (shouldRunToday()) {
    try {
      const result = createInvoice();
      Logger.log(`請求書を自動生成しました: ${result}`);
      sendNotificationEmail(true); // 成功通知
      return result;
    } catch (error) {
      Logger.log(`自動生成中にエラーが発生しました: ${error.message}`);
      sendNotificationEmail(false, error.message); // エラー通知
      return `エラー: ${error.message}`;
    }
  } else {
    Logger.log('今日は請求書生成日ではありません');
    return '処理をスキップしました（実行日ではありません）';
  }
}

/**
 * トリガーを設定する関数
 * スクリプトエディタから手動で一度だけ実行する
 * @returns 設定結果のメッセージ
 */
function setDailyTrigger(): string {
  try {
    // 既存のトリガーをすべて削除
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(function (trigger) {
      if (trigger.getHandlerFunction() === 'dailyTrigger') {
        ScriptApp.deleteTrigger(trigger);
      }
    });

    // 毎日午前9時に実行するトリガーを設定
    ScriptApp.newTrigger('dailyTrigger')
      .timeBased()
      .everyDays(1)
      .atHour(9)
      .create();
    return 'トリガーを設定しました（毎日午前9時に実行）';
  } catch (e) {
    return `トリガー設定中にエラーが発生しました: ${e.toString()}`;
  }
}

// Google Apps Script環境では関数はグローバルにエクスポートされる
// TypeScriptを使用している場合でも、ルートレベルで宣言された関数はGASの実行時に
// 自動的にグローバルスコープになるため、特別なエクスポート処理は不要
// トラブルシューティング関数もグローバルにエクスポートされる

/**
 * スクリプトプロパティを初期化する関数
 * 初回セットアップ時や設定変更時に使用
 * @returns 設定結果のメッセージ
 */
function initializeProperties(): string {
  try {
    // デフォルト値を設定
    setPropertyValue(
      'WORK_LOG_FILE_ID',
      '1vCw-kt1eGBlRmLqAu8urEoX-oNYJbbbgqGDfLW7h9JI'
    );
    setPropertyValue(
      'INVOICE_TEMPLATE_ID',
      '1gv2H0YM5qxhWm_bJTPjmFBK3YHXfFZrdRHXrG6bmSgg'
    );
    setPropertyValue(
      'INVOICE_WORK_FOLDER_ID',
      '1Bc6i4RccvPdU0U4Q9bZS7ogMpA0iu_UE'
    );
    setPropertyValue(
      'INVOICE_OUTPUT_FOLDER_ID',
      '1FKxvmPu4icl4XOWZC9JM1Xc1mpMXnn5A'
    );
    setPropertyValue('PAYEE_NAME', '東顕正');
    // 通知設定
    setPropertyValue('NOTIFICATION_EMAIL', 'codino.baggio10@gmail.com');

    return '設定を初期化しました';
  } catch (e) {
    return `設定の初期化に失敗しました: ${e.toString()}`;
  }
}

// テスト用のグローバル関数定義
// @ts-ignore
function doGet() {
  return HtmlService.createHtmlOutput('Invoice Creator API is working!');
}

// global-exports.tsからのグローバル関数は自動的にロードされる
// Google Apps Script環境（本番環境）では関数は自動的にグローバルになるため何もしない
