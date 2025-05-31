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
const getPropertyValue = (key: string, defaultValue: string): string => {
  const properties = PropertiesService.getScriptProperties();
  return properties.getProperty(key) || defaultValue;
};

/**
 * スクリプトプロパティを設定する
 */
const setPropertyValue = (key: string, value: string): void => {
  const properties = PropertiesService.getScriptProperties();
  properties.setProperty(key, value);
};

/**
 * 実行タイミングを判定する
 * 基本は月末の1日前、ただし土日の場合は木曜に前倒し
 * @returns 実行すべき日であればtrue
 */
const shouldRunToday = (): boolean => {
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
};

// 定数定義
const CONFIG = {
  // 作業記録スプレッドシート
  WORK_LOG: {
    FILE_ID: getPropertyValue(
      'WORK_LOG_FILE_ID',
      '1vCw-kt1eGBlRmLqAu8urEoX-oNYJbbbgqGDfLW7h9JI'
    ),
    SHEET_NAME: '作業記録',
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
    OUTPUT_FOLDER_ID: getPropertyValue(
      'INVOICE_OUTPUT_FOLDER_ID',
      '1FKxvmPu4icl4XOWZC9JM1Xc1mpMXnn5A'
    ),
    CELLS: {
      INVOICE_DATE: 'E4', // 請求日
      INVOICE_NUMBER: 'E5', // 請求書番号
      WORK_HOURS: 'C33', // 作業時間
      TOTAL_AMOUNT: 'F30', // 請求金額
    },
    OUTPUT_SHEET_NAME: 'ダウンロード用',
  },
  // 支払先情報
  PAYEE: {
    NAME: getPropertyValue('PAYEE_NAME', '東顕正'),
  },
};

/**
 * 請求書を作成する関数
 */
const createInvoice = () => {
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
    deleteOriginalFile(invoiceFile);

    Logger.log(`請求書PDFを作成しました: ${savedFile.getName()}`);
    return savedFile.getUrl();
  } catch (error) {
    Logger.log(`エラーが発生しました: ${error.message}`);
    throw error;
  }
};

/**
 * 当月の作業データを取得する
 */
const getCurrentMonthWorkData = (): WorkLogEntry[] => {
  const ss = SpreadsheetApp.openById(CONFIG.WORK_LOG.FILE_ID);
  const sheet = ss.getSheetByName(CONFIG.WORK_LOG.SHEET_NAME);

  if (!sheet) {
    throw new Error('作業記録シートが見つかりません');
  }

  // シートからデータを取得
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();

  // ヘッダー行をスキップ
  const data = values.slice(1);

  // 当月のデータを抽出
  const now = new Date();
  const currentMonth = now.getMonth();
  const currentYear = now.getFullYear();

  return data
    .map((row): WorkLogEntry => {
      return {
        id: row[CONFIG.WORK_LOG.COLUMNS.ID],
        date: new Date(row[CONFIG.WORK_LOG.COLUMNS.DATE]),
        hours: Number(row[CONFIG.WORK_LOG.COLUMNS.HOURS]),
        description: row[CONFIG.WORK_LOG.COLUMNS.DESC],
        createdAt: new Date(row[CONFIG.WORK_LOG.COLUMNS.CREATED]),
        updatedAt: new Date(row[CONFIG.WORK_LOG.COLUMNS.UPDATED]),
      };
    })
    .filter((entry) => {
      const entryMonth = entry.date.getMonth();
      const entryYear = entry.date.getFullYear();
      return entryMonth === currentMonth && entryYear === currentYear;
    });
};

/**
 * 合計作業時間を計算する
 */
const calculateTotalHours = (entries: WorkLogEntry[]): number => {
  return entries.reduce((total, entry) => total + entry.hours, 0);
};

/**
 * 請求書テンプレートを複製する
 */
const duplicateInvoiceTemplate = (): GoogleAppsScript.Drive.File => {
  const templateFile = DriveApp.getFileById(CONFIG.INVOICE.TEMPLATE_ID);

  // 日付文字列を生成（yyyymmdd形式）
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const lastDay = new Date(year, now.getMonth() + 1, 0).getDate();
  const dateStr = `${year}${month}${lastDay}`;

  // 新規ファイル名
  const newFileName = `請求書_${dateStr}`;

  // テンプレートを複製
  const folder = DriveApp.getFolderById(CONFIG.INVOICE.OUTPUT_FOLDER_ID);
  const newFile = templateFile.makeCopy(newFileName, folder);

  return newFile;
};

/**
 * 請求書に値を入力する
 */
const setInvoiceValues = (
  file: GoogleAppsScript.Drive.File,
  totalHours: number
): GoogleAppsScript.Spreadsheet.Sheet => {
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

  // 請求書番号（yyyymmdd）
  const invoiceNumber = Utilities.formatDate(lastDay, 'JST', 'yyyyMMdd');

  // 値を設定
  sheet.getRange(CONFIG.INVOICE.CELLS.INVOICE_DATE).setValue(dateStr);
  sheet.getRange(CONFIG.INVOICE.CELLS.INVOICE_NUMBER).setValue(invoiceNumber);
  sheet.getRange(CONFIG.INVOICE.CELLS.WORK_HOURS).setValue(totalHours);

  return sheet;
};

/**
 * 請求書の金額を取得する
 */
const getInvoiceAmount = (
  sheet: GoogleAppsScript.Spreadsheet.Sheet
): number => {
  const amountCell = sheet.getRange(CONFIG.INVOICE.CELLS.TOTAL_AMOUNT);
  return amountCell.getValue();
};

/**
 * PDFファイルを生成する
 */
const generatePDF = (
  file: GoogleAppsScript.Drive.File
): GoogleAppsScript.Base.Blob => {
  const ss = SpreadsheetApp.openById(file.getId());
  const sheet = ss.getSheetByName(CONFIG.INVOICE.OUTPUT_SHEET_NAME);

  if (!sheet) {
    throw new Error('ダウンロード用シートが見つかりません');
  }

  // PDFエクスポート用のURL
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
};

/**
 * 生成したPDFを保存する
 */
const saveInvoicePDF = (
  pdfBlob: GoogleAppsScript.Base.Blob,
  amount: number
): GoogleAppsScript.Drive.File => {
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
};

/**
 * 元のファイルを削除する
 */
const deleteOriginalFile = (file: GoogleAppsScript.Drive.File): void => {
  file.setTrashed(true);
};

/**
 * スクリプトプロパティを初期化する
 * 初回セットアップ時や設定変更時に使用
 */
const initializeProperties = () => {
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
    'INVOICE_OUTPUT_FOLDER_ID',
    '1FKxvmPu4icl4XOWZC9JM1Xc1mpMXnn5A'
  );
  setPropertyValue('PAYEE_NAME', '東顕正');

  return '設定を初期化しました';
};

/**
 * トリガーから毎日実行される関数
 * 実行タイミングを判定し、適切な日に請求書を生成する
 */
const dailyTrigger = () => {
  if (shouldRunToday()) {
    try {
      const result = createInvoice();
      Logger.log(`請求書を自動生成しました: ${result}`);
      return result;
    } catch (error) {
      Logger.log(`自動生成中にエラーが発生しました: ${error.message}`);
      return `エラー: ${error.message}`;
    }
  } else {
    Logger.log('今日は請求書生成日ではありません');
    return '処理をスキップしました（実行日ではありません）';
  }
};

/**
 * トリガーを設定する関数
 * スクリプトエディタから手動で一度だけ実行する
 */
const setDailyTrigger = () => {
  // 既存のトリガーをすべて削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach((trigger) => {
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
};

// Google Apps Scriptのグローバルに公開
// @ts-ignore
global.createInvoice = createInvoice;
// @ts-ignore
global.initializeProperties = initializeProperties;
// @ts-ignore
global.dailyTrigger = dailyTrigger;
// @ts-ignore
global.setDailyTrigger = setDailyTrigger;
