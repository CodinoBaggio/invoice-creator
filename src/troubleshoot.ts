/**
 * 作業記録スプレッドシートの存在確認とトラブルシューティング用スクリプト
 */

// デフォルトのID（CONFIGと同じ）
const DEFAULT_WORK_LOG_FILE_ID = '1vCw-kt1eGBlRmLqAu8urEoX-oNYJbbbgqGDfLW7h9JI';
const DEFAULT_SHEET_NAME = '作業記録';

/**
 * 作業記録スプレッドシートが存在するか確認する
 */
function checkWorkLogSpreadsheet() {
  // 設定からIDを取得
  const storedId =
    PropertiesService.getScriptProperties().getProperty('WORK_LOG_FILE_ID');
  const fileId = storedId || DEFAULT_WORK_LOG_FILE_ID;

  try {
    // スプレッドシートを開く
    const ss = SpreadsheetApp.openById(fileId);
    Logger.log('スプレッドシートを正常に開きました: ' + ss.getName());

    // シート一覧を取得
    const sheets = ss.getSheets();
    Logger.log('シート一覧:');
    sheets.forEach((sheet) => {
      Logger.log('- ' + sheet.getName());
    });

    // 作業記録シートを取得
    const workLogSheet = ss.getSheetByName(DEFAULT_SHEET_NAME);
    if (workLogSheet) {
      Logger.log(
        '作業記録シートが見つかりました。サイズ: ' +
          workLogSheet.getLastRow() +
          '行 x ' +
          workLogSheet.getLastColumn() +
          '列'
      );
      return {
        status: 'SUCCESS',
        message: '作業記録シートにアクセスできました。',
        sheetName: workLogSheet.getName(),
        rowCount: workLogSheet.getLastRow(),
      };
    } else {
      Logger.log('作業記録シートが見つかりません。利用可能なシート:');
      sheets.forEach((sheet, index) => {
        Logger.log(`${index + 1}. ${sheet.getName()}`);
      });

      return {
        status: 'ERROR',
        message:
          '作業記録シートが見つかりません。利用可能なシート: ' +
          sheets.map((s) => s.getName()).join(', '),
        available_sheets: sheets.map((s) => s.getName()),
      };
    }
  } catch (error) {
    Logger.log('エラーが発生しました: ' + error.message);
    return {
      status: 'ERROR',
      message: 'エラー: ' + error.message,
    };
  }
}

/**
 * スクリプトプロパティを確認する
 */
function checkScriptProperties() {
  const props = PropertiesService.getScriptProperties().getProperties();
  Logger.log('現在のスクリプトプロパティ:');

  const result = {};
  for (const key in props) {
    Logger.log(`${key}: ${props[key]}`);
    result[key] = props[key];
  }

  return result;
}

/**
 * ファイルに対するアクセス権を確認する
 */
function checkFileAccess(fileId) {
  fileId = fileId || DEFAULT_WORK_LOG_FILE_ID;

  try {
    // ドライブのファイルを確認
    const file = DriveApp.getFileById(fileId);
    const accessLevel = file.getSharingAccess();
    const permissions = file.getSharingPermission();

    Logger.log(`ファイル「${file.getName()}」へのアクセス:`);
    Logger.log(`- アクセスレベル: ${accessLevel}`);
    Logger.log(`- 権限: ${permissions}`);

    return {
      status: 'SUCCESS',
      fileName: file.getName(),
      accessLevel: accessLevel.toString(),
      permissions: permissions.toString(),
      owners: file.getOwner() ? file.getOwner().getEmail() : 'Unknown',
      url: file.getUrl(),
    };
  } catch (error) {
    Logger.log('エラーが発生しました: ' + error.message);
    return {
      status: 'ERROR',
      message: 'ファイルアクセスエラー: ' + error.message,
    };
  }
}
