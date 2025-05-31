// グローバル型定義
declare namespace NodeJS {
  interface Global {
    createInvoice(): string;
  }
}

// 既存のMimeType定義と競合を回避
declare interface MimeType extends GoogleAppsScript.Base.MimeType {}
