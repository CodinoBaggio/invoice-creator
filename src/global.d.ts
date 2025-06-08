// グローバル型定義
declare namespace NodeJS {
  interface Global {
    createInvoice(date?: string): string;
    initializeProperties(): void;
    dailyTrigger(): string;
    setDailyTrigger(): string;
    sendNotificationEmail(isSuccess: boolean, errorMessage?: string): void;
  }
}

// グローバルオブジェクト（Google Apps Script環境で使用）
interface Window {
  createInvoice(date?: string): string;
  initializeProperties(): void;
  dailyTrigger(): string;
  setDailyTrigger(): string;
  sendNotificationEmail(isSuccess: boolean, errorMessage?: string): void;
}

// 既存のMimeType定義と競合を回避
declare interface MimeType extends GoogleAppsScript.Base.MimeType {}
