# 請求書自動作成システム

Google Apps Script（GAS）を利用し、作業記録スプレッドシートから当月の作業時間を自動集計し、請求書テンプレートに値を差し込んで PDF 出力・保存する自動化スクリプトです。

## セットアップ手順

### 1. リポジトリのクローン

```bash
git clone https://github.com/yourusername/invoice-creator.git
cd invoice-creator
```

### 2. 依存関係のインストール

```bash
npm install
```

### 3. Google アカウントにログイン

```bash
npm run login
```

### 4. Google Apps Script プロジェクトの作成

```bash
npm run create
```

上記コマンド実行後、生成された`.clasp.json`の`scriptId`を確認してください。

### 5. コードのビルドとデプロイ

```bash
npm run deploy
```

## 利用方法

1. Google Apps Script のプロジェクトを開き、関数`createInvoice`を実行します。
2. 当月の作業時間が自動的に集計され、請求書が生成されます。
3. 指定された Google Drive フォルダに`yyyymmdd_金額_東顕正.pdf`形式で PDF が保存されます。

## 仕様詳細

- **作業記録スプレッドシート**:

  - ファイル ID: `1vCw-kt1eGBlRmLqAu8urEoX-oNYJbbbgqGDfLW7h9JI`
  - シート名: 作業記録
  - カラム構成:
    - A: 作業 ID
    - B: 作業日
    - C: 作業時間
    - D: 作業内容
    - E: 作成日
    - F: 更新日

- **請求書テンプレート**:
  - ファイル ID: `1gv2H0YM5qxhWm_bJTPjmFBK3YHXfFZrdRHXrG6bmSgg`
  - 出力先フォルダ ID: `1FKxvmPu4icl4XOWZC9JM1Xc1mpMXnn5A`
  - 設定セル:
    - E4: 請求日
    - E5: 請求書番号
    - C33: 作業時間
    - F30: 請求金額（税込）
  - PDF 出力対象: 「ダウンロード用」シート

## 作業記録の注意点

作業記録の「作業日」列は月次集計対象となります。正しく当月分のみが集計されるように、適切な日付を入力してください。
