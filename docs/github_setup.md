# GitHub.com でリポジトリを作成する手順

1. [GitHub](https://github.com/) にログインする

2. 右上のプラスアイコン(+)をクリックして、「New repository」を選択

3. 以下の情報を入力：

   - リポジトリ名: `invoice-creator`
   - 説明: `Google Apps Script による請求書自動作成システム`
   - 公開設定: Public または Private（希望に応じて）
   - README.md の初期化: チェックを外す（既に README を持っているため）
   - .gitignore の追加: チェックを外す（既に作成済み）
   - ライセンスの追加: 必要に応じて選択

4. 「Create repository」ボタンをクリック

5. その後表示される手順に従って、ローカルリポジトリをリモートリポジトリに接続：

```bash
git remote add origin https://github.com/YOUR-USERNAME/invoice-creator.git
git branch -M main
git push -u origin main
```

上記コマンドの `YOUR-USERNAME` を自分の GitHub ユーザー名に置き換えてください。
