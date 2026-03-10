---
title: "VSCodeでVBAを開発するための環境構築"
emoji: "💻"
type: "tech" # tech: 技術記事 / idea: アイデア
topics: ["VSCode", "VBA"]
published: false
---

## はじめに
VBAでの開発を行う際にブラウザで生成AIを利用しながら作業を行うことが多く
さらに開発スピードを向上やコード管理を容易にするためにVSCodeを利用した開発環境を整える。
開発環境の構築や設定方法等を忘れがちであるため備忘録として残します。

### やりたいこと
- エディタ内で開発を行えるようにしAIへコードのコピペ作業を無くしたい
- コードの修正内容や変更状態を管理するためGitを利用したい

## VSCodeの拡張機能
1. VBA(発行元 : serkonda7)
用途 : コードをハイライトし見やすくする
![VSCode拡張機能VBA](/images/vscode_env_vba/VSCode-VBA.png)
2. XVBA - Live Server VBA(発行元 : Local Smart)
用途 : Excel⇔VSCodeの連携を行う
:::message
    npmコマンドを使用するため事前にnode.jsが必要
:::
![VSCode拡張機能XVBA](/images/vscode_env_vba/VSCode-XVBA.png)

## 初期設定
#### VSCode
::::details 1. エンコードの設定(必要な人のみ)
設定画面より以下のFiles:Auto Guess Encodingにチェックを入れる
![Files Auto Guess Encoding](/images/vscode_env_vba/FilesAutoGuessEncoding.png)
:::message
VBAはShift-JISではないと日本語が文字化けしてしまうため、
VSCodeにShift-JISで開くコードかどうか判別できるようにする
:::
::::
2. フォルダをExcel VBAのプロジェクトフォルダにする
設定ファイルの作成
各ファイルを保存するフォルダを作成
![プロジェクトフォルダの作成](/images/vscode_env_vba/XVBA-BootStrap.gif)
3. config.jsonの設定
どのExcelと紐づけるか設定する
```json:config.json
    {
    "app_name": "XVBA",
    "description": "",
    "author": "",
    "email": "",
    "create_ate": "Mon Mar 09 2026 11:50:45 GMT+0900 (GMT+09:00)",
    "excel_file": "test.xlsm", // 紐づけたいExcelのファイル名
    "vba_folder": "vba-files",
    "ribbon_file": "customUI14",
    "ribbon_folder": "ribbons",
    "logs": "on",
    "xvba_packages": {},
    "xvba_dev_packages": {}
}
```
4. プロジェクトフォルダにて`npm install`を行い 
package.jsonから必要なパッケージをインストールする
5. Import-VBAの実施
対象のファイル内に保管されているモジュールやクラスのデータを取得する。
![XVBA-Import](/images/vscode_env_vba/XVBA-Import.png)
:::message
Import実行時に書き画面が表示された場合は"Development"を選択
![確認画面](/images/vscode_env_vba/Development.png)
:::
![vba-files](/images/vscode_env_vba/XVBA-vba-files.png =200x)
*Import後*
6. Export All VBAの実施
ファイルの編集の完了後、Excel内のモジュール等を更新する
![XVBA-ExportAll](/images/vscode_env_vba/XVBA-ExportAll.png)
:::message
Export実行時に書き画面が表示された場合は"Development"を選択
![確認画面](/images/vscode_env_vba/Development.png)
:::


## つまずいた所
::::details 1. Import-VBA
npm installが必要であり、一部バージョンの問題でエラーが発生
```diff json:package.json
{
    "name": "xvba-app",
    "version": "1.0.0",
    "description": "A XVBA App",
    "main": "index.js",
    "author": "LocalSmart",
    "license": "ISC",
-    "dependencies": {
-        "excel-types": "1.0.0",
-        "Xlog": "1.0.0"
-    },
    "devDependencies": {
        "@localsmart/xvba-cli": "^1.0.2"
    }
}
```
不要な部分を削除することで`npm install`が正常にできモジュールなどを取得できた。
::::

::::details 2. Export All VBA
動作環境によっては可能化もしれないが対象のxlsmファイルを開いてエクスポートを実行することで解決できた。
::::

## まとめ
今回、VBAのコードをGitで管理しやすくしコードの確認をAIで行いやすくするためにVSCodeに環境を構築したが、packageのインストールでつまずいたりと改めてnpmについて学習する機会にもなって良かったと思いました。
同様なところでつまずいた方の助けになればと思います。