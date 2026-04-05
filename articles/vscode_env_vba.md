---
title: "Excelマクロ⇔ローカルへのデータ連携ツールの作成記録"
emoji: "💻"
type: "tech" # tech: 技術記事 / idea: アイデア
topics: ["VSCode", "VBA"]
published: false
---

## はじめに
Excelのマクロを編集中によく以前のバージョンに戻したいと思う機会が多く、
VBAもGitを使用してバージョン管理を行いたいと思いました。
VSCodeの拡張機能である[XVBA]を使用することでも可能ですが、個人的に使用しにくかったため、Excel⇔ローカルへの変換作業を簡易に行えるツールを自作した際の記録となります。

### やりたいこと
- コードの修正内容や変更状態を管理するためGitを利用したい
- Excelからエクスポートを作業を楽にしたい
- Excelへのインポートを作業を楽にしたい

## XVBAを試した感想
XVBAを使ってみた感想ですが、node.jsを普段使用している方には問題はないと思いますが
普段使用しない自分のような人間には手間がかかる印象がありました。
また、正常に動作していたのに別の日では上手く動かなくなるなどもあり、
原因を調査などに時間を取られるのは面倒なこともあり今回は自作することにしました。

::::: details XVBAで行った設定
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
::::details 1. Import-VBA 実行エラー
npm installが必要であり、一部バージョンの問題でエラーが発生  
ここで削除した部分が原因で正常に動作しなくなった？
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
不要な部分を削除することで`npm install`が正常に完了し、モジュールとクラスの取得ができた。
::::

::::details 2. Export All VBA 実行エラー
動作環境によっては可能化もしれないが対象のxlsmファイルを開いてエクスポートを実行することで解決できた。
::::
:::::

## Excel⇔VSCodeの開発
### 開発環境
Windows11で作業を行い、追加のインストールなどを不要にしたかっため、PowerShellを使用して作成
OS: Windows 11
エディタ：VSCode
言語: PowerShell

### 要件
1. Excelから各モジュールをローカルへエクスポートする
2. ローカルで編集したファイルをExcelにインポートする
3. ツールがローカル内に存在すれば特定のコマンドで開発環境を構築を簡単にする
4. VSCodeのタスク機能を利用してGUIで操作を行えるようにする

### 完成画面
VSCodeで`Ctrl + Shift + B`で操作画面を表示
![操作画面](https://github.com/re-fuku/vscode-xl-converter/blob/main/assets/tasks.png?raw=true)

初期設定が完了すれば、1度コマンドを入力すれば上画面を表示し、
インポートとエクスポートの処理が簡単に行えるようになります。

### 使用方法
環境設定や使用方法については以下を参照

https://github.com/re-fuku/vscode-xl-converter.git


### エラー対応
:::details 1. try-catchでエラー
### 内容
コードを作成して実行確認をした際にtry-catchで{}が正常に閉じられていないというエラーが発生
### 原因
PowerShellがUTF-8でエンコードされていた
### 対応
ファイルをUTF-8 With BOMで保存しなおすことで解決
:::
:::details 2. PowerShell プロファイル設定
### 内容
VSCode内のターミナルで`code $PROFILE`でPowerShellのプロファイルの編集が行えるが
編集を行ったVSCodeとは異なるウインドウでプロファイルの内容が更新されていなかった。
### 原因
VSCode内では2種類のPowerShellが存在していた。
1. 通常のPowerShell
2. VSCodeの拡張機能のPowerShell⇐デバックの際に便利
    | 名称 | プロファイル名 |
    | ---- | ---- |
    | powershell (通常)| Microsoft.PowerShell_profile.ps1 |
    | PowerShell Extension (拡張機能)| Microsoft.VSCode_profile.ps1 |


:::

### 課題
1. Excelを開いたままインポートしても、Excel内のモジュールへの変更が行えなかった。
⇒Excelが開いていることにより編集ロックがかかっており外部(PowerShell)から操作を行うことが出来ないことが原因であると考えられる

## 学んだこと
- VSCodeのタスクは`tasks.json`内で定義されたものを使用するため事前に用意することができる
- PowerShellのプロファイルに関数を定義することでコマンドを自作できる
- PowerShellのプロファイルは拡張機能と通常のもので分けて作成されている



## 感想
VSCodeを使い慣れているので編集作業とGitでのバージョン管理は行いやすいが、
Excelを開いたままだとインポートができないのとデバックの際にはExcelとVBEを使用しないといけないため、開くウインドウが増えるなど今回の作業は手間が増えたように感じました。
Excelを使用する必要があるのでExcel内で完結できるようにアドインなどを開発してGitへの連携を行うのはありではないかと思いました。
また、VSCodeを普段使用していてもまだまだ仕組みやできることなど知らないことが多くとても勉強になりました。