---
title: "VBAでシフト管理表と日報表を改修"
emoji: "📓"
type: "tech" # tech: 技術記事 / idea: アイデア
topics: [VBA, Excel, SQL]
published: false
---

## はじめに
昔、知人の依頼でExcelとVBAで作成したシフト管理表と日報管理表を知人より相談があったため、改修の記録を残します。

## 実装した機能
::::details シフト管理表
- 月が替わるたびに表を1から作成しなおしていたので対象の年月を設定するとその月の罫線を自動で作成する
![表の自動作成](/images/taimeSchedule-dailyLog-app/preSample.gif)
- 勤務時間の入力時が1なら9:00~17:00などキーを入力して作成していたため、キーから勤務時間を計算し合算する
![勤務時間の計算](/images/taimeSchedule-dailyLog-app/preSumTime.gif)
::::