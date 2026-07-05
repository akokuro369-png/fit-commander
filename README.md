# FitCommander Phase 1

FitCommander の Phase 1 成果物です。Google Sheets の初期構築と Google Apps Script WebApp API を提供します。

## Files

- `setup_sheets.gs` - 7 シートを作成し、`exercise_master` に初期種目を投入します。
- `gas_api.gs` - PWA / iPhone ショートカットから利用する WebApp API です。

## Setup

1. Google Drive で新しい Google Spreadsheet を作成し、名前を `FitCommander` にします。
2. Spreadsheet で `拡張機能 > Apps Script` を開きます。
3. `setup_sheets.gs` の内容を Apps Script に貼り付け、`setupFitCommanderSheets()` を実行します。
4. Spreadsheet URL の `/d/` と `/edit` の間にある ID をコピーします。
5. `gas_api.gs` の `SHEET_ID` をコピーした Spreadsheet ID に置き換えます。
6. `gas_api.gs` を Apps Script に追加します。
7. Apps Script の `プロジェクトの設定` でタイムゾーンを `Asia/Tokyo` に設定します。
8. `デプロイ > 新しいデプロイ > ウェブアプリ` を選びます。
9. 実行ユーザーを `自分`、アクセスできるユーザーを利用形態に合わせて設定し、デプロイします。
10. 発行された WebApp URL を PWA と iPhone ショートカットに設定します。

既存デプロイ済み環境に advice を追加する場合は、`setup_sheets.gs` を再実行しないでください。データが消えます。Spreadsheet に `advice` シートを手動追加し、1行目に `date,text` を入れてから `gas_api.gs` を再貼り付けし、新バージョンで再デプロイしてください。

## API

GET actions:

- `?action=getExercises`
- `?action=getLastWorkout&exercise_id=ex_001`
- `?action=getCalendar&year=2026&month=7`
- `?action=getToday`

POST body examples:

PWA から POST する場合は `Content-Type: text/plain` で JSON 文字列を送ってください。GAS WebApp は OPTIONS プリフライトを処理できないため、`application/json` ではなく `text/plain` を使ってプリフライトを回避します。

```json
{
  "action": "addExercise",
  "body_part": "胸",
  "name": "インクラインベンチプレス",
  "equipment": "ダンベル",
  "is_cardio": false
}
```

```json
{
  "action": "addWorkout",
  "date": "2026-07-05",
  "exercise_id": "ex_001",
  "sets": [
    { "set_no": 1, "weight_kg": 60, "reps": 8 },
    { "set_no": 2, "weight_kg": 60, "reps": 7 }
  ]
}
```

```json
{
  "action": "addBodyLog",
  "date": "2026-07-05",
  "weight_kg": 68.5,
  "body_fat_pct": 16.2,
  "source": "manual"
}
```

```json
{
  "action": "addMeditation",
  "date": "2026-07-05",
  "duration_min": 10,
  "memo": "夜"
}
```

```json
{
  "action": "receiveHealth",
  "date": "2026-07-05",
  "steps": 8420,
  "weight_kg": 68.5,
  "body_fat_pct": 16.2,
  "source": "shortcut"
}
```


```json
{
  "action": "setAdvice",
  "date": "2026-07-06",
  "text": "明日は脚を中心に軽めのボリュームで。睡眠と歩数の様子を見ながら有酸素は10分から。"
}
```
## Daily Summary

`aggregateDaily()` は `workout_log`、`body_log`、`health_log`、`meditation_log` から `daily_summary` を再生成します。日次トリガーからは引数なしラッパーの `aggregateDailyJob()` を実行します。

集計項目:

- `body_parts` - その日に実施した部位
- `total_volume` - `weight_kg * reps` の合計
- `cardio_min` - 有酸素種目の `duration_min` 合計
- `steps` - `health_log` の歩数
- `weight_kg` / `body_fat_pct` - 同日の最新 body log
- `meditation_min` - 瞑想時間合計

日次トリガーを使う場合は、Apps Script のトリガーで `aggregateDailyJob` を毎日夜に実行してください。

## PWA Update Rule

index.html を更新してデプロイする場合は、sw.js の CACHE_NAME も必ずバンプしてください。



