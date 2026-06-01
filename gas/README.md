# GAS（デプロイ元）

Apps Script エディタにこのフォルダの内容を反映してください。

- `Code.gs` → **コード**
- `index.html` → **index**（HTML ファイル名）

Vercel 連携 API:

- `GET ?api=dayContext&token=...` … 今日のカレンダー + Workspace 同期データ
- `POST` `{ "api": "createDailyDraft", "done": [...], "kansou": "...", "token": "..." }` … Gmail 下書き

スクリプトプロパティ `WS_API_TOKEN` を設定すると、上記 API は token 必須になります。
