# Render 部署

這個資料夾是台籍預排特休機器人，可獨立部署到 Render。

## Render 建立服務

1. 新增 `Web Service`
2. 連接台籍機器人的 GitHub repository
3. 設定：

```text
Language: Node
Build Command: npm install
Start Command: node src/server.js
```

若使用 `render.yaml`，服務名稱預設為：

```text
line-tw-special-leave-bot
```

## Environment Variables

```text
LINE_CHANNEL_SECRET=
LINE_CHANNEL_ACCESS_TOKEN=
GOOGLE_SPREADSHEET_ID=
GOOGLE_SERVICE_ACCOUNT_EMAIL=
GOOGLE_PRIVATE_KEY=
TIME_ZONE=Asia/Taipei
SHEET_NAME=台籍特休
LOG_SHEET_NAME=台籍特休回覆紀錄
BINDING_SHEET_NAME=台籍Line綁定
MAX_PER_DATE=3
NEW_MARK=特休
ELIGIBLE_SHIFT_MARK=N1
ALLOW_CHANGE=false
WORKER_ID_PATTERN=[A-Z]{1,3}\d{3,4}
GROUP_TEAM_MAP_JSON={}
```

Render 不要填 `GOOGLE_SERVICE_ACCOUNT_JSON_PATH`，請改填 `GOOGLE_PRIVATE_KEY`。

## LINE Webhook

部署完成後，LINE Developers 的 Webhook URL 設成：

```text
https://你的-render網址.onrender.com/webhook
```

並確認：

```text
Use webhook: 開啟
Auto-reply messages: 關閉
```
