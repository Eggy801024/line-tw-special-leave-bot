# 台籍預排特休 LINE 機器人

這是一套獨立的台籍預排特休機器人，程式與部署檔都放在 `台籍` 資料夾內。

## 使用方式

第一次申請請輸入工號與日期：

```text
A1234 6/3
```

第一次綁定完成後，同一個 LINE 帳號之後可只輸入日期：

```text
6/4
```

若表格日期範圍明確，也可只輸入日號：

```text
4
```

查詢日期狀況：

```text
查詢 6/3
```

顯示說明：

```text
help
```

## 規則

- 系統只接受 Google Sheet 目前表格上已開放的日期。
- 同一個日期最多 3 名人員申請特休。
- 已申請後如需變更，預設需聯絡組長。
- 若多人同時留言，系統會排隊處理，避免同一日期超收。

## Google Sheet 表格格式

主表預設分頁名稱：

```text
台籍特休
```

表格需包含以下欄位：

```text
工號 | 姓名 | 組別
```

日期欄位可放在表頭同一列，或放在下一列。系統會依照表格上的日期欄位填入 `X`。

支援表格結構範例：

```text
工號 | 姓名 | 組別/日期 | A | A | B | B
     |      |           | 6/1 | 6/2 | 6/3 | 6/4
```

## 環境變數

請參考 `.env.example`。Render 部署時請在 Environment Variables 填入 LINE 與 Google Sheets 設定。

重要預設值：

```text
SHEET_NAME=台籍特休
LOG_SHEET_NAME=台籍特休回覆紀錄
BINDING_SHEET_NAME=台籍Line綁定
MAX_PER_DATE=3
NEW_MARK=特休
ELIGIBLE_SHIFT_MARK=N1
```
