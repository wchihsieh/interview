# 實習生面談預約系統

將預約資料同步寫入本地 Excel（`bookings.xlsx`）並可選擇同步至 Google Sheets。

---

## 快速啟動

### 1. 安裝依賴

```bash
npm install
```

### 2. 啟動後端伺服器

```bash
npm start
```

伺服器啟動後，開啟 `http://localhost:3001/interview-schedule.html` 使用預約系統。

> 每次預約成功後，資料會自動寫入同目錄下的 `bookings.xlsx`。

---

## Google Sheets 同步設定（選用）

### 步驟一：建立 Google Apps Script

1. 開啟您的目標 **Google Sheets**
2. 點選選單列「**擴充功能**」→「**Apps Script**」
3. 將 `Code.gs` 的全部內容貼入編輯器（取代原有內容）
4. 點擊「**儲存**」

### 步驟二：部署為網路應用程式

1. 點選「**部署**」→「**新增部署作業**」
2. 設定如下：
   - 類型：**網路應用程式**
   - 說明：`面談預約 API`
   - 執行身分：**我自己**
   - 存取權：**所有人**
3. 點選「**部署**」，授權後複製產生的 **部署 URL**

### 步驟三：設定環境變數

將部署 URL 設為環境變數後啟動伺服器：

```bash
# Windows PowerShell
$env:APPS_SCRIPT_URL="https://script.google.com/macros/s/你的部署ID/exec"
npm start

# Windows CMD
set APPS_SCRIPT_URL=https://script.google.com/macros/s/你的部署ID/exec
npm start
```

設定後，每次預約新增或取消，後端會自動同步至 Google Sheets。

---

## API 端點

| 方法 | 路徑 | 說明 |
|------|------|------|
| `GET` | `/api/bookings` | 取得所有預約資料 |
| `POST` | `/api/bookings` | 新增預約（body: `{ slotId, time, name }`）|
| `DELETE` | `/api/bookings/:slotId` | 取消預約 |
| `GET` | `/api/export` | 下載 Excel 檔案（`面談預約紀錄.xlsx`）|

---

## 說明

- **離線模式**：若後端未啟動，前端會自動降級使用 `localStorage` 暫存資料，不影響操作
- **Excel 路徑**：`bookings.xlsx`（與 `server.js` 同目錄）
- **Google Sheets 工作表名稱**：`預約紀錄`（由 `Code.gs` 自動建立）
