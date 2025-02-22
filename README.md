# Course Feedback Processor

A Google Apps Script tool for processing course feedback data.

## Features

- Extract feedback from spreadsheet responses automatically
- Process feedback data for multiple instructors 
- Generate structured feedback reports

## Setup

1. Create a new Google Apps Script project
2. Copy the content of `processFeedback.gs` into the project
3. Configure the following constants:
   - `FEEDBACK_SOURCE_SPREADSHEET_ID`: ID of the source spreadsheet
   - `HACKFOLDR_SPREADSHEET_ID`: ID of the output spreadsheet
   - `TEACHER_NAME`: Name of the instructor
   - `FEEDBACK_SHEET_NAME`: Name of the response sheet

## Core Functions

- `createTeacherFeedback()`: Main execution function
- `processFeedbackData()`: Process feedback data
- `getFeedbackColumn()`: Get feedback column
- `getTeacherSpreadsheetUrl()`: Get instructor's spreadsheet URL

## Notes

- Ensure proper Google Spreadsheet access permissions
- Form field names must match the expected format in the code

---

# 課程回饋處理工具

一個用於處理課程回饋資料的 Apps Script 工具。

## 功能

- 自動從試算表回應中提取課程回饋
- 處理多位講師的課程回饋資料
- 產生結構化的回饋報告

## 設定方式

1. 在 Google Apps Script 中建立新專案
2. 複製 `processFeedback.gs` 的內容到專案中
3. 設定以下常數：
   - `FEEDBACK_SOURCE_SPREADSHEET_ID`: 來源試算表的 ID
   - `HACKFOLDR_SPREADSHEET_ID`: 輸出試算表的 ID  
   - `TEACHER_NAME`: 講師姓名
   - `FEEDBACK_SHEET_NAME`: 回應工作表名稱

## 主要功能

- `createTeacherFeedback()`: 主要執行函式
- `processFeedbackData()`: 處理回饋資料
- `getFeedbackColumn()`: 取得回饋欄位
- `getTeacherSpreadsheetUrl()`: 取得講師的試算表網址

## 注意事項

- 請確保有適當的 Google Spreadsheet 存取權限
- 表單欄位名稱需符合程式中的預期格式 