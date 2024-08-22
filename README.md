# Excel 到 Wiki 表格轉換工具

這是一個使用 C# 開發的工具,可以將 Excel 文件轉換為 Wiki 表格(HTML 格式)。如果您的辦公環境無法使用其他工具或第三方庫,這個程式可以幫助您完成轉換工作。

![image](https://github.com/swwad/Excel2WikiTable/blob/master/01.png)
![image](https://github.com/swwad/Excel2WikiTable/blob/master/02.png)

## 功能特點

- 將 Excel 文件轉換為 HTML 格式的 Wiki 表格
- 支持批量處理多個 Excel 文件
- 處理合併儲存格
- 保留原始 Excel 表格的列寬
- 不依賴第三方庫,僅使用 Microsoft Office Interop Excel

## 系統需求

- Windows 作業系統
- .NET Framework
- Microsoft Excel (用於讀取 Excel 文件)

## 使用方法

1. 運行程式
2. 點擊 "選擇目錄" 按鈕,選擇包含 Excel 文件的資料夾
3. 點擊 "解析" 按鈕開始轉換過程
4. 轉換完成後,程式會在原始 Excel 文件所在的目錄中生成對應的 .md 文件

## 注意事項

- 程式會處理選定目錄及其子目錄中的所有 .xls 文件
- 轉換過程中可能會開啟多個 Excel 進程,程式會定期關閉這些進程以釋放資源
- 如果遇到權限問題或目錄訪問錯誤,程式會跳過該文件並繼續處理其他文件

## 開發者說明

本程式主要包含以下幾個關鍵類別和方法：

- `Excel2WikiTable`: 主要的表單類別,處理用戶界面和轉換邏輯
- `SimpleExcelCell`: 用於存儲 Excel 儲存格數據的簡單數據結構
- `StackBasedIteration`: 用於遍歷目錄結構的輔助類別

主要的轉換邏輯位於 `ProcessWorkSheet` 和 `MakeHtmlTable` 方法中。
