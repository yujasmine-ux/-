# WEB-預算上傳平台

使用 Apps Script 開發後端，串接 Google Sheets 試算表，並透過 HTML、CSS、JavaScript 設計美化後的前端介面，讓使用者可以手動或透過檔案上傳資料，實現資料自動同步與管理。


# 主要功能說明
1.請求統一處理：

建立一個新的主函數 handleRequest(e) 來集中處理所有來自前端的請求。

這讓 doGet 和 doPost 函數變得非常簡單，它們都只是調用 handleRequest。

根據前端傳來的 action 參數（例如 ?action=get），handleRequest 會自動執行對應的函數 (getSheetData 或 syncData)。

2.資料驗證：

在 syncData 函數中，我新增了對 item 內各個欄位的資料驗證。

使用 String(item.store_name || '').trim() 確保 store_name 是一個有效的字串。

使用 isNaN() 來檢查 year, month, amount 是否為有效的數字。

這樣可以防止前端傳來無效的資料（如空值或非數字）時，腳本寫入錯誤內容到試算表中。

3.優化更新邏輯：

原始程式碼的 for 迴圈在處理大量資料時效率較低。

新版本使用了 Map 來儲存現有資料，將尋找資料的時間複雜度從 O(n) 降低到 O(1)。這在試算表資料量變大時，能大幅提升更新速度。

4.一次性寫入：

新增資料不再使用 appendRow 一條條寫入。

我將所有待新增的資料收集到 newRows 陣列中，然後使用 setValues 一次性寫入。這可以減少對 Apps Script 服務的調用次數，從而提高整體效能。

5.更清晰的回傳訊息：

syncData 函數現在會回傳一個包含已新增和已更新筆數的詳細訊息，方便您在前端顯示。
