
# 從其內容傳遞網路 (CDN) 參考適用於 Office 的 JavaScript API


[JavaScript API for Office](../../reference/javascript-api-for-office.md) 程式庫包含 Office.js 檔案和關聯的主應用程式特定的 .js 檔案，例如 Excel-15.js 和 Outlook-15.js。 


參考 API 的最簡單方法是藉由將下列 `<script>` 新增至您的頁面的 `<head>` 標記，使用我們的 CDN：  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

CDN URL 中的 `/1/` 前面的 `office.js` 會指定使用 Office.js 第 1 版的最新的累加版本。因為適用於 Office 的 JavaScript API 會維護回溯相容性，最新的版本會繼續支援第 1 版之前導入的 API 成員。如果您需要升級現有專案，請參閱 [更新您的適用於 Office 的 JavaScript API 和資訊清單結構描述檔案的版本] (../docs/develop/update-your-javascript-api-for-office-and-manifest-schema-version.md)。 

如果您計劃從 Office 市集發佈 Office 增益集，您必須使用這個 CDN 參考。本機參考僅適用於內部、開發和偵錯案例。

> **重要事項：**在開發任何 Office 主應用程式的增益集時，從頁面的 `<head>` 區段中參考適用於 Office 的 JavaScript API 很重要。 如此一來，可確保在任何本文元素之前完全初始化 API。 Office 主應用程式需要增益集在啟用的 5 秒內初始化。 超過這個臨界值會使增益集宣告為沒有回應，並且會向使用者顯示錯誤訊息。       

## 其他資源



- [了解適用於 Office 的 JavaScript API](../../docs/develop/understanding-the-javascript-api-for-office.md)
    
- [Office 增益集平台概觀](../../docs/overview/office-add-ins.md)
    
- [Office 增益集開發生命週期](../../docs/design/add-in-development-lifecycle.md)
    
- [JavaScript API for Office](../../reference/javascript-api-for-office.md)
    
