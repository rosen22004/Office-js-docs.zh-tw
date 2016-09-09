
# 在 Office Online 中偵錯增益集


例如，如果您正在 Mac 上開發，您可以在未執行 Windows 或 Office 2013 或 Office 2016 桌面用戶端的電腦上建置和偵錯增益集。本文說明如何使用 Office Online 來測試和偵錯您的增益集。 

若要開始使用：


- 如果您沒有 Office 365 開發人員帳戶，或具有 SharePoint 網站的存取權，請取得一個帳戶。
    
     >**附註** 若要申請免費的 Office 365 開發人員帳號，請加入我們 [Office 365 開發人員程式](https://dev.office.com/devprogram)。
     
- 在 Office 365 (SharePoint Online) 上設定增益集目錄 增益集目錄是裝載 Office 增益集的主機文件庫的 SharePoint Online 中的專用網站集合。 如果您有自己的 SharePoint 網站，您可以設定增益集目錄文件庫。 如需相關資訊，請參閱[在 SharePoint 上發佈工作窗格和內容增益集至增益集目錄](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)。
    

## 從 Excel Online 或 Word Online 偵錯增益集

若要使用 Office Online 偵錯增益集︰


1. 將增益集部署到支援 SSL 的伺服器。
    
     >**附註︰**我們建議您使用 [Yeoman 產生器](https://github.com/OfficeDev/generator-office) 來建立和裝載增益集。
     
2. 在您的[增益集資訊清單檔案](../../docs/overview/add-in-manifests.md)中，更新 **SourceLocation** 元素值以加入絕對、而非相對 URI。 例如：
    
    ```xml
    <SourceLocation DefaultValue="https://localhost:44300/App/Home/Home.html" />
    ```
    
3. 在 SharePoint 的增益集目錄中，將資訊清單上載至 Office 增益集程式庫。
    
4. 在 Office 365 中從應用程式啟動程式中啟動 Excel Online 或 Word Online，並開啟新文件。
    
5. 在 [插入] 索引標籤上，選擇 [我的增益集]**** 或 [Office 增益集]****，在應用程式中插入和測試您的增益集。
    
6. 使用您的最愛瀏覽器工具偵錯工具來偵錯增益集。
    
    以下是解決您偵錯時可能遇到的問題的一些秘訣：
    
  - 您看到的一些 JavaScript 錯誤可能來自 Office Online。
    
  - 瀏覽器可能會顯示您將需要略過的不正確憑證錯誤。
    
  - 如果您在程式碼中設定中斷點，Office Online 可能會擲回錯誤來指出其無法儲存。
    

## 其他資源


- [開發 Office 增益集的最佳做法](../overview/add-in-development-best-practices.md)
    
- [提交給 Office 市集的應用程式和增益集的驗證原則 (1.9 版)](http://msdn.microsoft.com/library/cd90836a-523e-42f5-ab02-5123cdf9fefe%28Office.15%29.aspx)
    
- [建立有效的 Office 市集應用程式和增益集](http://msdn.microsoft.com/library/c66a6e6b-2e96-458f-8f8c-2a499fe942c9%28Office.15%29.aspx)
    
- [疑難排解 Office 增益集的使用者錯誤](../testing/testing-and-troubleshooting.md)
    
