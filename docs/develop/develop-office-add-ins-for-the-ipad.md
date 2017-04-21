
# <a name="develop-office-add-ins-for-the-ipad"></a>開發適用於 iPad 的 Office 增益集


下表列出要開發執行於 Office for iPad 中 Office 增益集所需進行的工作。


|**工作**|**描述**|**資源**|
|:-----|:-----|:-----|
|更新增益集以支援 Office.js 1.1 版。|將 JavaScript 檔案 (Office.js 和應用程式專屬 .js 檔案) 和在 Office 增益集專案中所使用的增益集資訊清單驗證檔案更新至 1.1 版。|[適用於 Office 的 JavaScript API 中的變更項目](../../reference/what's-changed-in-the-javascript-api-for-office.md)|
|套用 UI 設計最佳做法。|將您的增益集 UI 與 iOS 經驗順暢整合。|[針對 iOS 設計](https://developer.apple.com/library/ios/documentation/UserExperience/Conceptual/MobileHIG/)|
|套用增益集設計最佳做法。|確定您的增益集提供清楚的值、吸引人且一致地執行。|[開發 Office 增益集的最佳做法](../../docs/overview/add-in-development-best-practices.md)|
|針對觸控功能最佳化您的增益集。|讓您的 UI 在回應滑鼠及鍵盤以外，還能夠回應觸控輸入。|[套用 UX 設計原則](https://msdn.microsoft.com/EN-US/library/mt590883.aspx#Anchor_3)|
|讓增益集可供免費使用。|iPad 版 Office 是一個管道，透過它您可以接觸更多使用者並提升您的服務。這些新的使用者有可能成為您的客戶。|[驗證原則 10.8](http://msdn.microsoft.com/library/cd90836a-523e-42f5-ab02-5123cdf9fefe%28Office.15%29.aspx)|
|讓您的增益集可免費供商務使用。|您的增益集不得包含應用程式中購買、不提供試用版，沒有目的在追加銷售為付費版或連結到任何線上市集 (使用者可以購買或取得其他內容、應用程式或增益集) 的 UI。您的隱私權原則和使用規定頁面也不得包含任何商務 UI 或市集連結。|[驗證原則 3.4](http://msdn.microsoft.com/library/cd90836a-523e-42f5-ab02-5123cdf9fefe%28Office.15%29.aspx)|
|重新提交增益集至 Office 市集。|在「賣方儀表板」中，選取 [將此增益集放在 iPad 上的 Office 增益集目錄中]**** 核取方塊，並在 Apple ID 方塊中提供您的 Apple 開發人員 ID。檢閱 [Office 市集應用程式提供者](https://sellerdashboard.microsoft.com/Assets/Content/Agreements/en-US/Office_Store_Seller_Agreement_20120927.htm)合約來確定您了解合約。|[將 Office 和 SharePoint 增益集和 Office 365 Web 應用程式提交給 Office 市集](http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx)|
對於在其他平台上執行的 Office 應用程式，您的增益集可以保持不變。您也可以根據增益集執行所在的瀏覽器/裝置提供不同的 UI。若要偵測增益集是否在 iPad 上執行，您可以使用下列的 API： 

- var isTouchEnabled = [Office.context.touchEnabled](../../reference/shared/office.context.touchenabled.md)
    
- var allowCommerce = [Office.context.commerceAllowed](../../reference/shared/office.context.commerceallowed.md)
    

## <a name="best-practices-for-developing-office-add-ins-for-ios-and-mac"></a>開發 iOS 和 Mac 適用的 Office 增益集的最佳做法

開發在 iOS 上執行的增益集時，套用下列最佳做法︰


-  **使用 Visual Studio 來開發增益集。**
    
    如果使用 Visual Studio 開發增益集，在側載至 iPad 或 Mac 上之前，您可以在執行 Windows 的 Office 主應用程式中[設定中斷點並偵錯其程式碼](../get-started/create-and-debug-office-add-ins-in-visual-studio.md#Test)。因為在 Office for iOS 或 Office for Mac 中執行的增益集，與在 Office for Windows 執行的增益集支援相同的 API，您的增益集程式碼應該在兩種平台上以相同的方式執行。
    
-  **在增益集的資訊清單中指定 API 的需求，或利用執行階段檢查。**
    
    當您在增益集資訊清單中指定 API 的需求時，Office 會判斷主應用程式是否支援這些 API 成員。如果主應用程式中具備這些 API 成員，那麼您可以在該主應用程式中使用增益集。或者，您可以進行執行階段檢查，先判斷主應用程式中是否具備某方法，再套用增益集使用該方法。執行階段檢查可確保主應用程式中一律可以使用增益集，並在方法可用時提供額外的功能。如需詳細資訊，請參閱[指定 Office 主應用程式及 API 需求](../../docs/overview/specify-office-hosts-and-api-requirements.md)。
    
如需一般增益集程式開發最佳做法，請參閱[開發 Office 增益集的最佳做法](../../docs/overview/add-in-development-best-practices.md)。


## <a name="additional-resources"></a>其他資源
<a name="bk_addresources"> </a>


- [在 iPad 和 Mac 上側載 Office 增益集](../../docs/testing/sideload-an-office-add-in-on-ipad-and-mac.md)
    
- [在 iPad 和 Mac 上偵錯 Office 增益集](../../docs/testing/debug-office-add-ins-on-ipad-and-mac.md)
    
