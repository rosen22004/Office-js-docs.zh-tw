
# 更新您的適用於 Office 的 JavaScript API 和資訊清單結構描述檔案的版本



本文章說明如何將 JavaScript 檔案 (Office.js 和應用程式專屬 .js 檔案) 和您的 Office 增益集專案中的增益集資訊清單驗證檔案更新至 1.1 版。

## 使用最新的專案檔案

如果您使用 Visual Studio 來開發增益集，若要使用適用於 Office 的 JavaScript API [最新的 API 成員](../../reference/what's-changed-in-the-javascript-api-for-office.md)和[增益集資訊清單的 1.1 版功能](../../docs/overview/add-in-manifests.md) (已對 offappmanifest-1.1.xsd 經過驗證)，您必須下載並安裝 [Visual Studio 2015 和最新的 Office 開發人員工具](https://www.visualstudio.com/features/office-tools-vs)。

如果您使用文字編輯器或 Visual Studio 以外的 IDE 來開發增益集，您需要在增益集的資訊清單中更新 Office.js 的 CDN 的參考和參考的結構描述版本。

若要執行使用新的及更新的 Office.js API 和增益集資訊清單功能開發的增益集，您的客戶必須執行 Office 2013 SP1 更新版本的內部部署產品，並在適用時，執行 SharePoint Server 2013 SP1 和相關伺服器產品、Exchange Server 2013 Service Pack 1 (SP1) 或同等線上主控產品︰Office 365、SharePoint Online 和 Exchange Online。

若要下載 Office、SharePoint 和 Exchange SP1 產品，請參閱下列各項︰


- [Microsoft Office 2013 和相關的桌面產品所有的 Service Pack 1 (SP1) 更新的清單](http://support.microsoft.com/kb/2850036)
    
- [Microsoft SharePoint Server 2013 和相關的伺服器產品所有的 Service Pack 1 (SP1) 更新的清單](http://support.microsoft.com/kb/2850035)
    
- [Exchange Server 2013 Service Pack 1 的描述](http://support.microsoft.com/kb/2926248)
    

## 將使用 Visual Studio 建立的 Office 增益集專案更新，以使用最新的適用於 Office 的 JavaScript API 程式庫和 1.1 版增益集資訊清單結構描述


對於在 1.1 版的適用於 Office 的 JavaScript API 和增益集資訊清單結構描述之前建立的專案，您需要使用 **NuGet 封裝管理員**更新專案的檔案，然後更新增益集的 HTML 頁面以參考它們。 

請注意，更新程序會_以專案為基礎_套用 - 您必須為您要使用 1.1 版 Office.js 和增益集資訊清單結構描述的每個增益集專案重複更新程序。




### 將專案中適用於 Office 的 JavaScript API 程式庫檔案更新為最新版本


1. 在 Visual Studio 2015 中，開啟或建立新的 **Office 增益集**專案。
    
      - 在左窗格中，選擇 [更新]**** 並完成套件更新程序。
    
  - 繼續進行步驟 6。
    
2. 選擇 [工具]**** > **[NuGet 封裝管理員]** > **[管理方案的 Nuget 套件]**。
    
3. 在 **NuGet 封裝管理員**中，對 [套件來源]**** 選取 **nuget.org**，以及對 [篩選]**** 選取 [升級可供使用]****， 然後選取 [Microsoft.Office.js]。
    
4. 在左窗格中，選擇 [更新]**** 並完成套件更新程序。
    
5. 在增益集的 HTML 頁面的 **head** 標記中，註解化或刪除任何現有的 office.js 指令碼參考 (例如︰`<script src="https://appsforoffice.microsoft.com/lib/1.0/hosted/office.js" type="text/javascript"></script>)`)，並且現在參考更新的適用於 Office 的 JavaScript API 程式庫，就像這樣 (將版本值變更為 '1')。 

   >**附註：**CDN URL 中 'office.js' 前面的 '/1/' 指定要使用 Office.js 版本 1 內最新的累加版本。
    
```
  <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```


### 更新專案中的資訊清單檔案以使用結構描述 1.1 版


- 在您的專案的增益集資訊清單 (_projectname_ Manifest.xml) 檔案中，更新 **OfficeApp** 元素的 **xmlns** 屬性，將版本值變更為 '1.1' (保留 **xmlns** 屬性以外的屬性不變)。
    
```XML
  <OfficeApp xsi:type="ContentApp" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" >
```


>
  **附註：**將增益集資訊清單結構描述的版本更新為 1.1 之後，您必須移除 **Capabilities** 和 **Capability** 元素，並以 [Hosts 和 Host 元素](http://msdn.microsoft.com/library/cff9fbdf-a530-4f6e-91ca-81bcacd90dcd%28Office.15%29.aspx)或 [Requirements 和 Requirement 元素](../../docs/overview/specify-office-hosts-and-api-requirements.md)取代它們。

## 將使用文字編輯器或其他 IDE 建立的 Office 增益集專案更新，以使用最新的適用於 Office 的 JavaScript API 程式庫和 1.1 版增益集資訊清單結構描述


對於在 1.1 版的適用於 Office 的 JavaScript API 和增益集資訊清單結構描述之前建立的專案，您需要更新增益集的 HTML 頁面，以參考 1.1 版程式庫的 CDN，並更新增益集的資訊清單檔案以使用結構描述 1.1 版。 

更新程序會_以專案為基礎_套用 - 您必須為您要使用 1.1 版 Office.js 和增益集資訊清單結構描述的每個增益集專案重複更新程序。

您不需要適用於 Office 的 JavaScript 檔案的本機複本 (Office.js 和應用程式特定的 .js 檔案)，來開發 Office 增益集 (參考的 CDN Office.js 會在執行階段下載必要的檔案)，但如果您想要程式庫檔案的本機複本，則可以使用 [NuGet 命令列公用程式](http://docs.nuget.org/consume/installing-nuget)和 `Install-Package Microsoft.Office.js` 命令來下載它們。

 > **附註** 若要取得 1.1 版增益集資訊清單的一份 XSD (XML 結構描述定義)，請參閱 [Office 增益集資訊清單的結構描述參考 (v1.1)](../overview/add-in-manifests.md) 中的清單。


### 將專案中適用於 Office 的 JavaScript API 程式庫檔案更新以使用最新版本


1. 在文字編輯器或 IDE 中開啟增益集的 HTML 頁面。
    
2. 在增益集的 HTML 頁面的 **head** 標記中，註解化或刪除任何現有的 office.js 指令碼參考 (例如︰`<script src="https://appsforoffice.microsoft.com/lib/1.0/hosted/office.js" type="text/javascript"></script>)`)，並且現在參考更新的適用於 Office 的 JavaScript API 程式庫，就像這樣 (將版本值變更為 '1')。
    
```
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```


    The  `/1/` in front of `office.js` in the CDN URL specifies to use the latest incremental release within version 1 of Office.js.
    

### 更新專案中的資訊清單檔案以使用結構描述 1.1 版


- 在您的專案的增益集資訊清單 (_projectname_ Manifest.xml) 檔案中，更新 **OfficeApp** 元素的 **xmlns** 屬性，將版本值變更為 `1.1` (保留 **xmlns** 屬性以外的屬性不變)。
    
```XML
<OfficeApp xsi:type="ContentApp" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" >
```

>
  **附註：**將增益集資訊清單結構描述的版本更新為 1.1 之後，您必須移除 **Capabilities** 和 **Capability** 元素，並以 [Hosts 和 Host 元素](http://msdn.microsoft.com/library/cff9fbdf-a530-4f6e-91ca-81bcacd90dcd%28Office.15%29.aspx)或 [Requirements 和 Requirement 元素](../../docs/overview/specify-office-hosts-and-api-requirements.md)取代它們。
    

## 其他資源



- [指定 Office 主應用程式和 API 需求](../../docs/overview/specify-office-hosts-and-api-requirements.md)
    
- [了解適用於 Office 的 JavaScript API](../../docs/develop/understanding-the-javascript-api-for-office.md)
    
- [JavaScript API for Office](../../reference/javascript-api-for-office.md)
    
- [Office 增益集資訊清單的結構描述參考 (1.1 版)](../overview/add-in-manifests.md)
    
