
# 開始使用 Office Mix 的 LabsJS



LabsJS 內容會公開 API (labs.js)、範例、文件及相關的檔案，您可以用來開發互動式實驗室、將它們整合到 Office Mix，然後在 Microsoft PowerPoint 中呈現它們。這些實驗室其實是您使用 HTML5 和 labs.js JavaScript 程式庫所建立的 Office 增益集。

## LabsJS 內容

LabsJS 提供文件、範例實驗室，和建立及發佈您自己的 Office Mix 實驗室的必要檔案。


**必要檔案**


|**File**|**說明**|
|:-----|:-----|
|labs-1.0.4.js|Office Mix 實驗室開發的 LabsJS JavaScript API。這個檔案必須包含在您的專案中，使其與 Office Mix 整合。檔案也裝載於 <code>https://az592748.vo.msecnd.net/sdk/LabsJS-1.0.4/labs-1.0.4.js</code> 上的內容傳遞網路 (CDN)。當您發佈應用程式時，您必須連結到 CDN 上的檔案。|
|labs-1.0.4.d.ts|labs.js 的 TypeScript 定義檔。如此一來，可以輕易地將您的 TypeScript 程式碼整合 labs.js。定義檔也提供 labs.js 中包含的所有元件的廣泛概觀。您可以從 [http://www.typescriptlang.org/](http://www.typescriptlang.org/) 下載 TypeScript。定義檔是根據 TypeScript 0.9.1.1 版所建立。|
|歷程記錄|各種版本的 labs.js 程式庫的發行歷程記錄。|
|Labshost.html|可讓您針對 PowerPoint 內容外部的 Office Mix 檢視和偵錯實驗室的網頁。若要使用網頁，請在主要的輸入方塊中鍵入 URL，隨後會在框架中載入它。在右邊的輸入方塊中，會顯示在 PowerPoint 或 Office Mix 課程播放程式中執行時，API 與 Office Mix 之間交換的資料。也可以預先植入資料。請注意，＜範例＞一節中的範例實驗室顯示主機內容中執行的現有 Office Mix 增益集。|
|SampleManifest.xml|範例 Office 增益集資訊清單，做為建立您自己的應用程式資訊清單的範本。|
|Simplelab.html|使用 labs.js 建立的範例實驗室。允許選擇網頁和插入網頁，然後追蹤檢視它的使用者。|
|Simplelab.ts|用來建立 simplelab 範例的 TypeScript 檔案。|
|Simplelab.js|Simplelab 範例的 JavaScript 版本。此範例和 simplelab.ts 皆示範 LabsJS API 的使用。|

## 設定開發環境

Labs.js 程式庫做為 office.js 程式庫頂端的抽象層 (Office 增益集的 API)，因此您使用 labs.js 程式庫建立的實驗室實際上是 Office 增益集。為了要使用 labs.js 程式庫，及在 Office Mix 中執行這些實驗室，您必須先將自己設定為 Office 增益集開發人員。


### 註冊 Office 365 開發人員網站

第一個步驟是申請 Office 365 開發人員網站。這可讓您先裝載和測試實驗室，再將其提交給 Office 市集。網站可讓您將增益集發佈至 Office Mix，並在真實的環境中進行測試。

如需詳細資訊，請參閱[在 Office 365 上設定 SharePoint 增益集的開發環境](http://msdn.microsoft.com/library/b22ce52a-ae9e-4831-9b68-c9210af6dc54%28Office.15%29.aspx)。您只需要遵循前兩個步驟；可選擇安裝 "Napa" 開發人員工具。


### 在 SharePoint Online 上設定應用程式目錄

建立和佈建開發人員網站之後，接著在 SharePoint Online 上設定增益集目錄。如需詳細資訊，請參閱[在 Office 365 上設定增益集目錄](../../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)。

對於 Office Mix，您會使用增益集目錄，讓您可以將生產前的增益集插入課程中，並先進行端對端測試再將實驗室提交至存放區。


## 建立實驗室

如果要建立第一個實驗室，請依照[逐步解說](../../powerpoint/office-mix/creating-your-first-lab-for-office-mix.md)中的步驟，其中解釋如何建立簡單的真/假測驗。請參閱[逐步解說︰建立第一個 Office Mix 實驗室](../../powerpoint/office-mix/creating-your-first-lab-for-office-mix.md)。


## 發佈實驗室

建立實驗室後，您可以將其發佈並提交至存放區。


### 建立和上載您的應用程式資訊清單

應用程式資訊清單是描述 LabJS 實驗室的 XML 文件。它提供裝載實驗室的 URL 參考，並提供實驗室的詳細資訊，包括顯示名稱、描述、圖示、大小等等。

我們加入了範例資訊清單 "SampleManifest.xml"。如需資訊清單結構描述的詳細資訊，以及結構描述定義的連結，請參閱 [Office 增益集的 XML 資訊清單](../../../docs/overview/add-in-manifests.md)。

若要將資訊清單上載至 SharePoint 網站，請先瀏覽到您的應用程式目錄，通常位於 URL <code>https://\<your site\>/sites/AppCatalog</code>。 然後，選擇 [新增應用程式]**** 按鈕，並依照步驟來上載您的應用程式資訊清單。


### 更新您的 PowerPoint 2013 目錄

接著更新您的 PowerPoint 2013 目錄。您之後可使用您的開發人員帳戶登入。

更新 PowerPoint 2013 目錄以啟動。 啟動 PowerPoint 2013 並瀏覽功能表路徑 [檔案 > 選項 > 信任中心 > 信任中心設定 > 信任應用程式目錄]****。 從那裡，將參考加入至您的應用程式目錄，然後選擇 [確定]****。 PowerPoint 2013 會要求您先登出，讓變更生效。 登出。

最後，使用開發人員帳戶再重新登入。選擇 PowerPoint 2013 右上角的登入名稱，並使用您的開發人員帳戶登入。您現在可以插入您的增益集。


### 插入、發佈及檢視您的應用程式

若要將增益集插入目錄中，請選擇 [插入]**** 功能區，然後選擇 [應用程式]**** 區段中的 [存放區]****。 選擇 [我的組織]****，且您會在增益集類別中看到增益集。 選擇增益集，選取 [插入]****，並將增益集 (實驗室) 插入 PowerPoint 2013 文件中。

現在您也可以利用所有可用的 Office Mix 功能，使用新實驗室發佈課程。


 >**重要**：若要檢視應用程式，您必須從檢視課程的相同瀏覽器中登入到您的 SharePoint 目錄。SharePoint 目錄只允許來自已驗證的使用者存取，因此若要查看您的應用程式，您必須先行登入。 


### 將實驗室提交至 Office 市集

若要將實驗室提交至 Office 市集，請參閱[發佈 Office 增益集](../../publish/publish.md)


## 其他資源



- [Office Mix 增益集](../../powerpoint/office-mix/office-mix-add-ins.md)
    
- [Office 增益集](../../../docs/overview/office-add-ins.md)
    
- [建立第一個 Office Mix 實驗室](../../powerpoint/office-mix/creating-your-first-lab-for-office-mix.md)
    
