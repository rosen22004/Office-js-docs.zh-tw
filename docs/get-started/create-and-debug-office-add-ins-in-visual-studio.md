
# 在 Visual Studio 中建立和偵錯 Office 增益集




 >**附註：**這些指示以 Visual Studio 2015 為基礎。如果您正在使用另一個版本的 Visual Studio，程序可能會稍有不同。



## 在 Visual Studio 中建立 Office 增益集專案


若要開始，請確定您已安裝 [Office 開發人員工具](https://www.visualstudio.com/features/office-tools-vs.aspx)。 


1. 在 Visual Studio 功能表列中，選擇 [檔案]****  >  [新增]****  >  [專案]****。
    
2. 在 **Visual C#** 或 **Visual Basic** 下的專案類型清單中，展開 **Office/SharePoint**，選擇 [Web 增益集]****，然後選擇其中一個增益集專案。  
    
3. 為專案命名，然後選擇 [確定]**** 來建立專案。
    
4. Visual Studio 會即建立專案，而且其檔案將出現在**方案總管**中。預設的 Home.html 頁面隨即在 Visual Studio 中開啟。
    
在 Visual Studio 2015 中，某些增益集專案範本已更新，以反映其他功能︰


- 除了 Excel 試算表，內容增益集可以出現在 Access 及 PowerPoint 文件的內文中。您也可以選擇「基本專案」選項來建立具有最基本的起始程式碼的基本內容增益集專案，或「文件視覺效果專案」選項 (僅適用 Access 和 Excel) 以建立功能更完整的內容增益集，其中包含以視覺化方式檢視和繫結至資料的起始程式碼。
    
- Outlook 增益集包含的選項不只可用於在電子郵件或約會中併入增益集，還可在撰寫和讀取電子郵件或約會時，用於指定增益集是否可供使用。
    

 >**附註：**在 Visual Studio 中，大部分選項都可透過描述了解，除了 [電子郵件訊息]**** 核取方塊。 如果您想要建立不只是對郵件項目顯示，也會對會議邀請、回覆和取消顯示的 Outlook 增益集，請使用該核取方塊。

完成精靈後時，Visual Studio 會為您建立包含兩個專案的方案。



|**Project**|**說明**|
|:-----|:-----|
|增益集專案|僅包含一個 XML 資訊清單檔，其中包含描述增益集的所有設定。這些設定可以協助 Office 主應用程式判斷何時應啟動增益集，以及增益集應該顯示在何處。Visual Studio 會為您產生此檔案的內容，以便您可以立即執行專案及使用增益集。您隨時可以使用資訊清單編輯器來變更這些設定。|
|Web 應用程式專案|包含增益集的內容頁面，包括開發 Office 感知的 HTML 和 JavaScript 頁面所需的所有檔案和檔案參考。開發增益集時，Visual Studio 會在您的本機 IIS 伺服器上主控 Web 應用程式。當您準備好要發佈時，您必須找到要主控這個專案的伺服器。若要進一步了解 ASP.NET Web 應用程式專案，請參閱 [ASP.NET Web 專案](http://msdn.microsoft.com/en-us/library/cdcd712f-96b0-4165-8b5d-9d0566650a28%28Office.15%29.aspx)。|

## 修改增益集設定


若要修改增益集的設定，請編輯專案的 XML 資訊清單檔案。 在**方案總管**中展開增益集專案節點，接著展開包含 XML 清單的資料夾，然後選擇 XML 清單。 您可以指向檔案中的任何元素，以檢視描述元素用途的工具提示。 如需資訊清單檔案的更多相關資訊，請參閱 [Office 增益集的 XML 資訊清單](../../docs/overview/add-in-manifests.md)。


## 開發增益集的內容


增益集專案可讓您修改描述增益集的設定，而 Web 應用程式則提供會出現在增益集中的內容。 

Web 應用程式專案包含預設的 HTML 頁面和 JavaScript 檔案，供您開始使用。專案也包含加入至專案的所有頁面通用的 JavaScript 檔案。這些檔案很方便，因為它們包含其他 JavaScript 程式庫 (包含適用於 Office 的 JavaScript API) 的參考。 

隨著增益集變得更複雜，您可以加入更多的 HTML 和 JavaScript 檔案。您可以使用預設 HTML 和 JavaScript 檔案的內容，做為您可能想要加入至專案中其他頁面，以讓該頁面使用增益集的參考類型範例。下表說明預設的 HTML 和 JavaScript 檔案。



|**File**|**說明**|
|:-----|:-----|
|**Home.html**|位於專案的 **Home** 資料夾，這是增益集的預設 HTML 頁面。在文件、電子郵件訊息或約會項目中啟動增益集時，這個頁面會顯示為增益集內的第一頁。這個檔案很方便，因為它包含開始使用所需的所有檔案參考。當您準備好要建立第一個增益集時，只需將您的 HTML 程式碼加入至這個檔案。|
|**Home.js**|位於專案的 **Home** 資料夾中，這是與 Home.js 頁面相關聯的 JavaScript 檔案。您可以在 Home.js 檔案中放置 Home.html 頁面的任何特定行為的程式碼。Home.js 檔案中包含可幫助您開始的一些範例程式碼。|
|**App.js**|位於專案的 **Add-in** 資料夾中，這是整個增益集的預設 JavaScript 檔案。您可以在 App.js 檔案中放置增益集的多個頁面的行為通用的程式碼。App.js 檔案中包含可幫助您開始的一些範例程式碼。|

 >**附註：**您不一定要使用這些檔案。請隨意加入其他檔案至專案中，並改為使用那些檔案。如果您想要將另一個 HTML 檔案顯示為增益集的初始頁面，請開啟資訊清單編輯器，然後指向該檔案名稱的 **SourceLocation** 屬性。


## 偵錯增益集


當您準備啟動增益集，請檢閱建置和偵錯相關的屬性，然後啟動方案。


### 檢閱建置和偵錯屬性

啟動方案之前，請驗證 Visual Studio 會開啟您想要的主應用程式。該資訊會出現在專案的屬性頁，並出現與建置和偵錯增益集相關的其他幾個屬性。


### 開啟專案的屬性頁面


1. 在**方案總管**中，選擇專案名稱。
    
2. 在功能表列上，選擇 [檢視]**** > [屬性視窗]****。
    
下表描述專案的屬性。



|**屬性**|**說明**|
|:-----|:-----|
|**啟動動作**|指定是否要在 Office 桌面用戶端，或在指定的瀏覽器中的 Office Online 用戶端中偵錯增益集。|
|**起始文件** (僅限內容和工作窗格增益集)|指定當您啟動專案時要開啟的文件。|
|**Web 專案**|指定與增益集相關聯的 Web 專案的名稱。|
|**電子郵件地址** (僅限 Outlook 增益集)|指定 Exchange Server 或 Exchange Online 中，您想要測試 Outlook 增益集之使用者帳戶的電子郵件地址。|
|**EWS URL** (僅限 Outlook 增益集)|Exchange Web 服務 URL (例如︰https://www.contoso.com/ews/exchange.aspx)。 |
|**OWA URL** (僅限 Outlook 增益集)|Outlook Web App URL (例如︰https://www.contoso.com/owa)。|
|**使用者名稱** (僅限 Outlook 增益集)|指定 Exchange Server 或 Exchange Online 中，您的使用者帳戶的名稱。|
|**專案檔案**|指定包含組建、組態和與專案相關的其他資訊的檔案名稱。|
|**專案資料夾**|專案檔的位置。|

### 使用現有文件來偵錯增益集 (僅限內容和工作窗格增益集)


您可以將文件加入增益集專案。如果您有文件其中包含您想要搭配增益集使用的測試資料，當您啟動專案時，Visual Studio 會為您開啟該文件。


### 使用現有文件來偵錯增益集


1. 在**方案總管**中，選擇增益集專案資料夾。
    
     >**附註**  選擇增益集專案和非 Web 應用程式專案。
2. 在 [專案]**** 功能表中，選擇 [加入現有項目]****。
    
3. 在 [加入現有項目]**** 對話方塊中，找出並選取您想要加入的文件。
    
4. 選擇 [加入]**** 按鈕，將文件加入至您的專案。
    
5. 在**方案總管** 中，開啟專案的快顯功能表，然後選擇 [屬性]****。
    
    專案的屬性頁面隨即出現。
    
6. 在 [起始文件]****清單中，選擇您加入至專案中的文件，然後選擇 [確定]**** 按鈕來關閉屬性頁面。
    

### 啟動方案


當您啟動它時，Visual Studio 會自動建置方案。 您可以從**功能表**列選擇 [偵錯]**** > [開始]**** 來啟動方案。 


 >**附註：**如果 Internet Explorer中未啟用指令碼偵錯，您將無法在 Visual Studio 中啟動偵錯工具。 您可以透過開啟 [網際網路選項]**** 對話方塊，選擇 [進階]**** 索引標籤，然後清除 [停用指令碼偵錯 (Internet Explorer)]**** 和 [停用指令碼偵錯 (其他)]**** 核取方塊來啟用指令碼偵錯。

Visual Studio 會建置專案，並執行下列︰


1. 建立 XML 資訊清單檔案的複本，並將它加入至 _ProjectName_\Output 目錄。當您啟動 Visual Studio 並偵錯增益集時，主應用程式會取用這個複本。
    
2. 在您的電腦上建立一組登錄項目，以讓增益集可在主應用程式中出現。
    
3. 建置 Web 應用程式專案，然後將它部署到本機 IIS Web 伺服器 (http://localhost)。 
    
接下來，Visual Studio 會執行下列作業︰


1. 修改 XML 資訊清單檔案的 [SourceLocation](http://msdn.microsoft.com/en-us/library/e6ea8cd4-7c8b-1da7-d8f8-8d3c80a088bc%28Office.15%29.aspx) 元素，方法是將 ~remoteAppUrl 權杖取代為啟動頁的完整位置 (例如，http://localhost/MyAgave.html)。
    
2. 在 IIS Express 中啟動 Web 應用程式專案。
    
3. 隨即開啟主應用程式。 
    
建置專案時，Visual Studio 不會在 [OUTPUT]**** 視窗中顯示驗證錯誤。 Visual Studio 會在發生時於 [ERRORLIST]**** 視窗中報告錯誤和警告。 Visual Studio 也會透過在程式碼和文字編輯器中顯示不同顏色的波浪底線 (也稱為不規則曲線)，以報告驗證錯誤。 這些標記會通知您 Visual Studio 在程式碼中偵測到的問題。 如需詳細資訊，請參閱[程式碼和文字編輯器](http://go.microsoft.com/fwlink/?LinkID=128497)。 如需如何啟用或停用驗證的相關資訊，請參閱： 


- [選項、文字編輯器、JavaScript、IntelliSense](http://go.microsoft.com/fwlink/?LinkID=238779)
    
- [作法：設定在 Visual Web Developer 中進行 HTML 編輯的驗證選項](http://msdn.microsoft.com/en-us/library/vstudio/0byxkfet%28v=vs.100%29.aspx)
    
- [CSS，請參閱選項對話方塊 | 文字編輯器 | CSS | 驗證](http://go.microsoft.com/fwlink/?LinkID=238780)
    
若要檢閱專案中 XML 資訊清單檔案的驗證規則，請參閱 [Office 增益集的 XML 資訊清單](../../docs/overview/add-in-manifests.md)。


### 在 Excel、Word 或 Project 中顯示增益集，並逐步執行程式碼


如果您將增益集專案的 [起始文件]**** 屬性設定為 Excel 或 Word，Visual Studio 會建立新的文件，而增益集會出現。 如果您將增益集專案的 [起始文件]**** 屬性設定為使用現有的文件，Visual Studio 會開啟文件，但是您必須手動插入增益集。 如果您將 [起始文件]**** 設定為 **Microsoft Project**，您也必須手動插入增益集。


### 在 Excel 或 Word 中顯示 Office 增益集


1. 在 Excel 或 Word 中，於 [插入]**** 索引標籤上選擇 [Office 增益集]****。
    
2. 在出現的清單中，選擇您的增益集。
    

### 在 Project 中顯示 Office 增益集


1. 在 Project 中，於 [專案]**** 索引標籤上選擇 [Office 增益集]****。
    
2. 在出現的清單中，選擇您的增益集。
    
然後您可以在 Visual Studio 設定中斷點。然後，在與增益集互動時，可以逐步執行 HTML、JavaScript 和 C# 或 VB 程式碼檔案中的程式碼。


### 在 Outlook 中顯示 Outlook 增益集並逐步執行程式碼


若要在 Outlook 中檢視增益集，請開啟電子郵件訊息或約會項目。

只要啟動條件符合，Outlook 便會啟動項目的增益集。增益集列會顯示在檢查程式視窗或閱讀窗格的頂端，而您的 Outlook 增益集會在增益集列中以按鈕形式出現。如果增益集有增益集命令，則會在功能區出現按鈕，可能是在預設索引標籤，或是指定的自訂索引標籤中，且增益集不會出現在增益集列中。

若要檢視 Outlook 增益集，請選擇 Outlook 增益集的按鈕。

在 Visual Studio 中，您可以設定中斷點。然後，在與 Outlook 增益集互動時，可以逐步執行 HTML、JavaScript 和 C# 或 VB 程式碼檔案中的程式碼。 

您也可以變更您的程式碼並且在 Outlook 增益集檢閱這些變更的效果，而不需關閉 Office 增益集，然後再次啟動專案。 在 Outlook 中，只需開啟 Outlook 增益集的快顯功能表，然後選擇 [重新載入]****。


### 修改程式碼並繼續偵錯增益集，而不需重新啟動專案


您可以變更您的程式碼並且在增益集檢閱這些變更的效果，而不需關閉主應用程式，然後再次啟動專案。 變更您的程式碼之後，請開啟增益集的快顯功能表，然後選擇 [重新載入]****。 當您重新載入增益集，它就會變成與 Visual Studio 偵錯工具中斷連線。 因此，您可以檢視您的變更的效果，但在您將 Visual Studio 偵錯工具附加至所有可用的 Iexplore.exe 處理序之前，會無法逐步執行程式碼。


### 將 Visual Studio 偵錯工具附加至所有可用的 Iexplore.exe 處理序


1. 在 Visual Studio 中，選擇 [偵錯]**** > [附加至處理序]****。
    
2. 在 [附加至處理序]**** 對話方塊中，選擇所有可用的 **iexplore.exe** 處理序，然後選擇 [附加]**** 按鈕。
    

## 後續步驟

- [發佈 Office 增益集](../publish/publish.md)
    
