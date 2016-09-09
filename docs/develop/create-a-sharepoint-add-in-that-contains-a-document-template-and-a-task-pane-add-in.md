
# 建立包含文件範本和工作窗格增益集的 SharePoint 增益集


您可以建立包含文件範本 (例如費用報表) 的 SharePoint 增益集。文件可以包含與 SharePoint 資料互動的工作窗格增益集。例如，使用者可以藉由使用來自 Business Connectivity Services (BCS) 的資料來填入發票的欄位，或從 SharePoint 清單中選取費用類別來建立費用報表。

這個逐步解說會示範如何建立包含 Excel 活頁簿的 SharePoint 增益集。Excel 活頁簿包含的工作窗格增益集會使用 SharePoint 2013 提供的 REST 介面來填入下拉式清單方塊，並且在工作窗格增益集中會有 SharePoint 日期。


## 必要條件


開始前請安裝下列元件：




- SharePoint 開發環境︰
    
      - To develop SharePoint Add-ins that target SharePoint in Office 365, see [How to: Set up an environment for developing SharePoint Add-ins on Office 365](http://msdn.microsoft.com/en-us/library/office/apps/fp161179%28v=office.15%29).
    
  - 若要開發以 SharePoint 的內部部署安裝為目標的 SharePoint 增益集，請參閱[作法：設定 SharePoint 增益集的內部部署開發環境](http://msdn.microsoft.com/en-us/library/office/apps/fp179923%28v=office.15%29)。
    
- [Visual Studio 2015 和 Microsoft Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs)
    
- Excel 2013 或 Office 365 帳戶。
    

## 在 Visual Studio 中建立 SharePoint 增益集專案



1. 啟動 Visual Studio。
    
2. 在功能表列中，選擇 [檔案]**** > [新增]**** > [專案]****。
    
    [新增專案]**** 對話方塊隨即開啟。
    
3. 在範本窗格中，在您想要使用的語言節點下，展開 **Office/SharePoint**，然後選擇 [Office 增益集]****。
    
4. 在專案類型清單中，選擇 [SharePoint 增益集]****，將專案命名為 OfficeEnabledAddin，然後選擇 [確定]**** 按鈕。
    
    [新增 SharePoint 增益集]**** 對話方塊隨即出現。
    
5. 在 [您要使用哪個 SharePoint 網站來偵錯增益集?]**** 下拉式清單中，選擇或輸入 SharePoint 網站的 URL。
    
6. 在 [您要如何裝載您的 SharePoint 增益集?]**** 下拉式清單中，選擇 [由 SharePoint 裝載]****，然後選擇 [下一步]****。
    
     >**附註：**此狀況僅於 SharePoint 裝載及提供者裝載選項出現於 [您要如何裝載您的 SharePoint 增益集?]**** 下拉式清單中方可適用。
7. 在下一個頁面上，選取 **SharePoint 2013**，然後選擇 [完成]**** 按鈕以關閉對話方塊。
    

## 加入工作窗格增益集項目


接下來，將 Office 增益集加入專案。您可以加入您想要的任何類型的增益集。在這個逐步解說中，我們會加入工作窗格增益集。


1. 在**方案總管**中，選擇 **OfficeEnabledAddin** 專案節點。
    
2. 在 [專案]**** 功能表中，選擇 [加入新項目]****。
    
3. 在 [加入新項目]**** 對話方塊中，選取 **Office/SharePoint**，然後選擇 [Office 增益集]****。
    
4. 將工作窗格增益集命名為 MyTaskPaneAddin，然後選擇 [加入]**** 按鈕。
    
    [建立 Office 相關增益集]**** 對話方塊隨即開啟。
    
5. 在 [建立 Office 相關增益集]**** 對話方塊中，選取 [工作窗格]****，然後選擇 [下一步]****。 在下一個頁面上，清除 **Word** 和 **PowerPoint** 核取方塊，然後選擇 [下一步]****。
    
6. 在 [是否要讓您的 Office 增益集可以出現在新文件或現有文件中?]**** 頁面中，選擇 [建立新文件並插入我的增益集]****，然後選擇 [完成]****。
    
    Visual Studio 新增了文件庫，以及為文件庫新增了活頁簿範本。 活頁簿包含了工作窗格增益集。
    

## 新增文件庫


在此程序中，您將加入文件庫，並讓活頁簿成為文件庫的預設範本。


1. 在**方案總管**中，選擇 **OfficeEnabledAddin** 專案節點。
    
2. 在 [專案]**** 功能表中，選擇 [加入新項目]****。
    
3. 在 [加入新項目]**** 對話方塊中，選取 **Office/SharePoint**，然後選擇 [清單]****，將清單命名為 MyDocumentLibrary，然後選擇 [加入]****按鈕。
    
4. 在 **SharePoint 自訂精靈**，請選取 [建立自訂的清單範本和它的清單執行個體]**** 選項。
    
5. 在這個選項下方的下拉式清單中，選取 [文件庫]****，然後選擇 [下一步]**** 按鈕。
    
6. 在 [請為這個文件庫選擇範本。使用者在這個文件庫建立的文件，將會以該範本為基礎]**** 頁面上，選擇 [將下列文件作為此文件庫的範本使用]****，然後選擇 [瀏覽]**** 按鈕。
    
7. 在 [開啟]**** 對話方塊中，開啟 **OfficeDocuments** 資料夾，選取 **MyTaskPaneApp.xlsx** 檔案，選擇 [開啟]**** 按鈕，選擇 [完成]**** 按鈕，然後關閉清單設計工具。
    
8. 在**方案總管**中，選擇 **OfficeEnabledAddin** 專案節點。
    
9. 在 [檢視]**** 功能表中，選擇 [屬性視窗]****。
    
10. 在**方案總管**中，選擇 **AppManifest.xml** 檔案。
    
11. 選擇 [檢視]**** 及 [設計工具]****。
    
12. 在資訊清單設計工具中，將 [起始頁面]**** 的值設定為 ~appWebUrl/Lists/MyDocumentLibrary。 這將轉換成 OfficeEnabledAddin/Lists/MyDocumentLibrary 值。
    
     >**附註：**此 URL 指的是文件庫。 針對在增益集網站內指向項目的 Office 增益集資訊清單，您必須在任何 URL 的開頭使用 ~appWebUrl Token。 如需在 SharePoint 增益集專案中的 URL Token 相關資訊，請參閱 [SharePoint 增益集中的 URL 字串和 Token](http://msdn.microsoft.com/library/800ec8cd-a448-46bc-b41e-d4030eeb4048%28Office.15%29.aspx)。
13. 關閉資訊清單設計工具，以儲存變更。
    

## 在工作窗格中取用 SharePoint 資料


在此程序中，您會透過使用 SharePoint 2013 提供的 REST 介面來顯示站台使用者的清單。

在這個範例中，只顯示了 SharePoint 清單資料，但您可能會使用這類資料做為文件核准增益集的一部分。當使用者從該清單選擇名稱時，您的程式碼會在文件追蹤清單中設定檢閱者欄的值。與該清單相關聯的工作流程可以傳送檢閱通知給該使用者。或者，您可能會將所選取的名稱儲存至文件設定。然後，當使用者開啟文件時，只有當目前的使用者與儲存至文件設定中的使用者相同時，您才能在工作窗格增益集中顯示控制項。如需詳細資訊，請參閱下列主題：


- [使用 SharePoint 2013 REST 端點完成基本作業](http://msdn.microsoft.com/library/e3000415-50a0-426e-b304-b7de18f2f7d9%28Office.15%29.aspx)
    
- [使用 SharePoint 2013 中的 JavaScript 程式庫程式碼完成基本作業](http://msdn.microsoft.com/library/29089af8-dbc0-49b7-a1a0-9e311f49c826%28Office.15%29.aspx)
    
- [保存增益集狀態和設定](../../docs/develop/persisting-add-in-state-and-settings.md)
    

1. 在**方案總管**中，展開 **MyTaskPaneAddin** 資料夾，接著展開 **Home** 資料夾，然後選擇 **Home.html** 檔案。
    
    在程式碼編輯器中開啟 Home.html 檔案。
    
2. 在 `get-data-from-selection` 按鈕下方加入下列 HTML。
    
```HTML
  <p>Select Reviewer:</p> <select class="select" id="select-reviewer" name="D1"> </select>
```

3. 選擇 **Home.js** 檔案以在程式碼編輯器中開啟 Home.js 檔案。
    
4. 將下列宣告加入 Home.js 檔案頂端。
    
```js
  var appWebURL; var web;
```

5. 以下列程式碼取代 `Initialize` 函式。
    
    此程式碼會執行下列工作：
    
      - 在 jQuery 函式中使用 `getScript` 來載入 SP.Runtime.js 和 SP.js 檔案。 在載入檔案之後，您的程式會有權存取 SharePoint 的 JavaScript 物件模型。
    
  - 載入目前的網站物件。
    
  - 呼叫會取得站台的所有使用者的函式。您將會在下一個步驟中加入該函式的程式碼。
    



```js
   // The initialize function must be run each time a new page is loaded Office.initialize = function (reason) { $(document).ready(function () { app.initialize(); var scriptbase = "/_layouts/15/"; $.getScript(scriptbase + "SP.Runtime.js", function () { $.getScript(scriptbase + "SP.js", function () { getAppWeb(function () { getSPUsers(populateUsersDropDown); }); }); }); function getAppWeb(functionToExecuteOnReady) { var context = SP.ClientContext.get_current(); web = context.get_web(); context.load(web); context.executeQueryAsync(onSuccess, onFailure); function onSuccess() { appWebURL = web.get_url(); functionToExecuteOnReady(); } function onFailure(sender, args) { app.initialize(); app.showNotification("Failed to connect to SharePoint. Error: " + args.get_message()); } } $('#get-data-from-selection').click(getDataFromSelection); }); };
```

6. 將下列程式碼加入 Home.js 檔案底端。
    
    此程式碼使用了 SharePoint 2013 提供的 REST 介面來取得網站使用者的清單。 接著，此程式碼在下拉式清單中填入了每位使用者的名稱及 ID。
    


```js
  function getSPUsers(functionToExecuteOnReady) { var url = appWebURL + "/../_api/web/siteUsers"; jQuery.ajax({ url: url, type: "GET", headers: { "ACCEPT": "application/json;odata=verbose" }, success: onSuccess, error: onFailure }); function onSuccess(data) { var results = data.d.results; functionToExecuteOnReady(results); } function onFailure(jaXHR, textStatus, errorThrown) { var error = textStatus + " " + errorThrown; app.showNotification(error); } } function populateUsersDropDown(results) { for (var i = 0; i < results.length; i++) { var IDTemp = results[i].Id; $('#select-reviewer').append("<option value='" + IDTemp + "'>" + results[i].Title + "</option>"); } }
```

7. 在**方案總管**中，開啟 **AppManifest.xml** 檔案的快顯功能表，然後選擇 [檢視表設計工具]****。
    
8. 在設計工具中，選擇 [權限]**** 頁面。
    
9. 從下拉式清單中，於 [範圍]**** 欄選擇 **Web** 項目。
    
10. 從下拉式清單中，於 [權限]**** 欄選擇 [讀取]**** 項目。
    

## 偵錯工作窗格增益集


您可以藉由啟動文件，或藉由啟動 SharePoint 增益集然後從文件庫開啟文件來偵錯工作窗格增益集。


### 藉由啟動文件來偵錯工作窗格增益集




 >**附註：**因為在此程序會開啟 Excel，只有當系統上安裝了 Office 時才能運作。 否則，就會發生錯誤「這部電腦上未安裝與這種專案類型關聯的應用程式」。


1. 在程式碼編輯器中開啟 Home.js 檔案，然後在 `getDataFromSelection` 方法旁邊設定中斷點。
    
2. 在**方案總管**中，選擇 **OfficeEnabledApp** 專案節點。
    
3. 在 [檢視]**** 功能表中，選擇 [屬性視窗]****。
    
4. 在 [屬性] 視窗中，從 [起始動作]**** 下拉式清單中，選擇 [Office 桌面用戶端]**** 項目。 當您這麼做時，便會出現新屬性 [起始文件]****。
    
5. 從 [起始文件]**** 下拉式清單中，選擇 **OfficeDocuments\TaskPaneApp.xlsx** 項目。
    
6. 在 [偵錯]**** 功能表中，選擇 [開始偵錯]****。
    
    在增益集執行時，此設定讓工作窗格增益集中的活頁簿顯示。 活頁簿會開啟且工作窗格增益集會顯示。
    
7. 在工作窗格增益集中，選擇 [選取檢閱者]**** 下拉式清單以檢視 SharePoint 使用者的清單。
    
8. 在 Excel 活頁簿中，選取任何儲存格。
    
9. 在工作窗格增益集中，選擇 [從選取範圍取得資料]**** 按鈕。
    
    執行會在您於 `getDataFromSelection` 方法旁邊所設定的中斷點停止。
    

### 藉由啟動 SharePoint 來偵錯工作窗格增益集




 >**附註：**此程序會開啟 Excel Online。 只有當您有 Office 365 帳戶時才有作用。 請參閱[作法：設定用於開發 Office 365 上 SharePoint 增益集的環境](http://msdn.microsoft.com/en-us/library/office/apps/fp161179%28v=office.15%29)。


1. 在程式碼編輯器中開啟 Home.js 檔案，然後在 `getDataFromSelection` 方法旁邊設定中斷點。
    
2. 在**方案總管**中，選擇 **OfficeEnabledApp** 專案節點。
    
3. 在 [檢視]**** 功能表中，選擇 [屬性視窗]****。
    
4. 在 [屬性] 視窗中，從 [起始動作]**** 下拉式清單中，選擇 **Internet Explorer** 項目。
    
5. 在 [偵錯]**** 功能表中，**開始偵錯**。
    
    Visual Studio 會開啟 SharePoint 並顯示 [MyDocumentLibrary]**** 程式庫。
    
6. 在 SharePoint 中，於 [檔案]**** 索引標籤上，選擇 [新增文件]****。 
    
7. 瀏覽至您的專案中的活頁簿 MyTaskPaneApp.xlsx。
    
    活頁簿會開啟且工作窗格增益集會顯示。
    
8. 請確定您的瀏覽器中已啟用指令碼偵錯。 在 Internet Explorer 中，您可以透過開啟 [網際網路選項]**** 對話方塊，選擇 [進階]****索引標籤，然後清除 [停用指令碼偵錯 (Internet Explorer)]**** 和 [停用指令碼偵錯 (其他)]**** 核取方塊來啟用指令碼偵錯。
    
9. 在 Visual Studio 中，於 [偵錯]**** 功能表中，選擇 [附加至處理序]****。
    
10. 在 [附加至處理序]**** 對話方塊中，選擇所有可用的 **iexplore.exe** 處理序，然後選擇 [附加]**** 按鈕。
    
11. 在工作窗格增益集中，選擇 [選取檢閱者]**** 下拉式清單以檢視 SharePoint 使用者的清單。
    
    清單中資料是使用 REST 呼叫在 SharePoint 中擷取而得。
    
12. 在 Excel 活頁簿中，選擇任何儲存格。
    
13. 在工作窗格增益集中，選擇 [從選取範圍取得資料]**** 按鈕。
    
    執行會在您於 `getDataFromSelection` 方法旁邊所設定的中斷點停止。
    
     >**附註：**若活頁簿並未包含任何資料，您可以新增部分項目，方式是在活頁簿內的工具列上，依序選擇 [編輯活頁簿]**** > [於 Excel Online 中編輯]****。

## 封裝和發佈增益集


當您準備好要封裝您的增益集供發佈時，開啟 [發佈 Office 和 SharePoint 增益集]**** 精靈。


- 在**方案總管**中，開啟 SharePoint 增益集專案的快顯功能表，然後選擇 [發佈]****。
    
    便會顯示 [發佈 Office 和 SharePoint 增益集]**** 精靈。 如需詳細資訊，請參閱[使用 Visual Studio 發佈 SharePoint 增益集](http://msdn.microsoft.com/library/8137d0fa-52e2-4771-8639-60af80f693bb%28Office.15%29.aspx)。
    

## 其他資源


- [Office 增益集的設計指導方針](../../docs/design/add-in-design.md)
    
- [Office 增益集開發生命週期](../../docs/design/add-in-development-lifecycle.md)
    
- [發佈 Office 增益集](../publish/publish.md)
    
- [了解適用於 Office 的 JavaScript API](../../docs/develop/understanding-the-javascript-api-for-office.md)
    
- [Office 增益集的 XML 資訊清單](../../docs/overview/add-in-manifests.md)
    
- [Office 增益集 API 和結構描述參考](../../reference/reference.md)
    
