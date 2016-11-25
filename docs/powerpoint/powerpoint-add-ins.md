# <a name="powerpoint-add-ins"></a>PowerPoint 增益集

您可以使用 PowerPoint 增益集，針對各平台 (包括 Windows、iOS、Office Online 和 Mac ) 上使用者的簡報建置吸引人的解決方案。您可以建立的增益集類型有兩種︰

- 使用**內容增益集**，將動態 HTML5 內容新增到您的簡報。例如，請參閱[適用於 PowerPoint 的 LucidChart 圖表](https://store.office.com/en-us/app.aspx?assetid=WA104380117&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Productivity&homapppos=3&homchv=2&appredirect=false)增益集，您可用於將 LucidChart 中的互動式圖表插入您的投影片組中。
- 使用**工作窗格增益集**帶入參考資訊，或透過服務將資料插入投影片中。例如，請參閱 [Shutterstock 影像](https://store.office.com/en-us/app.aspx?assetid=WA104380169&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Editor%2527s%2BPicks&homapppos=0&homchv=1&appredirect=false)增益集，您可用於將專業相片新增至您的簡報。 


## <a name="powerpoint-add-in-scenarios"></a>PowerPoint 增益集案例

本文中的程式碼範例說明用於開發 PowerPoint 內容增益集的一些基本工作。 

為了顯示資訊，這些範例取決於 Visual StudioOffice 增益集專案範本包含的 `app.showNotification` 函式。如果您不使用 Visual Studio 開發增益集，您將需要使用自己的程式碼取代 `showNotification` 函式。其中幾個範例也取決於這個 `globals` 物件 (在這些函式的範圍外部宣告)︰`var globals = {activeViewHandler:0, firstSlideId:0};`

這些程式碼範例需要您的專案[參考 Office.js 1.1 版程式庫或更新版本](../../docs/develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md)。


## <a name="detect-the-presentations-active-view-and-handle-the-activeviewchanged-event"></a>偵測簡報的使用中檢視並處理 ActiveViewChanged 事件

如果您要建置內容增益集，您必須取得簡報的使用中檢視並處理 ActiveViewChanged 事件，做為 Office.Initialize 處理常式的一部分。


- `getActiveFileView` 函式呼叫 [Document.getActiveViewAsync](../../reference/shared/document.getactiveviewasync.md) 方法以傳回簡報的目前檢視為「編輯」(您可在其中編輯投影片的任何檢視，例如**標準**或**大綱模式**) 或「讀取」(**投影片放映**或**讀取檢視**) 檢視。


- `registerActiveViewChanged` 函式呼叫 [addHandlerAsync](../../reference/shared/document.addhandlerasync.md) 方法來註冊 [Document.ActiveViewChanged](../../reference/shared/document.activeviewchanged.md) 事件的處理常式。 
> 附註：在 PowerPoint Online 中，因為投影片放映模式被視為新的工作階段，所以永遠不會觸發 [Document.ActiveViewChanged](../../reference/shared/document.activeviewchanged.md) 事件。在此情況下，增益集必須在載入時擷取使用中檢視，如下所述。



```js

//general Office.initialize function. Fires on load of the add-in.
Office.initialize = function(){

    //Gets whether the current view is edit or read.
    var currentView = getActiveFileView();

    //register for the active view changed handler
    registerActiveViewChanged();

    //render the content based off of the currentView
    //....
}

function getActiveFileView()
{
    Office.context.document.getActiveViewAsync(function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification(asyncResult.value);
        }
    });

}


function registerActiveViewChanged() {
    Globals.activeViewHandler = function (args) {
        app.showNotification(JSON.stringify(args));
    }

    Office.context.document.addHandlerAsync(Office.EventType.ActiveViewChanged, Globals.activeViewHandler, 
        function (asyncResult) {
            if (asyncResult.status == "failed") {
           app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
            else {
            app.showNotification(asyncResult.status);
            }
        });
}
```
    

## <a name="navigate-to-a-particular-slide-in-the-presentation"></a>導覽至簡報中的特定投影片

`getSelectedRange` 函式呼叫 [Document.getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md) 方法來取得 `asyncResult.value` 傳回的 JSON 物件，其包含名為「投影片」的陣列，其中包含投影片選定範圍的識別碼、標題和索引 (或只是目前的投影片)。它也會將選取範圍中第一張投影片的識別碼儲存至全域變數。


```js
function getSelectedRange() {
    // Get the id, title, and index of the current slide (or selected slides) and store the first slide id */
    Globals.firstSlideId = 0;

    Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            Globals.firstSlideId = asyncResult.value.slides[0].id;
            app.showNotification(JSON.stringify(asyncResult.value));
        }
    });
}
```

`goToFirstSlide` 函式呼叫 [Document.goToByIdAsync](../../reference/shared/document.gotobyidasync.md) 方法，移至上述 `getSelectedRange` 函式儲存的第一張投影片的識別碼。




```js
function goToFirstSlide() {
    Office.context.document.goToByIdAsync(Globals.firstSlideId, Office.GoToType.Slide, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification("Navigation successful");
        }
    });
}
```


## <a name="navigate-between-slides-in-the-presentation"></a>在簡報中的投影片之間導覽

`goToSlideByIndex` 函式呼叫 **Document.goToByIdAsync** 方法以導覽至簡報中的下一張投影片。


```js
function goToSlideByIndex() {
    var goToFirst = Office.Index.First;
    var goToLast = Office.Index.Last;
    var goToPrevious = Office.Index.Previous;
    var goToNext = Office.Index.Next;

    Office.context.document.goToByIdAsync(goToNext, Office.GoToType.Index, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification("Navigation successful");
        }
    });
}
```

## <a name="get-the-url-of-the-presentation"></a>取得簡報的 URL

`getFileUrl` 函式呼叫 [Document.getFileProperties](../../reference/shared/document.getfilepropertiesasync.md) 方法來取得簡報檔案的 URL。


```js
function getFileUrl() {
    //Get the URL of the current file.
    Office.context.document.getFilePropertiesAsync(function (asyncResult) {
        var fileUrl = asyncResult.value.url;
        if (fileUrl == "") {
            app.showNotification("The file hasn't been saved yet. Save the file and try again");
        }
        else {
            app.showNotification(fileUrl);
        }
    });
}
```



## <a name="additional-resources"></a>其他資源
- [PowerPoint 程式碼範例](https://dev.office.com/code-samples#?filters=powerpoint)

- [如何依據內容和工作窗格增益集的文件來儲存增益集狀態和設定](../../docs/develop/persisting-add-in-state-and-settings.md#how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins)

- [在文件或試算表中的作用選取範圍內讀取和寫入資料](../../docs/develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)
    
- [從 PowerPoint 或 Word 增益集中，取得整份文件](../../docs/develop/get-the-whole-document-from-an-add-in-for-powerpoint-or-word.md)
    
- [使用您 PowerPoint 增益集的文件佈景主題](../powerpoint/use-document-themes-in-your-powerpoint-add-ins.md)
    