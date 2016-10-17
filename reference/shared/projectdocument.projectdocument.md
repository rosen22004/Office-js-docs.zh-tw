

# <a name="projectdocument-object"></a>ProjectDocument 物件
抽象類別，代表與 Office 增益集互動的專案文件 (使用中的專案)。

|||
|:-----|:-----|
|**主應用程式︰**|Project|
|**已新增於**|1.0|

```js
Office.context.document
```


## <a name="members"></a>成員


**方法**


|**名稱**|**描述**|
|:-----|:-----|
|[addHandlerAsync 方法](../../reference/shared/projectdocument.addhandlerasync.md)|針對 **ProjectDocument** 物件中的事件，以非同步方式加入事件處理常式。|
|[getMaxResourceIndexAsync 方法](../../reference/shared/projectdocument.getmaxresourceindexasync.md)|以非同步方式取得目前專案中資源集合的最大索引。|
|[getMaxTaskIndexAsync 方法](../../reference/shared/projectdocument.getmaxtaskindexasync.md)|以非同步方式取得目前專案中工作集合的最大索引。|
|[getProjectFieldAsync 方法](../../reference/shared/projectdocument.getprojectfieldasync.md)|以非同步方式取得使用中專案指定欄位的值。|
|[getResourceByIndexAsync 方法](../../reference/shared/projectdocument.getresourcebyindexasync.md)|以非同步方式取得資源集合中具有指定索引的資源 GUID。|
|[getResourceFieldAsync 方法](../../reference/shared/projectdocument.getresourcefieldasync.md)|以非同步方式取得指定資源的指定欄位值。|
|[getSelectedDataAsync 方法](../../reference/shared/projectdocument.getselecteddataasync.md)|以非同步方式取得甘特圖中目前的選取範圍內，一或多個儲存格中所包含的資料。|
|[getSelectedResourceAsync 方法](../../reference/shared/projectdocument.getselectedresourceasync.md)|以非同步方式取得所選取資源的 GUID。|
|[getSelectedTaskAsync 方法](../../reference/shared/projectdocument.getselectedtaskasync.md)|以非同步方式取得所選取工作的 GUID。|
|[getSelectedViewAsync 方法](../../reference/shared/projectdocument.getselectedviewasync.md)|以非同步方式取得使用中檢視的名稱與檢視類型。|
|[getTaskAsync 方法](../../reference/shared/projectdocument.gettaskasync.md)|以非同步方式取得工作名稱、指派給該工作的資源，和同步的 SharePoint 工作清單中工作的識別碼。|
|[getTaskByIndexAsync 方法](../../reference/shared/projectdocument.gettaskbyindexasync.md)|以非同步方式取得工作集合中具有指定索引的工作 GUID。|
|[getTaskFieldAsync 方法](../../reference/shared/projectdocument.gettaskfieldasync.md)|以非同步方式取得指定工作的指定欄位值。|
|[getWSSUrlAsync 方法](../../reference/shared/projectdocument.getwssurlasync.md)|以非同步方式取得同步化 SharePoint 工作清單的 URL。|
|[removeHandlerAsync 方法](../../reference/shared/projectdocument.removehandlerasync.md)|針對 **ProjectDocument** 物件中的事件，以非同步方式移除事件處理常式。|
|[setResourceFieldAsync 方法](../../reference/shared/projectdocument.setresourcefieldasync.md)|以非同步方式設定指定資源的指定欄位值。|
|[setTaskFieldAsync 方法](../../reference/shared/projectdocument.settaskfieldasync.md)|以非同步方式設定指定工作的指定欄位值。|

**事件**


|**名稱**|**描述**|
|:-----|:-----|
|[ResourceSelectionChanged 事件](../../reference/shared/projectdocument.resourceselectionchanged.event.md)|使用中專案的資源選取項目變更時，就會發生。|
|[TaskSelectionChanged 事件](../../reference/shared/projectdocument.taskselectionchanged.event.md)|使用中專案的工作選取項目變更時，就會發生。|
|[ViewSelectionChanged 事件](../../reference/shared/projectdocument.viewselectionchanged.event.md)|使用中專案的使用中檢視變更時，就會發生。|

## <a name="remarks"></a>備註

請勿直接呼叫或具現化指令碼中的 **ProjectDocument** 物件。


## <a name="example"></a>範例

下列範例會初始化增益集，然後取得可從 Project 文件內容取得的 [Document](../../reference/shared/document.md) 物件屬性。Project 文件是開啟的使用中專案。若要存取 **ProjectDocument** 物件的成員，請使用 **Office.context.document** 物件，如 **ProjectDocument** 方法和事件的程式碼範例所示。

此範例假設您的增益集參照 jQuery 程式庫，且已在頁面內文的內容 div 中定義下列頁面控制項。




```HTML
<span id="message"></span>
```




```js
(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            // Get information about the document.
            showDocumentProperties();
        });
    };

    // Get the document mode and the URL of the active project.
    function showDocumentProperties() {
        var output = String.format(
            'The document mode is {0}.<br/>The URL of the active project is {1}.',
            Office.context.document.mode,
            Office.context.document.url);
        $('#message').html(output);
    }
})();
```


## <a name="support-details"></a>支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此物件。空白儲存格表示 Office 主應用程式不支援此物件。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|
|:-----|:-----|:-----|
|**Project**|Y||

|||
|:-----|:-----|
|**增益集類型**|Task pane|
|**文件庫**|Office.js|
|**命名空間**|Office|

## <a name="support-history"></a>支援歷程記錄


|**版本**|**變更**|
|:-----|:-----|
|1.0|已導入|

## <a name="see-also"></a>請參閱



#### <a name="other-resources"></a>其他資源


[Project 的工作窗格增益集](../../docs/project/project-add-ins.md)
[Document 物件](../../reference/shared/document.md)

