
# ProjectDocument.getMaxTaskIndexAsync 方法
以非同步方式取得目前專案中工作集合的最大索引。

 **重要事項：**此 API 只適用於 Windows 桌面上的 Project 2016。

|||
|:-----|:-----|
|**主應用程式︰**|Project|
|**可用於[需求集合](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Selection|
|**已新增於**|1.1|

```js
Office.context.document.getMaxTaskIndexAsync([options][, callback]);
```


## 參數

_選項_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;下列**[選擇性參數](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)：**<br/><br/>&nbsp;&nbsp;&nbsp;&nbsp;_asyncContext_<br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;類型：**陣列**、**布林值**、**null**、**數字**、**物件**、**字串** 或 **未定義**<br/><br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;無變更的情況下，於 [AsyncResult](../../reference/shared/asyncresult.md) 物件中傳回的任一類型使用者定義項目。 選用。<br/><br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;例如，您可以傳遞 _asyncContext_ 引數，方法是使用格式 `{asyncContext: 'Some text'}` 或 `{asyncContext: <object>}`。

_回呼_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;類型：**函數**<br/>
&nbsp;&nbsp;&nbsp;&nbsp;方法呼叫傳回時所叫用的函數，其唯一的參數為 [AsyncResult](../../reference/shared/asyncresult.md) 類型。 選用。   

## 回呼值

當 _callback_ 函數執行時，該函數會收到 [AsyncResult](../../reference/shared/asyncresult.md) 物件，您可以從回呼函數的參數存取該物件。

若為 **getMaxTaskIndexAsync** 方法，傳回的 [AsyncResult](../../reference/shared/asyncresult.md) 物件會包含下列屬性：


|**名稱**|**說明**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|在選擇性 _asyncContext_ 參數中傳遞的資料 (如果有使用該參數)。|
|[錯誤](../../reference/shared/asyncresult.error.md)|錯誤的相關資訊 (如果 **status** 屬性等於 **failed**)。|
|[狀態](../../reference/shared/asyncresult.status.md)|非同步呼叫的 **succeeded** 或 **failed** 狀態。|
|[值](../../reference/shared/asyncresult.value.md)|在目前專案的工作集合中，最大的索引編號。|

## 備註

您可以使用傳回的值搭配 [getTaskByIndexAsync](../../reference/shared/projectdocument.gettaskbyindexasync.md) 方法，以取得工作 GUID。 0 索引工作代表專案摘要工作。


## 範例

下列程式碼範例會呼叫 **getMaxTaskIndexAsync**，以取得目前專案中之工作集合的最大索引。然後它會使用傳回的值搭配 [getTaskByIndexAsync](../../reference/shared/projectdocument.getselectedtaskasync.md) 方法，以取得每個工作 GUID。

此範例假設您的增益集參照 jQuery 程式庫，且已在頁面內文的內容 div 中定義下列頁面控制項。




```HTML
<input id="get-info" type="button" value="Get info" /><br />
<span id="message"></span>
```




```js
(function () {
    "use strict";
    var taskGuids = [];

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            // After the DOM is loaded, add-in-specific code can run.
            app.initialize();
            $('#get-info').click(getTaskInfo);
        });
    };

    // Get the maximum task index, and then get the task GUIDs.
    function getTaskInfo() {
        getMaxTaskIndex().then(
            function (data) {
                getTaskGuids(data);
            }
        );
    }

    // Get the maximum index of the tasks for the current project.
    function getMaxTaskIndex() {
        var defer = $.Deferred();
        Office.context.document.getMaxTaskIndexAsync(
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    onError(result.error);
                }
                else {
                    defer.resolve(result.value);
                }
            }
        );
        return defer.promise();
    }

    // Get each task GUID, and then display the GUIDs in the add-in.
    function getTaskGuids(maxTaskIndex) {
        var defer = $.Deferred();
        for (var i = 0; i <= maxTaskIndex; i++) {
            getTaskGuid(i);
        }
        return defer.promise();
        function getTaskGuid(index) {
            Office.context.document.getTaskByIndexAsync(index,
                function (result) {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        taskGuids.push(result.value);
                        if (index == maxTaskIndex) {
                            defer.resolve();
                            $('#message').html(taskGuids.toString());
                        }
                    }
                    else {
                        onError(result.error);
                    }
                }
            );
        }
    }
    function onError(error) {
        app.showNotification(error.name + ' ' + error.code + ': ' + error.message);
    }
})();
```


## 支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此方法。空白儲存格表示 Office 主應用程式不支援此方法。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|
|:-----|:-----|:-----|
|**Project**|Y||

|||
|:-----|:-----|
|**可用於需求集合**||
|**最低權限等級**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**增益集類型**|Task pane|
|**文件庫**|Office.js|
|**命名空間**|Office|

## 支援歷程記錄




|**版本**|**變更**|
|:-----|:-----|
|1.1|已導入|

## 請參閱



#### 其他資源


[getTaskByIndexAsync](../../reference/shared/projectdocument.gettaskbyindexasync.md)

[AsyncResult 物件](../../reference/shared/asyncresult.md)

[ProjectDocument 物件](../../reference/shared/projectdocument.projectdocument.md)
