

# <a name="projectdocument.viewselectionchanged-event"></a>ProjectDocument.ViewSelectionChanged 事件
使用中專案的使用中檢視變更時，就會發生。

|||
|:-----|:-----|
|**主應用程式︰**|Project|
|**可用於[需求集合](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Selection|
|**已新增於**|1.0|

```js
Office.EventType.ViewSelectionChanged
```


## <a name="remarks"></a>備註

 **ViewSelectionChanged** 為 [EventType](../../reference/shared/eventtype-enumeration.md) 列舉常數，可用於 [ProjectDocument.addHandlerAsync](../../reference/shared/projectdocument.addhandlerasync.md) 及 [ProjectDocument.removeHandlerAsync](../../reference/shared/projectdocument.removehandlerasync.md) 方法中，以新增或移除事件處理常式。


## <a name="example"></a>範例

下列程式碼範例會加入 **ViewSelectionChanged** 事件的處理常式。使用中檢視變更時，該處理常式會取得使用中檢視的名稱與類型。

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

            // After the DOM is loaded, add-in-specific code can run.
            Office.context.document.addHandlerAsync(
                Office.EventType.ViewSelectionChanged,
                getActiveView);
            getActiveView();
        });
    };

    // Get the name and type of the active view and display it in the add-in.
    function getActiveView() {
        Office.context.document.getSelectedViewAsync(
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    onError(result.error);
                }
                else {
                    var output = String.format(
                        'View name: {0}<br/>View type: {1}',
                        result.value.viewName, result.value.viewType);
                    $('#message').html(output);
                }
            }
        );
    }

    function onError(error) {
        $('#message').html(error.name + ' ' + error.code + ': ' + error.message);
    }
})();

```

如需範例示範如何使用 Project 增益集中的 **ViewSelectionChanged** 事件處理常式，請參閱[使用文字編輯器，建立您的第一個 Project 2013 工作窗格增益集](../../docs/project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md)。


## <a name="support-details"></a>支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此方法。空白儲存格表示 Office 主應用程式不支援此事件。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。



||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|
|:-----|:-----|:-----|
|**Project**|Y||

|||
|:-----|:-----|
|**可用於需求集合**||
|**最低權限等級**|[限制](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**增益集類型**|Task pane|
|**文件庫**|Office.js|
|**命名空間**|Office|

## <a name="support-history"></a>支援歷程記錄



|**版本**|**變更**|
|:-----|:-----|
|1.0|已導入|

## <a name="see-also"></a>請參閱



#### <a name="other-resources"></a>其他資源


[使用文字編輯器來建立第一個 Project 2013 的工作窗格增益集](../../docs/project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md)
[EventType 列舉](../../reference/shared/eventtype-enumeration.md)
[ProjectDocument.addHandlerAsync 方法](../../reference/shared/projectdocument.addhandlerasync.md)
[ProjectDocument.removeHandlerAsync 方法](../../reference/shared/projectdocument.removehandlerasync.md)
[ProjectDocument 物件](../../reference/shared/projectdocument.projectdocument.md)
