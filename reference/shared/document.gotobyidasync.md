
# <a name="document.gotobyidasync-method"></a>Document.goToByIdAsync 方法
移至文件中指定的物件或位置。

|||
|:-----|:-----|
|**主應用程式︰**|Excel、PowerPoint、Word|
|**可用於需求集合**|不在集合中|
|**已新增於**|1.1|

```js
Office.context.document.goToByIdAsync(id, goToType, [,options], callback);
```


## <a name="parameters"></a>參數



|**名稱**|**類型**|**描述**|**支援附註**|
|:-----|:-----|:-----|:-----|
| _id_|**字串**或**數字**|要前往之物件或位置的識別碼。必要。||
| _goToType_|[GoToType](../../reference/shared/gototype-enumeration.md)|要前往的位置類型。必要。||
| _options_|**object**|指定下列任何一項[選擇性參數](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)||
| _selectionMode_|[SelectionMode](../../reference/shared/selectionmode-enumeration.md)|指定位置是否由所選取 (反白顯示) 的 _id_ 參數指定。|**在 Excel 中：**<br/> **Office.SelectionMode.Selected** 會選取繫結中的所有內容，或具名項目。 <br/>**Office.SelectionMode.None** 對於文字繫結，會選取儲存格；對於矩陣細節、表格繫結和具名項目，會選取第一個資料儲存格 (非表格之標題列中的第一個儲存格)。<br/><br/> **在 PowerPoint 中：**<br/> **Office.SelectionMode.Selected** 會選取投影片上的投影片標題或第一個文字方塊。<br/> **Office.SelectionMode.None** 不會選取任何項目。<br/><br/> **在 Word 中：**<br/> **Office.SelectionMode.Selected** 會選取繫結中的所有內容。 <br/>**Office.SelectionMode.None** 對於文字繫結，會將游標移至文字開頭；對於矩陣繫結與表格繫結，會選取第一個資料儲存格 (非表格之標題列中的第一個儲存格)。|
| _asyncContext_|**陣列**、**布林值**、**null**、**數字**、**物件**、**字串**或**未定義**|無變更的情況下，於 **AsyncResult** 物件中傳回的任一類型使用者定義項目。||
| _callback_|**object**|回呼傳回時所叫用的函數，其唯一的參數為 **AsyncResult** 類型。||

## <a name="callback-value"></a>回呼值

傳遞至 _callback_ 參數的函數執行時，該函數會收到 [AsyncResult](../../reference/shared/asyncresult.md) 物件，您可以從回呼函數的唯一參數存取該物件。

在傳遞至 **goToByIdAsync** 方法的回呼函數中，您可以使用 **AsyncResult** 物件的屬性以傳回下列資訊。



|**屬性**|**用於...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|傳回目前的檢視。|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|判定作業成功或失敗。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|作業失敗時，存取提供錯誤資訊的 [Error](../../reference/shared/error.md) 物件。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|存取您的使用者定義**物件**或值 (如果您傳遞了其中一項做為 _asyncContext_ 參數)。|

## <a name="remarks"></a>備註

PowerPoint 不支援 **Master Views** 中的 **goToByIdAsync** 方法。


## <a name="example"></a>範例

 **根據 ID 移至繫結 (Word 和 Excel)**

下列範例顯示如何：


-  使用 **addFromSelectionAsync** 方法做為要使用則範例繫結，以[建立表格繫結](../../reference/shared/bindings.addfromselectionasync.md)。
    
-  **指定該繫結**做為要前往的繫結。
    
-  **傳遞匿名的回呼函數**，可將作業狀態傳回至 _goToByIdAsync_ 方法的 **callback** 參數。
    
-  在增益集頁面上**顯示值**。
    



```js
function gotoBinding() {
    //Create a new table binding for the selected table.
    Office.context.document.bindings.addFromSelectionAsync("table",{ id: "MyTableBinding" }, function (asyncResult) {
    if (asyncResult.status == "failed") {
              showMessage("Action failed with error: " + asyncResult.error.message);
           }
           else {
              showMessage("Added new binding with type: " + asyncResult.value.type +" and id: " + asyncResult.value.id);
           }
    });

    //Go to binding by id.
    Office.context.document.goToByIdAsync("MyTableBinding", Office.GoToType.Binding, function (asyncResult) {
        if (asyncResult.status == "failed") {
            showMessage("Action failed with error: " + asyncResult.error.message);
        }
        else {
            showMessage("Navigation successful");
        }
    });
}
```



 **前往試算表 (Excel) 中的表格**

下列範例顯示如何：


-  **依名稱指定表格**，做為要前往的表格。
    
-  **傳遞匿名的回呼函數**，可將作業狀態傳回至 _goToByIdAsync_ 方法的 **callback** 參數。
    
-  在增益集頁面上**顯示值**。
    



```js
function goToTable() {
    Office.context.document.goToByIdAsync("Table1", Office.GoToType.NamedItem, function (asyncResult) {
        if (asyncResult.status == "failed") {
            showMessage("Action failed with error: " + asyncResult.error.message);
        }
        else {
            showMessage("Navigation successful");
        }
    });
}
```



 **依 id 前往目前選取的投影片 (PowerPoint)**

下列範例顯示如何：


-  使用 **getSelectedDataAsync** 方法，取得目前選取之投影片的 [id](../../reference/shared/document.getselecteddataasync.md)。
    
-  **指定傳回的 id**，做為要前往的投影片。
    
-  **傳遞匿名的回呼函數**，可將作業狀態傳回至 _goToByIdAsync_ 方法的 **callback** 參數。
    
-  顯示由 `asyncResult.value` 傳回之 stringified JSON 物件的**值**，其包含有關在增益集頁面上所選取投影片的資訊。
    



```js
var firstSlideId = 0;
function gotoSelectedSlide() {
    //Get currently selected slide's id
    Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            firstSlideId = asyncResult.value.slides[0].id;
            app.showNotification(JSON.stringify(asyncResult.value));
        }
    });
    //Go to slide by id.
    Office.context.document.goToByIdAsync(firstSlideId, Office.GoToType.Slide, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification("Navigation successful");
        }
    });
}
```



 **依索引移至投影片 (PowerPoint)**

下列範例顯示如何：


-  **指定索引**，要前往之第一個、最後一個或下一個投影片的索引。
    
-  **傳遞匿名的回呼函數**，可將作業狀態傳回至 _goToByIdAsync_ 方法的 **callback** 參數。
    
-  在增益集頁面上**顯示值**。
    



```js
function goToSlideByIndex() {
    var goToFirst = Office.Index.First;
    var goToLast = Office.Index.Last;
    var goToPrevious = Office.Index.Previous;
    var goToNext = Office.Index.Next;

    Office.context.document.goToByIdAsync(goToNext, Office.GoToType.Index, function (asyncResult) {
        if (asyncResult.status == "failed") {
            showMessage("Action failed with error: " + asyncResult.error.message);
        }
        else {
            showMessage("Navigation successful");
        }
    });
}
```




## <a name="support-details"></a>支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此方法。空白儲存格表示 Office 主應用程式不支援此方法。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支援的主應用程式 (依平台排序)**


||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|
|**PowerPoint**|Y|Y|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**可用於需求集合**|不在集合中|
|**最低權限等級**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**增益集類型**|內容、工作窗格|
|**文件庫**|Office.js|
|**命名空間**|Office|

## <a name="support-history"></a>支援歷程記錄



|**版本**|**變更**|
|:-----|:-----|
|1.1|新增對 PowerPoint Online 的支援。|
|1.1|新增 iPad 版 Office 中對 Excel、PowerPoint 和 Word 的支援。|
|1.1|已導入|
