
# <a name="tablebinding.settableoptionsasync-method"></a>TableBinding.setTableOptionsAsync 方法
更新繫結表格上的表格格式設定選項。

|||
|:-----|:-----|
|**主應用程式︰**|Excel|
|**可用於[需求集合](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|不在集合中|
|**已新增於**|1.1|

```
bindingObj.setTableOptionsAsync(tableOptions [,options] , callback);
```


## <a name="parameters"></a>參數



|**名稱**|**類型**|**描述**|**支援附註**|
|:-----|:-----|:-----|:-----|
| _tableOptions_|**object**|包含定義要套用之表格選項的屬性名稱值組清單的物件常值。必要。||
| _options_|**object**|指定下列任何一項[選擇性參數](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)||
| _asyncContext_|**陣列**、**布林值**、**null**、**數字**、**物件**、**字串**或**未定義**|無變更的情況下，於 **AsyncResult** 物件中傳回的任一類型使用者定義項目。||
| _callback_|**object**|回呼傳回時所叫用的函數，其唯一的參數為 **AsyncResult** 類型。||

## <a name="callback-value"></a>回呼值

傳遞至 _callback_ 參數的函數執行時，該函數會收到 [AsyncResult](../../reference/shared/asyncresult.md) 物件，您可以從回呼函數的唯一參數存取該物件。

在傳遞至 **goToByIdAsync** 方法的回呼函數中，您可以使用 **AsyncResult** 物件的屬性以傳回下列資訊。



|**屬性**|**用於...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|設定表格選項時，因為沒有可擷取的資料或物件，所以一律傳回 **undefined**。|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|判定作業成功或失敗。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|作業失敗時，存取提供錯誤資訊的 [Error](../../reference/shared/error.md) 物件。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|存取您的使用者定義**物件**或值 (如果您傳遞了其中一項做為 _asyncContext_ 參數)。|

## <a name="example"></a>範例

下列範例顯示如何：


-  **建立物件常值**，其可指定[表格格式設定選項](../../docs/excel/format-tables-in-add-ins-for-excel.md)，以在繫結表格上更新。
    
-  在先前的繫結表格上**呼叫 setTableOptions** (具有 **id**`myBinding`)，傳遞帶有格式設定的物件做為 _tableOptions_ 參數。
    

```js
function updateTableFormatting(){
    var tableOptions = {bandedRows: true, filterButton: false, style: "TableStyleMedium3"}; 

    Office.select("bindings#myBinding").setTableOptionsAsync(tableOptions, function(asyncResult){});
}
```




## <a name="support-details"></a>支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此方法。空白儲存格表示 Office 主應用程式不支援此方法。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支援的主應用程式 (依平台排序)**


||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|

|||
|:-----|:-----|
|**可用於需求集合**|不在集合中。|
|**最低權限等級**|[WriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**增益集類型**|內容、工作窗格|
|**文件庫**|Office.js|
|**命名空間**|Office|

## <a name="support-history"></a>支援歷程記錄




|**版本**|**變更**|
|:-----|:-----|
|1.1|新增 iPad 版 Office 中對 Excel 的支援。|
|1.1|已導入|
