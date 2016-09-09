
# CustomXmlParts.getByNamespaceAsync 方法
以非同步方式取得符合指定的命名空間的自訂 XML 組件陣列。

|||
|:-----|:-----|
|**主機︰**|Word|
|**可用於[需求集合](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|CustomXmlParts|
|**已新增於**|1.1|

```
Office.context.document.customXmlParts.getByNamespaceAsync(ns [, options], callback);
```


## 參數



|**名稱**|**類型	**|**說明**|**支援附註**|
|:-----|:-----|:-----|:-----|
| _ns_|**字串**|命名空間 URI。必要。||
| _options_|**物件**|指定下列任何一項[選擇性參數](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)||
| _asyncContext_|**陣列**、**布林值**、**null**、**數字**、**物件**、**字串**或**未定義**|無變更的情況下，於 **AsyncResult** 物件中傳回的任一類型使用者定義項目。||
| _callback_|**物件**|回呼傳回時所叫用的函數，其唯一的參數為 **AsyncResult** 類型。||

## 回呼值

傳遞至 _callback_ 參數的函數執行時，該函數會收到 [AsyncResult](../../reference/shared/asyncresult.md) 物件，您可以從回呼函數的唯一參數存取該物件。

在傳遞至 **getByNamespaceAsync** 方法的回呼函數中，您可以使用 **AsyncResult** 物件的屬性以傳回下列資訊。



|**屬性**|**用途**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|以非同步方式存取符合指定命名空間之 **CustomXMLPart** 物件的[陣列](../../reference/shared/customxmlnode.getnodesasync.md)。|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|判定作業成功或失敗。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|作業失敗時，存取提供錯誤資訊的 [Error](../../reference/shared/error.md) 物件。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|存取您的使用者定義**物件**或值 (如果您傳遞了其中一項做為 _asyncContext_ 參數)。|

## 範例




```js
function showXMLPartsInNamespace() {
    Office.context.document.customXmlParts.getByNamespaceAsync("http://tempuri.org", function (eventArgs) {
        write("Found " + eventArgs.value.length + " parts with this namespace");
    }); 
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## 支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此方法。空白儲存格表示 Office 主應用程式不支援此方法。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**可用於需求集合**|CustomXmlParts|
|**最低權限等級**|[ReadWriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**增益集類型**|Task pane|
|**文件庫**|Office.js|
|**命名空間**|Office|

## 支援歷程記錄



****


|**版本**|**變更**|
|:-----|:-----|
|1.1|新增 iPad 版 Office 中對 Word 的支援。|
|1.0|已導入|
