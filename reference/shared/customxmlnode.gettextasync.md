
# <a name="customxmlnode.gettextasync-method"></a>CustomXmlNode.getTextAsync 方法
以非同步方式取得自訂 XML 組件中 XML 節點的文字。

|||
|:-----|:-----|
|**主應用程式︰**|Word|
|**可用於[需求集合](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|CustomXmlParts|
|**已新增於**|1.2|

```js
customXmlNodeObj.getTextAsync([asyncContext,]callback(asyncResult);
```


## <a name="parameters"></a>參數



|**名稱**|**類型**|**描述**|
|:-----|:-----|:-----|
| _asyncContext_|**object**|選擇性。[AsyncResult](../../reference/shared/asyncresult.md) 物件之 asyncContext 屬性上可用的使用者定義物件。當回呼是具名函數時，使用此物件以將物件或值提供給 **AsyncResult**。|
| _callback_|**object**|回呼傳回時所叫用的函數，其唯一的參數為 **AsyncResult** 類型。|

## <a name="callback-value"></a>回呼值

傳遞至 _callback_ 參數的函數執行時，該函數會收到 [AsyncResult](../../reference/shared/asyncresult.md) 物件，您可以從回呼函數的唯一參數存取該物件。

在傳遞至 **getTextAsync** 方法的回呼函數中，您可以使用 **AsyncResult** 物件的屬性以傳回下列資訊。



|**屬性**|**用於...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|存取包含參考節點之內部文字的**字串**。|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|表示作業成功或失敗。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|作業失敗時，存取提供錯誤資訊的 [Error](../../reference/shared/error.md) 物件。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|存取您的使用者定義**物件**或值 (如果您傳遞了其中一項做為 _asyncContext_ 參數)。如果未設定 _asyncContext_，此屬性會傳回未定義。|

## <a name="example"></a>範例

瞭解如何取得自訂 XML 組件中節點的文字值。


```js
// Get the built-in core properties XML part by using its ID. This results in a call to Word.
Office.context.document.customXmlParts.getByIdAsync("{6C3C8BC8-F283-45AE-878A-BAB7291924A1}", function (getByIdAsyncResult) {
    
    // Access the XML part.
    var xmlPart = getByIdAsyncResult.value;
    
    // Add namespaces to the namespace manager. These two calls result in two calls to Word.
    xmlPart.namespaceManager.addNamespaceAsync('cp', 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties', function () {
        xmlPart.namespaceManager.addNamespaceAsync('dc', 'http://purl.org/dc/elements/1.1/', function () {

            // Get XML nodes by using an Xpath expression. This results in a call to Word.
            xmlPart.getNodesAsync("/cp:coreProperties/dc:title", function (getNodesAsyncResult) {
                
                // Get the first node returned by using the Xpath expression. 
                var node = getNodesAsyncResult.value[0];
                
                // Get the text value of the node and use the asyncContext. This results in a call to Word. 
                // The results are logged to the browser console.
                node.getTextAsync({asyncContext: "StateNormal"}, function (getTextAsyncResult) {
                   console.log("Text of the title element = " + getTextAsyncResult.value;
                   console.log("The asyncContext value = " + getTextAsyncResult.asyncContext;
                });
            });
        });
    });
});
```


## <a name="support-details"></a>支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此方法。空白儲存格表示 Office 主應用程式不支援此方法。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**可用於需求集合**|CustomXmlParts|
|**最低權限等級**|[ReadWriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**增益集類型**|Task pane|
|**文件庫**|Office.js|
|**命名空間**|Office|

## <a name="support-history"></a>支援歷程記錄



****


|**版本**|**變更**|
|:-----|:-----|
|1.1|已新增 getTextAsync。|
