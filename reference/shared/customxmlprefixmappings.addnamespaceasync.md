
# <a name="customxmlprefixmappings.addnamespaceasync-method"></a>CustomXmlPrefixMappings.addNamespaceAsync 方法
以非同步的方式加入查詢項目時所使用的命名空間對應的前置詞。

|||
|:-----|:-----|
|**主應用程式︰**|Word|
|**可用於[需求集合](../../docs/overview/specify-office-hosts-and-api-requirements.md)**||
|**已新增於**|1.1|

```
customXmlPrefixMappingsObj.addNamespaceAsync(prefix, ns [, options], callback);
```


## <a name="parameters"></a>參數



|**名稱**|**類型**|**描述**|**支援附註**|
|:-----|:-----|:-----|:-----|
| _prefix_|**string**|指定要新增至前置詞對應清單的前置詞。必要。||
| _ns_|**string**|指定要指派給新增前置詞的命名空間。必要。||
| _options_|**object**|指定下列任何一項[選擇性參數](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)||
| _asyncContext_|**陣列**、**布林值**、**null**、**數字**、**物件**、**字串**或**未定義**|無變更的情況下，於 **AsyncResult** 物件中傳回的任一類型使用者定義項目。||
| _callback_|**object**|回呼傳回時所叫用的函數，其唯一的參數為 **AsyncResult** 類型。||

## <a name="callback-value"></a>回呼值

傳遞至 _callback_ 參數的函數執行時，該函數會收到 [AsyncResult](../../reference/shared/asyncresult.md) 物件，您可以從回呼函數的唯一參數存取該物件。

在傳遞至 **addNamespaceAsync** 方法的回呼函數中，您可以使用 **AsyncResult** 物件的屬性以傳回下列資訊。



|**屬性**|**用於...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|因為沒有可擷取的物件或資料，所以一律傳回 **undefined**。|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|判定作業成功或失敗。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|作業失敗時，存取提供錯誤資訊的 [Error](../../reference/shared/error.md) 物件。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|存取您的使用者定義**物件**或值 (如果您傳遞了其中一項做為 _asyncContext_ 參數)。|

## <a name="remarks"></a>備註

如果前置詞已經存在於命名空間管理員中，則此方法會覆寫該前置詞的對應，除非該前置詞是由資料存放區於內部新增或使用，若是如此則會傳回錯誤。


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
|1.1|新增 iPad 版 Office 中對 Word 的支援。|
|1.0|已導入|