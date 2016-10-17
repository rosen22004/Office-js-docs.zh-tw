
# <a name="customxmlpart.namespacemanager-property"></a>CustomXmlPart.namespaceManager 屬性
取得針對目前 CustomXMLPart 所用的命名空間前置詞對應集合 (CustomXMLPrefixMappings)。

|||
|:-----|:-----|
|**主應用程式︰**|Word|
|**可用於[需求集合](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|CustomXmlParts|
|**上次變更於**|1.1|

```js
var nsManager = customXmlPart.namespaceManager);
```


## <a name="return-value"></a>傳回值

[CustomXmlPrefixMappings](../../reference/shared/customxmlprefixmappings.customxmlprefixmappings.md) 物件。


## <a name="example"></a>範例




```js
function setXMLPartNamespaceManagerNamespace() {
    Office.context.document.customXmlParts.getByIdAsync("{3BC85265-09D6-4205-B665-8EB239A8B9A1}", function (result) {
        var xmlPart = result.value;
        xmlPart.namespaceManager.addNamespaceAsync("myPrefix", "myNamespace");
    });
}

```




## <a name="support-details"></a>支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此屬性。空白儲存格表示 Office 主應用程式不支援此屬性。

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
