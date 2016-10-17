
# <a name="customxmlpart-object"></a>CustomXMLPart 物件
代表 **CustomXMLParts** 集合中的單一 [CustomXMLPart](../../reference/shared/customxmlparts.customxmlparts.md)。

|||
|:-----|:-----|
|**主應用程式︰**|Word|
|**可用於[需求集合](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|CustomXmlParts|
|**上次變更於**|1.1|

```
Office.context.document.customXmlParts.getByIdAsync(id);
```


## <a name="members"></a>成員


**屬性**


|**名稱**|**描述**|
|:-----|:-----|
|[builtIn](../../reference/shared/customxmlpart.builtin.md)|取得一個值，指出 CustomXMLPart 是否為內建的。|
|[id](../../reference/shared/customxmlpart.id.md)|取得 CustomXMLPart 的 GUID|
|[namespaceManager](../../reference/shared/customxmlpart.namespacemanager.md)|取得針對目前 CustomXMLPart 所用的命名空間前置詞對應集合 (CustomXMLPrefixMappings)。|

**方法**


|**名稱**|**描述**|
|:-----|:-----|
|[addHandlerAsync](../../reference/shared/customxmlpart.addhandlerasync.md)|以非同步方式新增 **CustomXmlPart** 物件事件的事件處理常式。|
|[deleteAsync](../../reference/shared/customxmlpart.deleteasync.md)|以非同步方式從集合中刪除這個自訂 XML 組件。|
|[getNodesAsync](../../reference/shared/customxmlpart.getnodesasync.md)|以非同步方式從這個自訂 XML 組件中，取得任何符合指定 XPath 的 CustomXmlNodes。|
|[getXmlAsync](../../reference/shared/customxmlpart.getxmlasync.md)|以非同步方式取得此自訂 XML 組件內的 XML。|
|[removeHandlerAsync](../../reference/shared/customxmlpart.removehandlerasync.md)|移除 **CustomXmlPart** 物件事件的事件處理常式。|

**事件**


|**名稱**|**描述**|
|:-----|:-----|
|[dataNodeDeleted](../../reference/shared/customxmlpart.datanodedeleted.event.md)|刪除節點時，就會發生。|
|[dataNodeInserted](../../reference/shared/customxmlpart.datanodeinserted.event.md)|插入節點時，就會發生。|
|[dataNodeReplaced](../../reference/shared/customxmlpart.datanodereplaced.event.md)|取代節點時，就會發生。|

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
