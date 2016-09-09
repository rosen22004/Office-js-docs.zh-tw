
# CustomXmlNode 物件
代表在文件中樹狀目錄中的 XML 節點。

|||
|:-----|:-----|
|**主機︰**|Word|
|**可用於[需求集合](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|CustomXmlParts|
|**上次變更於**|1.1|

```js
CustomXmlNode
```


## 成員


**屬性**


|**名稱**|**說明**|
|:-----|:-----|
|[baseName](../../reference/shared/customxmlnode.basename.md)|取得節點的基本名稱，但不含命名空間前置詞 (如果有的話)。|
|[nodeType](../../reference/shared/customxmlnode.nodetype.md)|取得 **CustomXMLNode** 的類型。|
|[namespaceUri](../../reference/shared/customxmlnode.namespaceuri.md)|擷取 **CustomXMLPart** 的字串 GUID。|

**方法**


|**名稱**|**說明**|
|:-----|:-----|
|[getNodesAsync](../../reference/shared/customxmlnode.getnodesasync.md)|以非同步方式取得節點，做為符合相對的 XPath 運算式的 **CustomXMLNode** 物件的陣列。|
|[getNodeValueAsync](../../reference/shared/customxmlnode.getnodevalueasync.md)|以非同步方式取得節點的值。|
|[getTextAsync](customxmlnode.gettextasync.md)|以非同步方式取得自訂 XML 組件中 XML 節點的文字。|
|[getXmlAsync](../../reference/shared/customxmlnode.getxmlasync.md)|以非同步方式取得節點的 XML。|
|[setNodeValueAsync](../../reference/shared/customxmlnode.setnodevalueasync.md)|以非同步方式設定節點的值。|
|[setTextAsync](customxmlnode.settextasync.md)|以非同步方式設定自訂 XML 組件中 XML 節點的文字。|
|[setXmlAsync](../../reference/shared/customxmlnode.setxmlasync.md)|以非同步方式設定節點的 XML。|

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
