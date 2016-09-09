
# CustomXMLParts 物件
代表 [CustomXMLPart](../../reference/shared/customxmlpart.customxmlpart.md) 物件的集合。

|||
|:-----|:-----|
|**主機︰**|Word|
|**可用於[需求集合](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|CustomXmlParts|
|**上次變更於**|1.1|

```
Office.context.document.customXmlParts
```


## 成員


**方法**


|**名稱**|**說明**|
|:-----|:-----|
|[addAsync](../../reference/shared/customxmlparts.addasync.md)|以非同步方式將新的自訂 XML 組件加入至檔案中。|
|[getByIdAsync](../../reference/shared/customxmlparts.getbyidasync.md)|以非同步方式根據其 ID 取得自訂 XML 組件。|
|[getByNamespaceAsync](../../reference/shared/customxmlparts.getbynamespaceasync.md)|以非同步方式取得符合指定的命名空間的自訂 XML 組件陣列。|

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
