
# NodeDeletedEventArgs 物件
提供刪除引發 [dataNodeDeleted](../../reference/shared/customxmlpart.datanodedeleted.event.md) 事件之節點的相關資訊。

|||
|:-----|:-----|
|**主機︰**|Word|
|**可用於[需求集合](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|CustomXmlParts|
|**已新增於**|1.1|

```
NodeDeletedEventArgs
```


## 成員


**屬性**


|**名稱**|**說明**|
|:-----|:-----|
|[isUndoRedo](../../reference/shared/customxmlpart.isundoredo.md)|取得關於使用者是否將刪除節點，當做 [復原/取消復原] 動作的一部分的資訊。|
|[oldNextSibling](../../reference/shared/customxmlpart.oldnextsibling.md)|取得剛從 **CustomXMLPart** 物件中刪除的節點，先前的下個同層級。|
|[oldNode](../../reference/shared/customxmlpart.oldnode.md)|取得剛從 **CustomXmlPart** 物件中刪除的節點。|

## 支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此物件。空白儲存格表示 Office 主應用程式不支援此物件。

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




|**版本**|**變更**|
|:-----|:-----|
|1.1|新增 iPad 版 Office 中對 Word 的支援。|
|1.0|已導入|
