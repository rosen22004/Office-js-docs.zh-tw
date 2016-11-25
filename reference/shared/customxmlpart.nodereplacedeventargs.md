
# <a name="nodereplacedeventargs-object"></a>NodeReplacedEventArgs 物件
提供引發 [dataNodeReplaced](../../reference/shared/customxmlpart.datanodereplaced.event.md) 事件之取代節點的相關資訊。

|||
|:-----|:-----|
|**主應用程式︰**|Word|
|**可用於[需求集合](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|CustomXmlParts|
|**上次變更於**|1.1|

```
NodeReplacedEventArgs
```


## <a name="members"></a>成員


**屬性**


|**名稱**|**描述**|
|:-----|:-----|
|[isUndoRedo](../../reference/shared/customxmlpart.isundoredo.md)|取得關於使用者是否將插入取代的節點，當做復原或取消復原作業的一部分的資訊。|
|[newNode](../../reference/shared/customxmlpart.newnode.md)|取得新的節點。|
|[oldNode](../../reference/shared/customxmlpart.oldnode.md)|取得舊 (取代) 的節點。|

## <a name="support-details"></a>支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此物件。空白儲存格表示 Office 主應用程式不支援此物件。

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