
# <a name="nodedeletedeventargs.oldnode-property"></a>NodeDeletedEventArgs.oldNode 屬性
取得剛從 **CustomXmlPart** 物件中刪除的節點。

|||
|:-----|:-----|
|**主應用程式︰**|Word|
|**可用於[需求集合](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|CustomXmlParts|
|**上次變更於**|1.1|

```
var myNode = eventArgsObj.oldNode;
```


## <a name="return-value"></a>傳回值

[CustomXmlNode](../../reference/shared/customxmlnode.customxmlnode.md)，表示剛刪除節點。


## <a name="remarks"></a>備註

請注意，當您從文件中移除樹狀子目錄時，這個節點可能會有子項。此外，這個節點將是「中斷連線」的節點，這表示您可以從該節點向下查詢，但是不能沿著樹狀目錄向上查詢 - 該節點會顯示為獨立的節點。


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




|**版本**|**變更**|
|:-----|:-----|
|1.1|新增 iPad 版 Office 中對 Word 的支援。|
|1.0|已導入|
