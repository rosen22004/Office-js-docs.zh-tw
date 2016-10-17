
# <a name="context.touchenabled-property"></a>Context.touchEnabled 屬性
取得關於增益集是否在具有觸控功能的 Office 主應用程式中執行的資訊。

|||
|:-----|:-----|
|**主應用程式︰**|Excel、Word|
|**上次變更於**|1.1|

```
var isTouchEnabled = Office.context.touchEnabled;
```


## <a name="return-value"></a>傳回值

如果增益集是在觸控式裝置 (例如 iPad) 上執行，會傳回 **True**，否則會傳回 **False**。


## <a name="remarks"></a>備註

使用 **touchEnabled** 屬性決定增益集何時在觸控式裝置上執行，若有必要，可調整控制項的種類，以及增益集 UI 中的項目大小和間距，以配合觸控式互動。


## <a name="support-details"></a>支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此方法。空白儲存格表示 Office 主應用程式不支援此方法。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。

||**Office for iPad**|
|:-----|:-----|
|**Excel**|Y|
|**PowerPoint**|Y|
|**Word**|Y|

|||
|:-----|:-----|
|**最低權限等級**|[限制](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**增益集類型**|內容、工作窗格|
|**文件庫**|Office.js|
|**命名空間**|Office|

## <a name="support-history"></a>支援歷程記錄



****


|**版本**|**變更**|
|:-----|:-----|
|1.1|已導入。|
