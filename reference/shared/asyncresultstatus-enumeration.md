
# <a name="asyncresultstatus-enumeration"></a>AsyncResultStatus 列舉
指定非同步呼叫的結果。 

|||
|:-----|:-----|
|**主應用程式︰**|Access、Excel、Outlook、PowerPoint、Project、Word|
|**上次變更於**|1.1|

```
Office.AsyncResultStatus
```


## <a name="members"></a>成員


**值**


|**列舉**|**值**|**描述**|
|:-----|:-----|:-----|
|Office.AsyncResultStatus.Succeeded|"succeeded"|呼叫成功。|
|Office.AsyncResultStatus.Failed|"failed"|呼叫失敗。|

## <a name="remarks"></a>備註

傳回 [AsyncResult](../../reference/shared/asyncresult.status.md) 物件的 [status](../../reference/shared/asyncresult.md) 屬性。


## <a name="support-details"></a>支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此列舉。空白儲存格表示 Office 主應用程式不支援此列舉。


如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支援的主應用程式 (依平台排序)**


||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for iPad**|**裝置適用的 OWA**|**Office for Mac**|
|:-----|:-----|:-----|:-----|:-----|:-----|
|**Access**|Y|||||
|**Excel**|Y|Y|Y|||
|**Outlook**|Y|Y||Y|Y|
|**PowerPoint**|Y|Y|Y|||
|**Project**|Y|||||
|**Word**|Y||Y|||

|||
|:-----|:-----|
|**增益集類型**|內容、工作窗格、Outlook|
|**文件庫**|Office.js|
|**命名空間**|Office|

## <a name="support-history"></a>支援歷程記錄


|**版本**|**變更**|
|:-----|:-----|
|1.1|新增 iPad 版 Office 中對 Excel、PowerPoint 和 Word 的支援。|
|1.1|新增對 Access 增益集的支援。|
|1.0|已導入|