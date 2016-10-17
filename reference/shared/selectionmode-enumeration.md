
# <a name="selectionmode-enumeration"></a>SelectionMode 列舉
指定使用 [Document.goToByIdAsync](../../reference/shared/document.gotobyidasync.md) 方法時，是否要選取 (醒目提示) 要瀏覽至的位置。

|||
|:-----|:-----|
|**已導入至 Office.js 版本**|1.1|

|||
|:-----|:-----|
|**主應用程式︰**|Excel、PowerPoint、Word|
|**已新增於**|1.1|



```
Office.SelectionMode
```


## <a name="members"></a>成員


**值**


|**列舉**|**值**|**描述**|
|:-----|:-----|:-----|
|Office.SelectionMode.Selected|"selected"|將選取 (醒目提示) 的位置。|
|Office.SelectionMode.None|"none"|將游標移至位置的開頭。|

## <a name="support-details"></a>支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此方法。空白儲存格表示 Office 主應用程式不支援此方法。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支援的主應用程式 (依平台排序)**


||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|
|**PowerPoint**|Y|||
|**Word**|Y||Y|

|||
|:-----|:-----|
|**增益集類型**|內容、工作窗格|
|**文件庫**|Office.js|
|**命名空間**|Office|

## <a name="support-history"></a>支援歷程記錄



****


|**版本**|**變更**|
|:-----|:-----|
|1.1|已導入|
