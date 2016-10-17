
# <a name="filtertype-enumeration"></a>FilterType 列舉
指定擷取資料時，是否從主應用程式套用篩選條件。

|||
|:-----|:-----|
|**主應用程式︰**|Excel、Project、Word|
|**上次變更於**|1.1|

```js
Office.FilterType
```


## <a name="members"></a>成員


**值**


|**列舉**|**值**|**描述**|
|:-----|:-----|:-----|
|Office.FilterType.All|"all"|傳回所有資料 (不依主應用程式篩選)。|
|Office.FilterType.OnlyVisible|"onlyVisible"|只傳回可見的資料 (如同依主應用程式篩選)。|

## <a name="support-details"></a>支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此列舉。空白儲存格表示 Office 主應用程式不支援此列舉。


如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支援的主應用程式 (依平台排序)**


||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|
|**Project**|Y|||
|**Word**|Y||Y|

|||
|:-----|:-----|
|**增益集類型**|內容、工作窗格|
|**文件庫**|Office.js|
|**命名空間**|Office|

## <a name="support-history"></a>支援歷程記錄

|**版本**|**變更**|
|:-----|:-----|
|1.1|新增 iPad 版 Office 中對 Excel 和 Word 的支援。|
|1.0|已導入|
