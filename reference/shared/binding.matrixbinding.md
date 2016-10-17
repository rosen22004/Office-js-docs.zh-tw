
# <a name="matrixbinding-object"></a>MatrixBinding 物件
表示資料列和資料行的兩個維度中的繫結。 

|||
|:-----|:-----|
|**主應用程式︰**|Excel、Word|
|**可用於[需求集合](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|MatrixBindings|
|**上次變更於 Selection**|1.1|

```
MatrixBinding
```


**屬性**


|**名稱**|**描述**|
|:-----|:-----|
|[columnCount](../../reference/shared/binding.matrixbinding.columncount.md)|以整數值在矩陣資料結構中，取得資料欄的數目。|
|[rowCount](../../reference/shared/binding.matrixbinding.rowcount.md)|以整數值在矩陣資料結構中，取得資料列的數目。|

## <a name="remarks"></a>備註

**MatrixBinding** 物件從 [Binding](../../reference/shared/binding.id.md) 物件繼承 [id](../../reference/shared/binding.type.md) 屬性、[type](../../reference/shared/binding.getdataasync.md) 屬性、[getDataAsync](../../reference/shared/binding.setdataasync.md) 方法，以及 [setDataAsync](../../reference/shared/binding.md) 方法。


## <a name="support-details"></a>支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此方法。空白儲存格表示 Office 主應用程式不支援此方法。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支援的主應用程式 (依平台排序)**


||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**可用於需求集合**|MatrixBindings|
|**增益集類型**|內容、工作窗格|
|**文件庫**|Office.js|
|**命名空間**|Office|

## <a name="support-history"></a>支援歷程記錄



****


|**版本**|**變更**|
|:-----|:-----|
|1.1|新增 iPad 版 Office 中對 Excel 和 Word 的支援。|
|1.0|已導入|
