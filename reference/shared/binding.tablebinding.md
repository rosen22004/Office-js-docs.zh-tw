
# TableBinding 物件
表示資料列和資料行的兩個維度中的繫結，選擇性地包含標頭。

|||
|:-----|:-----|
|**主機︰**|Access、Excel、PowerPoint、Project、Word|
|**可用於[需求集合](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|TableBindings|
|**上次變更於 Selection**|1.1|

```
TableBinding
```


## 成員


**屬性**


|**名稱**|**說明**|**更新 Office.js v1.1**|
|:-----|:-----|:-----|
|[columnCount](../../reference/shared/binding.tablebinding.columncount.md)|取得指定之 **TableBinding** 物件中的欄數。|新增支援 Access 內容增益集中的表格繫結。|
|[hasHeaders](../../reference/shared/binding.tablebinding.hasheaders.md)|如果指定的 **TableBinding** 有標頭，則會傳回 True；否則會傳回 False。|新增支援 Access 內容增益集中的表格繫結。|
|[rowCount](../../reference/shared/binding.tablebinding.rowcount.md)|在指定之 **TableBinding** 物件中的列數。|基於效能原因，Access 的內容增益集中一律會傳回 -1。|

**方法**


|**名稱**|**說明**|**更新 Office.js v1.1**|
|:-----|:-----|:-----|
|[addColumnsAsync](../../reference/shared/binding.tablebinding.addcolumnsasync.md)|將資料欄和值加入表格中。||
|[addRowsAsync](../../reference/shared/binding.tablebinding.addrowsasync.md)|將資料列和值加入表格中。|新增支援 Access 內容增益集中的表格繫結。|
|[clearFormatsAsync](../../reference/shared/binding.tablebinding.clearformatsasync.md)|清除繫結檔案上的格式設定。|Excel 增益集之 Office.js v1.1 的新增功能。|
|[deleteAllDataValuesAsync](../../reference/shared/binding.tablebinding.deletealldatavaluesasync.md)|刪除表格中所有非標題列及其值，為主應用程式適當地移動。|新增支援 Access 內容增益集中的表格繫結。|
|[setDataAsync](../../reference/shared/binding.setdataasync.md)|將資料寫入指定的繫結物件所代表文件的繫結區段。|<ul><li>新增支援 Access 內容增益集中的表格繫結。</li><li>新增支援在 Excel 增益集中將資料寫入繫結表格時設定格式。</li></ul>|
|[setFormatsAsync](../../reference/shared/binding.tablebinding.setformatsasync.md)|設定繫結表格中指定的項目和資料上的儲存格與表格格式。|可設定 Excel 增益集中的表格格式。|
|[setTableOptionsAsync](../../reference/shared/binding.tablebinding.settableoptionsasync.md)|更新繫結表格上的表格格式設定選項。|可設定 Excel 增益集中的表格格式。|

## 備註

**TableBinding** 物件從 [Binding](../../reference/shared/binding.id.md) 抽象物件繼承 [id](../../reference/shared/binding.type.md) 屬性、[type](../../reference/shared/binding.getdataasync.md) 屬性、[getDataAsync](../../reference/shared/binding.setdataasync.md) 方法，以及 [setDataAsync](../../reference/shared/binding.md) 方法。

在 Excel 中建立表格繫結之後，使用者新增至表格的每個新列會自動包含在繫結中 (將增加 **rowCount**)。


## 支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此物件。空白儲存格表示 Office 主應用程式不支援此物件。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支援的主應用程式 (依平台排序)**


||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**可用於需求集合**|TableBindings|
|**最低權限等級**|[WriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**增益集類型**|內容、工作窗格|
|**文件庫**|Office.js|
|**命名空間**|Office|

## 支援歷程記錄




|**版本**|**變更**|
|:-----|:-----|
|1.1|新增 iPad 版 Office 中對 Excel 和 Word 的支援。|
|1.1|新增支援在 Excel 中[插入表格時設定格式](../../docs/excel/format-tables-in-add-ins-for-excel.md)。|
|1.1|新增對 Access 增益集的支援。|
|1.0|已導入|
