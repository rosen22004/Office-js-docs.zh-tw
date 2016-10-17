# <a name="excel-javascript-api-reference"></a>Excel JavaScript API 參考

您可以使用 Excel JavaScript API 來建置 Excel 2016 的增益集。下列清單顯示 API 中可用的高層次 Excel 物件。每個物件頁面連結都會描述物件可用的屬性、關聯性和方法。請瀏覽連結以了解詳細資訊。

* [Workbook](../../reference/excel/workbook.md)：這是最上層物件，包含相關的活頁簿物件，例如工作表、表格、範圍等等。也可以用來列出相關的參考。
* [Worksheet](../../reference/excel/worksheet.md)：Worksheets 集合的成員。Worksheets 集合包含活頁簿中的所有 Worksheet 物件。
    * [Worksheet 集合](../../reference/excel/worksheetcollection.md)：屬於活頁簿一部份的所有 Worksheet 物件的集合。
* [Range](../../reference/excel/range.md)：代表儲存格、列、欄，或是包含一個或多個連續儲存格區塊的儲存格選取範圍。
* [Table](../../reference/excel/table.md)：代表分類儲存格的集合，可讓您輕鬆管理資料。
    * [Table 集合](../../reference/excel/tablecollection.md)：活頁簿或工作表中的表格集合。
    * [TableColumn 集合](../../reference/excel/tablecolumncollection.md)：表格中所有欄的集合。
    * [TableRow 集合](../../reference/excel/tablerowcollection.md)：表格中所有列的集合。
* [Chart](../../reference/excel/chart.md)：代表工作表中的 Chart 物件，這是基礎資料的視覺表示法。
    * [Chart 集合](../../reference/excel/chartcollection.md)：工作表中圖表的集合。
* [TableSort](../../reference/excel/tablesort.md):代表在 Table 物件上排序作業的物件。
* [RangeSort](../../reference/excel/rangesort.md)：代表在 Range 物件上排序作業的物件。
* [篩選](../../reference/excel/filter.md)代表管理表格的欄位篩選的篩選物件。
* [保護工作表](../../reference/excel/worksheetprotection.md)：代表 worksheet 物件的保護。
* [工作表函數](../../reference/excel/functions.md)：代表從 JavaScript 中呼叫之 Microsoft Excel 工作表函數的容器。
* [NamedItem](../../reference/excel/nameditem.md)：代表一個儲存格範圍的已定義名稱或一個值。名稱可以是原始命名的物件、range 物件等。
    * [NamedItem 集合](../../reference/excel/nameditemcollection.md)：活頁簿中的 NamedItem 物件集合。
* [Binding](../../reference/excel/binding.md)：抽象類別，代表繫結至活頁簿的某個區段。
    * [Binding 集合](../../reference/excel/bindingcollection.md)：屬於活頁簿一部份的所有 Binding 物件的集合。
* [TrackedObject 集合](../../reference/excel/trackedobjectscollection.md)：可讓增益集跨 sync() 批次管理 range 物件參考。
* [Request Context](../../reference/excel/requestcontext.md)：RequestContext 物件可協助向 Excel 應用程式提出要求。


##### <a name="additional-resources"></a>其他資源

*  [Excel 增益集程式設計概觀](excel-add-ins-javascript-programming-overview.md)
*  [建立第一個 Excel 增益集](build-your-first-excel-add-in.md)
*  [Excel 的程式碼片段總管](http://officesnippetexplorer.azurewebsites.net/#/snippets/excel)

