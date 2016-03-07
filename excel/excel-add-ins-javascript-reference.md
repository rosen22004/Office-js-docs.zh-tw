# Excel 增益集 JavaScript API 參考

_適用版本：Excel 2016、Office 2016_

以下連結會顯示 API 中可用的高階 Excel 物件。每個物件頁面連結都會說明物件可用的屬性、關聯性和方法。請瀏覽下列連結，了解詳細資訊。
	
* [Workbook](resources/workbook.md)：這是最上層物件，包含相關的活頁簿物件，例如工作表、表格、範圍等等。也可以用來列出相關的參考。 
* [Worksheet](resources/worksheet.md)：Worksheets 集合的成員。Worksheets 集合包含活頁簿中的所有 Worksheet 物件。
	* [Worksheet 集合](resources/worksheetcollection.md)：屬於活頁簿一部份的所有 Worksheet 物件的集合。 
* [Range](resources/range.md)：代表儲存格、列、欄，或是包含一個或多個連續儲存格區塊的儲存格選取範圍。  
* [Table](resources/table.md)：代表分類儲存格的集合，可讓您輕鬆管理資料。 
	* [Table 集合](resources/tablecollection.md)：活頁簿或工作表中的表格集合。 
	* [TableColumn 集合](resources/tablecolumncollection.md)：表格中所有欄的集合。 
	* [TableRow 集合](resources/tablerowcollection.md)：表格中所有列的集合。 
* [Chart](resources/chart.md)：代表工作表中的 Chart 物件，這是基礎資料的視覺表示法。   
	* [Chart 集合](resources/chartcollection.md)：工作表中圖表的集合。	
* [NamedItem](resources/nameditem.md)：代表一個儲存格範圍的已定義名稱或一個值。名稱可以是原始命名的物件、range 物件等。
	* [NamedItem 集合](resources/nameditemcollection.md)：活頁簿中的 NamedItem 物件集合。
* [Binding](resources/binding.md)：抽象類別，代表繫結至活頁簿的某個區段。
	* [Binding 集合](resources/bindingcollection.md)：屬於活頁簿一部份的所有 Binding 物件的集合。 
* [TrackedObject 集合](resources/trackedobjectscollection.md)：可讓增益集跨 sync() 批次管理 range 物件參考。 
* [Request Context](resources/requestcontext.md)：RequestContext 物件可協助向 Excel 應用程式提出要求。


##### 其他資源

*  [Excel 增益集程式設計概觀](excel-add-ins-programming-overview.md)
*  [建立第一個 Excel 增益集](build-your-first-excel-add-in.md)
*  [Excel 的程式碼片段總管](http://officesnippetexplorer.azurewebsites.net/#/snippets/excel)
*  [Excel 增益集程式碼範例](excel-add-ins-code-samples.md) 


