# <a name="excel-javascript-api-requirement-sets"></a>Excel JavaScript API 需求集合

需求集合是 API 成員的具名群組。Office 增益集使用資訊清單中所指定的需求集合，或使用執行階段檢查，以判定 Office 主應用程式是否支援增益集所需的的 API。如需詳細資訊，請參閱[指定 Office 主應用程式及 API 需求](../../docs/overview/specify-office-hosts-and-api-requirements.md)。

Excel 增益集可在多種 Office 版本上執行，包含 Office 2016 for Windows、iPad 版 Office、Mac 版 Office 以及 Office Online。下表列出 Excel 需求集合、支援需求集合的 Office 主應用程式，以及這些應用程式的組建或版本號碼。

|  需求集合  |  Office 2016 for Windows*  |  Office 2016 for iPad  |  Mac 版 Office 2016  | Office Online  |  Office Online 伺服器  |
|:-----|-----|:-----|:-----|:-----|:-----|
| ExcelApi 1.5 **Beta**  | 版本 1702 (建置 TBD) 或更新版本| 即將推出 |  即將推出| 即將推出 | 即將推出|
| ExcelApi 1.4 **Beta** | 版本 1702 (建置 TBD) 或更新版本| 即將推出 |  即將推出| 即將推出 | 即將推出|
| ExcelApi 1.3  | 版本 1608 (組建 7369.2055) 或更新版本| 1.27 或更新版本 |  15.27 或更新版本| 2016 年 9 月 | 版本 1608 (組建 7601.6800) 或更新版本|
| ExcelApi 1.2  | 版本 1601 (組建 6741.2088) 或更新版本 | 1.21 或更新版本 | 15.22 或更新版本| 2016 年 1 月 ||
| ExcelApi 1.1  | 版本 1509 (組建 4266.1001) 或更新版本 | 1.19 或更新版本 | 15.20 或更新版本| 2016 年 1 月 ||

> **附註**：透過 MSI 安裝的 Office 2016 組建編號是 16.0.4266.1001。這個版本只會包含 ExcelApi 1.1 需求集合。

若要瞭解關於版本、組建編號及 Office Online Server 的詳細資訊，請參閱︰

- [Office 365 用戶端的更新通道版本和組建編號](https://technet.microsoft.com/en-us/library/mt592918.aspx)
- [我使用的是哪個版本的 Office？](https://support.office.com/en-us/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19?ui=en-US&rs=en-US&ad=US&fromAR=1)
- [您可以在其中找到 Office 365 用戶端應用程式的版本和組建編號](https://technet.microsoft.com/en-us/library/mt592918.aspx#Anchor_1)
- [Office Online Server 概觀](https://technet.microsoft.com/en-us/library/jj219437(v=office.16).aspx)

## <a name="office-common-api-requirement-sets"></a>Office 通用 API 需求集合
如需通用 API 需求集合的詳細資訊，請參閱 [Office 通用 API 需求集合](office-add-in-requirement-sets.md)。

## <a name="whats-new-in-excel-javascript-api-14"></a>Excel JavaScript API 1.4 的新功能
以下是需求集合 1.3 中 Excel JavaScript API 的新功能。

### <a name="named-item-add-and-new-properties"></a>新增具名項目及新屬性

新屬性
* `comment`
* `scope` 工作表或活頁簿設定範圍項目
* `worksheet`傳回具名項目限於其中的工作表。

新方法
* `add(name: string, reference: Range or string, comment: string)`新增名稱至指定範圍的集合。
* `addFormulaLocal(name: string, formula: string, comment: string)` 使用使用者的公式地區設定，新增名稱至指定範圍的集合。

### <a name="settings-api-in-in-excel-namespace"></a>Excel 命名空間中的設定 API

[Setting](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_1.4_OpenSpec/reference/excel/setting.md) 物件代表保留至文件之設定的機碼值組。現在，我們已在 Excel 命名空間下新增與設定相關的 API。這並不會提供新功能 - 無論這使保留承諾型批次 API 語法能有多容易減少與 Excel 相關聯的工作之通用 API 的相依性。

API 包括 `getItem()` 以透過機碼取得設定項目和 `add()` 以將指定機碼值組新增至活頁簿。

### <a name="others"></a>其他

* 設定表格欄名稱 (先前的版本僅允許讀取)。
* 新增表格欄至表格 (先前的版本僅允許讀取)。
* 一次新增多個列至表格 (先前的版本僅允許一次新增 1 列)。
* `range.getColumnsAfter(count: number)` 和 `range.getColumnsBefore(count: number)` 以取得目前的 Range 物件右/左邊的欄數。
* 取得項目或 null 物件的函數︰此功能可讓您使用機碼取得物件。如果物件不存在，傳回物件的 isNullObject 屬性為 true。這可使開發人員檢查物件是否存在，而不必透過例外狀況處理來處理物件。用於工作表、具名項目、繫結、圖表數列等

`worksheet.GetItemOrNullObject()`

### <a name="suspend-calculation"></a>擱置計算
擱置計算 (application.suspendCalculationUntilNextSync()) 直到呼叫下一個 "context.sync()"。一旦設定，開發人員便有責任重新計算活頁簿，以確保能夠傳播任何相依性。

此外，我們正在修正沒有重新計算已修改儲存格的 F9 重新計算錯誤。

|物件| 新功能| 描述|需求集合|
|:----|:----|:----|:----|
|[應用程式](../excel/application.md)|_方法_ > [suspendCalculationUntilNextSync()](../excel/application.md#suspendcalculationuntilnextsync)|擱置計算直至呼叫下一個 "context.sync()"。一旦設定，開發人員便有責任重新計算活頁簿，以確保能夠傳播任何相依性。|1.4|
|[bindingCollection](../excel/bindingcollection.md)|_方法_ > [getCount()](../excel/bindingcollection.md#getcount)|取得集合中的繫結數目。|1.4|
|[bindingCollection](../excel/bindingcollection.md)|_方法_ > [getItemOrNullObject(id: string)](../excel/bindingcollection.md#getitemornullobjectid-string)|依 ID 取得 Binding 物件。如果 Binding 物件不存在，會傳回 null 物件。|1.4|
|[chartCollection](../excel/chartcollection.md)|_方法_ > [getCount()](../excel/chartcollection.md#getcount)|傳回工作表中的圖表數目。|1.4|
|[chartCollection](../excel/chartcollection.md)|_方法_ > [getItemOrNullObject(id: string)](../excel/chartcollection.md#getitemornullobjectname-string)|使用其名稱取得圖表。如果有多個圖表具有相同的名稱，則會傳回第一個圖表。|1.4|
|[chartPointsCollection](../excel/chartpointscollection.md)|_方法_ > [getCount()](../excel/chartpointscollection.md#getcount)|傳回系列中的圖表點數目。|1.4|
|[chartSeriesCollection](../excel/chartseriescollection.md)|_方法_ > [getCount()](../excel/chartseriescollection.md#getcount)|傳回集合中的數列數目。|1.4|
|[namedItem](../excel/nameditem.md)|_屬性_ > 註解|表示與此名稱相關的註解。|1.4|
|[namedItem](../excel/nameditem.md)|_屬性_ > 註解|表示名稱是否限於活頁簿或或特定工作表。唯讀。可能的值為：Equal、Greater、GreaterEqual、Less、LessEqual、NotEqual。|1.4|
|[namedItem](../excel/nameditem.md)|_關聯性_ > 工作表|傳回具名項目限於其中的工作表。如果項目改為限於活頁簿，則擲回錯誤。唯讀。|1.4|
|[namedItem](../excel/nameditem.md)|_關聯性_ > worksheetOrNullObject|傳回具名項目限於其中的工作表。如果項目改為限於活頁簿，則傳回 null 物件。唯讀。|1.4|
|[namedItem](../excel/nameditem.md)|_方法_ > [delete()](../excel/nameditem.md#delete)|刪除指定的名稱。|1.4|
|[namedItem](../excel/nameditem.md)|_方法_ > [getRangeOrNullObject()](../excel/nameditem.md#getrangeornullobject)|傳回與名稱相關的 Range 物件。如果具名項目的類型不是範圍，則傳回 null 物件。|1.4|
|[namedItemCollection](../excel/nameditemcollection.md)|_方法_ > [Range 或 string, comment: string)](../excel/nameditemcollection.md#addname-string-reference-range-or-string-comment-string)|新增名稱至指定範圍的集合。|1.4|
|[namedItemCollection](../excel/nameditemcollection.md)|_方法_ > [addFormulaLocal(name: string, formula: string, comment: string)](../excel/nameditemcollection.md#addformulalocalname-string-formula-string-comment-string)|使用使用者的公式地區設定，新增名稱至指定範圍的集合。|1.4|
|[namedItemCollection](../excel/nameditemcollection.md)|_方法_ > [getCount()](../excel/nameditemcollection.md#getcount)|取得集合中的具名項目數目。|1.4|
|[namedItemCollection](../excel/nameditemcollection.md)|_方法_ > [getItemOrNullObject(id: string)](../excel/nameditemcollection.md#getitemornullobjectname-string)|使用其名稱取得 nameditem 物件。如果 nameditem 物件不存在，會傳回 null 物件。|1.4|
|[pivotTableCollection](../excel/pivottablecollection.md)|_方法_ > [getCount()](../excel/pivottablecollection.md#getcount)|取得集合中的樞紐分析表數目。|1.4|
|[pivotTableCollection](../excel/pivottablecollection.md)|_方法_ > [getItemOrNullObject(name: string)](../excel/pivottablecollection.md#getitemornullobjectname-string)|依名稱取得樞紐分析表。如果樞紐分析表不存在，會傳回 null 物件。|1.4|
|[範圍](../excel/range.md)|_方法_ > [getIntersectionOrNullObject(anotherRange:Range 或 string)](../excel/range.md#getintersectionornullobjectanotherrange-range-or-string)|取得範圍物件，代表特定範圍的矩形交集。如果找到沒有交集，則會傳回 null 物件。|1.4|
|[範圍](../excel/range.md)|_方法_ > [getUsedRangeOrNullObject(valuesOnly: bool)](../excel/range.md#getusedrangeornullobjectvaluesonly-bool)|傳回特定 Range 物件所使用的範圍。如果範圍內沒有已使用的儲存格，則此函數會傳回 null 物件。|1.4|
|[rangeViewCollection](../excel/rangeviewcollection.md)|_方法_ > [getCount()](../excel/rangeviewcollection.md#getcount)|取得集合中的 RangeView 物件數目。|1.4|
|[設定](../excel/setting.md)|_屬性_ > 索引鍵|傳回代表設定識別碼的索引鍵。唯讀。|1.4|
|[設定](../excel/setting.md)|_屬性_ > 值|表示儲存此設定的值。|1.4|
|[設定](../excel/setting.md)|_方法_ > [delete()](../excel/setting.md#delete)|刪除設定。|1.4|
|[settingCollection](../excel/settingcollection.md)|_屬性_ > 項目|設定物件的集合。唯讀。|1.4|
|[settingCollection](../excel/settingcollection.md)|_方法_ > [add(key: string, value: (any)[])](../excel/settingcollection.md#addkey-string-value-any)|將指定的設定新增或設定至活頁簿。|1.4|
|[settingCollection](../excel/settingcollection.md)|_方法_ > [getCount()](../excel/settingcollection.md#getcount)|取得集合中的設定數目。|1.4|
|[settingCollection](../excel/settingcollection.md)|_方法_ > [getItem(key: string)](../excel/settingcollection.md#getitemkey-string)|透過機碼取得設定項目。|1.4|
|[settingCollection](../excel/settingcollection.md)|_Method_ > [getItemOrNullObject(key: string)](../excel/settingcollection.md#getitemornullobjectkey-string)|透過機碼取得設定項目。如果設定不存在，會傳回 null 物件。|1.4|
|[settingsChangedEventArgs](../excel/settingschangedeventargs.md)|_關聯性_ > 設定|取得代表引發 SettingsChanged 事件之繫結的 Settings 物件。|1.4|
|[tableCollection](../excel/tablecollection.md)|_方法_ > [getCount()](../excel/tablecollection.md#getcount)|取得集合中的表格數目。|1.4|
|[tableCollection](../excel/tablecollection.md)|_方法_ > [getItemOrNullObject(key: number or string)](../excel/tablecollection.md#getitemornullobjectkey-number-or-string)|依名稱或 ID 取得表格。如果表格不存在，會傳回 null 物件。|1.4|
|[tableColumnCollection](../excel/tablecolumncollection.md)|_方法_ > [getCount()](../excel/tablecolumncollection.md#getcount)|取得表格中的欄數。|1.4|
|[tableColumnCollection](../excel/tablecolumncollection.md)|_方法_ > [getItemOrNullObject(key: number or string)](../excel/tablecolumncollection.md#getitemornullobjectkey-number-or-string)|依名稱或 ID 取得 column 物件。如果欄不存在，會傳回 null 物件。|1.4|
|[tableRowCollection](../excel/tablerowcollection.md)|_方法_ > [getCount()](../excel/tablerowcollection.md#getcount)|取得表格中的列數。|1.4|
|[活頁簿](../excel/workbook.md)|_關聯性_ > 設定|代表與活頁簿關聯的設定集合。唯讀。|1.4|
|[工作表](../excel/worksheet.md)|_相關性_ > 名稱|只限於目前工作表的名稱集合。唯讀。|1.4|
|[工作表](../excel/worksheet.md)|_方法_ > [getUsedRangeOrNullObject(valuesOnly: bool)](../excel/worksheet.md#getusedrangeornullobjectvaluesonly-bool)|使用的範圍是最小範圍，其中包含具有值或獲指派格式設定的任何儲存格。如果整個工作表空白，則此函數會傳回 null 物件。|1.4|
|[worksheetCollection](../excel/worksheetcollection.md)|_方法_ > [getCount(visibleOnly: bool)](../excel/worksheetcollection.md#getcountvisibleonly-bool)|取得集合中的工作表數目。|1.4|
|[worksheetCollection](../excel/worksheetcollection.md)|_方法_ > [getItemOrNullObject(key: string)](../excel/worksheetcollection.md#getitemornullobjectkey-string)|使用其名稱或 ID 取得 worksheet 物件。如果工作表不存在，會傳回 null 物件。|1.4|



## <a name="whats-new-in-excel-javascript-api-13"></a>Excel JavaScript API 1.3 的新功能
以下是需求集合 1.3 中 Excel JavaScript API 的新功能。

|物件| 新功能| 描述|需求集合|
|:----|:----|:----|:----|
|[繫結](../excel/binding.md)|_方法_ > [delete()](../excel/binding.md#delete)|刪除繫結。|1.3|
|[bindingCollection](../excel/bindingcollection.md)|_方法_ > [add(range:Range or string, bindingType: string, id: string)](../excel/bindingcollection.md#addrange-range-or-string-bindingtype-string-id-string)|將新的繫結新增至特定範圍。|1.3|
|[bindingCollection](../excel/bindingcollection.md)|_方法_ > [addFromNamedItem(name: string, bindingType: string, id: string)](../excel/bindingcollection.md#addfromnameditemname-string-bindingtype-string-id-string)|根據活頁簿中具名的項目，新增新的繫結。|1.3|
|[bindingCollection](../excel/bindingcollection.md)|_方法_ > [addFromSelection(bindingType: string, id: string)](../excel/bindingcollection.md#addfromselectionbindingtype-string-id-string)|根據目前的選取範圍，新增新的繫結。|1.3|
|[bindingCollection](../excel/bindingcollection.md)|_方法_ > [getItemOrNull(id: string)](../excel/bindingcollection.md#getitemornullid-string)|藉由識別碼取得繫結物件。如果繫結物件不存在，傳回物件的 isNull 屬性為 true。|1.3|
|[chartCollection](../excel/chartcollection.md)|_方法_ > [getItemOrNull(name: string)](../excel/chartcollection.md#getitemornullname-string)|使用其名稱取得圖表。如果有多個圖表具有相同的名稱，則會傳回第一個圖表。|1.3|
|[namedItemCollection](../excel/nameditemcollection.md)|_方法_ > [getItemOrNull(name: string)](../excel/nameditemcollection.md#getitemornullname-string)|使用其名稱取得 nameditem 物件。如果 nameditem 物件不存在，傳回物件的 isNull 屬性為 true。|1.3|
|[pivotTable](../excel/pivottable.md)|_屬性_ > 名稱|樞紐分析表的名稱。|1.3|
|[pivotTable](../excel/pivottable.md)|_關聯性_ > 工作表|包含目前樞紐分析表的工作表。唯讀。|1.3|
|[pivotTable](../excel/pivottable.md)|_方法_ > [refresh()](../excel/pivottable.md#refresh)|重新整理樞紐分析表。|1.3|
|[pivotTableCollection](../excel/pivottablecollection.md)|_屬性_ > 項目|pivotTable 物件的集合。唯讀。|1.3|
|[pivotTableCollection](../excel/pivottablecollection.md)|_方法_ > [getItem(name: string)](../excel/pivottablecollection.md#getitemname-string)|藉由名稱取得樞紐分析表。|1.3|
|[pivotTableCollection](../excel/pivottablecollection.md)|_方法_ > [getItemOrNull(name: string)](../excel/pivottablecollection.md#getitemornullname-string)|藉由名稱取得樞紐分析表。如果樞紐分析表不存在，傳回物件的 isNull 屬性為 true。|1.3|
|[範圍](../excel/range.md)|_方法_ > [getIntersectionOrNull(anotherRange:Range or string)](../excel/range.md#getintersectionornullanotherrange-range-or-string)|取得範圍物件，代表特定範圍的矩形交集。如果找到沒有交集，則會傳回 null 物件。|1.3|
|[範圍](../excel/range.md)|_方法_ > [getVisibleView()](../excel/range.md#getvisibleview)|代表目前範圍的可見資料列。|1.3|
|[rangeView](../excel/rangeview.md)|_屬性_ > cellAddresses|表示 RangeView 的儲存格位址。唯讀。|1.3|
|[rangeView](../excel/rangeview.md)|_屬性_ > columnWidth|傳回可見資料行的數目。唯讀。|1.3|
|[rangeView](../excel/rangeview.md)|_屬性_ > 公式|代表 A1 樣式標記法的公式。|1.3|
|[rangeView](../excel/rangeview.md)|_屬性_ > formulasLocal|以使用者的語言和數字格式地區設定，表示 A1 樣式標記法的公式。例如，英文的 "=SUM(A1, 1.5)" 公式在德文中會表示為 "=SUMME(A1; 1,5)"。|1.3|
|[rangeView](../excel/rangeview.md)|_屬性_ > formulasR1C1|代表 R1C1 樣式標記法的公式。|1.3|
|[rangeView](../excel/rangeview.md)|_屬性_ > 索引|傳回值，表示 RangeView 的索引。唯讀。|1.3|
|[rangeView](../excel/rangeview.md)|_屬性_ > numberFormat|代表特定儲存格的 Excel 數字格式代碼。|1.3|
|[rangeView](../excel/rangeview.md)|_屬性_ > rowCount|傳回可見資料列的數目。唯讀。|1.3|
|[rangeView](../excel/rangeview.md)|_屬性_ > 文字|所指定範圍的文字值。文字值與儲存格寬度無關。Excel UI 中出現的 # 替代符號不會影響 API 所傳回的文字值。唯讀。|1.3|
|[rangeView](../excel/rangeview.md)|_屬性_ > valueTypes|代表每個儲存格的資料類型。唯讀。可能的值為：Unknown、Empty、String、Integer、Double、Boolean、Error。|1.3|
|[rangeView](../excel/rangeview.md)|_屬性_ > 值|代表所指定範圍檢視的原始值。傳回的資料可能是 string、number 或 boolean 類型。包含錯誤的儲存格會傳回錯誤字串。|1.3|
|[rangeView](../excel/rangeview.md)|_關聯性_ > 資料列|代表與範圍關聯的範圍檢視集合。唯讀。|1.3|
|[rangeView](../excel/rangeview.md)|_方法_ > [getRange()](../excel/rangeview.md#getrange)|取得與目前的 RangeView 相關聯的父項範圍。|1.3|
|[rangeViewCollection](../excel/rangeviewcollection.md)|_屬性_ > 項目|rangeView 物件的集合。唯讀。|1.3|
|[rangeViewCollection](../excel/rangeviewcollection.md)|_方法_ > [getItemAt(index: number)](../excel/rangeviewcollection.md#getitematindex-number)|透過其索引取得 RangeView 資料列。以 0 開始編製索引。|1.3|
|[設定](../excel/setting.md)|_屬性_ > 索引鍵|傳回代表設定識別碼的索引鍵。唯讀。|1.3|
|[設定](../excel/setting.md)|_方法_ > [delete()](../excel/setting.md#delete)|刪除設定。|1.3|
|[settingCollection](../excel/settingcollection.md)|_屬性_ > 項目|設定物件的集合。唯讀。|1.3|
|[settingCollection](../excel/settingcollection.md)|_方法_ > [getItem(key: string)](../excel/settingcollection.md#getitemkey-string)|透過索引鍵取得設定項目。|1.3|
|[settingCollection](../excel/settingcollection.md)|_方法_ > [getItemOrNull(key: string)](../excel/settingcollection.md#getitemornullkey-string)|透過索引鍵取得設定項目。如果設定物件不存在，傳回物件的 isNull 屬性為 true。|1.3|
|[settingCollection](../excel/settingcollection.md)|_方法_ > [set(key: string, value: string)](../excel/settingcollection.md#setkey-string-value-string)|將指定的設定設定或新增至活頁簿。|1.3|
|[settingsChangedEventArgs](../excel/settingschangedeventargs.md)|_關聯性_ > settingCollection|取得代表引發 settingsChanged 事件之繫結的 Settings 物件。|1.3|
|[表格](../excel/table.md)|_屬性_ > highlightFirstColumn|指出第一個資料行是否包含特殊格式。|1.3|
|[表格](../excel/table.md)|_屬性_ > highlightFirstColumn|指出最後一個資料行是否包含特殊格式。|1.3|
|[表格](../excel/table.md)|_屬性_ > showBandedColumns|表示資料行是否顯示帶狀格式，其中奇數的資料行會以不同於偶數資料行的方式反白顯示，讓閱讀資料表更方便。|1.3|
|[表格](../excel/table.md)|_屬性_ > showBandedRows|表示資料列是否顯示帶狀格式，其中奇數的資料列會以不同於偶數資料列的方式反白顯示，讓閱讀資料表更方便。|1.3|
|[表格](../excel/table.md)|_屬性_ > showFilterButton|表示篩選按鈕是否在各個資料行標頭上方可見。只有在資料表包含標頭資料列時允許設定這個選項。|1.3|
|[tableCollection](../excel/tablecollection.md)|_方法_ > [getItemOrNull(key: number or string)](../excel/tablecollection.md#getitemornullkey-number-or-string)|依名稱或識別碼取得資料表。如果資料表不存在，傳回物件的 isNull 屬性為 true。|1.3|
|[tableColumnCollection](../excel/tablecolumncollection.md)|_方法_ > [getItemOrNull(key: number or string)](../excel/tablecolumncollection.md#getitemornullkey-number-or-string)|依名稱或識別碼取得資料行物件。如果資料行物件不存在，傳回物件的 isNull 屬性為 true。|1.3|
|[活頁簿](../excel/workbook.md)|_關聯性_ > 樞紐分析表|代表與活頁簿關聯的樞紐分析表集合。唯讀。|1.3|
|[活頁簿](../excel/workbook.md)|_關聯性_ > 設定|代表與活頁簿關聯的設定集合。唯讀。|1.3|
|[工作表](../excel/worksheet.md)|_關聯性_ > 樞紐分析表|代表屬於活頁簿一部份的樞紐分析表集合。唯讀。|1.3|

## <a name="whats-new-in-excel-javascript-api-12"></a>Excel JavaScript API 1.2 的新功能
以下是需求集合 1.2 中 Excel JavaScript API 的新功能。

|物件| 新功能| 描述|需求集合|
|:----|:----|:----|:----|
|[圖表](../excel/chart.md)|_屬性_ > 識別碼|根據圖表在集合中的位置，取得圖表。唯讀。|1.2|
|[圖表](../excel/chart.md)|_關聯性_ > 工作表|包含目前圖表的工作表。唯讀。|1.2|
|[圖表](../excel/chart.md)|_方法_ > [getImage(height: number, width: number, fittingMode: string)](../excel/chart.md#getimageheight-number-width-number-fittingmode-string)|藉由縮放圖表以符合指定的維度，以 base64 編碼的影像呈現圖表。|1.2|
|[篩選](../excel/filter.md)|_關聯性_ > 準則|目前在指定的欄位上套用的篩選。唯讀。|1.2|
|[篩選](../excel/filter.md)|_方法_ > [apply(criteria:FilterCriteria)](../excel/filter.md#applycriteria-filtercriteria)|在指定的欄位上套用指定的篩選準則。|1.2|
|[篩選](../excel/filter.md)|_方法_ > [applyBottomItemsFilter(count: number)](../excel/filter.md#applybottomitemsfiltercount-number)|套用 [底端項目] 篩選至指定元素數目的欄位。|1.2|
|[篩選](../excel/filter.md)|_方法_ > [applyBottomPercentFilter(percent: number)](../excel/filter.md#applybottompercentfilterpercent-number)|套用 [底部百分比] 篩選至指定元素百分比的欄位。|1.2|
|[篩選](../excel/filter.md)|_方法_ > [applyCellColorFilter(color: string)](../excel/filter.md#applycellcolorfiltercolor-string)|套用 [儲存格色彩] 篩選至指定色彩的欄位。|1.2|
|[篩選](../excel/filter.md)|_方法_ > [applyCustomFilter(criteria1: string, criteria2: string, oper: string)](../excel/filter.md#applycustomfiltercriteria1-string-criteria2-string-oper-string)|套用 [圖示] 篩選至指定準則字串的欄位。|1.2|
|[篩選](../excel/filter.md)|_方法_ > [applyDynamicFilter(criteria: string)](../excel/filter.md#applydynamicfiltercriteria-string)|套用 [動態] 篩選至欄位。|1.2|
|[篩選](../excel/filter.md)|_方法_ > [applyFontColorFilter(color: string)](../excel/filter.md#applyfontcolorfiltercolor-string)|套用 [字型色彩] 篩選至指定色彩的欄位。|1.2|
|[篩選](../excel/filter.md)|_方法_ > [applyIconFilter(icon:Icon)](../excel/filter.md#applyiconfiltericon-icon)|套用 [圖示] 篩選至指定圖示的欄位。|1.2|
|[篩選](../excel/filter.md)|_方法_ > [applyTopItemsFilter(count: number)](../excel/filter.md#applytopitemsfiltercount-number)|套用 [頂端項目] 篩選至指定元素數目的欄位。|1.2|
|[篩選](../excel/filter.md)|_方法_ > [applyTopPercentFilter(percent: number)](../excel/filter.md#applytoppercentfilterpercent-number)|套用 [頂端百分比] 篩選至指定元素百分比的欄位。|1.2|
|[篩選](../excel/filter.md)|_方法_ > [applyValuesFilter(values: ()[])](../excel/filter.md#applyvaluesfiltervalues-)|套用 [值] 篩選至指定值的欄位。|1.2|
|[篩選](../excel/filter.md)|_方法_ > [clear()](../excel/filter.md#clear)|清除指定欄位上的篩選。|1.2|
|[filterCriteria](../excel/filtercriteria.md)|_屬性_ > 色彩|用來篩選儲存格的 HTML 色彩字串。搭配使用 "cellColor" 和 "fontColor" 篩選。|1.2|
|[filterCriteria](../excel/filtercriteria.md)|_屬性_ > criterion1|用來篩選資料的第一個準則。用來做為「自訂」篩選案例中的運算子。|1.2|
|[filterCriteria](../excel/filtercriteria.md)|_屬性_ > criterion2|用來篩選資料的第二個準則。只用來做為「自訂」篩選案例中的運算子。|1.2|
|[filterCriteria](../excel/filtercriteria.md)|_屬性_ > dynamicCriteria|Excel.DynamicFilterCriteria 的動態準則設定為在此資料行上套用。與「動態」篩選搭配使用。可能的值為：未知、AboveAverage、AllDatesInPeriodApril、AllDatesInPeriodAugust、AllDatesInPeriodDecember、AllDatesInPeriodFebruray、AllDatesInPeriodJanuary、AllDatesInPeriodJuly、AllDatesInPeriodJune、AllDatesInPeriodMarch、AllDatesInPeriodMay、AllDatesInPeriodNovember、AllDatesInPeriodOctober、AllDatesInPeriodQuarter1、AllDatesInPeriodQuarter2、AllDatesInPeriodQuarter3、AllDatesInPeriodQuarter4、AllDatesInPeriodSeptember、BelowAverage、LastMonth、LastQuarter、LastWeek、LastYear、NextMonth、NextQuarter、NextWeek、NextYear、ThisMonth、ThisQuarter、ThisWeek、ThisYear、Today、Tomorrow、YearToDate、Yesterday。|1.2|
|[filterCriteria](../excel/filtercriteria.md)|_屬性_ > filterOn|篩選器用於判斷值是否仍看得見的屬性。可能的值為：BottomItems、BottomPercent、CellColor、Dynamic、FontColor、Values、TopItems、TopPercent、Icon、Custom。|1.2|
|[filterCriteria](../excel/filtercriteria.md)|_屬性_ > 運算子|使用「自訂」篩選時，用來結合準則 1 和 2 的運算子。可能的值為：And、Or。|1.2|
|[filterCriteria](../excel/filtercriteria.md)|_屬性_ > 值|要做為「值」篩選部分的值集合。|1.2|
|[filterCriteria](../excel/filtercriteria.md)|_關聯性_ > 圖示|用來篩選儲存格的圖示。與「圖示」篩選搭配使用。|1.2|
|[filterDatetime](../excel/filterdatetime.md)|_屬性_ > 日期|用來篩選資料的 ISO8601 格式的日期。|1.2|
|[filterDatetime](../excel/filterdatetime.md)|_屬性_ > 明確性|保留資料時應該使用多精確的日期。例如，如果日期是 2005-04-02 且明確性設定為「月」，篩選作業會保留日期在 2009 年 4 月份中的所有資料列。可能的值為：年、星期一、日、小時、分鐘、秒。|1.2|
|[formatProtection](../excel/formatprotection.md)|_屬性_ > formulaHidden|表示 Excel 是否在範圍的儲存格中隱藏公式。Null 值表示整個範圍沒有統一公式隱藏設定。|1.2|
|[formatProtection](../excel/formatprotection.md)|_屬性_ > 鎖定|表示 Excel 是否在物件中鎖定儲存格。Null 值表示整個範圍沒有統一鎖定設定。|1.2|
|[圖示](../excel/icon.md)|_屬性_ > 索引|代表指定集合中圖示的索引。|1.2|
|[圖示](../excel/icon.md)|_屬性_ > 設定|代表圖示是其一部分的集合。可能的值為：無效、ThreeArrows、ThreeArrowsGray、ThreeFlags、ThreeTrafficLights1、ThreeTrafficLights2、ThreeSigns、ThreeSymbols、ThreeSymbols2、FourArrows、FourArrowsGray、FourRedToBlack、FourRating、FourTrafficLights、FiveArrows、FiveArrowsGray、FiveRating、FiveQuarters、ThreeStars、ThreeTriangles、FiveBoxes。|1.2|
|[範圍](../excel/range.md)|_屬性_ > columnHidden|表示是否隱藏目前範圍的所有資料行。|1.2|
|[範圍](../excel/range.md)|_屬性_ > formulasR1C1|代表 R1C1 樣式標記法的公式。|1.2|
|[範圍](../excel/range.md)|_屬性_ > 隱藏|表示是否隱藏目前範圍的所有儲存格。唯讀。|1.2|
|[範圍](../excel/range.md)|_屬性_ > rowHidden|表示是否隱藏目前範圍的所有資料列。|1.2|
|[範圍](../excel/range.md)|_關聯性_ > 排序|代表目前範圍的範圍排序。唯讀。|1.2|
|[範圍](../excel/range.md)|_方法_ > [merge(across: bool)](../excel/range.md#mergeacross-bool)|合併範圍儲存格到工作表中的一個區域。|1.2|
|[範圍](../excel/range.md)|_方法_ > [unmerge()](../excel/range.md#unmerge)|取消將範圍儲存格合併至個別儲存格。|1.2|
|[rangeFormat](../excel/rangeformat.md)|_屬性_ > columnWidth|取得或設定範圍內所有資料行寬度。如果資料行寬度不一致，則會傳回 null。|1.2|
|[rangeFormat](../excel/rangeformat.md)|_屬性_ > rowHeight|取得或設定範圍內所有列的高度。如果不是統一的資料列高度，則會傳回 null。|1.2|
|[rangeFormat](../excel/rangeformat.md)|_關聯性_ > 保護|傳回範圍的格式保護物件。唯讀。|1.2|
|[rangeFormat](../excel/rangeformat.md)|_方法_ > [autofitColumns()](../excel/rangeformat.md#autofitcolumns)|根據資料行中的目前資料，變更目前範圍的資料行寬度來調整為最適寬度。|1.2|
|[rangeFormat](../excel/rangeformat.md)|_方法_ > [autofitRows()](../excel/rangeformat.md#autofitrows)|根據資料行中的目前資料，變更目前範圍的資料列高度來調整為最適高度。|1.2|
|[rangeReference](../excel/rangereference.md)|_屬性_ > 地址|代表目前範圍的可見資料列。|1.2|
|[rangeSort](../excel/rangesort.md)|_方法_ > [apply(fields:SortField[], matchCase: bool, hasHeaders: bool, orientation: string, method: string)](../excel/rangesort.md#applyfields-sortfield-matchcase-bool-hasheaders-bool-orientation-string-method-string)|執行排序作業。|1.2|
|[sortField](../excel/sortfield.md)|_屬性_ > 遞增|表示是否以遞增方式完成排序。|1.2|
|[sortField](../excel/sortfield.md)|_屬性_ > 色彩|如果排序是針對字型或儲存格色彩，表示色彩是條件的目標。|1.2|
|[sortField](../excel/sortfield.md)|_屬性_ > dataOption|表示此欄位額外的排序選項。可能的值為：一般、TextAsNumber。|1.2|
|[sortField](../excel/sortfield.md)|_屬性_ > 索引鍵|表示套用條件的資料行 (或資料列，視排序的方向而定)。表示為從第一個資料行 (或資料列) 的位移。|1.2|
|[sortField](../excel/sortfield.md)|_屬性_ > sortOn|表示這個條件的排序的類型。可能的值為：值、CellColor、FontColor、圖示。|1.2|
|[sortField](../excel/sortfield.md)|_關聯性_ > 圖示|如果排序是針對儲存格的圖示，表示圖示是條件的目標。|1.2|
|[表格](../excel/table.md)|_關聯性_ > 排序|代表資料表的排序。唯讀。|1.2|
|[表格](../excel/table.md)|_關聯性_ > 工作表|包含目前資料表的工作表。唯讀。|1.2|
|[表格](../excel/table.md)|_方法_ > [clearFilters()](../excel/table.md#clearfilters)|清除目前在資料表上套用的所有篩選器。|1.2|
|[表格](../excel/table.md)|_方法_ > [convertToRange()](../excel/table.md#converttorange)|將資料表轉換成一般儲存格範圍。所有的資料會保留。|1.2|
|[表格](../excel/table.md)|_方法_ > [reapplyFilters()](../excel/table.md#reapplyfilters)|重新套用目前在資料表上的所有篩選器。|1.2|
|[tableColumn](../excel/tablecolumn.md)|_關聯性_ > 篩選|擷取套用至資料行的篩選器。唯讀。|1.2|
|[tableSort](../excel/tablesort.md)|_屬性_ > matchCase|表示大小寫會影響料表的最後排序。唯讀。|1.2|
|[tableSort](../excel/tablesort.md)|_屬性_ > 方法|表示最後用於排序資料表的中文字元排序方法。唯讀。可能的值為：拼音、StrokeCount。|1.2|
|[tableSort](../excel/tablesort.md)|_關聯性_ > 欄位|表示用於最後排序資料表的目前條件。唯讀。|1.2|
|[tableSort](../excel/tablesort.md)|_方法_ > [apply(fields:SortField[], matchCase: bool, method: string)](../excel/tablesort.md#applyfields-sortfield-matchcase-bool-method-string)|執行排序作業。|1.2|
|[tableSort](../excel/tablesort.md)|_方法_ > [clear()](../excel/tablesort.md#clear)|清除資料表上目前的排序。雖然這不會修改資料表的順序，它會清除標頭按鈕的狀態。|1.2|
|[tableSort](../excel/tablesort.md)|_方法_ > [reapply()](../excel/tablesort.md#reapply)|將目前的排序參數重新套用至資料表。|1.2|
|[活頁簿](../excel/workbook.md)|_關聯性_ > 函式|代表包含此活頁簿的 Excel 應用程式執行個體。唯讀。|1.2|
|[工作表](../excel/worksheet.md)|_關聯性_ > 保護|傳回工作表的工作表保護物件。唯讀。|1.2|
|[worksheetProtection](../excel/worksheetprotection.md)|_屬性_ > 保護|表示工作表是否受到保護。唯讀。唯讀。|1.2|
|[worksheetProtection](../excel/worksheetprotection.md)|_關聯性_ > 選項|工作表保護選項。唯讀。|1.2|
|[worksheetProtection](../excel/worksheetprotection.md)|_方法_ > [protect(options:WorksheetProtectionOptions)](../excel/worksheetprotection.md#protectoptions-worksheetprotectionoptions)|保護工作表。如果工作表已經受到保護，則會失敗。|1.2|
|[worksheetProtection](../excel/worksheetprotection.md)|_方法_ > [unprotect()](../excel/worksheetprotection.md#unprotect)|取消保護工作表。|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_屬性_ > allowAutoFilter|代表工作表保護選項，允許使用自動篩選功能。|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_屬性_ > allowDeleteColumns|代表工作表保護選項，允許刪除資料行。|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_屬性_ > allowDeleteRows|代表工作表保護選項，允許刪除資料列。|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_屬性_ > allowFormatCells|代表工作表保護選項，允許格式化儲存格。|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_屬性_ > allowFormatColumns|代表工作表保護選項，允許格式化資料行。|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_屬性_ > allowFormatRows|代表工作表保護選項，允許格式化資料列。|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_屬性_ > allowInsertColumns|代表工作表保護選項，允許插入資料行。|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_屬性_ > allowInsertHyperlinks|代表工作表保護選項，允許插入超連結。|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_屬性_ > allowInsertRows|代表工作表保護選項，允許插入資料列。|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_屬性_ > allowPivotTables|代表工作表保護選項，允許使用樞紐分析表功能。|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_屬性_ > allowSort|代表工作表保護選項，允許使用排序功能。|1.2|

## <a name="excel-javascript-api-11"></a>Excel JavaScript API 1.1
Excel JavaScript API 1.1 是 API 的第一個版本。如需 API 的詳細資訊，請參閱 Excel JavaScript API 參考主題。  

## <a name="additional-resources"></a>其他資源

- [指定 Office 主應用程式和 API 需求](../../docs/overview/specify-office-hosts-and-api-requirements.md)
- [Office 增益集的 XML 資訊清單](../../docs/overview/add-in-manifests.md)
