# <a name="workbook-object-javascript-api-for-excel"></a>Workbook 物件 (適用於 Excel 的 JavaScript API)

活頁簿是最上層物件，其中包含相關的活頁簿物件，例如工作表、表格、範圍等等。

## <a name="properties"></a>屬性

無

## <a name="relationships"></a>關聯性
| 關聯性 | 類型	    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|應用程式|[Application](application.md)|代表包含此活頁簿的 Excel 應用程式執行個體。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|bindings|[BindingCollection](bindingcollection.md)|代表屬於活頁簿一部份的繫結集合。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|functions|[Functions](functions.md)|代表包含此活頁簿的 Excel 應用程式執行個體。唯讀。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|names|[NamedItemCollection](nameditemcollection.md)|代表活頁簿限定範圍具名項目 (具名的範圍和常數) 的集合。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|pivotTables|[PivotTableCollection](pivottablecollection.md)|代表與活頁簿關聯的樞紐分析表集合。唯讀。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|settings|[SettingCollection](settingcollection.md)|代表與活頁簿關聯的設定集合。唯讀。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|表格|[TableCollection](tablecollection.md)|代表與活頁簿關聯的表格集合。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|worksheets|[WorksheetCollection](worksheetcollection.md)|代表與活頁簿關聯的工作表集合。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|[getSelectedRange()](#getselectedrange)|[Range](range.md)|從活頁簿取得目前選取的範圍。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法詳細資料


### <a name="getselectedrange"></a>getSelectedRange()
從活頁簿取得目前選取的範圍。

#### <a name="syntax"></a>語法
```js
workbookObject.getSelectedRange();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
[Range](range.md)

#### <a name="examples"></a>範例

```js
Excel.run(function (ctx) { 
    var selectedRange = ctx.workbook.getSelectedRange();
    selectedRange.load('address');
    return ctx.sync().then(function() {
            console.log(selectedRange.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```