# <a name="charttitle-object-javascript-api-for-excel"></a>ChartTitle 物件 (適用於 Excel 的 JavaScript API)

代表圖表的圖表標題物件。

## <a name="properties"></a>屬性

| 屬性	       | 類型	    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|overlay|bool|布林值，表示圖表標題是否覆蓋圖表。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|text|string|代表圖表的標題文字。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|visible|bool|布林值，代表圖表標題物件的可見性。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
| 關聯性 | 類型	    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|format|[ChartTitleFormat](charttitleformat.md)|代表圖表標題的格式設定，其中包含填滿和字型格式。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>方法
無


## <a name="method-details"></a>方法詳細資料

### <a name="property-access-examples"></a>屬性存取範例

從 Chart1 取得圖表標題的 `text`。

```js
Excel.run(function (ctx) { 
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");    

var title = chart.title;
title.load('text');
return ctx.sync().then(function() {
        console.log(title.text);
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
});
```

將圖表標題的 `text` 設定為 "My Chart"，並讓它顯示在圖表頂端且不重疊。

```js
Excel.run(function (ctx) { 
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");    

chart.title.text= "My Chart"; 
chart.title.visible=true;
chart.title.overlay=true;

return ctx.sync().then(function() {
        console.log("Char Title Changed");
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
});
```
