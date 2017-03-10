# <a name="chartlineformat-object-javascript-api-for-excel"></a>ChartLineFormat 物件 (適用於 Excel 的 JavaScript API)

封裝線條元素的格式設定選項。

## <a name="properties"></a>屬性

| 屬性	       | 類型	    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|Color|string|代表圖表中線條色彩的 HTML 色彩代碼。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
無


## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|[clear()](#clear)|void|清除圖表項目的線條格式。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法詳細資料


### <a name="clear"></a>clear()
清除圖表項目的線條格式。

#### <a name="syntax"></a>語法
```js
chartLineFormatObject.clear();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
void

#### <a name="examples"></a>範例

清除名為 "Chart1" 之圖表的數值軸上主要格線的線條格式

```js
Excel.run(function (ctx) { 
    var gridlines = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").axes.valueAxis.majorGridlines;    
    gridlines.format.line.clear();
    return ctx.sync().then(function() {
            console.log("Chart Major Gridlines Format Cleared");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
### <a name="property-access-examples"></a>屬性存取範例

將圖表數值軸上的主要格線設定為紅色。

```js
Excel.run(function (ctx) {
    var gridlines = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").axes.valueAxis.majorGridlines;
    gridlines.format.line.color = "#FF0000";
    return ctx.sync().then(function () {
        console.log("Chart Gridlines Color Updated");
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```
