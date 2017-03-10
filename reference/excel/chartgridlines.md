# <a name="chartgridlines-object-javascript-api-for-excel"></a>ChartGridlines 物件 (適用於 Excel 的 JavaScript API)

代表圖表座標軸上的主要或次要格線。

## <a name="properties"></a>屬性

| 屬性	       | 類型	    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|visible|bool|布林值，代表座標軸格線是否可見。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
| 關聯性 | 類型	    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|format|[ChartGridlinesFormat](chartgridlinesformat.md)|代表圖表格線的格式。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>方法
無


## <a name="method-details"></a>方法詳細資料

### <a name="property-access-examples"></a>屬性存取範例

取得 Chart1 數值軸主要格線的 `visible`。

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");    
    var majGridlines = chart.axes.valueaxis.majorGridlines;
    majGridlines.load('visible');
    return ctx.sync().then(function() {
            console.log(majGridlines.visible);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

設定為顯示 Chart1 數值軸的主要格線。

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");    
    chart.axes.valueAxis.majorGridlines.visible = true;
    return ctx.sync().then(function() {
            console.log("Axis Gridlines Added ");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
