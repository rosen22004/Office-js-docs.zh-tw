# <a name="chartlegend-object-javascript-api-for-excel"></a>ChartLegend 物件 (適用於 Excel 的 JavaScript API)

代表圖表中的圖例。

## <a name="properties"></a>屬性

| 屬性	     | 類型	   |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|overlay|bool|布林值，指出圖表圖例是否應與圖表主體重疊。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|position|string|代表圖例在圖表上的位置。可能的值為：Top、Bottom、Left、Right、Corner、Custom。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|visible|bool|布林值，代表 ChartLegend 物件的可見性。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
| 關聯性 | 類型	   |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|format|[ChartLegendFormat](chartlegendformat.md)|代表圖表圖例的格式設定，其中包含填滿和字型格式。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法詳細資料


### <a name="loadparam-object"></a>load(param: object)
以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。

#### <a name="syntax"></a>語法
```js
object.load(param);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|:---|
|param|物件|選用。接受參數與關聯性名稱，做為分隔字串或陣列。或者提供 [loadOption](loadoption.md) 物件。|

#### <a name="returns"></a>傳回
void
### <a name="property-access-examples"></a>屬性存取範例

從 Chart1 取得圖表圖例的 `position`

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1"); 
    var legend = chart.legend;
    legend.load('position');
    return ctx.sync().then(function() {
            console.log(legend.position);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

設定為顯示 Chart1 的圖例，並將它設定在圖表的頂端。

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1"); 
    chart.legend.visible = true;
    chart.legend.position = "top"; 
    chart.legend.overlay = false; 
    return ctx.sync().then(function() {
            console.log("Legend Shown ");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
``` 
