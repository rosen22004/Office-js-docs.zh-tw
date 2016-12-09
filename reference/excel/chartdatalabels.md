# <a name="chartdatalabels-object-javascript-api-for-excel"></a>ChartDataLabels 物件 (適用於 Excel 的 JavaScript API)

代表圖表點上所有資料標籤的集合。

## <a name="properties"></a>屬性

| 屬性	     | 類型	   |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|position|string|DataLabelPosition 值，代表資料標籤的位置。可能的值為：None、Center、InsideEnd、InsideBase、OutsideEnd、Left、Right、Top、Bottom、BestFit、Callout。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|separator|string|代表圖表上資料標籤所使用分隔符號的字串。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showBubbleSize|bool|布林值，代表資料標籤的泡泡大小是否可見。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showCategoryName|bool|布林值，代表資料標籤的類別名稱是否可見。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showLegendKey|bool|布林值，代表資料標籤的圖例符號是否可見。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showPercentage|bool|布林值，代表資料標籤的百分比是否可見。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showSeriesName|bool|布林值，代表資料標籤的數列名稱是否可見。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showValue|bool|布林值，代表資料標籤的值是否可見。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
| 關聯性 | 類型	   |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|format|[ChartDataLabelFormat](chartdatalabelformat.md)|代表圖表資料標籤的格式，其中包含填滿和字型格式。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

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

讓數列名稱顯示在資料標籤中，並設定資料標籤的 `position` 為 "top"；

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1"); 
    chart.datalabels.showValue = true;
    chart.datalabels.position = "top";
    chart.datalabels.showSeriesName = true;
    return ctx.sync().then(function() {
            console.log("Datalabels Shown");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
