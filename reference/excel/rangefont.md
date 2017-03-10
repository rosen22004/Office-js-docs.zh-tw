# <a name="rangefont-object-javascript-api-for-excel"></a>RangeFont 物件 (適用於 Excel 的 JavaScript API)

此物件代表物件的字型屬性 (字型名稱、字型大小、色彩等)。

## <a name="properties"></a>屬性

| 屬性	       | 類型	    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|bold|bool|代表字型的粗體狀態。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|color|string|代表文字色彩的 HTML 色彩代碼。例如 #FF0000 代表紅色。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|italic|bool|代表字型的斜體狀態。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|字型名稱 (例如 "Calibri")|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|size|雙精確度|字型大小。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|underline|string|套用至字型的底線類型。可能的值為：None、Single、Double、SingleAccountant、DoubleAccountant。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
無


## <a name="methods"></a>方法
無


## <a name="method-details"></a>方法詳細資料

### <a name="property-access-examples"></a>屬性存取範例

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "F:G";
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var range = worksheet.getRange(rangeAddress);
    var rangeFont = range.format.font;
    rangeFont.load('name');
    return ctx.sync().then(function() {
        console.log(rangeFont.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
下列範例會設定字型名稱。 

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "F:G";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.format.font.name = 'Times New Roman';
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```