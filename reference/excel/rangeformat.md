# <a name="rangeformat-object-javascript-api-for-excel"></a>RangeFormat 物件 (適用於 Excel 的 JavaScript API)

格式物件，其中封裝了範圍的字型、填滿、框線、對齊方式及其他屬性。

## <a name="properties"></a>屬性

| 屬性	       | 類型	    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|columnWidth|雙精確度|取得或設定範圍內所有資料行寬度。如果資料行寬度不一致，則會傳回 null。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|horizontalAlignment|string|代表所指定物件的水平對齊方式。可能的值為：General、Left、Center、Right、Fill、Justify、CenterAcrossSelection、Distributed。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|rowHeight|雙精確度|取得或設定範圍內所有列的高度。如果不是統一的資料列高度，則會傳回 null。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|verticalAlignment|string|代表所指定物件的垂直對齊方式。可能的值為：Top、Center、Bottom、Justify、Distributed。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|wrapText|bool|表示 Excel 是否在物件中自動換行。Null 值表示整個範圍沒有統一換行設定。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
| 關聯性 | 類型	    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|borders|[RangeBorderCollection](rangebordercollection.md)|套用至選定之整體範圍的 border 物件集合。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|fill|[RangeFill](rangefill.md)|傳回整體範圍中定義的 fill 物件。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|font|[RangeFont](rangefont.md)|傳回整體範圍中定義的 font 物件。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|protection|[FormatProtection](formatprotection.md)|傳回範圍的格式保護物件。唯讀。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|[autofitColumns()](#autofitcolumns)|void|根據資料行中的目前資料，變更目前範圍的資料行寬度來調整為最適寬度。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[autofitRows()](#autofitrows)|void|根據資料行中的目前資料，變更目前範圍的資料列高度來調整為最適高度。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法詳細資料


### <a name="autofitcolumns"></a>autofitColumns()
根據資料行中的目前資料，變更目前範圍的資料行寬度來調整為最適寬度。

#### <a name="syntax"></a>語法
```js
rangeFormatObject.autofitColumns();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
void

### <a name="autofitrows"></a>autofitRows()
根據資料行中的目前資料，變更目前範圍的資料列高度來調整為最適高度。

#### <a name="syntax"></a>語法
```js
rangeFormatObject.autofitRows();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
void
### <a name="property-access-examples"></a>屬性存取範例

下列範例將選取所有範圍格式屬性。 

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "F:G";
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var range = worksheet.getRange(rangeAddress);
    range.load(["format/*", "format/fill", "format/borders", "format/font"]);
    return ctx.sync().then(function() {
        console.log(range.format.wrapText);
        console.log(range.format.fill.color);
        console.log(range.format.font.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

以下範例會設定字型名稱、填滿色彩，並自動換行。 

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "F:G";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.format.wrapText = true;
    range.format.font.name = 'Times New Roman';
    range.format.fill.color = '0000FF';
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

以下範例會在範圍周圍增加格線框線。

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "F:G";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.format.borders.getItem('InsideHorizontal').style = 'Continuous';
    range.format.borders.getItem('InsideVertical').style = 'Continuous';
    range.format.borders.getItem('EdgeBottom').style = 'Continuous';
    range.format.borders.getItem('EdgeLeft').style = 'Continuous';
    range.format.borders.getItem('EdgeRight').style = 'Continuous';
    range.format.borders.getItem('EdgeTop').style = 'Continuous';
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```