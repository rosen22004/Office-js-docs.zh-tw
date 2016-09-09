# RangeFormat 物件 (適用於 Excel 的 JavaScript API)

格式物件，其中封裝了範圍的字型、填滿、框線、對齊方式及其他屬性。

## 屬性

| 屬性	     | 類型	   |說明
|:---------------|:--------|:----------|
|columnWidth|雙精確度|取得或設定範圍內所有資料行寬度。如果資料行寬度不一致，則會傳回 null。|
|horizontalAlignment|string|代表所指定物件的水平對齊方式。可能的值為：General、Left、Center、Right、Fill、Justify、CenterAcrossSelection、Distributed。|
|rowHeight|雙精確度|取得或設定範圍內所有列的高度。如果不是統一的資料列高度，則會傳回 null。|
|verticalAlignment|string|代表所指定物件的垂直對齊方式。可能的值為：Top、Center、Bottom、Justify、Distributed。|
|wrapText|bool|指出 Excel 文字控制項已設定為在物件中自動換行。Null 值表示整個範圍不使用統一的自動換行文字設定。|

_請參閱屬性存取[範例。](#範例)_

## 關聯性
| 關聯性 | 類型	   |說明|
|:---------------|:--------|:----------|
|borders|[RangeBorderCollection](rangebordercollection.md)|套用至選定之整體範圍的 border 物件集合。唯讀。|
|fill|[RangeFill](rangefill.md)|傳回整體範圍中定義的 fill 物件。唯讀。|
|font|[RangeFont](rangefont.md)|傳回選定之整體範圍中定義的 font 物件。唯讀。|
|保護|[FormatProtection](formatprotection.md)|傳回範圍的格式保護物件。唯讀。|

## 方法

| 方法           | 傳回類型    |說明|
|:---------------|:--------|:----------|
|[autofitColumns()](#autofitcolumns)|void|根據資料行中的目前資料，變更目前範圍的資料行寬度來調整為最適寬度。|
|[autofitRows()](#autofitrows)|void|根據資料行中的目前資料，變更目前範圍的資料列高度來調整為最適高度。|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|

## 方法詳細資料


### autofitColumns()
根據資料行中的目前資料，變更目前範圍的資料行寬度來調整為最適寬度。

#### 語法
```js
rangeFormatObject.autofitColumns();
```

#### 參數
無

#### 傳回
void

### autofitRows()
根據資料行中的目前資料，變更目前範圍的資料列高度來調整為最適高度。

#### 語法
```js
rangeFormatObject.autofitRows();
```

#### 參數
無

#### 傳回
void

### load(param: object)
以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。

#### 語法
```js
object.load(param);
```

#### 參數
| 參數	    | 類型	   |說明|
|:---------------|:--------|:----------|
|param|物件|選用。接受參數與關聯性名稱，做為分隔字串或陣列。或者提供 [loadOption](loadoption.md) 物件。|

#### 傳回
void
### 屬性存取範例

此範例將列印一個範圍的所有格式屬性。 

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

以下範例會設定一個範圍的字型名稱及填滿色彩，並自動換行。 

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
    range.format.borders('InsideHorizontal').lineStyle = 'Continuous';
    range.format.borders('InsideVertical').lineStyle = 'Continuous';
    range.format.borders('EdgeBottom').lineStyle = 'Continuous';
    range.format.borders('EdgeLeft').lineStyle = 'Continuous';
    range.format.borders('EdgeRight').lineStyle = 'Continuous';
    range.format.borders('EdgeTop').lineStyle = 'Continuous';
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
