# 篩選物件 (適用於 Excel 的 JavaScript API)

_適用版本：Excel 2016、Excel Online、Excel for iOS、Office 2016_

管理表格欄位的篩選。

## 屬性

無

## 關聯性
| 關聯性 | 類型	|說明|
|:---------------|:--------|:----------|
|準則|[FilterCriteria](filtercriteria.md)|目前在指定的欄位上套用的篩選。唯讀。|

## 方法

| 方法	   | 傳回類型	|說明|
|:---------------|:--------|:----------|
|[apply(criteria:FilterCriteria)](#applycriteria-filtercriteria)|void|在指定的欄位上套用指定的篩選準則。使用任何下列的 helper 方法就可以達成相同的功能。|
|[applyBottomItemsFilter(count: number)](#applybottomitemsfiltercount-number)|void|套用 [底端項目] 篩選至指定元素數目的欄位。|
|[applyBottomPercentFilter(percent: number)](#applybottompercentfilterpercent-number)|void|套用 [底部百分比] 篩選至指定元素百分比的欄位。|
|[applyCellColorFilter(color: string)](#applycellcolorfiltercolor-string)|void|套用 [儲存格色彩] 篩選至指定色彩的欄位。|
|[applyCustomFilter(criteria1: string, criteria2: string, oper:FilterOperator)](#applycustomfiltercriteria1-string-criteria2-string-oper-filteroperator)|void|套用 [圖示] 篩選至指定準則字串的欄位。|
|[applyDynamicFilter(criteria: string)](#applydynamicfiltercriteria-string)|void|套用 [動態] 篩選至欄位。|
|[applyFontColorFilter(color: string)](#applyfontcolorfiltercolor-string)|void|套用 [字型色彩] 篩選至指定色彩的欄位。|
|[applyIconFilter(icon:Icon)](#applyiconfiltericon-icon)|void|套用 [圖示] 篩選至指定圖示的欄位。|
|[applyTopItemsFilter(count: number)](#applytopitemsfiltercount-number)|void|套用 [頂端項目] 篩選至指定元素數目的欄位。|
|[applyTopPercentFilter(percent: number)](#applytoppercentfilterpercent-number)|void|套用 [頂端百分比] 篩選至指定元素百分比的欄位。|
|[applyValuesFilter(values: ()[])](#applyvaluesfiltervalues-)|void|套用 [值] 篩選至指定值的欄位。|
|[clear()](#clear)|void|清除指定欄位上的篩選。|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|

## 方法詳細資料


### apply(criteria:FilterCriteria)
在指定的欄位上套用指定的篩選準則。使用任何下列的 helper 方法就可以達成相同的功能。 

#### 語法
```js
filterObject.apply(criteria);
```

#### 參數
| 參數	   | 類型	|說明|
|:---------------|:--------|:----------|
|準則|FilterCriteria|要套用的準則。|

#### 傳回
void

#### 範例
下列範例證明如何使用一般 apply() 方法套用自訂篩選。

```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    var filterCriteria = { 
		filterOn: Excel.FilterOn.custom,
		criterion1: ">50",
		operator: Excel.FilterOperator.and,
		criterion2: "<100"
    	} 
    column.filter.apply(filterCriteria);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### applyBottomItemsFilter(count: number)
套用 [底端項目] 篩選至指定元素數目的欄位。

#### 語法
```js
filterObject.applyBottomItemsFilter(count);
```

#### 參數
| 參數	   | 類型	|說明|
|:---------------|:--------|:----------|
|Count|number|從下至上顯示的元素數目。|

#### 傳回
void

#### 範例
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyBottomItemsFilter(3);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### applyBottomPercentFilter(percent: number)
套用 [底部百分比] 篩選至指定元素百分比的欄位。

#### 語法
```js
filterObject.applyBottomPercentFilter(percent);
```

#### 參數
| 參數	   | 類型	|說明|
|:---------------|:--------|:----------|
|百分比|number|從下至上顯示的元素百分比。|

#### 傳回
void

#### 範例
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyBottomPercentFilter(30);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
### applyCellColorFilter(color: string)
套用 [儲存格色彩] 篩選至指定色彩的欄位。


#### 語法
```js
filterObject.applyCellColorFilter(color);
```

#### 參數
| 參數	   | 類型	|說明|
|:---------------|:--------|:----------|
|Color|string|顯示的儲存格背景色彩。|

#### 傳回
void

#### 範例
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyCellColorFilter('red');
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### applyCustomFilter(criteria1: string, criteria2: string, oper:FilterOperator)
套用 [圖示] 篩選至指定準則字串的欄位。

#### 語法
```js
filterObject.applyCustomFilter(criteria1, criteria2, oper);
```

#### 參數
| 參數	   | 類型	|說明|
|:---------------|:--------|:----------|
|criteria1|string|第一個準則字串。|
|criteria2|string|選用。第二個準則字串。|
|運算子|FilterOperator|選用。說明如何聯結兩個準則的運算子。|

#### 傳回
void


#### 範例
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyCustomFilter('>50','<100','and');
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### applyDynamicFilter(criteria: string)
套用 [動態] 篩選至欄位。

#### 語法
```js
filterObject.applyDynamicFilter(criteria);
```

#### 參數
| 參數	   | 類型	|說明|
|:---------------|:--------|:----------|
|criteria|string|要套用的動態準則。可能的值為：未知、AboveAverage、AllDatesInPeriodApril、AllDatesInPeriodAugust、AllDatesInPeriodDecember、AllDatesInPeriodFebruray、AllDatesInPeriodJanuary、AllDatesInPeriodJuly、AllDatesInPeriodJune、AllDatesInPeriodMarch、AllDatesInPeriodMay、AllDatesInPeriodNovember、AllDatesInPeriodOctober、AllDatesInPeriodQuarter1、AllDatesInPeriodQuarter2、AllDatesInPeriodQuarter3、AllDatesInPeriodQuarter4、AllDatesInPeriodSeptember、BelowAverage、LastMonth、LastQuarter、LastWeek、LastYear、NextMonth、NextQuarter、NextWeek、NextYear、ThisMonth、ThisQuarter、ThisWeek、ThisYear、Today、Tomorrow、YearToDate、Yesterday|

#### 傳回
void

#### 範例
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyDynamicFilter(Excel.DynamicFilterCriteria.aboveAverage);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### applyFontColorFilter(color: string)
套用 [字型色彩] 篩選至指定色彩的欄位。

#### 語法
```js
filterObject.applyFontColorFilter(color);
```

#### 參數
| 參數	   | 類型	|說明|
|:---------------|:--------|:----------|
|Color|string|顯示的儲存格字型色彩。|

#### 傳回
void

#### 範例
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyFontColorFilter('red');
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### applyIconFilter(icon:Icon)
套用 [圖示] 篩選至指定圖示的欄位。

#### 語法
```js
filterObject.applyIconFilter(icon);
```

#### 參數
| 參數	   | 類型	|說明|
|:---------------|:--------|:----------|
|圖示|圖示|顯示的儲存格圖示。|

#### 傳回
void

#### 範例
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyIconFilter(Excel.icons.fiveArrows.yellowDownInclineArrow);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### applyTopItemsFilter(count: number)
套用 [頂端項目] 篩選至指定元素數目的欄位。

#### 語法
```js
filterObject.applyTopItemsFilter(count);
```

#### 參數
| 參數	   | 類型	|說明|
|:---------------|:--------|:----------|
|Count|number|從上至下顯示的元素數目。|

#### 傳回
void

#### 範例
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyTopItemsFilter(3);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### applyTopPercentFilter(percent: number)
套用 [頂端百分比] 篩選至指定元素百分比的欄位。

#### 語法
```js
filterObject.applyTopPercentFilter(percent);
```

#### 參數
| 參數	   | 類型	|說明|
|:---------------|:--------|:----------|
|百分比|number|從上至下顯示的元素百分比。|

#### 傳回
void

#### 範例
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyTopPercentFilter(30);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
### applyValuesFilter(values: ()[])
套用 [值] 篩選至指定值的欄位。

#### 語法
```js
filterObject.applyValuesFilter(values);
```

#### 參數
| 參數	   | 類型	|說明|
|:---------------|:--------|:----------|
|values|()[]|顯示的值清單。|

#### 傳回
void

#### 範例
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyValuesFilter(['a','b']);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### clear()
清除指定欄位上的篩選。

#### 語法
```js
filterObject.clear();
```

#### 參數
無

#### 傳回
void

#### 範例
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.clear();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### load(param: object)
以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。

#### 語法
```js
object.load(param);
```

#### 參數
| 參數	   | 類型	|說明|
|:---------------|:--------|:----------|
|param|物件|選用。接受參數與關聯性名稱，做為分隔字串或陣列。或者提供 [loadOption](loadoption.md) 物件。|

#### 傳回
void

