# <a name="filter-object-(javascript-api-for-excel)"></a>篩選物件 (適用於 Excel 的 JavaScript API)

_適用於：Excel 2016、Excel Online、Excel for iOS、Office 2016_

管理表格欄位的篩選。

## <a name="properties"></a>屬性

無

## <a name="relationships"></a>關聯性
| 關聯性 | 類型	   |描述|
|:---------------|:--------|:----------|
|criteria|[FilterCriteria](filtercriteria.md)|目前在指定的欄位上套用的篩選。唯讀。|

## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述|
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

## <a name="method-details"></a>方法詳細資料


### <a name="apply(criteria:-filtercriteria)"></a>apply(criteria:FilterCriteria)
在指定的欄位上套用指定的篩選準則。使用任何下列的 helper 方法就可以達成相同的功能。 

#### <a name="syntax"></a>語法
```js
filterObject.apply(criteria);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|準則|FilterCriteria|要套用的準則。|

#### <a name="returns"></a>傳回
void

#### <a name="example"></a>範例
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

### <a name="applybottomitemsfilter(count:-number)"></a>applyBottomItemsFilter(count: number)
套用 [底端項目] 篩選至指定元素數目的欄位。

#### <a name="syntax"></a>語法
```js
filterObject.applyBottomItemsFilter(count);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|Count|number|從下至上顯示的元素數目。|

#### <a name="returns"></a>傳回
void

#### <a name="example"></a>範例
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

### <a name="applybottompercentfilter(percent:-number)"></a>applyBottomPercentFilter(percent: number)
套用 [底部百分比] 篩選至指定元素百分比的欄位。

#### <a name="syntax"></a>語法
```js
filterObject.applyBottomPercentFilter(percent);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|百分比|number|從下至上顯示的元素百分比。|

#### <a name="returns"></a>傳回
void

#### <a name="example"></a>範例
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
### <a name="applycellcolorfilter(color:-string)"></a>applyCellColorFilter(color: string)
套用 [儲存格色彩] 篩選至指定色彩的欄位。


#### <a name="syntax"></a>語法
```js
filterObject.applyCellColorFilter(color);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|Color|string|顯示的儲存格背景色彩。|

#### <a name="returns"></a>傳回
void

#### <a name="example"></a>範例
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

### <a name="applycustomfilter(criteria1:-string,-criteria2:-string,-oper:-filteroperator)"></a>applyCustomFilter(criteria1: string, criteria2: string, oper:FilterOperator)
套用 [圖示] 篩選至指定準則字串的欄位。

#### <a name="syntax"></a>語法
```js
filterObject.applyCustomFilter(criteria1, criteria2, oper);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|criteria1|string|第一個準則字串。|
|criteria2|字串|選用。第二個準則字串。|
|運算子|FilterOperator|選用。說明如何聯結兩個準則的運算子。|

#### <a name="returns"></a>傳回
void


#### <a name="example"></a>範例
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

### <a name="applydynamicfilter(criteria:-string)"></a>applyDynamicFilter(criteria: string)
套用 [動態] 篩選至欄位。

#### <a name="syntax"></a>語法
```js
filterObject.applyDynamicFilter(criteria);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|criteria|字串|要套用的動態準則。可能的值為：未知、AboveAverage、AllDatesInPeriodApril、AllDatesInPeriodAugust、AllDatesInPeriodDecember、AllDatesInPeriodFebruray、AllDatesInPeriodJanuary、AllDatesInPeriodJuly、AllDatesInPeriodJune、AllDatesInPeriodMarch、AllDatesInPeriodMay、AllDatesInPeriodNovember、AllDatesInPeriodOctober、AllDatesInPeriodQuarter1、AllDatesInPeriodQuarter2、AllDatesInPeriodQuarter3、AllDatesInPeriodQuarter4、AllDatesInPeriodSeptember、BelowAverage、LastMonth、LastQuarter、LastWeek、LastYear、NextMonth、NextQuarter、NextWeek、NextYear、ThisMonth、ThisQuarter、ThisWeek、ThisYear、Today、Tomorrow、YearToDate、Yesterday|

#### <a name="returns"></a>傳回
void

#### <a name="example"></a>範例
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

### <a name="applyfontcolorfilter(color:-string)"></a>applyFontColorFilter(color: string)
套用 [字型色彩] 篩選至指定色彩的欄位。

#### <a name="syntax"></a>語法
```js
filterObject.applyFontColorFilter(color);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|Color|string|顯示的儲存格字型色彩。|

#### <a name="returns"></a>傳回
void

#### <a name="example"></a>範例
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

### <a name="applyiconfilter(icon:-icon)"></a>applyIconFilter(icon:Icon)
套用 [圖示] 篩選至指定圖示的欄位。

#### <a name="syntax"></a>語法
```js
filterObject.applyIconFilter(icon);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|圖示|圖示|顯示的儲存格圖示。|

#### <a name="returns"></a>傳回
void

#### <a name="example"></a>範例
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

### <a name="applytopitemsfilter(count:-number)"></a>applyTopItemsFilter(count: number)
套用 [頂端項目] 篩選至指定元素數目的欄位。

#### <a name="syntax"></a>語法
```js
filterObject.applyTopItemsFilter(count);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|Count|number|從上至下顯示的元素數目。|

#### <a name="returns"></a>傳回
void

#### <a name="example"></a>範例
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


### <a name="applytoppercentfilter(percent:-number)"></a>applyTopPercentFilter(percent: number)
套用 [頂端百分比] 篩選至指定元素百分比的欄位。

#### <a name="syntax"></a>語法
```js
filterObject.applyTopPercentFilter(percent);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|百分比|number|從上至下顯示的元素百分比。|

#### <a name="returns"></a>傳回
void

#### <a name="example"></a>範例
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
### <a name="applyvaluesfilter(values:-()[])"></a>applyValuesFilter(values: ()[])
套用 [值] 篩選至指定值的欄位。

#### <a name="syntax"></a>語法
```js
filterObject.applyValuesFilter(values);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|values|()[]|顯示的值清單。|

#### <a name="returns"></a>傳回
void

#### <a name="example"></a>範例
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

### <a name="clear()"></a>clear()
清除指定欄位上的篩選。

#### <a name="syntax"></a>語法
```js
filterObject.clear();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
void

#### <a name="example"></a>範例
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

### <a name="load(param:-object)"></a>load(param: object)
以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。

#### <a name="syntax"></a>語法
```js
object.load(param);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|param|物件|選用。接受參數與關聯性名稱，做為分隔字串或陣列。或者提供 [loadOption](loadoption.md) 物件。|

#### <a name="returns"></a>傳回
void
