# <a name="filter-object-javascript-api-for-excel"></a>Filter 物件 (適用於 Excel 的 JavaScript API)

管理表格欄位的篩選。

## <a name="properties"></a>屬性

無

## <a name="relationships"></a>關聯性
| 關聯性 | 類型	    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|準則|[FilterCriteria](filtercriteria.md)|目前在指定的欄位上套用的篩選。唯讀。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|[apply(criteria:FilterCriteria)](#applycriteria-filtercriteria)|void|在指定的欄位上套用指定的篩選準則。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[applyBottomItemsFilter(count: number)](#applybottomitemsfiltercount-number)|void|套用 [底端項目] 篩選至指定元素數目的欄位。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[applyBottomPercentFilter(percent: number)](#applybottompercentfilterpercent-number)|void|套用 [底部百分比] 篩選至指定元素百分比的欄位。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[applyCellColorFilter(color: string)](#applycellcolorfiltercolor-string)|void|套用 [儲存格色彩] 篩選至指定色彩的欄位。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[applyCustomFilter(criteria1: string, criteria2: string, oper: string)](#applycustomfiltercriteria1-string-criteria2-string-oper-string)|void|套用 [圖示] 篩選至指定準則字串的欄位。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[applyDynamicFilter(criteria: string)](#applydynamicfiltercriteria-string)|void|套用 [動態] 篩選至欄位。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[applyFontColorFilter(color: string)](#applyfontcolorfiltercolor-string)|void|套用 [字型色彩] 篩選至指定色彩的欄位。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[applyIconFilter(icon:Icon)](#applyiconfiltericon-icon)|無效|套用 [圖示] 篩選至指定圖示的欄位。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[applyTopItemsFilter(count: number)](#applytopitemsfiltercount-number)|void|套用 [頂端項目] 篩選至指定元素數目的欄位。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[applyTopPercentFilter(percent: number)](#applytoppercentfilterpercent-number)|void|套用 [頂端百分比] 篩選至指定元素百分比的欄位。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[applyValuesFilter(values: ()[])](#applyvaluesfiltervalues-)|void|套用 [值] 篩選至指定值的欄位。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[clear()](#clear)|void|清除指定欄位上的篩選。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法詳細資料


### <a name="applycriteria-filtercriteria"></a>apply(criteria:FilterCriteria)
在指定的欄位上套用指定的篩選準則。

#### <a name="syntax"></a>語法
```js
filterObject.apply(criteria);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|:---|
|準則|FilterCriteria|要套用的準則。|

#### <a name="returns"></a>傳回
void

### <a name="applybottomitemsfiltercount-number"></a>applyBottomItemsFilter(count: number)
套用 [底端項目] 篩選至指定元素數目的欄位。

#### <a name="syntax"></a>語法
```js
filterObject.applyBottomItemsFilter(count);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|:---|
|Count|number|從下至上顯示的元素數目。|

#### <a name="returns"></a>傳回
void

### <a name="applybottompercentfilterpercent-number"></a>applyBottomPercentFilter(percent: number)
套用 [底部百分比] 篩選至指定元素百分比的欄位。

#### <a name="syntax"></a>語法
```js
filterObject.applyBottomPercentFilter(percent);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|:---|
|百分比|number|從下至上顯示的元素百分比。|

#### <a name="returns"></a>傳回
void

### <a name="applycellcolorfiltercolor-string"></a>applyCellColorFilter(color: string)
套用 [儲存格色彩] 篩選至指定色彩的欄位。

#### <a name="syntax"></a>語法
```js
filterObject.applyCellColorFilter(color);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|:---|
|Color|string|顯示的儲存格背景色彩。|

#### <a name="returns"></a>傳回
void

### <a name="applycustomfiltercriteria1-string-criteria2-string-oper-string"></a>applyCustomFilter(criteria1: string, criteria2: string, oper: string)
套用 [圖示] 篩選至指定準則字串的欄位。

#### <a name="syntax"></a>語法
```js
filterObject.applyCustomFilter(criteria1, criteria2, oper);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|:---|
|criteria1|string|第一個準則字串。|
|criteria2|string|選用。第二個準則字串。|
|oper|string|選用。說明如何聯結兩個準則的運算子。可能的值為：And、Or|

#### <a name="returns"></a>傳回
void

### <a name="applydynamicfiltercriteria-string"></a>applyDynamicFilter(criteria: string)
套用 [動態] 篩選至欄位。

#### <a name="syntax"></a>語法
```js
filterObject.applyDynamicFilter(criteria);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|:---|
|criteria|string|要套用的動態準則。可能的值為：未知、AboveAverage、AllDatesInPeriodApril、AllDatesInPeriodAugust、AllDatesInPeriodDecember、AllDatesInPeriodFebruray、AllDatesInPeriodJanuary、AllDatesInPeriodJuly、AllDatesInPeriodJune、AllDatesInPeriodMarch、AllDatesInPeriodMay、AllDatesInPeriodNovember、AllDatesInPeriodOctober、AllDatesInPeriodQuarter1、AllDatesInPeriodQuarter2、AllDatesInPeriodQuarter3、AllDatesInPeriodQuarter4、AllDatesInPeriodSeptember、BelowAverage、LastMonth、LastQuarter、LastWeek、LastYear、NextMonth、NextQuarter、NextWeek、NextYear、ThisMonth、ThisQuarter、ThisWeek、ThisYear、Today、Tomorrow、YearToDate、Yesterday|

#### <a name="returns"></a>傳回
void

### <a name="applyfontcolorfiltercolor-string"></a>applyFontColorFilter(color: string)
套用 [字型色彩] 篩選至指定色彩的欄位。

#### <a name="syntax"></a>語法
```js
filterObject.applyFontColorFilter(color);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|:---|
|Color|string|顯示的儲存格字型色彩。|

#### <a name="returns"></a>傳回
void

### <a name="applyiconfiltericon-icon"></a>applyIconFilter(icon:Icon)
套用 [圖示] 篩選至指定圖示的欄位。

#### <a name="syntax"></a>語法
```js
filterObject.applyIconFilter(icon);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|:---|
|圖示|圖示|顯示的儲存格圖示。|

#### <a name="returns"></a>傳回
void

### <a name="applytopitemsfiltercount-number"></a>applyTopItemsFilter(count: number)
套用 [頂端項目] 篩選至指定元素數目的欄位。

#### <a name="syntax"></a>語法
```js
filterObject.applyTopItemsFilter(count);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|:---|
|Count|number|從上至下顯示的元素數目。|

#### <a name="returns"></a>傳回
void

### <a name="applytoppercentfilterpercent-number"></a>applyTopPercentFilter(percent: number)
套用 [頂端百分比] 篩選至指定元素百分比的欄位。

#### <a name="syntax"></a>語法
```js
filterObject.applyTopPercentFilter(percent);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|:---|
|百分比|number|從上至下顯示的元素百分比。|

#### <a name="returns"></a>傳回
void

### <a name="applyvaluesfiltervalues-"></a>applyValuesFilter(values: ()[])
套用 [值] 篩選至指定值的欄位。

#### <a name="syntax"></a>語法
```js
filterObject.applyValuesFilter(values);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|:---|
|values|()[]|顯示的值清單。|

#### <a name="returns"></a>傳回
void

### <a name="clear"></a>clear()
清除指定欄位上的篩選。

#### <a name="syntax"></a>語法
```js
filterObject.clear();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
void
