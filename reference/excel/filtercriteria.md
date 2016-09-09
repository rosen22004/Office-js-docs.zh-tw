# FilterCriteria 物件 (適用於 Excel 的 JavaScript API)

_適用版本：Excel 2016、Excel Online、Excel for iOS、Office 2016_

表示套用到資料行的篩選準則。

## 屬性

| 屬性	     | 類型	   |說明
|:---------------|:--------|:----------|
|Color|string|用來篩選儲存格的 HTML 色彩字串。搭配使用 "cellColor" 和 "fontColor" 篩選。|
|criterion1|string|用來篩選資料的第一個準則。用來做為「自訂」篩選案例中的運算子。|
|criterion2|string|用來篩選資料的第二個準則。只用來做為「自訂」篩選案例中的運算子。|
|dynamicCriteria|string|Excel.DynamicFilterCriteria 的動態準則設定為在此資料行上套用。與「動態」篩選搭配使用。可能的值為：未知、AboveAverage、AllDatesInPeriodApril、AllDatesInPeriodAugust、AllDatesInPeriodDecember、AllDatesInPeriodFebruray、AllDatesInPeriodJanuary、AllDatesInPeriodJuly、AllDatesInPeriodJune、AllDatesInPeriodMarch、AllDatesInPeriodMay、AllDatesInPeriodNovember、AllDatesInPeriodOctober、AllDatesInPeriodQuarter1、AllDatesInPeriodQuarter2、AllDatesInPeriodQuarter3、AllDatesInPeriodQuarter4、AllDatesInPeriodSeptember、BelowAverage、LastMonth、LastQuarter、LastWeek、LastYear、NextMonth、NextQuarter、NextWeek、NextYear、ThisMonth、ThisQuarter、ThisWeek、ThisYear、Today、Tomorrow、YearToDate、Yesterday。|
|filterOn|string|篩選器用於判斷值是否仍看得見的屬性。可能的值為：   BottomItems、BottomPercent、CellColor、Dynamic、FontColor、Values、TopItems、TopPercent、Icon、Custom |
|values|object[]|要做為「值」篩選部分的值集合。|

## 關聯性
| 關聯性 | 類型	   |描述|
|:---------------|:--------|:----------|
|圖示|[圖示](icon.md)|用來篩選儲存格的圖示。與「圖示」篩選搭配使用。|
|運算子|[FilterOperator](filteroperator.md)|使用「自訂」篩選時，用來結合準則 1 和 2 的運算子。|

## 方法

| 方法           | 傳回類型    |說明|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|

## 方法詳細資料


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
