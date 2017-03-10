# <a name="filtercriteria-object-javascript-api-for-excel"></a>FilterCriteria 物件 (適用於 Excel 的 JavaScript API)

表示套用到資料行的篩選準則。

## <a name="properties"></a>屬性

| 屬性	       | 類型	    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|Color|string|用來篩選儲存格的 HTML 色彩字串。搭配使用 "cellColor" 和 "fontColor" 篩選。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|criterion1|string|用來篩選資料的第一個準則。用來做為「自訂」篩選案例中的運算子。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|criterion2|string|用來篩選資料的第二個準則。只用來做為「自訂」篩選案例中的運算子。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|dynamicCriteria|string|Excel.DynamicFilterCriteria 的動態準則設定為在此資料行上套用。與「動態」篩選搭配使用。可能的值為：未知、AboveAverage、AllDatesInPeriodApril、AllDatesInPeriodAugust、AllDatesInPeriodDecember、AllDatesInPeriodFebruray、AllDatesInPeriodJanuary、AllDatesInPeriodJuly、AllDatesInPeriodJune、AllDatesInPeriodMarch、AllDatesInPeriodMay、AllDatesInPeriodNovember、AllDatesInPeriodOctober、AllDatesInPeriodQuarter1、AllDatesInPeriodQuarter2、AllDatesInPeriodQuarter3、AllDatesInPeriodQuarter4、AllDatesInPeriodSeptember、BelowAverage、LastMonth、LastQuarter、LastWeek、LastYear、NextMonth、NextQuarter、NextWeek、NextYear、ThisMonth、ThisQuarter、ThisWeek、ThisYear、Today、Tomorrow、YearToDate、Yesterday。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|filterOn|string|篩選器用於判斷值是否仍看得見的屬性。可能的值為：BottomItems、BottomPercent、CellColor、Dynamic、FontColor、Values、TopItems、TopPercent、Icon、Custom。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|operator|string|使用「自訂」篩選時，用來結合準則 1 和 2 的運算子。可能的值為：And、Or。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|values|object[]|要做為「值」篩選部分的值集合。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
| 關聯性 | 類型	    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|圖示|[Icon](icon.md)|用來篩選儲存格的圖示。與「圖示」篩選搭配使用。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>方法
無

