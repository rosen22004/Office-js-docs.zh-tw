# <a name="sortfield-object-javascript-api-for-excel"></a>SortField 物件 (適用於 Excel 的 JavaScript API)

表示排序作業中的條件。

## <a name="properties"></a>屬性

| 屬性	       | 類型	    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|ascending|bool|表示是否以遞增方式完成排序。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|color|string|如果排序是針對字型或儲存格色彩，表示色彩是條件的目標。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|dataOption|string|表示此欄位額外的排序選項。可能的值為：一般、TextAsNumber。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|key|int|表示套用條件的資料行 (或資料列，視排序的方向而定)。表示為從第一個資料行 (或資料列) 的位移。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|sortOn|string|表示這個條件的排序的類型。可能的值為：值、CellColor、FontColor、圖示。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
| 關聯性 | 類型	    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|圖示|[Icon](icon.md)|如果排序是針對儲存格的圖示，表示圖示是條件的目標。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>方法
無

