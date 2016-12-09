# <a name="rangesort-object-javascript-api-for-excel"></a>RangeSort 物件 (適用於 Excel 的 JavaScript API)

管理在 Range 物件上的排序作業。

## <a name="properties"></a>屬性

無

## <a name="relationships"></a>關聯性
無


## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|[apply(fields:SortField[], matchCase: bool, hasHeaders: bool, orientation: string, method: string)](#applyfields-sortfield-matchcase-bool-hasheaders-bool-orientation-string-method-string)|void|執行排序作業。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法詳細資料


### <a name="applyfields-sortfield-matchcase-bool-hasheaders-bool-orientation-string-method-string"></a>apply(fields:SortField[], matchCase: bool, hasHeaders: bool, orientation: string, method: string)
執行排序作業。

#### <a name="syntax"></a>語法
```js
rangeSortObject.apply(fields, matchCase, hasHeaders, orientation, method);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|:---|
|欄位|SortField[]|用來排序的條件清單。|
|matchCase|bool|選擇性。是否有大小寫影響的字串排序。|
|hasHeaders|bool|選擇性。範圍是否有標頭。|
|方向|string|選擇性。作業是排序資料列或資料行。可能的值為：資料列，資料行|
|方法|string|選擇性。適用於中文字元的排序方法。可能的值為：拼音、StrokeCount|

#### <a name="returns"></a>傳回
void
