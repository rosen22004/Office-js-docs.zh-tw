# <a name="rangeview-object-javascript-api-for-excel"></a>RangeView 物件 (適用於 Excel 的 JavaScript API)

RangeView 表示父項範圍的一組可見儲存格。

## <a name="properties"></a>屬性

| 屬性	       | 類型	    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|cellAddresses|object[][]|表示 RangeView 的儲存格位址。唯讀。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|columnCount|Int|傳回可見資料行的數目。唯讀。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|formulas|object[][]|代表 A1 樣式標記法的公式。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|formulasLocal|object[][]|以使用者的語言和數字格式地區設定，表示 A1 樣式標記法的公式。例如，英文的 "=SUM(A1, 1.5)" 公式在德文中會表示為 "=SUMME(A1; 1,5)"。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|formulasR1C1|object[][]|代表 R1C1 樣式標記法的公式。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|index|Int|傳回值，表示 RangeView 的索引。唯讀。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|numberFormat|object[][]|代表特定儲存格的 Excel 數字格式代碼。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|rowCount|Int|傳回可見資料列的數目。唯讀。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|text|object[][]|所指定範圍的文字值。文字值與儲存格寬度無關。Excel UI 中出現的 # 替代符號不會影響 API 所傳回的文字值。唯讀。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|valueTypes|string|代表每個儲存格的資料類型。唯讀。可能的值為：Unknown、Empty、String、Integer、Double、Boolean、Error。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|values|object[][]|代表所指定範圍檢視的原始值。傳回的資料可能是 string、number 或 boolean 類型。包含錯誤的儲存格會傳回錯誤字串。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
| 關聯性 | 類型	    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|rows|[RangeViewCollection](rangeviewcollection.md)|代表與範圍關聯的範圍檢視集合。唯讀。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|[getRange()](#getrange)|[Range](range.md)|取得與目前 RangeView 相關聯的父項範圍。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法詳細資料


### <a name="getrange"></a>getRange()
取得與目前 RangeView 相關聯的父項範圍。

#### <a name="syntax"></a>語法
```js
rangeViewObject.getRange();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
[範圍](range.md)
