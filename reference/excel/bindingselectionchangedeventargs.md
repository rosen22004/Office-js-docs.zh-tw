# <a name="bindingselectionchangedeventargs-object-javascript-api-for-excel"></a>BindingSelectionChangedEventArgs 物件 (適用於 Excel 的 JavaScript API)

提供引發 SelectionChanged 事件之繫結的相關資訊。

## <a name="properties"></a>屬性

| 屬性	       | 類型	    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|columnCount|Int|取得選取的資料欄數目。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|rowCount|Int|取得選取的資料列數目。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|startColumn|Int|取得選取範圍首欄的索引 (以零為基礎)。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|startRow|Int|取得選取範圍首列的索引 (以零為基礎)。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
| 關聯性 | 類型	    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|繫結|[繫結](binding.md)|取得代表引發 SelectionChanged 事件之繫結的 Binding 物件。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>方法
無

