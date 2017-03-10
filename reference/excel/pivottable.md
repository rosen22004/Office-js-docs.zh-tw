# <a name="pivottable-object-javascript-api-for-excel"></a>PivotTable 物件 (適用於 Excel 的 JavaScript API)

代表 Excel 樞紐分析表。

## <a name="properties"></a>屬性

| 屬性	       | 類型	    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|名稱|string|樞紐分析表的名稱。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
| 關聯性 | 類型	    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|工作表|[Worksheet](worksheet.md)|包含目前樞紐分析表的工作表。唯讀。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|[refresh()](#refresh)|void|重新整理樞紐分析表。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法詳細資料


### <a name="refresh"></a>refresh()
重新整理樞紐分析表。

#### <a name="syntax"></a>語法
```js
pivotTableObject.refresh();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
void
