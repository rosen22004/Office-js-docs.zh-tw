# <a name="worksheetprotectionoptions-object-javascript-api-for-excel"></a>WorksheetProtectionOptions 物件 (適用於 Excel 的 JavaScript API)

表示工作表保護中的選項。

## <a name="properties"></a>屬性

| 屬性	     | 類型	   |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|allowAutoFilter|bool|表示工作表保護選項，允許使用自動篩選功能。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|allowDeleteColumns|bool|表示工作表保護選項，允許刪除資料行。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|allowDeleteRows|bool|表示工作表保護選項，允許刪除資料列。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|allowFormatCells|bool|表示工作表保護選項，允許格式化儲存格。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|allowFormatColumns|bool|表示工作表保護選項，允許格式化資料行。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|allowFormatRows|bool|表示工作表保護選項，允許格式化資料列。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|allowInsertColumns|bool|表示工作表保護選項，允許插入資料行。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|allowInsertHyperlinks|bool|表示工作表保護選項，允許插入超連結。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|allowInsertRows|bool|表示工作表保護選項，允許插入資料列。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|allowPivotTables|bool|表示工作表保護選項，允許使用樞紐分析表功能。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|allowSort|bool|表示工作表保護選項，允許使用排序功能。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
無


## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法詳細資料


### <a name="loadparam-object"></a>load(param: object)
以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。

#### <a name="syntax"></a>語法
```js
object.load(param);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|:---|
|param|物件|選用。接受參數與關聯性名稱，做為分隔字串或陣列。或者提供 [loadOption](loadoption.md) 物件。|

#### <a name="returns"></a>傳回
void
