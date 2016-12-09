# <a name="pivottablecollection-object-javascript-api-for-excel"></a>PivotTableCollection 物件 (適用於 Excel 的 JavaScript API)

表示屬於活頁簿或工作表一部份的所有樞紐分析表集合。

## <a name="properties"></a>屬性

| 屬性	     | 類型	   |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|items|[PivotTable[]](pivottable.md)|pivotTable 物件的集合。唯讀。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
無


## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|[getItem(name: string)](#getitemname-string)|[PivotTable](pivottable.md)|藉由名稱取得樞紐分析表。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNull(name: string)](#getitemornullname-string)|[PivotTable](pivottable.md)|藉由名稱取得樞紐分析表。如果樞紐分析表不存在，傳回物件的 isNull 屬性為 true。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[refreshAll()](#refreshall)|void|重新整理集合中的所有樞紐分析表。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法詳細資料


### <a name="getitemname-string"></a>getItem(name: string)
藉由名稱取得樞紐分析表。

#### <a name="syntax"></a>語法
```js
pivotTableCollectionObject.getItem(name);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|:---|
|Name|string|要擷取之樞紐分析表的名稱。|

#### <a name="returns"></a>傳回
[PivotTable](pivottable.md)

### <a name="getitemornullname-string"></a>getItemOrNull(name: string)
藉由名稱取得樞紐分析表。如果樞紐分析表不存在，傳回物件的 isNull 屬性為 true。

#### <a name="syntax"></a>語法
```js
pivotTableCollectionObject.getItemOrNull(name);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|:---|
|Name|string|要擷取之樞紐分析表的名稱。|

#### <a name="returns"></a>傳回
[PivotTable](pivottable.md)

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

### <a name="refreshall"></a>refreshAll()
重新整理集合中的所有樞紐分析表。

#### <a name="syntax"></a>語法
```js
pivotTableCollectionObject.refreshAll();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
void
