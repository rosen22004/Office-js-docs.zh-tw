# <a name="pivottablecollection-object-javascript-api-for-excel"></a>PivotTableCollection 物件 (適用於 Excel 的 JavaScript API)

代表屬於活頁簿或工作表一部份的所有樞紐分析表集合。

## <a name="properties"></a>屬性

| 屬性	       | 類型	    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|項目|[PivotTable[]](pivottable.md)|pivotTable 物件的集合。唯讀。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
無


## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|[getCount()](#getcount)|Int|取得集合中的樞紐分析表數目。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(name: string)](#getitemname-string)|[PivotTable](pivottable.md)|藉由名稱取得樞紐分析表。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNullObject(name: string)](#getitemornullobjectname-string)|[PivotTable](pivottable.md)|依名稱取得樞紐分析表。如果樞紐分析表不存在，會傳回 null 物件。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[refreshAll()](#refreshall)|void|重新整理集合中的所有樞紐分析表。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法詳細資料


### <a name="getcount"></a>getCount()
取得集合中的樞紐分析表數目。

#### <a name="syntax"></a>語法
```js
pivotTableCollectionObject.getCount();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
Int

### <a name="getitemname-string"></a>getItem(name: string)
藉由名稱取得樞紐分析表。

#### <a name="syntax"></a>語法
```js
pivotTableCollectionObject.getItem(name);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|:---|
|Name|string|要擷取之樞紐分析表的名稱。|

#### <a name="returns"></a>傳回
[PivotTable](pivottable.md)

### <a name="getitemornullobjectname-string"></a>getItemOrNullObject(name: string)
依名稱取得樞紐分析表。如果樞紐分析表不存在，會傳回 null 物件。

#### <a name="syntax"></a>語法
```js
pivotTableCollectionObject.getItemOrNullObject(name);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|:---|
|Name|string|要擷取之樞紐分析表的名稱。|

#### <a name="returns"></a>傳回
[PivotTable](pivottable.md)

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
