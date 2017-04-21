# <a name="rangeviewcollection-object-javascript-api-for-excel"></a>RangeViewCollection 物件 (適用於 Excel 的 JavaScript API)

代表 rangeView 物件的集合。

## <a name="properties"></a>屬性

| 屬性       | 類型	    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|項目|[RangeView[]](rangeview.md)|rangeView 物件的集合。唯讀。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
無


## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|[getCount()](#getcount)|Int|取得集合中的物件數目。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[RangeView](rangeview.md)|透過其索引取得 RangeView 資料列。以 0 開始編製索引。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法詳細資料


### <a name="getcount"></a>getCount()
取得集合中的物件數目。

#### <a name="syntax"></a>語法
```js
rangeViewCollectionObject.getCount();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
Int

### <a name="getitematindex-number"></a>getItemAt(index: number)
透過其索引取得 RangeView 資料列。以 0 開始編製索引。

#### <a name="syntax"></a>語法
```js
rangeViewCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|
|index|number|可見資料列的索引。|

#### <a name="returns"></a>傳回
[RangeView](rangeview.md)
