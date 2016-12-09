# <a name="settingcollection-object-javascript-api-for-excel"></a>SettingCollection object (JavaScript API for Excel)

代表屬於活頁簿一部份的 worksheet 物件集合。

## <a name="properties"></a>屬性

| 屬性	     | 類型	   |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|items|[Setting[]](setting.md)|設定物件的集合。唯讀。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
無


## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|[getItem(key: string)](#getitemkey-string)|[Setting](setting.md)|透過索引鍵取得設定項目。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNull(key: string)](#getitemornullkey-string)|[Setting](setting.md)|透過索引鍵取得設定項目。如果設定物件不存在，傳回物件的 isNull 屬性為 true。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[set(key: string, value: string)](#setkey-string-value-string)|[Setting](setting.md)|將指定的設定設定或新增至活頁簿。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法詳細資料


### <a name="getitemkey-string"></a>getItem(key: string)
透過索引鍵取得設定項目。

#### <a name="syntax"></a>語法
```js
settingCollectionObject.getItem(key);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|:---|
|索引鍵|string|設定的索引鍵。|

#### <a name="returns"></a>傳回
[Setting](setting.md)

### <a name="getitemornullkey-string"></a>getItemOrNull(key: string)
透過索引鍵取得設定項目。如果設定物件不存在，傳回物件的 isNull 屬性為 true。

#### <a name="syntax"></a>語法
```js
settingCollectionObject.getItemOrNull(key);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|:---|
|索引鍵|string|設定的索引鍵。|

#### <a name="returns"></a>傳回
[Setting](setting.md)

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

### <a name="setkey-string-value-string"></a>set(key: string, value: string)
將指定的設定設定或新增至活頁簿。

#### <a name="syntax"></a>語法
```js
settingCollectionObject.set(key, value);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|:---|
|索引鍵|string|新設定的索引鍵。|
|value|string|新設定的值。|

#### <a name="returns"></a>傳回
[Setting](setting.md)
