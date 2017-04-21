# <a name="settingcollection-object-javascript-api-for-excel"></a>SettingCollection 物件 (適用於 Excel 的 JavaScript API)

代表屬於活頁簿一部份的 Worksheet 物件集合。

## <a name="properties"></a>屬性

| 屬性       | 類型	    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|項目|[Setting[]](setting.md)|設定物件的集合。唯讀。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
無


## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|[add(key: string, value: (any)[])](#addkey-string-value-any)|[設定](setting.md)|將指定的設定新增或設定至活頁簿。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getCount()](#getcount)|Int|取得集合中的設定數目。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(key: string)](#getitemkey-string)|[Setting](setting.md)|透過機碼取得設定項目。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNullObject(key: string)](#getitemornullobjectkey-string)|[Setting](setting.md)|透過機碼取得設定項目。如果設定不存在，會傳回 null 物件。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法詳細資料


### <a name="addkey-string-value-any"></a>add(key: string, value: (any)[])
將指定的設定新增或設定至活頁簿。

#### <a name="syntax"></a>語法
```js
settingCollectionObject.add(key, value);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|
|索引鍵|string|新設定的索引鍵。|
|數值|(any)[]|新設定的值。|

#### <a name="returns"></a>傳回
[設定](setting.md)

### <a name="getcount"></a>getCount()
取得集合中的設定數目。

#### <a name="syntax"></a>語法
```js
settingCollectionObject.getCount();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
Int

### <a name="getitemkey-string"></a>getItem(key: string)
透過索引鍵取得設定項目。

#### <a name="syntax"></a>語法
```js
settingCollectionObject.getItem(key);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|
|索引鍵|string|設定的索引鍵。|

#### <a name="returns"></a>傳回
[設定](setting.md)

### <a name="getitemornullobjectkey-string"></a>getItemOrNullObject(key: string)
透過機碼取得設定項目。如果設定不存在，會傳回 null 物件。

#### <a name="syntax"></a>語法
```js
settingCollectionObject.getItemOrNullObject(key);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|
|索引鍵|string|設定的索引鍵。|

#### <a name="returns"></a>傳回
[Setting](setting.md)
