# <a name="formatprotection-object-(javascript-api-for-excel)"></a>FormatProtection 物件 (適用於 Excel 的 JavaScript API)

_適用於：Excel 2016、Excel Online、Excel for iOS、Office 2016_

代表 range 物件的格式保護。

## <a name="properties"></a>屬性

| 屬性	     | 類型	   |描述
|:---------------|:--------|:----------|
|formulaHidden|bool|表示 Excel 是否在範圍的儲存格中隱藏公式。Null 值表示整個範圍沒有統一公式隱藏設定。|
|鎖定|bool|表示 Excel 是否在物件中鎖定儲存格。Null 值表示整個範圍沒有統一鎖定設定。|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
無


## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|

## <a name="method-details"></a>方法詳細資料


### <a name="load(param:-object)"></a>load(param: object)
以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。

#### <a name="syntax"></a>語法
```js
object.load(param);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|param|物件|選用。接受參數與關聯性名稱，做為分隔字串或陣列。或者提供 [loadOption](loadoption.md) 物件。|

#### <a name="returns"></a>傳回
void
