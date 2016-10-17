# <a name="sortfield-object-(javascript-api-for-excel)"></a>SortField 物件 (適用於 Excel 的 JavaScript API)

_適用於：Excel 2016、Excel Online、Excel for iOS、Office 2016_

表示排序作業中的條件。

## <a name="properties"></a>屬性

| 屬性	     | 類型	   |描述
|:---------------|:--------|:----------|
|遞增|bool|表示是否以遞增方式完成排序。|
|色彩|字串|如果排序是針對字型或儲存格色彩，表示色彩是條件的目標。|
|dataOption|string|表示此欄位額外的排序選項。可能的值為：一般、TextAsNumber。|
|索引鍵|int|表示套用條件的資料行 (或資料列，視排序的方向而定)。表示為從第一個資料行 (或資料列) 的位移。|
|sortOn|字串|表示這個條件的排序的類型。可能的值為：Value、CellColor、FontColor、Icon。|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
| 關聯性 | 類型	   |描述|
|:---------------|:--------|:----------|
|icon|[Icon](icon.md)|如果排序是針對儲存格的圖示，表示圖示是條件的目標。|

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
