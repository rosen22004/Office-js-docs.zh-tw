# <a name="boundingbox-object-javascript-api-for-visio"></a>BoundingBox 物件 (適用於 Visio 的 JavaScript API)

適用於：_Visio Online_

代表圖案的 BoundingBox。

## <a name="properties"></a>屬性

| 屬性	       | 類型	    |描述|
|:---------------|:--------|:----------|
|Height|Int|圖形周框方塊上下邊緣之間的距離，不包括與形狀相關的任何資料圖形。|
|width|Int|圖形周框方塊左右邊緣之間的距離，不包括與形狀相關的任何資料圖形。|
|x|Int|指定周框方塊的 x 座標的整數。|
|y|Int|指定周框方塊的 y 座標的整數。|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
無


## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|

## <a name="method-details"></a>方法詳細資料


### <a name="loadparam-object"></a>load(param: object)
以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。

#### <a name="syntax"></a>語法
```js
object.load(param);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|:---|
|param|物件|選用。接受參數與關聯性名稱，做為分隔字串或陣列。或者提供 [loadOption](loadoption.md) 物件。|

#### <a name="returns"></a>傳回
void
