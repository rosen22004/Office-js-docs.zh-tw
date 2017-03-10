# <a name="selection-object-javascript-api-for-visio"></a>Selection 物件 (適用於 Visio 的 JavaScript API)

適用於：_Visio Online_

代表頁面內的選取範圍。

## <a name="properties"></a>屬性

無

## <a name="relationships"></a>關聯性
| 關聯性 | 類型	    |描述|
|:---------------|:--------|:----------|
|圖形|[ShapeCollection](shapecollection.md)|取得唯讀選取範圍的圖形。|

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
