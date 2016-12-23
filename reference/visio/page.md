# <a name="page-object-javascript-api-for-visio"></a>Page 物件 (適用於 Visio 的 JavaScript API)

適用於：_Visio Online_
>**附註：**Visio JavaScript API 目前是預覽模式，可能有所異動。Visio JavaScript API 目前不支援在生產環境中使用。

代表頁面類別。

## <a name="properties"></a>屬性

| 屬性	     | 類型	   |描述| 意見反應|
|:---------------|:--------|:----------|:---|
|索引|Int|頁面索引。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-page-index)|
|isBackground|bool|不論是背景頁面與否。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-page-isBackground)|
|name|string|頁面名稱。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-page-name)|

## <a name="relationships"></a>關聯性
| 關聯性 | 類型	   |描述| 意見反應|
|:---------------|:--------|:----------|:---|
|圖形|[ShapeCollection](shapecollection.md)|頁面內的圖形。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-page-shapes)|
|檢視|[PageView](pageview.md)|傳回頁面的檢視。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-page-view)|

## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述| 意見反應|
|:---------------|:--------|:----------|:---|
|[activate()](#activate)|無效|將文件頁面設定為使用中的頁面。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-page-activate)|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-page-load)|

## <a name="method-details"></a>方法詳細資料


### <a name="activate"></a>activate()
將文件頁面設定為使用中的頁面。

#### <a name="syntax"></a>語法
```js
pageObject.activate();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
void

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
