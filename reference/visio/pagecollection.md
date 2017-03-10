# <a name="pagecollection-object-javascript-api-for-visio"></a>PageCollection 物件 (適用於 Visio 的 JavaScript API)

適用於：_Visio Online_

代表屬於文件一部份的頁面物件集合。

## <a name="properties"></a>屬性

| 屬性	       | 類型	    |描述|
|:---------------|:--------|:----------|
|項目|[Page[]](page.md)|Page 物件的集合。唯讀。|

## <a name="relationships"></a>關聯性
無


## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述|
|:---------------|:--------|:----------|
|[getCount()](#getcount)|Int|取得集合中的頁面數目。|
|[getItem(key: number 或 string)](#getitemkey-number-or-string)|[頁面](page.md)|取得使用其機碼 (名稱或 ID) 的頁面。|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|

## <a name="method-details"></a>方法詳細資料


### <a name="getcount"></a>getCount()
取得集合中的頁面數目。

#### <a name="syntax"></a>語法
```js
pageCollectionObject.getCount();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
Int

### <a name="getitemkey-number-or-string"></a>getItem(key: number 或 string)
取得使用其機碼 (名稱或 ID) 的頁面。

#### <a name="syntax"></a>語法
```js
pageCollectionObject.getItem(key);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|:---|
|key|number 或 string|機碼是要擷取之頁面的名稱或 ID。|

#### <a name="returns"></a>傳回
[Page](page.md)

#### <a name="examples"></a>範例
```js
Visio.run(function (ctx) { 
    var pageName = 'Page-1';
    var page = ctx.document.pages.getItem(pageName);
    page.activate();
    return ctx.sync();
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

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
