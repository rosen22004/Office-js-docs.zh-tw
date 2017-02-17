# <a name="document-object-javascript-api-for-visio"></a>Document 物件 (適用於 Visio 的 JavaScript API)

適用於：_Visio Online_
>**附註：**Visio JavaScript API 目前是預覽模式，可能有所異動。Visio JavaScript API 目前不支援在生產環境中使用。

代表文件類別。

## <a name="properties"></a>屬性

無

## <a name="relationships"></a>關聯性
| 關聯性 | 類型	    |描述| 意見反應|
|:---------------|:--------|:----------|:---|
|應用程式|[應用程式](application.md)|代表包含此文件的 Visio 應用程式執行個體。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-document-application)|
|pages|[PageCollection](pagecollection.md)|代表與文件關聯的頁面集合。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-document-pages)|

## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述| 意見反應|
|:---------------|:--------|:----------|:---|
|[getActivePage()](#getactivepage)|[頁面](page.md)|傳回文件的使用中頁面。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-document-getActivePage)|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-document-load)|
|[setActivePage(PageName: string)](#setactivepagepagename-string)|無效|設定文件的使用中頁面。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-document-setActivePage)|

## <a name="method-details"></a>方法詳細資料


### <a name="getactivepage"></a>getActivePage()
傳回文件的使用中頁面。

#### <a name="syntax"></a>語法
```js
documentObject.getActivePage();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
[Page](page.md)

#### <a name="examples"></a>範例
```js
Visio.run(function (ctx) { 
    var document = ctx.document;
    var activePage = document.getActivePage();
    activePage.load();
    return ctx.sync().then(function () {
    console.log("pageName: " +activePage.name);
      });   
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
無效

### <a name="setactivepagepagename-string"></a>setActivePage(PageName: string)
設定文件的使用中頁面。

#### <a name="syntax"></a>語法
```js
documentObject.setActivePage(PageName);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|:---|
|PageName|string|頁面名稱|

#### <a name="returns"></a>傳回
void

#### <a name="examples"></a>範例
```js
Visio.run(function (ctx) { 
    var document = ctx.document;
    var pageName = "Page-1";
    document.setActivePage(pageName);
    return ctx.sync();
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="property-access-examples"></a>屬性存取範例
```js
Visio.run(function (ctx) { 
    var pages = ctx.document.pages;
    var pageCount = pages.getCount();
    return ctx.sync().then(function () {
        console.log("Pages Count: " +pageCount.value);
        });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="property-access-examples"></a>屬性存取範例
```js
Visio.run(function (ctx) { 
    var application = ctx.document.application;
    application.showToolbars = false;
    return ctx.sync();
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

