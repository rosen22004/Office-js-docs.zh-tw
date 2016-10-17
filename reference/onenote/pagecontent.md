# <a name="pagecontent-object-(javascript-api-for-onenote)"></a>PageConten 物件 (適用於 OneNote 的 JavaScript API)

_適用於：OneNote Online_  


代表頁面上的區域，包含最上層的內容類型，例如 Outline 或 Image。可以指派 XY 座標位置給 PageContent 物件。

## <a name="properties"></a>屬性

| 屬性	     | 類型	   |描述|意見反應|
|:---------------|:--------|:----------|:-------|
|id|string|取得 PageContent 物件的識別碼。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-id)|
|left|double|取得或設定 PageContent 物件的左邊 (X-軸) 位置。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-left)|
|top|double|取得或設定 PageContent 物件的上方 (Y-軸) 位置。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-top)|
|type|string|取得 PageContent 物件的類型。唯讀。可能的值為：Outline、Image、Other。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-type)|

## <a name="relationships"></a>關聯性
| 關聯性 | 類型	   |描述| 意見反應|
|:---------------|:--------|:----------|:-------|
|image|[Image](image.md)|取得 PageContent 物件中的 Image。如果 PageContentType 不是 Image，則擲回例外狀況。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-image)|
|ink|[FloatingInk](floatingink.md)|取得 PageContent 物件中的筆跡。如果 PageContentType 不是 Ink，則擲回例外狀況。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-ink)|
|outline|[Outline](outline.md)|取得 PageContent 物件中的 Outline。如果 PageContentType 不是 Outline，則擲回例外狀況。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-outline)|
|parentPage|[Page](page.md)|取得包含 PageContent 物件的頁面。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-parentPage)|

## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述| 意見反應|
|:---------------|:--------|:----------|:-------|
|[delete()](#delete)|void|刪除 PageContent 物件。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-delete)|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-load)|

## <a name="method-details"></a>方法詳細資料


### <a name="delete()"></a>delete()
刪除 PageContent 物件。

#### <a name="syntax"></a>語法
```js
pageContentObject.delete();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
void

#### <a name="examples"></a>範例
```js
OneNote.run(function (context) {

    var page = context.application.getActivePage();
    var pageContents = page.contents;

    var firstPageContent = pageContents.getItemAt(0);
    firstPageContent.load('type');

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            if(firstPageContent.isNull === false) {
                firstPageContent.delete();
                return context.sync();
            }
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```
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
