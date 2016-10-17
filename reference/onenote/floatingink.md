# <a name="floatingink-object-(javascript-api-for-onenote)"></a>FloatingInk 物件 (適用於 OneNote 的 JavaScript API)

_適用於：OneNote Online_  


表示一組筆跡線條。

## <a name="properties"></a>屬性

| 屬性	     | 類型	   |描述|意見反應|
|:---------------|:--------|:----------|:-------|
|id|字串|取得 FloatingInk 物件的識別碼。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-floatingInk-id)|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
| 關聯性 | 類型	   |描述| 意見反應|
|:---------------|:--------|:----------|:-------|
|inkStrokes|[InkStrokeCollection](inkstrokecollection.md)|取得 FloatingInk 物件的線條。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-floatingInk-inkStrokes)|
|pageContent|[PageContent](pagecontent.md)|取得 FloatingInk 物件的 PageContent 上階。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-floatingInk-pageContent)|

## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述| 意見反應|
|:---------------|:--------|:----------|:-------|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-floatingInk-load)|

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
### <a name="property-access-examples"></a>屬性存取範例

**id**
```js
OneNote.run(function(context) {

    // Gets the active page.
    var page = context.application.getActivePage();
    var contents = page.contents;
    
    // Load page contents and their types.
    page.load('contents/type');
    return context.sync()
        .then(function(){
        
            // Load every ink content.
            $.each(contents.items, function(i, content) {
                if (content.type == "Ink")
                {
                    content.load('ink/id');
                }                           
            })
            return context.sync();
        })
        .then(function(){
        
            // Log ID of every ink content.
            $.each(contents.items, function(i, content) {
                if (content.type == "Ink")
                {
                    console.log(content.ink.id);
                }                           
            })              
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
}); 
```
