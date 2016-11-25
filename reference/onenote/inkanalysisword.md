# <a name="inkanalysisword-object-(javascript-api-for-onenote)"></a>InkAnalysisWord 物件 (適用於 OneNote 的 JavaScript API)

_適用於：OneNote Online_  


表示由筆跡線條形成的已識別文字的筆跡分析資料。

## <a name="properties"></a>屬性

| 屬性	     | 類型	   |描述|意見反應|
|:---------------|:--------|:----------|:-------|
|id|字串|取得 InkAnalysisWord 物件的識別碼。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisWord-id)|
|languageId|字串|這個 inkAnalysisWord 中已辨識語言的識別碼。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisWord-languageId)|
|wordAlternates|字串|在這個筆跡文字中以可能的順序辨識的文字。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisWord-wordAlternates)|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
| 關聯性 | 類型	   |描述| 意見反應|
|:---------------|:--------|:----------|:-------|
|line|[InkAnalysisLine](inkanalysisline.md)|父 InkAnalysisLine 的參考。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisWord-line)|
|strokePointers|[InkStrokePointer](inkstrokepointer.md)|辨識為這個筆跡分析文字一部分的筆跡線條的弱式參考。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisWord-strokePointers)|

## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述| 意見反應|
|:---------------|:--------|:----------|:-------|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisWord-load)|

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

**wordAlternates 和 languageId**
```js
OneNote.run(function (ctx) {        
    var app = ctx.application;
    
    // Gets the active page.
    var page = app.getActivePage();
    
    page.load('inkAnalysisOrNull/paragraphs/lines/words');
    return ctx.sync()
        .then(function() {
            var inkParagraphs = page.inkAnalysisOrNull.paragraphs;
            $.each(inkParagraphs.items, function(i, inkParagraph) {
                var inkLines = inkParagraph.lines;
                $.each(inkLines.items, function(j, inkLine) {
                    var inkWords = inkLine.words;
                    $.each(inkWords.items, function(k, inkWord) {
                    
                        // Log language Id of the word
                        console.log(inkWord.languageId);
                        
                        // Log every ink analyzed words.
                        $.each(inkWord.wordAlternates, function(l, word) {
                            console.log(word);                                  
                        })
                    })
                })
            })
        })
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
}); 
```