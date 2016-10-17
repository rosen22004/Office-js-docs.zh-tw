# <a name="outline-object-(javascript-api-for-onenote)"></a>Outline 物件 (適用於 OneNote 的 JavaScript API)

_適用於：OneNote Online_  


代表 Paragraph 物件的容器。

## <a name="properties"></a>屬性

| 屬性	     | 類型	   |描述|意見反應|
|:---------------|:--------|:----------|:-------|
|id|字串|取得 Outline 物件的識別碼。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-outline-id)|

## <a name="relationships"></a>關聯性
| 關聯性 | 類型	   |描述| 意見反應|
|:---------------|:--------|:----------|:-------|
|pageContent|[PageContent](pagecontent.md)|取得包含大綱的 PageContent 物件。這個物件會定義在頁面上大綱的位置。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-outline-pageContent)|
|paragraphs|[ParagraphCollection](paragraphcollection.md)|取得大綱中 Paragraph 物件的集合。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-outline-paragraphs)|

## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述| 意見反應|
|:---------------|:--------|:----------|:-------|
|[appendHtml(html: string)](#appendhtmlhtml-string)|void|將指定的 HTML 加入大綱的底部。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-outline-appendHtml)|
|[appendImage(base64EncodedImage: string, width: double, height: double)](#appendimagebase64encodedimage-string-width-double-height-double)|[Image](image.md)|將指定的影像加入大綱的底部。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-outline-appendImage)|
|[appendRichText(paragraphText: string)](#appendrichtextparagraphtext-string)|[RichText](richtext.md)|將指定的文字加入大綱的底部。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-outline-appendRichText)|
|[appendTable(rowCount: number, columnCount: number, values: string[][])](#appendtablerowcount-number-columncount-number-values-string)|[Table](table.md)|將具有指定列和欄數的表格新增至大綱的底部。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-outline-appendTable)|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-outline-load)|

## <a name="method-details"></a>方法詳細資料


### <a name="appendhtml(html:-string)"></a>appendHtml(html: string)
將指定的 HTML 加入大綱的底部。

#### <a name="syntax"></a>語法
```js
outlineObject.appendHtml(html);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|HTML|字串|要附加的 HTML 字串。請參閱 OneNote 增益集 JavaScript API 的[支援的 HTML](../../docs/onenote/onenote-add-ins-page-content.md#supported-html)。|

#### <a name="returns"></a>傳回
void

#### <a name="examples"></a>範例
```js
OneNote.run(function (context) {

    // Gets the active page.
    var activePage = context.application.getActivePage();

    // Get pageContents of the activePage. 
    var pageContents = activePage.contents;

    // Queue a command to load the pageContents to access its data.
    context.load(pageContents);

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            if (pageContents.items.length != 0 && pageContents.items[0].type == "Outline")
            {
                // First item is an outline.
                outline = pageContents.items[0].outline;

                // Queue a command to append a paragraph to the outline.
                outline.appendHtml("<p>new paragraph</p>");

                // Run the queued commands.
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


### <a name="appendimage(base64encodedimage:-string,-width:-double,-height:-double)"></a>appendImage(base64EncodedImage: string, width: double, height: double)
將指定的影像加入大綱的底部。

#### <a name="syntax"></a>語法
```js
outlineObject.appendImage(base64EncodedImage, width, height);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|base64EncodedImage|string|要附加的 HTML 字串。|
|width|double|選用。以點為單位的寬度。預設值為 null，且會遵守影像寬度。|
|height|double|選用。以點為單位的高度。預設值為 null，且會遵守影像高度。|

#### <a name="returns"></a>傳回
[Image](image.md)

### <a name="appendrichtext(paragraphtext:-string)"></a>appendRichText(paragraphText: string)
將指定的文字加入大綱的底部。

#### <a name="syntax"></a>語法
```js
outlineObject.appendRichText(paragraphText);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|paragraphText|string|要附加的 HTML 字串。|

#### <a name="returns"></a>傳回
[RichText](richtext.md)

### <a name="appendtable(rowcount:-number,-columncount:-number,-values:-string[][])"></a>appendTable(rowCount: number, columnCount: number, values: string[][])
將具有指定列和欄數的表格新增至大綱的底部。

#### <a name="syntax"></a>語法
```js
outlineObject.appendTable(rowCount, columnCount, values);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|rowCount|number|必要。表格中的列數。|
|columnCount|number|必要。表格中的欄數。|
|values|string[][]|選用。選擇性的 2 維陣列。如果陣列中指定對應的字串，則會填滿儲存格。|

#### <a name="returns"></a>傳回
[Table](table.md)

#### <a name="examples"></a>範例
```js
OneNote.run(function (context) {

    // Gets the active page.
    var activePage = context.application.getActivePage();

    // Get pageContents of the activePage. 
    var pageContents = activePage.contents;

    // Queue a command to load the pageContents to access its data.
    context.load(pageContents);

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            if (pageContents.items.length != 0 && pageContents.items[0].type == "Outline") {
                // First item is an outline.
                var outline = pageContents.items[0].outline;

                // Queue a command to append a paragraph to the outline.
                outline.appendTable(2, 2, [[1, 2],[3, 4]]);

                // Run the queued commands.
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
