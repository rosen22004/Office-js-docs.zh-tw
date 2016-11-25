# <a name="paragraph-object-(javascript-api-for-onenote)"></a>Paragraph 物件 (適用於 OneNote 的 JavaScript API)

_適用於：OneNote Online_  


在頁面上可見內容的容器。段落可以包含任何一種 ParagraphType 類型的內容。

## <a name="properties"></a>屬性

| 屬性	     | 類型	   |描述|意見反應|
|:---------------|:--------|:----------|:-------|
|id|string|取得 Paragraph 物件的識別碼。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-id)|
|type|string|取得 Paragraph 物件的類型。唯讀。可能的值為：RichText、Image、Table、Other。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-type)|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
| 關聯性 | 類型	   |描述| 意見反應|
|:---------------|:--------|:----------|:-------|
|image|[Image](image.md)|取得段落中的 Image 物件。如果 ParagraphType 不是 Image，則擲回例外狀況。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-image)|
|inkWords|[InkWordCollection](inkwordcollection.md)|取得段落中的 Ink 集合。如果 ParagraphType 不是 Ink，則擲回例外狀況。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-inkWords)|
|outline|[Outline](outline.md)|取得包含段落的 Outline 物件。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-outline)|
|paragraphs|[ParagraphCollection](paragraphcollection.md)|在這個段落底下的段落集合。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-paragraphs)|
|parentParagraph|[Paragraph](paragraph.md)|取得父段落物件。如果父段落不存在，則擲回。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-parentParagraph)|
|parentParagraphOrNull|[Paragraph](paragraph.md)|取得父段落物件。如果父段落不存在，則傳回 null。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-parentParagraphOrNull)|
|parentTableCell|[TableCell](tablecell.md)|取得包含 Paragraph (如果存在) 的 TableCell 物件。如果父項不是 TableCell，則擲回 ItemNotFound。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-parentTableCell)|
|parentTableCellOrNull|[TableCell](tablecell.md)|取得包含 Paragraph (如果存在) 的 TableCell 物件。如果父項不是 TableCell，則傳回 null。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-parentTableCellOrNull)|
|richText|[RichText](richtext.md)|取得段落中的 RichText 物件。如果 ParagraphType 不是 RichText，則擲回例外狀況。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-richText)|
|table|[Table](table.md)|取得段落中的 Table 物件。如果 ParagraphType 不是 Table，則擲回例外狀況。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-table)|

## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述| 意見反應|
|:---------------|:--------|:----------|:-------|
|[delete()](#delete)|void|刪除段落|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-delete)|
|[insertHtmlAsSibling(insertLocation: string, html: string)](#inserthtmlassiblinginsertlocation-string-html-string)|void|插入指定的 HTML 內容|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-insertHtmlAsSibling)|
|[insertImageAsSibling(insertLocation: string, base64EncodedImage: string, width: double, height: double)](#insertimageassiblinginsertlocation-string-base64encodedimage-string-width-double-height-double)|[Image](image.md)|在指定插入位置插入影像。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-insertImageAsSibling)|
|[insertRichTextAsSibling(insertLocation: string, paragraphText: string)](#insertrichtextassiblinginsertlocation-string-paragraphtext-string)|[RichText](richtext.md)|在指定插入位置插入段落文字。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-insertRichTextAsSibling)|
|[insertTableAsSibling(insertLocation: string, rowCount: number, columnCount: number, values: string[][])](#inserttableassiblinginsertlocation-string-rowcount-number-columncount-number-values-string)|[Table](table.md)|將具有指定列和欄數的表格新增到目前段落之前或之後。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-insertTableAsSibling)|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-load)|

## <a name="method-details"></a>方法詳細資料


### <a name="delete()"></a>delete()
刪除段落

#### <a name="syntax"></a>語法
```js
paragraphObject.delete();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
void

#### <a name="examples"></a>範例
```js
OneNote.run(function (context) {

    // Get the collection of pageContent items from the page.
    var pageContents = context.application.getActivePage().contents;

    // Get the first PageContent on the page
    // Assuming its an outline, get the outline's paragraphs.
    var pageContent = pageContents.getItemAt(0);
    
    var paragraphs = pageContent.outline.paragraphs;
    
    var firstParagraph = paragraphs.getItemAt(0);
    
    // Queue a command to load the id and type of the first paragraph
    firstParagraph.load("id,type");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            
            // Queue a command to delete the first paragraph                 
            firstParagraph.delete();
            
            // Run the command to delete it
            return context.sync();
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### <a name="inserthtmlassibling(insertlocation:-string,-html:-string)"></a>insertHtmlAsSibling(insertLocation: string, html: string)
插入指定的 HTML 內容

#### <a name="syntax"></a>語法
```js
paragraphObject.insertHtmlAsSibling(insertLocation, html);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|insertLocation|string|新內容與目前段落的相對位置。可能的值為：之前、之後|
|HTML|string|HTML 字串，描述內容的視覺化呈現。請參閱 OneNote 增益集 JavaScript API 的[支援的 HTML](../../docs/onenote/onenote-add-ins-page-content.md#supported-html)。|

#### <a name="returns"></a>傳回
void

#### <a name="examples"></a>範例
```js
OneNote.run(function (context) {

    // Get the collection of pageContent items from the page.
    var pageContents = context.application.getActivePage().contents;

    // Get the first PageContent on the page
    // Assuming its an outline, get the outline's paragraphs.
    var pageContent = pageContents.getItemAt(0);
    var paragraphs = pageContent.outline.paragraphs;
    var firstParagraph = paragraphs.getItemAt(0);

    // Queue a command to load the id and type of the first paragraph
    firstParagraph.load("id,type");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // Queue commands to insert before and after the first paragraph
            firstParagraph.insertHtmlAsSibling("Before", "<p>ContentBeforeFirstParagraph</p>");
            firstParagraph.insertHtmlAsSibling("After", "<p>ContentAfterFirstParagraph</p>");
            
            // Run the command to run inserts
            return context.sync();
        });
))
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### <a name="insertimageassibling(insertlocation:-string,-base64encodedimage:-string,-width:-double,-height:-double)"></a>insertImageAsSibling(insertLocation: string, base64EncodedImage: string, width: double, height: double)
在指定插入位置插入影像。

#### <a name="syntax"></a>語法
```js
paragraphObject.insertImageAsSibling(insertLocation, base64EncodedImage, width, height);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|insertLocation|string|表格與目前段落的相對位置。可能的值為：之前、之後|
|base64EncodedImage|string|要附加的 HTML 字串。|
|width|double|選用。以點為單位的寬度。預設值為 null，且會遵守影像寬度。|
|height|double|選用。以點為單位的高度。預設值為 null，且會遵守影像高度。|

#### <a name="returns"></a>傳回
[Image](image.md)

#### <a name="examples"></a>範例
```js
OneNote.run(function (context) {

    // Get the collection of pageContent items from the page.
    var pageContents = context.application.getActivePage().contents;

    // Get the first PageContent on the page
    // Assuming its an outline, get the outline's paragraphs.
    var pageContent = pageContents.getItemAt(0);
    var paragraphs = pageContent.outline.paragraphs;
    var firstParagraph = paragraphs.getItemAt(0);

    // Queue a command to load the id and type of the first paragraph
    firstParagraph.load("id,type");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // Queue commands to insert before and after the first paragraph
            firstParagraph.insertImageAsSibling("Before", "R0lGODlhDwAPAKECAAAAzMzM/////wAAACwAAAAADwAPAAACIISPeQHsrZ5ModrLlN48CXF8m2iQ3YmmKqVlRtW4MLwWACH+H09wdGltaXplZCBieSBVbGVhZCBTbWFydFNhdmVyIQAAOw==");
            firstParagraph.insertImageAsSibling("After", "R0lGODlhDwAPAKECAAAAzMzM/////wAAACwAAAAADwAPAAACIISPeQHsrZ5ModrLlN48CXF8m2iQ3YmmKqVlRtW4MLwWACH+H09wdGltaXplZCBieSBVbGVhZCBTbWFydFNhdmVyIQAAOw==");
            
            // Run the command to insert images
            return context.sync();
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### <a name="insertrichtextassibling(insertlocation:-string,-paragraphtext:-string)"></a>insertRichTextAsSibling(insertLocation: string, paragraphText: string)
在指定插入位置插入段落文字。

#### <a name="syntax"></a>語法
```js
paragraphObject.insertRichTextAsSibling(insertLocation, paragraphText);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|insertLocation|string|表格與目前段落的相對位置。可能的值為：之前、之後|
|paragraphText|string|要附加的 HTML 字串。|

#### <a name="returns"></a>傳回
[RichText](richtext.md)

#### <a name="examples"></a>範例
```js
OneNote.run(function (context) {

    // Get the collection of pageContent items from the page.
    var pageContents = context.application.getActivePage().contents;

    // Get the first PageContent on the page
    // Assuming its an outline, get the outline's paragraphs.
    var pageContent = pageContents.getItemAt(0);
    var paragraphs = pageContent.outline.paragraphs;
    var firstParagraph = paragraphs.getItemAt(0);

    // Queue a command to load the id and type of the first paragraph
    firstParagraph.load("id,type");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // Queue commands to insert before and after the first paragraph
            firstParagraph.insertRichTextAsSibling("Before", "Text Appears Before Paragraph");
            firstParagraph.insertRichTextAsSibling("After", "Text Appears After Paragraph");
            
            // Run the command to insert text contents
            return context.sync();
        });
})  
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
}); 
```


### <a name="inserttableassibling(insertlocation:-string,-rowcount:-number,-columncount:-number,-values:-string[][])"></a>insertTableAsSibling(insertLocation: string, rowCount: number, columnCount: number, values: string[][])
將具有指定列和欄數的表格新增到目前段落之前或之後。

#### <a name="syntax"></a>語法
```js
paragraphObject.insertTableAsSibling(insertLocation, rowCount, columnCount, values);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|insertLocation|string|表格與目前段落的相對位置。可能的值為：之前、之後|
|rowCount|number|表格中的列數。|
|columnCount|number|表格中的欄數。|
|values|string[][]|選用。選擇性的 2 維陣列。如果陣列中指定對應的字串，則會填滿儲存格。|

#### <a name="returns"></a>傳回
[Table](table.md)

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

**id 和 type**
```js
OneNote.run(function (context) {

    // Get the collection of pageContent items from the page.
    var pageContents = context.application.getActivePage().contents;
    
    // Queue a command to load the outline property of each pageContent.
    pageContents.load("outline");
        
    // Get the first PageContent on the page, and then get its Outline.
    var pageContent = pageContents._GetItem(0);
    var paragraphs = pageContent.outline.paragraphs;
            
    // Queue a command to load the id and type of each paragraph.
    paragraphs.load("id,type");
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            
            // Write the text.                  
            $.each(paragraphs.items, function(index, paragraph) {
                console.log("Paragraph type: " + paragraph.type);
                console.log("Paragraph ID: " + paragraph.id);
            });
        });
})      
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
}); 
```

**paragraphs**
```js
OneNote.run(function(context) {
    var app = context.application;
    
    // Gets the active outline
    var outline = app.getActiveOutline();
    
    // load nested paragraphs and their types.
    outline.load("paragraphs/type");
    
    return context.sync().then(function () {
        var paragraphs = outline.paragraphs.items;
        
        var promise;
        // for each nested paragraphs, load tables only
        for (var i = 0; i < paragraphs.length; i++) {
            var paragraph = paragraphs[i];
            if (paragraph.type == "Table") {
                paragraph.load("table/id");
                promise =  context.sync().then(function() {
                    console.log(paragraph.table.id);
                });
            }
        }
        return promise;
    })
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```
