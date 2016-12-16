# <a name="paragraph-object-javascript-api-for-word"></a>Paragraph 物件 (適用於 Word 的 JavaScript API)

代表選取範圍、範圍、內容控制項或文件內文中的單一段落。

_適用於：Word 2016、Word for iPad、Word for Mac、Word Online_

## <a name="properties"></a>屬性
| 屬性	     | 類型	   |描述
|:---------------|:--------|:----------|
|outlineLevel|int|取得或設定段落的大綱層級。|
|style|string|取得或設定段落所使用的樣式。這是預先安裝或自訂樣式的名稱。 [Word-Add-in-DocumentAssembly][paragraph.style]範例示範如何設定段落樣式。|
|文字|string|取得段落的文字。唯讀。|

## <a name="relationships"></a>關聯性
| 關聯性 | 類型	   |描述|
|:---------------|:--------|:----------|
|alignment|**Alignment**|取得或設定段落的對齊方式。此值可以是 "left"、"centered"、"right" 或 "justified"。|
|contentControls|[ContentControlCollection](contentcontrolcollection.md)|取得段落中內容控制項物件的集合。唯讀。|
|firstLineIndent|**float**|取得或設定首行縮排或首行凸排的值，以點為單位。使用正值可以設定首行縮排、使用負值可以設定首行凸排。|
|font|[Font](font.md)|取得段落的文字格式。使用此選項可取得及設定字型名稱、大小、色彩及其他屬性。唯讀。|
|inlinePictures|[InlinePictureCollection](inlinepicturecollection.md)|取得段落中 inlinePicture 物件的集合。集合不包含浮動圖像。唯讀。|
|leftIndent|**float**|取得或設定段落的左邊縮排值，以點為單位。|
|lineSpacing|**float**|取得或設定所指定段落的行距，以點為單位。在 Word UI 中，此值除以 12。|
|lineUnitAfter|**float**|取得或設定段落後的間距量，以格線為單位。|
|lineUnitBefore|**float**|取得或設定段落前的間距量，以格線為單位。|
|parentContentControl|[ContentControl](contentcontrol.md)|取得包含段落的內容控制項。如果沒有父代內容控制項，則傳回 null。唯讀。|
|rightIndent|**float**|取得或設定段落的右邊縮排值，以點為單位。|
|spaceAfter|**float**|取得或設定段落後的間距，以點為單位。|
|spaceBefore|**float**|取得或設定段落前的間距，以點為單位。|

## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述|
|:---------------|:--------|:----------|
|[clear()](#clear)|void|清除 paragraph 物件的內容。使用者可對已清除的內容執行復原作業。|
|[delete()](#delete)|void|刪除文件中的段落及其內容。|
|[getHtml()](#gethtml)|string|取得 paragraph 物件的 HTML 表示法。|
|[getOoxml()](#getooxml)|string|取得 paragraph 物件的 Office Open XML (OOXML) 表示法。|
|[insertBreak(breakType: BreakType, insertLocation: InsertLocation)](#insertbreakbreaktype-breaktype-insertlocation-insertlocation)|void|在指定的位置插入中斷符號。除了換行符號可以插入至任何 body 物件，其他中斷符號只能插入到主文件內文所包含的段落中。InsertLocation 值可以是 'After' 或 'Before'。|
|[insertContentControl()](#insertcontentcontrol)|[ContentControl](contentcontrol.md)|以 RTF 內容控制項圍繞 paragraph 物件。|
|[insertFileFromBase64(base64File: string, insertLocation:InsertLocation)](#insertfilefrombase64base64file-string-insertlocation-insertlocation)|[Range](range.md)|在目前段落的指定位置插入文件。InsertLocation 值可以是 'Start' 或 'End'。|
|[insertHtml(html: string, insertLocation:InsertLocation)](#inserthtmlhtml-string-insertlocation-insertlocation)|[Range](range.md)|在段落的指定位置插入 HTML。InsertLocation 值可以是 'Replace'、'Start' 或 'End'。|
|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation:InsertLocation)](#insertinlinepicturefrombase64base64encodedimage-string-insertlocation-insertlocation)|[InlinePicture](inlinepicture.md)|在段落的指定位置插入圖片。InsertLocation 值可以是 'Before'、'After'、'Start' 或 'End'。|
|[insertOoxml(ooxml: string, insertLocation:InsertLocation)](#insertooxmlooxml-string-insertlocation-insertlocation)|[Range](range.md)|在段落的指定位置插入 OOXML 或 wordProcessingML。InsertLocation 值可以是 'Replace'、'Start' 或 'End'。|
|[insertParagraph(paragraphText: string, insertLocation:InsertLocation)](#insertparagraphparagraphtext-string-insertlocation-insertlocation)|[Paragraph](paragraph.md)|在指定的位置插入段落。InsertLocation 值可以是 'Before' 或 'After'。|
|[insertText(text: string, insertLocation:InsertLocation)](#inserttexttext-string-insertlocation-insertlocation)|[Range](range.md)|在段落的指定位置插入文字。InsertLocation 值可以是 'Replace'、'Start' 或 'End'。|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|
|[search(searchText: string, searchOptions:ParamTypeStrings.SearchOptions)](#searchsearchtext-string-searchoptions-paramtypestringssearchoptions)|[SearchResultCollection](searchresultcollection.md)|以指定的 searchOptions 在 paragraph 物件的範圍中執行搜尋。搜尋結果將是 range 物件的集合。|
|[select(selectionMode: SelectionMode)](#selectselectionmode-selectionmode)|void|選取段落並將 Word UI 導覽至該處。選取模式可以是'Select'、'Start' 或 'End'。'Select' 為預設值。|

## <a name="method-details"></a>方法詳細資料

### <a name="clear"></a>clear()
清除 paragraph 物件的內容。使用者可對已清除的內容執行復原作業。

#### <a name="syntax"></a>語法
```js
paragraphObject.clear();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
void

#### <a name="examples"></a>範例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the style property for all of the paragraphs.
    context.load(paragraphs, 'style');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Queue a command to clear the contents of the first paragraph.
        paragraphs.items[0].clear();

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Cleared the contents of the first paragraph.');
        });
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="delete"></a>delete()
刪除文件中的段落及其內容。

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
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the text property for all of the paragraphs.
    context.load(paragraphs, 'text');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Queue a command to delete the first paragraph.
        paragraphs.items[0].delete();

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Deleted the first paragraph.');
        });
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="gethtml"></a>getHtml()
取得 paragraph 物件的 HTML 表示法。

#### <a name="syntax"></a>語法
```js
paragraphObject.getHtml();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
字串

#### <a name="examples"></a>範例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the style property for all of the paragraphs.
    context.load(paragraphs, 'style');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Queue a a set of commands to get the HTML of the first paragraph.
        var html = paragraphs.items[0].getHtml();

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Paragraph HTML: ' + html.value);
        });
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="getooxml"></a>getOoxml()
取得 paragraph 物件的 Office Open XML (OOXML) 表示法。

#### <a name="syntax"></a>語法
```js
paragraphObject.getOoxml();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
字串

#### <a name="examples"></a>範例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the style property for the top 2 paragraphs.
    context.load(paragraphs, {select: 'style', top: 2} );

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Queue a a set of commands to get the OOXML of the first paragraph.
        var ooxml = paragraphs.items[0].getOoxml();

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Paragraph OOXML: ' + ooxml.value);
        });
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="insertbreakbreaktype-breaktype-insertlocation-insertlocation"></a>insertBreak(breakType: BreakType, insertLocation: InsertLocation)
在指定的位置插入中斷符號。除了換行符號可以插入至任何 body 物件，其他中斷符號只能插入到主文件內文所包含的段落中。InsertLocation 值可以是 'Before' 或 'After'。

#### <a name="syntax"></a>語法
```js
paragraphObject.insertBreak(breakType, insertLocation);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|breakType|BreakType|必要。要加入至文件的中斷類型。|
|insertLocation|InsertLocation|必要。此值可以是 'Before' 或 'After'。|

#### <a name="returns"></a>傳回
void

#### <a name="additional-details"></a>其他詳細資料
您不能在頁首、頁尾、註腳、章節附註、註解和文字方塊中插入中斷符號。

#### <a name="examples"></a>範例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the style property for the top 2 paragraphs.
    // We never perform an empty load. We always must request a property.
    context.load(paragraphs, {select: 'style', top: 2} );

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];

        // Queue a command to insert a page break after the first paragraph.
        paragraph.insertBreak('page', 'After');

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Inserted a page break after the paragraph.');
        });
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="insertcontentcontrol"></a>insertContentControl()
以 RTF 內容控制項圍繞 paragraph 物件。

#### <a name="syntax"></a>語法
```js
paragraphObject.insertContentControl();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
[ContentControl](contentcontrol.md)

#### <a name="examples"></a>範例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the style property for the top 2 paragraphs.
    // We never perform an empty load. We always must request a property.
    context.load(paragraphs, {select: 'style', top: 2} );

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];

        // Queue a command to wrap the first paragraph in a rich text content control.
        paragraph.insertContentControl();

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Wrapped the first paragraph in a content control.');
        });
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

#### <a name="additional-information"></a>其他資訊
[Word-Add-in-DocumentAssembly][paragraph.insertContentControl]範例示範如何使用 insertContentControl 方法。

### <a name="insertfilefrombase64base64file-string-insertlocation-insertlocation"></a>insertFileFromBase64(base64File: string, insertLocation:InsertLocation)
在目前段落的指定位置插入文件。InsertLocation 值可以是 ‘Replace’、'Start' 或 'End'。

#### <a name="syntax"></a>語法
```js
paragraphObject.insertFileFromBase64(base64File, insertLocation);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|base64File|string|必要。要插入的檔案 base64 編碼檔案內容。|
|insertLocation|InsertLocation|必要。此值可以是 ‘Replace’、'Start' 或 'End'。|

#### <a name="returns"></a>傳回
[Range](range.md)

#### <a name="examples"></a>範例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the style property for all of the paragraphs.
    context.load(paragraphs, 'style');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];

        // Queue a command to insert base64 encoded .docx at the beginning of the first paragraph.
        // This won't work unless you have a definition for getBase64().
        paragraph.insertFileFromBase64(getBase64(), Word.InsertLocation.start);

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Inserted base64 encoded content at the beginning of the first paragraph.');
        });
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="inserthtmlhtml-string-insertlocation-insertlocation"></a>insertHtml(html: string, insertLocation:InsertLocation)
在段落的指定位置插入 HTML。InsertLocation 值可以是 'Replace'、'Start' 或 'End'。

#### <a name="syntax"></a>語法
```js
paragraphObject.insertHtml(html, insertLocation);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|HTML|string|必要。要插入至段落的 HTML。|
|insertLocation|InsertLocation|必要。此值可以是 'Replace'、'Start' 或 'End'。|

#### <a name="returns"></a>傳回
[Range](range.md)

#### <a name="examples"></a>範例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the style property for the top 2 paragraphs.
    // We never perform an empty load. We always must request a property.
    context.load(paragraphs, {select: 'style', top: 2} );

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];

        // Queue a command to insert HTML content at the end of the first paragraph.
        paragraph.insertHtml('<strong>Inserted HTML.</strong>', Word.InsertLocation.end);

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Inserted HTML content at the end of the first paragraph.');
        });
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});

```

### <a name="insertinlinepicturefrombase64base64encodedimage-string-insertlocation-insertlocation"></a>insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation:InsertLocation)
在段落的指定位置插入圖片。InsertLocation 值可以是 'Replace'、'Start' 或 'End'。

#### <a name="syntax"></a>語法
```js
paragraphObject.insertInlinePictureFromBase64(base64EncodedImage, insertLocation);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|base64EncodedImage|string|必要。要插入至段落的 HTML。|
|insertLocation|InsertLocation|必要。此值可以是 'Replace'、'Start' 或 'End'。|

#### <a name="returns"></a>傳回
[InlinePicture](inlinepicture.md)

#### <a name="examples"></a>範例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the style property for all of the paragraphs.
    context.load(paragraphs, 'style');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];

        var b64encodedImg = "iVBORw0KGgoAAAANSUhEUgAAAB4AAAANCAIAAAAxEEnAAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAACFSURBVDhPtY1BEoQwDMP6/0+XgIMTBAeYoTqso9Rkx1zG+tNj1H94jgGzeNSjteO5vtQQuG2seO0av8LzGbe3anzRoJ4ybm/VeKEerAEbAUpW4aWQCmrGFWykRzGBCnYy2ha3oAIq2MloW9yCCqhgJ6NtcQsqoIKdjLbFLaiACnYyf2fODbrjZcXfr2F4AAAAAElFTkSuQmCC";

        // Queue a command to insert a base64 encoded image at the beginning of the first paragraph.
        paragraph.insertInlinePictureFromBase64(b64encodedImg, Word.InsertLocation.start);

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Added an image to the first paragraph.');
        });
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

#### <a name="additional-information"></a>其他資訊
[Word-Add-in-DocumentAssembly][paragraph.insertpicture] 範例提供如何將影像插入段落的另一個範例。

### <a name="insertooxmlooxml-string-insertlocation-insertlocation"></a>insertOoxml(ooxml: string, insertLocation:InsertLocation)
在段落的指定位置插入 OOXML 或 wordProcessingML。InsertLocation 值可以是 'Replace'、'Start' 或 'End'。

#### <a name="syntax"></a>語法
```js
paragraphObject.insertOoxml(ooxml, insertLocation);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|ooxml|string|必要。要插入至段落的 OOXML 或 wordProcessingML。|
|insertLocation|InsertLocation|必要。此值可以是 'Replace'、'Start' 或 'End'。|

#### <a name="returns"></a>傳回
[Range](range.md)

#### <a name="known-issues"></a>已知問題
這個方法會在 Word Online 中導致較長的延遲時間，可能會影響增益集的使用者經驗。我們建議您只有在沒有其他解決方案可以使用時，才使用這個方法。 

#### <a name="examples"></a>範例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the style property for the top 2 paragraphs.
    // We never perform an empty load. We always must request a property.
    context.load(paragraphs, {select: 'style', top: 2} );

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];

        // Queue a command to insert Ooxml content into the first paragraph.
        var ooxmlContent = "<pkg:package xmlns:pkg='http://schemas.microsoft.com/office/2006/xmlPackage'><pkg:part pkg:name='/_rels/.rels' pkg:contentType='application/vnd.openxmlformats-package.relationships+xml' pkg:padding='512'><pkg:xmlData><Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name='/word/document.xml' pkg:contentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'><pkg:xmlData><w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' ><w:body><w:p><w:pPr><w:spacing w:before='360' w:after='0' w:line='480' w:lineRule='auto'/><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr></w:pPr><w:r><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr><w:t>This text has formatting directly applied to achieve its font size, color, line spacing, and paragraph spacing.</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>";
        paragraph.insertOoxml(ooxmlContent, Word.InsertLocation.end);

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Inserted OOXML at the end of the first paragraph.');
        });
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

#### <a name="additional-information"></a>其他資訊
如需使用 OOXML 的指示，請閱讀[使用 Office Open XML 為 Word 建立更佳的增益集](https://msdn.microsoft.com/en-us/library/office/dn423225.aspx)。

### <a name="insertparagraphparagraphtext-string-insertlocation-insertlocation"></a>insertParagraph(paragraphText: string, insertLocation:InsertLocation)
在指定的位置插入段落。InsertLocation 值可以是 'Before' 或 'After'。

#### <a name="syntax"></a>語法
```js
paragraphObject.insertParagraph(paragraphText, insertLocation);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|paragraphText|string|必要。要插入的段落文字。|
|insertLocation|InsertLocation|必要。此值可以是 'Before' 或 'After'。|

#### <a name="returns"></a>傳回
[Paragraph](paragraph.md)

#### <a name="examples"></a>範例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the style property for the top 2 paragraphs.
    // We never perform an empty load. We always must request a property.
    context.load(paragraphs, {select: 'style', top: 2} );

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];

        // Queue a command to insert the paragraph after the current paragraph.
        paragraph.insertParagraph('Content of a new paragraph', Word.InsertLocation.after);

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Inserted a new paragraph at the end of the first paragraph.');
        });
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="inserttexttext-string-insertlocation-insertlocation"></a>insertText(text: string, insertLocation:InsertLocation)
在段落的指定位置插入文字。InsertLocation 值可以是 'Replace'、'Start' 或 'End'。

#### <a name="syntax"></a>語法
```js
paragraphObject.insertText(text, insertLocation);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|文字|string|必要。要插入的文字。|
|insertLocation|InsertLocation|必要。此值可以是 'Replace'、'Start' 或 'End'。|

#### <a name="returns"></a>傳回
[Range](range.md)

#### <a name="examples"></a>範例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the style property for the top 2 paragraphs.
    // We never perform an empty load. We always must request a property.
    context.load(paragraphs, {select: 'style', top: 2} );

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];

        // Queue a command to insert text into the end of the paragraph.
        paragraph.insertText('New text inserted into the paragraph.', Word.InsertLocation.end);

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Inserted text at the end of the first paragraph.');
        });
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
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
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|param|物件|選用。接受參數與關聯性名稱，做為分隔字串或陣列。或者提供 [loadOption](loadoption.md) 物件。|

#### <a name="returns"></a>傳回
void

#### <a name="examples"></a>範例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the style property for the top 2 paragraphs.
    // We never perform an empty load. We always must request a property.
    context.load(paragraphs, {select: 'style', top: 2} );

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];

        // Queue a command to load font information for the paragraph.
        context.load(paragraph, 'font/size, font/name, font/color');

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            // Show the results of the load method. Here we show the
            // property values on the paragraph object. Note that we
            // requested the style property in the first load command.
            var results = "<strong>Paragraph</strong>--" +
                          "--Font size: " + paragraph.font.size +
                          "--Font name: " + paragraph.font.name +
                          "--Font color: " + paragraph.font.color +
                          "--Style: " + paragraph.style;

            console.log(results);
        });
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="searchsearchtext-string-searchoptions-paramtypestringssearchoptions"></a>search(searchText: string, searchOptions:ParamTypeStrings.SearchOptions)
以指定的 searchOptions 在 paragraph 物件的範圍中執行搜尋。搜尋結果將是 range 物件的集合。

#### <a name="syntax"></a>語法
```js
paragraphObject.search(searchText, searchOptions);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|searchText|string|必要。搜尋文字。|
|[searchOptions](searchoptions.md)|ParamTypeStrings.SearchOptions|選用。搜尋選項。|

#### <a name="returns"></a>傳回
[SearchResultCollection](searchresultcollection.md)

### <a name="selectselectionmode-selectionmode"></a>select(selectionMode: SelectionMode)
選取段落並將 Word UI 導覽至該處。

#### <a name="syntax"></a>語法
```js
paragraphObject.select(selectionMode);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|selectionMode|SelectionMode|選用。選取模式可以是'Select'、'Start' 或 'End'。'Select' 為預設值。|

#### <a name="returns"></a>傳回
void

#### <a name="examples"></a>範例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;

    // Queue a commmand to load the style property for all of the paragraphs.
    context.load(paragraphs, 'style');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        // Queue a command to get the last paragraph a create a
        // proxy paragraph object.
        var paragraph = paragraphs.items[paragraphs.items.length - 1];

        // Queue a command to select the paragraph. The Word UI will
        // move to the selected paragraph.
        paragraph.select();

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Selected the last paragraph.');
        });
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

## <a name="support-details"></a>支援詳細資料
在執行階段檢查使用[需求集](../office-add-in-requirement-sets.md)以確認您的應用程式受到 Word 主應用程式版本的支援。如需有關 Office 主應用程式及伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


[paragraph.insertContentControl]: https://github.com/OfficeDev/Word-Add-in-DocumentAssembly/blob/master/WordAPIDocAssemblySampleWeb/App/Home/Home.js#L161 "插入內容控制項"
[paragraph.style]: https://github.com/OfficeDev/Word-Add-in-DocumentAssembly/blob/master/WordAPIDocAssemblySampleWeb/App/Home/Home.js#L172 "設定樣式"
[paragraph.insertpicture]: https://github.com/OfficeDev/Word-Add-in-DocumentAssembly/blob/master/WordAPIDocAssemblySampleWeb/App/Home/Home.js#L236 "插入圖片"
