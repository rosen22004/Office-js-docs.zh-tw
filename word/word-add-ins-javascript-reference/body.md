# Body 物件 (適用於 Word 的 JavaScript API)

代表文件或區段的內文。

_適用版本：Word 2016、Word for iPad、Word for Mac_

## 屬性
| 屬性	   | 類型	|說明
|:---------------|:--------|:----------|
|Style|string|取得或設定內文所使用的樣式。這是預先安裝或自訂樣式的名稱。|
|text|string|取得內文的文字。可以使用 insertText 方法來插入文字。唯讀。|

_請參閱屬性存取[範例。](#property-access-examples)_

## 關聯性
| 關聯性 | 類型	|說明|
|:---------------|:--------|:----------|
|contentControls|[ContentControlCollection](contentcontrolcollection.md)|取得內文中 RTF 內容控制項物件的集合。唯讀。|
|font|[Font](font.md)|取得內文的文字格式。使用此選項可取得及設定字型名稱、大小、色彩及其他屬性。唯讀。|
|inlinePictures|[InlinePictureCollection](inlinepicturecollection.md)|取得內文中 inlinePicture 物件的集合。集合不包含浮動圖像。唯讀。|
|paragraphs|[ParagraphCollection](paragraphcollection.md)|取得內文中 paragraph 物件的集合。唯讀。|
|parentContentControl|[ContentControl](contentcontrol.md)|取得包含內文的內容控制項。如果沒有父代內容控制項，則傳回 null。唯讀。|

## 方法

| 方法		   | 傳回類型	|說明|
|:---------------|:--------|:----------|
|[clear()](#clear)|void|清除 body 物件的內容。使用者可對已清除的內容執行復原作業。|
|[getHtml()](#gethtml)|string|取得 body 物件的 HTML 表示法。|
|[getOoxml()](#getooxml)|string|取得 body 物件的 OOXML (Office Open XML) 表示法。|
|[insertBreak(breakType: BreakType, insertLocation: InsertLocation)](#insertbreakbreaktype-breaktype-insertlocation-insertlocation)|void|在指定的位置插入中斷符號。除了換行符號可以插入至任何 body 物件，其他中斷符號只能插入到主文件內文中。InsertLocation 值可以是 'Start' 或 'End'。|
|[insertContentControl()](#insertcontentcontrol)|[ContentControl](contentcontrol.md)|以 RTF 內容控制項圍繞 body 物件。|
|[insertFileFromBase64(base64File: string, insertLocation:InsertLocation)](#insertfilefrombase64base64file-string-insertlocation-insertlocation)|[Range](range.md)|在內文的指定位置插入文件。InsertLocation 值可以是 'Replace'、'Start' 或 'End'。|
|[insertHtml(html: string, insertLocation:InsertLocation)](#inserthtmlhtml-string-insertlocation-insertlocation)|[Range](range.md)|在指定的位置插入 HTML。InsertLocation 值可以是 'Replace'、'Start' 或 'End'。|
|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation:InsertLocation)](#insertInlinePictureFromBase64base64EncodedImage-string-insertlocation-insertlocation)|[InlinePicture](inlinepicture.md)|在內文的指定位置插入圖片。InsertLocation 值可以是 'Start' 或 'End'。 |
|[insertOoxml(ooxml: string, insertLocation:InsertLocation)](#insertooxmlooxml-string-insertlocation-insertlocation)|[Range](range.md)|在指定的位置插入 OOXML 或 wordProcessingML。InsertLocation 值可以是 'Replace'、'Start' 或 'End'。|
|[insertParagraph(paragraphText: string, insertLocation:InsertLocation)](#insertparagraphparagraphtext-string-insertlocation-insertlocation)|[Paragraph](paragraph.md)|在指定的位置插入段落。InsertLocation 值可以是 'Start' 或 'End'。|
|[insertText(text: string, insertLocation:InsertLocation)](#inserttexttext-string-insertlocation-insertlocation)|[Range](range.md)|在內文的指定位置插入文字。InsertLocation 值可以是 'Replace'、'Start' 或 'End'。|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|
|[search(searchText: string, searchOptions:ParamTypeStrings.SearchOptions)](#searchsearchtext-string-searchoptions-paramtypestringssearchoptions)|[SearchResultCollection](searchresultcollection.md)|以指定的 searchOptions 在 body 物件的範圍中執行搜尋。搜尋結果將是 range 物件的集合。|
|[select(selectionMode: SelectionMode)](#selectselectionmode-selectionmode)|void|選取內文並將 Word UI 導覽至該處。SelectionMode 值可以是 'Select'、'Start' 或 'End'。|

## 方法詳細資料

### clear()
清除 body 物件的內容。使用者可對已清除的內容執行復原作業。

#### 語法
```js
bodyObject.clear();
```

#### 參數
無

#### 傳回
void

#### 範例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the document body.
    var body = context.document.body;
    
    // Queue a commmand to clear the contents of the body.
    body.clear();
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Cleared the body contents.');
    });  
})
.catch(function (error) {
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

[Silly stories](https://aka.ms/sillystorywordaddin) 增益集範例示範如何使用 **clear** 方法清除文件的內容。

### getHtml()
取得 body 物件的 HTML 表示法。

#### 語法
```js
bodyObject.getHtml();
```

#### 參數
無

#### 會傳回
字串

#### 範例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the document body.
    var body = context.document.body;
    
    // Queue a commmand to get the HTML contents of the body.
    var bodyHTML = body.getHtml();
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log("Body HTML contents: " + bodyHTML.value);
    });  
})
.catch(function (error) {
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### getOoxml()
取得 body 物件的 OOXML (Office Open XML) 表示法。

#### 語法
```js
bodyObject.getOoxml();
```

#### 參數
無

#### 會傳回
字串

#### 範例 
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the document body.
    var body = context.document.body;
    
    // Queue a commmand to get the OOXML contents of the body.
    var bodyOOXML = body.getOoxml();
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log("Body OOXML contents: " + bodyOOXML.value);
    });  
})
.catch(function (error) {
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### insertBreak(breakType: BreakType, insertLocation: InsertLocation)
在指定的位置插入中斷符號。除了換行符號可以插入至任何 body 物件，其他中斷符號只能插入到主文件內文中。InsertLocation 值可以是 'Start' 或 'End'。

#### 語法
```js
bodyObject.insertBreak(breakType, insertLocation);
```

#### 參數
| 參數	   | 類型	|說明|
|:---------------|:--------|:----------|
|breakType|BreakType|必要。要加入至內文的中斷類型。|
|insertLocation|InsertLocation|必要。此值可以是 'Start' 或 'End'。|

#### 傳回
void

#### 其他詳細資料
除了換行符號以外，您不能在頁首、頁尾、註腳、章節附註、註解和文字方塊中插入其他中斷符號。  

#### 範例
```js
// Run a batch operation against the Word object model.
Word.run(function (ctx) {
    
    // Create a proxy object for the document body.
    var body = ctx.document.body;
    
    // Queue a commmand to insert a page break at the start of the document body.
    body.insertBreak(Word.BreakType.page, Word.InsertLocation.start);
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return ctx.sync().then(function () {
        console.log('Added a page break at the start of the document body.');
    });  
})
.catch(function (error) {
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```
### insertContentControl()
以 RTF 內容控制項圍繞 body 物件。

#### 語法
```js
bodyObject.insertContentControl();
```

#### 參數
無

#### 傳回
[ContentControl](contentcontrol.md)

#### 範例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the document body.
    var body = context.document.body;
    
    // Queue a commmand to wrap the body in a content control.
    body.insertContentControl();
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Wrapped the body in a content control.');
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```
### insertFileFromBase64(base64File: string, insertLocation:InsertLocation)
在內文的指定位置插入文件。InsertLocation 值可以是 'Replace'、'Start' 或 'End'。

#### 語法
```js
bodyObject.insertFileFromBase64(base64File, insertLocation);
```

#### 參數
| 參數	   | 類型	|說明|
|:---------------|:--------|:----------|
|base64File|string|必要。要插入的 base64 編碼檔案內容。|
|insertLocation|InsertLocation|必要。此值可以是 'Replace'、'Start' 或 'End'。|

#### 傳回
[Range](range.md)

#### 範例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the document body.
    var body = context.document.body;
    
    // Queue a commmand to insert base64 encoded .docx at the beginning of the content body.
    // You will need to implement getBase64() to pass in a string of a base64 encoded docx file.
    body.insertFileFromBase64(getBase64(), Word.InsertLocation.start);
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Added base64 encoded text to the beginning of the document body.');
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

[Silly stories](https://aka.ms/sillystorywordaddin) 增益集範例示範如何使用 **insertFileFromBase64** 方法從服務插入 docx 檔案。

### insertHtml(html: string, insertLocation:InsertLocation)
在指定的位置插入 HTML。InsertLocation 值可以是 'Replace'、'Start' 或 'End'。

#### 語法
```js
bodyObject.insertHtml(html, insertLocation);
```

#### 參數
| 參數	   | 類型	|說明|
|:---------------|:--------|:----------|
|HTML|string|必要。要插入至文件的 HTML。|
|insertLocation|InsertLocation|必要。此值可以是 'Replace'、'Start' 或 'End'。|

#### 傳回
[Range](range.md)

#### 範例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the document body.
    var body = context.document.body;
    
    // Queue a commmand to insert HTML in to the beginning of the body.
    body.insertHtml('<strong>This is text inserted with body.insertHtml()</strong>', Word.InsertLocation.start);
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('HTML added to the beginning of the document body.');
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation:InsertLocation)
在內文的指定位置插入圖片。InsertLocation 值可以是 'Start' 或 'End'。

#### 語法
bodyObject.insertInlinePictureFromBase64(image, insertLocation);

#### 參數
| 參數	   | 類型	|說明|
|:---------------|:--------|:----------|
|base64EncodedImage|string|必要。要插入至內文的 base64 編碼影像。|
|insertLocation|InsertLocation|必要。此值可以是 'Start' 或 'End'。|

#### 傳回
[InlinePicture](inlinepicture.md)

### insertOoxml(ooxml: string, insertLocation:InsertLocation)
在指定的位置插入 OOXML 或 wordProcessingML。InsertLocation 值可以是 'Replace'、'Start' 或 'End'。

#### 語法
```js
bodyObject.insertOoxml(ooxml, insertLocation);
```

#### 參數
| 參數	   | 類型	|說明|
|:---------------|:--------|:----------|
|ooxml|string|必要。要插入的 OOXML 或 wordProcessingML。|
|insertLocation|InsertLocation|必要。此值可以是 'Replace'、'Start' 或 'End'。|

#### 傳回
[Range](range.md)

#### 範例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the document body.
    var body = context.document.body;
    
    // Queue a commmand to insert OOXML in to the beginning of the body.
    body.insertOoxml("<pkg:package xmlns:pkg='http://schemas.microsoft.com/office/2006/xmlPackage'><pkg:part pkg:name='/_rels/.rels' pkg:contentType='application/vnd.openxmlformats-package.relationships+xml' pkg:padding='512'><pkg:xmlData><Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name='/word/document.xml' pkg:contentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'><pkg:xmlData><w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' ><w:body><w:p><w:pPr><w:spacing w:before='360' w:after='0' w:line='480' w:lineRule='auto'/><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr></w:pPr><w:r><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr><w:t>This text has formatting directly applied to achieve its font size, color, line spacing, and paragraph spacing.</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>", Word.InsertLocation.start);
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('OOXML added to the beginning of the document body.');
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

#### 其他資訊
如需使用 OOXML 的指示，請閱讀[使用 Office Open XML 為 Word 建立更佳的增益集](https://msdn.microsoft.com/en-us/library/office/dn423225.aspx)。[Word-Add-in-DocumentAssembly][body.insertOoxml] 範例示範如何使用此 API 來組合文件。

### insertParagraph(paragraphText: string, insertLocation:InsertLocation)
在指定的位置插入段落。InsertLocation 值可以是 'Start' 或 'End'。

#### 語法
```js
bodyObject.insertParagraph(paragraphText, insertLocation);
```

#### 參數
| 參數	   | 類型	|說明|
|:---------------|:--------|:----------|
|paragraphText|string|必要。要插入的段落文字。|
|insertLocation|InsertLocation|必要。此值可以是 'Start' 或 'End'。|

#### 傳回
[段落](paragraph.md)

#### 範例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the document body.
    var body = context.document.body;
    
    // Queue a commmand to insert the paragraph at the end of the document body.
    body.insertParagraph('Content of a new paragraph', Word.InsertLocation.end);
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Paragraph added at the end of the document body.');
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

#### 其他資訊
[Word-Add-in-DocumentAssembly][body.insertParagraph] 範例示範如何使用 insertParagraph 方法來組合文件。

### insertText(text: string, insertLocation:InsertLocation)
在內文的指定位置插入文字。InsertLocation 值可以是 'Replace'、'Start' 或 'End'。

#### 語法
```js
bodyObject.insertText(text, insertLocation);
```

#### 參數
| 參數	   | 類型	|說明|
|:---------------|:--------|:----------|
|文字|string|必要。要插入的文字。|
|insertLocation|InsertLocation|必要。此值可以是 'Replace'、'Start' 或 'End'。|

#### 傳回
[Range](range.md)

#### 範例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the document body.
    var body = context.document.body;
    
    // Queue a commmand to insert text in to the beginning of the body.
    body.insertText('This is text inserted with body.insertText()', Word.InsertLocation.start);
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Text added to the beginning of the document body.');
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```
### load(param: object)
以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。

#### 語法
```js
object.load(param);
```

#### 參數
| 參數	   | 類型	|說明|
|:---------------|:--------|:----------|
|param|object|選用。接受參數與關聯性名稱，做為分隔字串或陣列。或者提供 [loadOption](loadoption.md) 物件。|

#### 傳回
void

#### 範例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the document body.
    var body = context.document.body;
    
    // Queue a commmand to load font and style information for the document body.
    context.load(body, 'font/size, font/name, font/color, style');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        // Show the results of the load method. Here we show the
        // property values on the body object.
        var results = 'Font size: ' + body.font.size +
                      '; Font name: ' + body.font.name +
                      '; Font color: ' + body.font.color +
                      '; Body style: ' + body.style;

        console.log(results);
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```
### search(searchText: string, searchOptions:ParamTypeStrings.SearchOptions)
以指定的搜尋選項在 body 物件的範圍中執行搜尋。搜尋結果將是 range 物件的集合。

#### 語法
```js
bodyObject.search(searchText, searchOptions);
```

#### 參數
| 參數	   | 類型	|說明|
|:---------------|:--------|:----------|
|searchText|string|必要。搜尋文字。|
|[searchOptions](searchoptions.md)|ParamTypeStrings.SearchOptions|選用。搜尋選項。|

#### 傳回
[SearchResultCollection](searchresultcollection.md)

#### 範例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the document body.
    var body = context.document.body;
    
    // Queue a commmand to search the document.
    var searchResults = context.document.body.search('video', {matchCase: false});

    // Queue a commmand to load the results.
    context.load(searchResults, 'text, font');

    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        var results = 'Found count: ' + searchResults.items.length + 
                      '; we highlighted the results.';

        // Queue a command to change the font for each found item. 
        for (var i = 0; i < searchResults.items.length; i++) {
          searchResults.items[i].font.color = '#FF0000'    // Change color to Red
          searchResults.items[i].font.highlightColor = '#FFFF00';
          searchResults.items[i].font.bold = true;
        }
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
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

#### 其他資訊
[Word-Add-in-DocumentAssembly][body.search] 範例提供如何搜尋文件的另一個範例。

### select(selectionMode: SelectionMode)
選取內文並將 Word UI 導覽至該處。SelectionMode 值可以是 'Select'、'Start' 或 'End'。

#### 語法
```js
bodyObject.select(selectionMode);
```

#### 參數
| 參數	   | 類型	|說明|
|:---------------|:--------|:----------|
|selectionMode|SelectionMode|選用。選取模式可以是 'Select'、'Start' 或 'End'。'Select' 為預設值。|

#### 傳回
void

#### 範例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the document body.
    var body = context.document.body;
    
    // Queue a commmand to select the document body. The Word UI will 
    // move to the selected document body.
    body.select();
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Selected the document body.');
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

## 屬性存取範例

### 取得 body 物件的文字屬性
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the document body.
    var body = context.document.body;
    
    // Queue a commmand to load the text in document body.
    context.load(body, 'text');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log("Body contents: " + body.text);
    });  
})
.catch(function (error) {
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```
### 取得 body 物件的樣式及字型大小、字型名稱和字型色彩屬性。

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the document body.
    var body = context.document.body;
    
    // Queue a commmand to load font and style information for the document body.
    context.load(body, 'font/size, font/name, font/color, style');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        // Show the results of the load method. Here we show the
        // property values on the body object.
        var results = 'Font size: ' + body.font.size +
                      '; Font name: ' + body.font.name +
                      '; Font color: ' + body.font.color +
                      '; Body style: ' + body.style;

        console.log(results);
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

## 支援詳細資料

在執行階段檢查使用[需求集](https://msdn.microsoft.com/EN-US/library/office/mt590206.aspx)以確認您的應用程式受到 Word 主應用程式版本的支援。如需有關 Office 主應用程式及伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](https://msdn.microsoft.com/EN-US/library/office/dn833104.aspx)。 


[body.insertOoxml]: https://github.com/OfficeDev/Word-Add-in-DocumentAssembly/blob/master/WordAPIDocAssemblySampleWeb/App/Home/Home.js#L127 "insert OOXML"[body.insertParagraph]: https://github.com/OfficeDev/Word-Add-in-DocumentAssembly/blob/master/WordAPIDocAssemblySampleWeb/App/Home/Home.js#L153 "insert paragraph" [body.search]: https://github.com/OfficeDev/Word-Add-in-DocumentAssembly/blob/master/WordAPIDocAssemblySampleWeb/App/Home/Home.js#L261 "body search"
