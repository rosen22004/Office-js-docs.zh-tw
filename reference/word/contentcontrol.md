# <a name="contentcontrol-object-(javascript-api-for-word)"></a>ContentControl 物件 (適用於 Word 的 JavaScript API)

代表內容控制項。內容控制項是指文件中具有界限且可能具有標籤的區域，這些區域會做為特定內容類型的容器。個別的內容控制項可能含有內容，例如影像、表格或格式化文字的段落。目前僅支援 RTF 內容控制項。

_適用於：Word 2016、Word for iPad、Word for Mac、Word Online_

## <a name="properties"></a>屬性
| 屬性	     | 類型	   |描述
|:---------------|:--------|:----------|
|cannotDelete|bool|取得或設定值，指出使用者是否可以刪除內容控制項。與 removeWhenEdited 互斥。|
|cannotEdit|bool|取得或設定值，指出使用者是否可以編輯內容控制項的內容。|
|color|string|取得或設定內容控制項的色彩。Color 以 "#RRGGBB" 格式設定，或使用色彩名稱。|
|placeholderText|string|取得或設定內容控制項的預留位置文字。內容控制項為空時，將顯示暗灰色文字。|
|removeWhenEdited|bool|取得或設定值，指出在編輯內容控制項後是否可以將其移除。與 cannotDelete 互斥。|
|Style|string|取得或設定內容控制項所使用的樣式。這是預先安裝或自訂樣式的名稱。|
|Tag|string|取得或設定用以識別內容控制項的標記。[Silly stories](https://aka.ms/sillystorywordaddin) 增益集範例示範如何使用 **tag** 屬性。|
|文字|string|取得內容控制項的文字。唯讀。|
|標題|string|取得或設定內容控制項的標題。|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
| 關聯性 | 類型	   |描述|
|:---------------|:--------|:----------|
|appearance|**ContentControlAppearance**|取得或設定內容控制項的外觀。此值可以是 'boundingBox'、'tags' 或 'hidden'。|
|contentControls|[ContentControlCollection](contentcontrolcollection.md)|取得內容控制項中內容控制項物件的集合。唯讀。|
|font|[Font](font.md)|取得內容控制項的文字格式。使用此選項可取得及設定字型名稱、大小、色彩及其他屬性。唯讀。|
|id|**uint**|取得代表內容控制項識別碼的整數。唯讀。|
|inlinePictures|[InlinePictureCollection](inlinepicturecollection.md)|取得內容控制項中 inlinePicture 物件的集合。集合不包含浮動圖像。唯讀。|
|paragraphs|[ParagraphCollection](paragraphcollection.md)|取得內容控制項中 paragraph 物件的集合。唯讀。|
|parentContentControl|[ContentControl](contentcontrol.md)|取得包含內容控制項的內容控制項。如果沒有父代內容控制項，則傳回 null。唯讀。|
|類型|**ContentControlType**|取得內容控制項類型。目前僅支援 RTF 內容控制項。唯讀。|

## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述|
|:---------------|:--------|:----------|
|[clear()](#clear)|void|清除內容控制項的內容。使用者可對已清除的內容執行復原作業。|
|[delete(keepContent: bool)](#deletekeepcontent-bool)|void|刪除內容控制項和其內容。如果 keepContent 設定為 true，則不能刪除內容。|
|[getHtml()](#gethtml)|string|取得內容控制項物件的 HTML 表示法。|
|[getOoxml()](#getooxml)|string|取得內容控制項物件的 Office Open XML (OOXML) 表示法。|
|[insertBreak(breakType: BreakType, insertLocation: InsertLocation)](#insertbreakbreaktype-breaktype-insertlocation-insertlocation)|void|在指定的位置插入中斷符號。除了換行符號可以插入至任何 body 物件，其他中斷符號只能插入到主文件內文所包含的物件中。InsertLocation 值可以是 'Before'、'After'、'Start' 或 'End'。|
|[insertFileFromBase64(base64File: string, insertLocation:InsertLocation)](#insertfilefrombase64base64file-string-insertlocation-insertlocation)|[Range](range.md)|在目前內容控制項的指定位置插入文件。InsertLocation 值可以是 'Replace'、'Start' 或 'End'。|
|[insertHtml(html: string, insertLocation:InsertLocation)](#inserthtmlhtml-string-insertlocation-insertlocation)|[Range](range.md)|在內容控制項的指定位置插入 HTML。InsertLocation 值可以是 'Replace'、'Start' 或 'End'。|
|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation:InsertLocation)](#insertInlinePictureFromBase64base64EncodedImage-string-insertlocation-insertlocation)|[InlinePicture](inlinepicture.md)|在內容控制項的指定位置插入內嵌圖片。InsertLocation 值可以是 'Replace'、'Start' 或 'End'。 |
|[insertOoxml(ooxml: string, insertLocation:InsertLocation)](#insertooxmlooxml-string-insertlocation-insertlocation)|[Range](range.md)|在內容控制項的指定位置插入 OOXML 或 wordProcessingML。InsertLocation 值可以是 'Replace'、'Start' 或 'End'。|
|[insertParagraph(paragraphText: string, insertLocation:InsertLocation)](#insertparagraphparagraphtext-string-insertlocation-insertlocation)|[Paragraph](paragraph.md)|在指定的位置插入段落。InsertLocation 值可以是 'Before'、'After'、'Start' 或 'End'。|
|[insertText(text: string, insertLocation:InsertLocation)](#inserttexttext-string-insertlocation-insertlocation)|[Range](range.md)|在內容控制項的指定位置插入文字。InsertLocation 值可以是 'Replace'、'Start' 或 'End'。|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|
|[search(searchText: string, searchOptions:ParamTypeStrings.SearchOptions)](#searchsearchtext-string-searchoptions-paramtypestringssearchoptions)|[SearchResultCollection](searchresultcollection.md)|以指定的 searchOptions 在內容控制項物件的範圍中執行搜尋。搜尋結果將是 range 物件的集合。|
|[select(selectionMode: SelectionMode)](#selectselectionmode-selectionmode)|void|選取內容控制項。這會讓 Word 捲動至該選取範圍。選取模式可以是 'Select'、'Start' 或 'End'。|

## <a name="method-details"></a>方法詳細資料

### <a name="clear()"></a>clear()
清除內容控制項的內容。使用者可對已清除的內容執行復原作業。

#### <a name="syntax"></a>語法
```js
contentControlObject.clear();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
void

#### <a name="examples"></a>範例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the content controls collection.
    contentControls.load('text');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        if (contentControls.items.length === 0) {
            console.log("There isn't a content control in this document.");
        } else {
            
            // Queue a command to clear the contents of the first content control.
            contentControls.items[0].clear();
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync().then(function () {
                console.log('Content control cleared of contents.');
            });      
        }
            
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});

```

### <a name="delete(keepcontent:-bool)"></a>delete(keepContent: bool)
刪除內容控制項和其內容。如果 keepContent 設定為 true，則不能刪除內容。

#### <a name="syntax"></a>語法
```js
contentControlObject.delete(keepContent);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|keepContent|bool|必要。指出是否應隨著內容控制項一併刪除內容。如果 keepContent 設定為 true，則不能刪除內容。|

#### <a name="returns"></a>傳回
void

#### <a name="examples"></a>範例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the content controls collection.
    contentControls.load('text');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        if (contentControls.items.length === 0) {
            console.log("There isn't a content control in this document.");
        } else {
            
            // Queue a command to delete the first content control. The
            // contents will remain in the document.
            contentControls.items[0].delete(true);
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync().then(function () {
                console.log('Content control cleared of contents.');
            });      
        }
            
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```


### <a name="gethtml()"></a>getHtml()
取得內容控制項物件的 HTML 表示法。

#### <a name="syntax"></a>語法
```js
contentControlObject.getHtml();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>會傳回
字串

#### <a name="examples"></a>範例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection that contains a specific tag.
    var contentControlsWithTag = context.document.contentControls.getByTag('Customer-Address');
    
    // Queue a command to load the tag property for all of content controls. 
    context.load(contentControlsWithTag, 'tag');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControlsWithTag.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to get the HTML contents of the first content control.
            var html = contentControlsWithTag.items[0].getHtml();
        
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Content control HTML: ' + html.value);
            });
        }
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="getooxml()"></a>getOoxml()
取得內容控制項物件的 Office Open XML (OOXML) 表示法。

#### <a name="syntax"></a>語法
```js
contentControlObject.getOoxml();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>會傳回
字串

#### <a name="examples"></a>範例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the id property for all of the content controls. 
    context.load(contentControls, 'id');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControls.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to get the OOXML contents of the first content control.
            var ooxml = contentControls.items[0].getOoxml();
        
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Content control OOXML: ' + ooxml.value);
            });
        }
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="insertbreak(breaktype:-breaktype,-insertlocation:-insertlocation)"></a>insertBreak(breakType: BreakType, insertLocation: InsertLocation)
在指定的位置插入中斷符號。除了換行符號可以插入至任何 body 物件，其他中斷符號只能插入到主文件內文所包含的物件中。InsertLocation 值可以是 'Before'、'After'、'Start' 或 'End'。

#### <a name="syntax"></a>語法
```js
contentControlObject.insertBreak(breakType, insertLocation);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|breakType|BreakType|必要。中斷類型 (breakType.md)|
|insertLocation|InsertLocation|必要。此值可以是 'Before'、'After'、'Start' 或 'End'。|

#### <a name="returns"></a>傳回
void

#### <a name="additional-details"></a>其他詳細資料
除了換行符號以外，您不能在頁首、頁尾、註腳、章節附註、註解和文字方塊所包含的物件中插入其他中斷符號。  

#### <a name="examples"></a>範例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a commmand to load the id property for all of content controls. 
    context.load(contentControls, 'id');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion. We now will have 
    // access to the content control collection.
    return context.sync().then(function () {
        if (contentControls.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to insert a page break after the first content control. 
            contentControls.items[0].insertBreak('page', "After");
            
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion. 
            return context.sync()
                .then(function () {
                    console.log('Inserted a page break after the first content control.');    
            });
        }
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="insertfilefrombase64(base64file:-string,-insertlocation:-insertlocation)"></a>insertFileFromBase64(base64File: string, insertLocation:InsertLocation)
在目前內容控制項的指定位置插入文件。InsertLocation 值可以是 'Replace'、'Start' 或 'End'。

#### <a name="syntax"></a>語法
```js
contentControlObject.insertFileFromBase64(base64File, insertLocation);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|base64File|string|必要。要插入的檔案 base64 編碼內容。|
|insertLocation|InsertLocation|必要。此值可以是 'Replace'、'Start' 或 'End'。|

#### <a name="returns"></a>傳回
[Range](range.md)

### <a name="inserthtml(html:-string,-insertlocation:-insertlocation)"></a>insertHtml(html: string, insertLocation:InsertLocation)
在內容控制項的指定位置插入 HTML。InsertLocation 值可以是 'Replace'、'Start' 或 'End'。

#### <a name="syntax"></a>語法
```js
contentControlObject.insertHtml(html, insertLocation);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|HTML|string|必要。要插入至內容控制項的 HTML。|
|insertLocation|InsertLocation|必要。此值可以是 'Replace'、'Start' 或 'End'。|

#### <a name="returns"></a>傳回
[Range](range.md)

#### <a name="examples"></a>範例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the id property for all of the content controls. 
    context.load(contentControls, 'id');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControls.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to put HTML into the contents of the first content control.
            contentControls.items[0].insertHtml('<strong>HTML content inserted into the content control.</strong>', 'Start');
        
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Inserted HTML in the first content control.');
            });
        }
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="insertinlinepicturefrombase64(base64encodedimage:-string,-insertlocation:-insertlocation)"></a>insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation:InsertLocation)
在內容控制項的指定位置插入內嵌圖片。InsertLocation 值可以是 'Replace'、'Start' 或 'End'。

#### <a name="syntax"></a>語法
contentControlObject.insertInlinePictureFromBase64(image, insertLocation);

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|base64EncodedImage|string|必要。要插入至內容控制項的 base64 編碼影像。|
|insertLocation|InsertLocation|必要。此值可以是 'Replace'、'Start' 或 'End'。|

#### <a name="returns"></a>傳回
[InlinePicture](inlinepicture.md)



### <a name="insertooxml(ooxml:-string,-insertlocation:-insertlocation)"></a>insertOoxml(ooxml: string, insertLocation:InsertLocation)
在內容控制項的指定位置插入 OOXML 或 wordProcessingML。InsertLocation 值可以是 'Replace'、'Start' 或 'End'。

#### <a name="syntax"></a>語法
```js
contentControlObject.insertOoxml(ooxml, insertLocation);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|ooxml|string|必要。要插入至內容控制項的 OOXML 或 wordProcessingML。|
|insertLocation|InsertLocation|必要。此值可以是 'Replace'、'Start' 或 'End'。|

#### <a name="returns"></a>傳回
[Range](range.md)

#### <a name="examples"></a>範例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the id property for all of the content controls. 
    context.load(contentControls, 'id');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControls.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to put OOXML into the contents of the first content control.
            contentControls.items[0].insertOoxml("<pkg:package xmlns:pkg='http://schemas.microsoft.com/office/2006/xmlPackage'><pkg:part pkg:name='/_rels/.rels' pkg:contentType='application/vnd.openxmlformats-package.relationships+xml' pkg:padding='512'><pkg:xmlData><Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name='/word/document.xml' pkg:contentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'><pkg:xmlData><w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' ><w:body><w:p><w:pPr><w:spacing w:before='360' w:after='0' w:line='480' w:lineRule='auto'/><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr></w:pPr><w:r><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr><w:t>This text has formatting directly applied to achieve its font size, color, line spacing, and paragraph spacing.</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>", "End");
        
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Inserted OOXML in the first content control.');
            });
        }
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

### <a name="insertparagraph(paragraphtext:-string,-insertlocation:-insertlocation)"></a>insertParagraph(paragraphText: string, insertLocation:InsertLocation)
在指定的位置插入段落。InsertLocation 值可以是 'Before'、'After'、'Start' 或 'End'。

#### <a name="syntax"></a>語法
```js
contentControlObject.insertParagraph(paragraphText, insertLocation);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|paragraphText|string|必要。要插入的段落文字。|
|insertLocation|InsertLocation|必要。此值可以是 'Before'、'After'、'Start' 或 'End'。|

#### <a name="returns"></a>傳回
[Paragraph](paragraph.md)

#### <a name="examples"></a>範例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the id property for all of the content controls. 
    context.load(contentControls, 'id');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControls.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to insert a paragraph after the first content control. 
            contentControls.items[0].insertParagraph('Text of the inserted paragraph.', 'After');
        
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Inserted a paragraph after the first content control.');
            });
        }
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="inserttext(text:-string,-insertlocation:-insertlocation)"></a>insertText(text: string, insertLocation:InsertLocation)
在內容控制項的指定位置插入文字。InsertLocation 值可以是 'Replace'、'Start' 或 'End'。

#### <a name="syntax"></a>語法
```js
contentControlObject.insertText(text, insertLocation);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|文字|string|必要。要插入至內容控制項的文字。|
|insertLocation|InsertLocation|必要。此值可以是 'Replace'、'Start' 或 'End'。|

#### <a name="returns"></a>傳回
[Range](range.md)

#### <a name="examples"></a>範例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the id property for all of the content controls. 
    context.load(contentControls, 'id');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControls.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to replace text in the first content control. 
            contentControls.items[0].insertText('Replaced text in the first content control.', 'Replace');
        
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Replaced text in the first content control.');
            });
        }
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

[Silly stories](https://aka.ms/sillystorywordaddin) 增益集範例示範如何使用 **insertText** 方法。

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

#### <a name="examples"></a>範例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy range object for the current selection.
    var range = context.document.getSelection();
    
    // Queue a commmand to create the content control.
    var myContentControl = range.insertContentControl();
    myContentControl.tag = 'Customer-Address';
    myContentControl.title = ' has t';
    myContentControl.style = 'Heading 2';
    myContentControl.insertText('One Microsoft Way, Redmond, WA 98052', 'replace');
    myContentControl.cannotEdit = true;
    
    // Queue a command to load the id property for the content control you created.
    context.load(myContentControl, 'id');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Created content control with id: ' + myContentControl.id);
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="search(searchtext:-string,-searchoptions:-paramtypestrings.searchoptions)"></a>search(searchText: string, searchOptions:ParamTypeStrings.SearchOptions)
以指定的 searchOptions 在內容控制項物件的範圍中執行搜尋。搜尋結果將是 range 物件的集合。

#### <a name="syntax"></a>語法
```js
contentControlObject.search(searchText, searchOptions);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|searchText|string|必要。搜尋文字。|
|[searchOptions](searchoptions.md)|ParamTypeStrings.SearchOptions|選用。搜尋選項。|

#### <a name="returns"></a>傳回
[SearchResultCollection](searchresultcollection.md)

### <a name="select(selectionmode:-selectionmode)"></a>select(selectionMode: SelectionMode)
選取內容控制項。這會讓 Word 捲動至該選取範圍。選取模式可以是 'Select'、'Start' 或 'End'。

#### <a name="syntax"></a>語法
```js
contentControlObject.select(selectionMode);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|selectionMode|SelectionMode|選用。選取模式可以是 'Select'、'Start' 或 'End'。'Select' 為預設值。|

#### <a name="returns"></a>傳回
void

#### <a name="examples"></a>範例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the id property for all of the content controls. 
    context.load(contentControls, 'id');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControls.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to select the first content control.
            contentControls.items[0].select();
        
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Selected the first content control.');
            });
        }
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

## <a name="property-access-examples"></a>屬性存取範例

### <a name="load-all-of-the-content-control-properties"></a>載入所有內容控制項屬性
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the id property for all of the content controls. 
    context.load(contentControls, 'id');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControls.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to load the properties on the first content control. 
            contentControls.items[0].load(  'appearance,' +
                                            'cannotDelete,' +
                                            'cannotEdit,' +
                                            'id,' +
                                            'placeHolderText,' +
                                            'removeWhenEdited,' +
                                            'title,' +
                                            'text,' +
                                            'type,' +
                                            'style,' +
                                            'tag,' +
                                            'font/size,' +
                                            'font/name,' +
                                            'font/color');             
        
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Property values of the first content control:' + 
                        '   ----- appearance: ' + contentControls.items[0].appearance + 
                        '   ----- cannotDelete: ' + contentControls.items[0].cannotDelete +
                        '   ----- cannotEdit: ' + contentControls.items[0].cannotEdit +
                        '   ----- color: ' + contentControls.items[0].color +
                        '   ----- id: ' + contentControls.items[0].id +
                        '   ----- placeHolderText: ' + contentControls.items[0].placeholderText +
                        '   ----- removeWhenEdited: ' + contentControls.items[0].removeWhenEdited +
                        '   ----- title: ' + contentControls.items[0].title +
                        '   ----- text: ' + contentControls.items[0].text +
                        '   ----- type: ' + contentControls.items[0].type +
                        '   ----- style: ' + contentControls.items[0].style +
                        '   ----- tag: ' + contentControls.items[0].tag +
                        '   ----- font size: ' + contentControls.items[0].font.size +
                        '   ----- font name: ' + contentControls.items[0].font.name +
                        '   ----- font color: ' + contentControls.items[0].font.color);
            });
        }
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
