# <a name="tablecell-object-(javascript-api-for-onenote)"></a>TableCell 物件 (適用於 OneNote 的 JavaScript API)

_適用於：OneNote Online_  


代表 OneNote 表格中的儲存格。

## <a name="properties"></a>屬性

| 屬性	     | 類型	   |描述|意見反應|
|:---------------|:--------|:----------|:-------|
|cellIndex|int|取得列中的儲存格索引。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-cellIndex)|
|id|字串|取得儲存格的識別碼。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-id)|
|rowIndex|int|取得表格中儲存格之列的索引。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-rowIndex)|
|shadingColor|字串|取得或設定儲存格的網底色彩。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-shadingColor)|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
| 關聯性 | 類型	   |描述| 意見反應|
|:---------------|:--------|:----------|:-------|
|paragraphs|[ParagraphCollection](paragraphcollection.md)|取得 TableCell 中 Paragraph 物件的集合。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-paragraphs)|
|parentRow|[TableRow](tablerow.md)|取得儲存格的父列。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-parentRow)|

## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述| 意見反應|
|:---------------|:--------|:----------|:-------|
|[appendHtml(html: string)](#appendhtmlhtml-string)|void|將指定的 HTML 加入 TableCell 的底部。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-appendHtml)|
|[appendImage(base64EncodedImage: string, width: double, height: double)](#appendimagebase64encodedimage-string-width-double-height-double)|[Image](image.md)|將指定的影像新增至表格儲存格。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-appendImage)|
|[appendRichText(paragraphText: string)](#appendrichtextparagraphtext-string)|[RichText](richtext.md)|將指定的文字新增至表格儲存格。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-appendRichText)|
|[appendTable(rowCount: number, columnCount: number, values: string[][])](#appendtablerowcount-number-columncount-number-values-string)|[Table](table.md)|將具有指定列和欄數的表格新增至表格儲存格。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-appendTable)|
|[clear()](#clear)|void|清除儲存格的內容。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-clear)|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-load)|

## <a name="method-details"></a>方法詳細資料


### <a name="appendhtml(html:-string)"></a>appendHtml(html: string)
將指定的 HTML 加入 TableCell 的底部。

#### <a name="syntax"></a>語法
```js
tableCellObject.appendHtml(html);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|HTML|字串|要附加的 HTML 字串。請參閱 OneNote 增益集 JavaScript API 的[支援的 HTML](../../docs/onenote/onenote-add-ins-page-content.md#supported-html)。|

#### <a name="returns"></a>傳回
void

#### <a name="examples"></a>範例
```js
OneNote.run(function(ctx) {
    var app = ctx.application;
    var outline = app.getActiveOutline();
    
    // Queue a command to load outline.paragraphs and their types.
    ctx.load(outline, "paragraphs, paragraphs/type");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return ctx.sync().then(function () {
        var paragraphs = outline.paragraphs;
        
        // for each table, get a table cell at row one and column two and add "Hello".
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                var cell = table.getCell(1 /*Row Index*/, 2 /*Column Index*/);
                cell.appendHtml("<p>Hello</p>");
            }
        }
        return ctx.sync();
    })
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});


### appendImage(base64EncodedImage: string, width: double, height: double)
Adds the specified image to table cell.

#### Syntax
```js
tableCellObject.appendImage(base64EncodedImage, width, height);
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
將指定的文字新增至表格儲存格。

#### <a name="syntax"></a>語法
```js
tableCellObject.appendRichText(paragraphText);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|paragraphText|string|要附加的 HTML 字串。|

#### <a name="returns"></a>傳回
[RichText](richtext.md)

#### <a name="examples"></a>範例
```js
OneNote.run(function(ctx) {
    var app = ctx.application;
    var outline = app.getActiveOutline();
    var appendedRichText = null;
    
    // Queue a command to load outline.paragraphs and their types.
    ctx.load(outline, "paragraphs, paragraphs/type");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return ctx.sync().then(function () {
        var paragraphs = outline.paragraphs;
        
        // for each table, get a table cell at row one and column two and add "Hello".
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                var cell = table.getCell(1 /*Row Index*/, 2 /*Column Index*/);
                appendedRichText = cell.appendRichText("Hello");
            }
        }
        return ctx.sync();
    })
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="appendtable(rowcount:-number,-columncount:-number,-values:-string[][])"></a>appendTable(rowCount: number, columnCount: number, values: string[][])
將具有指定列和欄數的表格新增至表格儲存格。

#### <a name="syntax"></a>語法
```js
tableCellObject.appendTable(rowCount, columnCount, values);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|rowCount|number|必要。表格中的列數。|
|columnCount|number|必要。表格中的欄數。|
|values|string[][]|選用。選擇性的 2 維陣列。如果陣列中指定對應的字串，則會填滿儲存格。|

#### <a name="returns"></a>傳回
[Table](table.md)

### <a name="clear()"></a>clear()
清除儲存格的內容。

#### <a name="syntax"></a>語法
```js
tableCellObject.clear();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
void

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
**id、cellIndex、rowIndex**
```js
OneNote.run(function(ctx) {
    var app = ctx.application;
    var outline = app.getActiveOutline();
    
    // Queue a command to load outline.paragraphs and their types.
    ctx.load(outline, "paragraphs, paragraphs/type");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return ctx.sync().then(function () {
        var paragraphs = outline.paragraphs;
        
        // for each table, get a table cell at row one and column two.
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                var cell = table.getCell(1 /*Row Index*/, 2 /*Column Index*/);
                
                // Queue a command to load the table cell.
                ctx.load(cell);
                ctx.sync().then(function() {
                    console.log("Cell Id: " + cell.id);
                    console.log("Cell Index: " + cell.cellIndex);
                    console.log("Cell's Row Index: " + cell.rowIndex);
                });
            }
        }
        return ctx.sync();
    })
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

**parentTable、cells**
```js
ParentTable, ParentRow, Paragraphs
OneNote.run(function(ctx) {
    var app = ctx.application;
    var outline = app.getActiveOutline();
    
    // Queue a command to load outline.paragraphs and their types.
    ctx.load(outline, "paragraphs, paragraphs/type");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return ctx.sync().then(function () {
        var paragraphs = outline.paragraphs;
        
        // for each table, get a table cell at row one and column two.
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                var cell = table.getCell(1 /*Row Index*/, 2 /*Column Index*/);
                
                // Queue a command to load parentTable, parentRow and paragraphs of the table cell.
                ctx.load(cell, "parentTable, parentRow, paragraphs");
                
                ctx.sync().then(function() {
                    console.log("Parent Table Id: " + cell.parentTable.id);
                    console.log("Parent Row Id: " + cell.parentRow.id);
                    var paragraphs = cell.paragraphs;
                    
                    for (var i = 0; i < paragraphs.items.length; i++) {
                        console.log("Paragraph Id: " + paragraphs.items[i].id);
                    }
                });
            }
        }
        return ctx.sync();
    })
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```
