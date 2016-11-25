# <a name="page-object-(javascript-api-for-onenote)"></a>Page 物件 (適用於 OneNote 的 JavaScript API)

_適用於：OneNote Online_   


代表 OneNote 頁面。

## <a name="properties"></a>屬性

| 屬性	     | 類型	   |描述|意見反應|
|:---------------|:--------|:----------|:-------|
|clientUrl|字串|頁面的用戶端 URL。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-clientUrl)|
|id|字串|取得頁面的識別碼。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-id)|
|pageLevel|int|取得或設定頁面的縮排層次。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-pageLevel)|
|title|string|取得或設定頁面的標題。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-title)|
|webUrl|字串|頁面的網頁 URL。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-webUrl)|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
| 關聯性 | 類型	   |描述| 意見反應|
|:---------------|:--------|:----------|:-------|
|contents|[PageContentCollection](pagecontentcollection.md)|頁面上 PageContent 物件的集合。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-contents)|
|inkAnalysisOrNull|[InkAnalysis](inkanalysis.md)|頁面上筆跡的文字解譯。如果沒有筆跡分析資訊，則傳回 null。唯讀。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-inkAnalysisOrNull)|
|parentSection|[Section](section.md)|取得包含頁面的節。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-parentSection)|

## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述| 意見反應|
|:---------------|:--------|:----------|:-------|
|[addOutline(left: double, top: double, html:String)](#addoutlineleft-double-top-double-html-string)|[Outline](outline.md)|將大綱加入至頁面中的指定位置。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-addOutline)|
|[copyToSection(destinationSection:Section)](#copytosectiondestinationsection-section)|[Page](page.md)|將此頁面複製到指定區段。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-copyToSection)|
|[insertPageAsSibling(location: string, title: string)](#insertpageassiblinglocation-string-title-string)|[Page](page.md)|在目前頁面的前或後，插入新頁面。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-insertPageAsSibling)|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-load)|

## <a name="method-details"></a>方法詳細資料


### <a name="addoutline(left:-double,-top:-double,-html:-string)"></a>addOutline(left: double, top: double, html:String)
將大綱加入至頁面中的指定位置。

#### <a name="syntax"></a>語法
```js
pageObject.addOutline(left, top, html);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|left|double|大綱左上方邊角中的左側位置。|
|top|double|大綱左上方邊角中的頂端位置。|
|HTML|String|HTML 字串，描述大綱的視覺化呈現。請參閱 OneNote 增益集 JavaScript API 的[支援的 HTML](../../docs/onenote/onenote-add-ins-page-content.md#supported-html)。|

#### <a name="returns"></a>傳回
[Outline](outline.md)

#### <a name="examples"></a>範例
```js
OneNote.run(function (context) {

    // Gets the active page.
    var page = context.application.getActivePage();

    // Queue a command to add an outline with given html. 
    var outline = page.addOutline(200, 200,
"<p>Images and a table below:</p> \
 <img src=\"data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAUAAAAFCAYAAACNbyblAAAAHElEQVQI12P4//8/w38GIAXDIBKE0DHxgljNBAAO9TXL0Y4OHwAAAABJRU5ErkJggg==\"> \
 <img src=\"http://imagenes.es.sftcdn.net/es/scrn/6653000/6653659/microsoft-onenote-2013-01-535x535.png\"> \
 <table> \
   <tr> \
     <td>Jill</td> \
     <td>Smith</td> \
     <td>50</td> \
   </tr> \
   <tr> \
     <td>Eve</td> \
     <td>Jackson</td> \
     <td>94</td> \
   </tr> \
 </table>"     
        );

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
});
```


### <a name="copytosection(destinationsection:-section)"></a>copyToSection(destinationSection:Section)
將此頁面複製到指定區段。

#### <a name="syntax"></a>語法
```js
pageObject.copyToSection(destinationSection);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|destinationSection|區段|要複製此頁面的區段。|

#### <a name="returns"></a>傳回
[Page](page.md)

#### <a name="examples"></a>範例
```js
OneNote.run(function(ctx) {
    var app = ctx.application;
    
    // Gets the active notebook.
    var notebook = app.getActiveNotebook();
    
    // Gets the active page.
    var page = app.getActivePage();
    
    // Queue a command to load sections under the notebook.
    notebook.load('sections');
    
    var newPage;
    
    // Run the queued commands, and return a promise to indicate task completion.
    return ctx.sync()
        .then(function() {
            var section = notebook.sections.items[0];
            
            // copy page to the section.
            newPage = page.copyToSection(section);
            newPage.load('id');
            return ctx.sync();
        })
        .then(function() {
            console.log(newPage.id);
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="insertpageassibling(location:-string,-title:-string)"></a>insertPageAsSibling(location: string, title: string)
在目前頁面的前或後，插入新頁面。

#### <a name="syntax"></a>語法
```js
pageObject.insertPageAsSibling(location, title);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|Location|字串|新頁面與目前頁面的相對位置。可能的值為：之前、之後|
|標題|string|新頁面的標題。|

#### <a name="returns"></a>傳回
[Page](page.md)

#### <a name="examples"></a>範例
```js
OneNote.run(function (context) {

    // Gets the active page.
    var activePage = context.application.getActivePage();

    // Queue a command to add a new page after the active page. 
    var newPage = activePage.insertPageAsSibling("After", "Next Page");

    // Queue a command to load the newPage to access its data.
    context.load(newPage);

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            console.log("page is created with title: " + newPage.title);
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
### <a name="property-access-examples"></a>屬性存取範例

**contents**
```js
OneNote.run(function (context) {

    // Gets the active page.
    var activePage = context.application.getActivePage();

    // Queue a command to add a new page after the active page. 
    var pageContents = activePage.contents;

    // Queue a command to load the pageContents to access its data.
    context.load(pageContents);

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            for(var i=0; i < pageContents.items.length; i++)
            {
                var pageContent = pageContents.items[i];
                if (pageContent.type == "Outline")
                {
                    console.log("Found an outline");
                }
                else if (pageContent.type == "Image")
                {
                    console.log("Found an image");
                }
                else if (pageContent.type == "Other")
                {
                    console.log("Found a type not supported yet.");
                }
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

**webUrl**
```js
OneNote.run(function (context) {

    var app = context.application;
    
    // Gets the active page.
    var page = app.getActivePage();
    
    // Queue a command to load the webUrl of the page.
    page.load("webUrl");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            console.log(page.webUrl);
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

**inkAnalysisOrNull**
```js
OneNote.run(function (ctx) {        
    var app = ctx.application;
    
    // Gets the active page.
    var page = app.getActivePage();
    
    // Load ink words
    page.load('inkAnalysisOrNull/paragraphs/lines/words');
    
    return ctx.sync()
        .then(function() {
            if (!page.inkAnalysisOrNull.isNull)
                console.log(page.inkAnalysisOrNull.paragraphs.length);
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```
