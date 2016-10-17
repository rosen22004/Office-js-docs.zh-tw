# <a name="application-object-(javascript-api-for-onenote)"></a>應用程式物件 (適用於 OneNote 的 JavaScript API)

_適用於：OneNote Online_


代表最上層物件，會包含所有的全域可定址 OneNote 物件，例如筆記本、使用中的筆記本中和使用中的節。

## <a name="properties"></a>屬性

無

## <a name="relationships"></a>關聯性
| 關聯性 | 類型	   |描述| 意見反應|
|:---------------|:--------|:----------|:-------|
|Notebooks|[NotebookCollection](notebookcollection.md)|取得在 OneNote 應用程式執行個體中開啟的筆記本集合。在 OneNote Online 中，應用程式執行個體內，一次只能開啟一個筆記本。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-notebooks)|

## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述| 意見反應|
|:---------------|:--------|:----------|:-------|
|[getActiveNotebook()](#getactivenotebook)|[Notebook](notebook.md)|取得使用中的筆記本 (如果有的話)。如果沒有使用中的筆記本，會擲回 ItemNotFound。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-getActiveNotebook)|
|[getActiveNotebookOrNull()](#getactivenotebookornull)|[Notebook](notebook.md)|取得使用中的筆記本 (如果有的話)。如果沒有使用中的筆記本，會傳回 null。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-getActiveNotebookOrNull)|
|[getActiveOutline()](#getactiveoutline)|[Outline](outline.md)|取得使用中的大綱 (如果有的話)，如果沒有使用中的大綱，會擲回 ItemNotFound。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-getActiveOutline)|
|[getActiveOutlineOrNull()](#getactiveoutlineornull)|[Outline](outline.md)|取得使用中的大綱 (如果有的話)，否則會傳回 null。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-getActiveOutlineOrNull)|
|[getActivePage()](#getactivepage)|[Page](page.md)|取得使用中的頁面 (如果有的話)。如果沒有使用中的頁面，會擲回 ItemNotFound。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-getActivePage)|
|[getActivePageOrNull()](#getactivepageornull)|[Page](page.md)|取得使用中的頁面 (如果有的話)。如果沒有使用中的頁面，會傳回 null。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-getActivePageOrNull)|
|[getActiveSection()](#getactivesection)|[Section](section.md)|取得使用中的區段 (如果有的話)。如果沒有使用中的區段，會擲回 ItemNotFound。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-getActiveSection)|
|[getActiveSectionOrNull()](#getactivesectionornull)|[Section](section.md)|取得使用中的區段 (如果有的話)。如果沒有使用中的區段，會傳回 null。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-getActiveSectionOrNull)|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-load)|
|[navigateToPage(page:Page)](#navigatetopagepage-page)|void|在應用程式執行個體中開啟指定的頁面。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-navigateToPage)|
|[navigateToPageWithClientUrl(url: string)](#navigatetopagewithclienturlurl-string)|[Page](page.md)|取得指定的頁面，並且在應用程式執行個體中開啟。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-navigateToPageWithClientUrl)|

## <a name="method-details"></a>方法詳細資料


### <a name="getactivenotebook()"></a>getActiveNotebook()
取得使用中的筆記本 (如果有的話)。如果沒有使用中的筆記本，會擲回 ItemNotFound。

#### <a name="syntax"></a>語法
```js
applicationObject.getActiveNotebook();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
[Notebook](notebook.md)

#### <a name="examples"></a>範例
```js
OneNote.run(function (context) {
        
    // Get the active notebook.
    var notebook = context.application.getActiveNotebook();
            
    // Queue a command to load the notebook. 
    // For best performance, request specific properties.           
    notebook.load('id,name');
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
                    
            // Show some properties.
            console.log("Notebook name: " + notebook.name);
            console.log("Notebook ID: " + notebook.id);
            
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### <a name="getactivenotebookornull()"></a>getActiveNotebookOrNull()
取得使用中的筆記本 (如果有的話)。如果沒有使用中的筆記本，會傳回 null。

#### <a name="syntax"></a>語法
```js
applicationObject.getActiveNotebookOrNull();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
[Notebook](notebook.md)

#### <a name="examples"></a>範例
```js
OneNote.run(function (context) {

    // Get the active notebook.
    var notebook = context.application.getActiveNotebookOrNull();

    // Queue a command to load the notebook. 
    // For best performance, request specific properties.           
    notebook.load('id,name');

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // check if active notebook is set.
            if (!notebook.isNull) {
                console.log("Notebook name: " + notebook.name);
                console.log("Notebook ID: " + notebook.id);
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


### <a name="getactiveoutline()"></a>getActiveOutline()
取得使用中的大綱 (如果有的話)，如果沒有使用中的大綱，會擲回 ItemNotFound。

#### <a name="syntax"></a>語法
```js
applicationObject.getActiveOutline();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
[Outline](outline.md)

#### <a name="examples"></a>範例
```js
OneNote.run(function (context) {

    // get active outline.
    var outline = context.application.getActiveOutline();

    // Queue a command to load the id of the outline.         
    outline.load('id');

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // Show some properties.
            console.log("outline id: " + outline.id);
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### <a name="getactiveoutlineornull()"></a>getActiveOutlineOrNull()
取得使用中的大綱 (如果有的話)，否則會傳回 null。

#### <a name="syntax"></a>語法
```js
applicationObject.getActiveOutlineOrNull();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
[Outline](outline.md)

#### <a name="examples"></a>範例
```js
OneNote.run(function (context) {

    // get active outline.
    var outline = context.application.getActiveOutlineOrNull();

    // Queue a command to load the id of the outline.         
    outline.load('id');

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            if (!outline.isNull) {
                console.log("outline id: " + outline.id);
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


### <a name="getactivepage()"></a>getActivePage()
取得使用中的頁面 (如果有的話)。如果沒有使用中的頁面，會擲回 ItemNotFound。

#### <a name="syntax"></a>語法
```js
applicationObject.getActivePage();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
[Page](page.md)

#### <a name="examples"></a>範例
```js
OneNote.run(function (context) {
        
    // Get the active page.
    var page = context.application.getActivePage();
            
    // Queue a command to load the page. 
    // For best performance, request specific properties.           
    page.load('id,title');
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
                    
            // Show some properties.
            console.log("Page title: " + page.title);
            console.log("Page ID: " + page.id);
            
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### <a name="getactivepageornull()"></a>getActivePageOrNull()
取得使用中的頁面 (如果有的話)。如果沒有使用中的頁面，會傳回 null。

#### <a name="syntax"></a>語法
```js
applicationObject.getActivePageOrNull();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
[Page](page.md)

#### <a name="examples"></a>範例
```js
OneNote.run(function (context) {

    // Get the active page.
    var page = context.application.getActivePageOrNull();

    // Queue a command to load the page. 
    // For best performance, request specific properties.           
    page.load('id,title');

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            
            if (!page.isNull) {
                // Show some properties.
                console.log("Page title: " + page.title);
                console.log("Page ID: " + page.id);
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


### <a name="getactivesection()"></a>getActiveSection()
取得使用中的區段 (如果有的話)。如果沒有使用中的區段，會擲回 ItemNotFound。

#### <a name="syntax"></a>語法
```js
applicationObject.getActiveSection();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
[Section](section.md)

#### <a name="examples"></a>範例
```js
OneNote.run(function (context) {
        
    // Get the active section.
    var section = context.application.getActiveSection();
            
    // Queue a command to load the section. 
    // For best performance, request specific properties.           
    section.load('id,name');
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
                    
            // Show some properties.
            console.log("Section name: " + section.name);
            console.log("Section ID: " + section.id);
            
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### <a name="getactivesectionornull()"></a>getActiveSectionOrNull()
取得使用中的區段 (如果有的話)。如果沒有使用中的區段，會傳回 null。

#### <a name="syntax"></a>語法
```js
applicationObject.getActiveSectionOrNull();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
[Section](section.md)

#### <a name="examples"></a>範例
```js
OneNote.run(function (context) {

    // Get the active section.
    var section = context.application.getActiveSectionOrNull();

    // Queue a command to load the section. 
    // For best performance, request specific properties.           
    section.load('id,name');

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            if (!section.isNull) {
                // Show some properties.
                console.log("Section name: " + section.name);
                console.log("Section ID: " + section.id);
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

### <a name="navigatetopage(page:-page)"></a>navigateToPage(page:Page)
在應用程式執行個體中開啟指定的頁面。

#### <a name="syntax"></a>語法
```js
applicationObject.navigateToPage(page);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|page|Page|要開啟的頁面。|

#### <a name="returns"></a>傳回
void

#### <a name="examples"></a>範例
```js        
OneNote.run(function (context) {
        
    // Get the pages in the current section.
    var pages = context.application.getActiveSection().pages;
            
    // Queue a command to load the pages. 
    // For best performance, request specific properties.           
    pages.load('id');
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
                    
            // This example loads the first page in the section.
            var page = pages.items[0];
                        
            // Open the page in the application.                    
            context.application.navigateToPage(page);
                    
            // Run the queued command.
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


### <a name="navigatetopagewithclienturl(url:-string)"></a>navigateToPageWithClientUrl(url: string)
取得指定的頁面，並且在應用程式執行個體中開啟。

#### <a name="syntax"></a>語法
```js
applicationObject.navigateToPageWithClientUrl(url);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|URL|字串|要開啟頁面的用戶端 URL。|

#### <a name="returns"></a>傳回
[Page](page.md)

#### <a name="examples"></a>範例
```js
OneNote.run(function (context) {

    // Get the pages in the current section.
    var pages = context.application.getActiveSection().pages;

    // Queue a command to load the pages. 
    // For best performance, request specific properties.           
    pages.load('clientUrl');

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // This example loads the first page in the section.
            var page = pages.items[0];

            // Open the page in the application.                    
            context.application.navigateToPageWithClientUrl(page.clientUrl);

            // Run the queued command.
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
