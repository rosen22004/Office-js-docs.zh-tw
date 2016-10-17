# <a name="section-object-(javascript-api-for-onenote)"></a>Section 物件 (適用於 OneNote 的 JavaScript API)

_適用於：OneNote Online_   


代表 OneNote 的節。節可以包含頁面。

## <a name="properties"></a>屬性

| 屬性	     | 類型	   |描述|意見反應|
|:---------------|:--------|:----------|:-------|
|clientUrl|字串|區段的用戶端 URL。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-section-clientUrl)|
|id|字串|取得節的識別碼。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-section-id)|
|name|字串|取得節的名稱。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-section-name)|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
| 關聯性 | 類型	   |描述| 意見反應|
|:---------------|:--------|:----------|:-------|
|notebook|[Notebook](notebook.md)|取得包含節的筆記本。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-section-notebook)|
|pages|[PageCollection](pagecollection.md)|節中的頁面集合。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-section-pages)|
|parentSectionGroup|[SectionGroup](sectiongroup.md)|取得包含節的節群組。如果節是筆記本的直接子項，則擲回 ItemNotFound。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-section-parentSectionGroup)|
|parentSectionGroupOrNull|[SectionGroup](sectiongroup.md)|取得包含節的節群組。如果節是筆記本的直接子項，則傳回 null。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-section-parentSectionGroupOrNull)|

## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述| 意見反應|
|:---------------|:--------|:----------|:-------|
|[addPage(title: string)](#addpagetitle-string)|[Page](page.md)|在節的結尾加入新頁面。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-section-addPage)|
|[copyToNotebook(destinationNotebook:Notebook)](#copytonotebookdestinationnotebook-notebook)|[Section](section.md)|將此區段複製到指定筆記本。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-section-copyToNotebook)|
|[copyToSectionGroup(destinationSectionGroup: SectionGroup)](#copytosectiongroupdestinationsectiongroup-sectiongroup)|[Section](section.md)|將此區段複製到指定區段群組。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-section-copyToSectionGroup)|
|[insertSectionAsSibling(location: string, title: string)](#insertsectionassiblinglocation-string-title-string)|[Section](section.md)|在目前這一節的前或後，插入新節。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-section-insertSectionAsSibling)|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-section-load)|

## <a name="method-details"></a>方法詳細資料


### <a name="addpage(title:-string)"></a>addPage(title: string)
在節的結尾加入新頁面。

#### <a name="syntax"></a>語法
```js
sectionObject.addPage(title);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|標題|string|新頁面的標題。|

#### <a name="returns"></a>傳回
[Page](page.md)

#### <a name="examples"></a>範例
```js
OneNote.run(function (context) {
            
    // Queue a command to add a page to the current section.
    var page = context.application.getActiveSection().addPage("Wish list");
            
    // Queue a command to load the id and title of the new page. 
    // This example loads the new page so it can read its properties later.           
    page.load('id,title');
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
             
            // Display the properties.       
            console.log("Page name: " + page.title);
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


### <a name="copytonotebook(destinationnotebook:-notebook)"></a>copyToNotebook(destinationNotebook:Notebook)
將此區段複製到指定筆記本。

#### <a name="syntax"></a>語法
```js
sectionObject.copyToNotebook(destinationNotebook);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|destinationNotebook|筆記本|要複製此區段的筆記本。|

#### <a name="returns"></a>傳回
[Section](section.md)

#### <a name="examples"></a>範例
```js
OneNote.run(function (context) {
    var app = context.application;
    
    // Gets the active Notebook.
    var notebook = app.getActiveNotebook();
    
    // Gets the active Section.
    var section = app.getActiveSection();
    
    var newSection;
    
    return context.sync()
        .then(function() {
            newSection = section.copyToNotebook(notebook);
            newSection.load('id');
            return context.sync();
        })
        .then(function() {
            console.log(newSection.id);
        });
})
.catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### <a name="copytosectiongroup(destinationsectiongroup:-sectiongroup)"></a>copyToSectionGroup(destinationSectionGroup: SectionGroup)
將此區段複製到指定區段群組。

#### <a name="syntax"></a>語法
```js
sectionObject.copyToSectionGroup(destinationSectionGroup);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|destinationSectionGroup|SectionGroup|要複製此區段的區段群組。|

#### <a name="returns"></a>傳回
[Section](section.md)

#### <a name="examples"></a>範例
```js
OneNote.run(function (ctx) {
    var app = ctx.application;
    
    // Gets the active Notebook.
    var notebook = app.getActiveNotebook();
    
    // Gets the active Section.
    var section = app.getActiveSection();
    
    var newSection;
    
    return ctx.sync()
        .then(function() {
            var firstSectionGroup = notebook.sectionGroups.items[0];
            newSection = section.copyToSectionGroup(firstSectionGroup);
            newSection.load('id');
            return ctx.sync();
        })
        .then(function() {
            console.log(newSection.id);
        });
})
.catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### <a name="insertsectionassibling(location:-string,-title:-string)"></a>insertSectionAsSibling(location: string, title: string)
在目前這一節的前或後，插入新節。

#### <a name="syntax"></a>語法
```js
sectionObject.insertSectionAsSibling(location, title);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|Location|字串|新節與目前這一節的相對位置。可能的值為：之前、之後|
|標題|string|新區段的名稱。|

#### <a name="returns"></a>傳回
[Section](section.md)

#### <a name="examples"></a>範例
```js
OneNote.run(function (context) {
            
    // Queue a command to insert a section after the current section.
    var section = context.application.getActiveSection().insertSectionAsSibling("After", "New section");
            
    // Queue a command to load the id and name of the new section. 
    // This example loads the new section so it can read its properties later.           
    section.load('id,name');
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
             
            // Display the properties.       
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
OneNote.run(function (context) {
        
    // Get the current section.
    var section = context.application.getActiveSection();
            
    // Queue a command to load the section. 
    // For best performance, request specific properties.           
    section.load("id");
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
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

**name 和 notebook**
```js
OneNote.run(function (context) {
        
    // Get the current section.
    var section = context.application.getActiveSection();
            
    // Queue a command to load the section with the specified properties. 
    section.load("name,notebook/name");
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            console.log("Section name: " + section.name);
            console.log("Parent notebook name: " + section.notebook.name);
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

**parentSectionGroupOrNull**
```js
OneNote.run(function (context) {
    // Queue a command to add a page to the current section.
    var section = context.application.getActiveSection();
    section.load('clientUrl,notebook');
    var sectionGroup = section.parentSectionGroupOrNull;
    
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            if(sectionGroup.isNull === false)
            {
                // If a parent section group exists, queue a command to add a section in it!
                sectionGroup.addSection("NewSectionInSectionGroup");
            }
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
    
