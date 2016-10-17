# <a name="contentcontrolcollection-object-(javascript-api-for-word)"></a>ContentControlCollection 物件 (適用於 Word 的 JavaScript API)

包含 ContentControl 物件的集合。內容控制項是指文件中具有界限且可能具有標籤的區域，這些區域會做為特定內容類型的容器。個別的內容控制項可能含有內容，例如影像、表格或格式化文字的段落。目前僅支援 RTF 內容控制項。

_適用於：Word 2016、Word for iPad、Word for Mac、Word Online_

## <a name="properties"></a>屬性
| 屬性	     | 類型	   |描述
|:---------------|:--------|:----------|
|項目|[ContentControl[]](contentcontrol.md)|ContentControl 物件的集合。唯讀。|

## <a name="relationships"></a>關聯性
無


## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述|
|:---------------|:--------|:----------|
|[getById(id: number)](#getbyidid-number)|[ContentControl](contentcontrol.md)|依識別碼取得內容控制項。|
|[getByTag(tag: string)](#getbytagtag-string)|[ContentControlCollection](contentcontrolcollection.md)|取得具有指定之標記的內容控制項。|
|[getByTitle(title: string)](#getbytitletitle-string)|[ContentControlCollection](contentcontrolcollection.md)|取得具有指定之標題的內容控制項。|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|

## <a name="method-details"></a>方法詳細資料

### <a name="getbyid(id:-number)"></a>getById(id: number)
依識別碼取得內容控制項。

#### <a name="syntax"></a>語法
```js
contentControlCollectionObject.getById(id);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|id|number|必要。內容控制項識別碼。|

#### <a name="returns"></a>傳回
[ContentControl](contentcontrol.md)

#### <a name="examples"></a>範例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the content control that contains a specific id.
    var contentControl = context.document.contentControls.getById(30086310);

    // Queue a command to load the text property for a content control.
    context.load(contentControl, 'text');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('The content control with that Id has been found in this document.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="getbytag(tag:-string)"></a>getByTag(tag: string)
取得具有指定之標記的內容控制項。

#### <a name="syntax"></a>語法
```js
contentControlCollectionObject.getByTag(tag);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|Tag|string|必要。內容控制項的標記設定。|

#### <a name="returns"></a>傳回
[ContentControlCollection](contentcontrolcollection.md)

#### <a name="examples"></a>範例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the content controls collection that contains a specific tag.
    var contentControlsWithTag = context.document.contentControls.getByTag('Customer-Address');

    // Queue a command to load the text property for all of content controls with a specific tag.
    context.load(contentControlsWithTag, 'text');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControlsWithTag.items.length === 0) {
            console.log("There isn't a content control with a tag of Customer-Address in this document.");
        } else {
            console.log('The first content control with the tag of Customer-Address has this text: ' + contentControlsWithTag.items[0].text);
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
The [Word-Add-in-DocumentAssembly][contentControls.getByTag] 範例提供使用 getByTag 方法的另一個範例。


### <a name="getbytitle(title:-string)"></a>getByTitle(title: string)
取得具有指定之標題的內容控制項。

#### <a name="syntax"></a>語法
```js
contentControlCollectionObject.getByTitle(title);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|標題|string|必要。內容控制項的標題。|

#### <a name="returns"></a>傳回
[ContentControlCollection](contentcontrolcollection.md)

#### <a name="examples"></a>範例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the content controls collection that contains a specific title.
    var contentControlsWithTitle = context.document.contentControls.getByTitle('Enter Customer Address Here');

    // Queue a command to load the text property for all of content controls with a specific title.
    context.load(contentControlsWithTitle, 'text');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControlsWithTitle.items.length === 0) {
            console.log("There isn't a content control with a title of 'Enter Customer Address Here' in this document.");
        } else {
            console.log("The first content control with the title of 'Enter Customer Address Here' has this text: " + contentControlsWithTitle.items[0].text);
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
[Word-Add-in-DocumentAssembly][contentControls.getByTitle] 範例提供使用 getByTitle 方法的另一個範例。

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

[Silly stories](https://aka.ms/sillystorywordaddin) 增益集範例示範如何使用 **load** 方法，搭配 **tag** 和 **title** 屬性，載入內容控制項集合。

## <a name="support-details"></a>支援詳細資料
在執行階段檢查使用[需求集](../office-add-in-requirement-sets.md)以確認您的應用程式受到 Word 主應用程式版本的支援。如需有關 Office 主應用程式及伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


[contentControls.getByTag]: https://github.com/OfficeDev/Word-Add-in-DocumentAssembly/blob/master/WordAPIDocAssemblySampleWeb/App/Home/Home.js#L300 "依標籤取得"
[contentControls.getByTitle]: https://github.com/OfficeDev/Word-Add-in-DocumentAssembly/blob/master/WordAPIDocAssemblySampleWeb/App/Home/Home.js#L331 "依標題取得"

