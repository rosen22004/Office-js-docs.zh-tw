# <a name="section-object-(javascript-api-for-word)"></a>Section 物件 (適用於 Word 的 JavaScript API)

代表 Word 文件中的區段。

_適用於：Word 2016、Word for iPad、Word for Mac、Word Online_

## <a name="properties"></a>屬性
無

## <a name="relationships"></a>關聯性
| 關聯性 | 類型	   |描述|
|:---------------|:--------|:----------|
|body|[Body](body.md)|取得區段的內文。這不包括 headerfooter 和其他區段中繼資料。唯讀。|

## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述|
|:---------------|:--------|:----------|
|[getFooter(type: HeaderFooterType)](#getfootertype-headerfootertype)|[Body](body.md)|取得區段的其中一個頁尾。|
|[getHeader(type: HeaderFooterType)](#getheadertype-headerfootertype)|[Body](body.md)|取得區段的其中一個頁首。|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|

## <a name="method-details"></a>方法詳細資料

### <a name="getfooter(type:-headerfootertype)"></a>getFooter(type: HeaderFooterType)
取得區段的其中一個頁尾。

#### <a name="syntax"></a>語法
```js
sectionObject.getFooter(type);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|類型|HeaderFooterType|必要。要傳回的頁尾類型。此值可以是：'primary'、'firstPage' 或 'evenPages'。|

#### <a name="returns"></a>傳回
[Body](body.md)

#### <a name="examples"></a>範例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy sectionsCollection object.
    var mySections = context.document.sections;
    
    // Queue a commmand to load the sections.
    context.load(mySections, 'body/style');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Create a proxy object the primary footer of the first section. 
        // Note that the footer is a body object.
        var myFooter = mySections.items[0].getFooter("primary");
        
        // Queue a command to insert text at the end of the footer.
        myFooter.insertText("This is a footer.", Word.InsertLocation.end);
        
        // Queue a command to wrap the header in a content control.
        myFooter.insertContentControl();
                              
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log("Added a footer to the first section.");
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
### <a name="getheader(type:-headerfootertype)"></a>getHeader(type: HeaderFooterType)
取得區段的其中一個頁首。

#### <a name="syntax"></a>語法
```js
sectionObject.getHeader(type);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|類型|HeaderFooterType|必要。要傳回的頁首類型。此值可以是：'primary'、'firstPage' 或 'evenPages'。|

#### <a name="returns"></a>傳回
[Body](body.md)

#### <a name="examples"></a>範例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy sectionsCollection object.
    var mySections = context.document.sections;
    
    // Queue a commmand to load the sections.
    context.load(mySections, 'body/style');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Create a proxy object the primary header of the first section. 
        // Note that the header is a body object.
        var myHeader = mySections.items[0].getHeader("primary");
        
        // Queue a command to insert text at the end of the header.
        myHeader.insertText("This is a header.", Word.InsertLocation.end);
        
        // Queue a command to wrap the header in a content control.
        myHeader.insertContentControl();
                              
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log("Added a header to the first section.");
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

## <a name="support-details"></a>支援詳細資料
在執行階段檢查使用[需求集](../office-add-in-requirement-sets.md)以確認您的應用程式受到 Word 主應用程式版本的支援。如需有關 Office 主應用程式及伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。