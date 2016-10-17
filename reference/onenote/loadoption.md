# <a name="object-load-options"></a>物件載入選項 

代表一個可以傳遞至 load 方法的物件，以指定在執行 sync() 方法 (用以同步處理 OneNote 物件和增益集內相對應 JavaScript proxy 物件之間的狀態) 時要載入的屬性集和關聯。此物件需要 select 和 expand 參數等選項，以指定要載入至物件的屬性集，同時也允許在集合上分頁。

也可以提供包含要載入之屬性和關聯的字串，或提供包含要載入之屬性和關聯清單的陣列。請參閱下列的範例。

```js   
object.load('<var1>,<relationship1/var2>');

// Pass the parameter as an array.
object.load(["var1", "relationship1/var2"]);
```

## <a name="properties"></a>屬性
| 屬性	     | 類型	   |說明|
|:---------------|:--------|:----------|
|select|Object|提供在同步呼叫時要載入的參數/關聯性名稱的逗點分隔清單或陣列，例如 "property1, relationship1", [ "property1", "relationship1"]。選用。|
|expand|Object|提供在同步呼叫時要載入的關聯性名稱的逗點分隔清單或陣列，例如 "relationship1, relationship2", [ "relationship1", "relationship2"]。選用。|
|top|int|指定結果中所包含查詢集合內的項目數。選用。|
|skip|int|指定結果中要略過不予包含的集合項目數。如果指定 `top`，則結果的選取範圍會在略過指定的項目數後開始。選用。|

#### <a name="examples"></a>範例

在範例中，取得在目前節中前五頁的頁面標題和縮排層次。

```js
OneNote.run(function (context) { 
    
    // Get the pages in the current section.
    var pages = context.application.getActiveSection().pages;
            
    // Queue a command to load the pages.           
    pages.load({ "select":"title,pageLevel", "top":5, "skip":0 });
    return context.sync()
        .then(function() {
            
            // Iterate through the collection of pages.    
            $.each(pages.items, function(index, page) {
                
                // Show some properties.
                console.log("Page title: " + page.title);
                console.log("Indentation level: " + page.pageLevel);
                
            });
        }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        })
    });
```
