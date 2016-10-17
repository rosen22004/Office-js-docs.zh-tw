# <a name="word-javascript-api-reference"></a>Word JavaScript API 參考資料

Word 會提供一組豐富的 API 可供您用來建立增益集，而該增益集可與文件內容和中繼資料進行互動。使用這些 API 來建立可與 Word 整合並將其擴充的吸引人的體驗。您可以匯入和匯出內容、從不同的資料來源組合新文件，並與文件工作流程整合，來建立自訂的文件解決方案。

您可以使用兩個 JavaScript API 與 Word 文件中的物件和中繼資料進行互動︰

- Word JavaScript API - 在 Office 2016 中推出。
- [適用於 Office 的 JavaScript API](../javascript-api-for-office.md) (Office.js) - 在 Office 2013 中推出。

## <a name="word-javascript-api"></a>Word JavaScript API

Word JavaScript API 由 Office.js 載入。Word JavaScript API 改變了您與文件、段落等物件互動的方式。Word JavaScript API 不是提供個別非同步 API 來擷取和更新各物件，而是提供對應到 Word 中執行之實際物件的「代理」("proxy") JavaScript 物件。您可以與這些 proxy 物件互動，方法是同步讀取和寫入其屬性，以及呼叫同步方法以在其上執行作業。與 proxy 物件的互動不會在執行中指令碼內立即實現。此 **context.sync** 方法會透過執行佇列的指示以及擷取已載入供指令碼使用之 Word 物件的屬性，來同步處理執行中 JavaScript 和 Office 中實際物件之間的狀態。

## <a name="javascript-api-for-office"></a>JavaScript API for Office

您可以從下列位置參考 Office.js︰

* https://appsforoffice.microsoft.com/lib/1/hosted/office.js - 針對實際執行的增益集，請使用此資源。
* https://appsforoffice.microsoft.com/lib/beta/hosted/office.js - 當您想嘗試預覽功能時，請使用此資源。

如果您正在使用 [Visual Studio](https://www.visualstudio.com/products/free-developer-offers-vs)，您可以下載 [Office 開發人員工具](https://www.visualstudio.com/features/office-tools-vs.aspx)以取得包含 Office.js 的專案範本。您也可以使用[取得 Office.js 的 NuGet](https://www.nuget.org/packages/Microsoft.Office.js/)。

如果您使用 TypeScript 並有 npm，您可以在命令列介面中輸入下列命令以取得 TypeScript 定義︰```typings install office-js --ambient```。

## <a name="running-word-add-ins"></a>執行 Word 增益集

若要執行增益集，請使用 Office.initialize 事件控制碼。如需增益集初始化的詳細資訊，請參閱[了解 API](../../docs/develop/understanding-the-javascript-api-for-office.md)。

針對 Word 2016 的增益集，執行時會在 **Word.run()** 方法中傳遞函數。**Run** 方法中傳遞的函數必須具有 context 引數。此 [context 物件](../../reference/word/requestcontext.md)不同於您從 Office 物件取得的內容物件，但同樣都是用來與 Word 執行階段環境互動。下列範例顯示如何使用 **Word.run()** 方法來初始化並執行 Word 增益集。

```js
    (function () {
        "use strict";

        // The initialize event handler must be run on each page to initialize Office JS.
        // You can add optional custom initialization code that will run after OfficeJS
        // has initialized.
        Office.initialize = function (reason) {
            // The reason object tells how the add-in was initialized. The values can be:
            // inserted - the add-in was inserted to an open document.
            // documentOpened - the add-in was already inserted in to the document and the document was opened.

            // Checks for the DOM to load using the jQuery ready function.
            $(document).ready(function () {
                // Set your optional initialization code.
                // You can also load saved settings from the Office object.
            });
        };

        // Run a batch operation against the Word JavaScript API object model.
        // Use the context argument to get access to the Word document.
        Word.run(function (context) {

            // Create a proxy object for the document.
            var thisDocument = context.document;
            // ...
        })
    })();
```

### <a name="synchronizing-word-documents-with-word-javascript-api-proxy-objects"></a>使用 Word JavaScript API proxy 物件同步處理 Word 文件

Word JavaScript API 物件模型與 Word 中的物件鬆散結合。Word JavaScript API 物件是 Word 文件中物件的 proxy。除非同步處理文件狀態，否則對 proxy 物件採取的動作都不會在 Word 中實現。相反地，除非同步處理文件狀態，否則 Word 文件的狀態不會在 proxy 物件中實現。若要同步處理文件狀態，您需執行 **context.sync()** 方法。下列範例建立 proxy body 物件與佇列命令，以載入 proxy body 物件上的 text 屬性，並且使用 **context.sync()** 方法將 Word 文件的內文與 body proxy 物件同步處理。

```js
    // Run a batch operation against the Word object model.
    Word.run(function (context) {

        // Create a proxy object for the document body.
        // The body object hasn't been set with any property values.
        var body = context.document.body;

        // Queue a command to load the text property for the proxy document body object.
        context.load(body, 'text');

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log("Body contents: " + body.text);
        });
    })
```

### <a name="executing-a-batch-of-commands"></a>執行命令批次

Word proxy 物件具有可以存取和更新物件模型的方法。系統會依照批次中方法的佇列順序，循序執行這些方法。呼叫 context.sync() 時，便會執行批次中所有排入佇列的命令。

下列範例顯示命令佇列的運作方式。呼叫 **context.sync()** 時，會在 Word 中執行[命令以載入](../../reference/word/loadoption.md)內文文字。接著，執行將文字插入 Word 內文的命令。隨後將結果傳回 body proxy 物件。Word JavaScript API 中 **body.text** 屬性的值，會是文字插入 Word 文件<u>之前</u>的 Word 文件內文的值。


```js
    // Run a batch operation against the Word JavaScript API.
    Word.run(function (context) {

        // Create a proxy object for the document body.
        var body = context.document.body;

        // Queue a command to load the text property of the proxy body object.
        context.load(body, 'text');

        // Queue a command to insert text into the end of the Word document body.
        body.insertText('This is text inserted after loading the body.text property',
                        Word.InsertLocation.end);

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log("Body contents: " + body.text);
        });
    })
```

## <a name="open-word-api-specifications"></a>開放式 Word API 規格

我們設計和開發新的 Word 增益集 API 時，我們會將其放在[開放式 API 規格](../../reference/openspec.md)頁面中，可供您提出意見反應。了解 Word JavaScript API 即將推出的新功能，並對我們的設計規格提出意見反應。

## <a name="additional-resources"></a>其他資源

* [Word 增益集概觀](../../docs/word/word-add-ins-programming-overview.md )
* [Office 增益集平台概觀](../../docs/overview/office-add-ins.md)
* [GitHub 上的 Word 增益集範例](https://github.com/OfficeDev?utf8=%E2%9C%93&query=Word)
