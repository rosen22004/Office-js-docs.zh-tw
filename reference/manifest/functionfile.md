# <a name="functionfile-element"></a>FunctionFile 元素

透過增益集命令為增益集所公開的作業指定原始程式碼檔，以執行 JavaScript 函式而非顯示 UI。**FunctionFile** 元素是 [DesktopFormFactor](./desktopformfactor.md) 或 [MobileFormFactor](./mobileformfactor.md) 的子元素。**FunctionFile** 元素的 **resid** 屬性是設定為 **Resources** 元素中 **Url** 元素的 **id** 屬性，包含要包含或載入無 UI 增益集命令按鈕所使用的所有 JavaScript 函式的 HTML 檔案的 URL，如 [Control 元素](control.md)所定義。

下列是 **FunctionFile** 元素的範例。


```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- information about this extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```

**FunctionFile** 元素指示的 HTML 檔案中的 JavaScript 必須呼叫 `Office.initialize`，並定義採用單一參數的命名函式︰`event`。函式應該使用 [item.notificationMessages](../../reference/outlook/Office.context.mailbox.item.md) API，以指出進度、成功或失敗給使用者。它也應該在完成執行後呼叫 [event.completed](../../reference/shared/event.completed.md)。函式的名稱是在無 UI 按鈕的 **FunctionName** 元素中使用。

下列是定義 **trackMessage** 函式的 HTML 檔範例。

```js
Office.intialize = function () {
    doAuth();
}

function trackMessage (event) {
    var buttonId = event.source.id;    
    var itemId = Office.context.mailbox.item.id;
    // save this message
    event.completed();
}
```

下列程式碼示範如何實作 **FunctionName** 使用的函式。

```js
// The initialize function must be run each time a new page is loaded.
(function () {
    Office.initialize = function (reason) {
        // If you need to initialize something you can do so here.
    };
})();

// Your function must be in the global namespace.
function writeText(event) {

    // Implement your custom code here. The following code is a simple example.

    Office.context.document.setSelectedDataAsync("ExecuteFunction works. Button ID=" + event.source.id,
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === "failed") {
                // Show error message.
            }
            else {
                // Show success message.
            }
        });
    // Calling event.completed is required. event.completed lets the platform know that processing has completed.
    event.completed();
}
```

 >**重要**  **event.completed** 呼叫表示您已經成功處理事件。呼叫函式多次時，例如按相同的增益集命令多次，所有的事件會自動進入佇列。第一個事件會自動執行，而保留佇列的其他事件。當您的函式呼叫 **event.completed** 時，會執行下一個佇列的函式呼叫。您必須實作 **event.completed**；否則將無法執行您的函式。