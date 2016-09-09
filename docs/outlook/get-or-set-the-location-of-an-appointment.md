
# 在 Outlook 中撰寫約會時，取得或設定位置

Office 的 JavaScript API 提供非同步方法 ([getAsync](../../reference/outlook/Location.md) 與 [setAsync](../../reference/outlook/Location.md)) 來取得及設定使用者正在撰寫的郵件或約會的位置。這些非同步方法僅供撰寫增益集使用。若要使用這些方法，請確定您已正確設定適用於 Outlook 的增益集資訊清單以啟動撰寫表單中的增益集，如[建立撰寫格式的 Outlook 增益集](../outlook/compose-scenario.md)所述。

[location](../../reference/outlook/Office.context.mailbox.item.md) 屬性在約會的撰寫和讀取表單中可供讀取權限使用。在讀取表單中，您可以直接從父物件存取屬性，如下︰




```js
item.location
```

但在撰寫表單中，因為使用者及您的增益集可能會同時插入或變更位置，您必須使用非同步方法 **getAsync** 來取得位置，如下所示︰




```js
item.location.getAsync
```

**location** 屬性僅在約會的撰寫表單中，而非在閱讀表單中可供撰寫權限使用。

如同適用於 Office 的 JavaScript API 中大部分的非同步方法，**getAsync** 和 **setAsync** 接受選擇性輸入參數。如需有關指定這些選擇性輸入參數的詳細資訊，請參閱[Office 增益集中的非同步程式設計](../../docs/develop/asynchronous-programming-in-office-add-ins.md)。


## 若要取得位置


本章節會顯示取得使用者正在撰寫的約會位置，及顯示位置的程式碼範例。這個程式碼範例假設增益集資訊清單中啟動約會撰寫表單中的增益集的規則，如下所示。


```XML
<Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>

```

若要使用 **item.location.getAsync**，請提供檢查非同步呼叫的狀態和結果的回撥方法。您可以透過 _asyncContext_ 選擇性參數提供任何必要的引數給回撥方法。您可以使用回撥的輸出參數 _asyncResult_ 來取得狀態、結果及任何錯誤。如果非同步呼叫成功，您可以使用 [AsyncResult.value](../../reference/outlook/simple-types.md) 屬性取得位置做為字串。




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the location of the item being composed.
        getLocation();
    });
}

// Get the location of the item that the user is composing.
function getLocation() {
    item.location.getAsync(
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully got the location, display it.
                write ('The location is: ' + asyncResult.value);
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## 若要設定位置


本章節會顯示設定使用者正在撰寫的約會位置的程式碼範例。類似於先前的範例，這個程式碼範例假設增益集資訊清單中啟動約會撰寫表單中的增益集的規則。

若要使用 **item.location.setAsync**，在資料參數中指定最多 255 個字元的字串。選擇性地，您可以在 _asyncContext_ 參數中提供回撥方法及回撥方法的任何引數。您應該檢查回呼的 _asyncResult_ 輸出參數中的狀態、結果和任何錯誤訊息。如果非同步呼叫成功，**setAsync** 插入指定的位置做為純文字，覆寫任何該項目現有的位置。




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Check for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set the location of the item being composed.
        setLocation();
    });
}

// Set the location of the item that the user is composing.
function setLocation() {
    item.location.setAsync(
        'Conference room A',
        { asyncContext: { var1: 1, var2: 2 } },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully set the location.
                // Do whatever appropriate for your scenario
                // using the arguments var1 and var2 as applicable.
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## 其他資源



- [在 Outlook 中取得並設定撰寫格式的項目資料](../outlook/get-and-set-item-data-in-a-compose-form.md)
    
- [取得並設定讀取或撰寫格式的 Outlook 項目資料](../outlook/item-data.md)
    
- [建立撰寫格式的 Outlook 增益集](../outlook/compose-scenario.md)
    
- [Office 增益集中的非同步程式設計](../../docs/develop/asynchronous-programming-in-office-add-ins.md)
    
- [在 Outlook 中撰寫約會或郵件時，取得、設定或新增收件者。](../outlook/get-set-or-add-recipients.md)
    
- [在 Outlook 中撰寫約會或郵件時，取得或設定主旨](../outlook/get-or-set-the-subject.md)
    
- [在 Outlook 中撰寫約會或郵件時，在本文中插入資料](../outlook/insert-data-in-the-body.md)
    
- [在 Outlook 中撰寫約會時，取得或設定時間](../outlook/get-or-set-the-time-of-an-appointment.md)
    
