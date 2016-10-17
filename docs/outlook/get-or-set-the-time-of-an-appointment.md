
# <a name="get-or-set-the-time-when-composing-an-appointment-in-outlook"></a>在 Outlook 中撰寫約會時，取得或設定時間

Office 的 JavaScript API 提供非同步方法 ([Time.getAsync](../../reference/outlook/Time.md) 與 [Time.setAsync](../../reference/outlook/Time.md)) 來取得及設定使用者正在撰寫的郵件或約會的開始或結束時間。這些非同步方法僅供撰寫增益集使用。若要使用這些方法，請確定您已正確設定適用於 Outlook 的增益集資訊清單以啟動撰寫表單中的增益集，如[建立撰寫格式的 Outlook 增益集](../outlook/compose-scenario.md)中所述。

[start](../../reference/outlook/Office.context.mailbox.item.md) 和 [end](../../reference/outlook/Office.context.mailbox.item.md) 屬性在撰寫和讀取表單中可供約會使用。在讀取表單中，您可以直接從父物件存取屬性，如下︰




```
item.start
```

以及於︰




```
item.end
```

但在撰寫表單中，因為使用者及您的增益集可能會同時插入或變更時間，您必須使用非同步方法 **getAsync** 來取得開始或結束時間，如下所示︰




```
item.start.getAsync
```

且：




```
item.end.getAsync
```

如同適用於 Office 的 JavaScript API 中大部分的非同步方法，**getAsync** 和 **setAsync** 接受選擇性輸入參數。如需有關指定這些選擇性輸入參數的詳細資訊，請參閱 [Office 增益集中的非同步程式設計](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-inline)中的[將選擇性參數傳遞至非同步方法](../../docs/develop/asynchronous-programming-in-office-add-ins.md)。


## <a name="to-get-the-start-or-end-time"></a>若要取得開始或結束時間


本章節會顯示取得使用者正在撰寫的約會開始時間，及顯示時間的程式碼範例。您可以使用相同的程式碼，並以 **end** 屬性取代 **start** 屬性來取得結束時間。這個程式碼範例假設增益集資訊清單中啟動約會撰寫表單中的增益集的規則，如下所示。


```XML
<Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>

```

若要使用 **item.start.getAsync** 或 **item.end.getAsync**，請提供檢查非同步呼叫的狀態和結果的回撥方法。您可以透過 _asyncContext_ 選擇性參數提供任何必要的引數給回撥方法。您可以使用回撥的輸出參數 _asyncResult_ 來取得狀態、結果及任何錯誤。如果非同步呼叫成功，您可以使用 **AsyncResult.value** 屬性取得 UTC 格式的開始時間做為 [Date](../../reference/outlook/simple-types.md) 物件。




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the start time of the item being composed.
        getStartTime();
    });
}

// Get the start time of the item that the user is composing.
function getStartTime() {
    item.start.getAsync(
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully got the start time, display it, first in UTC and 
                // then convert the Date object to local time and display that.
                write ('The start time in UTC is: ' + asyncResult.value.toString());
                write ('The start time in local time is: ' + asyncResult.value.toLocaleString());
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## <a name="to-set-the-start-or-end-time"></a>若要設定開始或結束時間


本章節會顯示設定使用者正在撰寫的約會開始時間，或使用者正在撰寫的郵件的程式碼範例。您可以使用相同的程式碼，並以 **end** 屬性取代 **start** 屬性來設定結束時間。請注意，如果約會撰寫表單已經有現有的開始時間，接下來設定開始時間會調整結束時間以維持約會的任何先前的期間。如果約會撰寫表單已經有現有的結束時間，接下來設定結束時間會調整期間及結束時間。如果約會已經設定為全天事件，則設定開始時間會將結束時間調整為晚 24 小時，並取消選取撰寫表單中全天事件的 UI。

類似於先前的範例，這個程式碼範例假設增益集資訊清單中啟動約會撰寫表單中的增益集的規則。

若要使用 **item.start.setAsync** 或 **item.end.setAsync**，在 **dateTime** 參數中以 UTC 指定_日期_值。若您收到根據使用者在用戶端輸入的日期，可以使用 [mailbox.convertToUtcClientTime](../../reference/outlook/Office.context.mailbox.md) 將值轉換成 UTC 的 **Date** 物件。您可以在 _asyncContext_ 參數中提供選擇性的回撥方法及回撥方法的任何引數。您應該檢查回呼的 _asyncResult_ 輸出參數中的狀態、結果和任何錯誤訊息。如果非同步呼叫成功，**setAsync** 插入指定的主旨做為純文字，覆寫任何該項目的現有開始或結束時間。




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set the start time of the item being composed.
        setStartTime();
    });
}

// Set the start time of the item that the user is composing.
function setStartTime() {
    var startDate = new Date("September 27, 2012 12:30:00");
    
    item.start.setAsync(
        startDate,
        { asyncContext: { var1: 1, var2: 2 } },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully set the start time.
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


## <a name="additional-resources"></a>其他資源



- [在 Outlook 中取得並設定撰寫格式的項目資料](../outlook/get-and-set-item-data-in-a-compose-form.md)
    
- [取得並設定讀取或撰寫格式的 Outlook 項目資料](../outlook/item-data.md)
    
- [建立撰寫格式的 Outlook 增益集](../outlook/compose-scenario.md)
    
- [Office 增益集中的非同步程式設計](../../docs/develop/asynchronous-programming-in-office-add-ins.md)
    
- [在 Outlook 中撰寫約會或郵件時，取得、設定或新增收件者。](../outlook/get-set-or-add-recipients.md)
    
- [在 Outlook 中撰寫約會或郵件時，取得或設定主旨](../outlook/get-or-set-the-subject.md)
    
- [在 Outlook 中撰寫約會或郵件時，在本文中插入資料](../outlook/insert-data-in-the-body.md)
    
- [在 Outlook 中撰寫約會時，取得或設定位置](../outlook/get-or-set-the-location-of-an-appointment.md)
    
