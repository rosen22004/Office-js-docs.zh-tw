
# 保存增益集狀態和設定

Office 增益集基本上是在瀏覽器控制項的無狀態環境中執行的 Web 應用程式。因此，您的增益集可能需要保存資料，以在使用增益集的各個工作階段中保有特定作業或功能的持續性。例如，增益集可能有自訂的設定或必須儲存並在下次初始化時重新載入的其他值，例如使用者偏好的檢視或預設位置。

若要這樣做，您可以︰


- 使用適用於 Office 的 JavaScript API 的成員，其會在因增益集類型而定的位置的屬性包中將資料儲存為名稱/值組。
    
- 使用基礎瀏覽器控制項所提供的技術︰瀏覽器的 Cookie 或 HTML5 Web 儲存區 ([localStorage](http://msdn.microsoft.com/en-us/library/cc848902%28v=vs.85%29.aspx) 或 [sessionStorage](http://msdn.microsoft.com/en-us/library/cc197020%28v=vs.85%29.aspx))。
    
本文著重於如何使用適用於 Office 的 JavaScript API 來保存增益集的狀態。如需使用瀏覽器 Cookie 和網頁儲存的範例，請參閱 [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)。

## 使用適用於 Office 的 JavaScript API 保存增益集狀態和設定


適用於 Office 的 JavaScript API 提供 [Settings](../../reference/shared/settings.md)、[RoamingSettings](../../reference/outlook/RoamingSettings.md) 和 [CustomProperties](../../reference/outlook/CustomProperties.md) 物件，用於儲存各個工作階段增益集的狀態，如下表所述。在所有情況下，儲存的設定值會與建立它們的增益集的[識別碼](http://msdn.microsoft.com/en-us/library/67c4344a-935c-09d6-1282-55ee61a2838b%28Office.15%29.aspx)關聯。



|**物件**|**增益集類型支援**|**儲存位置**|**Office 主應用程式支援**|
|:-----|:-----|:-----|:-----|
|[Settings](../../reference/shared/settings.md)|內容和工作窗格|文件、試算表或簡報增益集正在使用。內容和工作窗格增益集設定可供建立它們的增益集從儲存所在的文件中使用。**重要事項：**不要利用 **Settings** 物件來儲存密碼和其他機密個人識別資訊 (PII)。使用者看不到儲存的資料，但它會隨著文件的一部分儲存，透過直接讀取文件的檔案格式即可存取。您應該限制增益集使用的 PII，並只有在主控增益集做為使用者保護資源的伺服器上，才儲存增益集所需的任何 PII。|Word、Excel 或 PowerPoint **附註︰**Project 2013 的工作窗格增益集不支援用來儲存增益集的狀態或設定的**設定** API。不過，對於在 Project (以及其他的 Office 主應用程式) 中執行的增益集，您可以使用瀏覽器 Cookie 或 Web 儲存區之類的技術。如需有關這些技術的詳細資訊，請參閱 [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)。 |
|[RoamingSettings](../../reference/outlook/RoamingSettings.md)|Outlook|安裝增益集所在的使用者的 Exchange 伺服器信箱。因為這些設定會儲存在使用者的伺服器信箱，它們可以隨著使用者「漫遊」，並且當增益集在存取該使用者的任何支援的用戶端主應用程式或瀏覽器的信箱內容中執行時，可供增益集使用。Outlook 增益集漫遊設定僅供建立它們的增益集使用，而且僅能從安裝增益集所在的信箱使用。|Outlook|
|[CustomProperties](../../reference/outlook/CustomProperties.md)|Outlook|增益集正在使用的訊息、約會或會議要求項目。Outlook 增益集項目自訂屬性僅供建立它們的增益集使用，而且僅能從儲存增益集所在的項目使用。|Outlook|

## 設定資料在執行階段時是在記憶體中管理


使用 **Settings**、**CustomProperties** 或 **RoamingSettings** 物件存取的屬性包中的資料，會在內部儲存為包含名稱/值組的序列化 JavaScript 物件標記法 (JSON) 物件。每個值的名稱 (機碼) 必須是 **string**，儲存值可以是 JavaScript **string**、**number**、**date** 或 **object**，但不能是 **function**。

屬性包結構的這個範例包含名為 `firstName`、`location` 和 `defaultView` 的三個定義的 **string** 值。




```
{
"firstName":"Erik",
"location":"98052",
"defaultView":"basic"
}
```

在先前的增益集工作階段期間儲存設定屬性包之後，可以在初始化增益集時載入，或在增益集的目前工作階段之後的任何點載入。在工作階段期間，設定會完全在記憶體中使用對應於您要建立的類型設定(**Settings**、**CustomProperties** 或 **RoamingSettings**) 的物件的 **get**、**set** 和 **remove** 方法管理。 


 >**重要：**若要將在增益集目前工作階段進行的任何新增、更新或刪除保存在儲存位置中，您必須呼叫用來處理這種設定的對應物件的 **saveAsync** 方法。**get**、**set** 和 **remove** 方法只能在設定屬性包的記憶體內部複本上作業。如果增益集關閉而未呼叫 **saveAsync**，該工作階段期間對設定所做的任何變更將會遺失。 


## 如何依據內容和工作窗格增益集的文件來儲存增益集狀態和設定


若要保存 Word、Excel 或 PowerPoint 的內容或工作窗格增益集的狀態或自訂設定，您會使用 [Settings](../../reference/shared/settings.md) 物件和其方法。使用 **Settings** 物件的方法建立的屬性包，僅供建立它的內容或工作窗格增益集的執行個體使用，而且僅能從儲存它所在的文件使用。

**Settings** 物件會自動載入為 [Document](../../reference/shared/document.md) 物件的一部分，並且在工作窗格或內容增益集啟動時可供使用。在 **Document** 物件具現化之後，您可以使用**Document** 物件的 [settings](../../reference/shared/document.settings.md) 屬性來存取 **Settings** 物件。在工作階段的存留期內，您可以只使用 **Settings.get**、**Settings.set**，和 **Settings.remove** 方法，從屬性包的記憶體中複本讀取、寫入或移除保存的設定和增益集的狀態。

因為 set 和 remove 方法只對設定屬性包的記憶體中複本運作，若要將新的或變更的設定儲存回增益集相關聯的文件，您必須呼叫 [Settings.saveAsync](../../reference/shared/settings.saveasync.md) 方法。


### 建立或更新設定值

下列程式碼範例示範如何使用 [Settings.set](../../reference/shared/settings.set.md) 方法來建立名為 `'themeColor'` 值 `'green'` 的設定。set 方法的第一個參數為設定區分大小寫的 _name_ (Id)，用來設定或建立。第二個參數是設定的_值_。


```
Office.context.document.settings.set('themeColor', 'green');
```

 會建立具有指定名稱的設定 (如果它不存在)，或是會更新它的值 (如果它存在)。使用 **Settings.saveAsync** 方法來保存新的或更新的設定至文件。


### 取得設定的值

下列範例顯示如何使用 [Settings.get](../../reference/shared/settings.get.md) 方法來取得名為 "themeColor" 之設定的值。**get** 方法的唯一參數為設定區分大小寫的 _name_。


```js
write('Current value for mySetting: ' + Office.context.document.settings.get('themeColor'));

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

 **get** 方法會傳回針對傳入的設定 _name_ 先前所儲存的值。如果設定不存在，則這個方法會傳回 **null**。


### 移除設定

下列範例示範如何使用 [Settings.remove](../../reference/shared/settings.removehandlerasync.md) 方法來移除名為 "themeColor" 的設定。**remove** 方法的唯一參數為設定區分大小寫的 _name_。


```
Office.context.document.settings.remove('themeColor');
```

如果設定不存在，則不會發生任何事。使用 **Settings.saveAsync** 方法來保存從文件中設定的移除項目。


### 儲存您的設定

若要儲存在目前的工作階段期間，增益集對設定屬性包的記憶體中複本所做的任何新增、變更或刪除，您必須呼叫 [Settings.saveAsync](../../reference/shared/settings.saveasync.md) 方法，將它們儲存在文件中。**saveAsync** 方法的唯一的參數是 _callback_，這是一個具有單一參數的回撥函式。 


```js
Office.context.document.settings.saveAsync(function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Settings save failed. Error: ' + asyncResult.error.message);
    } else {
        write('Settings saved.');
    }
});
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

傳遞至 **saveAsync** 方法做為 _callback_ 參數的匿名函式會在作業完成時執行。回撥的 _asyncResult_ 參數提供包含作業狀態之 **AsyncResult** 物件的存取。在範例中，函式會檢查 **AsyncResult.status** 屬性，查看儲存作業是否成功或失敗，然後在增益集頁面中顯示結果。


## 如何為 Outlook 增益集將設定儲存於使用者信箱做為漫遊設定


Outlook 增益集可以使用 [RoamingSettings](../../reference/outlook/RoamingSettings.md) 物件來儲存使用者信箱的特定增益集狀態和設定資料。此資料只可供該 Outlook 增益集代表執行增益集的使用者存取。資料會儲存在使用者的 Exchange Server 信箱上，在該使用者登入其帳戶並執行 Outlook 增益集時可供存取。


### 載入漫遊設定


Outlook 增益集通常會在 [Office.initialize](../../reference/shared/office.initialize.md) 事件處理常式中載入漫遊設定。下列的 JavaScript 程式碼範例會示範如何載入現有的漫遊設定。


```
var _mailbox;
var _settings;

// The initialize function is required for all add-ins.
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, add-in-specific code can run.
   // Initialize instance variables to access API objects.
    _mailbox = Office.context.mailbox;
    _settings = Office.context.roamingSettings;
    });
}

```


### 建立或指派漫遊設定


延續先前的範例，下列 `setAppSetting` 函數顯示如何使用 [RoamingSettings.set](../../reference/outlook/RoamingSettings.md) 方法以今天的日期設定或更新名為 `cookie` 的設定。然後，它會使用 [RoamingSettings.saveAsync](../../reference/outlook/RoamingSettings.md) 方法，將所有漫遊的設定儲存回 Exchange Server。


```
// Set an add-in setting.
function setAppSetting() {
    _settings.set("cookie", Date());
    _settings.saveAsync(saveMyAppSettingsCallback);
}

// Saves all roaming settings.
function saveMyAppSettingsCallback(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        // Handle the failure.
    }
}
```

**saveAsync** 方法會以非同步方式儲存漫遊設定，並採用一個選擇性的回撥函式。這個程式碼範例會傳遞名為 `saveMyAppSettingsCallback` 的回撥函式到 **saveAsync** 方法。非同步呼叫傳回時，`saveMyAppSettingsCallback` 函式的 _asyncResult_ 參數會提供 [AsyncResult](../../reference/outlook/simple-types.md) 物件的存取，您可以使用該物件利用 **AsyncResult.status** 屬性來判斷作業成功或失敗。


### 移除漫遊設定


並且延伸之前的範例，下列 `removeAppSetting` 函數會顯示如何使用 [RoamingSettings.remove](../../reference/outlook/RoamingSettings.md) 方法來移除 `cookie` 設定並將所有漫遊設定儲存回 Exchange Server。


```
// Remove an application setting.
function removeAppSetting()
{
    _settings.remove("cookie");
    _settings.saveAsync(saveMyAppSettingsCallback);
}
```


## 如何將 Outlook 增益集的每個項目的設定儲存為自訂屬性


自訂屬性可讓您的 Outlook 增益集儲存其正在使用的項目的相關資訊。例如，如果 Outlook 增益集在郵件中透過會議建議建立了約會，您可以使用自訂屬性來儲存會議已建立的事實。如此可確保如果再次開啟郵件，Outlook 增益集不會再次提供建立約會的機會。

在可以對特定訊息、約會或會議要求項目使用自訂屬性之前，您必須藉由呼叫 [Item](../../reference/outlook/Office.context.mailbox.item.md) 物件的 **loadCustomPropertiesAsync** 方法將屬性載入記憶體。如果有任何自訂屬性已針對目前的項目進行設定，則在此時會從 Exchanger 伺服器載入。載入屬性之後，您可以使用 [CustomProperties](../../reference/outlook/CustomProperties.md) 物件的 [set](../../reference/outlook/RoamingSettings.md) 和 **get** 的方法來加入、更新和擷取記憶體中的屬性。若要儲存您對項目的自訂屬性進行的任何變更，您必須使用 [saveAsync](../../reference/outlook/CustomProperties.md) 方法來將對項目所做的變更保存至 Exchange 伺服器上。


### 自訂屬性範例

下列範例會針對使用自訂屬性的 Outlook 增益集顯示一組簡化的函式。針對使用自訂屬性的 Outlook 增益集，您可以使用這個範例做為起點。 

使用這些函式的 Outlook 增益集會藉由呼叫 `_customProps` 變數上的 **get** 方法來擷取任何自訂屬性，如下列範例所示。




```
var property = _customProps.get("propertyName");
```

此範例包含下列功能︰



|**函數名稱**|**說明**|
|:-----|:-----|
| `Office.initialize`|初始化增益集，並從 Exchange 伺服器載入目前項目的自訂屬性。|
| `customPropsCallback`|取得從 Exchange 伺服器所傳回的自訂屬性，並儲存以供日後使用。|
| `updateProperty`|設定或更新特定屬性，然後將變更儲存至 Exchange 伺服器。|
| `removeProperty`|移除特定屬性，然後將移除動作保存至 Exchange 伺服器。|
| `saveCallback`|對 `updateProperty` 和 `removeProperty` 函式中 **saveAsync** 方法之呼叫的回撥。|



```
var _mailbox;
var _customProps;

// The initialize function is required for all add-ins.
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, add-in-specific code can run.
    _mailbox = Office.context.mailbox;
    _mailbox.item.loadCustomPropertiesAsync(customPropsCallback);
    });
}

// Get the item's custom properties from the server and save for later use.
function customPropsCallback(asyncResult) {
    _customProps = asyncResult.value;
}

// Sets or updates the specified property, and then saves the change 
// to the server.
function updateProperty(name, value) {
    _customProps.set(name, value);
    _customProps.saveAsync(saveCallback);
}

// Removes the specified property, and then persists the removal 
// to the server.
function removeProperty(name) {
   _customProps.remove(name);
   _customProps.saveAsync(saveCallback);
}

// Callback for calls to saveAsync method. 
function saveCallback(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        // Handle the failure.
    }
}
```


## 其他資源



- [了解適用於 Office 的 JavaScript API](../../docs/develop/understanding-the-javascript-api-for-office.md)
    
- [Outlook 增益集](../outlook/outlook-add-ins.md)
    
- [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)
    
