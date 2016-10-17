
# <a name="get-and-set-add-in-metadata-for-an-outlook-add-in"></a>取得和設定 Outlook 增益集的增益集中繼資料

您可以使用下列其中一項來管理您 Outlook 增益集中的自訂資料︰

- 漫遊設定，可管理使用者信箱的自訂資料。
    
- 自訂屬性，可管理使用者信箱中項目的自訂資料。
    
這兩種都可為僅可透過 Outlook 增益集存取的自訂資料授與存取權，但每一種方法會與其他方法分別儲存資料。也就是自訂屬性無法存取透過漫遊設定儲存的資料，反之亦然。資料會儲存在該信箱的伺服器上，並且在增益集支援的所有表單外觀尺寸上後續的 Outlook 工作階段中可供存取。 

## <a name="custom-data-per-mailbox:-roaming-settings"></a>每個信箱的自訂資料：漫遊設定


您可以使用 [RoamingSettings](../../reference/outlook/RoamingSettings.md) 物件指定使用者的 Exchange 信箱的特定資料。這類資料的範例包括使用者的個人資料和喜好設定。當郵件增益集在任何設計為供其執行的裝置 (桌面、平板電腦或智慧型手機) 上漫遊時，可以存取漫遊設定。

 此資料的變更會儲存在這些設定的記憶體中複本以供目前的 Outlook 工作階段使用。您更新所有漫遊的設定之後應該明確儲存它們，這樣使用者下一次在相同或任何其他支援的裝置上開啟您的增益集時可供使用。


### <a name="roaming-settings-format"></a>漫遊設定格式


**RoamingSettings** 物件中的資料儲存為序列化的 JavaScript 物件標記法 (JSON) 字串。下列是結構的範例，假設有三個定義的漫遊設定，名為 `add-in_setting_name_0`、`add-in_setting_name_1` 和 `add-in_setting_name_2`。


```js
{
  "add-in_setting_name_0":"add-in_setting_value_0",
  "add-in_setting_name_1":"add-in_setting_value_1",
  "add-in_setting_name_2":"add-in_setting_value_2"
}
```


### <a name="loading-roaming-settings"></a>載入漫遊設定


郵件增益集通常會在 [Office.initialize](../../reference/shared/office.initialize.md) 事件處理常式中載入漫遊設定。下列的 JavaScript 程式碼範例會顯示如何載入現有的漫遊設定，並取得「customerName」及「customerBalance」2 個設定的值：


```js
var _mailbox;
var _settings;
var _customerName;
var _customerBalance;

// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Initialize instance variables to access API objects.
  _mailbox = Office.context.mailbox;
  _settings = Office.context.roamingSettings;
  _customerName = _settings.get("customerName");
  _customerBalance = _settings.get("customerBalance");
}

```


### <a name="creating-or-assigning-a-roaming-setting"></a>建立或指派漫遊設定


延續先前的範例，下列 JavaScript 函數 `setAddInSetting` 顯示如何使用 [RoamingSettings.set](../../reference/outlook/RoamingSettings.md) 方法以今天的日期設定名為 `cookie` 的設定，並使用 [RoamingSettings.saveAsync](../../reference/outlook/RoamingSettings.md) 方法將所有漫遊設定儲存回伺服器來保存資料。如果設定尚未存在，**set** 方法會建立設定，並將設定指派至指定的值。**saveAsync** 方法會以非同步的方式儲存漫遊設定。這個程式碼範例會將回撥方法 `saveMyAddInSettingsCallback` 傳遞至 **saveAsync**。當非同步呼叫完成時，會使用參數 _asyncResult_ 呼叫 `saveMyAddInSettingsCallback`。這個參數是 [AsyncResult](../../reference/outlook/simple-types.md) 物件，包含非同步呼叫的結果和任何相關詳細資料。您可以使用選擇性的 _userContext_ 參數從非同步呼叫傳遞任何狀態資訊至回撥函數。


```js
// Set a roaming setting.
function setAddInSetting() {
  _settings.set("cookie", Date());
  // Save roaming settings for the mailbox
  // to the server so that they will be available
  // in the next session.
  _settings.saveAsync(saveMyAddInSettingsCallback);
}

// Callback method after saving custom roaming settings.
function saveMyAddInSettingsCallback(asyncResult) {
  if (asyncResult.status == Office.AsyncResultStatus.Failed) {
    // Handle the failure.
  }
}
```


### <a name="removing-a-roaming-setting"></a>移除漫遊設定


並且延伸之前的範例，下列 JavaScript 函數 `removeAddInSetting` 會顯示如何使用 [RoamingSettings.remove](../../reference/outlook/RoamingSettings.md) 方法來移除 `cookie` 設定並將所有漫遊設定儲存回 Exchange Server。


```js
// Remove an add-in setting.
function removeAddInSetting()
{
  _settings.remove("cookie");
  // Save changes to the roaming settings for the mailbox
  // to the server so that they will be available
  // in the next session.
  _settings.saveAsync(saveMyAddInSettingsCallback);
}
```


## <a name="custom-data-per-item-in-a-mailbox:-custom-properties"></a>信箱中每個項目的自訂資料︰自訂屬性


您可以使用 [CustomProperties](../../reference/outlook/CustomProperties.md) 物件指定使用者信箱中項目的特定資料。例如，郵件增益集可分類某些郵件，並使用自訂屬性 `messageCategory` 來通知類別。或者，如果郵件增益集從郵件中的會議建議建立約會，您可以使用自訂屬性來追蹤每一個約會。如此可確保如果使用者再次開啟郵件，郵件增益集不會提供第二次建立約會的機會。

類似於漫遊設定，自訂屬性的變更會儲存在目前 Outlook 工作階段屬性的記憶體中複本。若要確定這些自訂屬性在下一個工作階段中可供使用，將所有的自訂屬性儲存到伺服器。

只能使用 **CustomProperties** 物件來存取這些新增集特定、項目特定的自訂屬性。這些屬性與 Outlook 物件模型中自訂、以 MAPI 為基礎的 [UserProperties](http://msdn.microsoft.com/library/20b49c86-d74f-9bda-382c-559af278c148%28Office.15%29.aspx) 以及 extended properties in Exchange Web 服務 (EWS) 中的延伸屬性不同。您無法使用 Outlook 物件模型或 EWS 存取 **CustomProperties**。

不過，郵件增益集可以使用 EWS [GetItem](http://msdn.microsoft.com/en-us/library/e3590b8b-c2a7-4dad-a014-6360197b68e4%28Office.15%29.aspx) 作業來取得以 MAPI 為基礎的延伸屬性。使用回撥權杖在伺服器端，或使用 **mailbox.makeEwsRequestAsync** 方法在用戶端存取 [GetItem](../../reference/outlook/Office.context.mailbox.md)。在 **GetItem** 要求中，指定您在屬性設定中需要的自訂延伸屬性。郵件增益集也可以使用 **makeEwsRequestAsync** 和 EWS [CreateItem](http://msdn.microsoft.com/library/78a52120-f1d0-4ed7-8748-436e554f75b6%28Office.15%29.aspx) 和 [UpdateItem](http://msdn.microsoft.com/library/5d027523-e0bc-4da2-b60b-0cb9fc1fdfe4%28Office.15%29.aspx) 作業來建立及修改延伸屬性。




### <a name="using-custom-properties"></a>使用自訂屬性


您必須呼叫 [loadCustomPropertiesAsync](../../reference/outlook/Office.context.mailbox.item.md) 方法來載入自訂屬性後，才能使用它們。如果任何自訂屬性已針對目前的項目進行設定，則其目前是從 Exchanger 伺服器載入。在您建立屬性包之後，您可以使用 [set](../../reference/outlook/CustomProperties.md) 及 [get](../../reference/outlook/CustomProperties.md) 方法來新增並擷取自訂屬性。若要儲存您對屬性包的任何變更，您必須使用 [saveAsync](../../reference/outlook/CustomProperties.md) 方法來保存在 Exchange Server 上的變更。


 >**附註**  由於 Outlook for Mac 不會快取自訂屬性，所以如果使用者的網路連線中斷，Outlook for Mac 中的郵件增益集便無法存取其自訂屬性。


### <a name="custom-properties-example"></a>自訂屬性範例


下列範例會針對使用自訂屬性的 Outlook 增益集顯示一組簡化的方法。針對使用自訂屬性的增益集，您可以使用這個範例做為起點。 

此範例包含下列方法︰


- [Office.initialize](../../reference/shared/office.initialize.md) -- 初始化增益集，並從 Exchange Server 載入自訂屬性包。
    
-  **customPropsCallback** -取得從伺服器傳回並將它儲存以供日後使用的自訂屬性包。
    
-  **updateProperty** -- 設定或更新特定的屬性，然後將變更儲存至伺服器。
    
-  **removeProperty** -- 從屬性包移除特定屬性，然後將移除項目儲存到伺服器。
    



```js
var _mailbox;
var _customProps;

// The initialize function is required for all add-ins.
Office.initialize = function () {
  _mailbox = Office.context.mailbox;
  _mailbox.item.loadCustomPropertiesAsync(customPropsCallback);
}

// Callback function from loading custom properties.
function customPropsCallback(asyncResult) {
  if (asyncResult.status == Office.AsyncResultStatus.Failed) {
    // Handle the failure.
  }
  else {
    // Successfully loaded custom properties,
    // can get them from the asyncResult argument.
    _customProps = asyncResult.value;
  }
}

// Get individual custom property.
function getProperty() {
  var myProp = customProps.get("myProp");
}

// Set individual custom property.
function updateProperty(name, value) {
  _customProps.set(name, value);
  // Save all custom properties to server.
  _customProps.saveAsync(saveCallback);
}

// Remove a custom property.
function removeProperty(name) {
  _customProps.remove(name);
  // Save all custom properties to server.
  _customProps.saveAsync(saveCallback);
}

// Callback function from saving custom properties.
function saveCallback() {
  if (asyncResult.status == Office.AsyncResultStatus.Failed) {
    // Handle the failure.
  }
}
```


## <a name="additional-resources"></a>其他資源

    
- 
  [MAPI 屬性概觀](http://msdn.microsoft.com/library/02e5b23f-1bdb-4fbf-a27d-e3301a359573%28Office.15%29.aspx)
    
- 
  [Outlook 屬性概觀](http://msdn.microsoft.com/library/242c9e89-a0c5-ff89-0d2a-410bd42a3461%28Office.15%29.aspx)
    
- [從 Outlook 增益集呼叫 Web 服務](../outlook/web-services.md)
    
- 
  [Exchange 的 EWS 中的屬性與延伸屬性](http://msdn.microsoft.com/library/68623048-060e-4602-b3fa-62617a94cf72%28Office.15%29.aspx)
    
- 
  [Exchange 的 EWS 中的屬性集和回應圖案](http://msdn.microsoft.com/library/04a29804-6067-48e7-9f5c-534e253a230e%28Office.15%29.aspx)
    


