
# <a name="asyncresult-object"></a>AsyncResult 物件
會封裝非同步要求結果的物件，包括狀態及錯誤資訊 (如果要求失敗)。

|||
|:-----|:-----|
|**主應用程式︰**|Access、Excel、Outlook、PowerPoint、Project、Word|
|**上次變更於**|1.1|

```
AsyncResult
```


## <a name="members"></a>成員


**屬性**


|**名稱**|**描述**|
|:-----|:-----|
|**[asyncContext](../../reference/shared/asyncresult.asynccontext.md)**|取得傳遞至叫用方法之選擇性 _asyncContext_ 參數的使用者定義項目，並保留傳遞時的狀態。|
|**[error](../../reference/shared/asyncresult.error.md)**|如果發生任何錯誤，取得提供錯誤描述的 **Error** 物件。|
|**[status](../../reference/shared/asyncresult.status.md)**|取得非同步作業的狀態。|
|**[value](../../reference/shared/asyncresult.value.md)**|如果有的話，取得這個非同步作業的裝載或內容。|

## <a name="remarks"></a>備註

當傳遞至「Async」方法之 _callback_ 參數的函數執行時，該函數會收到 [AsyncResult](../../reference/shared/asyncresult.md) 物件，您可以從回呼函數的唯一參數存取該物件。

下列範例適用於內容與工作窗格增益集。此範例顯示 [Document](../../reference/shared/document.getselecteddataasync.md) 物件之 **getSelectedDataAsync** 方法的呼叫。




```js
Office.context.document.getSelectedDataAsync("text", {valueFormat:"unformatted", filterType:"all"}, 
   function (result) {
      if (result.status === "success")      
         var dataValue = result.value; // Get selected data.
         write('Selected data is ' + dataValue);
      else {            
         var err = result.error; 
         write(err.name + ": " + err.message);
      }
   });
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}

```

以 _callback_ 引數 ( `function (result){...}`) 傳遞的匿名函數具有名稱為  _result_ 的單一參數，其在函數執行時提供 **AsyncResult** 物件的存取。當呼叫 **getSelectedDataAsync** 方法完成時，會執行回呼函數，而且下列程式碼行會存取 **AsyncResult** 物件的 **value** 屬性，以傳回文件中選取的資料：

 `var dataValue = result.value;`

請注意，函數中的程式碼行使用回呼函數的 _result_ 參數，以存取 **AsyncResult** 物件的 **status** 與 **error** 屬性。

**AsyncResult** 物件可從以引數傳遞至下列方法之 _callback_ 參數的函數中獲得：



|**父物件**|**方法**|
|:-----|:-----|
|**Document** (僅限 Excel、 PowerPoint、Project 及 Word)|[getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md)|
||[setSelectedDataAsync](../../reference/shared/document.setselecteddataasync.md)|
|**Bindings** (僅限 Excel 與 Word)|[addFromPromptAsync](../../reference/shared/bindings.addfrompromptasync.md)|
||[addFromSelectionAsync](../../reference/shared/bindings.addfromselectionasync.md)|
||[getAllAsync](../../reference/shared/bindings.getallasync.md)|
||[getByIdAsync](../../reference/shared/bindings.getbyidasync.md)|
||[releaseByIdAsync](../../reference/shared/bindings.releasebyidasync.md)|
|**Binding** (僅限 Excel 與 Word)|[getDataAsync](../../reference/shared/binding.getdataasync.md)|
||[setDataAsync](../../reference/shared/binding.setdataasync.md)|
||[removeHandlerAsync](../../reference/shared/binding.removehandlerasync.md)|
|**TableBinding** (僅限 Excel 與 Word)||
||[addRowsAsync](../../reference/shared/binding.tablebinding.addrowsasync.md)|
||[deleteAllDataValuesAsync](../../reference/shared/binding.tablebinding.deletealldatavaluesasync.md)|
|**Settings** (僅限 Excel、PowerPoint 及 Word)|[refreshAsync](../../reference/shared/settings.refreshasync.md)|
||[saveAsync](../../reference/shared/settings.saveasync.md)|
|**CustomXmlNode** (僅限 Word)|[getNodesAsync](../../reference/shared/customxmlnode.getnodesasync.md)|
||[getNodeValueAsync](../../reference/shared/customxmlnode.getnodevalueasync.md)|
||[getXmlAsync](../../reference/shared/customxmlnode.getxmlasync.md)|
||[setNodeValueAsync](../../reference/shared/customxmlnode.setnodevalueasync.md)|
||[setXmlAsync](../../reference/shared/customxmlnode.setxmlasync.md)|
|**CustomXmlPart** (僅限 Word)|[deleteAsync](../../reference/shared/customxmlpart.deleteasync.md)|
||[getNodesAsync](../../reference/shared/customxmlpart.getnodesasync.md)|
||[getXmlAsync](../../reference/shared/customxmlpart.getxmlasync.md)|
|**CustomXmlParts** (僅限 Word)|[addAsync](../../reference/shared/customxmlparts.addasync.md)|
||[getByIdAsync](../../reference/shared/customxmlparts.getbyidasync.md)|
||[getByNamespaceAsync](../../reference/shared/customxmlparts.getbynamespaceasync.md)|
|**CustomXmlPrefixMappings** (僅限 Word)|[addNamespaceAsync](../../reference/shared/customxmlprefixmappings.addnamespaceasync.md)|
||[getNamespaceAsync](../../reference/shared/customxmlprefixmappings.getnamespaceasync.md)|
||[getPrefixAsync](../../reference/shared/customxmlprefixmappings.getprefixasync.md)|
|**Mailbox** (僅限 Outlook)|
  [getUserIdentityTokenAsync](http://msdn.microsoft.com/library/c658518b-6867-41a0-99cf-810303e4c539%28Office.15%29.aspx)|
||
  [makeEwsRequestAsync](http://msdn.microsoft.com/library/2ec380e0-4a67-4146-92a6-6a39f65dc6f2%28Office.15%29.aspx)|
|**CustomProperties** (僅限 Outlook)|
  [saveAsync](http://msdn.microsoft.com/library/690d5aa9-62b5-4e5c-9548-62dfdbb5fa56%28Office.15%29.aspx)|
|**Item** (僅限 Outlook)|
  [loadCustomPropertiesAsync](http://msdn.microsoft.com/library/dfbec151-8ea7-4915-b723-09ea1396a261%28Office.15%29.aspx)|
|**RoamingSettings** (僅限 Outlook)|
  [saveAsync](http://msdn.microsoft.com/library/a616f71c-a447-423f-a0d2-e9d6f1ac32f8%28Office.15%29.aspx)|

## <a name="support-details"></a>支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此方法。空白儲存格表示 Office 主應用程式不支援此方法。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。



| |**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for iPad**|**裝置適用的 OWA**|**Outlook for Mac**|
|:-----|:-----|:-----|:-----|:-----|:-----|
|**Access**||Y||||
|**Excel**|Y|Y|Y|||
|**Outlook**|Y|Y||Y|Y|
|**PowerPoint**|Y|Y|Y|||
|**Project**|Y|||||
|**Word**|Y|Y|Y|||

|||
|:-----|:-----|
|**增益集類型**|內容、工作窗格、Outlook|
|**文件庫**|Office.js|
|**命名空間**|Office|

## <a name="support-history"></a>支援歷程記錄



****


|**版本**|**變更**|
|:-----|:-----|
|1.1|新增 iPad 版 Office 中對 Excel、PowerPoint 和 Word 的支援。|
|1.1|新增對 Access 增益集的支援。|
|1.0|已導入|
