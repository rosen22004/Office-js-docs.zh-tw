
# <a name="settings.get-method"></a>Settings.get 方法
擷取指定的設定。

|||
|:-----|:-----|
|**主應用程式︰**|Access、Excel、PowerPoint、Word|
|**可用於[需求集合](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|設定|
|**上次變更於**|1.1|

```js
var mySetting = Office.context.document.settings.get(name);
```


## <a name="parameters"></a>參數



_name_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;類型：**字串**

&nbsp;&nbsp;&nbsp;&nbsp;要擷取之設定的區分大小寫名稱。

    



## <a name="return-value"></a>傳回值

屬性名稱對應至 JSON 序列化值的**物件**。


## <a name="example"></a>範例




```js
function displayMySetting() {
    write('Current value for mySetting: ' + Office.context.document.settings.get('mySetting'));
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## <a name="support-details"></a>支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此屬性。空白儲存格表示 Office 主應用程式不支援此屬性。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。



||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**PowerPoint**|Y|Y|Y|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**可用於需求集合**|Settings|
|**最低權限等級**|[限制](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**增益集類型**|內容、工作窗格|
|**文件庫**|Office.js|
|**命名空間**|Office|

## <a name="support-history"></a>支援歷程記錄




|**版本**|**變更**|
|:-----|:-----|
|1.1|新增對 PowerPoint Online 的支援。|
|1.1|新增 iPad 版 Office 中對 Excel、PowerPoint 和 Word 的支援。|
|1.1|新增支援在 Access 內容增益集中建立設定。|
|1.0|已導入|
