
# TableBinding.rowCount 屬性
以整數值取得表格中的列數。

|||
|:-----|:-----|
|**主機︰**|Access、Excel、Word|
|**可用於[需求集合](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|TableBindings|
|**上次變更於 Selection**|1.1|

```
var rowCount = bindingObj.rowCount;
```


## 傳回值

在指定之 [TableBinding](../../reference/shared/binding.tablebinding.md) 物件中的列數。


## 備註

當您藉由在 Excel 2013 和 Excel Online 中選取單一列 (使用 **[插入]** 索引標籤上的**[表格]**)，插入空白表格時，Office 主應用程式會建立標頭的單一列，後面跟著單一空白列。不過，如果您的增益集指令碼為此新插入的表格建立繫結 (例如，藉由使用 [addFromSelectionAsync](../../reference/shared/bindings.addfromselectionasync.md) 方法)，然後檢查 **rowCount** 屬性的值，則傳回的值將會根據是否在 Excel 2013 或 Excel Online 中開啟試算表而有所不同。


- 在桌面上的 Excel，**rowCount** 將會傳回 0 (空白列之後的標頭不予計入)。
    
- 在 Excel Online 中，**rowCount** 將會傳回 1 (空白列之後的標頭不予計入)。
    
您可以檢查 `rowCount == 1`，以解決指令碼中的差異，如果為真，則檢查列是否包含所有空白字串。

在 Access 的內容增益集中，基於效能原因，**rowCount** 屬性一律會傳回 -1。


## 範例




```js
function showBindingRowCount() {
    Office.context.document.bindings.getByIdAsync("myBinding", function (asyncResult) {
        write("Rows: " + asyncResult.value.rowCount);
    });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## 支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此屬性。空白儲存格表示 Office 主應用程式不支援此屬性。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支援的主應用程式 (依平台排序)**


||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**可用於需求集合**|TableBindings|
|**最低權限等級**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**增益集類型**|內容、工作窗格|
|**文件庫**|Office.js|
|**命名空間**|Office|

## 支援歷程記錄



****


|**版本**|**變更**|
|:-----|:-----|
|1.1|新增 iPad 版 Office 中對 Excel 和 Word 的支援|
|1.1|新增對 Access 增益集的支援。|
|1.0|已導入|
