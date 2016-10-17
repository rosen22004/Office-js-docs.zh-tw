
# <a name="tabledata.headers-property"></a>TableData.headers 屬性
取得或設定資料表中的標題。

|||
|:-----|:-----|
|**主應用程式︰**|Excel、Word|
|**可用於[需求集合](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|TableBindings|
|**上次變更於**|1.1|

```
var hasHeaders = tableBindingObj.headers;
```


## <a name="return-value"></a>傳回值

 如果資料表有標題會傳回 **true**；沒有則傳回 **false**。 


## <a name="remarks"></a>備註

若要指定標題，您必須指定對應到資料表結構之陣列的陣列。例如，若要指定兩欄資料表的標題，您要將 **header** 屬性設定為 ` [['header1', 'header2']]`。

如果您將 **headers** 屬性指定為 **null** (或在架構 **TableData** 物件時空下屬性)，您的程式碼執行時會發生下列結果︰


- 您插入新資料表時，會建立資料表的預設資料行標題。
    
- 您覆寫或更新現有的資料表時，未變更現有的標題。
    

## <a name="example"></a>範例

下列範例會建立具一個標題與三個資料列的單欄式資料表。


```js
function createTableData() {
    var tableData = new Office.TableData();
    tableData.headers = [['header1']];
    tableData.rows = [['row1'], ['row2'], ['row3']];
    return tableData;
}

```


## <a name="support-details"></a>支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此屬性。空白儲存格表示 Office 主應用程式不支援此屬性。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。

||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**可用於需求集合**|TableBindings|
|**最低權限等級**|[限制](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**增益集類型**|內容、工作窗格|
|**文件庫**|Office.js|
|**命名空間**|Office|

## <a name="support-history"></a>支援歷程記錄




|**版本**|**變更**|
|:-----|:-----|
|1.1|新增對 Word Online 的支援。|
|1.1|新增 iPad 版 Office 中對 Excel 和 Word 的支援。|
|1.0|已導入|
