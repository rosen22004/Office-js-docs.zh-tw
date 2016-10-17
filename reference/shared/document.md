
# <a name="document-object"></a>Document 物件
代表在與增益集互動之文件的抽象類別。

|||
|:-----|:-----|
|**主應用程式︰**|Access、Excel、PowerPoint、Project、Word|
|**已新增於**|1.0|
|**上次變更於**|1.1|

```
Office.context.document
```


## <a name="members"></a>成員


**屬性**


|**名稱**|**描述**|**支援附註**|
|:-----|:-----|:-----|
|[bindings](../../reference/shared/document.bindings.md)|取得提供文件中所定義的繫結存取的物件。|在 1.1 中，新增對 Access 內容增益集的支援。|
|[customXmlParts](../../reference/shared/document.customxmlparts.md)|取得在文件中代表自訂 XML 組件的物件。||
|[mode](../../reference/shared/document.mode.md)|取得文件所在的模式。|在 1.1 中，新增對 Access 內容增益集的支援。|
|[settings](../../reference/shared/document.settings.md)|取得代表目前文件的內容或工作窗格增益集的已儲存自訂設定的物件。|在 1.1 中，新增對 Access 內容增益集的支援。|
|[url](../../reference/shared/document.url.md)|取得主應用程式目前已開啟的文件 URL。|在 1.1 中，新增對 Access 內容增益集的支援。|

**方法**


|**名稱**|**描述**|**支援附註**|
|:-----|:-----|:-----|
|[addHandlerAsync](../../reference/shared/document.addhandlerasync.md)|新增 **Document** 物件事件的事件處理常式。||
|[getActiveViewAsync](../../reference/shared/document.getactiveviewasync.md)|傳回簡報的目前檢視。|在 1.1 中，新增支援 [PowerPoint 的增益集](../../docs/powerpoint/powerpoint-add-ins.md)。|
|[getFileAsync](../../reference/shared/document.getfileasync.md)|傳回整份文件檔案，配量最多為 4194304 位元組 (4MB)。|在 1.1 中，新增支援在 PowerPoint 和 Word 增益集中取得 PDF 檔案。|
|[getFilePropertiesAsync](../../reference/shared/document.getfilepropertiesasync.md)|取得目前文件的檔案屬性。在此版本中，只能取得文件的 URL。|在 1.1 中，新增取得 Excel、Word 和 PowerPoint 增益集中的文件 URL。|
|[getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md)|讀取文件目前選取中包含的資料。|在 1.1 中，新增支援在 PowerPoint 增益集中取得投影片所選取範圍的 id、標題和索引。|
|[goToByIdAsync](../../reference/shared/document.gotobyidasync.md)|移至文件中指定的物件或位置。|在 1.1 中，新增支援在 Excel 和 PowerPoint 增益集中文件內的導覽。|
|[removeHandlerAsync](../../reference/shared/document.removehandlerasync.md)|移除 **Document** 物件事件的事件處理常式。||
|[setSelectedDataAsync](../../reference/shared/document.setselecteddataasync.md)|將資料寫入文件中目前的選取範圍。|在 1.1 中，新增支援[在 Excel 增益集中寫入資料時，在選取的表格上設定格式](../../docs/excel/format-tables-in-add-ins-for-excel.md)。|

**件**


|**名稱**|**描述**|**支援附註**||
|:-----|:-----|:-----|:-----|
|[ActiveViewChanged](../../reference/shared/document.activeviewchanged.md)|使用者變更文件目前的檢視時，就會發生。|在 1.1 中，新增支援 PowerPoint 的增益集。||
|[SelectionChanged](../../reference/shared/document.selectionchanged.event.md)|文件中的選取項目變更時，就會發生。|||

## <a name="remarks"></a>備註

您在指令碼中未直接具現化 **Document** 物件。若要呼叫 **Document** 物件的成員，以與目前文件或工作表互動，請在指令碼中使用 `Office.context.document`。


## <a name="example"></a>範例

下列範例使用**Document** 物件的 **getSelectedDataAsync** 方法，以擷取使用者目前選擇做為文字，然後將其顯示在增益集頁面中。


```js

// Display the user's current selection.
function showSelection() {
    Office.context.document.getSelectedDataAsync(
        "text",                        // coercionType
        {valueFormat: "unformatted",   // valueFormat
        filterType: "all"},            // filterType
        function (result) {            // callback
            var dataValue; 
            dataValue = result.value;
            write('Selected data is: ' + dataValue);
        });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## <a name="support-details"></a>支援詳細資料


支援 **Document** 物件的每個 API 成員在 Office 主應用程式之間有所不同。如需瞭解主機支援資訊，請參閱每個成員主題的「支援詳細資料」一節。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


|||
|:-----|:-----|
|**已新增於**|1.0|
|**上次變更於**|1.1|
|**增益集類型**|內容、工作窗格|
|**文件庫**|Office.js|
|**命名空間**|Office|
