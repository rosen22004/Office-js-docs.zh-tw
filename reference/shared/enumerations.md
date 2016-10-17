
# <a name="enumerations"></a>列舉

您可以藉由使用列舉的完整格式名稱 ( `Office.CoercionType.Text`)，或其對應的文字值 ( `"text"`)，指定列舉值。例如，下列方法呼叫會使用列舉名稱：


```js
Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, {valueFormat:Office.ValueFormat.Unformatted, filterType:Office.FilterType.All},
   function (result) {
      if (result.status === Office.AsyncResultStatus.Success)
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


以下是使用列舉文字值的相同呼叫：




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
```


## <a name="reference"></a>參考資料



|**名稱**|**定義**|
|:-----|:-----|
|[ActiveView](activeview-enumeration.md)|指定文件的使用中檢視狀態，例如使用者是否可以編輯文件。|
|[AsyncResultStatus](asyncresultstatus-enumeration.md)|指定非同步呼叫的結果。|
|
  [AttachmentType](http://msdn.microsoft.com/library/83883a47-a937-4afb-a55e-e789057335c4%28Office.15%29.aspx)|指定電子郵件訊息或會議要求的附件類型。Outlook 2013 不支援此列舉。|
|[BindingType](bindingtype-enumeration.md)|指定應傳回的繫結物件類型。|
|
  [BodyType](http://msdn.microsoft.com/library/31350fe6-4c42-4cbb-a5b2-4fb2d360fa11%28Office.15%29.aspx)|指定的約會或訊息本文的文字類型。|
|[CoercionType](coerciontype-enumeration.md)|指定如何強制轉型所傳回或由叫用方法設定的資料。|
|[CustomXMLNodeType](customxmlnodetype-enumeration.md)|指定節點類型。|
|[DocumentMode](documentmode-enumeration.md)|指定相關應用程式中的文件是唯讀或讀寫。 |
|
  [EntityType](http://msdn.microsoft.com/library/0035be38-8a65-4693-bcc4-0a8dd7b1495b%28Office.15%29.aspx)|指定實體的類型。|
|[EventType](eventtype-enumeration.md)|指定引發的事件種類。|
|[FileType](filetype-enumeration.md)|指定文件要傳回的格式。|
|[GoToType](gototype-enumeration.md)|指定要瀏覽至的位置或物件類型。|
|[FilterType](filtertype-enumeration.md)|指定擷取資料時，是否從主應用程式套用篩選條件。|
|[InitializationReason](initializationreason-enumeration.md)|指定增益集是否剛插入或之前已經包含在文件中。|
|
  [ItemType](http://msdn.microsoft.com/library/e0bb23fd-f360-4b0f-b72c-1cf08d4cab3f%28Office.15%29.aspx)|指定項目的類型。|
|
  [notificationMessageType](http://msdn.microsoft.com/library/ff00c89d-0019-4545-a95b-7ed0db712ce9%28Office.15%29.aspx)|指定約會或訊息的通知訊息。|
|[ProjectProjectFields](projectprojectfields-enumeration.md)|指定可作為 [getProjectFieldAsync](projectdocument.getprojectfieldasync.md) 方法參數的專案欄位。|
|[ProjectResourceFields](projectresourcefields-enumeration.md)|指定可作為 [getResourceFieldAsync](projectdocument.gettaskfieldasync.md) 方法參數的資源欄位。|
|[ProjectTaskFields](projecttaskfields-enumeration.md)|指定可作為 [getTaskFieldAsync](projectdocument.gettaskfieldasync.md) 方法參數的工作欄位。|
|[ProjectViewTypes](projectviewtypes-enumeration.md)|指定 [getSelectedViewAsync](projectdocument.getselectedviewasync.md) 方法可以辨識的檢視類型。|
|
  [RecipientType](http://msdn.microsoft.com/library/6e7c4029-6e52-47f6-98d2-4cd3ce7bd8b4%28Office.15%29.aspx)|指定約會的收件者類型。|
|
  [ResponseType](http://msdn.microsoft.com/library/b3e723ca-4be0-4846-ad97-0eecab4355eb%28Office.15%29.aspx)|指定會議邀請的回應。|
|[SelectionMode](selectionmode-enumeration.md)|指定使用 [Document.goToByIdAsync](document.gotobyidasync.md) 方法時，是否要選取 (醒目提示) 要瀏覽至的位置。|
|
  [SourceProperty](http://msdn.microsoft.com/library/6a209a7f-57cd-4dc3-869e-07b0f5928b28%28Office.15%29.aspx)|指定由叫用方法所傳回的資料來源。|
|[Table](table-enumeration.md)|指定_資料表格式化方法_的 [cellFormat](../../docs/excel/format-tables-in-add-ins-for-excel.md) 參數中的 `cells:` 屬性列舉值。|
|[ValueFormat](valueformat-enumeration.md)|指定叫用方法所傳回的值，例如數字和日期，是否以其套用的格式設定傳回。|

## <a name="support-details"></a>支援詳細資料


支援 Office 主應用程式之間的每個列舉有所不同。如需瞭解主機支援資訊，請參閱每個列舉主題的「支援詳細資料」一節。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


|||
|:-----|:-----|
|**增益集類型**|內容、工作窗格、Outlook|
|**文件庫**|Office.js|
|**命名空間**|Office|
