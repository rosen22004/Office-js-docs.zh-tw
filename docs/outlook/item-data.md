
# <a name="get-and-set-outlook-item-data-in-read-or-compose-forms"></a>取得並設定讀取或撰寫格式的 Outlook 項目資料

自 Office 增益集資訊清單結構描述 1.1 版開始，當使用者在檢視或撰寫項目時，Outlook 可以啟動增益集。依增益集是否在讀取或撰寫表單中啟動而定，可供項目上的增益集使用的屬性也會有所不同。例如，僅針對已傳送的項目 (隨後在讀取表單中檢視的項目) 而非在項目已建立 (在撰寫表單中) 定義 [dateTimeCreated](../../reference/outlook/Office.context.mailbox.item.md) 和 [dateTimeModified](../../reference/outlook/Office.context.mailbox.item.md) 屬性。另一個範例為 [bcc](../../reference/outlook/Office.context.mailbox.item.md) 屬性，其僅在郵件撰寫 (在撰寫表單中) 時才有意義，且使用者在讀取表單中無法存取。

表 1 顯示適用於 Office 的 JavaScript API 中的項目層級屬性，其在郵件增益集的每個讀取及撰寫模式中可供使用。通常，讀取表單中可用的這些屬性是唯讀，而撰寫表單中的屬性是讀寫，除了 [itemId](../../reference/outlook/Office.context.mailbox.item.md) 和 [conversationId](../../reference/outlook/Office.context.mailbox.item.md) 屬性之外，其無論如何一律是唯讀。針對撰寫表單中其餘的可用項目層級屬性，因為增益集及使用者可能會同時讀取或撰寫相同的屬性，在撰寫模式中取得或設定它們的方法為非同步，因此這些屬性傳回的物件類型在撰寫表單中也會與讀取表單中不同。如需有關使用非同步方法在撰寫模式中取得或設定項目層級屬性的詳細資訊，請參閱[在 Outlook 中取得並設定撰寫格式的項目資料](../outlook/get-and-set-item-data-in-a-compose-form.md)。


**表 1.撰寫和讀取表單中可用的項目屬性**


|**項目類型**|**屬性**|**讀取表單中的屬性類型**|**撰寫表單中的屬性類型**|
|:-----|:-----|:-----|:-----|
|約會與郵件|[dateTimeCreated](../../reference/outlook/Office.context.mailbox.item.md)|JavaScript **Date** 物件|屬性無法使用|
|約會與郵件|[dateTimeModified](../../reference/outlook/Office.context.mailbox.item.md)|JavaScript **Date** 物件|屬性無法使用|
|約會與郵件|[itemClass](../../reference/outlook/Office.context.mailbox.item.md)|String|屬性無法使用|
|約會與郵件|[itemId](../../reference/outlook/Office.context.mailbox.item.md)|String|屬性無法使用|
|約會與郵件|[itemType](../../reference/outlook/Office.context.mailbox.item.md)|[ItemType](../../reference/outlook/Office.MailboxEnums.md) 列舉中的字串|屬性無法使用|
|約會與郵件|[attachments](../../reference/outlook/Office.context.mailbox.item.md)|[AttachmentDetails](../../reference/outlook/simple-types.md)|屬性無法使用|
|約會與郵件|[body](../../reference/outlook/Office.context.mailbox.item.md)|[Body](../../reference/outlook/Body.md)|[Body](../../reference/outlook/Body.md)|
|約會|[end](../../reference/outlook/Office.context.mailbox.item.md)|JavaScript **Date** 物件|[Time](../../reference/outlook/Time.md)|
|約會|[location](../../reference/outlook/Office.context.mailbox.item.md)|String|[Location](../../reference/outlook/Location.md)|
|約會與郵件|[normalizedSubject](../../reference/outlook/Office.context.mailbox.item.md)|String|屬性無法使用|
|約會|[optionalAttendees](../../reference/outlook/Office.context.mailbox.item.md)|[EmailAddressDetails](../../reference/outlook/simple-types.md)|[Recipients](../../reference/outlook/Recipients.md)|
|約會|[organizer](../../reference/outlook/Office.context.mailbox.item.md)|EmailAddressDetails|屬性無法使用|
|約會|[requiredAttendees](../../reference/outlook/Office.context.mailbox.item.md)|EmailAddressDetails|收件者|
|約會|[resources](../../reference/outlook/Office.context.mailbox.item.md)|String|屬性無法使用|
|約會|[start](../../reference/outlook/Office.context.mailbox.item.md)|JavaScript **Date** 物件|時間|
|約會與郵件|[subject](../../reference/outlook/Office.context.mailbox.item.md)|String|[Subject](../../reference/outlook/Subject.md)|
|郵件|[bcc](../../reference/outlook/Office.context.mailbox.item.md)|屬性無法使用|收件者|
|郵件|[cc](../../reference/outlook/Office.context.mailbox.item.md)|EmailAddressDetails|收件者|
|郵件|[conversationId](../../reference/outlook/Office.context.mailbox.item.md)|String|字串 (唯讀)|
|郵件|[from](../../reference/outlook/Office.context.mailbox.item.md)|EmailAddressDetails|屬性無法使用|
|郵件|[internetMessageId](../../reference/outlook/Office.context.mailbox.item.md)|Integer|屬性無法使用|
|郵件|[sender](../../reference/outlook/Office.context.mailbox.item.md)|EmailAddressDetails|屬性無法使用|
|郵件|[to](../../reference/outlook/Office.context.mailbox.item.md)|EmailAddressDetails|收件者|

## <a name="using-exchange-server-callback-tokens-from-a-read-add-in"></a>使用 Exchange Server 從讀取增益集回撥權杖


如果 Outlook 增益集在讀取表單中啟動，您可以取得 Exchange 回撥權杖。這個權杖可在伺服器端程式碼中使用以透過 Exchange Web 服務 (EWS) 存取完整項目。藉由在增益集資訊清單中指定 **ReadItem** 權限，您可以使用 [mailbox.getCallbackTokenAsync](../../reference/outlook/Office.context.mailbox.md) 方法來取得 Exchange 回撥權杖、[mailbox.ewsUrl](../../reference/outlook/Office.context.mailbox.md) 屬性來取得 EWS 端點的 URL 以存取使用者的信箱，及 [item.itemId](../../reference/outlook/Office.context.mailbox.item.md) 來取得 EWS ID 以存取選取的項目。您接著可以將回撥權杖、EWS 端點 URL 及 EWS 項目 ID 傳遞至伺服器端程式碼，以存取 [GetItem](http://msdn.microsoft.com/en-us/library/e3590b8b-c2a7-4dad-a014-6360197b68e4%28Office.15%29.aspx) 作業，以取得更多項目的屬性。


## <a name="accessing-ews-from-a-read-or-compose-add-in"></a>從讀取或撰寫增益集存取 EWS


您也可以使用 [mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md) 方法，直接從增益集存取 Exchange Web 服務 (EWS) 作業 [GetItem](http://msdn.microsoft.com/en-us/library/e3590b8b-c2a7-4dad-a014-6360197b68e4%28Office.15%29.aspx) 和 [UpdateItem](http://msdn.microsoft.com/en-us/library/5d027523-e0bc-4da2-b60b-0cb9fc1fdfe4%28Office.15%29.aspx)。您可以使用這些作業來取得及設定指定的項目中的許多屬性。這個方法可用於 Outlook 增益集，無論增益集是否已在讀取或撰寫表單中啟動，只要您在增益集資訊清單中指定 **ReadWriteMailbox** 權限。如需使用 **makeEwsRequestAsync** 來存取 EWS 作業的詳細資訊，請參閱[從 Outlook 增益集呼叫 Web 服務](../outlook/web-services.md)


## <a name="additional-resources"></a>其他資源



- [Outlook 增益集](../outlook/outlook-add-ins.md)
    
- [在 Outlook 中取得並設定撰寫格式的項目資料](../outlook/get-and-set-item-data-in-a-compose-form.md)
    
- [從 Outlook 增益集呼叫 Web 服務](../outlook/web-services.md)
    


