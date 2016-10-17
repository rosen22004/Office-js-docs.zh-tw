
# <a name="get-and-set-item-data-in-a-compose-form-in-outlook"></a>在 Outlook 中取得並設定撰寫格式的項目資料
了解如何在撰寫案例中取得或設定 Outlook 增益集項目的各種屬性，包括其收件者、主旨、本文和約會位置和時間。




## <a name="getting-and-setting-item-properties-for-a-compose-add-in"></a>取得並設定撰寫增益集的項目屬性


在撰寫表單中，您可以取得大部分與讀取表單 (例如出席者、收件者、主旨和本文等) 中相同類型項目上公開的屬性，並且可以取得幾個僅在撰寫表單而非讀取表單中 (本文、密件副本) 相關的額外屬性。 

對於大多數的這些屬性而言，因為 Outlook 增益集和使用者可能可以在相同時間在使用者介面中修改相同的屬性，因此取得及設定這些的方法都是非同步。表 1 列出項目層級屬性和對應的非同步方法，可在撰寫表單中取得及設定它們。[item.itemType](../../reference/outlook/Office.context.mailbox.item.md) 和 [item.conversationId](../../reference/outlook/Office.context.mailbox.item.md) 屬性是例外狀況，因為使用者無法修改它們。您可以直接從父系物件以讀取表單中相同的方式在撰寫表單中透過程式設計取得它們。

除了在適用於 Office 的 JavaScript API 中存取項目屬性以外，您可以使用 Exchange Web 服務 (EWS) 來存取項目層級屬性。具有 **ReadWriteMailbox** 權限，您可以使用 [mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md) 方法來存取 EWS 作業、[GetItem](http://msdn.microsoft.com/en-us/library/e3590b8b-c2a7-4dad-a014-6360197b68e4%28Office.15%29.aspx) 和 [UpdateItem](http://msdn.microsoft.com/en-us/library/5d027523-e0bc-4da2-b60b-0cb9fc1fdfe4%28Office.15%29.aspx)，以取得及設定更多項目或使用者信箱中項目的屬性。**makeEwsRequestAsync** 在撰寫和讀取表單中可供使用。如需 **ReadWriteMailbox** 權限，以及透過 Office 增益集的平台存取 EWS 的相關資訊，請參閱[了解 Outlook 增益集的權限](../outlook/understanding-outlook-add-in-permissions.md)和[從 Outlook 增益集呼叫 Web 服務](../outlook/web-services.md)。


**表 1.在撰寫表單中取得或設定項目屬性的非同步方法**


|**屬性**|**屬性類型**|**取得的非同步方法**|**設定的非同步方法**|
|:-----|:-----|:-----|:-----|
|[bcc](../../reference/outlook/Office.context.mailbox.item.md)|[Recipients](../../reference/outlook/Recipients.md)|[Recipients.getAsync](../../reference/outlook/Recipients.md)|[Recipients.addAsync](../../reference/outlook/Recipients.md)[Recipients.setAsync](../../reference/outlook/Recipients.md)|
|[body](../../reference/outlook/Office.context.mailbox.item.md)|[Body](../../reference/outlook/Body.md)|[Body.getAsync](../../reference/outlook/Body.md)|[Body.prependAsync](../../reference/outlook/Body.md)[Body.setAsync](../../reference/outlook/Body.md)[Body.setSelectedDataAsync](../../reference/outlook/Body.md)|
|[cc](../../reference/outlook/Office.context.mailbox.item.md)|Recipients|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[end](../../reference/outlook/Office.context.mailbox.item.md)|[Time](../../reference/outlook/Time.md)|[Time.getAsync](../../reference/outlook/Time.md)|[Time.setAsync](../../reference/outlook/Time.md)|
|[location](../../reference/outlook/Office.context.mailbox.item.md)|[Location](../../reference/outlook/Location.md)|[Location.getAsync](../../reference/outlook/Location.md)|[Location.setAsync](../../reference/outlook/Location.md)|
|[optionalAttendees](../../reference/outlook/Office.context.mailbox.item.md)|Recipients|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[requiredAttendees](../../reference/outlook/Office.context.mailbox.item.md)|Recipients|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[start](../../reference/outlook/Office.context.mailbox.item.md)|Time|Time.getAsync|Time.setAsync|
|[subject](../../reference/outlook/Office.context.mailbox.item.md)|[Subject](../../reference/outlook/Subject.md)|[Subject.getAsync](../../reference/outlook/Subject.md)|[Subject.setAsync](../../reference/outlook/Subject.md)|
|[to](../../reference/outlook/Office.context.mailbox.item.md)|Recipients|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|



## <a name="additional-resources"></a>其他資源



- [建立撰寫格式的 Outlook 增益集](../outlook/compose-scenario.md)
    
- [了解 Outlook 增益集的權限](../outlook/understanding-outlook-add-in-permissions.md)
    
- [從 Outlook 增益集呼叫 Web 服務](../outlook/web-services.md)
    
- [取得並設定讀取或撰寫格式的 Outlook 項目資料](../outlook/item-data.md)
    


