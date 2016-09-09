
# 了解 Outlook 增益集的權限

Outlook 增益集在其資訊清單中指定所需的權限層級。可用的層級為 **Restricted**,  **ReadItem**、**ReadWriteItem** 或 **ReadWriteMailbox**。這些權限的層級是累積的︰**Restricted** 是最低的層級，而且每一個較高的層級包含所有較低層級的權限。**ReadWriteMailbox** 包括所有支援的權限。

您從 Office 市集安裝郵件增益集之前，可以看到它所要求的權限。您也可以查看 Exchange 系統管理中心內已安裝增益集的必要權限。


## 受限制的權限



  **Restricted** 是最基本的權限層級。在資訊清單中的 **Permissions** 元素中指定 [Restricted](http://msdn.microsoft.com/en-us/library/c20cdf29-74b0-564c-e178-b75d148b36d1%28Office.15%29.aspx) 來要求此權限。如果增益集不在其資訊清單中要求特定權限，Outlook 會依預設指派此權限給郵件增益集。


### 可以執行


- [僅從項目的主旨或本文取得特定實體](../outlook/match-strings-in-an-item-as-well-known-entities.md) (電話號碼、地址、URL)。
    
- 指定需要在讀取或撰寫表單中的目前項目為特定項目類別的 [ItemIs activation rule](../outlook/manifests/activation-rules.md#itemis-rule)，或符合選取項目中任何支援的已知實體 (電話號碼、地址、URL) 的較小子集的 [ItemHasKnownEntity rule](../outlook/match-strings-in-an-item-as-well-known-entities.md)。
    
- 存取**不**關於使用者或項目的特定資訊的任何屬性和方法。(請參閱下一節查看可以執行的成員清單。)
    

### 無法執行


- 使用連絡人、電子郵件地址、會議建議或工作建議實體的 [ItemHasKnownEntity](http://msdn.microsoft.com/en-us/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx) 規則。
    
- 使用 [ItemHasAttachment](http://msdn.microsoft.com/en-us/library/031db7be-8a25-5185-a9c3-93987e10c6c2%28Office.15%29.aspx) 或 [ItemHasRegularExpressionMatch](http://msdn.microsoft.com/en-us/library/bfb726cd-81b0-a8d5-644f-2ca90a5273fc%28Office.15%29.aspx) 規則。
    
- 存取下面清單中關於使用者或項目資訊的成員。嘗試存取這份清單中的成員會傳回 **null** 和錯誤訊息中的結果，其說明 Outlook 需要郵件增益集具有更高的權限。
    
      - [item.addFileAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.addItemAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.attachments](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.bcc](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.body](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.cc](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.from](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.getRegExMatches](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.getRegExMatchesByName](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.optionalAttendees](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.organizer](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.removeAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.requiredAttendees](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.resources](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.sender](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.to](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [mailbox.getCallbackTokenAsync](../../reference/outlook/Office.context.mailbox.md)
    
  - [mailbox.getUserIdentityTokenAsync](../../reference/outlook/Office.context.mailbox.md)
    
  - [mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md)
    
  - [mailbox.userProfile](../../reference/outlook/Office.context.mailbox.userProfile.md)
    
  - [本文](../../reference/outlook/Body.md)及其所有子成員
    
  - [位置](../../reference/outlook/Location.md)及其所有子成員
    
  - [收件者](../../reference/outlook/Recipients.md)及其所有子成員
    
  - [主旨](../../reference/outlook/Subject.md)及其所有子成員
    
  - [時間](../../reference/outlook/Time.md)及其所有子成員
    

## ReadItem 權限


**ReadItem** 權限是權限模型中下一個層級的權限。在資訊清單中的 **Permissions** 元素中指定 **ReadItem** 來要求此權限。


### 可以執行


- [讀取讀取或 [撰寫表單](../outlook/get-and-set-item-data-in-a-compose-form.md)中目前項目的所有屬性](../outlook/item-data.md)，例如，讀取表單中的 [item.to](../../reference/outlook/Office.context.mailbox.item.md) 和撰寫表單中的 [item.to.getAsync](../../reference/outlook/Recipients.md)。
    
- [取得回撥權杖以取得項目附件](../outlook/get-attachments-of-an-outlook-item.md)或完整的項目。
    
- [寫入由該項目上的增益集設定的自訂屬性](http://msdn.microsoft.com/library/30217d63-7615-4f3f-8618-c91e4e60cd43%28Office.15%29.aspx)。
    
- [從項目的主旨或本文取得所有現有的已知實體](../outlook/match-strings-in-an-item-as-well-known-entities.md)，而非只是子集。
    
- 使用所有 [ItemHasKnownEntity](../outlook/manifests/activation-rules.md#itemhasknownentity-rule) 規則中的[已知實體](http://msdn.microsoft.com/en-us/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx)，或 [ItemHasRegularExpressionMatch](../outlook/manifests/activation-rules.md#itemhasregularexpressionmatch-rule) 規則中的[規則運算式](http://msdn.microsoft.com/en-us/library/bfb726cd-81b0-a8d5-644f-2ca90a5273fc%28Office.15%29.aspx)。下列範例會遵循結構描述 1.1 版。如果在所選郵件的主旨或本文中找到一個或多個已知實體，它會顯示啟動增益集的規則：
    

```XML
<Permissions>ReadItem</Permissions>
    <Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" FormType = "Read" ItemType="Message" />
    <Rule xsi:type="RuleCollection" Mode="Or">
        <Rule xsi:type="ItemHasKnownEntity" 
            EntityType="PhoneNumber" />
        <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
        <Rule xsi:type="ItemHasKnownEntity" EntityType="Url" />
        <Rule xsi:type="ItemHasKnownEntity" 
            EntityType="MeetingSuggestion" />
        <Rule xsi:type="ItemHasKnownEntity" 
            EntityType="TaskSuggestion" />
        <Rule xsi:type="ItemHasKnownEntity" 
            EntityType="EmailAddress" />
        <Rule xsi:type="ItemHasKnownEntity" EntityType="Contact" />
</Rule>
```


### 無法執行

存取 **mailbox.makeEWSRequestAsync** 或以下撰寫方法︰


- [item.addFileAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md)
    
- [item.addItemAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md)
    
- [item.bcc.addAsync](../../reference/outlook/Recipients.md)
    
- [item.bcc.setAsync](../../reference/outlook/Recipients.md)
    
- [item.body.prependAsync](../../reference/outlook/Body.md)
    
- [item.body.setAsync](../../reference/outlook/Body.md)
    
- [item.body.setSelectedDataAsync](../../reference/outlook/Body.md)
    
- [item.cc.addAsync](../../reference/outlook/Recipients.md)
    
- [item.cc.setAsync](../../reference/outlook/Recipients.md)
    
- [item.end.setAsync](../../reference/outlook/Time.md)
    
- [item.location.setAsync](../../reference/outlook/Location.md)
    
- [item.optionalAttendees.addAsync](../../reference/outlook/Recipients.md)
    
- [item.optionalAttendees.setAsync](../../reference/outlook/Recipients.md)
    
- [item.removeAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md)
    
- [item.requiredAttendees.addAsync](../../reference/outlook/Recipients.md)
    
- [item.requiredAttendees.setAsync](../../reference/outlook/Recipients.md)
    
- [item.start.setAsync](../../reference/outlook/Time.md)
    
- [item.subject.setAsync](../../reference/outlook/Subject.md)
    
- [item.to.addAsync](../../reference/outlook/Recipients.md)
    
- [item.to.setAsync](../../reference/outlook/Recipients.md)
    

## ReadWriteItem 權限


在資訊清單中的 **Permissions** 元素中指定 **ReadWriteItem** 來要求此權限。使用寫入方法 ( **Message.to.addAsync** 或 **Message.to.setAsync**) 的撰寫表單中啟用的郵件增益集必須至少使用這個層級的權限。


### 可以執行


- [讀取和寫入在 Outlook 中檢視或撰寫的項目的所有項目層級屬性](../outlook/item-data.md)。
    
- [新增或移除該項目的附件](../outlook/add-and-remove-attachments-to-an-item-in-a-compose-form.md)。
    
- 使用適用於 Office 的 JavaScript API 的所有其他成員 (可以用於郵件增益集)，除了 **Mailbox.makeEWSRequestAsync** 以外。
    

### 無法執行

使用 **Mailbox.makeEWSRequestAsync**。


## ReadWriteMailbox 權限


**ReadWriteMailbox** 權限是最高層級的權限。在資訊清單中的 **Permissions** 元素中指定 **ReadWriteMailbox** 來要求此權限。

除了 **ReadWriteItem** 權限所支援的項目外，您可以使用 **Mailbox.makeEWSRequestAsync**, 來存取支援的 Exchange Web 服務 (EWS) 作業執行下列動作︰


- 讀取和寫入使用者的信箱中任何項目的所有屬性。
    
- 建立、讀取及撰寫至任何資料夾或該信箱中的項目。
    
- 從該信箱傳送項目
    
透過 **mailbox.makeEWSRequestAsync**，您可以存取下列 EWS 作業︰


- [CopyItem](http://msdn.microsoft.com/en-us/library/bcc68f9e-d511-4c29-bba6-ed535524624a%28Office.15%29.aspx)
    
- [CreateFolder](http://msdn.microsoft.com/en-us/library/6f6c334c-b190-4e55-8f0a-38f2a018d1b3%28Office.15%29.aspx)
    
- [CreateItem](http://msdn.microsoft.com/en-us/library/78a52120-f1d0-4ed7-8748-436e554f75b6%28Office.15%29.aspx)
    
- [FindConversation](http://msdn.microsoft.com/en-us/library/2384908a-c203-45b6-98aa-efd6a4c23aac%28Office.15%29.aspx)
    
- [FindFolder](http://msdn.microsoft.com/en-us/library/7a9855aa-06cc-45ba-ad2a-645c15b7d031%28Office.15%29.aspx)
    
- [FindItem](http://msdn.microsoft.com/en-us/library/ebad6aae-16e7-44de-ae63-a95b24539729%28Office.15%29.aspx)
    
- [GetConversationItems](http://msdn.microsoft.com/en-us/library/8ae00a99-b37b-4194-829c-fe300db6ab99%28Office.15%29.aspx)
    
- [GetFolder](http://msdn.microsoft.com/en-us/library/355bcf93-dc71-4493-b177-622afac5fdb9%28Office.15%29.aspx)
    
- [GetItem](http://msdn.microsoft.com/en-us/library/e3590b8b-c2a7-4dad-a014-6360197b68e4%28Office.15%29.aspx)
    
- [MarkAsJunk](http://msdn.microsoft.com/en-us/library/1f71f04d-56a9-4fee-a4e7-d1034438329e%28Office.15%29.aspx)
    
- [MoveItem](http://msdn.microsoft.com/en-us/library/dcf40fa7-7796-4a5c-bf5b-7a509a18d208%28Office.15%29.aspx)
    
- [SendItem](http://msdn.microsoft.com/en-us/library/337b89ef-e1b7-45ed-92f3-8abe4200e4c7%28Office.15%29.aspx)
    
- [UpdateFolder](http://msdn.microsoft.com/en-us/library/3494c996-b834-4813-b1ca-d99642d8b4e7%28Office.15%29.aspx)
    
- [UpdateItem](http://msdn.microsoft.com/en-us/library/5d027523-e0bc-4da2-b60b-0cb9fc1fdfe4%28Office.15%29.aspx)
    
嘗試使用不支援的作業會導致錯誤回應。


## 其他資源



- [Outlook 增益集的隱私權、權限和安全性](../outlook/../../docs/develop/privacy-and-security.md)
    
- [使 Outlook 項目中的字串與已知的實體相符](../outlook/match-strings-in-an-item-as-well-known-entities.md)
    
