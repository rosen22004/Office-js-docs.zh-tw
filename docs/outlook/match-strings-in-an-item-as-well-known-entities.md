

# 使 Outlook 項目中的字串與已知的實體相符


在傳送郵件或會議要求項目之前，Exchange Server 會剖析項目的內容、識別並戳記主旨及本文中的特定字串，其類似於 Exchange 的已知實體，例如，電子郵件地址、電話號碼和 URL。郵件及會議要求在 Outlook Inbox 收件匣中使用戳記的已知實體由 Exchange Server 所傳遞。 

您可以使用適用於 Office 的 JavaScript API，來取得符合特定已知實體的這些字串，以供進一步處理。您也可以在增益集資訊清單的規則中指定已知的實體，以便使用者在檢視包含該實體的符合項目的項目時，Outlook 可以啟動增益集。然後，您可以解壓縮及在實體的相符項目上採取動作。 

能夠從選取的郵件或約會識別或擷取這些執行個體很方便。例如，您可以建立反向電話查詢服務做為 Outlook 增益集。增益集可以擷取項目主旨或本文中類似於電話號碼的字串、反向查閱並顯示每個電話號碼的註冊擁有人。

本主題會介紹這些已知的項目，根據已知的實體以及如何擷取與已使用啟動規則中的實體無關的實體相符項目來顯示啟動規則的範例。


## 已知實體的支援


在寄件者傳送項目之後及 Exchange 將項目傳遞至收件者之前，Exchange Server 會在郵件或會議要求項目中戳記已知實體。因此，只有在 Exchange 中已經完成傳輸的項目會加上戳記，當使用者在檢視這類項目時，Outlook 可以根據這些戳記來啟動增益集。相反地，當使用者在撰寫或檢視位於 [寄件備份] 資料夾中的項目時，因為項目尚未經過傳輸，Outlook 無法根據已知的實體來啟動增益集。 

同樣地，您無法在所要撰寫或在 [寄件備份] 資料夾中的項目擷取已知實體，因為這些項目尚未經過傳輸且未戳記。如需有關支援啟動的項目類型的詳細資訊，請參閱 [Outlook 增益集的啟用規則](../outlook/manifests/activation-rules.md#activation-rules-for-outlook-add-ins)。

下表列出 Exchange Server 及 Outlook 支援和辨識的實體 (因此名為「已知實體」)，以及每個實體的執行個體的物件類型。做為這些實體其中一個的字串自然語言辨識是根據受過大量資料訓練的學習模型。因此，辨識不具有決定性。如需有關辨識的條件的詳細資訊，請參閱[使用已知實體的秘訣](#使用已知實體的秘訣)。

 **表 1.支援的實體及其類型**



|**實體類型**|**辨識的條件**|**物件類型**|
|:-----|:-----|:-----|
|**Address**|美國街道地址；例如：1234 Main Street, Redmond, WA 07722。一般而言，若要識別地址，必須遵循美國郵寄地址的結構，包含街道號碼、街道名稱、城市、州和郵遞區號呈現。可以在一或多個行中指定地址。|JavaScript  **String** 物件|
|**連絡人**|自然語言中所識別的人員資訊的參考。連絡人的辨識視內容而定。例如，郵件結尾處的簽章，或人員的名稱出現在一些下列資訊附近︰電話號碼、地址、電子郵件地址及 URL。|[Contact](../../reference/outlook/simple-types.md) 物件|
|**EmailAddress**|SMTP 電子郵件地址。|JavaScript  **String** 物件|
|**MeetingSuggestion**|事件或會議的參考。例如，Exchange 2013 會將下列文字識別為會議建議：_我們明天見面一起用午餐。_|[MeetingSuggestion](../../reference/outlook/simple-types.md) 物件|
|**PhoneNumber**|美國電話號碼；例如︰_(235) 555-0110_|[PhoneNumber](../../reference/outlook/simple-types.md) 物件|
|**TaskSuggestion**|在電子郵件中可採取動作的句子。例如：_請更新試算表。_|[TaskSuggestion](../../reference/outlook/simple-types.md) 物件|
|**URL**|明確指定網路位置與 web 資源的識別項的網址。Exchange Server 不需要在網址中的存取通訊協定，且無法識別內嵌在連結文字中做為 **Url** 實體的執行個體的 URL。Exchange Server 可以符合下列的範例︰_www.youtube.com/user/officevideos_ _http://www.youtube.com/user/officevideos_|JavaScript  **String** 物件|
圖 1 說明 Exchange Server 和 Outlook 如何支援增益集的已知實體，以及增益集可以使用已知實體進行的作業。如需有關如何使用這些實體的詳細資訊，請參閱[在增益集中擷取實體](#在增益集中擷取實體)及[根據實體存在與否啟動增益集](#根據實體存在與否啟動增益集)。


**圖 1.Exchange Server、Outlook 和增益集如何支援已知的實體**

![Support and use of well-known entities in mail app](../../images/mod_off15_mailapp_wellknownentities_curvedlines.png)


## 擷取實體的權限


若要擷取 JavaScript 程式碼中的實體，或根據特定已知實體存在與否啟動增益集，請確定您已在增益集資訊清單中要求適當的權限。

指定預設受限制的權限可讓您的增益集擷取 **Address**、**MeetingSuggestion** 或 **TaskSuggestion** 實體。 若要擷取任何其他的實體，請指定讀取項目、讀寫項目或讀寫信箱權限。 若要在資訊清單中如此做，請使用 [Permissions](http://msdn.microsoft.com/en-us/library/c20cdf29-74b0-564c-e178-b75d148b36d1%28Office.15%29.aspx) 元素並指定適當的權限 - **Restricted**、**ReadItem**、**ReadWriteItem** 或 **ReadWriteMailbox** - 如下列範例所示︰




```XML
<Permissions>ReadItem</Permissions>
```


## 在增益集中擷取實體


只要使用者所檢視項目的主旨或本文包含 Exchange 與 Outlook 可以辨識為已知實體的字串，這些執行個體便可供增益集使用。即使未根據已知實體啟動增益集，它們還是可供使用。您可以利用適當的權限，使用 **getEntities** 或 **getEntitiesByType** 方法來擷取出現在目前郵件或約會的已知實體。**getEntities** 方法會傳回 [Entities](../../reference/outlook/simple-types.md) 物件的陣列，該陣列在項目中包含所有已知實體。如果您對實體的特定類型有興趣，使用會僅傳回您所需實體之陣列的 **getEntitiesByType** 方法。[EntityType](../../reference/outlook/Office.MailboxEnums.md) 列舉代表所有您可以擷取的已知實體的類型。

在呼叫 **getEntities** 之後，您便可以使用 **Entities** 物件的對應屬性來取得實體類型的執行個體陣列。依實體類型而定，陣列中的執行個體可以只是字串，或可以對應至特定的物件。如圖 1 中所示的範例，若要取得項目中的地址，存取由 `getEntities().addresses[]` 傳回的陣列。**Entities.addresses** 屬性會傳回 Outlook 視為郵寄地址的字串陣列。同樣地，**Entities.contacts** 會屬性傳回 Outlook 視為連絡資訊的 **Contact** 物件陣列。表 1 列出每個受支援實體的執行個體的物件類型。

下列範例會顯示如何擷取在郵件中找到的任何地址。




```
// Get the address entities from the item.
var entities = Office.context.mailbox.item.getEntities();
// Check to make sure that address entities are present.
if (null != entities &amp;&amp; null != entities.addresses &amp;&amp; undefined != entities.addresses) {
   //Addresses are present, so use them here.
}

```


## 根據實體存在與否啟動增益集


使用已知實體的另一個方法是讓 Outlook 根據目前檢視項目的主旨或本文中一或多個實體類型存在與否啟動增益集。您可以在增益集資訊清單中指定 **ItemHasKnownEntity** 來執行這項操作。[KnownEntityType](http://msdn.microsoft.com/en-us/library/432d413b-9fcc-eb50-cfea-0ed10a43bd52%28Office.15%29.aspx) 簡單類型代表由 **ItemHasKnownEntity** 規則所支援的不同類型的已知實體。啟動增益集之後，您也可以為您的目的擷取這類實體的執行個體，如前一節[在增益集中擷取實體](#在增益集中擷取實體)中所述。 

您可以選擇性地在 **ItemHasKnownEntity** 規則中套用規則運算式，以便進一步篩選實體的執行個體，並且使 Outlook 僅在實體的執行個體子集上啟動增益集。例如，您可以在包含以「98」為開頭的華盛頓州郵遞區號的郵件中指定街道地址實體的篩選器。若要在實體執行個體上套用篩選器，請使用 **ItemHasKnownEntity** 類型的 **Rule** 元素中的 [RegExFilter](http://msdn.microsoft.com/en-us/library/56dfc32e-2b8c-1724-05be-5595baf38aa3%28Office.15%29.aspx) 和 [FilterName](http://msdn.microsoft.com/en-us/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx) 屬性。

類似於其他的啟用規則，您可以指定多個規則來形成增益集的規則集合。下列範例會在 2 個規則上套用「AND」作業：**ItemIs** 規則和 **ItemHasKnownEntity** 規則。每當目前項目為郵件，且 Outlook 識別該項目的主旨或本文中的地址時，此規則集合會啟動增益集。




```XML
<Rule xsi:type="RuleCollection" Mode="And">
   <Rule xsi:type="ItemIs" ItemType="Message" />
   <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
</Rule>
```

下列範例會使用目前項目的 **getEntitiesByType** 將變數 `addresses` 設定為前述規則集合的結果。




```
var addresses = Office.context.mailbox.item.getEntitiesByType(Office.MailboxEnums.EntityType.Address);
```

每當目前項目的主旨或本文中有 URL，且 URL 包含字串「youtube」(無論字串的大小寫) 時，下列 **ItemHasKnownEntity** 規則範例會啟動增益集。




```XML
<Rule xsi:type="ItemHasKnownEntity" 
    EntityType="Url" 
    RegExFilter="youtube"
    FilterName="youtube"
    IgnoreCase="true"/>
```

下列範例會使用目前項目的 **getFilteredEntitiesByName(name)** 來設定變數 `videos` 以取得與前述 **ItemHasKnownEntity** 規則中的規則運算式相符的結果陣列。




```
var videos = Office.context.mailbox.item.getFilteredEntitiesByName(youtube);
```


## 使用已知實體的秘訣


如果您在增益集中使用已知的實體，則應該要注意幾個事實和限制。只要使用者在讀取的項目中包含已知實體的符合項目，無論您是否使用 **ItemHasKnownEntity** 規則，便啟用增益集，套用下列項目︰


1. 僅在字串是英文時，您才可以擷取已知實體的字串。
    
2. 您可以從項目本文的前 2,000 個字元擷取已知的實體，但超出該限制則不可。此大小限制有助於平衡功能和效能的需求，因此不會藉由剖析及識別大型郵件和約會中的已知實體的執行個體而拖累 Exchange Server 和 Outlook。請注意，這項限制與增益集是否指定 **ItemHasKnownEntity** 規則無關。如果增益集確實使用這種規則，則也請注意以下項目 2 中規則處理限制的 Outlook 豐富型用戶端。
    
3. 您可以從任何人員而非信箱擁有者所召集的會議的約會來擷取實體。您無法從非會議或由信箱擁有者召集的會議的行事曆項目來擷取實體。
    
4. 您僅可從郵件而非會議擷取 **MeetingSuggestion** 類型的實體。
    
5. 您可以擷取明確存在於項目本文中的 URL，而非內嵌在 HTML 項目本文中超連結文字的 URL。請考慮改用 **ItemHasRegularExpressionMatch** 規則以取得明確及內嵌的 URL。指定 **BodyAsHTML** 做為 _PropertyName_，及符合 URL 的規則運算式做為 _RegExValue_。
    
6. 您無法從 [寄件備份] 資料夾中的項目擷取實體。
    
此外，如果您使用 [ItemHasKnownEntity](http://msdn.microsoft.com/en-us/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx) 規則，且可能會影響您在其他方面期望您增益集啟動的案例，則套用下列項目︰


1. 當使用 **ItemHasKnownEntity** 規則時，預期 Outlook 僅符合英文的實體字串，無論資訊清單中指定的預設地區設定。
    
2. 當增益集在 Outlook 的豐富型用戶端上執行時，預期 Outlook 將 **ItemHasKnownEntity** 規則套用至項目本文的前 1 MB，而非超過該限制的本文其餘部分。
    
3. 您不能使用 **ItemHasKnownEntity** 規則來啟動 [寄件備份] 資料夾中項目的增益集。
    

## 其他資源



- [建立讀取格式的 Outlook 增益集](../outlook/read-scenario.md)
    
- [從 Outlook 項目擷取實體字串](../outlook/extract-entity-strings-from-an-item.md)
    
- [Outlook 增益集的啟用規則](../outlook/manifests/activation-rules.md)
    
- [使用規則運算式的啟用規則來顯示 Outlook 增益集](../outlook/use-regular-expressions-to-show-an-outlook-add-in.md)
    
- [了解 Outlook 增益集的權限](../outlook/understanding-outlook-add-in-permissions.md)
    
