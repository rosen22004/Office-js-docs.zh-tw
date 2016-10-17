
# <a name="limits-for-activation-and-javascript-api-for-outlook-add-ins"></a>適用於 Outlook 增益集的 JavaScript API 和啟動的限制

若要為 Outlook 增益集的使用者提供令人滿意的經驗，您應該留意特定啟動和 API 使用方針，並實作您的增益集以保持在這些限制內。這些指導方針存在，因此個別的增益集無法要求 Exchange Server 或 Outlook 花費過長的時間處理其啟動規則或呼叫適用於 Office 的 JavaScript API，影響整體 Outlook 和其他增益集的使用者體驗。這些限制會套用到增益集資訊清單中的設計啟動規則，及使用自訂屬性、漫遊設定、收件者、Exchange Web 服務 (EWS) 要求和回應，以及非同步呼叫。 

 >**附註** 如果您的增益集在 Outlook 豐富型用戶端上執行，您也必須確認增益集在執行階段資源使用狀況的限制內執行。 


## <a name="limits-for-activation-rules"></a>啟動規則的限制


設計 Outlook 增益集的啟動規則時，請遵循以下方針︰


- 將資訊清單的大小限制為 256 KB。如果您超過該上限，則無法安裝 Exchange 信箱的 Outlook 增益集。

- 針對增益集指定最多 15 個啟動規則。如果您超過該上限，則無法安裝增益集。
    
- 如果您在選取項目的本文使用 [ItemHasKnownEntity](http://msdn.microsoft.com/en-us/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx) 規則，則預期 Outlook 豐富型用戶端僅對本文的前 1 MB 套用規則，而超過該限制的非本文其餘部分。如果相符項目僅存在本文的前 1 MB 之後，則不會啟用增益集。如果您預期的是這種情況，請重新設計您的啟動條件。
    
- 如果您在 **ItemHasKnownEntity** 或 [ItemHasRegularExpressionMatch](http://msdn.microsoft.com/en-us/library/bfb726cd-81b0-a8d5-644f-2ca90a5273fc%28Office.15%29.aspx) 規則中使用規則運算式，請注意下列常套用於任何 Outlook 主應用程式的限制和指導方針，以及在表 1、2 和 3 中所述的會依主應用程式而有所不同︰
    
      - 在增益集的啟用規則中最多只能指定 5 個規則運算式。如果您超過該上限，則無法安裝增益集。
    
  - 指定規則運算式使得您預期的結果會由前 50 個相符項目內的 **getRegExMatches** 方法呼叫傳回。
    
  - 在規則運算式中指定前查詢的判斷提示，而非後查詢 (?<=text) 及負後查詢 (?<!text)。
    

表 1 列出限制，並說明在 Outlook 豐富型用戶端和 Outlook Web App 或裝置用 OWA 之間的規則運算式支援差異。支援與裝置及項目本文的任何特定類型無關。


 **表 1.在規則運算式的支援中的一般差異**


|**Outlook 豐富型用戶端**|**Outlook Web App 或裝置用 OWA**|
|:-----|:-----|
|使用提供做為 Visual Studio 標準範本程式庫一部分的 C++ 規則運算式引擎。這個引擎遵守 ECMAScript 5 標準。 |使用屬於 JavaScript 的規則運算式評估 (由瀏覽器提供)，且支援 ECMAScript 5 的超集。|
|因為不同的 regex 引擎，根據預先定義的字元類別預期包含自訂字元類別的 regex 在 Outlook 豐富型用戶端中可能會傳回與 Outlook Web App 或裝置用 OWA 不同的結果。<br/><br/>例如，regex "[\s\S]{0,100}" 符合空格或非空格的單一字元的任何數字 (介於 0 到 100)。這個 regex 在 Outlook 豐富型用戶端中會傳回與 Outlook Web App 和裝置用 OWA 不同的結果。您應該將 regex 重寫為 ""(\s\|\S){0,100}" 做為因應措施。此因應措施 regex 符合空格或非空格的單一字元的任何數字 (介於 0 到 100)。<br/><br/>您應該在每個 Outlook 主應用程式上測試徹底每個 regex，而如果 regex 傳回不同的結果，則重寫 regex。 |您應該在每個 Outlook 主應用程式上測試徹底每個 regex，而如果 regex 傳回不同的結果，則重寫 regex。|
|根據預設，會將增益集的所有規則運算式的評估限制為 1 秒。超過這個限制會導致最多 3 倍的重新評估。在重新評估外，Outlook 豐富型用戶端會停用增益集，使其無法在任何 Outlook 主應用程式中的相同信箱執行。<br/><br/>系統管理員可以使用 **OutlookActivationAlertThreshold** 及 **OutlookActivationManagerRetryLimit** 登錄機碼來覆寫這些評估限制。|不支援與 Outlook 的豐富型用戶端中相同的資源監視或登錄設定。但是，具有需要在 Outlook 豐富型用戶端上大量評估時間的規則運算式的增益集已針對所有 Outlook 主應用程式上相同的信箱停用。|

表 2 列出限制，並說明在每一個 Outlook 套用規則運算式的項目本文部分中的差異。這些限制部分會依裝置類型及項目本文而定，如果規則運算式套用在項目本文上。

**表 2.評估的項目本文的大小限制**


||**Outlook 豐富型用戶端**|**Outlook Web App、裝置用 OWA、OWA for iPad 或 OWA for iPhone**|**Outlook Web App**|
|:-----|:-----|:-----|:-----|
|表單的外觀尺寸|任何支援的裝置|Android 智慧型手機、iPad 或 iPhone|任何支援的裝置，除了 Android 智慧型手機、iPad 或 iPhone|
|純文字項目本文|在本文前 1 MB 的資料上套用 regex，而非超過該限制的其餘本文上。|僅在 < 16,000 個字元時啟動增益集。|僅在 < 500,000 個字元時啟動增益集。|
|HTML 項目本文|在本文前 512 KB 的資料上套用 regex，而非超過該限制的其餘本文上。(實際的字元數取決於編碼，其範圍可從每個字元 1 到 4 個位元組。)|在前 64,000 個字元 (包含 HTML 標記字元) 上套用 regex，而非超過該限制的其餘本文上。|僅在 < 500,000 個字元時啟動增益集。|

表 3 列出限制，並說明在每一個 Outlook 主應用程式評估規則運算式後傳回的相符項目中的差異。支援與裝置的任何特定類型無關，但如果在項目本文上套用規則運算式，可能會取決於項目本文的類型。

** 3.傳回的符合項目的限制**


||**Outlook 豐富型用戶端**|**Outlook Web App 或裝置用 OWA**|
|:-----|:-----|:-----|
|傳回的相符項目的順序|假設 **getRegExMatches** 在 Outlook 豐富型用戶端中傳回套用在相同項目上的相同規則運算式的相符項目與 Outlook Web App 或裝置用 OWA 不同。|假設 **getRegExMatches** 在 Outlook 豐富型用戶端中傳回的相符項目與 Outlook Web App 或裝置用 OWA 不同。|
|純文字項目本文|**getRegExMatches** 傳回最多 1,536 (1.5 KB) 個字元的任何符合項目，最多 50 個符合項目。<br/><br/>**附註**：**getRegExMatches** 在傳回的陣列中不會以任何特定的順序傳回符合項目。一般而言，假設相同項目上套用的相同規則運算式的 Outlook 豐富型用戶端中的相符項目的順序與 Outlook Web App 或裝置用 OWA 不同。|**getRegExMatches** 傳回最多 3,072 (3 KB) 個字元的任何符合項目，最多 50 個符合項目。|
|HTML 項目本文|**getRegExMatches** 傳回最多 3,072 (3 KB) 個字元的任何符合項目，最多 50 個符合項目。<br/> <br/> **附註**：**getRegExMatches** 在傳回的陣列中不會以任何特定的順序傳回符合項目。一般而言，假設相同項目上套用的相同規則運算式的 Outlook 豐富型用戶端中的相符項目的順序與 Outlook Web App 或裝置用 OWA 不同。|**getRegExMatches** 傳回最多 3,072 (3 KB) 個字元的任何符合項目，最多 50 個符合項目。|

## <a name="limits-for-javascript-api"></a>JavaScript API 的限制


除了啟用規則的前述指導原則，每個 Outlook 主應用程式會在 JavaScript 物件模型中強制某些限制，如表 4 中所述。


**表 4.使用適用於 Office 的 JavaScript API 取得或設定某些資料的限制**


|**功能**|**限制**|**相關的 API**|**描述**|
|:-----|:-----|:-----|:-----|
|自訂屬性|2,500 個字元|[CustomProperties](../../reference/outlook/CustomProperties.md) 物件<br/> <br/>[item.loadCustomPropertiesAsync](../../reference/outlook/Office.context.mailbox.item.md) 方法|約會或郵件項目的所有自訂屬性的限制。如果增益集的所有自訂屬性的總大小超過這個限制，所有的 Outlook 主應用程式會傳回錯誤。|
|漫遊設定|32 KB 字元數|[RoamingSettings](../../reference/outlook/RoamingSettings.md) 物件<br/><br/> [context.roamingSettings](../../reference/outlook/Office.context.md) 屬性|增益集的所有漫遊設定的限制。如果您的設定超過此限制，所有的 Outlook 主應用程式會傳回錯誤。|
|解壓縮已知的實體|2,000 字元數|[item.getEntities](../../reference/outlook/Office.context.mailbox.item.md) 方法<br/> <br/>[item.getEntitiesByType](../../reference/outlook/Office.context.mailbox.item.md) 方法<br/> <br/>[item.getFilteredEntitiesByName](../../reference/outlook/Office.context.mailbox.item.md) 方法|Exchange Server 在項目本文上擷取已知實體的限制。Exchange Server 會忽略超出該限制的實體。請注意，這項限制與增益集是否使用 **ItemHasKnownEntity** 規則無關。|
|Exchange Web 服務|1 MB 字元數|[mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md) 方法|要求或回應 **Mailbox.makeEwsRequestAsync** 呼叫的限制。|
|收件者|100 位收件者|[item.requiredAttendees](../../reference/outlook/Office.context.mailbox.item.md) 屬性<br/> <br/>[item.optionalAttendees](../../reference/outlook/Office.context.mailbox.item.md) 屬性<br/> <br/>[item.resources](../../reference/outlook/Office.context.mailbox.item.md) 屬性<br/> <br/>[item.to](../../reference/outlook/Office.context.mailbox.item.md) 屬性<br/> <br/>[item.cc](../../reference/outlook/Office.context.mailbox.item.md) 屬性<br/> <br/>[Recipients.addAsync](../../reference/outlook/Recipients.md) 方法<br/> <br/>[Recipient.getAsync](../../reference/outlook/Recipients.md) 方法<br/> <br/>[Recipient.setAsync](../../reference/outlook/Recipients.md) 方法|在每個屬性中指定的收件者的限制。|
|顯示名稱|255 個字元|[EmailAddressDetails.displayName](../../reference/outlook/simple-types.md) 屬性<br/><br/> [收件者](../../reference/outlook/Recipients.md) 物件<br/><br/> **item.requiredAttendees** 屬性<br/><br/> **item.optionalAttendees** 屬性 <br/><br/>**item.resources** 屬性 <br/><br/>**item.to** 屬性 <br/><br/>**item.cc** 屬性|約會或郵件中的顯示名稱的長度限制。|
|設定主旨|255 個字元|[mailbox.displayNewAppointmentForm](../../reference/outlook/Office.context.mailbox.md) 方法<br/><br/> [Subject.setAsync](../../reference/outlook/Subject.md) 方法|新約會表單中的主旨或設定約會或郵件的主旨的限制。|
|設定位置|255 個字元|[Location.setAsync](../../reference/outlook/Location.md) 方法|設定約會的地點或會議邀請的限制。|
|新約會表單中的本文|32 KB 字元數|**Mailbox.displayNewAppointmentForm** 方法|新約會表單中的本文的限制。|
|顯示現有項目的本文|32 KB 字元數|[mailbox.displayAppointmentForm](../../reference/outlook/Office.context.mailbox.md) 方法<br/><br/> [mailbox.displayMessageForm](../../reference/outlook/Office.context.mailbox.md) 方法|針對 Outlook Web App 及裝置用 OWA：現有約會或郵件表單中的本文的限制。|
|設定本文|1 MB 字元數|[Body.prependAsync](../../reference/outlook/Body.md) 方法<br/> <br/>[Body.setAsync](../../reference/outlook/Body.md)<br/><br/>[Body.setSelectedDataAsync](../../reference/outlook/Body.md) 方法|設定約會或郵件項目本文的限制。|
|附件數|Outlook Web App 及裝置用 OWA 上 499 個檔案|[item.addFileAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md) 方法|可以附加至項目以傳送的檔案數的限制。Outlook Web App 及裝置用 OWA 通常會限制透過使用者介面和 **addFileAttachmentAsync** 附加最多 499 個檔案。Outlook 豐富型用戶端不會特別限制檔案附件的數目。不過，所有的 Outlook 主應用程式會觀察使用者的 Exchange Server 已設定的附件大小的限制。請參閱下一個資料列查看「附件的大小」。|
|附件的大小|取決於 Exchange Server|**item.addFileAttachmentAsync** 方法|項目的所有附件的大小有所限制，系統管理員可以在使用者信箱的 Exchange Server 上設定。針對 Outlook 豐富型用戶端，這會限制項目的附件數目。針對 Outlook Web App 及裝置用 OWA，兩個限制中較小的 - 附件數與所有附件的大小 - 會限制項目的實際附件。|
|附件檔名|255 個字元|**item.addFileAttachmentAsync** 方法|要加入項目的附件的檔案名稱長度限制。|
|附件 URI|2048 個字元|**item.addFileAttachmentAsync** 方法|要以附件方式加入項目的檔名 URI 的限制。|
|附件 ID|100 個字元|[item.addItemAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md) 方法<br/><br/> [item.removeAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md) 方法|要加入項目或從項目中移除的附件的 ID 長度限制。|
|非同步呼叫|3 個呼叫|**item.addFileAttachmentAsync** 方法<br/><br/>**item.addItemAttachmentAsync** 方法<br/><br/><br/>**item.removeAttachmentAsync** 方法<br/><br/> [Body.getTypeAsync](../../reference/outlook/Body.md) 方法<br/><br/>**Body.prependAsync** 方法<br/><br/>**Body.setSelectedDataAsync** 方法<br/><br/> [CustomProperties.saveAsync](../../reference/outlook/CustomProperties.md) 方法<br/><br/><br/> [item.LoadCustomPropertiesAysnc](../../reference/outlook/Office.context.mailbox.item.md) 方法<br/><br/><br/> [Location.getAsync](../../reference/outlook/Location.md) 方法<br/><br/>**Location.setAsync** 方法<br/><br/> [mailbox.getCallbackTokenAsync](../../reference/outlook/Office.context.mailbox.md) 方法<br/><br/> [mailbox.getUserIdentityTokenAsync](../../reference/outlook/Office.context.mailbox.md) 方法<br/><br/> [mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md) 方法<br/><br/>**Recipients.addAsync** 方法<br/><br/> [Recipients.getAsync](../../reference/outlook/Recipients.md) 方法<br/><br/>**Recipients.setAsync** 方法<br/><br/> [RoamingSettings.saveAsync](../../reference/outlook/RoamingSettings.md) 方法<br/><br/> [Subject.getAsync](../../reference/outlook/Subject.md) 方法<br/><br/>**Subject.setAsync** 方法<br/><br/> [Time.getAsync](../../reference/outlook/Time.md) 方法<br/><br/> [Time.setAsync](../../reference/outlook/Time.md) 方法|針對 Outlook Web App 或裝置用 OWA：任一時間同時非同步呼叫數目的限制，因為瀏覽器僅允許有限的對伺服器非同步呼叫數目。 |

## <a name="additional-resources"></a>其他資源



- [部署和安裝 Outlook 增益集以進行測試](../outlook/testing-and-tips.md)
    
- [Outlook 增益集的隱私權、權限和安全性](../outlook/../../docs/develop/privacy-and-security.md)
    
