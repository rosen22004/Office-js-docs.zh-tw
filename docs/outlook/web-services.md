
# 從 Outlook 增益集呼叫 Web 服務

增益集可以從正在執行 Exchange Server 2013 的電腦使用 Exchange Web 服務 (EWS)，其為伺服器上可供使用的 web 服務 (該伺服器提供增益集 UI 的來源位置)，或在網際網路上可供使用的 web 服務。本文提供的範例會顯示如何透過 Outlook 增益集會如何從 EWS 要求資訊。

您呼叫 web 服務的方式會根據 web 服務所在的位置而有所不同。表 1 列出您可以根據位置呼叫 web 服務的不同方法。


**表 1.從 Outlook 增益集呼叫 Web 服務的方法**


|**Web 服務位置**|**要呼叫 web 服務的方法**|
|:-----|:-----|
|裝載用戶端信箱的 Exchange Server|使用 [mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md) 方法來呼叫增益集所支援的 EWS 作業。裝載用戶端信箱的 Exchange Server 也會公開 EWS。|
|提供增益集 UI 來源位置的 web 伺服器|使用標準 JavaScript 技術來呼叫 web 服務。UI 框架中的 JavaScript 程式碼在提供 UI 的 web 伺服器的內容中執行。因此，它可以呼叫該伺服器上的 web 服務，而不會引起跨網站指令碼的錯誤。|
|所有其他位置|提供 UI 來源位置的 web 伺服器上建立 web 服務的 proxy。如果您不提供 proxy，跨網站指令碼錯誤將會防止增益集執行。提供 proxy 的方法之一是使用 JSON/P。如需詳細資訊，請參閱 [Office 增益集的隱私權和安全性](../../docs/develop/privacy-and-security.md)。|

## 使用 makeEwsRequestAsync 方法存取 EWS 作業


您可以使用 [mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md) 方法對主控使用者信箱的 Exchange Server 進行 EWS 要求。

EWS 在 Exchange server 上支援不同的作業；例如複製、尋找、更新或傳送項目的項目層級作業，以及建立、取得或更新資料夾的資料夾層級作業。若要執行 EWS 作業，針對該作業建立 XML SOAP 要求。當作業完成時，會取得包含與作業相關的資料的 XML SOAP 回應。EWS SOAP 要求和回應遵循 Messages.xsd 檔案中定義的結構描述。如其他的 EWS 結構描述檔案，Message.xsd 檔案位於裝載 EWS 的 IIS 虛擬目錄中。 

若要使用 **makeEwsRequestAsync** 方法來起始 EWS 作業，請提供下列項目︰


- 該 EWS 作業的 SOAP 要求的 XML，做為 _data_ 參數的引數。
    
- 回呼方法 (做為 _callback_ 引數)
    
- 該回呼方法的任何選擇性輸入資料 (做為 _userContext_ 引數)
    
當 EWS SOAP 要求完成時，Outlook 會使用一個引數呼叫回呼方法，該引數為 [AsyncResult](../../reference/outlook/simple-types.md) 物件。回呼方法可以存取 **AsyncResult** 物件的兩個屬性：**value** 屬性，其包含 EWS 作業的 XML SOAP 回應，及 (選擇性地) **asyncContext** 屬性，其包含以 **userContext** 參數傳遞的任何資料。一般而言，回呼方法會剖析 SOAP 回應中的 XML，以取得任何相關的資訊，然後據以處理該資訊。


## 剖析 EWS 回應的秘訣


從 EWS 作業剖析 SOAP 回應時，請注意下列的瀏覽器相關問題︰


- 使用 DOM 方法 **getElementsByTagName** 時，指定標記名稱的前置詞，以包含 Internet Explorer 的支援。
    
     **getElementsByTagName** 的效果會依瀏覽器類型而有所不同。 例如，EWS 回應可以包含下列 XML (為清楚顯示，採用格式化及縮寫)：
    
```XML
      <t:ExtendedProperty><t:ExtendedFieldURI PropertySetId="00000000-0000-0000-0000-000000000000" 
    PropertyName="MyProperty" 
    PropertyType="String"/>
    <t:Value>{
    ...
    }</t:Value></t:ExtendedProperty>
```

 如下列所示的程式碼可在如 Chrome 等瀏覽器上執行，以取得以 **ExtendedProperty** 標記括住的 XML：

```js
    var mailbox = Office.context.mailbox;
    mailbox.makeEwsRequestAsync(mailbox.item.itemId), function(result) {
        var response = $.parseXML(result.value);
        var extendedProps = response.getElementsByTagName("ExtendedProperty");
```


   
 在 Internet Explorer 上，您必須包含標記名稱的 `t:` 前置詞，如下所示︰

```js
    var mailbox = Office.context.mailbox;
    mailbox.makeEwsRequestAsync(mailbox.item.itemId), function(result) {
        var response = $.parseXML(result.value);
        var extendedProps = response.getElementsByTagName("t:ExtendedProperty");
```

- 使用 DOM 屬性 **textContent** 在 EWS 回應中取得標記的內容，如下所示︰
    
```
      content = $.parseJSON(value.textContent);
```

 其他屬性 (例如 **innerHTML**) 在 Internet Explorer 上可能不適用於 EWS 回應中的某些標記。
    

## 範例


下列範例會呼叫 **makeEwsRequestAsync** 以使用 [GetItem](http://msdn.microsoft.com/en-us/library/e3590b8b-c2a7-4dad-a014-6360197b68e4%28Office.15%29.aspx) 作業來取得項目主旨。此範例包含下列三個功能︰


-  `getSubjectRequest` -- 取項目 ID 做為輸入，並傳回 SOAP 要求的 XML 來呼叫 **GetItem** 以取得指定的項目。
    
-  `sendRequest` -- 呼叫 `getSubjectRequest` 以取得選取項目的 SOAP 要求，然後將 SOAP 要求和回呼方法 `callback`傳遞至 **makeEwsRequestAsync** 以取得指定項目的主旨。
    
-  `callback` -- 處理 SOAP 回應，其包含任何主旨和關於指定項目的其他資訊。
    

```js
function getSubjectRequest(id) {
   // Return a GetItem operation request for the subject of the specified item. 
   var result = 
'<?xml version="1.0" encoding="utf-8"?>' +
'<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
'               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
'               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
'               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
'  <soap:Header>' +
'    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
'  </soap:Header>' +
'  <soap:Body>' +
'    <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
'      <ItemShape>' +
'        <t:BaseShape>IdOnly</t:BaseShape>' +
'        <t:AdditionalProperties>' +
'            <t:FieldURI FieldURI="item:Subject"/>' +
'        </t:AdditionalProperties>' +
'      </ItemShape>' +
'      <ItemIds><t:ItemId Id="' + id + '"/></ItemIds>' +
'    </GetItem>' +
'  </soap:Body>' +
'</soap:Envelope>';

   return result;
}





function sendRequest() {
   // Create a local variable that contains the mailbox.
   var mailbox = Office.context.mailbox;

   mailbox.makeEwsRequestAsync(getSubjectRequest(mailbox.item.itemId), callback);
}

function callback(asyncResult)  {
   var result = asyncResult.value;
   var context = asyncResult.context;

   // Process the returned response here.
}


```


## 增益集支援的 EWS 作業


Outlook 增益集可以透過 **makeEwsRequestAsync** 方法存取 EWS 中可用的作業子集。如果您不熟悉 EWS 作業，以及如何使用 **makeEwsRequestAsync** 方法來存取作業，請以 SOAP 要求範例開始自訂 _data_ 引數。以下說明如何使用 **makeEwsRequestAsync** 方法︰


1. 在 XML 中，以適當的值取代任何項目 ID 和相關的 EWS 作業屬性。
    
2. 包含 SOAP 要求做為 _makeEwsRequestAsync_ 的 **data** 參數的引數。
    
3. 指定回呼方法及呼叫 **makeEwsRequestAsync**。
    
4. 在回呼方法中，確認在 SOAP 回應中作業的結果。
    
5. 根據您的需要使用 EWS 作業的結果。
    
下表列出的增益集所支援的 EWS 作業。若要參閱 SOAP 要求和回應的範例，請選擇每個作業的連結。如需有關 EWS 作業的詳細資訊，請參閱 [Exchange 中的 EWS 作業](http://msdn.microsoft.com/library/cf6fd871-9a65-4f34-8557-c8c71dd7ce09%28Office.15%29.aspx)。


**表 2.支援的 EWS 作業**


|**EWS 作業**|**說明**|
|:-----|:-----|
|[CopyItem 作業](http://msdn.microsoft.com/library/bcc68f9e-d511-4c29-bba6-ed535524624a%28Office.15%29.aspx)|複製指定的項目，並將新的項目放在 Exchange 儲存區的指定資料夾中。|
|[CreateFolder 作業](http://msdn.microsoft.com/library/6f6c334c-b190-4e55-8f0a-38f2a018d1b3%28Office.15%29.aspx)|在 Exchange 儲存區的指定位置中建立資料夾。|
|[CreateItem 作業](http://msdn.microsoft.com/library/78a52120-f1d0-4ed7-8748-436e554f75b6%28Office.15%29.aspx)|在 Exchange 儲存區中建立指定項目。|
|[FindConversation 作業](http://msdn.microsoft.com/library/2384908a-c203-45b6-98aa-efd6a4c23aac%28Office.15%29.aspx)|在 Exchange 儲存區的指定資料夾中列舉一個會談清單。|
|[FindFolder 作業](http://msdn.microsoft.com/library/7a9855aa-06cc-45ba-ad2a-645c15b7d031%28Office.15%29.aspx)|尋找已識別資料夾的子資料夾，並傳回一組描述子資料夾組的屬性。|
|[FindItem 作業](http://msdn.microsoft.com/library/ebad6aae-16e7-44de-ae63-a95b24539729%28Office.15%29.aspx)|識別位於 Exchange 儲存區的指定資料夾中的項目。|
|[GetConversationItems 作業](http://msdn.microsoft.com/library/8ae00a99-b37b-4194-829c-fe300db6ab99%28Office.15%29.aspx)|取得一或多組在對話中組織的節點的項目。|
|[GetFolder 作業](http://msdn.microsoft.com/library/355bcf93-dc71-4493-b177-622afac5fdb9%28Office.15%29.aspx)|從 Exchange 儲存區取得資料夾的指定屬性和內容。|
|[GetItem 作業](http://msdn.microsoft.com/library/e3590b8b-c2a7-4dad-a014-6360197b68e4%28Office.15%29.aspx)|從 Exchange 儲存區取得項目的指定屬性和內容。|
|[MarkAsJunk 作業](http://msdn.microsoft.com/library/1f71f04d-56a9-4fee-a4e7-d1034438329e%28Office.15%29.aspx)|將電子郵件移至 [垃圾郵件] 資料夾，並據以從封鎖的寄件者清單中新增或移除郵件的寄件者。|
|[MoveItem 作業](http://msdn.microsoft.com/library/dcf40fa7-7796-4a5c-bf5b-7a509a18d208%28Office.15%29.aspx)|在 Exchange 儲存區中將項目移到單一目的地的資料夾中。|
|[SendItem 作業](http://msdn.microsoft.com/library/337b89ef-e1b7-45ed-92f3-8abe4200e4c7%28Office.15%29.aspx)|傳送位於 Exchange 儲存區中的電子郵件。|
|[UpdateFolder 作業](http://msdn.microsoft.com/library/3494c996-b834-4813-b1ca-d99642d8b4e7%28Office.15%29.aspx)|在 Exchange 儲存區中修改現有資料夾的屬性。|
|[UpdateItem 作業](http://msdn.microsoft.com/library/5d027523-e0bc-4da2-b60b-0cb9fc1fdfe4%28Office.15%29.aspx)|在 Exchange 儲存區中修改現有項目的屬性。|

## makeEwsRequestAsync 方法的驗證和權限考量


當您使用 **makeEwsRequestAsync** 方法時，會使用目前使用者的電子郵件帳戶的認證來驗證要求。**makeEwsRequestAsync** 方法會管理您的認證，讓您不需要使用您的要求提供驗證認證。


 >
  **附註**  伺服器系統管理員必須使用 [New-WebServicesVirtualDirctory](http://technet.microsoft.com/en-us/library/bb125176.aspx) 或 [Set-WebServicesVirtualDirecory](http://technet.microsoft.com/en-us/library/aa997233.aspx) cmldet 將 Client Access server EWS 目錄上的 _OAuthAuthentication_ 參數設定為 **true**，以便啟用 **makeEwsRequestAsync** 方法來進行 EWS 要求。

增益集必須指定其增益集資訊清單中的 **ReadWriteMailbox** 權限以使用 **makeEwsRequestAsync** 方法。如需如何使用 **ReadWriteMailbox** 權限的相關資訊，請參閱 [了解 Outlook 增益集權限](../outlook/understanding-outlook-add-in-permissions.md#readwritemailbox-permission) 中的 [ReadWriteMailbox 權限](../outlook/understanding-outlook-add-in-permissions.md)一節。


## 其他資源



- [Outlook 增益集](../outlook/outlook-add-ins.md)
    
- [Office 增益集的隱私權和安全性](../../docs/develop/privacy-and-security.md)
    
- [解決 Office 增益集中的相同原始來源原則的限制](../../docs/develop/addressing-same-origin-policy-limitations.md)
    
- [Exchange EWS 參考](http://msdn.microsoft.com/library/2a873474-1bb2-4cb1-a556-40e8c4159f4a%28Office.15%29.aspx)
    
- [Exchange 中的 Outlook 和 EWS 郵件應用程式](http://msdn.microsoft.com/library/821c8eb9-bb58-42e8-9a3a-61ca635cba59%28Office.15%29.aspx)
    
請參閱下列資訊以了解使用 ASP.NET Web API 建立增益集的後端服務：


- [使用 ASP.NET Web API 建立 Office 增益集的 web 服務](http://blogs.msdn.com/b/officeapps/archive/2013/06/10/create-a-web-service-for-an-app-for-office-using-the-asp-net-web-api.aspx)
    
- [使用 ASP.NET Web API 建置 HTTP 服務的基本概念](http://www.asp.net/web-api)
    
