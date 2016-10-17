
# <a name="get-attachments-of-an-outlook-item-from-the-server"></a>從伺服器取得 Outlook 項目的附件

Outlook 增益集無法將選取項目的附件直接傳遞至您的伺服器上執行的遠端服務。相反地，增益集可以使用附件 API 來傳送附件相關資訊到遠端服務。然後，服務可以直接連絡 Exchange Server 以擷取附件。

若要將附件資訊傳送至遠端服務，您可以使用下列屬性和函數︰


- 
  [Office.context.mailbox.ewsUrl](../../reference/outlook/Office.context.mailbox.md) 屬性 -- 在主控信箱的 Exchange Server 上提供 Exchange Web Services (EWS) 的 URL。您的服務會使用這個 URL 來呼叫 [ExchangeService.GetAttachments](http://msdn.microsoft.com/en-us/library/office/dn600509%28v=exchg.80%29.aspx)[EWS 受管理 API](http://msdn.microsoft.com/library/c2267733-6f4f-49e5-9614-1e4a24c3af1a%28Office.15%29.aspx) 方法或 [GetAttachment](http://msdn.microsoft.com/en-us/library/24d10a15-b942-415e-9024-a6375708f326%28Office.15%29.aspx) EWS 作業。
    
- [Office.context.mailbox.item.attachments](../../reference/outlook/Office.context.mailbox.item.md) 屬性 -- 取得 [AttachmentDetails](../../reference/outlook/simple-types.md) 物件的陣列，項目的每個附件使用一個。
    
- [Office.context.mailbox.getCallbackTokenAsync](../../reference/outlook/Office.context.mailbox.md) 函數 -- 對主控信箱的 Exchange Server 進行非同步呼叫，以取得伺服器傳回 Exchange Server 以驗證附件要求的回撥權杖。
    

## <a name="using-the-attachments-api"></a>使用附件 API


若要使用附件 API 以從 Exchange 信箱取得附件，請執行下列步驟︰ 


1. 當使用者正在檢視包含附件的郵件或約會時顯示增益集。
    
2. 從 Exchange 伺服器取得回呼權杖。
    
3. 將回呼權杖和附件資訊傳送到遠端服務。
    
4. 使用 **ExchangeService.GetAttachments** 方法或 **GetAttachment** 作業從 Exchange Server 取得附件。
    
這些步驟中每個都在下列各節中使用 [Outlook-Add-in-JavaScript-GetAttachments](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-GetAttachments) 範例中的程式碼詳細討論。


 >**附註** 已縮短這些範例中的程式碼，以強調附件資訊。這個範例包含額外的程式碼，以利用遠端伺服器驗證增益集並管理要求的狀態。


### <a name="activate-the-add-in"></a>啟動增益集


當選取的項目具有附件時，您可以使用增益集資訊清單檔中的 [ItemHasAttachment](http://msdn.microsoft.com/en-us/library/031db7be-8a25-5185-a9c3-93987e10c6c2%28Office.15%29.aspx) 規則來顯示您的增益集，如下列範例中所示。


```XML
<Rule xsi:type="ItemHasAttachment" />
```


### <a name="get-a-callback-token"></a>取得回撥權杖


[Office.context.mailbox](../../reference/outlook/Office.context.mailbox.md) 物件提供 **getCallbackTokenAsync** 函數來取得遠端伺服器可用來以 Exchange Server 驗證的權杖。下列程式碼會顯示增益集中的函數，其可啟動非同步要求來取得回撥權杖及取得回應的回撥函數。回撥權杖會儲存在下一節中所定義的服務要求物件。


```
function getAttachmentToken() {
    if (serviceRequest.attachmentToken == "" {
        Office.context.mailbox.getCallbackTokenAsync(attachmentTokenCallback);
    }
};
function attachmentTokenCallback(asyncResult, userContext) {
    if (asyncResult.status === "succeeded") {
        // Cache the result from the server.
        serviceRequest.attachmentToken = asyncResult.value;
        serviceRequest.state = 3;
        testAttachments();
    } else {
        showToast("Error", "Could not get callback token: " + asyncResult.error.message);
    }
};
```


### <a name="send-attachment-information-to-the-remote-service"></a>將附件資訊傳送到遠端服務


您的增益集呼叫的遠端服務會定義如何將附件資訊傳送至服務的細節。在這個範例中，遠端服務是使用 Visual Studio 2013 所建立的 Web API 應用程式。遠端服務會預期 JSON 物件中的附件資訊。下列程式碼會初始化包含附件資訊的物件。


```
// Initialize a context object for the add-in.
//   Set the fields that are used on the request
//   object to default values.
serviceRequest = new Object();
serviceRequest.attachmentToken = "";
serviceRequest.ewsUrl = Office.context.mailbox.ewsUrl;
serviceRequest.attachments = new Array();
```

`Office.context.mailbox.item.attachments` 屬性包含一系列的 **AttachmentDetails** 物件，項目的每個附件使用一個。在大部分的情況下，增益集只能傳遞 **AttachmentDetails** 物件的附件 ID 屬性至遠端服務。如果遠端服務需要更多關於附件的詳細資料，您可以傳遞全部或部分的 **AttachmentDetails** 物件。下列程式碼會定義方法，將整個 **AttachmentDetails** 陣列放置在 `serviceRequest` 物件中，並傳送要求到遠端服務。




```js
    function makeServiceRequest() {
      // Format the attachment details for sending.
      for (var i = 0; i < mailbox.item.attachments.length; i++) {
        serviceRequest.attachments[i] = JSON.parse(JSON.stringify(mailbox.item.attachments[i].$0_0));
      }

      $.ajax({
        url: '../../api/Default',
        type: 'POST',
        data: JSON.stringify(serviceRequest),
        contentType: 'application/json;charset=utf-8'
      }).done(function (response) {
        if (!response.isError) {
          var names = "<h2>Attachments processed using " +
                        serviceRequest.service +
                        ": " +
                        response.attachmentsProcessed +
                        "</h2>";
          for (i = 0; i < response.attachmentNames.length; i++) {
            names += response.attachmentNames[i] + "<br />";
          }
          document.getElementById("names").innerHTML = names;
        } else {
          app.showNotification("Runtime error", response.message);
        }
      }).fail(function (status) {

      }).always(function () {
        $('.disable-while-sending').prop('disabled', false);
      })
    };

```


### <a name="get-the-attachments-from-the-exchange-server"></a>從 Exchange Server 伺服器取得附件


您的遠端服務可以使用 [GetAttachments](http://msdn.microsoft.com/en-us/library/office/dn600509%28v=exchg.80%29.aspx) EWS 受管理 API 方法或 [GetAttachment](http://msdn.microsoft.com/library/24d10a15-b942-415e-9024-a6375708f326%28Office.15%29.aspx) EWS 作業，從伺服器擷取附件。服務應用程式需要兩個物件以將 JSON 字串還原序列化成可在伺服器上使用的 .NET Framework 物件。下列程式碼顯示還原序列化物件的定義。


```C#



namespace AttachmentsSample
{
  public class AttachmentSampleServiceRequest
  {
    public string attachmentToken { get; set; }
    public string ewsUrl { get; set; }
    public string service { get; set; }
    public AttachmentDetails [] attachments { get; set; }
  }

  public class AttachmentDetails
  {
    public string attachmentType { get; set; }
    public string contentType { get; set; }
    public string id { get; set; }
    public bool isInline { get; set; }
    public string name { get; set; }
    public int size { get; set; }
  }
}
```


#### <a name="use-the-ews-managed-api-to-get-the-attachments"></a>使用 EWS 受管理 API 取得附件

如果您在在遠端服務中使用 [EWS 受管理 API](http://go.microsoft.com/fwlink/?LinkID=255472)，您可以使用 [GetAttachments](http://msdn.microsoft.com/en-us/library/office/dn600509%28v=exchg.80%29.aspx) 方法，以建構、傳送及接收 EWS SOAP 要求來取得附件。我們建議您使用 EWS 受管理 API，因為它需要較少行的程式碼，且提供更具直覺性介面來呼叫 EWS。下列程式碼會提出一個要求來擷取所有的附件，並傳回已處理附件的計數和名稱。


```C#
    private AttachmentSampleServiceResponse GetAtttachmentsFromExchangeServerUsingEWSManagedApi(AttachmentSampleServiceRequest request)
    {
      var attachmentsProcessedCount = 0;
      var attachmentNames = new List<string>();

      // Create an ExchangeService object, set the credentials and the EWS URL.
      ExchangeService service = new ExchangeService();
      service.Credentials = new OAuthCredentials(request.attachmentToken);
      service.Url = new Uri(request.ewsUrl);

      var attachmentIds = new List<string>();

      foreach (AttachmentDetails attachment in request.attachments)
      {
        attachmentIds.Add(attachment.id);
      }

      // Call the GetAttachments method to retrieve the attachments on the message.
      // This method results in a GetAttachments EWS SOAP request and response
      // from the Exchange server.
      var getAttachmentsResponse =
        service.GetAttachments(attachmentIds.ToArray(),
                               null,
                               new PropertySet(BasePropertySet.FirstClassProperties,
                                               ItemSchema.MimeContent));

      if (getAttachmentsResponse.OverallResult == ServiceResult.Success)
      {
        foreach (var attachmentResponse in getAttachmentsResponse)
        {
          attachmentNames.Add(attachmentResponse.Attachment.Name);

          // Write the content of each attachment to a stream.
          if (attachmentResponse.Attachment is FileAttachment)
          {
            FileAttachment fileAttachment = attachmentResponse.Attachment as FileAttachment;
            Stream s = new MemoryStream(fileAttachment.Content);
            // Process the contents of the attachment here.
          }

          if (attachmentResponse.Attachment is ItemAttachment)
          {
            ItemAttachment itemAttachment = attachmentResponse.Attachment as ItemAttachment;
            Stream s = new MemoryStream(itemAttachment.Item.MimeContent.Content);
            // Process the contents of the attachment here.
          }

          attachmentsProcessedCount++;
        }
      }

      // Return the names and number of attachments processed for display
      // in the add-in UI.
      var response = new AttachmentSampleServiceResponse();
      response.attachmentNames = attachmentNames.ToArray();
      response.attachmentsProcessed = attachmentsProcessedCount;

      return response;
    }


```


#### <a name="use-ews-to-get-the-attachments"></a>使用 EWS 取得附件

如果您在遠端服務中使用 EWS，您必須建構 [GetAttachment](http://msdn.microsoft.com/library/24d10a15-b942-415e-9024-a6375708f326%28Office.15%29.aspx) SOAP 要求以從 Exchange Server 取得附件。下列程式碼會傳回提供 SOAP 要求的字串。遠端服務使用 **String.Format** 方法將附件的附件 ID 插入字串。


```C#
    private const string GetAttachmentSoapRequest =
@"<?xml version=""1.0"" encoding=""utf-8""?>
<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""
xmlns:xsd=""http://www.w3.org/2001/XMLSchema""
xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/""
xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types"">
<soap:Header>
<t:RequestServerVersion Version=""Exchange2013"" />
</soap:Header>
  <soap:Body>
    <GetAttachment xmlns=""http://schemas.microsoft.com/exchange/services/2006/messages""
    xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types"">
      <AttachmentShape/>
      <AttachmentIds>
        <t:AttachmentId Id=""{0}""/>
      </AttachmentIds>
    </GetAttachment>
  </soap:Body>
</soap:Envelope>";

```

最後，下列方法會執行使用 EWS  **GetAttachment** 要求從 Exchange Server 取得附件的工作。這項實作會針對每個附件提出個別要求，並傳回已處理附件的計數。個別的 **ProcessXmlResponse** 方法中處理每個回應，接下來會定義。




```C#
    private AttachmentSampleServiceResponse GetAttachmentsFromExchangeServerUsingEWS(AttachmentSampleServiceRequest request)
    {
      var attachmentsProcessedCount = 0;
      var attachmentNames = new List<string>();

      foreach (var attachment in request.attachments)
      {
        // Prepare a web request object.
        HttpWebRequest webRequest = WebRequest.CreateHttp(request.ewsUrl);
        webRequest.Headers.Add("Authorization",
          string.Format("Bearer {0}", request.attachmentToken));
        webRequest.PreAuthenticate = true;
        webRequest.AllowAutoRedirect = false;
        webRequest.Method = "POST";
        webRequest.ContentType = "text/xml; charset=utf-8";

        // Construct the SOAP message for the GetAttachment operation.
        byte[] bodyBytes = Encoding.UTF8.GetBytes(
          string.Format(GetAttachmentSoapRequest, attachment.id));
        webRequest.ContentLength = bodyBytes.Length;

        Stream requestStream = webRequest.GetRequestStream();
        requestStream.Write(bodyBytes, 0, bodyBytes.Length);
        requestStream.Close();

        // Make the request to the Exchange server and get the response.
        HttpWebResponse webResponse = (HttpWebResponse)webRequest.GetResponse();

        // If the response is okay, create an XML document from the reponse
        // and process the request.
        if (webResponse.StatusCode == HttpStatusCode.OK)
        {
          var responseStream = webResponse.GetResponseStream();

          var responseEnvelope = XElement.Load(responseStream);

          // After creating a memory stream containing the contents of the 
          // attachment, this method writes the XML document to the trace output.
          // Your service would perform it's processing here.
          if (responseEnvelope != null)
          {
            var processResult = ProcessXmlResponse(responseEnvelope);
            attachmentNames.Add(string.Format("{0} {1}", attachment.name, processResult));

          }

          // Close the response stream.
          responseStream.Close();
          webResponse.Close();

        }
        // If the response is not OK, return an error message for the 
        // attachment.
        else
        {
          var errorString = string.Format("Attachment \"{0}\" could not be processed. " +
            "Error message: {1}.", attachment.name, webResponse.StatusDescription);
          attachmentNames.Add(errorString);
        }
        attachmentsProcessedCount++;
      }

      // Return the names and number of attachments processed for display
      // in the add-in UI.
      var response = new AttachmentSampleServiceResponse();
      response.attachmentNames = attachmentNames.ToArray();
      response.attachmentsProcessed = attachmentsProcessedCount;

      return response;
    }

```

來自 **GetAttachment** 作業的每個回應會傳送至 **ProcessXmlResponse** 方法。這個方法會檢查回應是否有錯誤。如果沒有找到任何錯誤，它會處理檔案附件和項目附件。**ProcessXmlResponse** 方法會執行大量工作來處理附件。




```C#
    // This method processes the response from the Exchange server.
    // In your application the bulk of the processing occurs here.
    private string ProcessXmlResponse(XElement responseEnvelope)
    {
      // First, check the response for web service errors.
      var errorCodes = from errorCode in responseEnvelope.Descendants
                       ("{http://schemas.microsoft.com/exchange/services/2006/messages}ResponseCode")
                       select errorCode;
      // Return the first error code found.
      foreach (var errorCode in errorCodes)
      {
        if (errorCode.Value != "NoError")
        {
          return string.Format("Could not process result. Error: {0}", errorCode.Value);
        }
      }

      // No errors found, proceed with processing the content.
      // First, get and process file attachments.
      var fileAttachments = from fileAttachment in responseEnvelope.Descendants
                        ("{http://schemas.microsoft.com/exchange/services/2006/types}FileAttachment")
                            select fileAttachment;
      foreach(var fileAttachment in fileAttachments)
      {
        var fileContent = fileAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}Content");
        var fileData = System.Convert.FromBase64String(fileContent.Value);
        var s = new MemoryStream(fileData);
        // Process the file attachment here. 
      }

      // Second, get and process item attachments.
      var itemAttachments = from itemAttachment in responseEnvelope.Descendants
                            ("{http://schemas.microsoft.com/exchange/services/2006/types}ItemAttachment")
                            select itemAttachment;
      foreach(var itemAttachment in itemAttachments)
      {
        var message = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}Message");
        if (message != null)
        {
         // Process a message here.
          break;
        }
        var calendarItem = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}CalendarItem");
        if (calendarItem != null)
        {
          // Process calendar item here.
          break;
        }
        var contact = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}Contact");
        if (contact != null)
        {
          // Process contact here.
          break;
        }
        var task = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}Tontact");
        if (task != null)
        {
          // Process task here.
          break;
        }
        var meetingMessage = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}MeetingMessage");
        if (meetingMessage != null)
        {
          // Process meeting message here.
          break;
        }
        var meetingRequest = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}MeetingRequest");
        if (meetingRequest != null)
        {
          // Process meeting request here.
          break;
        }
        var meetingResponse = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}MeetingResponse");
        if (meetingResponse != null)
        {
          // Process meeting response here.
          break;
        }
        var meetingCancellation = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}MeetingCancellation");
        if (meetingCancellation != null)
        {
          // Process meeting cancellation here.
          break;
        }
      }
     
      return string.Empty;
    }

```


## <a name="additional-resources"></a>其他資源



- [建立讀取格式的 Outlook 增益集](../outlook/read-scenario.md)
    
- 
  [探索 Exchange 中的 EWS Managed API、EWS 和 Web 服務](http://msdn.microsoft.com/library/0bc6f81d-cc10-42b0-ba5d-6f22ff55d51c%28Office.15%29.aspx)
    
- 
  [開始使用 EWS Managed API 用戶端應用程式](http://msdn.microsoft.com/library/c2267733-6f4f-49e5-9614-1e4a24c3af1a%28Office.15%29.aspx)
    
- [Outlook-Power-Hour_Code-Samples](https://github.com/OfficeDev/Outlook-Power-Hour-Code-Samples)：`MyAttachments` 和 `AttachmentsDemo`
    
