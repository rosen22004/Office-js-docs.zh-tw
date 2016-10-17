
# <a name="inside-the-exchange-identity-token"></a>Exchange 識別權杖的內容
了解 Exchange 2013 識別權杖中所包含內容。



Exchange Server 傳送至 Outlook 增益集的驗證識別權杖對您的增益集而言是不透明的；您不需要查看權杖內部即可在您的伺服器上傳送它。但是，當您在撰寫與 Outlook 增益集互動的 web 服務程式碼時，您需要知道識別權杖的內容。

## <a name="what-is-an-identity-token?"></a>識別權杖是什麼？


識別權杖是 Base-64 URL 編碼的字串，由傳送它的 Exchange Server 自我簽署。權杖未加密，您用來驗證簽章的公開金鑰會儲存在發出此權杖的 Exchange Server 上。權杖包含三個部分︰標頭、內容和簽章。在權杖字串中，各部分是由「.」字元所分隔，讓您可輕鬆分割權杖。

Exchange 2013 針對識別權杖使用 JSON Web 權杖 (JWT)。如需 JWT 權杖的相關資訊，請參閱 [JSON Web 權杖 (JWT) 網際網路草稿](http://self-issued.info/docs/draft-goland-json-web-token-00.html).


### <a name="identity-token-header"></a>識別權杖標頭

標頭會識別權杖，並可讓您的 web 服務知道要呈現的是哪一種權杖。下列範例會顯示識別權杖的標頭的外觀。

```js
{ "typ" : "JWT", "alg" : "RS256", "x5t" : "Un6V7lYN-rMgaCoFSTO5z707X-4" }
```

下表描述識別權杖標頭的部分。


**識別權杖標頭的部分**


|**宣告**|**值**|**描述**|
|:-----|:-----|:-----|
|typ|"JWT"|將權杖識別為 JSON Web 權杖。Exchange Server 提供的所有識別權杖皆為 JWT 權杖。|
|alg|"RS256"|用來建立簽章的雜湊演算法。Exchange Server 提供的所有權杖使用 RS-256 演算法。|
|x5t|憑證指紋|權杖的 X.509 指紋。|

### <a name="identity-token-payload"></a>識別權杖內容

內容包含識別電子郵件帳戶及識別傳送權杖的 Exchange Server 的驗證宣告。下列範例會顯示內容區段的外觀。
```js

{ 
   "aud" : "https://mailhost.contoso.com/IdentityTest.html", 
   "iss" : "00000002-0000-0ff1-ce00-000000000000@mailhost.contoso.com", 
   "nbf" : "1331579055", 
   "exp" : "1331607855", 
   "appctxsender":"00000002-0000-0ff1-ce00-000000000000@mailhost.context.com",
   "isbrowserhostedapp":"true",
"appctx" : { 
     "msexchuid" : "53e925fa-76ba-45e1-be0f-4ef08b59d389@mailhost.contoso.com" "version" : "ExIdTok.V1" "amurl" :         "https://mailhost.contoso.com:443/autodiscover/metadata/json/1" 
     } 
}
```
下表列出識別權杖內容的部分。


**識別權杖內容的部分**


|**宣告**|**描述**|
|:-----|:-----|
|aud|要求權杖的增益集 URL。權杖僅在從用戶端瀏覽器中執行的增益集傳送時才有效。如果增益集使用 Office 增益集資訊清單結構描述 1.1 版，則此 URL 為 **ItemRead** 或 **ItemEdit**表單類型下的第一個 **SourceLocation** 元素中指定的 URL，視何者先發生於增益集資訊清單中的 [FormSettings](http://msdn.microsoft.com/en-us/library/0d1a311d-939d-78c1-e968-89ddf7ebc4b4%28Office.15%29.aspx) 元素部分。|
|iss|發出此權杖的 Exchange Server 的唯一識別碼。此 Exchange Server 所發出的所有權杖都會有相同的識別碼。|
|nbf|有效的權杖開始的日期和時間。值為 1970 年 1 月 1 日以來的秒數。 |
|exp|有效的權杖到期的日期和時間。值為 1970 年 1 月 1 日以來的秒數。|
|appctxsender|傳送應用程式內容的 Exchange Server 的唯一識別碼。|
|isbrowserhostedapp|表示增益集是否在瀏覽器中主控。|
|appctx|權杖的應用程式內容。 |
appctx 宣告中的資訊為您提供電子郵件帳戶的位址，以及唯一的識別碼。下表列出 appctx 宣告的部分。



|**appctx 宣告組件**|**描述**|
|:-----|:-----|
|msexchuid|與電子郵件帳戶和 Exchange Server 相關聯的唯一識別碼。|
|版本|權杖的版本號碼。針對執行 Exchange 2013 的伺服器所提供的所有權杖，值為「ExIdTok.V1」。|
|amurl|包含用來簽署權杖的 X.509 憑證的公開金鑰的驗證中繼資料文件的 URL。如需有關如何使用驗證中繼資料文件詳細資訊，請參閱[驗證 Exchange 識別權杖](../outlook/validate-an-identity-token.md)。|

### <a name="identity-token-signature"></a>識別權杖簽章

簽章的建立方法為雜湊標頭和內容區段與標頭中指定的演算法，並使用位於內容中指定位置的伺服器上的自我簽署 X509 憑證。您的 web 服務可驗證此簽章，以協助確保識別權杖是來自您希望傳送的伺服器。


## <a name="additional-resources"></a>其他資源



- [使用 Exchange 識別權杖來驗證 Outlook 增益集](../outlook/authentication.md)
    
- [在 Exchange 中使用識別權杖以從 Outlook 增益集呼叫服務](../outlook/call-a-service-by-using-an-identity-token.md)
    
- [使用 Exchange 權杖驗證程式庫](../outlook/use-the-token-validation-library.md)
    
- [驗證 Exchange 識別權杖](../outlook/validate-an-identity-token.md)
    
