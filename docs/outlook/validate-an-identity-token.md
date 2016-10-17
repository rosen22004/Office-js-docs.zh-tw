
# <a name="validate-an-exchange-identity-token"></a>驗證 Exchange 識別權杖

Outlook 增益集可以傳送識別權杖給您，但您必須驗證權杖以確保其來自您所預期的 Exchange server 之後，才可以信任要求。本文中的範例會告訴您如何使用以 C#; 撰寫的驗證物件來驗證 Exchange 識別權杖；不過，您可以使用任何程式設計語言來執行驗證。[JSON Web 權杖 (JWT) 網際網路草稿](http://self-issued.info/docs/draft-goland-json-web-token-00.mdl)中敘述驗證權杖所需的步驟。 

我們建議您使用四個步驟的程序來驗證識別權杖，並取得使用者的唯一識別項。首先，從以 Base64 編碼的字串擷取 JSON Web 權杖 (JWT)。第二，請確定權杖格式正確，其適用於您的 Outlook 增益集、尚未過期，而且您可以擷取驗證中繼資料文件的有效 URL。接下來，從 Exchange Server 擷取驗證中繼資料文件，並附加至識別權杖的簽章。最後，利用驗證中繼資料文件的 URL 雜湊使用者的 Exchange ID，計算使用者的唯一識別項。整體而言，程序看起來會很複雜，但每個個別的步驟相當簡單。您可以在 [Outlook-Add-in-JavaScript-ValidateIdentityToken](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ValidateIdentityToken)下載包含這些來自 web 的範例的解決方案。
 




## <a name="set-up-to-validate-your-identity-token"></a>設定以驗證識別權杖


本文中的程式碼範例取決於 Windows Identity Foundation (WIF)，以及利用 JSON 權杖的處理常式延伸 WIF 的 DLL。您可以從下列位置下載必要的組件︰


- [Windows Identity Foundation](http://msdn.microsoft.com/en-us/security/aa570351)
    
- [32 位元應用程式的 Windows.IdentityModel.Extensions.dll](http://download.microsoft.com/download/0/1/D/01D06854-CA0C-46F1-ADBA-EBF86010DCC6/MicrosoftIdentityExtensions-32.msi)
    
- [64 位元應用程式的 Windows.IdentityModel.Extensions.dll](http://download.microsoft.com/download/0/1/D/01D06854-CA0C-46F1-ADBA-EBF86010DCC6/MicrosoftIdentityExtensions-64.msi)
    

## <a name="extract-the-json-web-token"></a>擷取 JSON Web 權杖


**Decode** 原廠方法會從 Exchange Server 將 JWT 分割成三個組成權杖的字串，然後使用 **Base64Decode** 方法 (如第二個範例所示) 將 JWT 標頭和內容解碼成 JSON 字串。字串會傳遞至 **JsonToken** 建構函式，其中會驗證 JWT 的內容且會傳回新的 **JsonToken** 物件執行個體。


```C#
    public static JsonToken Decode(string rawToken)
    {
      string[] tokenParts = rawToken.Split('.');

      if (tokenParts.Length != 3)
      {
        throw new ApplicationException("Token must have three parts separated by '.' characters.");
      }

      string encodedHeader = tokenParts[0];
      string encodedPayload = tokenParts[1];
      string signature = tokenParts[2];

      string decodedHeader = Base64Decode(encodedHeader);
      string decodedPayload = Base64Decode(encodedPayload);

      JavaScriptSerializer serializer = new JavaScriptSerializer();

      Dictionary<string, string> header = serializer.Deserialize<Dictionary<string, string>>(decodedHeader);
      Dictionary<string, string> payload = serializer.Deserialize<Dictionary<string, string>>(decodedPayload);

      return new JsonToken(header, payload, signature);
    }
```

**Base64Decode** 方法會實作 [JSON Web 權杖 (JWT) 網際網路草稿](http://self-issued.info/docs/draft-goland-json-web-token-00.mdl)中的「無填補實作 base64url 編碼的附註」附錄中所述的解碼邏輯。




```C#
    public static Encoding TextEncoding = Encoding.UTF8;

    private static char Base64PadCharacter = '=';
    private static char Base64Character62 = '+';
    private static char Base64Character63 = '/';
    private static char Base64UrlCharacter62 = '-';
    private static char Base64UrlCharacter63 = '_';

    private static byte[] DecodeBytes(string arg)
    {
      if (String.IsNullOrEmpty(arg))
      {
        throw new ApplicationException("String to decode cannot be null or empty.");
      }

      StringBuilder s = new StringBuilder(arg);
      s.Replace(Base64UrlCharacter62, Base64Character62);
      s.Replace(Base64UrlCharacter63, Base64Character63);

      int pad = s.Length % 4;
      s.Append(Base64PadCharacter, (pad == 0) ? 0 : 4 - pad);

      return Convert.FromBase64String(s.ToString());
    }

    private static string Base64Decode(string arg)
    {
      return TextEncoding.GetString(DecodeBytes(arg));
    }
```


## <a name="parse-the-jwt"></a>剖析 JWT


**JsonToken** 物件的建構函式會檢查 JWT 的結構與內容以判斷是否有效。在您要求驗證中繼資料文件之前，最好這麼做。如果 JWT 不包含適當的宣告，或已超過其存留時間，您可以避免呼叫 Exchange Server 與相關聯的延遲。

建構函式會呼叫公用程式方法，判斷是否出現不同的宣告且宣告在範圍中。如果沒有問題，公用程式方法會擲回應用程式例外狀況。如果沒有擲回任何例外狀況，**IsValid** 屬性會設定為 **true** 且權杖會準備好簽章驗證。

每個公用程式方法會在本文稍後進一步說明。




```C#
    public JsonToken(Dictionary<string, string> header, Dictionary<string, string> payload, string signature)
    {

      // Assume that the token is invalid to start out.
      this.IsValid = false;

      // Set the private dictionaries that contain the claims.
      this.headerClaims = header;
      this.payloadClaims = payload;
      this.signature = signature;

      // If there is no "appctx" claim in the token, throw an ApplicationException.
      if (!this.payloadClaims.ContainsKey(AuthClaimTypes.AppContext))
      {
        throw new ApplicationException(String.Format("The {0} claim is not present.", AuthClaimTypes.AppContext));
      }

      appContext = new JavaScriptSerializer().Deserialize<Dictionary<string, string>>(payload[AuthClaimTypes.AppContext]);


      // Validate the header fields.
      this.ValidateHeader();

      // Determine whether the token is within its valid time.
      this.ValidateLifetime();

      // Validate that the token was sent to the correct URL.
      this.ValidateAudience();

      // Validate the token version.
      this.ValidateVersion();

      // Make sure that the appctx contains an authentication
      // metadata location.
      this.ValidateMetadataLocation();

      // If the token passes all the validation checks, we
      // can assume that it is valid.
      this.IsValid = true;
    }
```


### <a name="validateheader-method"></a>ValidateHeader 方法

**ValidateHeader** 方法會檢查以確定必要的宣告是否位於權杖的標頭，以及宣告具有正確的值。標頭必須設定如下；否則，方法會擲回應用程式例外狀況並結束。

```js
{ "typ" : "JWT", "alg" : "RS256", "x5t" : "<thumbprint>" }
```

```C#
    private void ValidateHeaderClaim(string key, string value)
    {
      if (!this.headerClaims.ContainsKey(key))
      {
        throw new ApplicationException(String.Format("Header does not contain \"{0}\" claim.", key));
      }

      if (!value.Equals(this.headerClaims[key]))
      {
        throw new ApplicationException(String.Format("\"{0}\" claim must be \"{0}\".", key, value));
      }
    }

    private void ValidateHeader()
    {
      ValidateHeaderClaim(AuthClaimTypes.TokenType, Config.TokenType);
      ValidateHeaderClaim(AuthClaimTypes.Algorithm, Config.Algorithm);
    
      if (!this.headerClaims.ContainsKey(AuthClaimTypes.x509Thumprint))
      {
        throw new ApplicationException(String.Format("Header does not contain \"{0}\" claim.", AuthClaimTypes.x509Thumprint));
      }
    }


```


### <a name="validatelifetime-method"></a>ValidateLifetime 方法

JWT 中會提供兩個日期：「nbf」(not before) 提供權杖變成有效的日期和時間，以及「exp」提供權杖到期的時間。只有出現在這兩個日期之間的權杖才可視為有效。若要調整伺服器和用戶端之間的時鐘設定中的微小差異，這個方法會在權杖的時間中最多 5 分鐘前後內驗證權杖。


```C#
    private void ValidateLifetime()
    {
      if (!this.payloadClaims.ContainsKey(AuthClaimTypes.ValidFrom))
      {
        throw new ApplicationException(
          String.Format("The \"{0}\" claim is missing from the token.", AuthClaimTypes.ValidFrom));
      }

      if (!this.payloadClaims.ContainsKey(AuthClaimTypes.ValidTo))
      {
        throw new ApplicationException(
          String.Format("The \"{0}\" claim is missing from the token.", AuthClaimTypes.ValidTo));
      }

      DateTime unixEpoch = new DateTime(1970, 1, 1, 0, 0, 0,DateTimeKind.Utc);

      TimeSpan padding = new TimeSpan(0, 5, 0);

      DateTime validFrom = unixEpoch.AddSeconds(int.Parse(this.payloadClaims[AuthClaimTypes.ValidFrom]));
      DateTime validTo = unixEpoch.AddSeconds(int.Parse(this.payloadClaims[AuthClaimTypes.ValidTo]));

      DateTime now = DateTime.UtcNow;

      if (now < (validFrom - padding))
      {
        throw new ApplicationException(String.Format("The token is not valid until {0}.", validFrom));
      }

      if (now > (validTo + padding))
      {
        throw new ApplicationException(String.Format("The token is not valid after {0}.", validFrom));
      }
    }
```

**validFrom** (nbf) 及 **validTo** (exp) 日期會以自 1970 年 1 月 1 日 Unix 期間起的秒數傳送。使用 UTC 計算日期和時間以防止任何在 Exchange Server 與執行驗證程式碼的伺服器之間的時區差異問題。


### <a name="validateaudience-method"></a>ValidateAudience 方法

識別權杖僅對要求它的增益集有效。**ValidateAudience** 方法會檢查權杖中的對象宣告，以確保它符合 Outlook 增益集的預期 URL。


```C#
    private void ValidateAudience()
    {
      if (!this.payloadClaims.ContainsKey(AuthClaimTypes.Audience))
      {
        throw new ApplicationException(String.Format("The \"{0}\" claim is missing from the application context.", AuthClaimTypes.Audience));
      }

      string location = Config.Audience.Replace("/", "-").Replace("\\", "-");
      string audience = this.payloadClaims[AuthClaimTypes.Audience].Replace("/", "-").Replace("\\", "-");

      if (!location.Equals(audience))
      {
        throw new ApplicationException(String.Format(
          "The audience URL does not match. Expected {0}; got {1}.",
          Config.Audience, this.payloadClaims[AuthClaimTypes.Audience]));
      }
    }

```


### <a name="validateversion-method"></a>ValidateVersion 方法

**ValidateVersion** 方法會檢查識別權杖的版本，並確保其符合預期的版本。不同版本的權杖可以執行不同的宣告。檢查版本可確保預期的宣告會是識別權杖。


```js
    private void ValidateVersion()
    {
      if (!this.appContext.ContainsKey(AuthClaimTypes.MsExchExtensionVersion))
      {
        throw new ApplicationException(String.Format("The \"{0}\" claim is missing from the token.", AuthClaimTypes.MsExchExtensionVersion));
      }

      if (!Config.Version.Equals(this.appContext[AuthClaimTypes.MsExchExtensionVersion]))
      {
        throw new ApplicationException(String.Format(
          "The version does not match. Expected {0}; got {1}.",
          Config.Version, this.appContext[AuthClaimTypes.MsExchExtensionVersion]));
      }
    }

```


### <a name="validatemetadatalocation-method"></a>ValidateMetadataLocation 方法

儲存在 Exchange Server 上的驗證中繼資料物件包含驗證識別權杖中包含的簽章所需的資訊。**ValidateMetadataLocation** 方法會確保識別權杖中有驗證中繼資料 URL 宣告，實際驗證簽章會在下一個步驟中發生。


```C#
    private void ValidateMetadataLocation()
    {
      if (!this.appContext.ContainsKey(AuthClaimTypes.MsExchAuthMetadataUrl))
      {
        throw new ApplicationException(String.Format("The \"{0}\" claim is missing from the token.", AuthClaimTypes.MsExchAuthMetadataUrl));
      }
    }

```


## <a name="validate-the-identity-token-signature"></a>驗證識別權杖簽章


在知道 JWT 包含驗證簽章所需的宣告之後，您可以使用 Windows Identity Foundation (WIF) 和 WIF 副檔名來驗證權杖上的簽章。您需要下列資訊以驗證簽章︰


- 傳送自 Exchange Server 的原始 Base-64 URL 編碼的驗證簽章字串。
    
- 來自 JWT 的驗證中繼資料文件位置。
    
- 來自 JWT 的對象 URL。
    
在本範例中，**IdentityToken** 物件的建構函式會從 Exchange Server 取得驗證中繼資料文件，並驗證識別權杖的簽章。如果識別權杖是有效的，您可以使用 **IdentityToken** 物件執行個體取得識別權杖中包含的唯一使用者識別碼。




```C#
    public IdentityToken(string rawToken, string audience, string authMetadataEndpoint)
    {
      X509Certificate2 currentCertificate = null;

      currentCertificate = AuthMetadata.GetSigningCertificate(new Uri(authMetadataEndpoint));

      JsonWebSecurityTokenHandler jsonTokenHandler =
          GetSecurityTokenHandler(audience, authMetadataEndpoint, currentCertificate);

      SecurityToken jsonToken = jsonTokenHandler.ReadToken(rawToken);
      JsonWebSecurityToken webToken = (JsonWebSecurityToken)jsonToken;

      SigningCertificateThumbprint = currentCertificate.Thumbprint;
      Issuer = webToken.Issuer;
      Audience = webToken.Audience;
      ValidTo = webToken.ValidTo;
      ValidFrom = webToken.ValidFrom;
      foreach (JsonWebTokenClaim claim in webToken.Claims)
      {
        if (claim.ClaimType.Equals(AuthClaimTypes.AppContextSender))
        {
          ApplicationContextSender = claim.Value;
        }

        if (claim.ClaimType.Equals(AuthClaimTypes.IsBrowserHostedApp))
        {
          IsBrowserHostedApp = claim.Value == "true";
        }

        if (claim.ClaimType.Equals(AuthClaimTypes.AppContext))
        {
          string[] appContextClaims = claim.Value.Split(',');
          Dictionary<string, string> appContext =
              new JavaScriptSerializer().Deserialize<Dictionary<string, string>>(claim.Value);
          AuthenticationMetaDataUrl = appContext[AuthClaimTypes.MsExchAuthMetadataUrl];
          ExchangeID = appContext[AuthClaimTypes.MsExchImmutableId];
          TokenVersion = appContext[AuthClaimTypes.MsExchTokenVersion];
        }
      }
    }


```

**IdentityToken** 物件建構函式中大部分的程式碼會利用來自 Exchange Server 的宣告在執行個體上設定屬性。建構函式會呼叫 **GetSecurityTokenHandler** 方法來取得驗證 Exchange 識別權杖的權杖處理常式。**GetSecurityTokenHandler** 方法會呼叫兩個公用程式方法，**GetMetadataDocument** 和 **GetSigningCertificate**，它會執行從 Exchange Server 取得簽章憑證的工作。下列各節將描述這些方法。


### <a name="getsecuritytokenhandler-method"></a>GetSecurityTokenHandler 方法

**GetSecurityTokenHandler** 方法會傳回驗證識別權杖的 WIF 權杖處理常式。方法中大多數的程式碼會初始化權杖處理常式來進行驗證；但是，該方法會呼叫 **GetSigningCertificate** 方法以擷取用來簽署來自 Exchange Server 的權杖的 X.509 憑證。


```C#
    private JsonWebSecurityTokenHandler GetSecurityTokenHandler(string audience,
        string authMetadataEndpoint,
        X509Certificate2 currentCertificate)
    {
      JsonWebSecurityTokenHandler jsonTokenHandler = new JsonWebSecurityTokenHandler();
      jsonTokenHandler.Configuration = new SecurityTokenHandlerConfiguration();

      jsonTokenHandler.Configuration.AudienceRestriction = new AudienceRestriction(AudienceUriMode.Always);
      jsonTokenHandler.Configuration.AudienceRestriction.AllowedAudienceUris.Add(
        new Uri(audience, UriKind.RelativeOrAbsolute));

      jsonTokenHandler.Configuration.CertificateValidator = X509CertificateValidator.None;

      jsonTokenHandler.Configuration.IssuerTokenResolver =
        SecurityTokenResolver.CreateDefaultSecurityTokenResolver(
          new ReadOnlyCollection<SecurityToken>(new List<SecurityToken>(
            new SecurityToken[]
            {
              new X509SecurityToken(currentCertificate)
            })), false);

      ConfigurationBasedIssuerNameRegistry issuerNameRegistry = new ConfigurationBasedIssuerNameRegistry();
      issuerNameRegistry.AddTrustedIssuer(currentCertificate.Thumbprint, Config.ExchangeApplicationIdentifier);
      jsonTokenHandler.Configuration.IssuerNameRegistry = issuerNameRegistry;

      return jsonTokenHandler;
    }
```


### <a name="getsigningcertificate-method"></a>GetSigningCertificate 方法

**GetSigningCertificate** 方法會呼叫 **GetMetadataDocument** 方法以擷取來自 Exchange Server 的驗證中繼資料，然後傳回驗證中繼資料文件中的第一個 X.509 憑證。如果文件不存在，方法會擲回應用程式例外狀況。


```C#
    private X509Certificate2 GetSigningCertificate(Uri authMetadataEndpoint)
    {
      JsonAuthMetadataDocument document = GetMetadataDocument(authMetadataEndpoint);

      if (null != document.keys &amp;&amp; document.keys.Length > 0)
      {
        JsonKey signingKey = document.keys[0];

        if (null != signingKey &amp;&amp; null != signingKey.keyValue)
        {
          return new X509Certificate2(Encoding.UTF8.GetBytes(signingKey.keyValue.value));
        }
      }

      throw new ApplicationException("The metadata document does not contain a signing certificate.");
    }

```


### <a name="getmetadatadocument-method"></a>GetMetadataDocument 方法

驗證中繼資料文件包含您要在 Exchange 識別權杖上驗證的簽章的資訊。文件會以 JSON 字串傳送。**GetMetatDataDocument** 方法會要求來自 Exchange 識別權杖中指定位置的文件，並傳回封裝 JSON 字串做為物件的物件。如果 URL 未包含驗證中繼資料文件，方法會擲回應用程式例外狀況。


```C#
    private JsonAuthMetadataDocument GetMetadataDocument(Uri authMetadataEndpoint)
    {
      // Uncomment the next line if your Exchange server uses the default
      // self-signed certificate.
      // ServicePointManager.ServerCertificateValidationCallback = Config.CertificateValidationCallback;

      byte[] acsMetadata;
      using (WebClient webClient = new WebClient())
      {
        acsMetadata = webClient.DownloadData(authMetadataEndpoint);
      }
      string jsonResponseString = Encoding.UTF8.GetString(acsMetadata);

      JsonAuthMetadataDocument document = new JavaScriptSerializer().Deserialize<JsonAuthMetadataDocument>(jsonResponseString);

      if (null == document)
      {
        throw new ApplicationException(String.Format("No authentication metadata document found at {0}.", authMetadataEndpoint));
      }

      return document;
    }
```

依預設，Exchange Server 會使用自我簽署的 X.509 憑證來驗證驗證中繼資料文件的要求。除非您安裝追蹤回到根伺服器的憑證，您必須建立憑證驗證回呼方法，否則驗證中繼資料文件的要求將會失敗。 

.NET Framework System.Net 命名空間中的 **ServicePointManager** 類別可讓您連接驗證回呼方法，方法為設定 **ServerCertificateValidationCallback** 屬性。在 [驗證 X509 憑證](http://msdn.microsoft.com/en-us/library/dd633677%28EXCHG.80%29.aspx)文章中，您可以看到適用於開發和測試的憑證驗證回呼方法的範例。


 **安全性附註**  如果您使用憑證驗證回呼方法，必須確定它符合組織的安全性需求。


## <a name="compute-the-unique-id-for-an-exchange-account"></a>計算 Exchange 帳戶的唯一識別碼


您可以雜湊驗證中繼資料文件 URL 與帳戶的 Exchange 識別項，來建立 Exchange 帳戶的唯一識別項。當您具有此唯一識別項時，可以使用它來建立 Outlook 增益集 web 服務的單一登入 (SSO) 系統。如需有關使用 SSO 的唯一識別項的詳細資訊，請參閱[使用 Exchange 的識別權杖來驗證使用者](../outlook/authenticate-a-user-with-an-identity-token.md)

**UniqueUserIdentification** 屬性會建立 Exchange ID 的改良式 SHA256 雜湊和驗證中繼資料 URL，方法為使用來自 **System.Security.Cryptography** 命名空間的標準 SHA256 提供者。


 **安全性附註**  您必須使用 Exchange ID 雜湊驗證中繼資料文件，以建立帳戶的唯一識別項。只使用 Exchange ID 可以公開您的服務給未經授權的使用者。一如往常，當處理驗證及安全性時，您必須確保使用以這個方法建立的唯一識別項符合您的應用程式的安全性需求。




```C#
    // Salt to apply when creating unique ID.
    private byte[] Salt = new byte[] {<Provide random salt bytes here };

    private string ComputeUniqueIdentification()
    {
      byte[] inputBytes = Encoding.ASCII.GetBytes(string.Concat(ExchangeID, AuthenticationMetaDataUrl));

      // Combine input bytes and salt.
      byte[] saltedInput = new byte[Salt.Length + inputBytes.Length];
      Salt.CopyTo(saltedInput, 0);
      inputBytes.CopyTo(saltedInput, Salt.Length);

      // Compute the unique key.
      byte[] hashedBytes = SHA256CryptoServiceProvider.Create().ComputeHash(saltedInput);

      // Convert the hashed value to a string and return.
      return BitConverter.ToString(hashedBytes);
    }

    public string UniqueUserIdentification
    {
      get { return ComputeUniqueIdentification(); }
    }


```


## <a name="utility-objects"></a>公用程式物件


本文中的程式碼範例根據幾個公用程式物件而定，這些物件對所使用的常數提供好記的名稱。下表列出公用程式物件。


**表 1.公用程式物件**


|**物件**|**描述**|
|:-----|:-----|
|**AuthClaimsType**|將權杖驗證程式碼所使用的宣告識別項收集在單一位置。|
|**Config**|提供要驗證識別權杖的常數。 |
|**JsonAuthMetadataDocument**|封裝從 Exchange Server 所傳送的 JSON 驗證中繼資料文件。|

### <a name="authclaimtypes-object"></a>AuthClaimTypes 物件

**AuthClaimTypes** 物件將權杖驗證程式碼所使用的宣告識別項收集在單一位置。它包括標準 JWT 宣告以及 Exchange 識別權杖中的特定宣告。


```C#
  public class AuthClaimTypes
  {
    public const string NameIdentifier =
        JsonWebTokenConstants.ReservedClaims.NameIdentifier;
    public const string MsExchImmutableId = "msexchuid";
    public const string MsExchTokenVersion = "version";
    public const string MsExchAuthMetadataUrl = "amurl";

    public const string AppContext =
        JsonWebTokenConstants.ReservedClaims.AppContext;
    public const string Audience =
        JsonWebTokenConstants.ReservedClaims.Audience;
    public const string Issuer =
        JsonWebTokenConstants.ReservedClaims.Issuer;
    public const string ValidFrom =
        JsonWebTokenConstants.ReservedClaims.NotBefore;
    public const string ValidTo =
        JsonWebTokenConstants.ReservedClaims.ExpiresOn;

    public const string AppContextSender = "appctxsender";
    public const string IsBrowserHostedApp = "isbrowserhostedapp";

    public const string TokenType = "typ";
    public const string Algorithm = "alg";
    public const string x509Thumbprint = "x5t";      
  }
```


### <a name="config-object"></a>Config 物件

**Config** 物件包含用來驗證識別權杖的常數，以及伺服器沒有可追蹤回到根憑證的 X509 憑證時，您可以使用的憑證驗證回呼方法。


 
  **安全性附註**  僅在您的伺服器使用預設的自我簽署的憑證時，才需要安全性憑證回呼方法。當憑證是自我簽署時，這個範例中的回呼方法會傳回 **false**，所以您必須以符合組織的安全性需求的回呼方法取代它。如需適用於開發和測試的憑證驗證回呼方法的範例，請參閱[驗證 X509 憑證](http://msdn.microsoft.com/en-us/library/dd633677%28EXCHG.80%29.aspx)文章。


```C#
  public static class Config
  {
    public static string Algorithm = "RS256";
    public static string Audience = @"https:\\localhost:44300\Pages\IdentityTest.html";
    public static string TokenType = "JWT";
    public static string Version = "ExIdTok.V1";

    public static string ExchangeApplicationIdentifier = "Exchange";

    internal static bool CertificateValidationCallback(
    object sender,
    System.Security.Cryptography.X509Certificates.X509Certificate certificate,
    System.Security.Cryptography.X509Certificates.X509Chain chain,
    System.Net.Security.SslPolicyErrors sslPolicyErrors)
    {
      // If the certificate is a valid, signed certificate, return true.
      if (sslPolicyErrors == System.Net.Security.SslPolicyErrors.None)
      {
        return true;
      }

      // If there are errors in the certificate chain, look at each error to determine the cause.
      else
      {
        return false;
      }
    }
  }
```


### <a name="jsonauthmetadatadocument-object"></a>JsonAuthMetadataDocument 物件

**JsonAuthMetadataDocument** 物件會透過屬性公開驗證中繼資料文件的內容。


```C#
using System;

namespace IdentityTest
{
  public class JsonAuthMetadataDocument
  {
    public string id { get; set; }
    public string version { get; set; }
    public string name { get; set; }
    public string realm { get; set; }
    public string serviceName { get; set; }
    public string issuer { get; set; }
    public string [] allowedAudiences { get; set; }
    public JsonKey[] keys;
    public JsonEndpoint[] endpoints;
  }

  public class JsonEndpoint
  {
    public string location { get; set; }
    public string protocol { get; set; }
    public string usage { get; set; }
  }

  public class JsonKey
  {
    public string usage { get; set; }
    public JsonKeyValue keyValue { get; set; }
  }

  public class JsonKeyValue
  {
    public string type { get; set; }
    public string value { get; set; }
  }
}

```


## <a name="additional-resources"></a>其他資源



- [使用 Exchange 識別權杖來驗證 Outlook 增益集](../outlook/authentication.md)
    
- [Exchange 識別權杖的內容](../outlook/inside-the-identity-token.md)
    
