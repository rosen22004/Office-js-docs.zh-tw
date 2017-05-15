# <a name="create-an-aspnet-office-add-in-that-uses-single-sign-on-preview"></a>建立使用單一登入的 ASP.NET Office 增益集 (預覽)

可以登入 Office 與 Office Web 增益集的使用者，可以利用這個登入程序為增益集和 Microsoft Graph 的使用者授權，而不需要使用者登入第二次。如需概觀，請參閱[在 Office 增益集啟用 SSO](../../docs/develop/sso-in-office-add-ins.md)。

本文會引導您完成啟用在增益集單一登入 (SSO) 的程序，該增益集是使用 ASP.NET、OWIN，以及 Microsoft 驗證程式庫 (MSAL) 所建立。 

> **附註：**如需有關以 Node.js 為基礎之增益集的類似文章，請參閱[建立使用單一登入的 Node.js Office 增益集](../../docs/develop/create-sso-office-add-ins-nodejs.md)。

## <a name="prerequisites"></a>必要條件

* Visual Studio 2017 15.3 版 (26424.2-Preview) 或更新版本。

* Office 2016 1704 版、組建 8027.nnnn 或更新版本 (Office 365 訂閱版本，有時候稱為「隨選即用」)。您必須是 Office 測試人員才能取得這個版本。如需詳細資訊，請參閱[成為 Office 測試人員](https://products.office.com/en-us/office-insider?tab=tab-1)。

## <a name="set-up-the-starter-project"></a>設定起始專案

1. 複製或下載位於 [Office 增益集 ASPNET SSO](https://github.com/officedev/office-add-in-aspnet-sso) 的存放庫。 

1. 開啟 [Before]**** 資料夾，並且開啟 Visual Studio 中的 .sln 檔案。這是起始專案。未直接連接至 SSO 或授權的 UI 和增益集的其他方面已完成。 

    > 附註：相同的存放庫中也有已完成版本的範例。如果您完成這篇文章的程序，它就如同您擁有的增益集，例外的是已完成專案具有程式碼註解，在這篇文章中可能是多餘的文字。若要使用已完成的版本，只要開啟 .sln 檔案，並且遵循本文中的指示，但是略過**撰寫用戶端**和**撰寫伺服器端**區段。

1. 專案開啟之後，在 Visual Studio 中建置它，這樣會讓 Visual Studio 安裝 packages.config 檔案中列出的套件。這個作業可能會耗費數秒鐘到數分鐘的時間，取決於電腦的本機套件快取中有多少套件。

1. 完全建置專案之後，按下 F5 鍵。PowerPoint 隨即開啟，且在 [首頁]**** 功能區中有 **SSO ASP.NET** 群組。 

1. 按下此群組中的 [顯示增益集]**** 按鈕，以在工作窗格中查看增益集的 UI。工作窗格中的按鈕尚未連線。 
2. 在 Visual Studio 中停止偵錯工具。

## <a name="register-the-add-in-with-azure-ad-v20-endpoint"></a>使用 Azure AD v2.0 端點註冊增益集

1. 瀏覽至 [https://apps.dev.microsoft.com/?test=build2017](https://apps.dev.microsoft.com/?test=build2017)。 

1. 使用系統管理員認證登入您的 Office 365 租用。例如，MyName@contoso.onmicrosoft.com

1. 按一下 [新增應用程式]****。

1. 出現提示時，使用「Office-Add-in-ASPNET-SSO」作為應用程式名稱，然後按一下 [建立應用程式]****。

1. 當應用程式的組態頁面開啟時，複製**應用程式 ID** 並且儲存。您將在稍後的程序中使用它。 

    > 附註：當其他應用程式，例如 Office 主應用程式 (例如 PowerPoint、Word、Excel) 尋求應用程式的授權存取時，這個 ID 是「對象」值。當應用程式轉為尋求 Microsoft Graph 的授權存取時，它也是應用程式的「用戶端 ID」。

1. 在 [應用程式祕密]**** 區段中，按下 [產生新的密碼]****。快顯對話方塊隨即開啟，並且顯示新的密碼 (也稱為「應用程式祕密」)。立即複製密碼，並且將它與應用程式 ID 一起儲存。**您在稍後的程序中需要它。然後關閉對話方塊。

1. 在 [平台]**** 區段中，按一下 [新增平台]****。 

1. 在開啟的對話方塊中，選取 [Web API]****。

1. 已產生表單的**應用程式 ID URI**，“api://{App ID GUID}”。以 “localhost:44355” 取代 GUID。整個 ID 應該會是 `api://localhost:44355`。(**應用程式 ID URI** 的**範圍**名稱網域部分將會自動變更為相符項目。它應該會是 `api://localhost:44355/access_as_user`。)

1. 在 [預先授權應用程式]**** 區段中，有空白的 [應用程式 ID]**** 方塊。在方塊中輸入下列 ID (這是 Microsoft Office 的 ID)：`d3590ed6-52b3-4102-aeff-aad2292ab01c`。

1. 開啟 [應用程式 ID]**** 旁的 [範圍]**** 下拉式清單，並且勾選 `api://localhost:44355/access_as_user` 的方塊。

1. 在 [平台]**** 區段的上方附近，再次按一下 [新增平台]****，並且選取 [Web]****。

1. 在 [平台]****底下新的 [Web]**** 區段中，輸入下列項目作為**重新導向 URL**：`https://localhost:44355`。 

    > 附註：本文撰寫時，**Web API** 平台有時候會從 [平台]**** 區段中消失，特別是當頁面在 **Web** 平台已新增且註冊網頁已儲存之後重新整理時**。為了再次確保您的 **Web API** 平台仍然是註冊的一部分，按一下頁面底端附近的 [編輯應用程式資訊清單]**** 按鈕。您應該會在資訊清單的 **identifierUris** 屬性中看到 `api://localhost:44355` 字串。也會有 **oauth2Permissions** 屬性，它的**值**子屬性具有值 `access_as_user`。

1. 向下捲動到 [Microsoft Graph 權限]**** 區段、[委派的權限]**** 子區段。使用 [新增]**** 按鈕以開啟 [選取權限]**** 對話方塊。

1. 在對話方塊中，勾選下列權限的方塊 (有些可能已經預設勾選)︰ 
 * Files.Read.All
 * offline_access
 * openid
 * 設定檔

1. 按一下對話方塊底端的 [確定]****。

1. 按一下註冊頁面底端的 [儲存]****。

## <a name="grant-admin-consent-to-the-add-in"></a>授與增益集的系統管理員同意權限

1. 如果增益集不是在 Visual Studio 中執行，請按下 F5 鍵來執行它。它必須在 IIS 中執行，讓這個程序順利完成。 

1. 在下列字串中，將預留位置 "{application_ID}" 取代為您註冊增益集時所複製的應用程式 ID。

    `https://login.microsoftonline.com/common/adminconsent?client_id={application_ID}&state=12345`

1. 將產生的 URL 貼至瀏覽器網址列，並且瀏覽到它。

1. 出現提示時，使用系統管理員認證登入您的 Office 365 租用。

1. 然後系統會提示您授與增益集的權限，以存取您的 Microsoft Graph 資料。按一下 [接受]****。 

1. 然後瀏覽器視窗/索引標籤會重新導向至您在註冊增益集時指定的**重新導向 URL**，讓增益集的首頁在瀏覽器中開啟。 

2. 在瀏覽器的網址列中，您會看到「租用戶」查詢參數與 GUID 值。這是您 Office 365 租用的 ID。複製並儲存此值。您將在稍後的步驟中使用它。

3. 關閉視窗/索引標籤。

1. 在 Visual Studio 中停止偵錯工具。

## <a name="configure-the-add-in"></a>設定增益集

1. 在下列字串中，將預留位置 “{tenant_ID}” 取代為您稍早取得的 Office 365 租用戶 ID。如果由於任何原因，您稍早並沒有取得 ID，請使用[尋找您的 Office 365 租用戶 ID](https://support.office.com/en-us/article/Find-your-Office-365-tenant-ID-6891b561-a52d-4ade-9f39-b492285e2c9b) 中的其中一個方法以取得它。

    `https://login.microsoftonline.com/{tenant_ID}/v2.0`

1. 在 Visual Studio 中，開啟 web.config。在 **appSettings** 區段中有一些機碼，您必須為其指派值。

1. 使用您在步驟 1 中建構的字串，作為名稱為 “ida:Issuer” 之機碼的值。請確認在值中沒有任何空白。

1. 將下列值給予對應的機碼︰

|機碼|值|
|:-----|:-----|
|ida:ClientID|當您註冊增益集時取得的應用程式 ID。|
|ida:Audience|當您註冊增益集時取得的應用程式 ID。|
|ida:Password|當您註冊增益集時取得的密碼。|


以下是您所變更的四個機碼的範例。*請注意 ClientID 和對象是相同的*。

    ```xml
    <add key=”ida:ClientID" value="12345678-1234-1234-1234-123456789012" />
    <add key="ida:Audience" value="12345678-1234-1234-1234-123456789012" />
    <add key="ida:Password" value="rFfv17ezsoGw5XUc0CDBHiU" />
    <add key="ida:Issuer" value="https://login.microsoftonline.com/aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee/v2.0" />
    ```

> **附註：**保留 **appSettings** 區段中的其他設定。


1. 儲存後關閉檔案。

1. 在增益集專案中，開啟增益集資訊清單檔案 “Office-Add-in-ASPNET-SSO.xml”。

1. 捲動至檔案底端。

1. 在結束 </VersionOverrides> 標記的正上方，您會發現下列標記︰

    ```xml
    <WebApplicationId>{application_GUID here}</WebApplicationId>
    <WebApplicationResource>api://localhost:44355<WebApplicationResource>
    <WebApplicationScopes>
        <WebApplicationScope>profile</WebApplicationScope>
        <WebApplicationScope>openid</WebApplicationScope>
        <WebApplicationScope>offline_access</WebApplicationScope>
        <WebApplicationScope>files.read.all</WebApplicationScope>
    </WebApplicationScopes>
    ```

1. 將標記中的預留位置 “{application_GUID here}” 取代為您註冊增益集時所複製的應用程式 ID。這是您用於 web.config 中 ClientID 和對象的相同 ID。

    >附註： 
    >
    >* **WebApplicationResource** 值是您將 Web API 平台新增至增益集的註冊時所設定的**應用程式 ID URI**。
    >* 如果增益集是透過 Office 市集銷售，**WebApplicationScopes** 區段只用於產生同意對話方塊。

1. 儲存後關閉檔案。

## <a name="code-the-client-side"></a>撰寫用戶端

1. 開啟 [指令碼]**** 資料夾中的 Home.js 檔案。在其中已經有一些程式碼︰

    * 對 `Office.initialize` 方法的指派會轉為將控制碼指派給 `getGraphAccessTokenButton` 按鈕點擊事件。
    * `showResult` 方法，在工作窗格底端顯示從 Microsoft Graph 傳回的資料 (或錯誤訊息)。

1. 在 `Office.initialize` 指派的下方，新增以下程式碼。請注意有關這段程式碼的下列各項︰ 

    * `getAccessTokenAsync` 是 Office.js 中新的 API，可讓增益集向 Office 主應用程式 (Excel、PowerPoint、Word 等等) 要求增益集的存取權杖 (針對已登入 Office 的使用者)。Office 主應用程式接下來會要求 Azure AD 2 端點以取得權杖。因為您在註冊時，將 Office 主應用程式預先授權給增益集，所以 Azure AD 會傳送權杖。 
    * 如果沒有任何使用者登入 Office，Office 主應用程式會提示使用者登入。 
    * 選項參數會將 `forceConsent` 設定為 false，因此系統不會提示使用者同意將 Office 主應用程式的存取權給予您的增益集。

    ```js
    function getOneDriveItems() {
    Office.context.auth.getAccessTokenAsync({ forceConsent: false },
        function (result) {
            if (result.status === "succeeded") {
                // TODO1: Use the access token to get Microsoft Graph data.
            }
            else {
                console.log("Code: " + result.error.code);
                console.log("Message: " + result.error.message);
                console.log("name: " + result.error.name);
                document.getElementById("getGraphAccessTokenButton").disabled = true;
            }
        });
    }
    ```

1. 使用下列行取代 TODO1。您會在稍後的步驟中建立 `getData` 方法和伺服器端 “/api/values” 路由。相對 URL 用於端點，因為它必須裝載於與您的增益集相同的網域。

    ```js
    accessToken = result.value;
    getData("/api/values", accessToken);
    ```

1. 在 `getOneDriveFiles` 方法下方，新增下列項目。這個公用程式方法會呼叫指定的 Web API 端點，並且傳遞與 Office 主應用程式用來取得增益集存取權的相同存取權杖給它。在伺服器端，此存取權杖將用於「代表」流程，以取得 Microsoft Graph 的存取權杖。 

    ```js
    function getData(relativeUrl, accessToken) {
        $.ajax({
            url: relativeUrl,
            headers: { "Authorization": "Bearer " + accessToken },
            type: "GET",
        })
        .done(function (result) {
            showResult(result);
        })
        .fail(function (result) {
            console.log(result.error);
        });
    }
    ```

1. 儲存後關閉檔案。

## <a name="code-the-server-side"></a>撰寫伺服器端

### <a name="configure-the-owin-middleware"></a>設定 OWIN 中介軟體

1. 開啟專案根目錄中的 Startup.cs 檔案。 

1. 如果尚未存在，將關鍵字 `partial` 新增至啟動類別的宣告。它看起來應該像這樣︰

    `public partial class Startup`

1. 將下列行新增至 `Configure` 方法的主體。您會在稍後的步驟中建立 `ConfigureAuth` 方法。

    `ConfigureAuth(app);`

1. 儲存後關閉檔案。

1. 以滑鼠右鍵按一下 [App_Start]**** 資料夾，然後選取 [新增 | 類別]****。 

1. 在 [新增項目]**** 對話方塊中，將檔案命名為 **Startup.Auth.cs**，然後按一下 [新增]****。

1. 將新檔案中的命名空間名稱縮短為 `Office_Add_in_ASPNET_SSO_WebAPI`。

1. 確保所有下列 `using` 陳述式是在檔案頂端。 

   ```
    using Owin;
    using System.IdentityModel.Tokens;
    using System.Configuration;
    using Microsoft.Owin.Security.OAuth;
    using Microsoft.Owin.Security.Jwt;
    using Office_Add_in_ASPNET_SSO_WebAPI.App_Start;
    ```

1. 如果尚未存在，將關鍵字 `partial` 新增至 `Startup` 類別的宣告。它看起來應該像這樣︰

    `public partial class Startup`

1. 將下列方法新增至 `Startup` 類別。這個方法會指定 OWIN 中介軟體如何驗證存取權杖，該存取權杖是從 client-side Home.js 檔案中的 `getData` 方法傳遞給它。呼叫以 `[Authorize]` 屬性裝飾的 Web API 端點時，都會觸發授權程序。

    ```
    public void ConfigureAuth(IAppBuilder app)
    {
        // TODO2: Configure the validation settings
        // TODO3: Specify the type of authorization and the discovery endpoint
        // of the secure token service.
    }
    ```

1. 使用下列項目取代 TODO2。附註：

    * 程式碼指示 OWIN，以確保來自 Office 主應用程式之存取權杖中所指定的對象和權杖簽發者 (由用戶端呼叫 `getData` 來傳遞) 必須符合在 web.config 中指定的值。
    * 將 `SaveSigninToken` 設定為 `true` 會讓 OWIN 儲存來自 Office 主應用程式的未經處理權杖。增益集需要它以使用「代表」流程取得 Microsoft Graph 的存取權杖。
    * 範圍未由 OWIN 中介軟體驗證。存取權杖的範圍，其中應該包含 `access_as_user`，在控制站中進行驗證。

    ```
    var tvps = new TokenValidationParameters
        {
            ValidAudience = ConfigurationManager.AppSettings["ida:Audience"],
            ValidIssuer = ConfigurationManager.AppSettings["ida:Issuer"],
            SaveSigninToken = true
        };
    ```

1. 使用下列項目取代 TODO3。附註：

    * 會呼叫方法 `UseOAuthBearerAuthentication`，而不是比較常見的 `UseWindowsAzureActiveDirectoryBearerAuthentication`，因為後者與 Azure AD V2 端點不相容。
    * 傳遞給方法的探索 URL 是 OWIN 中介軟體取得指示的位置，該指示是用來取得機碼，當它驗證從 Office 主應用程式收到的存取權杖上的簽章時需要此機碼。

    ```
    app.UseOAuthBearerAuthentication(new OAuthBearerAuthenticationOptions
            {
                AccessTokenFormat = new JwtFormat(tvps, new OpenIdConnectCachingSecurityTokenProvider("https://login.microsoftonline.com/common/v2.0/.well-known/openid-configuration"))
            });
    ```

1. 儲存後關閉檔案。

### <a name="create-the-apivalues-controller"></a>建立 /api/values 控制站

1. 開啟檔案 **Controllers\ValueController.cs**。 

1. 確保下列 `using` 陳述式是在檔案頂端。

    ```
    using Microsoft.Identity.Client;
    using System.IdentityModel.Tokens;
    using System.Collections.Generic;
    using System.Configuration;
    using System.Linq;
    using System.Security.Claims;
    using System.Threading.Tasks;
    using System.Web.Http;
    using Office_Add_in_ASPNET_SSO_WebAPI.Helpers;
    using Office_Add_in_ASPNET_SSO_WebAPI.Models;
    ```

1. 在宣告 `ValuesController` 的行上方，新增屬性 `[Authorize]`。這可確保每當呼叫控制站方法時，您的增益集都會執行在最後一個程序中設定的授權程序，只有具有您的增益集的有效存取權杖的呼叫端可以叫用控制站的方法。 

1. 將下列方法新增至 `ValuesController`：

    ```
    // GET api/values
    public async Task<IEnumerable<string>> Get()
    {
        // TODO4: Validate the scopes of the access token.
    }
    ```

1. 使用下列程式碼取代 TODO4，以驗證在權杖中指定的範圍包括 `access_as_user`。 

    ```
    string[] addinScopes = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/scope").Value.Split(' ');
    if (addinScopes.Contains("access_as_user"))
    {
        // TODO5: Get the raw token that the add-in page received from the Office host.
        // TODO6: Get the access token for MS Graph.
        // TODO7: Get the names of files and folders in OneDrive for Business by using the Microsoft Graph API.
        // TODO8: Remove excess information from the data and send the data to the client.
    }
    return new string[] { "Error", "Microsoft Office does not have permission to get Microsoft Graph data on behalf of the current user." };
    ```

1. 使用下列程式碼取代 TODO5，該程式碼會將從 Office 主應用程式收到的未經處理存取權杖轉換為 `UserAssertion` 物件，該物件會傳遞至其他方法。

    ```
    var bootstrapContext = ClaimsPrincipal.Current.Identities.First().BootstrapContext as BootstrapContext;
    UserAssertion userAssertion = new UserAssertion(bootstrapContext.Token);
    ```

1. 使用下列程式碼取代 TODO6。附註：

    * 您的增益集不再扮演資源 (或對象) 的角色，Office 主應用程式和使用者需要該角色的存取權。現在它本身是用戶端，需要 Microsoft Graph 的存取權。`ConfidentialClientApplication` 是 MSAL「用戶端內容」物件。 
    * `ConfidentialClientApplication` 建構函式的第三個參數是重新導向 URL，實際上不會在「代表」流程中使用，但是它是使用正確 URL 的良好做法。第四個和第五個參數可以用來定義持續性存放區，可以跨不同工作階段使用增益集重複使用未過期的權杖。這個範例不會實作任何持續性儲存體。
    * `ConfidentialClientApplication.AcquireTokenOnBehalfOfAsync` 方法會先在 MSAL 快取 (在記憶體中) 查尋相符的存取權杖。只有在沒有相符項目時，它才會使用 Azure AD V2 端點起始「代理」流程。

    ```
    ClientCredential clientCred = new ClientCredential(ConfigurationManager.AppSettings["ida:Password"]);
    ConfidentialClientApplication cca =
                    new ConfidentialClientApplication(ConfigurationManager.AppSettings["ida:ClientID"],
                                                      "https://localhost:44355", clientCred, null, null);
    string[] graphScopes = { "profile", "Files.Read.All" };
    AuthenticationResult result = await cca.AcquireTokenOnBehalfOfAsync(graphScopes, userAssertion, "https://login.microsoftonline.com/common/oauth2/v2.0");
    ```

1. 使用下列項目取代 TODO7。附註：

    * `GraphApiHelper` 和 `ODataHelper` 類別是在 [協助程式]**** 資料夾的檔案中定義。`OneDriveItem` 類別是在 [模型]**** 資料夾的檔案中定義。這些類別的詳細討論與授權或 SSO 不相關，所以超出本文的範圍。
    * 藉由僅向 Microsoft Graph 要求實際需要的資料，讓效能獲得改善，因此程式碼會使用 ` $select` 查詢參數來指定我們只要名稱屬性，以及使用 `$top` 參數來指定我們只要檔案名稱的前 3 個資料夾。

    ```
    var fullOneDriveItemsUrl = GraphApiHelper.GetOneDriveItemNamesUrl("?$select=name&$top=3");
    var getFilesResult = await ODataHelper.GetItems<OneDriveItem>(fullOneDriveItemsUrl, result.AccessToken);
    ```

1. 使用下列項目取代 TODO8。請注意，雖然上述程式碼只要求 OneDrive 項目的名稱**屬性，但是 Microsoft Graph 永遠會包含 OneDrive 項目的 eTag** 屬性。為了減少傳送至用戶端的裝載，下列程式碼僅會使用項目名稱重新建構結果。

    ```
    List<string> itemNames = new List<string>();
    foreach (OneDriveItem item in getFilesResult)
    {
      itemNames.Add(item.Name);
    }                    
    return itemNames;
    ```

## <a name="run-the-add-in"></a>執行增益集

1. 請確定您在商務用 OneDrive 中有一些檔案。

1. 在 Visual Studio 中，按下 F5。PowerPoint 隨即開啟，且在 [首頁]**** 功能區中有 **SSO ASP.NET** 群組。 

1. 按下此群組中的 [顯示增益集]**** 按鈕，以在工作窗格中查看增益集的 UI。 

1. 按下 [從 OneDrive 取得我的檔案]**** 按鈕。如果您未登入 Office，系統會提示您登入。

1. 在您登入之後，商務用 OneDrive 上檔案和資料夾的清單會出現在按鈕下方。這可能需要超過 15 秒的時間，尤其是第一次執行時。 



