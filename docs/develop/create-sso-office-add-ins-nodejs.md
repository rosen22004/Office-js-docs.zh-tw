# <a name="create-a-nodejs-office-add-in-that-uses-single-sign-on-preview"></a>建立使用單一登入的 Node.js Office 增益集 (預覽)

可以登入 Office 與 Office Web 增益集的使用者，可以利用這個登入程序為增益集和 Microsoft Graph 的使用者授權，而不需要使用者登入第二次。如需概觀，請參閱[在 Office 增益集啟用 SSO](../../docs/develop/sso-in-office-add-ins.md)。

本文會引導您完成啟用在增益集單一登入 (SSO) 的程序，該增益集是使用 Node.js 和運算式所建立。 

> **附註：**如需有關以 ASP.NET 為基礎的增益集之類似文章，請參閱[建立使用單一登入的 ASP.NET Office 增益集](../../docs/develop/create-sso-office-add-ins-aspnet.md)。

## <a name="prerequisites"></a>必要條件

* [節點和 npm](https://nodejs.org/en/)，版本 6.9.4 或更新版本。
* [Git Bash](https://git-scm.com/downloads) (或其他 git 用戶端。)
* TypeScript 2.2.2 版或更新版本。
* Office 2016 1704 版、組建 8027.nnnn 或更新版本 (Office 365 訂閱版本，有時候稱為「隨選即用」)。您必須是 Office 測試人員才能取得這個版本。如需詳細資訊，請參閱[成為 Office 測試人員](https://products.office.com/en-us/office-insider?tab=tab-1)。

## <a name="set-up-the-starter-project"></a>設定起始專案

1. 複製或下載位於 [Office 增益集 NodeJS SSO](https://github.com/officedev/office-add-in-nodejs-sso) 的存放庫。 


    > **附註：**範例有兩種版本。 
    > 
    > * [Before]**** 資料夾是起始專案。未直接連接至 SSO 或授權的 UI 和增益集的其他方面已完成。這篇文件的後續章節會引導您進行完成它的程序。 
    > * 如果您完成這篇文章的程序，**已完成**版本範例就如同您擁有的增益集，例外的是已完成專案具有程式碼註解，在這篇文章中可能是多餘的文字。若要使用已完成的版本，只要遵循本文中的指示，但是以 "Completed" 取代 "Before"，並且略過**撰寫用戶端**和**撰寫伺服器端**區段。

1. 開啟 [Before]**** 資料夾中的 Git bash 主控台。

2. 在主控台中輸入 `npm install` 以安裝 package.json 檔案中的所有分項相依性。

3. 在主控台中輸入 `npm run build ` 以建置專案。 
     > 附註：您可能會看到一些建置錯誤，指出某些變數已宣告但未使用。請忽略這些錯誤。它們事實上是 "Before" 版本範例遺漏某些程式碼 (稍後會新增) 的副作用。

## <a name="register-the-add-in-with-azure-ad-v2-endpoint"></a>使用 Azure AD V2 端點註冊增益集

1. 瀏覽至 [https://apps.dev.microsoft.com/?test=build2017](https://apps.dev.microsoft.com/?test=build2017)。 

1. 使用系統管理員認證登入您的 Office 365 租用。例如，MyName@contoso.onmicrosoft.com

1. 按一下 [新增應用程式]****。

1. 出現提示時，使用「Office-Add-in-NodeJS-SSO」作為應用程式名稱，然後按一下 [建立應用程式]****。

1. 當應用程式的組態頁面開啟時，複製**應用程式 ID** 並且儲存。您將在稍後的程序中使用它。 

    > 附註：當其他應用程式，例如 Office 主應用程式 (例如 PowerPoint、Word、Excel) 尋求應用程式的授權存取時，這個 ID 是「對象」值。當應用程式轉為尋求 Microsoft Graph 的授權存取時，它也是應用程式的「用戶端 ID」。

1. 在 [應用程式祕密]**** 區段中，按下 [產生新的密碼]****。快顯對話方塊隨即開啟，並且顯示新的密碼 (也稱為「應用程式祕密」)。立即複製密碼，並且將它與應用程式 ID 一起儲存。**您在稍後的程序中需要它。然後關閉對話方塊。

1. 在 [平台]**** 區段中，按一下 [新增平台]****。 

1. 在開啟的對話方塊中，選取 [Web API]****。

1. 已產生表單的**應用程式 ID URI**，“api://{App ID GUID}”。以 “localhost:3000” 取代 GUID。整個 ID 應該會是 `api://localhost:3000`。(**應用程式 ID URI** 的**範圍**名稱網域部分將會自動變更為相符項目。它應該會是 `api://localhost:3000/access_as_user`。)

1. 這個步驟和下一個步驟會將 Office 主應用程式存取權給予您的增益集。在 [預先授權應用程式]**** 區段中，有空白的 [應用程式 ID]**** 方塊。在方塊中輸入下列 ID (這是 Microsoft Office 的 ID)：`d3590ed6-52b3-4102-aeff-aad2292ab01c`。

1. 開啟 [應用程式 ID]**** 旁的 [範圍]**** 下拉式清單，並且勾選 `api://localhost:3000/access_as_user` 的方塊。

1. 在 [平台]**** 區段的上方附近，再次按一下 [新增平台]****，並且選取 [Web]****。

1. 在 [平台]****底下新的 [Web]**** 區段中，輸入下列項目作為**重新導向 URL**：`https://localhost:3000`。 

    > 附註：本文撰寫時，**Web API** 平台有時候會從 [平台]**** 區段中消失，特別是當頁面在 **Web** 平台已新增且註冊網頁已儲存之後重新整理時**。為了再次確保您的 **Web API** 平台仍然是註冊的一部分，按一下頁面底端附近的 [編輯應用程式資訊清單]**** 按鈕。您應該會在資訊清單的 **identifierUris** 屬性中看到 `api://localhost:3000` 字串。也會有 **oauth2Permissions** 屬性，它的**值**子屬性具有值 `access_as_user`。

1. 向下捲動到 [Microsoft Graph 權限]**** 區段、[委派的權限]**** 子區段。使用 [新增]**** 按鈕以開啟 [選取權限]**** 對話方塊。

1. 在對話方塊中，勾選下列權限的方塊 (有些可能已經預設勾選)︰ 
    * Files.Read.All
    * 設定檔


1. 按一下對話方塊底端的 [確定]****。

1. 按一下註冊頁面底端的 [儲存]****。

## <a name="grant-admin-consent-to-the-add-in"></a>授與增益集的系統管理員同意權限

1. 在下列字串中，將預留位置 "{application_ID}" 取代為您註冊增益集時所複製的應用程式 ID。

    `https://login.microsoftonline.com/common/adminconsent?client_id={application_ID}&state=12345`

1. 將產生的 URL 貼至瀏覽器網址列，並且瀏覽到它。

1. 出現提示時，使用系統管理員認證登入您的 Office 365 租用。

1. 然後系統會提示您授與增益集的權限，以存取您的 Microsoft Graph 資料。按一下 [接受]****。 

1. 然後瀏覽器視窗/索引標籤會重新導向至您在註冊增益集時指定的**重新導向 URL**，因此，如果增益集正在執行，增益集的首頁會在瀏覽器中開啟。如果增益集未執行，您將得到錯誤，指出找不到 localhost:3000 的資源，或者無法開啟。但是已嘗試重新導向的事實，表示系統管理員同意處理程序已順利完成**。不論首頁是否開啟，或者您是否遇到錯誤，您都可以繼續進行下一個步驟。

2. 在瀏覽器的網址列中，您會看到「租用戶」查詢參數與 GUID 值。這是您 Office 365 租用的 ID。複製並儲存此值。您將在稍後的步驟中使用它。

3. 關閉視窗/索引標籤。

## <a name="configure-the-add-in"></a>設定增益集

1. 在您的程式碼編輯器中開啟 src\server.ts 檔案。在頂端附近，有對於 `AuthModule` 類別之建構函式的呼叫。建構函式中有一些您需要為其指派值的字串參數。

2. 對於 `client_id` 屬性，以您在註冊增益集時儲存的應用程式 ID 取代預留位置 `{client GUID}`。當您完成之後，應該只有在單引號中的 GUID。此處不應有任何 "{}" 字元。

3. 對於 `client_secret` 屬性，以您在註冊增益集時儲存的應用程式秘密取代預留位置 `{client secret}`。

4. 對於 `audience` 屬性，以您在註冊增益集時儲存的應用程式 ID 取代預留位置 `{audience GUID}`。(您指派給 `client_id` 屬性的相同值。)
  
3. 在指派給 `issuer` 屬性的字串中，您會看到預留位置 *{O365 tenant GUID}*。以在最後一個程序結尾儲存的 Office 365 租用 ID 取代這個項目。如果由於任何原因，您稍早並沒有取得 ID，請使用[尋找您的 Office 365 租用戶 ID](https://support.office.com/en-us/article/Find-your-Office-365-tenant-ID-6891b561-a52d-4ade-9f39-b492285e2c9b) 中的其中一個方法以取得它。當您完成之後，`issuer` 屬性值看起來應該如下所示︰

    `https://login.microsoftonline.com/12345678-1234-1234-1234-123456789012/v2.0`

    > **附註：**保留 `AuthModule` 建構函式中的其他參數。

1. 儲存後關閉檔案。

1. 在專案的根目錄中，開啟增益集資訊清單檔案 “Office-Add-in-NodeJS-SSO.xml”。

1. 捲動至檔案底端。

1. 在結束 `</VersionOverrides>` 標記的正上方，您會發現下列標記︰

    ```
    <WebApplicationId>{application_GUID here}</WebApplicationId>
    <WebApplicationResource>api://localhost:3000<WebApplicationResource>
    <WebApplicationScopes>
        <WebApplicationScope>profile</WebApplicationScope>
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

1. 在 [公用]**** 資料夾中開啟 program.js 檔案。在其中已經有一些程式碼︰

    * 對 `Office.initialize` 方法的指派會轉為將控制碼指派給 `getGraphAccessTokenButton` 按鈕點擊事件。
    * `showResult` 方法，在工作窗格底端顯示從 Microsoft Graph 傳回的資料 (或錯誤訊息)。

1. 在 `Office.initialize` 指派的下方，新增以下程式碼。請注意有關這段程式碼的下列各項︰ 

     * `getAccessTokenAsync` 是 Office.js 中新的 API，可讓增益集向 Office 主應用程式 (Excel、PowerPoint、Word 等等) 要求增益集的存取權杖 (針對已登入 Office 的使用者)。Office 主應用程式接下來會要求 Azure AD 2 端點以取得權杖。因為您在註冊時，將 Office 主應用程式預先授權給增益集，所以 Azure AD 會傳送權杖。 
     * 如果沒有任何使用者登入 Office，Office 主應用程式會提示使用者登入。 
     * 選項參數會將 `forceConsent` 設定為 false，因此系統不會提示使用者同意將 Office 主應用程式的存取權給予您的增益集。

    ```
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

    ```
    accessToken = result.value;
    getData("/api/onedriveitems", accessToken);
    ```

1. 在 `getOneDriveFiles` 方法下方，新增下列項目。這個公用程式方法會呼叫指定的 Web API 端點，並且傳遞與 Office 主應用程式用來取得增益集存取權的相同存取權杖給它。在伺服器端，此存取權杖將用於「代表」流程，以取得 Microsoft Graph 的存取權杖。 

    ```
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

有兩個伺服器端檔案需要修改。 
- src\auth.js 提供授權協助程式函式。它已經有泛型成員，用於各種不同的授權流程。我們必須將函式新增至其中，它會實作「代表」流程。
- src\server.js 檔案具有基本成員，必須執行伺服器並且說明中介軟體。我們必須將函式新增至其中，作為首頁和 Web API 以取得 Microsoft Graph 資料。

### <a name="create-a-method-to-exchange-tokens"></a>建立方法來交換權杖

1. 開啟 \src\auth.ts 檔案。將以下方法新增至 `AuthModule` 類別。請注意有關這段程式碼的下列各項︰
    * jwt 參數是應用程式的存取權杖。在「代表」流程中，它與 AAD 交換以取得資源的存取權杖。
    * 範圍參數具有預設值，但是在這個範例中，它將會以呼叫程式碼覆寫。
    * 資源參數是選用的。它不應該在 STS 是 AAD V2 端點時使用。後者會從範圍推斷資源，如果資源是在 HTTP 要求中傳送，則它會傳回錯誤。 
    

    ```
    private async exchangeForToken(jwt: string, scopes: string[] = ['openid'], resource?: string) {
        try {
            // TODO2: Construct the parameters that will be sent in the body of the 
            //        HTTP Request to the STS that starts the "on behalf of" flow.
            // TODO3: Send the request to the STS.
            // TODO4: Process the response and persist the access token to resource.
        }
        catch (exception) {
            throw new UnauthorizedError('Unable to obtain an access token to the resource' 
                                        + ' ' + exception.message, 
                                        exception);
        }
    }
    ```

2. 使用下列程式碼取代 TODO2。關於此程式碼，請注意︰
    * 支援「代表」流程的 STS 預期在 HTTP 要求主體中有特定屬性/值組。此程式碼會建構物件，該物件將會變成要求的主體。 
    * 資源屬性只有在資源已傳遞至方法時會新增至主體。

    ```
    const v2Params = {
            client_id: this.clientId,
            client_secret: this.clientSecret,
            grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
            assertion: jwt,
            requested_token_use: 'on_behalf_of',
            scope: scopes.join(' ')
        };
        let finalParams = {};
        if (resource) {
            // In JavaScript we could just add the resource property to the v2Params
            // object, but that won't compile in TypeScript.
            let v1Params  = { resource: resource };  
            for(var key in v2Params) { v1Params[key] = v2Params[key]; }
            finalParams = v1Params;
        } else {
            finalParams = v2Params;
        } 
    ```

3. 使用下列程式碼取代 TODO3，該程式碼會將 HTTP 要求傳送至 STS 的權杖端點。

    ```
    const res = await fetch(`${this.stsDomain}/${this.tenant}/${this.tokenURLsegment}`, {
        method: 'POST',
        body: form(finalParams),
        headers: {
            'Accept': 'application/json',
            'Content-Type': 'application/x-www-form-urlencoded'
        }
    }); 
    ```

4. 使用下列程式碼取代 TODO4。請注意，程式碼除了傳回保留資源的存取權杖，還會保留它及其到期時間。呼叫程式碼可以避免對 STS 的不必要呼叫，方法是重複使用資源的未到期存取權杖。您會在下一節看到如何執行這些作業。

    ```
    if (res.status !== 200) {
        const exception = await res.json();
        throw exception;
    }
    const json = await res.json();
    // Persist the token and it's expiration time.
    const resourceToken = json['access_token'];
    ServerStorage.persist('ResourceToken', resourceToken);
    const expiresIn = json['expires_in'];  // seconds until token expires.
    const resourceTokenExpiresAt = moment().add(expiresIn, 'seconds');
    ServerStorage.persist('ResourceTokenExpiresAt', resourceTokenExpiresAt);
    return resourceToken; 
    ```

5. 儲存檔案，但是不要關閉它。

### <a name="create-a-method-to-get-access-to-the-resource-using-the-on-behalf-of-flow"></a>建立方法來取得使用「代表」流程的資源存取權

1. 同樣在 src/auth.ts 中，將以下方法新增至 `AuthModule` 類別。請注意有關這段程式碼的下列各項︰
    * 上方註解是關於 `exchangeForToken` 方法的參數也適用於這個方法的參數。
    * 方法會先檢查持續性儲存體以取得在後續幾分鐘內未到期之資源的存取權杖。它只會在必要時呼叫最後一節中建立的方法。

    ```
    async acquireTokenOnBehalfOf(jwt: string, scopes: string[] = ['openid'], resource?: string) {
        const resourceTokenExpirationTime = ServerStorage.retrieve('ResourceTokenExpiresAt');
        if (moment().add(1, 'minute').diff(resourceTokenExpirationTime) < 1 ) {
            return ServerStorage.retrieve('ResourceToken');
        } else if (resource) {
            return this.exchangeForToken(jwt, scopes, resource);
        } else {
            return this.exchangeForToken(jwt, scopes);
        }
    } 
    ```

2. 儲存後關閉檔案。

### <a name="create-the-endpoints-that-will-serve-the-add-ins-home-page-and-data"></a>建立端點作為增益集的首頁和資料

1. 開啟 src\server.ts 檔案。 

2. 將下列方法新增至檔案底端。此方法會作為增益集的首頁。增益集資訊清單會指定首頁 URL。

    ```
    app.get('/index.html', handler(async (req, res) => {
        return res.sendfile('index.html');
    })); 
    ```

3. 將下列方法新增至檔案底端。此方法會處理 `onedriveitems` API 的任何要求。
    ```
    app.get('/api/onedriveitems', handler(async (req, res) => {
        // TODO5: Initialize the AuthModule object and validate the access token 
        //        that the client-side received from the Office host.
        // TODO6: Get a token to Microsoft Graph from either persistent storage 
        //        or the "on behalf of" flow.
        // TODO7: Use the token to get data from Microsoft Graph.
        // TODO8: Send to the client only the data that it actually needs.
    })); 
    ```

4. 使用下列程式碼取代 TODO5，該程式碼會驗證從 Office 主應用程式收到的存取權杖。`verifyJWT` 方法是在 src\auth.ts 檔案中定義的。它永遠會驗證對象和簽發者。我們會使用選擇性參數來指定我們也想要它驗證存取權杖中的範圍是 `access_as_user`。這是使用者和 Office 主應用程式需要的增益集唯一權限，以使用「代表流程」方式取得 Microsoft Graph 的存取權杖。 

    ```
    await auth.initialize();
    const { jwt } = auth.verifyJWT(req, { scp: 'access_as_user' }); 
    ```

5. 使用下列行取代 TODO6。請注意有關這段程式碼的下列各項︰

    * 對 `acquireTokenOnBehalfOf` 的呼叫不包含資源參數，因為我們已使用 AAD V2 端點建構 `AuthModule` 物件 (`auth`)，該端點不支援資源屬性。
    * 呼叫的第二個參數會指定取得商務用 OneDrive 上使用者檔案和資料夾清單所需的增益集權限。

    `const graphToken = await auth.acquireTokenOnBehalfOf(jwt, ['profile', 'Files.Read.All']);`

6. 使用下列行取代 TODO7。請注意有關這段程式碼的下列各項︰

    * MSGraphHelper 類別是在 src\msgraph-helper.ts 中定義的。 
    * 我們會將必須傳回的資料減至最低，方法是指定我們只想要名稱屬性和前 3 個項目。

    `const graphData = await MSGraphHelper.getGraphData(graphToken, "/me/drive/root/children", "?$select=name&$top=3");`

7. 使用下列程式碼取代 TODO8。請注意，Microsoft Graph 會針對每個項目傳回某些 OData 中繼資料和 **eTag** 屬性，即使 `name` 是唯一要求的屬性。程式碼只會將項目名稱傳送至用戶端。

    ```
    const itemNames: string[] = [];
    const oneDriveItems: string[] = graphData['value'];
    for (let item of oneDriveItems){
        itemNames.push(item['name']);
    }
    return res.json(itemNames);
    ```

8. 儲存後關閉檔案。

## <a name="deploy-the-add-in"></a>部署增益集

現在，您需要讓 Office 知道哪裡可以找到此增益集。

1. 建立網路共用，或[在網路上共用資料夾](https://technet.microsoft.com/en-us/library/cc770880.aspx)。

2. 將一份 Office-Add-in-NodeJS-SSO.xml 資訊清單檔，從專案的根目錄放入共用資料夾中。

3. 啟動 PowerPoint 並且開啟文件。

4. 選擇 [檔案]**** 索引標籤，然後選擇 [選項]****。

5. 選擇 [信任中心]****，然後選擇 [信任中心設定]**** 按鈕。

6. 選擇 [受信任的增益集目錄]****。

7. 在 [目錄 URL]**** 欄位中，輸入包含 Office-Add-in-NodeJS-SSO.xml 的資料夾共用的網路路徑，然後選擇 [新增目錄]****。

8. 選取 [顯示於功能表中]**** 核取方塊，然後選擇 [確定]****。

9. 接著會顯示訊息，通知您下次啟動 Microsoft Office 時就會套用您的設定。關閉 PowerPoint。

## <a name="build-and-run-the-project"></a>建置及執行專案

有兩種方式可以建置及執行專案，取決於您是否使用 Visual Studio 程式碼。針對這兩個方式，當您對程式碼進行變更時，會建置專案和自動重新建置及重新執行。

1. 如果您不是使用 Visual Studio 程式碼︰ 
 2. 開啟節點終端機並且瀏覽至專案的根資料夾。
 3. 在終端機中，輸入 **npm run build**。 
 4. 開啟第二個節點終端機並且瀏覽至專案的根資料夾。
 5. 在終端機中，輸入 **npm run start**。

2. 如果您是使用 VS 程式碼：
 3. 在 VS 程式碼中開啟專案。
 4. 按下 CTRL-SHIFT-B 來建置專案。
 5. 按下 F5 以在偵錯工作階段中執行專案。


## <a name="add-the-add-in-to-an-office-document"></a>將增益集新增至 Office 文件

1. 重新啟動 PowerPoint 並且開啟或建立簡報。 

2. 在 PowerPoint 的 [開發人員]**** 索引標籤上，選擇 [我的增益集]****。

3. 選取 [共用資料夾]**** 索引標籤。

4. 選擇 [SSO NodeJS 範例]****，然後選取 [確定]****。

5. 在 [首頁]**** 功能區上是稱為 **SSO NodeJS** 的新群組，具有標示為 [顯示增益集]**** 的按鈕和圖示。 

## <a name="test-the-add-in"></a>測試增益集

> **附註：**預覽版本 `getAccessTokenAsync` API 只支援工作或學校 (Office 365) 身分識別。如果您使用個人身分識別 (Microsoft 帳戶) 登入 Office，請先登出再繼續作業。**若要測試增益集，您必須從 Office 完全登出，或者使用工作或學校帳戶登入。

1. 請確定您在商務用 OneDrive 帳戶中有一些檔案或資料夾。

2. 按一下 [顯示增益集]**** 按鈕以開啟增益集。

2. 增益集會開啟歡迎畫面。按一下 [從 OneDrive 取得我的檔案]**** 按鈕。

2. 如果您登入 Office，OneDrive 上檔案和資料夾的清單會出現在按鈕下方。第一次執行時可能需要超過 15 秒鐘的時間。

3. 如果您未登入 Office，快顯功能表會開啟，並且提示您登入。在您完成登入之後，您的檔案和資料夾清單會在幾秒鐘後出現。您沒有按下按鈕第二次。**


