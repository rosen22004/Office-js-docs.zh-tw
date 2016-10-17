
# <a name="call-a-service-from-an-outlook-add-in-by-using-an-identity-token-in-exchange"></a>在 Exchange 中使用識別權杖以從 Outlook 增益集呼叫服務

識別權杖會為每個您可用來個人化您所提供的服務的客戶提供的唯一識別項。藉由使用會傳回字串至 Outlook 增益集的非同步方法呼叫，您的程式碼可以向 Exchange Server 要求識別權杖。字串包含 JSON Web 權杖 (JWT) 的識別權杖。增益集不需要解壓縮權杖。相反地，它會將權杖傳遞到您的 web 服務，讓您的服務可以驗證來自增益集的要求。

支援增益集的 web 服務必須在裝載增益集 HTML 和 JavaScript 來源檔的相同伺服器上執行。這可防止跨網站指令碼錯誤。如果您的應用程式需要的話，則您的伺服器可以將要求 Proxy 至其他 web 服務。

將識別權杖加入增益集傳送的服務要求很簡單。您要求權杖、使用權杖，然後再使用 web 服務回應。以下是您使用 **XmlHttpRequest** 方法傳送給您伺服器的簡單 XML 文件的外觀。

## <a name="request-a-token-from-your-exchange-server"></a>從您的 Exchange Server 要求權杖


這個簡單的增益集初始化方法使用 **getUserIdentityTokenAsync** 方法，從 Exchange Server 要求識別權杖。_getUserIdentityToken_ 參數是對伺服器的非同步要求傳回時所呼叫的函數。請參閱下一個步驟以了解回撥方法。


```js
var _mailbox;
var _xhr;
// The initialize function is required for all add-ins.
Office.initialize = function () {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
        _mailbox = Office.context.mailbox;
    _mailbox.getUserIdentityTokenAsync(getUserIdentityTokenCallback);
    });
}

```


## <a name="use-the-identity-token"></a>使用識別權杖


**getUserIdentityTokenAsync** 方法的回撥函數具有一個在其 **value** 屬性中包含使用者識別權杖的參數。

這個回撥函數會建立 **XMLHttpRequest** 物件來呼叫 web 服務。將 **XMLHttpRequest** 物件上的 **onreadystatechange** 屬性設定為增益集取得來自 web 服務的回應時所需執行的函數名稱。




```js
function getUserIdentityTokenCallback(asyncResult) {
    var token = asyncResult.value;

    _xhr = new XMLHttpRequest();
    _xhr.open("POST", "https://localhost:44300/IdentityTestService/UnpackTokenJSON");
    _xhr.setRequestHeader("Content-Type", "application/json; charset=utf-8");
    _xhr.onreadystatechange = readyStateChange;

    var request = new Object();
    request.token = token;
    request.phoneNumbers = _mailbox.item.getEntities().phoneNumbers;

    _xhr.send(JSON.stringify(request));
}
```


## <a name="use-the-web-service-response"></a>使用 web 服務回應


這是處理來自 web 服務的回應的其他簡單函數。其依照 **XHMHttpResponse** 回撥函數的標準模式。它會等候整個來自 web 服務的回應傳入，然後將回應的內容放在增益集的 UI 上。這個函數剖析的回應是來自 web 服務的回應。如需有關此回應的資訊，請參閱[驗證 Exchange 識別權杖](../outlook/validate-an-identity-token.md)。 


```js
function readyStateChange() {
    if (_xhr.readyState == 4 &amp;&amp; _xhr.status == 200) {

        var response = JSON.parse(_xhr.responseText);

        if (undefined == response.error) {
            document.getElementById("msexchuid").value = response.token.msexchuid;
            document.getElementById("amurl").value = response.token.amurl;
            document.getElementById("uniqueID").value = response.token.uniqueID;
            document.getElementById("iss").value = response.token.iss;
            document.getElementById("x5t").value = response.token.x5t;
            document.getElementById("nbf").value = response.token.nbf;
            document.getElementById("exp").value = response.token.exp;
        }
        else {
            document.getElementById("error").value = response.error;
        }
    }
}
```


## <a name="example:-calling-a-web-service-with-identity-tokens"></a>範例：使用識別權杖呼叫 web 服務


識別權杖會提供關於呼叫您的服務到您的伺服器上執行的 web 服務用戶端的識別資訊。若要使用識別權杖，需要具備下列項目︰


- 要求來自 Exchange Server 的識別權杖，並傳送至您的 web 服務的 Outlook 增益集。本主題中的資訊將協助您建立該增益集。
    
- 在針對可驗證識別權杖的增益集提供 UI 的伺服器上執行的 web 服務。您會在下列其中一個主題中找到您建立 web 服務所需的資訊︰
    
      - [使用 Exchange 權杖驗證程式庫](../outlook/use-the-token-validation-library.md) -- 如果您在使用我們提供的驗證程式庫。
    
  - [驗證 Exchange 識別權杖](../outlook/validate-an-identity-token.md) -- 如果您在撰寫您自己的驗證程式碼。
    

### <a name="code-for-the-sample-add-in"></a>範例增益集的程式碼


本文所述的增益集需要下列檔案︰


- IdentityTest.js - 為增益集提供商務邏輯的 JavaScript 檔案。
    
- IdentityTest.html - 為增益集提供 UI 的 HTML 檔案。
    
您也需要識別測試 web 服務。如需有關該 web 服務的資訊，請參閱[驗證 Exchange 識別權杖](../outlook/validate-an-identity-token.md)。


#### <a name="identitytest.js"></a>IdentityTest.js

下列範例顯示 IdentityTest.js 檔案。


```js
var _mailbox;
var _xhr;

// The initialize function is required for all add-ins.
Office.initialize = function () {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    _mailbox = Office.context.mailbox;
    _mailbox.getUserIdentityTokenAsync(getUserIdentityTokenCallback);
    });
}
function getUserIdentityTokenCallback(asyncResult) {
    var token = asyncResult.value;

    _xhr = new XMLHttpRequest();
    _xhr.open("POST", "https://localhost:44300/IdentityTestService/UnpackTokenJSON");
    _xhr.setRequestHeader("Content-Type", "application/json; charset=utf-8");
    _xhr.onreadystatechange = readyStateChange;

    var request = new Object();
    request.token = token;
    request.phoneNumbers = _mailbox.item.getEntities().phoneNumbers;

    _xhr.send(JSON.stringify(request));
}

function readyStateChange() {
    if (_xhr.readyState == 4 &amp;&amp; _xhr.status == 200) {

        var response = JSON.parse(_xhr.responseText);

        if (undefined == response.error) {
            document.getElementById("msexchuid").value = response.token.msexchuid;
            document.getElementById("amurl").value = response.token.amurl;
            document.getElementById("uniqueID").value = response.token.uniqueID;
            document.getElementById("iss").value = response.token.iss;
            document.getElementById("x5t").value = response.token.x5t;
            document.getElementById("nbf").value = response.token.nbf;
            document.getElementById("exp").value = response.token.exp;
        }
        else {
            document.getElementById("error").value = response.error;
        }
    }
}
```


#### <a name="identitytest.html"></a>IdentityTest.html

下列範例顯示 IdentityTest.html 檔案。


```HTML
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
    <title>Identity Test</title>

    <link rel="stylesheet" type="text/css" href="../Content/Office.css" />
    <link rel="stylesheet" type="text/css" href="../Content/App.css" />

    <script src="../Scripts/jquery-1.6.2.js"></script>
    <script src="../Scripts/Office/MicrosoftAjax.js"></script>
    <script src="../Scripts/Office/Office.js"></script>

    <!-- Add your JavaScript to the following JavaScript file -->
    <script src="../Scripts/IdentityTest.js"></script>
</head>
<body>
    <div id="SectionContent">
        <table style="width: 80%;">
            <tr>
                <th>Claim
                </th>
                <th>Contents
                </th>
            </tr>
            <tr>
                <td style="width: 25%;">Error:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="error" value="None" />
                </td>
            </tr>
            <tr>
                <td style="width: 25%;">User Exchange ID:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="msexchuid" />
                </td>
            </tr>
            <tr>
                <td style="width: 25%;">Authentication Metadata URL:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="amurl" />
                </td>
            </tr>
            <tr>
                <td style="width: 25%;">Unique identifier:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="uniqueID" />
                </td>
            </tr>
          </tr>
            <tr>
                <td style="width: 25%;">Audience:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="aud" />
                </td>
            </tr>
            <tr>
                <td style="width: 25%;">Issuer:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="iss" />
                </td>
            </tr>
            <tr>
                <td style="width: 25%;">Certificate thumbprint:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="x5t" />
                </td>
            </tr>
            <tr>
                <td style="width: 25%;">Valid from:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="nbf" />
                </td>
            </tr>
            <tr>
                <td style="width: 25%;">Valid to:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="exp" />
                </td>
            </tr>
        </table>
    </div>
</body>
</html>
```


## <a name="next-steps"></a>後續步驟


既然您知道如何要求識別權杖，您需要使用要求的伺服器端的權杖。下列文章會協助您快速入門︰


- [使用 Exchange 權杖驗證程式庫](../outlook/use-the-token-validation-library.md)
    
- [驗證 Exchange 識別權杖](../outlook/validate-an-identity-token.md)
    
- [使用 Exchange 的識別權杖來驗證使用者](../outlook/authenticate-a-user-with-an-identity-token.md)
    

## <a name="additional-resources"></a>其他資源



- [使用 Exchange 識別權杖來驗證 Outlook 增益集](../outlook/authentication.md)
    
- [Exchange 識別權杖的內容](../outlook/inside-the-identity-token.md)
    
