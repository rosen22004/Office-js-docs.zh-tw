
# 使用 Exchange 的識別權杖來驗證使用者

您可以對資訊服務實作單一登入 (SSO) 驗證配置，該服務可讓使用 Outlook 增益集的客戶使用他們的 Exchange 伺服器的認證來連線到您的服務。本文章顯示如何使用簡單的 **Dictionary** 物件式使用者資料存放區來比對憑證。

 >**附註：**這只是 SSO 的一個簡單範例，不應在您的實際執行程式碼中使用。一如往常，處理識別和驗證時，必須確定您的程式碼符合組織的安全性需求。


## 使用 SSO 驗證的必要條件


若要對 SSO 使用識別權杖，您的服務應用程式必須具有有效的識別權杖。您可以在下列文章中了解識別權杖，以及如何要求和驗證識別權杖︰


- [Exchange 識別權杖的內容](../outlook/inside-the-identity-token.md)
    
- [在 Exchange 中使用識別權杖以從 Outlook 增益集呼叫服務](../outlook/call-a-service-by-using-an-identity-token.md)
    
- [使用 Exchange 權杖驗證程式庫](../outlook/use-the-token-validation-library.md) - 如果您使用 Managed 程式碼，或[驗證 Exchange 識別權杖](../outlook/validate-an-identity-token.md) - 如果您正在撰寫自己的權杖驗證方法。
    

## 驗證使用者


下列程式碼範例顯示簡易的驗證物件，其符合由具有一組服務憑證的識別權杖所表示的唯一識別。**TokenAuthentication** 類別提供方法 **GetResponseFromService**，該方法會傳回先前已驗證的權杖，或要求使用者提供可驗證並與識別權杖關聯的認證。程式碼尚未完成；它會假設您將提供下列物件和方法。



|**物件/方法**|**說明**|
|:-----|:-----|
|**LocalCredentials** 物件|代表您的服務的使用者認證。物件的結構根據您的服務需求而定。|
|**IdentityToken** 物件|包含 Outlook 增益集傳送至您的服務的使用者識別權杖。該物件至少必須包含使用者的唯一 Exchange 識別碼，以及發出權杖之伺服器的驗證中繼資料 URL。這個範例會使用[驗證 Exchange 識別權杖](../outlook/validate-an-identity-token.md)文章中定義的識別權杖物件。|
|**JsonResponse** 物件|代表來自您的服務的回應。可將物件序列化為 JSON 物件。|
|**CallService** 方法|使用 **LocalCredentials** 物件呼叫您的服務，該物件包含使用者用於服務的認證和包含服務要求資料的物件。如果憑證有效，則這個方法會傳回包含要求的結果的 **JsonReponse** 物件。如果憑證無效，則這個方法會傳回 **null**。|
|**GetCredentialsResponse** 方法|傳回您的郵件 Office 增益集會視為服務憑證要求的 **JsonReponse** 物件。|
|**LocalCredentialsAreValid** 方法|如果提供給服務的認證有效，傳回 **true**；否則它會傳回 **false**。|

 >**附註：**這只是對於如何使用識別權杖的一個建議。一如往常，處理識別和驗證時，必須確定您的程式碼符合組織的安全性需求。


```C#
    public class TokenAuthentication
    {
        // This example uses a Dictionary object to store local credentials. Your application should use
        // a data store that is appropriate to the security requirements of your organization.
        private Dictionary<string, LocalCredentials> AuthenticationCache = new Dictionary<string, LocalCredentials>();

        // Salt to apply when creating unique ID.
        private byte[] Salt = new byte[] {25, 139, 201, 13};

        private JsonResponse CallService(LocalCredentials credentials, object data)
        {
            // Calls the local service to get the response for the user.
            return null;
        }

        private JsonResponse GetCredentialsResponse()
        {
            // Creates a response that tells the Outlook add-in to
            // request the user's credentials for the service.
            return null;
        }

        private bool LocalCredentialsAreValid(LocalCredentials credentials)
        {
            // Returns true if the service recognizes the credentials provided.
            return false;
        }

        private string ComputeSHA256Hash(string uniqueId, string authenticationMetadataUrl, byte[] salt)
        {
            byte[] inputBytes = Encoding.ASCII.GetBytes(string.Concat(uniqueId, authenticationMetadataUrl));

            // Combine input bytes and salt.
            byte[] saltedInput = new byte[salt.Length + inputBytes.Length];
            salt.CopyTo(saltedInput, 0);
            inputBytes.CopyTo(saltedInput, salt.Length);

            // Compute the unique key.
            byte[] hashedBytes = SHA256CryptoServiceProvider.Create().ComputeHash(saltedInput);

            // Convert the hashed value to a string and return.
            return BitConverter.ToString(hashedBytes);
        }

        public JsonResponse GetResponseFromService(IdentityToken token, LocalCredentials credentials, object data)
        {
            JsonResponse response = null;
            // This method should never be called with a null token.
            if (null == token)
            {
                throw new ArgumentNullException("token");
            }

            if (null == credentials)
            {
                string uniqueKey = ComputeSHA256Hash(token.ExchangeID, token.AuthenticationMetadataUrl, Salt);
                if (!AuthenticationCache.ContainsKey(uniqueKey))
                {
                    // The user's credentials are not in the authentication cache. Ask
                    // for the credentials.
                    response = GetCredentialsResponse();
                }
                else
                {
                    // The user's credentials are in the cache; make a request.
                    var serviceResponse = CallService(AuthenticationCache[uniqueKey], data);

                    if (null == serviceResponse)
                    {
                        // There was a problem with the stored credentials. For example,
                        // the user has ended their subscription to the service, or the
                        // credentials have expired. Get new credentials.
                        response = GetCredentialsResponse();
                    }
                    else
                    {
                        // The service returned a response to the user. Return the
                        // service response.
                        response = serviceResponse;
                    }
                }
            }
            else
            {
                // If the credentials are not null, it's a request to add an identity
                // to the authentication cache. Check to determine whether the local credentials
                // sent to the service are known.
                if (LocalCredentialsAreValid(credentials))
                {
                    // The local credentials are known. Add them to the 
                    // cached credentials.
                    string uniqueKey = ComputeSHA256Hash(token.ExchangeID, token.AuthenticationMetadataUrl, Salt);
                    AuthenticationCache.Add(uniqueKey, credentials);

                    // Get a response from the service.
                    var serviceResponse = CallService(AuthenticationCache[uniqueKey], data);

                    if (null == serviceResponse)
                    {
                        // There was a problem with the stored credentials.
                        response = GetCredentialsResponse();
                    }
                    else
                    {
                        // Return the service response to the user.
                        response = serviceResponse;
                    }
                }
            }

            return response;
        }
    }}
```


## 使用 Managed 驗證程式庫來驗證使用者


如果您使用 Managed 程式庫來驗證識別權杖，則不需要計算唯一索引鍵。**AppIdentityToken** 類別上的 **UniqueUserIdentification** 屬性可直接作為使用者的唯一索引鍵。下列程式碼範例顯示在前一個範例中對 **GetResponseFromService** 方法的修改，您必須先進行這些修改，才能使用 **AppIdentityToken** 類別。


```js
        public JsonResponse GetResponseFromService(AppIdentityToken token, LocalCredentials credentials, object data)
        {
            JsonResponse response = null;
            // This method should never be called with a null token.
            if (null == token)
            {
                throw new ArgumentNullException("token");
            }

            if (null == credentials)
            {
                string uniqueKey = token.UniqueUserIdentitification;
                if (!AuthenticationCache.ContainsKey(uniqueKey))
                {
                    // The user's credentials are not in the authentication cache. Ask
                    // for the credentials.
                    response = GetCredentialsResponse();
                }
                else
                {
                    // User's credentials are in the cache. Make a request.
                    var serviceResponse = CallService(AuthenticationCache[uniqueKey], data);

                    if (null == serviceResponse)
                    {
                        // There was a problem with the stored credentials. For example,
                        // the user has ended their subscription to the service, or the
                        // credentials have expired. Get new credentials.
                        response = GetCredentialsResponse();
                    }
                    else
                    {
                        // The service returned a response to the user. Return the
                        // service response.
                        response = serviceResponse;
                    }
                }
            }
            else
            {
                // If the credentials are not null, it's a request to add an identity
                // to the authentication cache. Check to determine whether the local credentials
                // sent to the service are known.
                if (LocalCredentialsAreValid(credentials))
                {
                    // The local credentials are known. Add them to the 
                    // cached credentials. 
                    string uniqueKey = token.UniqueUserIdentitification;
                    AuthenticationCache.Add(uniqueKey, credentials);

                    // Get a response from the service.
                    var serviceResponse = CallService(AuthenticationCache[uniqueKey], data);

                    if (null == serviceResponse)
                    {
                        // There was a problem with the stored credentials.
                        response = GetCredentialsResponse();
                    }
                    else
                    {
                        // Return the service response to the user.
                        response = serviceResponse;
                    }
                }
            }

            return response;
        }
```


## 其他資源



- [使用 Exchange 識別權杖來驗證 Outlook 增益集](../outlook/authentication.md)
    
- [在 Exchange 中使用識別權杖以從 Outlook 增益集呼叫服務](../outlook/call-a-service-by-using-an-identity-token.md)
    
- [使用 Exchange 權杖驗證程式庫](../outlook/use-the-token-validation-library.md)
    
- [驗證 Exchange 識別權杖](../outlook/validate-an-identity-token.md)
    
