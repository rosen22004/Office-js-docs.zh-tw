
# 使用 Exchange Web 服務受管理 API 權杖驗證程式庫

您可以使用增益集從執行 Exchange Server 2013 或 Exchange Online 的伺服器要求的驗證權杖，找出 Outlook 增益集的用戶端。格式化成 JSON Web 權杖的權杖在 Exchange Server 上提供電子郵件帳戶的唯一識別項。Exchange Web Services (EWS) Managed API 提供協助程式類別以簡化驗證權杖的使用。

## 使用驗證程式庫的先決條件

若要驗證 Exchange 身分識別權杖，您必須安裝 [EWS 受管理 API 程式庫](https://www.nuget.org/packages/Microsoft.Exchange.WebServices)。

## 驗證 Exchange 識別權杖

EWS Managed API 驗證程式庫會提供 **AppIdentityToken** 類別來管理 Exchange 識別權杖。下列方法示範如何建立 **AppIdentityToken** 執行個體及呼叫 **Validate** 方法來驗證該權杖為有效。這個方法會採用下列參數：

- *rawToken*:從 [**Office.context.mailbox.getUserIdentityTokenAsync**](http://dev.office.com/reference/add-ins/outlook/Office.context.mailbox) 方法的 Outlook 增益集中傳回的權杖的字串表示。
- *hostUri*:稱為 **getUserIdentityTokenAsync** 的 Outlook 增益集中的頁面的完整 URI。

```C#
// Required to use the validation library.
using Microsoft.Exchange.WebServices.Auth.Validate;

private AppIdentityToken CreateAndValidateIdentityToken(string rawToken, string hostUri)
{
    try
    {
        AppIdentityToken token = (AppIdentityToken)AuthToken.Parse(rawToken);
        token.Validate(new Uri(hostUri));

        return token;
    }
    catch (TokenValidationException ex)
    {
        throw new ApplicationException("A client identity token validation error occurred.", ex);
    }
}
```

## 其他資源

- [使用 Exchange 識別權杖來驗證 Outlook 增益集](../outlook/authentication.md)  
- [Exchange 識別權杖的內容](../outlook/inside-the-identity-token.md)
- [驗證 Exchange 識別權杖](../outlook/validate-an-identity-token.md)
    
