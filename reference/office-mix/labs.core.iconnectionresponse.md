
# <a name="labs.core.iconnectionresponse"></a>Labs.Core.IConnectionResponse

 _**適用於︰**Office 的應用程式 | Office 增益集 | Office Mix | PowerPoint_

從連線呼叫傳回回應資訊。

```
interface IConnectionResponse
```


## <a name="properties"></a>屬性


|||
|:-----|:-----|
| `initializationInfo: Core.IConfigurationInfo`|初始化組態資訊，如果應用程式尚未初始化則為 **null**。|
| `mode: Core.LabMode`|實驗室目前正在執行的模式。|
| `hostVersion: Core.IVersion`|伺服器的版本資訊 ([Labs.Core.IVersion](../../reference/office-mix/labs.core.iversion.md))。|
| `userInfo: Core.IUserInfo`|使用者的相關資訊 ([Labs.Core.IUserInfo](../../reference/office-mix/labs.core.iuserinfo.md))。|
