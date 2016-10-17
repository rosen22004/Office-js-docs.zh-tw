
# <a name="labs.core.icomponent"></a>Labs.Core.IComponent

 _**適用於︰**Office 的應用程式 | Office 增益集 | Office Mix | PowerPoint_

代表實驗室元件的基底類別。

```
interface IComponent extends Core.ILabObject, Core.IUserData
```


## <a name="properties"></a>屬性


|||
|:-----|:-----|
| `name: string`|元件的名稱。|
| `values: {[type:string]: Core.IValue[]}`|與元件相關聯的值屬性對應。|
