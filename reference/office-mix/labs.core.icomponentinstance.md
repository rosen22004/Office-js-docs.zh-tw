
# <a name="labs.core.icomponentinstance"></a>Labs.Core.IComponentInstance

 _**適用於︰**Office 的應用程式 | Office 增益集 | Office Mix | PowerPoint_

實驗室元件之執行個體的基底類別。

```
interface IComponentInstance extends Core.ILabObject, Core.IUserData
```


## <a name="properties"></a>屬性


|||
|:-----|:-----|
| `componentId: string`|這個執行個體相關聯的元件的 ID。|
| `name: string`|元件的名稱。|
| `values: {[type:string]: Core.IValueInstance[]}`|與元件相關聯的值屬性對應。|

## <a name="remarks"></a>備註

元件執行個體是使用者之元件的具現化。它包含實驗室的特定執行之元件的已轉換檢視。此檢視可能排除隱藏的資訊 (答案、提示等等)，並也包含用來識別不同執行個體的 ID。

