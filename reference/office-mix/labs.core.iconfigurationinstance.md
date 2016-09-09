
# Labs.Core.IConfigurationInstance

 _**適用於︰**Office 的應用程式 | Office 增益集 | Office Mix | PowerPoint_

實驗室組態執行個體的基底類別。執行個體是指定使用者之組態的具現化，且包含實驗室的特定執行之組態的轉譯檢視。此檢視可能排除隱藏的資訊 (例如，提示和答案)，並也包含用來識別不同執行個體的 ID。

```
interface IConfigurationInstance extends Core.IUserData
```


## 屬性


|||
|:-----|:-----|
| `appVersion: Core.IVersion`|與此組態相關聯的實驗室版本。|
| `components: Core.IComponentInstance[]`|實驗室的關聯元件。|
| `name: string`|實驗室的名稱。|
| `timeline: Core.ITimelineConfiguration`|實驗室的時間表組態。|
