
# Labs.Core.IConfiguration

 _**適用於︰**Office 的應用程式 | Office 增益集 | Office Mix | PowerPoint_

實驗室組態資料結構。

```
interface IConfiguration extends Core.IUserData
```


## 屬性


|||
|:-----|:-----|
| `appVersion: Core.IVersion`|與此組態相關聯的應用程式版本。|
| `components: Core.IComponent[]`|實驗室隨附的元件。|
| `name: string`|實驗室的名稱。|
| `timeline: Core.ITimelineConfiguration`|實驗室的時間表組態。|
| `analytics: Core.IAnalyticsConfiguration`|實驗室的分析組態。|
