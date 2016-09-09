
# Labs.Components.IDynamicComponentInstance

 _**適用於︰**Office 的應用程式 | Office 增益集 | Office Mix | PowerPoint_

動態元件的執行個體。

```
interface IDynamicComponentInstance extends Labs.Core.IComponentInstance
```


## 屬性


|名稱|說明|
|:-----|:-----|
| `generatedComponentTypes: string[]`|可能會產生包含此動態元件之元件類型的陣列。|
| `maxComponents: number`|此動態元件將產生的元件最大數目。或如果沒有任何端點，則為 **Labs.Components.Infinite**。|
