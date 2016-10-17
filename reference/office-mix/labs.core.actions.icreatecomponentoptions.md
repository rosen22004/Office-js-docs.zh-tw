
# <a name="labs.core.actions.icreatecomponentoptions"></a>Labs.Core.Actions.ICreateComponentOptions

 _**適用於︰**Office 的應用程式 | Office 增益集 | Office Mix | PowerPoint_

建立新元件。

```
interface ICreateComponentOptions extends Core.IActionOptions
```


## <a name="properties"></a>屬性


|||
|:-----|:-----|
| `componentId: string`|叫用建立元件動作的元件。|
| `component: Core.IComponent`|要建立的 [Labs.Core.IComponent](../../reference/office-mix/labs.core.icomponent.md) 元件|
| `correlationId?: string`|用來跨實驗室的所有執行個體關聯此元件的選用欄位。可讓主機在相同的元件上識別不同的嘗試。|
