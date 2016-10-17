
# <a name="labs.components.iinputcomponent"></a>Labs.Components.IInputComponent

 _**適用於︰**Office 的應用程式 | Office 增益集 | Office Mix | PowerPoint_

啟用與輸入元件的互動。

```
interface IInputComponent extends Labs.Core.IComponent
```


## <a name="properties"></a>屬性


|名稱|描述|
|:-----|:-----|
| `maxScore: number`|輸入元件的最大允許分數。|
| `timeLimit: number`|輸入問題的時間限制。|
| `hasAnswer: boolean`|如果元件有答案則為 **True**。|
| `answer: any`|元件問題的答案 (若有)。|
| `secure: boolean`|如果輸入元件安全，則為 **True**。|
