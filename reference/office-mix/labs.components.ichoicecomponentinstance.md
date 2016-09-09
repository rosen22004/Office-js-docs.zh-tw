
# Labs.Components.IChoiceComponentInstance

 _**適用於︰**Office 的應用程式 | Office 增益集 | Office Mix | PowerPoint_

選擇元件的執行個體。

```
interface IChoiceComponentInstance extends Labs.Core.IComponentInstance
```


## 屬性


|名稱|說明|
|:-----|:-----|
| `choices: Components.IChoice[]`|表示與問題與相關聯之選項清單的陣列。|
| `timeLimit: number`|完成問題的時間限制。|
| `maxAttempts: number`|問題所允許的最多嘗試次數。|
| `maxScore: number`|問題的最高分數。|
| `hasAnswer: boolean`|如果問題有答案則為 **True**。|
| `answer: any`|問題的答案。如果支援多個答案，則為一個陣列；如果只支援一個答案，則為單一 ID。|
| `secure: boolean`|測驗是否安全，表示從使用者抑制安全欄位。|
