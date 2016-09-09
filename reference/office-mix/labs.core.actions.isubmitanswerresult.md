
# Labs.Core.Actions.ISubmitAnswerResult

 _**適用於︰**Office 的應用程式 | Office 增益集 | Office Mix | PowerPoint_

提交嘗試答案的結果。

```
interface ISubmitAnswerResult extends Core.IActionResult
```


## 屬性


|||
|:-----|:-----|
| `submissionId: string`|提交相關聯的 ID。由伺服器提供。|
| `complete: boolean`|如果嘗試因目前的提交而完成，則傳回 **true**。|
| `score: any`|與提交相關聯的分數資訊。|
