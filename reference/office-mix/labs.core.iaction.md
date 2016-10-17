
# <a name="labs.core.iaction"></a>Labs.Core.IAction

 _**適用於︰**Office 的應用程式 | Office 增益集 | Office Mix | PowerPoint_

代表實驗室動作，為使用者與指定實驗室的互動。

```
interface IAction
```


## <a name="properties"></a>屬性


|||
|:-----|:-----|
| `type: string`|使用者所採取的動作類型。|
| `options: Core.IActionOptions`|使用者所採取的動作所傳送的 [Labs.Core.IActionOptions](../../reference/office-mix/labs.core.iactionoptions.md) 選項。|
| `result: Core.IActionResult`|動作的 [Labs.Core.IActionResult](../../reference/office-mix/labs.core.iactionresult.md) 結果。|
| `time: number`|動作的完成時間 (以毫秒為單位表示)，即 1970 年 1 月 1 日 00:00:00 UTC 以來經過的時間。|
