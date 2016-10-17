
# <a name="labs.components.inputcomponentsubmission"></a>Labs.Components.InputComponentSubmission

 _**適用於︰**Office 的應用程式 | Office 增益集 | Office Mix | PowerPoint_

表示提交到輸入元件。

```
class InputComponentSubmission
```


## <a name="properties"></a>屬性


|屬性	|描述|
|:-----|:-----|
| `public var answer: Components.InputComponentAnswer`|提交的關聯答案 ([Labs.Components.InputComponentAnswer](../../reference/office-mix/labs.components.inputcomponentanswer.md))。|
| `public var result: Components.InputComponentResult`|提交的關聯結果 ([Labs.Components.InputComponentResult](../../reference/office-mix/labs.components.inputcomponentresult.md))。|
| `public var time: number`|收到提交的時間。|

## <a name="methods"></a>方法




### <a name="constructor"></a>建構函式

 `function constructor(answer: Components.InputComponentAnswer, result: Components.InputComponentResult, time: number)`

建立 **InputComponentSubmission** 類別的新執行個體。

 **參數**


|參數|描述|
|:-----|:-----|
| _answer_|提交的相關解答。|
| _result_|提交的結果。|
| _time_|收到提交的時間。|
