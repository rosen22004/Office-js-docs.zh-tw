
# <a name="labs.components.inputcomponentattempt"></a>Labs.Components.InputComponentAttempt

 _**適用於︰**Office 的應用程式 | Office 增益集 | Office Mix | PowerPoint_

表示嘗試與輸入元件互動。

```
class InputComponentAttempt extends Components.ComponentAttempt
```


## <a name="methods"></a>方法




### <a name="constructor"></a>建構函式

 `function constructor(labs: Labs.LabsInternal, componentId: string, attemptId: string, values: {[type:string]: Labs.Core.IValueInstance[]})`

建立 **InputComponentAttempt** 類別的新執行個體。

 **參數**


|參數|描述|
|:-----|:-----|
| _labs_|嘗試所關聯的實驗室 ([Labs.LabsInternal](http://msdn.microsoft.com/library/599fb2c4-bb16-4422-84ad-10ed85a14018.aspx))。|
| _componentID_|嘗試所關聯之元件的 ID。|
| _attemptId_|特定嘗試的 ID。|
| _values_|包含值執行個體的陣列 ([Labs.Core.IValueInstance](../../reference/office-mix/labs.core.ivalueinstance.md))。|

### <a name="processaction"></a>processAction

 `public function processAction(action: Labs.Core.IAction): void`

逐一查看指定嘗試的擷取動作，並填入實驗室的狀態。

 **參數**


|參數|描述|
|:-----|:-----|
| _action_|實驗室狀態的關聯動作。|

### <a name="getsubmissions"></a>getSubmissions

 `public function getSubmissions(): Components.InputComponentSubmission[]`

擷取先前已針對指定嘗試提交的所有提交。


### <a name="submit"></a>提交

 `public function submit(answer: Components.InputComponentAnswer, result: Components.InputComponentResult, callback: Labs.Core.ILabCallback<Components.InputComponentSubmission>): void`

提交實驗室分級的新答案，且不會使用主機來計算成績。

 **參數**


|參數|描述|
|:-----|:-----|
| _answer_|嘗試的相關解答。|
| _result_|提交的相關結果。|
| _callback_|接收提交後引發的回呼函式。|
