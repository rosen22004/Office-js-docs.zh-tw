
# Labs.Components.ChoiceComponentAttempt

 _**適用於︰**Office 相關應用程式 | Office 增益集 | Office Mix | PowerPoint_

表示選擇元件上的嘗試。

```
class ChoiceComponentAttempt extends Components.ComponentAttempt
```


## 方法




### 建構函式

 `function constructor(labs: Labs.LabsInternal, componentId: string, attemptId: string, values: {[type:string]: Labs.Core.IValueInstance[]})`

建立 **ChoiceComponentAttempt** 類別的新執行個體。

 **參數**


|**名稱**|**說明**|
|:-----|:-----|
| _labs_|用於使用嘗試的 [Labs.LabsInternal](http://msdn.microsoft.com/library/599fb2c4-bb16-4422-84ad-10ed85a14018.aspx) 執行個體。|
| _attemptId_|與嘗試關聯的 ID。|
| _values_|與嘗試關聯的值。|

### 逾時

 `public function timeout(callback: Labs.Core.ILabCallback<void>): void`

表示實驗室已逾時。

 **參數**


|**名稱**|**說明**|
|:-----|:-----|
| _callback_|伺服器收到逾時訊息後所引發的回呼函式。|

### getSubmissions

 `public function getSubmissions(): Components.ChoiceComponentSubmission[]`

擷取先前已為指定嘗試所提交的所有提交。


### 提交

 `public function submit(answer: Components.ChoiceComponentAnswer, result: Components.ChoiceComponentResult, callback: Labs.Core.ILabCallback<Components.ChoiceComponentSubmission>): void`

提交實驗室分級的新答案，且不會使用主機來計算成績。

 **參數**


|**名稱**|**說明**|
|:-----|:-----|
| _answer_|嘗試的答案。|
| _result_|提交的結果。|
| _callback_|接收提交後引發的回呼函式。|

### processAction

 `public function processAction(action: Labs.Core.IAction): void`

啟始處理 [Labs.Core.IAction](../../reference/office-mix/labs.core.iaction.md) 動作。

