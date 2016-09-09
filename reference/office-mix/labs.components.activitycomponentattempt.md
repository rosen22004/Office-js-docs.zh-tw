
# Labs.Components.ActivityComponentAttempt

 _**適用於︰**Office 相關應用程式 | Office 增益集 | Office Mix | PowerPoint_

表示完成活動元件時的嘗試。

```
class Permissions
```


## 方法




### 建構函式

 `function constructor(labs: Labs.LabsInternal, componentId: string, attemptId: string, values: {[type:string]: Labs.Core.IValueInstance[]})`

建立 **ActivityComponentAttempt** 類別的新執行個體。

 **參數**


|**名稱**|**說明**|
|:-----|:-----|
| _labs_|元件所關聯的實驗室執行個體 ([Labs.LabsInternal](http://msdn.microsoft.com/library/599fb2c4-bb16-4422-84ad-10ed85a14018.aspx))。|
| _componentId_|嘗試所關聯之元件的 ID。|
| _attemptId_|嘗試的 ID。|
| _values_|與元件產生關聯的值 (如果有的話)。|

### 完成

 `public function complete(callback: Labs.Core.ILabCallback<void>): void`

活動已完成的指標。

 **參數**


|**名稱**|**說明**|
|:-----|:-----|
| _callback_|完成活動後所叫用的回呼函式。|

### processAction

 `public function processAction(action: Labs.Core.IAction): void`

透過指定的嘗試擷取動作所執行的函式，然後填入實驗室的狀態。

 **參數**


|**名稱**|**說明**|
|:-----|:-----|
| _action_|動作執行個體 ([Labs.Core.IAction](../../reference/office-mix/labs.core.iaction.md))。|
