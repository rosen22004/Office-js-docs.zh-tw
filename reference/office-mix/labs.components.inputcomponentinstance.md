
# Labs.Components.InputComponentInstance

 _**適用於︰**Office 的應用程式 | Office 增益集 | Office Mix | PowerPoint_

代表輸入元件的執行個體。

```
class InputComponentInstance extends Labs.ComponentInstance<Components.InputComponentAttempt>
```


## 屬性


|屬性	|說明|
|:-----|:-----|
| `public var component: Components.IInputComponentInstance`|這個類別所表示的基礎 [Labs.Components.IInputComponentInstance](../../reference/office-mix/labs.components.iinputcomponentinstance.md) 物件。|

## 方法




### 建構函式

 `function constructor(component: Components.IInputComponentInstance)`

建立一個新 [Labs.Components.IInputComponentInstance](../../reference/office-mix/labs.components.iinputcomponentinstance.md) 執行個體。

 **參數**


|參數|說明|
|:-----|:-----|
| _component_|從中建立此類別的 [Labs.Components.IInputComponentInstance](../../reference/office-mix/labs.components.iinputcomponentinstance.md)。|

### buildAttempt

 `public function buildAttempt(createAttemptAction: Labs.Core.IAction): Components.InputComponentAttempt`

建立新的 [Labs.Components.InputComponentAttempt](../../reference/office-mix/labs.components.inputcomponentattempt.md)。實作基底類別上定義的抽象方法。

 **參數**


|參數|說明|
|:-----|:-----|
| _createAttemptResult_|建立嘗試動作的結果。|
