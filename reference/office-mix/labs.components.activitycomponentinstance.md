
# Labs.Components.ActivityComponentInstance

 _**適用於︰**Office 的應用程式 | Office 增益集 | Office Mix | PowerPoint_

表示活動元件的目前執行個體。

```
class ActivityComponentInstance extends Labs.ComponentInstance<Components.ActivityComponentAttempt>
```


## 屬性


|**名稱**|**說明**|
|:-----|:-----|
| `public var component: Components.IActivityComponentInstance`|這個類別表示的基礎 [Labs.Components.IActivityComponentInstance](../../reference/office-mix/labs.components.iactivitycomponentinstance.md)|

## 方法




### 建構函式

 `function constructor(component: Components.IActivityComponentInstance)`

建立 [Labs.Components.IActivityComponentInstance](../../reference/office-mix/labs.components.iactivitycomponentinstance.md) 類別的新執行個體。

 **參數**


|**名稱**|**說明**|
|:-----|:-----|
| _component_|**IActivityComponentInstance**，用來從這個類別建立這個類別。|

### buildAttempt

 `public function buildAttempt(createAttemptAction: Labs.Core.IAction): Components.ActivityComponentAttempt`

建立新的 **ActivityComponentAttempt** 執行個體，並實作基底類別上定義的抽象方法

 **參數**


|**名稱**|**說明**|
|:-----|:-----|
| _createAttemptResult_|建立嘗試動作的結果。|
