
# <a name="labs.components.choicecomponentinstance"></a>Labs.Components.ChoiceComponentInstance

 _**適用於︰**Office 的應用程式 | Office 增益集 | Office Mix | PowerPoint_

表示選擇元件上的執行個體。

```
class ChoiceComponentInstance extends Labs.ComponentInstance<Components.ChoiceComponentAttempt>
```


## <a name="properties"></a>屬性


|屬性	|描述|
|:-----|:-----|
| `public var component: Components.IChoiceComponentInstance`|這個類別表示的基礎 [Labs.Components.IChoiceComponentInstance](../../reference/office-mix/labs.components.ichoicecomponentinstance.md)|

## <a name="methods"></a>方法




### <a name="constructor"></a>建構函式

 `function constructor(component: Components.IChoiceComponentInstance)`

建立 **ChoiceComponentInstance** 類別的新執行個體。

 **參數**


|參數|描述|
|:-----|:-----|
| _component_|從中建立此類別的 [Labs.Components.IChoiceComponentInstance](../../reference/office-mix/labs.components.ichoicecomponentinstance.md) 物件。|

### <a name="buildattempt"></a>buildAttempt

 `public function buildAttempt(createAttemptAction: Labs.Core.IAction): Components.ChoiceComponentAttempt`

建立新的 **ChoiceComponentAttempt** 執行個體，並實作基底類別上定義的抽象方法。

 **參數**


|參數|描述|
|:-----|:-----|
| _createAttemptResult_|來自建立嘗試動作的結果。|
