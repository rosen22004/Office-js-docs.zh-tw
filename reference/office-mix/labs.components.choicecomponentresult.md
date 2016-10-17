
# <a name="labs.components.choicecomponentresult"></a>Labs.Components.ChoiceComponentResult

 _**適用於︰**Office 的應用程式 | Office 增益集 | Office Mix | PowerPoint_

選擇元件提交的結果。

```
class ChoiceComponentResult
```


## <a name="properties"></a>屬性


|屬性	|描述|
|:-----|:-----|
| `public var score: any`|提交的相關分數。|
| `public var complete: boolean`|結果是否完成嘗試。如果結果完成嘗試，則為 **True**。|

## <a name="methods"></a>方法




### <a name="constructor"></a>建構函式

 `function constructor(score: any, complete: boolean)`

建立 **ChoiceComponentResult** 類別的新執行個體。

 **參數**


|參數|描述|
|:-----|:-----|
| _score_|結果的分數。|
| _complete_|指出結果是否完成嘗試。|
