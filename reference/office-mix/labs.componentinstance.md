
# <a name="labs.componentinstance"></a>Labs.ComponentInstance

 _**適用於︰**Office 的應用程式 | Office 增益集 | Office Mix | PowerPoint_

代表元件的執行個體，為使用者在執行階段的指定元件具現化。物件包含實驗室的特定執行之元件的轉譯檢視。

```
class ComponentInstance<T> extends Labs.ComponentInstanceBase
```


## <a name="properties"></a>屬性

無。


## <a name="methods"></a>方法




### <a name="constructor"></a>建構函式

 `function constructor()`

初始化 **ComponentInstance** 類別的新執行個體。


### <a name="createattempt"></a>createAttempt

 `public function createAttempt(callback: Labs.Core.ILabCallback<T>): void`

在元件內容中建立新嘗試。

 **參數**


|**名稱**|**描述**|
|:-----|:-----|
| _callback_|建立嘗試後所引發的回呼。|

### <a name="getattempts"></a>getAttempts

 `public function getAttempts(callback: Labs.Core.ILabCallback<T[]>): void`

擷取與指定元件相關聯的所有嘗試。

 **參數**


|**名稱**|**描述**|
|:-----|:-----|
| _callback_|擷取嘗試後所引發的回呼。|

### <a name="getcreateattemptoptions"></a>getCreateAttemptOptions

 `public function getCreateAttemptOptions(): Labs.Core.Actions.ICreateAttemptOptions`

擷取預設的建立嘗試選項。可由衍生類別覆寫。


### <a name="buildattempt"></a>buildAttempt

 `public function buildAttempt(createAttemptResult: Labs.Core.IAction): T`

從指定的動作建置嘗試。應該由衍生類別實作。

 **參數**


|**名稱**|**描述**|
|:-----|:-----|
| _createAttemptResult_|指定嘗試的建立嘗試動作。|
