
# Labs.Components.ComponentAttempt

 _**適用於︰**Office 相關應用程式 | Office 增益集 | Office Mix | PowerPoint_

元件上嘗試的基底類別。

```
class ComponentAttempt
```


## 屬性


|**名稱**|**說明**|
|:-----|:-----|
| `public var _componentId: string`|指定元件的 ID。|
| `public var _id: string`|相關聯的實驗室 ID。|
| `public var _labs: Labs.LabsInternal`|用來與基礎 [Labs.Core.ILabHost](../../reference/office-mix/labs.core.ilabhost.md) 互動的實驗室 ([Labs.LabsInternal](http://msdn.microsoft.com/library/599fb2c4-bb16-4422-84ad-10ed85a14018.aspx)) 物件。|
| `public var _resumed: boolean`|如果實驗室已恢復指定嘗試的進度，則為 **True**。|
| `public var _state: Labs.ProblemState`|enum [Labs.ProblemState](../../reference/office-mix/labs.problemstate.md) 所提供之嘗試的目前狀態。|
| `public var _values: { [type:string]: Labs.ValueHolder<any>[]}`|[Labs.ValueHolder](../../reference/office-mix/labs.valueholder.md) 物件包含的嘗試相關值 (如果有的話)。|

## 方法




### 建構函式

 `(labs: Labs.LabsInternal, componentId: string, attemptId: string, values: {[type:string]: Labs.Core.IValueInstance[]})`

建立 ComponentAttempt 類別的新執行個體，並提供輸入參數值。

 **參數**


|**名稱**|**說明**|
|:-----|:-----|
| _labs_|用於使用嘗試的 [Labs.LabsInternal](http://msdn.microsoft.com/library/599fb2c4-bb16-4422-84ad-10ed85a14018.aspx) 執行個體。|
| _attemptId_|與嘗試關聯的 ID。|
| _values_|嘗試的關聯值陣列 ([Labs.Core.IValueInstance](../../reference/office-mix/labs.core.ivalueinstance.md))。|

### isResumed

 `public function isResumed(): boolean`

指出實驗室是否已恢復的布林函式。如果實驗室已恢復則為 **True**。

 **參數**

無。


### 繼續

 `public function resume(callback: Labs.Core.ILabCallback<void>): void`

指出實驗室是否已在指定的嘗試上恢復進度，並在此程序的過程中載入現有的資料。必須先恢復嘗試才可以使用。

 **參數**


|**名稱**|**說明**|
|:-----|:-----|
| _callback_|恢復嘗試後所引發的回呼函式。|

### getState

 `public function getState(): Labs.ProblemState`

擷取實驗室的狀態。

 **參數**

無。


### processAction

 `public function processAction(action: Labs.Core.IAction): void`

執行與嘗試相關的動作。

 **參數**

無。


### getValues

 `public function getValues(key: string): Labs.ValueHolder<any>[]`

擷取與嘗試關聯的值

 **參數**


|**名稱**|**說明**|
|:-----|:-----|
| _key_|與值對應中的值相關聯的機碼。|
