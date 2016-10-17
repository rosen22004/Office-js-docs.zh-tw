
# <a name="labs.labinstance"></a>Labs.LabInstance

 _**適用於︰**Office 的應用程式 | Office 增益集 | Office Mix | PowerPoint_

為目前的使用者設定的實驗室執行個體。使用此物件以記錄和擷取使用者的實驗室資料。

```
class LabInstance
```


## <a name="variables"></a>變數


|||
|:-----|:-----|
| `public var data: any`|用來保留使用者資料的容器變數。|
| `public var components: Labs.ComponentInstanceBase[]`|組成實驗室執行個體的元件。|

## <a name="methods"></a>方法




### <a name="getstate"></a>getState

 `public function getState(callback: Labs.Core.ILabCallback<any>): void`

擷取指定使用者之實驗室的目前狀態。

 **參數**


|||
|:-----|:-----|
| _callback_|擷取實驗室狀態後時所引發的回呼函式。|

### <a name="setstate"></a>setState

 `public function setState(state: any, callback: Labs.Core.ILabCallback<void>): void`

針對指定的使用者設定實驗室的狀態。

 **參數**


|||
|:-----|:-----|
| _state_|要設定的狀態。|
| _callback_|設定狀態後所引發的回呼函式。|

### <a name="done"></a>完成

 `public function done(callback: Labs.Core.ILabCallback<void>): void`

表示使用者已完成採取實驗室的指標函式。

 **參數**


|||
|:-----|:-----|
| _callback_|完成實驗室後所引發的回呼函式。|
