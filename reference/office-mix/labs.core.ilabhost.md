
# Labs.Core.ILabHost

 _**適用於︰**Office 相關應用程式 | Office 增益集 | Office Mix | PowerPoint_

提供將 Labs.js 連線到主機的抽象層。

```
interface ILabHost
```


## 方法


### getSupportedVersions

 `getSupportedVersions(): Core.ILabHostVersionInfo[]`

擷取實驗室主機所支援的版本。

 **參數**

無。


### 連線

 `connect(versions: Core.ILabHostVersionInfo[], callback: Core.ILabCallback<Core.IConnectionResponse>)`

初始化與主機的連線。

 **參數**


|||
|:-----|:-----|
| _版本_|用戶端可使用的主機版本清單。|
| _callback_|完成連線時引發的回呼函式。|

### 中斷連線

 `disconnect(callback: Core.ILabCallback<void>)`

終止與主機的通訊。

 **參數**


|||
|:-----|:-----|
| _completionStatus_|中斷連線時的實驗室狀態。|
| _callback_|中斷連線完成時所引發的回呼函式。|

### 開啟

 `on(handler: (string: any, any: any): void)`

新增事件處理常式以處理來自主機的郵件。解析的承諾會傳回主機。

 **參數**


|||
|:-----|:-----|
| _處理常式_|事件處理常式。|

### sendMessage

 `sendMessage(type: string, options: Core.IMessage, callback: Core.ILabCallback<Core.IMessageResponse>)`

傳送訊息給主機。

 **參數**


|||
|:-----|:-----|
| _類型_|正在傳送的郵件類型。|
| _options_|郵件選項。|
| _callback_|收到郵件後所引發的回呼函式。|

### 建立

 `create(options: Core.ILabCreationOptions, callback: Core.ILabCallback<void>)`

建立實驗室。儲存主機資訊，並設定空間來儲存組態和其他元素。

 **參數**


|||
|:-----|:-----|
| _options_|建立作業過程所傳遞的選項。|
| _callback_|建立實驗室後所引發的回呼函式。|

### getConfiguration

 `getConfiguration(callback: Core.ILabCallback<Core.IConfiguration>)`

從主機擷取目前的實驗室組態。

 **參數**


|||
|:-----|:-----|
| _callback_|用來擷取組態資訊的回呼函式。|

### setConfiguration

 `setConfiguration(configuration: Core.IConfiguration, callback: Core.ILabCallback<void>)`

在主機上設定新的實驗室組態。

 **參數**


|||
|:-----|:-----|
| _組態_|設定的實驗室組態。|
| _callback_|設定組態後所引發的回呼函式。|

### getConfigurationInstance

 `getConfigurationInstance(callback: Core.ILabCallback<Core.IConfigurationInstance>)`

擷取實驗室的執行個體組態。

 **參數**


|||
|:-----|:-----|
| _callback_|擷取組態執行個體後所引發的回呼函式。|

### getState

 `getState(callback: Core.ILabCallback<any>)`

擷取指定使用者之實驗室的目前狀態。

 **參數**


|||
|:-----|:-----|
| _completionStatus_|傳回目前的實驗室狀態的回呼函式。|

### setState

 `setState(state: any, callback: Core.ILabCallback<void>)`

針對指定的使用者設定實驗室的狀態。

 **參數**


|||
|:-----|:-----|
| _state_|實驗室狀態。|
| _callback_|設定狀態後所引發的回呼函式。|

### takeAction

 `takeAction(type: string, options: Core.IActionOptions, callback: Core.ILabCallback<Core.IAction>)`

嘗試動作。

 **參數**


|||
|:-----|:-----|
| _類型_|動作的類型。|
| _options_|與動作一起提供的選項。|
| _callback_|傳回最後一個執行動作的回呼函式。|

### takeAction

 `takeAction(type: string, options: Core.IActionOptions, result: Core.IActionResult, callback: Core.ILabCallback<Core.IAction>)`

已完成的執行動作。

 **參數**


|||
|:-----|:-----|
| _類型_|動作的類型。|
| _options_|與動作一起提供的選項。|
| _result_|動作的結果。|
| _callback_|傳回最後一個執行動作的回呼函式。|

### getActions

 `getActions(type: string, options: Core.IGetActionOptions, callback: Core.ILabCallback<Core.IAction[]>)`

嘗試動作。

 **參數**


|||
|:-----|:-----|
| _類型_|取得動作的類型。|
| _options_|與取得動作一起提供的選項。|
| _callback_|傳回已完成動作清單的回呼函式。|
