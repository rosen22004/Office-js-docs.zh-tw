
# <a name="labs.connect-(overload)"></a>Labs.connect (overload)

 _**適用於︰**Office 的應用程式 | Office 增益集 | Office Mix | PowerPoint_

初始化與主機的連線。

```
function connect(labHost: Core.ILabHost, callback: Core.ILabCallback<Core.IConnectionResponse>)
```


## <a name="parameters"></a>參數


|||
|:-----|:-----|
| _labHost_|選用。連接到的 [Labs.Core.ILabHost](../../reference/office-mix/labs.core.ilabhost.md) 執行個體。如果未指定主機，將使用 [Labs.DefaultHostBuilder](../../reference/office-mix/labs.defaulthostbuilder.md) 建構一個主機。|
| _callback_|建立連後後引發回撥。|

## <a name="return-value"></a>傳回值

傳回至主機的連線。

