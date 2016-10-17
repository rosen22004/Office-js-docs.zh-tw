
# <a name="labs.timeline"></a>Labs.Timeline

 _**適用於︰**Office 的應用程式 | Office 增益集 | Office Mix | PowerPoint_

提供 labs.js 時刻表功能的存取。

```
class Timeline
```


## <a name="methods"></a>方法




### <a name="method"></a>方法

 `function constructor(labsInternal: Labs.LabsInternal)`

建立 **Timeline** 類別的新執行個體。


### <a name="next"></a>下一頁

 `public function next(completionStatus: Labs.Core.ICompletionStatus, callback: Labs.Core.ILabCallback<void>): void`

表示時刻表應該前進到下一張投影片。

 **參數**


|||
|:-----|:-----|
| _completionStatus_|表示實驗室的目前狀態。|
| _callback_|實驗室移動到下一張投影片時所引發的回呼函式。|
