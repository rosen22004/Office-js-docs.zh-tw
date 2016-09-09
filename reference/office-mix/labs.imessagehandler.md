
# Labs.IMessageHandler

 _**適用於︰**Office 的應用程式 | Office 增益集 | Office Mix | PowerPoint_

可讓您定義事件處理常式的介面。

```
interface IMessageHandler(origin: Window, data: any, callback: Labs.Core.ILabCallback<any>): void
```


## 

 **參數**


|||
|:-----|:-----|
| `origin`|發出郵件的實驗室視窗。|
| `data`|郵件的內容。|
| `callback`|收到郵件後所引發的回呼函式。|
