
# <a name="labs.core.ilabcallback"></a>Labs.Core.ILabCallback

 _**適用於︰**Office 的應用程式 | Office 增益集 | Office Mix | PowerPoint_

處理 Labs.js 回呼方法的介面。

```
interface ILabCallback<T>
```


## <a name="callback-signature"></a>回呼簽章

 `(err: any, data: T): void`

 **回呼參數**


|||
|:-----|:-----|
| _err_|如果未發生錯誤則傳回 **Null**。如果發生錯誤則為非 **null**。|
| _data_|使用回呼傳回的資料。|
