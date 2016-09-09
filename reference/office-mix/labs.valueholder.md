
# Labs.ValueHolder

 _**適用於︰**Office 相關應用程式 | Office 增益集 | Office Mix | PowerPoint_

保留和追蹤指定的實驗室值的容器物件。值可儲存在本機或伺服器上。

```
class ValueHolder<T>
```


## 變數


|||
|:-----|:-----|
| `public var isHint: boolean`|如果值是提示則為 **True**。|
| `public var hasBeenRequested: boolean`|如果值已被實驗室要求，則為 **True**。|
| `public var hasValue: boolean`|如果值容器目前具有所要值則為 **True**。|
| `public var value: T`|容器中保留的值。|
| `public var id: string`|值的 ID。|

## 方法




### getValue

 `public function getValue(callback: Labs.Core.ILabCallback<T>): void`

擷取指定的值。

 **參數**


|||
|:-----|:-----|
| _callback_|傳回指定值的回呼函式。|

### provideValue

 `public function provideValue(value: T): void`

將值提供給值容器的內部方法。

 **參數**


|||
|:-----|:-----|
| _value_|要提供給值容器的值。|
