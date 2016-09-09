
# Labs.Core.IValueInstance

 _**適用於︰**Office 的應用程式 | Office 增益集 | Office Mix | PowerPoint_

包含數值資料的 [Labs.Core.IValue](../../reference/office-mix/labs.core.ivalue.md) 物件執行個體 (若有)。

```
interface IValueInstance
```


## 屬性


|||
|:-----|:-----|
| `valueId: string`|此執行個體代表之值的 ID。|
| `isHint: boolean`|若將此值視為提示，則為布林值 **true**。|
| `hasValue: boolean`|如果執行個體資訊包含值，則為布林值 **true**。|
| `value?: any`|值。視其是否已隱藏，可能會設定這個參數。|
