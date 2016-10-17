
# <a name="labs.registerdeserializer"></a>Labs.registerDeserializer

 _**適用於︰**Office 的應用程式 | Office 增益集 | Office Mix | PowerPoint_

將指定的 JSON 物件還原序列化到物件。只應由元件作者使用。

```
function registerDeserializer(type: string, deserialize: (json: Core.ILabObject): any): void
```


## <a name="parameters"></a>參數


|**名稱**|**描述**|
|:-----|:-----|
|json|要還原序列化的 [Labs.Core.ILabObject](../../reference/office-mix/labs.core.ilabobject.md)。|

## <a name="return-value"></a>傳回值

傳回 [Labs.Core.ILabObject](../../reference/office-mix/labs.core.ilabobject.md) 執行個體。

