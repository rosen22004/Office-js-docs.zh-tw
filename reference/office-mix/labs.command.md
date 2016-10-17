
# <a name="labs.command"></a>Labs.Command

 _**適用於︰**Office 的應用程式 | Office 增益集 | Office Mix | PowerPoint_

用來在用戶端與主機之間傳遞訊息的一般命令。

```
class Command
```


## <a name="properties"></a>屬性


|**名稱**|**描述**|
|:-----|:-----|
| `public var type: string`|命令的類型。|
| `public var commandData: any`|與命令關聯的選擇性資料。|

## <a name="methods"></a>方法




### <a name="constructor"></a>建構函式

 `function constructor(type: string, commandData?: any)`

描述

 **參數**


|||
|:-----|:-----|
| `type`|命令的類型。|
| `commandData`|與命令關聯的選擇性資料。|
