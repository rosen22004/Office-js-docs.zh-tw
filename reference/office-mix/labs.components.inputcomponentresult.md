
# Labs.Components.InputComponentResult

 _**適用於︰**Office 的應用程式 | Office 增益集 | Office Mix | PowerPoint_

輸入元件提交的結果。

```
class InputComponentResult
```


## 屬性


|屬性	|說明|
|:-----|:-----|
| `public var score: any`|提交的相關分數。|
| `public var complete: boolean`|表示提交的結果是否導致嘗試完成。如果嘗試完成則為 **True**。|

## 方法




### 建構函式

 `function constructor(score: any, complete: boolean)`

建立 **InputComponentResult** 類別的新執行個體。

 **參數**


|參數|說明|
|:-----|:-----|
| _score_|結果的關聯分數。|
| _完成_|如果結果完成嘗試，則為布林值 **true**。|
