
# Labs.LabEditor

 _**適用於︰**Office 相關應用程式 | Office 增益集 | Office Mix | PowerPoint_

**LabEditor** 物件可讓您編輯指定的實驗室，以及取得和設定實驗室所關聯的組態資料。

```
class LabEditor
```


## 方法


### getConfiguration

 `public function getConfiguration(callback: Labs.Core.ILabCallback<Labs.Core.IConfiguration>): void`

擷取目前的實驗室組態。

 **參數**


|**名稱**|**說明**|
|:-----|:-----|
| _callback_|擷取組態後所引發的回呼函式。|

### setConfiguration

 `public function getConfiguration(callback: Labs.Core.ILabCallback<Labs.Core.IConfiguration>): void`

設定新的實驗室組態。

 **參數**


|**名稱**|**說明**|
|:-----|:-----|
| _組態_|要設定的組態。|
| _callback_|設定組態後所引發的回呼函式。|

### 完成

 `public function done(callback: Labs.Core.ILabCallback<void>): void`

表示使用者已完成編輯實驗室。

 **參數**


|**名稱**|**說明**|
|:-----|:-----|
| _callback_|完成實驗室編輯器後所引發的回呼函式。|
