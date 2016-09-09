 

# Office

Office 命名空間會提供共用的介面，可為所有 Office 應用程式中的增益集所使用。此清單會列出這些由 Outlook 增益集所使用的介面。如需 Office 命名空間的完整清單，請參閱 [共用 API](../../shared/shared-api.md)。

##### 需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|適用的 Outlook 模式| 撰寫或讀取|

### 命名空間

[內容](Office.context.md):提供來自 Office 增益集 API 內容的命名空間，用於 Outlook 增益集 API 的共用介面。

[MailboxEnums](Office.MailboxEnums.md):包括 ItemType、EntityType、AttachmentType、RecipientType、ResponseType 和 ItemNotificationMessageType 的列舉。

### 成員

####  AsyncResultStatus︰字串

指定非同步呼叫的結果。

##### 類型：

*   字串

##### 屬性：

|名稱| 類型	| 描述|
|---|---|---|
|`Succeeded`| String|呼叫成功。|
|`Failed`| String|呼叫失敗。|

##### 需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|適用的 Outlook 模式| 撰寫或讀取|
####  CoercionType︰字串

指定如何強制轉型所傳回或由叫用方法設定的資料。

##### 類型：

*   字串

##### 屬性：

|名稱| 類型	| 描述|
|---|---|---|
|`Html`| 字串|要求以 HTML 格式傳回資料。|
|`Text`| 字串|要求以文字格式傳回資料。|

##### 需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|適用的 Outlook 模式| 撰寫或讀取|
####  SourceProperty︰字串

指定由叫用方法所傳回的資料來源。

##### 類型：

*   字串

##### 屬性：

|名稱| 類型	| 描述|
|---|---|---|
|`Body`| 字串|資料來源是來自郵件本文。|
|`Subject`| String|資料來源是來自郵件主旨。|

##### 需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|適用的 Outlook 模式| 撰寫或讀取|
