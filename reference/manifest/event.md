# <a name="event-element"></a>事件元素
定義某個增益集中的事件處理常式。

> **附註：**僅 Office 365 中的 Outlook 網頁版支援 `Event` 元素。

## <a name="attributes"></a>屬性

|  屬性  |  必要  |  描述  |
|:-----|:-----|:-----|
|  [類型](#type-attribute)  |  是  | 指定要處理的事件。 |
|  [FunctionExecution](#functionexecution-attribute)  |  是  | 指定事件處理常式的執行樣式，非同步或同步執行樣式。目前只支援同步事件處理常式。 |
|  [FunctionName](#functionname-attribute)  |  是  | 指定事件處理常式函數名稱。 |

## <a name="type-attribute"></a>類型屬性
必要。指定哪一個事件會叫用事件處理常式。這個屬性的可能值詳列於下表中。

|  事件類型  |  描述  |
|:-----|:-----|
|  `ItemSend`  |  使用者傳送郵件或會議邀請時，便會叫用事件處理常式。  |

## <a name="functionexecution-attribute"></a>FunctionExecution 屬性
必要。必須設定為 `synchronous`。

## <a name="functionname-attribute"></a>FunctionName 屬性
必要。指定事件處理常式的函數名稱。此值必須符合增益集的 [函數檔案](./functionfile.md) 中的函數名稱。

```xml
<Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" /> 
```