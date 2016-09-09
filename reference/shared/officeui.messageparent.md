# UI.messageParent 方法

從對話方塊將訊息傳遞至其父系/opener 頁面。 呼叫此 API 的頁面必須位於與父系頁面相同的網域。 

## 語法

```js
Office.context.ui.messageParent("Message from Dialog box");
```

## 參數
| 參數	    | 類型	   |說明|
|:---------------|:--------|:----------|
|messageObject|字串或布林值|從對話方塊接受訊息，以傳遞至增益集。|

## 傳回
void

## 範例
如需範例，請參閱 [DisplayDialogAsync 方法](officeui.displaydialogasync.md)主題。

