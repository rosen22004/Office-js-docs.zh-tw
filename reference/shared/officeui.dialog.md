#UI.Dialog 物件
呼叫 [displayDialogAsync](officeui.displaydialogasync.md) 方法時傳回的物件。

## 成員
| 成員	       | 類型	   |描述|
|:---------------|:--------|:----------|
|關閉|函數|可讓增益集關閉其對話方塊。|
|addEventHandler|函數|註冊事件處理常式。 兩個支援的事件為︰ <ul><li>DialogMessageReceived。 當對話方塊傳送訊息至其父系時，就會觸發。</li><li>DialogEventReceived。 當對話方塊已關閉或卸載時，就會觸發。</li></ul> |


### close()
從父系頁面呼叫以關閉對應的對話方塊。     
```js    
[dialogObject].close();    
``` 

#### 參數    
無 

#### 傳回    
void  


#### 範例
如需範例，請參閱 [DisplayDialogAsync 方法](officeui.displaydialogasync.md)主題。
