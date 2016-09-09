
# Office.cast.item 屬性
提供撰寫或讀取模式訊息和約會專用的 IntelliSense。

|||
|:-----|:-----|
|**主機︰**|Outlook|
|**可用於[需求集合](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|信箱|
|**上次變更於**|1.0|



|||
|:-----|:-----|
|**適用的 Outlook 模式**|僅限在 Visual Studio 中設計時使用|

```js
Office.cast.item.toAppointmentCompose(Office.context.mailbox.item);
```

```js
Office.cast.item.toAppointmentRead(Office.context.mailbox.item);
```

```js
Office.cast.item.toAppointment(Office.context.mailbox.item);
```

```js
Office.cast.item.toItemCompose(Office.context.mailbox.item);
```

```js
Office.cast.item.toItemRead(Office.context.mailbox.item);
```

```js
Office.cast.item.toMessageCompose(Office.context.mailbox.item);
```

```js
Office.cast.item.toMessageRead(Office.context.mailbox.item);
```

```js
Office.cast.item.toMessage(Office.context.mailbox.item);
```


## 傳回值

一組方法，可讓您為 Outlook 增益集選取適當的 IntelliSense。


## 備註

這個屬性及其方法僅支援 IntelliSense 在 Visual Studio 上開發 Outlook 增益集。對於其他開發工具沒有任何會效用。

在 Visual Studio 中設計時，可使用 **Office.cast.item** 方法，以針對 **Office.context.mailbox.item** 屬性提供特定 IntelliSense。例如當您使用 **toAppointmentCompose** 方法時，IntelliSense 只會顯示適用於撰寫模式的 **Appointment** 方法和屬性。

在執行階段，**Office.cast.item** 方法對 Outlook 增益集沒有任何效用。


## 範例

下列範例會使用  **toMessageCompose** 方法來轉換 **Office.context.mailbox.item** 屬性，因此它只會針對撰寫模式的 **Message** 物件顯示 IntelliSense。轉換之後，`message` 變數只會針對可以在撰寫模式中使用的方法和屬性顯示 IntelliSense。


```js
var message = Office.cast.item.toMessageCompose(Office.context.mailbox.item);

```


## 支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此方法。空白儲存格表示 Office 主應用程式不支援此方法。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。

||Office for Windows desktop|Office Online (在瀏覽器中)|Mac 版 Outlook|
|:-----|:-----|:-----|:-----|
|**Outlook**|Y|Y|Y|

|||
|:-----|:-----|
|**可用於需求集合**|信箱|
|**最低權限等級**|[限制](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**增益集類型**|Outlook|
|**文件庫**|Office.js|
|**命名空間**|Office|

## 支援歷程記錄



|**版本**|**變更**|
|:-----|:-----|
|1.0|已導入|
