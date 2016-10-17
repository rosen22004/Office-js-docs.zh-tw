

# <a name="event"></a>事件

`event` 物件會作為參數傳遞給由無 UI 的指令按鈕所叫用的增益集函數。該物件可讓增益集識別所點選的按鈕，並向主機發出信號，指出其已完成處理。

例如，假設以下是在增益集資訊清單中定義的按鈕︰

```
<Control xsi:type="Button" id="eventTestButton">
  <Label resid="eventButtonLabel" />
  <Tooltip resid="eventButtonTooltip" />
  <Supertip>
    <Title resid="eventSuperTipTitle" />
    <Description resid="eventSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="blue-icon-16" />
    <bt:Image size="32" resid="blue-icon-32" />
    <bt:Image size="80" resid="blue-icon-80" />
  </Icon>
  <Action xsi:type="ExecuteFunction">
    <FunctionName>testEventObject</FunctionName>
  </Action>
</Control>
```

按鈕的 `id` 屬性已設定為 `eventTestButton`，並且將叫用在增益集中所定義 `testEventObject` 函數。該函數的外觀如下所示︰

```
function testEventObject(event) {
  // The event object implements the Event interface

  // This value will be "eventTestButton"
  var buttonId = event.source.id;

  // Signal to the host app that processing is complete.
  event.completed();
}
```

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](./tutorial-api-requirement-sets.md)| 1.3|
|[最低權限等級](../../docs/outlook/understanding-outlook-add-in-permissions.md)| 限制|
|適用的 Outlook 模式| 撰寫或讀取|

### <a name="members"></a>成員

####  <a name="source-:object"></a>source :Object

取得叫用方法的增益集命令按鈕識別碼。

`source` 屬性會傳回具有下列屬性的物件。

| 屬性 | 描述 |
| --- | --- |
| `id` | `id` 元素之 `Control` 屬性的值，其定義增益集資訊清單中的增益集命令按鈕。 |

當有多個按鈕叫用同一個函數，但是您必須根據點選的按鈕採取不同的動作時，可以使用這個值。

##### <a name="type:"></a>類型：

*   物件

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](./tutorial-api-requirement-sets.md)| 1.3|
|[最低權限等級](../../docs/outlook/understanding-outlook-add-in-permissions.md)| 限制|
|適用的 Outlook 模式| 撰寫或讀取|

##### <a name="example"></a>範例

```
// Function is used by two buttons:
// button1 and button2
function multiButton (event) {
  // Check which button was clicked
  var buttonId = event.source.id;

  if (buttonId === 'button1') {
    doButton1Action();
  else {
    doButton2Action();
  }

  event.completed();
}
```

### <a name="methods"></a>方法

####  <a name="completed()"></a>completed()

指出增益集已完成由增益集命令按鈕所觸發的處理。

這個方法必須在由增益集命令叫用的函數結束時呼叫，且增益集命令的定義為 `Action` 元素的 `xsi:type` 屬性已設定為 `ExecuteFunction`。呼叫這個方法會向主機用戶端發出訊號，指出函數已完成，而且它可以清除任何與叫用函數相關的狀態。例如，如果使用者在這個方法呼叫之前先關閉 Outlook，Outlook 就會發出警告，指出仍有執行中的函數。

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](./tutorial-api-requirement-sets.md)| 1.3|
|[最低權限等級](../../docs/outlook/understanding-outlook-add-in-permissions.md)| 限制|
|適用的 Outlook 模式| 撰寫或讀取|

##### <a name="example"></a>範例

```
function processItem (event) {
  // Do some processing

  event.completed();
}
```