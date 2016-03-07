# Binding 物件 (適用於 Excel 的 JavaScript API)

_適用版本：Excel 2016、Excel Online、Office 2016_

代表活頁簿中定義的 Office.js 繫結。

## 屬性

| 屬性	   | 類型	|說明
|:---------------|:--------|:----------|
|id|string|代表繫結識別碼。唯讀。|
|type|string|傳回繫結的類型。唯讀。可能的值為：Range、Table、Text。|

_請參閱屬性存取[範例。](#property-access-examples)_

## 關聯性
無


## 方法

| 方法		   | 傳回類型	|說明|
|:---------------|:--------|:----------|
|[getRange()](#getrange)|[Range](range.md)|傳回繫結所代表的範圍。如果繫結不是正確的類型，則會擲回錯誤。|
|[getTable()](#gettable)|[Table](table.md)|傳回繫結所代表的表格。如果繫結不是正確的類型，則會擲回錯誤。|
|[getText()](#gettext)|string|傳回繫結所代表的文字。如果繫結不是正確的類型，則會擲回錯誤。|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|

## 方法詳細資料

### getRange()
傳回繫結所代表的範圍。如果繫結不是正確的類型，則會擲回錯誤。

#### 語法
```js
bindingObject.getRange();
```

#### 參數
無

#### 會傳回
[Range](range.md)

#### 範例
下面的範例會使用 binding 物件，以取得相關的範圍。

```js
Excel.run(function (ctx) { 
	var binding = ctx.workbook.bindings.getItemAt(0);
	var range = binding.getRange();
	range.load('cellCount');
	return ctx.sync().then(function() {
		console.log(range.cellCount);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getTable()
傳回繫結所代表的表格。如果繫結不是正確的類型，則會擲回錯誤。

#### 語法
```js
bindingObject.getTable();
```

#### 參數
無

#### 會傳回
[Table](table.md)

#### 範例
```js
Excel.run(function (ctx) { 
	var binding = ctx.workbook.bindings.getItemAt(0);
	var table = binding.getTable();
	table.load('name');
	return ctx.sync().then(function() {
			console.log(table.name);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getText()
傳回繫結所代表的文字。如果繫結不是正確的類型，則會擲回錯誤。

#### 語法
```js
bindingObject.getText();
```

#### 參數
無

#### 會傳回
字串

#### 範例

```js
Excel.run(function (ctx) { 
	var binding = ctx.workbook.bindings.getItemAt(0);
	var text = binding.getText();
	ctx.load('text');
	return ctx.sync().then(function() {
		console.log(text);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### load(param: object)
以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。

#### 語法
```js
object.load(param);
```

#### 參數
| 參數	   | 類型	|說明|
|:---------------|:--------|:----------|
|param|object|選用。接受參數與關係名稱，做為分隔字串或陣列。或者接受 [loadOption](loadoption.md) 物件。|

#### 傳回
void
### 屬性存取範例

```js
Excel.run(function (ctx) { 
	var binding = ctx.workbook.bindings.getItemAt(0);
	binding.load('type');
	return ctx.sync().then(function() {
		console.log(binding.type);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

