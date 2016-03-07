# NamedItem 物件 (適用於 Excel 的 JavaScript API)

_適用版本：Excel 2016、Excel Online、Office 2016_

代表一個儲存格範圍或一個值的已定義名稱。名稱可以是原始命名物件 (如下列類型所示)、range 物件，以及範圍的參照。此物件可用來取得與名稱相關聯的 range 物件。

## 屬性

| 屬性	   | 類型	|說明
|:---------------|:--------|:----------|
|Name|string|物件的名稱。唯讀。|
|type|string|指出與名稱相關聯的參考類型。唯讀。可能的值為：String、Integer、Double、Boolean、Range。|
|value|object|代表名稱的定義參考公式，例如 =Sheet14!$B$2:$H$12, =4.75 等。唯讀。|
|visible|bool|指定物件是否可見。|

_請參閱屬性存取[範例。](#property-access-examples)_

## 關聯性
無


## 方法

| 方法		   | 傳回類型	|說明|
|:---------------|:--------|:----------|
|[getRange()](#getrange)|[Range](range.md)|傳回與名稱相關的 range 物件。如果具名項目的類型不是範圍，則擲回例外狀況。|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|

## 方法詳細資料

### getRange()
傳回與名稱相關的 range 物件。如果具名項目的類型不是範圍，則擲回例外狀況。

#### 語法
```js
namedItemObject.getRange();
```

#### 參數
無

#### 會傳回
[Range](range.md)

#### 範例

傳回與名稱相關的 range 物件。如果名稱不是 `Range` 類型則傳回 `null`。附註：此 API 目前僅支援 Workbook 範圍項目。

```js
Excel.run(function (ctx) { 
	var names = ctx.workbook.names;
	var range = names.getItem('MyRange').getRange();
	range.load('address');
	return ctx.sync().then(function() {
			console.log(range.address);
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
|param|object|選用。接受參數與關聯性名稱，做為分隔字串或陣列。或者提供 [loadOption](loadoption.md) 物件。|

#### 傳回
void
### 屬性存取範例

```js
Excel.run(function (ctx) { 
	var names = ctx.workbook.names;
	var namedItem = names.getItem('MyRange');
	namedItem.load('type');
	return ctx.sync().then(function() {
			console.log(namedItem.type);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

