# TableCollection 物件 (適用於 Excel 的 JavaScript API)

_適用版本：Excel 2016、Excel Online、Office 2016_

代表屬於活頁簿一部份的所有表格集合。

## 屬性

| 屬性	   | 類型	|說明
|:---------------|:--------|:----------|
|Count|int|傳回工作表中的表格數目。唯讀。|
|items|[Table[]](table.md)|Table 物件的集合。唯讀。|

_請參閱屬性存取[範例。](#property-access-examples)_

## 關聯性
無


## 方法

| 方法		   | 傳回類型	|說明|
|:---------------|:--------|:----------|
|[add(address: string, hasHeaders: bool)](#addaddress-string-hasheaders-bool)|[Table](table.md)|建立新表格。範圍來源位址將決定表格加入至哪一個工作表。如果無法加入表格 (例如因為位址無效，或是此表格會與其他表格重疊)，則會擲回錯誤。|
|[getItem(key: number or string)](#getitemkey-number-or-string)|[Table](table.md)|依名稱或 ID 取得表格。|
|[getItemAt(index: number)](#getitematindex-number)|[Table](table.md)|根據表格在集合中的位置，取得表格。|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|

## 方法詳細資料

### add(address: string, hasHeaders: bool)
建立新表格。範圍來源位址將決定表格加入至哪一個工作表。如果無法加入表格 (例如因為位址無效，或是此表格會與其他表格重疊)，則會擲回錯誤。

#### 語法
```js
tableCollectionObject.add(address, hasHeaders);
```

#### 參數
| 參數	   | 類型	|說明|
|:---------------|:--------|:----------|
|地址|string|代表資料來源的 range 物件的位址或名稱。如果位址不含工作表名稱，則會使用目前作用中的工作表。|
|hasHeaders|bool|布林值，指出是要匯入的資料是否具有欄標籤。如果來源不含標頭 (亦即此屬性設為 false)，Excel 會自動產生標頭，並將資料向下移一列。|

#### 傳回
[Table](table.md)

#### 範例

```js
Excel.run(function (ctx) { 
	var table = ctx.workbook.tables.add('Sheet1!A1:E7', true);
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
### getItem(key: number or string)
依名稱或 ID 取得表格。

#### 語法
```js
tableCollectionObject.getItem(key);
```

#### 參數
| 參數	   | 類型	|說明|
|:---------------|:--------|:----------|
|key|number 或 string|要擷取之表格的名稱或 ID。|

#### 傳回
[Table](table.md)

#### 範例

```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var table = ctx.workbook.tables.getItem(tableName);
	return ctx.sync().then(function() {
			console.log(table.index);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


#### 範例

```js
Excel.run(function (ctx) { 
	var table = ctx.workbook.tables.getItemAt(0);
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

### getItemAt(index: number)
根據表格在集合中的位置，取得表格。

#### 語法
```js
tableCollectionObject.getItemAt(index);
```

#### 參數
| 參數	   | 類型	|說明|
|:---------------|:--------|:----------|
|index|number|要擷取之物件的索引值。以 0 開始編製索引。|

#### 傳回
[Table](table.md)

#### 範例

```js
Excel.run(function (ctx) { 
	var table = ctx.workbook.tables.getItemAt(0);
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
	var tables = ctx.workbook.tables;
	tables.load('items');
	return ctx.sync().then(function() {
		console.log("tables Count: " + tables.count);
		for (var i = 0; i < tables.items.length; i++)
		{
			console.log(tables.items[i].name);
		}
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

取得表格的數目。

```js
Excel.run(function (ctx) { 
	var tables = ctx.workbook.tables;
	tables.load('count');
	return ctx.sync().then(function() {
		console.log(tables.count);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```