# RequestContext 物件 (適用於 Excel 的 JavaScript API)

_適用版本：Excel 2016、Excel Online、Office 2016_

RequestContext 物件可協助向 Excel 應用程式提出要求。由於 Office 增益集和 Excel 應用程式在兩個不同的處理程序中執行，因此需要使用要求內容以便從增益集存取 Excel 及相關的物件，例如工作表、表格等。 

## 屬性
無

## 方法

| 方法         | 傳回類型    |說明|
|:---------------|:--------|:----------|
|[load(object: object, option: object)](#loadobject-object-option-object)  |void     |以參數中指定的屬性和選項填滿 JavaScript 層中建立的 proxy 物件。|

## API 規格

### load(object: object, option: object)
以參數中指定的屬性和選項填滿 JavaScript 層中建立的 proxy 物件。

#### 語法
```js
requestContextObject.load(object, loadOption);
```

#### 參數
| 參數       | 類型    |描述|
|:----------------|:--------|:----------|
|物件|object|選用。指定要載入之物件的名稱。|
|option|[loadOption](loadoption.md)|選用。指定載入選項，例如 select、expand、skip 和 top。如需詳細資訊，請參閱 loadOption 物件。|

#### 傳回
void

##### 範例

下列範例會從一個範圍載入屬性值，並將它們複製到另一個範圍。

```js
Excel.run(function (ctx) { 
	var range = ctx.workbook.worksheets.getActiveWorksheet().getRange("A1:A2");
	ctx.load(range, "values");
	return ctx.sync().then(function() {
		var myvalues=range.values;
		ctx.workbook.worksheets.getActiveWorksheet().getRange("B1:B2").values = myvalues;
		console.log(range.values);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
})
```

