# <a name="nameditemcollection-object-javascript-api-for-excel"></a>NamedItemCollection 物件 (適用於 Excel 的 JavaScript API)

屬於活頁簿或工作表一部份的所有 NamedItem 物件的集合，視到達方式而定。

## <a name="properties"></a>屬性

| 屬性	       | 類型	    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|項目|[NamedItem[]](nameditem.md)|NamedItem 物件的集合。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
無


## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|[add(name: string, reference:Range 或 string, comment: string)](#addname-string-reference-range-or-string-comment-string)|[NamedItem](nameditem.md)|新增名稱至指定範圍的集合。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[addFormulaLocal (name: string, formula: string, comment: string)](#addformulalocalname-string-formula-string-comment-string)|[NamedItem](nameditem.md)|使用使用者的公式地區設定，新增名稱至指定範圍的集合。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getCount()](#getcount)|Int|取得集合中的具名項目數目。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(name: string)](#getitemname-string)|[NamedItem](nameditem.md)|使用其名稱取得 nameditem 物件|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNullObject(name: string)](#getitemornullobjectname-string)|[NamedItem](nameditem.md)|使用其名稱取得 nameditem 物件。如果 nameditem 物件不存在，會傳回 null 物件。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法詳細資料


### <a name="addname-string-reference-range-or-string-comment-string"></a>add(name: string, reference:Range 或 string, comment: string)
新增名稱至指定範圍的集合。

#### <a name="syntax"></a>語法
```js
namedItemCollectionObject.add(name, reference, comment);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|:---|
|Name|string|具名項目的名稱。|
|參考資料|Range 或 string|名稱參照的公式或範圍。|
|註解|string|選用。與具名項目相關的註解|

#### <a name="returns"></a>傳回
[NamedItem](nameditem.md)

### <a name="addformulalocalname-string-formula-string-comment-string"></a>addFormulaLocal(name: string, formula: string, comment: string)
使用使用者的公式地區設定，新增名稱至指定範圍的集合。

#### <a name="syntax"></a>語法
```js
namedItemCollectionObject.addFormulaLocal(name, formula, comment);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|:---|
|Name|string|具名項目的「名稱」。|
|公式|string|位於名稱參照的使用者地區設定的公式。|
|註解|string|選用。與具名項目相關的註解|

#### <a name="returns"></a>傳回
[NamedItem](nameditem.md)

### <a name="getcount"></a>getCount()
取得集合中的具名項目數目。

#### <a name="syntax"></a>語法
```js
namedItemCollectionObject.getCount();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
Int

### <a name="getitemname-string"></a>getItem(name: string)
使用其名稱取得 nameditem 物件

#### <a name="syntax"></a>語法
```js
namedItemCollectionObject.getItem(name);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|:---|
|Name|string|nameditem 名稱。|

#### <a name="returns"></a>傳回
[NamedItem](nameditem.md)

#### <a name="examples"></a>範例

```js
Excel.run(function (ctx) { 
    var sheetName = 'Sheet1';
    var nameditem = ctx.workbook.names.getItem(sheetName);
    nameditem.load('type');
    return ctx.sync().then(function() {
            console.log(nameditem.type);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
### <a name="getitemornullobjectname-string"></a>getItemOrNullObject(name: string)
使用其名稱取得 nameditem 物件。如果 nameditem 物件不存在，會傳回 null 物件。

#### <a name="syntax"></a>語法
```js
namedItemCollectionObject.getItemOrNullObject(name);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|:---|
|Name|string|nameditem 名稱。|

#### <a name="returns"></a>傳回
[NamedItem](nameditem.md)
### <a name="property-access-examples"></a>屬性存取範例

```js
Excel.run(function (ctx) { 
    var nameditems = ctx.workbook.names;
    nameditems.load('items');
    return ctx.sync().then(function() {
        for (var i = 0; i < nameditems.items.length; i++)
        {
            console.log(nameditems.items[i].name);
            console.log(nameditems.items[i].index);
        }
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


