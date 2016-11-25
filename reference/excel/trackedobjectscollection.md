# <a name="trackedobjectscollection-object-(javascript-api-for-office-2016)"></a>TrackedObjectsCollection 物件 (適用於 Office 2016 的 JavaScript API)

可讓增益集跨 sync() 批次管理 range 物件參考。一般而言，Excel.run() 可以自動跨批次維護參考，無需明確地追蹤它們。不過，如果增益集案例需要追蹤 range 物件並手動調整以反映基礎 Excel 範圍的目前狀態，便可使用這個集合來標記此類需要追蹤的物件。請注意，如果將 range 物件標記為追蹤，即使發生錯誤，在使用時也需要明確地將其移除以便從 Excel 中釋放記憶體。

## <a name="properties"></a>屬性
無

## <a name="relationships"></a>關聯性

無

## <a name="methods"></a>方法

TrackedObjectsCollection 物件有下列定義的方法。

| 方法     | 傳回類型    |描述|
|:-----------------|:--------|:----------|
|[add(rangeObject:Range)](#addrangeobject-range)| Null             |在範圍上建立新的參考。|
|[remove(rangeObject:Range)](#removerangeobject-range)| Null             |移除範圍上的參考。  |
|[removeAll()](#removeallrangeobject-range)| Null|移除裝置上由增益集建立的所有參考。|


## <a name="api-specification"></a>API 規格 

### <a name="add(rangeobject:-range)"></a>add(rangeObject: range)
加入 range 物件至 trackedObjectsCollection。如此會追蹤跨批次要求的所有基礎變更，任何後續更新也都會套用至 range 物件的目前狀態。 

#### <a name="syntax"></a>語法
```js
trackedObjectsCollection.add(rangeObject);
```

#### <a name="parameters"></a>參數

參數	       | 類型	   | 描述
--------------- | ------ | ------------
`rangeObject`  | [Range](range.md)| 要加入至 trackedObjectCollection 的 range 物件。

#### <a name="returns"></a>傳回
Null

#### <a name="examples"></a>範例

```js
var sheetName = "Sheet1";
var rangeAddress = "A1:B2";
var ctx = new Excel.RequestContext();
var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
ctx.trackedObjectsCollection.add(range);
ctx.load(range);

Excel.run(function (ctx) { 
    range.insert("Down");
    Console.log(range.address); // Address should be updated to A3:B4
    return ctx.sync(); 
});
```


### <a name="remove(rangeobject:-range)"></a>remove(rangeObject: range)

從集合中移除參考物件。這會釋放維護追蹤物件之狀態所需的記憶體和資源。請注意，如果將 range 物件標記為追蹤，即使發生錯誤，也需要明確地將其移除。

#### <a name="syntax"></a>語法
```js
trackedObjectsCollection.remove(rangeObject);
```

#### <a name="parameters"></a>參數

參數	       | 類型	   | 描述
--------------- | ------ | ------------
`rangeObject`  | [Range](range.md)| 要從 trackedObjectCollection 移除的 range 物件。

#### <a name="returns"></a>傳回
Null

#### <a name="examples"></a>範例


```js
var sheetName = "Sheet1";
var rangeAddress = "A1:B2";
var ctx = new Excel.RequestContext();
var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
ctx.trackedObjectsCollection.add(range);
ctx.load(range);

Excel.run(function (ctx) { 
    range.insert("Down");
    Console.log(range.address); // Address should be updated to A3:B4
    ctx.trackedObjectsCollection.remove(range); 
    return ctx.sync(); 
});
```

### <a name="removeall(rangeobject:-range)"></a>removeAll(rangeObject: range)

移除裝置上由增益集建立的所有參考。

#### <a name="syntax"></a>語法
```js
trackedObjectsCollection.removeAll();
```

#### <a name="parameters"></a>參數

無

#### <a name="returns"></a>傳回
Null

#### <a name="examples"></a>範例

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:B2";
    var ctx = new Excel.RequestContext();
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    ctx.trackedObjectsCollection.add(range);
    ctx.load(range);
    range.insert("Down");
    Console.log(range.address); // Address should be updated to A3:B4
    ctx.trackedObjectsCollection.removeAll(); 
    return ctx.sync(); 
});
```