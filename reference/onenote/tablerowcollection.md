# TableRowCollection 物件 (適用於 OneNote 的 JavaScript API)

_適用於：OneNote Online_  


含有 TableRow 物件的集合。

## 屬性

| 屬性	     | 類型	   |說明|意見反應|
|:---------------|:--------|:----------|:-------|
|Count|int|傳回此集合中的 TableCell 數目。 唯讀。|[執行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRowCollection-count)|
|項目|[TableRow[]](tablerow.md)|TableRow 物件的集合。 唯讀。|[執行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRowCollection-items)|

_請參閱屬性存取[範例。](#範例)_

## 關聯性
無


## 方法

| 方法           | 傳回類型    |說明| 意見反應|
|:---------------|:--------|:----------|:-------|
|[getItem(index: number 或 string)](#getitemindex-number-或-string)|[TableRow](tablerow.md)|藉由識別碼或藉由其集合中的索引，來取得 TableRow 物件。 唯讀。|[執行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRowCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[TableRow](tablerow.md)|按照 TableRow 在集合中的位置，取得 TableRow。|[執行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRowCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|[執行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRowCollection-load)|

## 方法詳細資料


### getItem(index: number 或 string)
藉由識別碼或藉由其集合中的索引，來取得 TableRow 物件。 唯讀。

#### 語法
```js
tableRowCollectionObject.getItem(index);
```

#### 參數
| 參數	    | 類型	   |說明|
|:---------------|:--------|:----------|
|Index|number 或 string|識別 TableRow 物件之索引位置的數值。|

#### 傳回
[TableRow](tablerow.md)

### getItemAt(index: number)
按照 TableRow 在集合中的位置，取得 TableRow。

#### 語法
```js
tableRowCollectionObject.getItemAt(index);
```

#### 參數
| 參數	    | 類型	   |說明|
|:---------------|:--------|:----------|
|index|number|要擷取之物件的索引值。以 0 開始編製索引。|

#### 傳回
[TableRow](tablerow.md)

### load(param: object)
以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。

#### 語法
```js
object.load(param);
```

#### 參數
| 參數	    | 類型	   |說明|
|:---------------|:--------|:----------|
|param|物件|選用。接受參數與關聯性名稱，做為分隔字串或陣列。或者提供 [loadOption](loadoption.md) 物件。|

#### 傳回
void
