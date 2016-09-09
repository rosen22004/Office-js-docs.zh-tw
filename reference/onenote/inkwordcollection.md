# InkWordCollection 物件 (適用於 OneNote 的 JavaScript API)

_適用於：OneNote Online_  


代表 InkWord 物件的集合。

## 屬性

| 屬性	     | 類型	   |說明|意見反應|
|:---------------|:--------|:----------|:-------|
|Count|int|傳回頁面中的 InkWords 數目。 唯讀。|[執行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkWordCollection-count)|
|items|[InkWord[]](inkword.md)|InkWord 物件的集合。 唯讀。|[執行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkWordCollection-items)|

_請參閱屬性存取[範例。](#範例)_

## 關聯性
無


## 方法

| 方法           | 傳回類型    |說明| 意見反應|
|:---------------|:--------|:----------|:-------|
|[getItem(index: number or string)](#getitemindex-number-or-string)|[InkWord](inkword.md)|藉由識別碼或藉由其集合中的索引，來取得 InkWord 物件。 唯讀。|[執行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkWordCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[InkWord](inkword.md)|在集合中 InkWord 的位置上取得它。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkWordCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|[執行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkWordCollection-load)|

## 方法詳細資料


### getItem(index: number or string)
藉由識別碼或藉由其集合中的索引，來取得 InkWord 物件。 唯讀。

#### 語法
```js
inkWordCollectionObject.getItem(index);
```

#### 參數
| 參數	    | 類型	   |說明|
|:---------------|:--------|:----------|
|Index|number 或 string|InkWord 物件的識別碼，或 InkWord 物件在集合中的索引位置。|

#### 傳回
[InkWord](inkword.md)

### getItemAt(index: number)
在集合中 InkWord 的位置上取得它。

#### 語法
```js
inkWordCollectionObject.getItemAt(index);
```

#### 參數
| 參數	    | 類型	   |說明|
|:---------------|:--------|:----------|
|index|number|要擷取之物件的索引值。以 0 開始編製索引。|

#### 傳回
[InkWord](inkword.md)

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
