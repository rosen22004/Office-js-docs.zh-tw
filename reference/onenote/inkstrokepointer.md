# InkStrokePointer 物件 (適用於 OneNote 的 JavaScript API)

_適用於：OneNote Online_  


筆跡線條物件及其內容父項的弱式參考

## 屬性

| 屬性	     | 類型	   |說明|意見反應|
|:---------------|:--------|:----------|:-------|
|contentId|string|代表對應到此線條的頁面內容物件的識別碼|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkStrokePointer-contentId)|
|inkStrokeId|string|代表筆跡線條的識別碼|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkStrokePointer-inkStrokeId)|

_請參閱屬性存取[範例。](#範例)_

## 關聯性
無


## 方法

| 方法           | 傳回類型    |說明| 意見反應|
|:---------------|:--------|:----------|:-------|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|[執行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkStrokePointer-load)|

## 方法詳細資料


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
