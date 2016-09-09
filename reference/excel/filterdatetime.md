# FilterDatetime 物件 (適用於 Excel 的 JavaScript API)

_適用版本：Excel 2016、Excel Online、Excel for iOS、Office 2016_

表示如何篩選值時篩選日期。

## 屬性

| 屬性	     | 類型	   |描述
|:---------------|:--------|:----------|
|date|string|用來篩選資料的 ISO8601 格式的日期。|
|明確性|string|保留資料時應該使用多精確的日期。例如，如果日期是 2005-04-02 且明確性設定為「月」，篩選作業會保留日期在 2009 年 4 月份中的所有資料列。可能的值為：年、星期一、日、小時、分鐘、秒。|

_請參閱屬性存取[範例。](#範例)_

## 關聯性
無


## 方法

| 方法           | 傳回類型    |說明|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|

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
