# ChartAxes 物件 (適用於 Excel 的 JavaScript API)

_適用版本：Excel 2016、Excel Online、Office 2016_

代表圖表座標軸。

## 屬性

無

## 關聯性
| 關聯性 | 類型	|說明|
|:---------------|:--------|:----------|
|categoryAxis|[ChartAxis](chartaxis.md)|代表圖表中的類別軸。唯讀。|
|seriesAxis|[ChartAxis](chartaxis.md)|代表 3 維圖表的數列座標軸。唯讀。|
|valueAxis|[ChartAxis](chartaxis.md)|代表圖表中的數值軸。唯讀。|

## 方法

| 方法		   | 傳回類型	|說明|
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
| 參數	   | 類型	|說明|
|:---------------|:--------|:----------|
|param|object|選用。接受參數與關聯性名稱，做為分隔字串或陣列。或者提供 [loadOption](loadoption.md) 物件。|

#### 傳回
void

