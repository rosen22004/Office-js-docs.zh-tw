# ChartTitleFormat 物件 (適用於 Excel 的 JavaScript API)

可供您存取圖表標題的 Office Art 格式設定。

## 屬性

無

## 關聯性
| 關聯性 | 類型	   |說明|
|:---------------|:--------|:----------|
|fill|[ChartFill](chartfill.md)|代表物件的填滿格式，其中包括背景格式資訊。唯讀。|
|font|[ChartFont](chartfont.md)|代表物件的字型屬性 (字型名稱、字型大小、色彩等)。唯讀。|

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
