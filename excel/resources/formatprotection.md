# FormatProtection 物件 (適用於 Excel 的 JavaScript API)

_適用版本：Excel 2016、Excel Online、Excel for iOS、Office 2016_

代表 range 物件的格式保護。

## 屬性

| 屬性	  | 類型	| 說明
|:---------------|:--------|:----------||formulaHidden|bool|表示 Excel 是否在範圍的儲存格中隱藏公式。Null 值表示整個範圍沒有統一公式隱藏設定。||鎖定|bool|表示 Excel 是否鎖定物件中的儲存格。Null 值表示整個範圍沒有統一鎖定設定。|_請參閱屬性存取 [範例。](#property-access-examples)_

## 關聯性
無


## 方法

| 方法		  | 傳回類型	|描述||:---------------|:--------|:----------||[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|

## 方法詳細資料


### load(param: object)
以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。

#### 語法
```js
object.load(param);
```

#### 參數
| 參數	  | 類型	|描述||:---------------|:--------|:----------||參數|物件|選用。接受參數與關聯性名稱，做為分隔字串或陣列。或者提供 [loadOption](loadoption.md) 物件。|

#### 傳回
void

