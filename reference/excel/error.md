# OfficeExtension.Error 物件 (適用於 Excel 的 JavaScript API)

表示當您使用 Excel JavaScript API 時發生的錯誤。

_適用版本：Excel 2016、Excel Online、Excel for iOS、Office 2016_

## 屬性
| 屬性	     | 類型	   |說明
|:---------------|:--------|:----------|
|code|string|取得指出錯誤類型的值。值可以是 "AccessDenied"、"ActivityLimitReached"、"BadPassword"、"GeneralException"、"InsertDeleteConflict"、"InvalidArgument"、"InvalidBinding"、"InvalidOperation"、"InvalidReference"、"InvalidSelection"、"ItemAlreadyExists"、"ItemNotFound"、"NotImplemented" 或 "UnsupportedOperation"。 |
|debugInfo|string|取得指出當錯誤發生時，會發生什麼事的值。這個值只適用於在開發/偵錯期間。  |
|訊息 |string| 取得對應於錯誤程式碼之當地人們可以讀取的字串。|
|name |string| 取得永遠是 "OfficeExtension.Error" 的值。 |
|traceMessages |string[]| 取得對應於 context.trace(); 設定之檢測訊息的值陣列 |

## 方法

| 方法           | 傳回類型    |說明|
|:---------------|:--------|:----------|
|[toString()](#tostring)|string|以下列格式傳回錯誤碼和訊息值："{0}: {1}", code, message。|

## 方法詳細資料

### toString()
以下列格式傳回錯誤碼和訊息值："{0}: {1}", code, message。

#### 語法
```js
error.toString()
```

#### 參數
無

#### 會傳回
string
