# <a name="officeextension.error-object-(javascript-api-for-onenote)"></a>OfficeExtension.Error 物件 (適用於 OneNote 的 JavaScript API)

_適用於：OneNote Online_


表示當您使用 OneNote JavaScript API 時發生的錯誤。

## <a name="properties"></a>屬性
| 屬性	     | 類型	   |描述
|:---------------|:--------|:----------|
|code|string|取得指出錯誤類型的值。值可以是 "InvalidArgument"、"GeneralException"、"ItemNotFound" 或 "UnsupportedOperationForObjectType"。 |
|debugInfo|字串|取得指出當錯誤發生時，會發生什麼事的值。這個值只適用於在開發/偵錯期間。  |
|訊息 |string| 取得對應於錯誤程式碼之當地人們可以讀取的字串。|
|name |string| 取得永遠是 "OfficeExtension.Error" 的值。 |
|traceMessages |string[]| 取得對應於 context.trace(); 設定之檢測訊息的值陣列 |

## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述|
|:---------------|:--------|:----------|
|[toString()](#tostring)|string|以下列格式傳回錯誤碼和訊息值："{0}: {1}", code, message。|

## <a name="method-details"></a>方法詳細資料

### <a name="tostring()"></a>toString()
以下列格式傳回錯誤碼和訊息值："{0}: {1}", code, message。

#### <a name="syntax"></a>語法
```js
error.toString()
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>會傳回
string
