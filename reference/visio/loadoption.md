# <a name="object-load-options-javascript-api-for-visio"></a>物件載入選項 (適用於 Visio 的 JavaScript API)

代表可以傳遞至 Load 方法的物件，以指定在執行 **sync()** 方法 (用以同步處理 Visio 物件和對應的 JavaScript Proxy 物件之間的狀態) 時要載入的屬性集和關聯。此物件需要 Select 和 Expand 參數等選項，以指定要載入至物件的屬性集，同時也允許在集合上分頁。

也可以提供包含要載入之屬性和關聯的字串，或提供包含要載入之屬性和關聯清單的陣列。請參閱下列的範例。

```js
object.load  ('<var1>,<relation1/var2>');

// Pass the parameter as an array.
object.load (["var1", "relation1/var2"]);
```

## <a name="properties"></a>屬性

| 屬性	 | 類型	  | 描述 |
|:---------|:------|:------------|
|select    |物件 |提供在 executeAsync 呼叫時要載入的參數/關聯性名稱的逗點分隔清單或陣列，例如 "property1, relationship1", [ "property1", "relationship1"]。選用。|
|expand    |物件 |提供在 executeAsync 呼叫時要載入的關聯性名稱的逗點分隔清單或陣列，例如 "relationship1, relationship2", [ "relationship1", "relationship2"]。選用。|
|top       |int    |指定結果中所包含查詢集合內的項目數。選用。|
|skip      |Int    |指定結果中要略過不予包含的集合項目數。如果指定 **top**，則結果的選取範圍會在略過指定的項目數後開始。選用。|

