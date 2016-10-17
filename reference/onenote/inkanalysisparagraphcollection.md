# <a name="inkanalysisparagraphcollection-object-(javascript-api-for-onenote)"></a>InkAnalysisParagraphCollection 物件 (適用於 OneNote 的 JavaScript API)

_適用於：OneNote Online_  


代表 InkAnalysisParagraph 物件的集合。

## <a name="properties"></a>屬性

| 屬性	     | 類型	   |描述|意見反應|
|:---------------|:--------|:----------|:-------|
|Count|int|傳回頁面中的 InkAnalysisParagraphs 數目。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisParagraphCollection-count)|
|items|[InkAnalysisParagraph[]](inkanalysisparagraph.md)|inkAnalysisParagraph 物件的集合。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisParagraphCollection-items)|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
無


## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述| 意見反應|
|:---------------|:--------|:----------|:-------|
|[getItem(index: number 或 string)](#getitemindex-number-or-string)|[InkAnalysisParagraph](inkanalysisparagraph.md)|藉由識別碼或藉由其集合中的索引，來取得 InkAnalysisParagraph 物件。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisParagraphCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[InkAnalysisParagraph](inkanalysisparagraph.md)|在集合中 InkAnalysisParagraph 的位置上取得它。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisParagraphCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisParagraphCollection-load)|

## <a name="method-details"></a>方法詳細資料


### <a name="getitem(index:-number-or-string)"></a>getItem(index: number or string)
藉由識別碼或藉由其集合中的索引，來取得 InkAnalysisParagraph 物件。唯讀。

#### <a name="syntax"></a>語法
```js
inkAnalysisParagraphCollectionObject.getItem(index);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|Index|number 或 string|InkAnalysisParagraph 物件的識別碼，或其在集合中 InkAnalysisParagraph 物件的索引位置。|

#### <a name="returns"></a>傳回
[InkAnalysisParagraph](inkanalysisparagraph.md)

### <a name="getitemat(index:-number)"></a>getItemAt(index: number)
在集合中 InkAnalysisParagraph 的位置上取得它。

#### <a name="syntax"></a>語法
```js
inkAnalysisParagraphCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|index|number|要擷取之物件的索引值。以 0 開始編製索引。|

#### <a name="returns"></a>傳回
[InkAnalysisParagraph](inkanalysisparagraph.md)

### <a name="load(param:-object)"></a>load(param: object)
以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。

#### <a name="syntax"></a>語法
```js
object.load(param);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|param|物件|選用。接受參數與關聯性名稱，做為分隔字串或陣列。或者提供 [loadOption](loadoption.md) 物件。|

#### <a name="returns"></a>傳回
void
