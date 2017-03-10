# <a name="setting-object-javascript-api-for-excel"></a>Setting 物件 (適用於 Excel 的 JavaScript API)

Setting 表示設定的機碼值組會保存至文件。

## <a name="properties"></a>屬性

| 屬性	       | 類型	    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|索引鍵|string|傳回代表設定識別碼的索引鍵。唯讀。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|數值|物件|表示儲存此設定的值。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
無


## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|[delete()](#delete)|void|刪除設定。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法詳細資料


### <a name="delete"></a>delete()
刪除設定。

#### <a name="syntax"></a>語法
```js
settingObject.delete();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
void
