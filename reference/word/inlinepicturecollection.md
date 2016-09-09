# InlinePictureCollection 物件 (適用於 Word 的 JavaScript API)

包含 [inlinePicture](inlinepicture.md) 物件的集合。

_適用版本：Word 2016、Word for iPad、Word for Mac_

## 屬性
| 屬性	     | 類型	   |說明
|:---------------|:--------|:----------|
|項目|[InlinePicture[]](inlinepicture.md)|InlinePicture 物件的集合。唯讀。|

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

## 支援詳細資料
在執行階段檢查使用[需求集](../office-add-in-requirement-sets.md)以確認您的應用程式受到 Word 主應用程式版本的支援。如需有關 Office 主應用程式及伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。
