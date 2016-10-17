# <a name="inlinepicture-object-(javascript-api-for-word)"></a>InlinePicture 物件 (適用於 Word 的 JavaScript API)

代表文字間圖片。

_適用於：Word 2016、Word for iPad、Word for Mac、Word Online_

## <a name="properties"></a>屬性
| 屬性	     | 類型	   |描述
|:---------------|:--------|:----------|
|altTextDescription|string|取得或設定字串，代表與內嵌影像相關聯的替代文字|
|altTextTitle|string|取得或設定字串，包含內嵌影像的標題。|
|hyperlink|string|取得或設定與內嵌影像相關聯的超連結。|
|lockAspectRatio|bool|取得或設定值，指出調整大小時是否保留內嵌影像的原始比例。|

## <a name="relationships"></a>關聯性
| 關聯性 | 類型	   |描述|
|:---------------|:--------|:----------|
|height|**float**|取得或設定描述內嵌影像高度的數字。以點為單位。 |
|parentContentControl|[ContentControl](contentcontrol.md)|取得包含內嵌影像的內容控制項。如果沒有父代內容控制項，則傳回 null。唯讀。|
|paragraph|[paragraph](paragraph.md)|取得包含內嵌影像的段落。唯讀。
|width|**float**|取得或設定描述內嵌影像寬度的數字。以點為單位。|

## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述|
|:---------------|:--------|:----------|
|[delete()](#delete)|void|刪除文件中的圖片。|
|[getBase64ImageSrc()](#getbase64imagesrc)|object|取得物件，其值是內嵌影像的 base64 編碼字串表示法。|
|[insertBreak(breakType: BreakType, insertLocation: InsertLocation)](#insertbreakbreaktype-breaktype-insertlocation-insertlocation)|void|在指定的位置插入中斷符號。InsertLocation 值可以是 'Before' 或 'After'。|
|[insertContentControl()](#insertcontentcontrol)|[ContentControl](contentcontrol.md)|以 RTF 內容控制項圍繞文字間圖片。|
|[insertFileFromBase64(base64File: string, insertLocation:InsertLocation)](#insertfilefrombase64base64file-string-insertlocation-insertlocation)|[Range](range.md)|在內文的指定位置插入文件。InsertLocation 值可以是 'Before' 或 'After'。|
|[insertHtml(html: string, insertLocation:InsertLocation)](#inserthtmlhtml-string-insertlocation-insertlocation)|[Range](range.md)|在指定的位置插入 HTML。InsertLocation 值可以是 'Before' 或 'After'。|
|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation:InsertLocation)](#insertInlinePictureFromBase64base64EncodedImage-string-insertlocation-insertlocation)|[InlinePicture](inlinepicture.md)|在內文的指定位置插入圖片。InsertLocation 值可以是 'Replace'、'Before' 或 'After'。 |
|[insertOoxml(ooxml: string, insertLocation:InsertLocation)](#insertooxmlooxml-string-insertlocation-insertlocation)|[Range](range.md)|在指定的位置插入 OOXML。InsertLocation 值可以是 'Before' 或 'After'。|
|[insertParagraph(paragraphText: string, insertLocation:InsertLocation)](#insertparagraphparagraphtext-string-insertlocation-insertlocation)|[Paragraph](paragraph.md)|在指定的位置插入段落。InsertLocation 值可以是 'Before' 或 'After'。|
|[insertText(text: string, insertLocation:InsertLocation)](#inserttexttext-string-insertlocation-insertlocation)|[Range](range.md)|在內文的指定位置插入文字。InsertLocation 值可以是 'Before' 或 'After'。|
|[select(selectionMode: SelectionMode)](#selectselectionmode-selectionmode)|void|選取圖片並將 Word UI 導覽至該處。SelectionMode 值可以是 'Select'、'Start' 或 'End'。|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|

## <a name="method-details"></a>方法詳細資料

### <a name="delete()"></a>delete()
刪除文件中的圖片。

#### <a name="syntax"></a>語法
```js
inlinePictureObject.delete();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
void

### <a name="getbase64imagesrc()"></a>getBase64ImageSrc()
取得物件，其值是內嵌影像的 base64 編碼字串表示法。

#### <a name="syntax"></a>語法
```js
var base64 = inlinePictureObject.getBase64ImageSrc();
return context.sync().then(function () {    
    console.log("base64 string is " + base64.value);
});

```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
物件 



### <a name="insertbreak(breaktype:-breaktype,-insertlocation:-insertlocation)"></a>insertBreak(breakType: BreakType, insertLocation: InsertLocation)

#### <a name="syntax"></a>語法
```js
inlinePictureObject.insertBreak(breakType, insertLocation);
```
#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|breakType|BreakType|必要。要加入至內文的中斷類型。|
|insertLocation|InsertLocation|必要。此值可以是 'Before' 或 'After'。|

#### <a name="returns"></a>傳回
void

### <a name="insertcontentcontrol()"></a>insertContentControl()
以 RTF 內容控制項圍繞文字間圖片。

#### <a name="syntax"></a>語法
```js
inlinePictureObject.insertContentControl();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
[ContentControl](contentcontrol.md)

### <a name="insertfilefrombase64(base64file:-string,-insertlocation:-insertlocation)"></a>insertFileFromBase64(base64File: string, insertLocation:InsertLocation)
在內文的指定位置插入文件。InsertLocation 值可以是 'Before' 或 'After'。

#### <a name="syntax"></a>語法
```js
inlinePictureObject.insertFileFromBase64(base64File, insertLocation);
```
#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|base64File|string|必要。Docx 檔案的 base64 編碼內容。|
|insertLocation|InsertLocation|必要。此值可以是 'Before' 或 'After'。|

#### <a name="returns"></a>傳回
[Range](range.md)

### <a name="inserthtml(html:-string,-insertlocation:-insertlocation)"></a>insertHtml(html: string, insertLocation:InsertLocation)
在指定的位置插入 HTML。InsertLocation 值可以是 'Before' 或 'After'。

#### <a name="syntax"></a>語法
```js
inlinePictureObject.insertHtml(html, insertLocation);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|HTML|string|必要。要插入至文件的 HTML。|
|insertLocation|InsertLocation|必要。此值可以是 'Before' 或 'After'。|

#### <a name="returns"></a>傳回
[Range](range.md)


### <a name="insertinlinepicturefrombase64(base64encodedimage:-string,-insertlocation:-insertlocation)"></a>insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation:InsertLocation)
在內文的指定位置插入圖片。InsertLocation 值可以是 'Before' 或 'After'。

#### <a name="syntax"></a>語法
inlinePictureObject.insertInlinePictureFromBase64(image, insertLocation);

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|base64EncodedImage|string|必要。要插入至內文的 base64 編碼影像。|
|insertLocation|InsertLocation|必要。此值可以是 'Before' 或 'After'。|

#### <a name="returns"></a>傳回
[InlinePicture](inlinepicture.md)


### <a name="insertooxml(ooxml:-string,-insertlocation:-insertlocation)"></a>insertOoxml(ooxml: string, insertLocation:InsertLocation)
在指定的位置插入 OOXML。InsertLocation 值可以是 'Before' 或 'After'。

#### <a name="syntax"></a>語法
```js
inlinePictureObject.insertOoxml(ooxml, insertLocation);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|ooxml|string|必要。要插入的 OOXML。|
|insertLocation|InsertLocation|必要。此值可以是 'Before' 或 'After'。|

#### <a name="returns"></a>傳回
[Range](range.md)

### <a name="insertparagraph(paragraphtext:-string,-insertlocation:-insertlocation)"></a>insertParagraph(paragraphText: string, insertLocation:InsertLocation)
在指定的位置插入段落。InsertLocation 值可以是 'Before' 或 'After'。

#### <a name="syntax"></a>語法
```js
inlinePictureObject.insertParagraph(paragraphText, insertLocation);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|paragraphText|string|必要。要插入的段落文字。|
|insertLocation|InsertLocation|必要。此值可以是 'Before' 或 'After'。|

#### <a name="returns"></a>傳回
[Paragraph](paragraph.md)

### <a name="inserttext(text:-string,-insertlocation:-insertlocation)"></a>insertText(text: string, insertLocation:InsertLocation)
在內文的指定位置插入文字。InsertLocation 值可以是 'Before' 或 'After'。

#### <a name="syntax"></a>語法
```js
inlinePictureObject.insertText(text, insertLocation);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|文字|string|必要。要插入的文字。|
|insertLocation|InsertLocation|必要。此值可以是 'Before' 或 'After'。|

#### <a name="returns"></a>傳回
[Range](range.md)

### <a name="select(selectionmode:-selectionmode)"></a>select(selectionMode: SelectionMode)
選取圖片並將 Word UI 導覽至該處。SelectionMode 值可以是 'Select'、'Start' 或 'End'。

#### <a name="syntax"></a>語法
```js
inlinePictureObject.select(selectionMode);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|selectionMode|SelectionMode|選用。選取模式可以是 'Select'、'Start' 或 'End'。'Select' 為預設值。|

#### <a name="returns"></a>傳回
void

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

## <a name="support-details"></a>支援詳細資料
在執行階段檢查使用[需求集](../office-add-in-requirement-sets.md)以確認您的應用程式受到 Word 主應用程式版本的支援。如需有關 Office 主應用程式及伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。
