# <a name="getstarted-element"></a>GetStarted 元素

提供在 Word、Excel、PowerPoint 及 OneNote 主應用程式中安裝增益集時，顯示的圖說文字所使用的資訊。**GetStarted** 元素是 [DesktopFormFactor](./desktopformfactor.md) 的子元素。

## <a name="child-elements"></a>子元素

| 元素                       | 必要 | 描述                                        |
|:------------------------------|:--------:|:---------------------------------------------------|
| [Title](#title)               | 是      | 定義增益集公開功能的位置。     |
| [Description](#description)   | 是      | 檔案中包含 JavaScript 函式的 URL。|
| [LearnMoreUrl](#learnmoreurl) | 不可以       | 詳細說明增益集的頁面的 URL。   |


## <a name="title"></a>標題 
必要。用於圖說文字頂端的標題。**Resid** 屬性參考[資源](./resources.md)區段中 [ShortStrings](./resources.md#shortstrings) 元素中的有效識別碼。

## <a name="description"></a>描述
必要。圖說文字的描述/本文內容。**resid** 屬性參考[資源](./resources.md)區段中 [LongStrings](./resources.md#longstrings) 元素中的有效識別碼。

## <a name="learnmoreurl"></a>LearnMoreUrl
必要。使用者可以深入了解增益集的頁面的 URL。**resid** 屬性參考[資源](./resources.md)區段中 [Urls](./resources.md#urls) 元素中的有效識別碼。

> **附註：****LearnMoreUrl** 目前不會在 Word、Excel 或 PowerPoint 用戶端中轉譯。我們建議您針對所有用戶端新增這個 URL，以便在可供使用時轉譯 URL。 
