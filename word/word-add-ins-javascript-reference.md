# Word 增益集 JavaScript 參考 

針對適用於 Word 的 JavaScript API，尋找 Word 增益集的 API 參考。

_適用版本：Word 2016、Word for iPad、Word for Mac_

## 在本章節中

以下是 Word JavaScript API 的主要物件。

* [Body](word-add-ins-javascript-reference/body.md)：代表文件或區段的內文。
* [ContentControl](word-add-ins-javascript-reference/contentcontrol.md)：內容的容器。這是指文件中具有界限且可能具有標籤的區域，這些區域會做為特定內容類型的容器。例如，內容控制項可能含有格式化文字的段落及其他內容控制項之類的內容。您可以透過文件、文件內文、段落、範圍或內容控制項的內容控制項集合，來存取內容控制項。
* [Document](word-add-ins-javascript-reference/document.md)：最上層物件。Document 物件包含一或多個 [section](word-add-ins-javascript-reference/section.md)，這是包含文件內容以及頁首/頁尾資訊的主體。
* [Font](word-add-ins-javascript-reference/font.md)：提供內文、內容控制項、段落或範圍的文字格式設定。
* [Image](word-add-ins-javascript-reference/inlinepicture.md)：代表錨定至一個段落的文字間圖片。
* [Paragraph](word-add-ins-javascript-reference/paragraph.md)：代表選取範圍、範圍或文件中的單一段落。您可以透過選取範圍、範圍或文件中的段落集合來存取段落。 
* [Range](word-add-ins-javascript-reference/range.md)：代表文件中的連續區域。當您取得選取範圍、將內容插入至內文、將內容插入至內容控制項、將內容插入至段落，或取得搜尋結果時，即可取得 Range 物件。無需變更選取範圍，即可定義與管理範圍。
* [Section](word-add-ins-javascript-reference/section.md)：定義不同的頁首和頁尾，以及文件的其他頁面版面配置設定。您可從 Document 物件存取 section。 
* [Selection](word-add-ins-javascript-reference/document.md#getselection)：Document 物件可讓您存取文件中的使用者選取範圍，如果未選取任何項目則可存取目前插入點。

## 歡迎您提供意見

我們很重視您的意見。 

* 查看文件，並在此存放庫中直接[送出問題](https://github.com/OfficeDev/office-js-docs/issues)，即可告知我們您找到的任何問題。
* 請告訴我們您的程式設計經驗、您希望未來版本提供哪些功能、程式碼範例等等。使用[這個網站](http://officespdev.uservoice.com/)可輸入您的建議和想法。

## 其他資源

* [Word 增益集](word-add-ins.md)
* [Word 增益集程式設計指南](word-add-ins-programming-guide.md)
* [Office 增益集](https://msdn.microsoft.com/en-us/library/office/jj220060.aspx)
* [Office 增益集入門](http://dev.office.com/getting-started/addins)
* &lt;a herf="https://github.com/OfficeDev?utf8=%E2%9C%93&amp;query=Word"&gt;GitHub 上的 Word 增益集&lt;/a&gt;
* [Word 的程式碼片段總管](http://officesnippetexplorer.azurewebsites.net/#/snippets/word)
