# <a name="breaktype-javascript-api-for-word"></a>BreakType (適用於 Word 的 JavaScript API)

指定中斷符號的形式。

_適用於：Word 2016、Word for iPad、Word for Mac、Word Online_

以下是 API 上支援的中斷類型。

| **值**         | **類型** | **描述**     |
|:-----------------|:--------|:----|
|行| | 分行符號。 |
|頁面| | 在插入點的分頁符號。|
|sectionNext| | 下一頁的分節符號。下一步的類型將過時。|
|sectionContinuous| | 沒有對應分頁符號的新節。|
|sectionEven| string | 開始於下一個偶數頁的下一節分節符號。如果分節符號位於偶數頁，Word 便會在下一個奇數頁保留空白。|
|sectionOdd| string | 開始於下一個奇數頁的下一節分節符號。如果分節符號位於奇數頁，Word 便會在下一個偶數頁保留空白。|

## <a name="support-details"></a>支援詳細資料
在執行階段檢查使用[需求集](../office-add-in-requirement-sets.md)以確認您的應用程式受到 Word 主應用程式版本的支援。如需有關 Office 主應用程式及伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。
