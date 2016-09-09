# BreakType (適用於 Word 的 JavaScript API)

指定中斷符號的形式。

_適用版本：Word 2016、Word for iPad、Word for Mac_

以下是 API 上支援的中斷類型。

| **值**         | **類型** | **描述**     |
|:-----------------|:--------|:----|
|column| | 在插入點的分欄符號。 |
|line| | 分行符號。 |
|lineClearLeft| | 分行符號。 |
|lineClearRight| | 分行符號。 |
|next| | 位於下一頁的分節符號。 |
|page| | 在插入點的分頁符號。|
|sectionContinuous| | 沒有對應分頁符號的新節。|
|sectionEven| string | 開始於下一個偶數頁的下一節分節符號。如果分節符號位於偶數頁，Word 便會在下一個奇數頁保留空白。|
|sectionOdd| string | 開始於下一個奇數頁的下一節分節符號。如果分節符號位於奇數頁，Word 便會在下一個偶數頁保留空白。|
|textWrapping| string | 結束目前這一行，並強制文字從圖片、表格或其他項目下方繼續。文字會在下一個空白行 (不含靠右或靠左邊界對齊的表格) 中繼續。|

## 支援詳細資料
在執行階段檢查使用[需求集](../office-add-in-requirement-sets.md)以確認您的應用程式受到 Word 主應用程式版本的支援。如需有關 Office 主應用程式及伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。