# <a name="uiclosecontainer-method"></a>UI.closeContainer 方法

此方法會將執行 JavaScript 的 UI 容器關閉。此方法的行為是由下表指定。

| 呼叫來源為以下項目時： | 行為 |
|:-----------------|:---------|
| 無 UI 命令按鈕 | 沒有效果。[displayDialogAsync](officeui.displaydialogasync.md) 所開啟的任何對話視窗將會維持開啟狀態。 |
| 工作窗格 | 該工作窗格將會關閉。`displayDialogAsync` 所開啟的任何對話也會關閉。如果該工作窗格支援釘選，而且使用者將其釘選，就會取消其釘選。 |
| 模組副檔名 | 沒有效果。 |

## <a name="syntax"></a>語法

```js
Office.context.ui.closeContainer();
```

## <a name="returns"></a>傳回
void