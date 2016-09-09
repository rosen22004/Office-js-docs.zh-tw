
# ProjectViewTypes 列舉
指定 **[getSelectedViewAsync](../../reference/shared/projectdocument.getselectedviewasync.md)** 方法可以辨識的檢視類型。

|||
|:-----|:-----|
|**主機︰**|Project|
|**已新增於**|1.0|

```
ProjectViewTypes={
    Gantt           : 1, 
    NetworkDiagram  : 2, 
    TaskDiagram     : 3, 
    TaskForm        : 4, 
    TaskSheet       : 5, 
    ResourceForm    : 6, 
    ResourceSheet   : 7, 
    ResourceGraph   : 8, 
    TeamPlanner     : 9, 
    TaskDetails     : 10, 
    TaskNameForm    : 11, 
    ResourceNames   : 12, 
    Calendar        : 13, 
    TaskUsage       : 14, 
    ResourceUsage   : 15, 
    Timeline        : 16
}
```


## 成員


****


|**成員	**|**說明**|
|:-----|:-----|
|**Gantt**|甘特圖檢視。|
|**NetworkDiagram**|網狀圖檢視。|
|**TaskDiagram**|任務圖檢視。|
|**TaskForm**|任務表單檢視。|
|**TaskSheet**|任務工作表檢視。|
|**ResourceForm**|資源表單檢視。|
|**ResourceSheet**|資源工作表檢視。|
|**ResourceForm**|資源表單檢視。|
|**ResourceGraph**|資源圖表檢視。|
|**TeamPlanner**|團隊規劃檢視。|
|**TaskDetails**|任務詳細資訊檢視。|
|**TaskNameForm**|任務名稱表單檢視。|
|**ResourceNames**|資源名稱檢視。|
|**行事曆**|行事曆檢視。|
|**TaskUsage**|任務分派狀況檢視。|
|**ResourceUsage**|資源使用狀況檢視。|
|**時間表**|時間表檢視。|

## 備註

**[getSelectedViewAsync](../../reference/shared/projectdocument.getselectedviewasync.md)** 方法會回傳 **ProjectViewTypes** 常數值，與對應至使用中檢視的名稱。


## 支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此列舉。空白儲存格表示 Office 主應用程式不支援此列舉。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支援的主應用程式 (依平台排序)**


||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|
|:-----|:-----|:-----|
|**Project**|Y||

|||
|:-----|:-----|
|**增益集類型**|Task pane|
|**文件庫**|Office.js|
|**命名空間**|Office|

## 支援歷程記錄



****


|**版本**|**變更**|
|:-----|:-----|
|1.0|已導入|

## 請參閱



#### 其他資源


[getSelectedViewAsync 方法](../../reference/shared/projectdocument.getselectedviewasync.md)
