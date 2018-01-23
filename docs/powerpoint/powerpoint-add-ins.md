# <a name="powerpoint-add-ins"></a>PowerPoint 外接程序

你可以使用 PowerPoint 外接程序为用户演示文稿构建跨平台（包括 Windows、iOS、Office Online 和 Mac）出色解决方案。可以创建以下两种类型的外接程序：

- 使用**内容外接程序**向演示文稿添加动态 HTML5 内容。有关示例，请参阅可用于将交互关系图从 LucidChart 插入面板的 [PowerPoint 的 LucidChart 关系图](https://store.office.com/en-us/app.aspx?assetid=WA104380117&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Productivity&homapppos=3&homchv=2&appredirect=false)外接程序。
- 使用**任务窗格外界程序**引入参考信息或通过服务将数据插入幻灯片。有关示例，请参阅可用于在演示文稿中添加专业照片的 [Shutterstock 图像](https://store.office.com/en-us/app.aspx?assetid=WA104380169&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Editor%2527s%2BPicks&homapppos=0&homchv=1&appredirect=false)外接程序。 

>
  **注意：**生成外接程序时，如果计划将外接程序[发布](../publish/publish.md)到 Office 应用商店，请务必遵循 [Office 应用商店验证策略](https://msdn.microsoft.com/en-us/library/jj220035.aspx)。例如，外接程序必须适用于支持你定义的方法的所有平台，才能通过验证（有关详细信息，请参阅[第 4.12 部分](https://msdn.microsoft.com/en-us/library/jj220035.aspx#Anchor_3)以及 [Office 外接程序主机和可用性](https://dev.office.com/add-in-availability)页）。

## <a name="powerpoint-add-in-scenarios"></a>PowerPoint 外接程序应用场景

本文中的代码示例展示了开发 PowerPoint 内容外接程序涉及的一些基本任务。 

这些示例依赖 `app.showNotification` 函数来显示信息，该函数包含在 Visual Studio Office 外接程序项目模板中。如果你没打算使用 Visual Studio 开发外接程序，则需要将 `showNotification` 函数替换为你自己的代码。其中一些示例还依赖在这些函数的作用域外声明的 `globals` 对象：`var globals = {activeViewHandler:0, firstSlideId:0};`

这些代码示例要求项目[引用 Office.js v1.1 库或更高版本](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md)。


## <a name="detect-the-presentations-active-view-and-handle-the-activeviewchanged-event"></a>检测演示文稿的活动视图并处理 ActiveViewChanged 事件

若要生成内容外接程序，则需要获取演示文稿的活动视图，并在 Office.Initialize 处理程序期间处理 ActiveViewChanged 事件。


- `getActiveFileView` 函数将调用 [Document.getActiveViewAsync](http://dev.office.com/reference/add-ins/shared/document.getactiveviewasync) 方法，以返回演示文稿的当前视图是“编辑”（你可在其中编辑幻灯片的任何视图，如**普通**或**大纲视图**）还是“阅读”（**幻灯片放映**或**阅读视图**）视图。


- `registerActiveViewChanged` 函数调用 [addHandlerAsync](http://dev.office.com/reference/add-ins/shared/document.addhandlerasync) 方法，注册 [Document.ActiveViewChanged](http://dev.office.com/reference/add-ins/shared/document.activeviewchanged) 事件的处理程序。 
> 注意：在 PowerPoint Online 中，[Document.ActiveViewChanged](http://dev.office.com/reference/add-ins/shared/document.activeviewchanged) 事件永远不会触发，因为幻灯片放映模式被视为新会话。在这种情况下，外接程序必须在加载时提取活动视图，如下所述。



```js

//general Office.initialize function. Fires on load of the add-in.
Office.initialize = function(){

    //Gets whether the current view is edit or read.
    var currentView = getActiveFileView();

    //register for the active view changed handler
    registerActiveViewChanged();

    //render the content based off of the currentView
    //....
}

function getActiveFileView()
{
    Office.context.document.getActiveViewAsync(function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification(asyncResult.value);
        }
    });

}


function registerActiveViewChanged() {
    Globals.activeViewHandler = function (args) {
        app.showNotification(JSON.stringify(args));
    }

    Office.context.document.addHandlerAsync(Office.EventType.ActiveViewChanged, Globals.activeViewHandler, 
        function (asyncResult) {
            if (asyncResult.status == "failed") {
           app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
            else {
            app.showNotification(asyncResult.status);
            }
        });
}
```
    

## <a name="navigate-to-a-particular-slide-in-the-presentation"></a>转到演示文稿中的特定幻灯片

`getSelectedRange` 函数将调用 [Document.getSelectedDataAsync](http://dev.office.com/reference/add-ins/shared/document.getselecteddataasync) 方法，以获取 `asyncResult.value` 返回的、包含名为“slides”的阵列的 JSON 对象，该阵列中包含所选幻灯片范围（或仅当前幻灯片）的 ID、标题和索引。它还会将所选范围内第一张幻灯片的 ID 保存到一个全局变量。


```js
function getSelectedRange() {
    // Get the id, title, and index of the current slide (or selected slides) and store the first slide id */
    Globals.firstSlideId = 0;

    Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            Globals.firstSlideId = asyncResult.value.slides[0].id;
            app.showNotification(JSON.stringify(asyncResult.value));
        }
    });
}
```

`goToFirstSlide` 函数将调用 [Document.goToByIdAsync](http://dev.office.com/reference/add-ins/shared/document.gotobyidasync) 方法，以转到上述 `getSelectedRange` 函数存储的第一张幻灯片的 ID。




```js
function goToFirstSlide() {
    Office.context.document.goToByIdAsync(Globals.firstSlideId, Office.GoToType.Slide, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification("Navigation successful");
        }
    });
}
```


## <a name="navigate-between-slides-in-the-presentation"></a>在演示文稿的幻灯片之间导航

`goToSlideByIndex` 函数调用 **Document.goToByIdAsync** 方法，转到演示文稿的下一张幻灯片。


```js
function goToSlideByIndex() {
    var goToFirst = Office.Index.First;
    var goToLast = Office.Index.Last;
    var goToPrevious = Office.Index.Previous;
    var goToNext = Office.Index.Next;

    Office.context.document.goToByIdAsync(goToNext, Office.GoToType.Index, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification("Navigation successful");
        }
    });
}
```

## <a name="get-the-url-of-the-presentation"></a>获取演示文稿的 URL

`getFileUrl` 函数调用 [Document.getFileProperties](http://dev.office.com/reference/add-ins/shared/document.getfilepropertiesasync) 方法，获取演示文稿文件的 URL。


```js
function getFileUrl() {
    //Get the URL of the current file.
    Office.context.document.getFilePropertiesAsync(function (asyncResult) {
        var fileUrl = asyncResult.value.url;
        if (fileUrl == "") {
            app.showNotification("The file hasn't been saved yet. Save the file and try again");
        }
        else {
            app.showNotification(fileUrl);
        }
    });
}
```



## <a name="additional-resources"></a>其他资源
- [PowerPoint 代码示例](https://dev.office.com/code-samples#?filters=powerpoint)

- [如何按文档保留内容和任务窗格外接程序的外接程序状态和设置](../develop/persisting-add-in-state-and-settings.md#how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins)

- [读取数据并将其写入文档或电子表格中的活动选择区](../develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)
    
- [从 PowerPoint 或 Word 相关外接程序中获取整个文档](../develop/get-the-whole-document-from-an-add-in-for-powerpoint-or-word.md)
    
- [在 PowerPoint 外接程序中使用文档主题](../powerpoint/use-document-themes-in-your-powerpoint-add-ins.md)
    
