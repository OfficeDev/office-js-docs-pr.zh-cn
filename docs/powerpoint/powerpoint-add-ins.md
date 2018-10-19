---
title: PowerPoint 加载项
description: ''
ms.date: 10/16/2018
ms.openlocfilehash: 390497e74d4dc52b9d400f242850ab72bdb0eabc
ms.sourcegitcommit: a6d6348075c1abed76d2146ddfc099b0151fe403
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/19/2018
ms.locfileid: "25640076"
---
# <a name="powerpoint-add-ins"></a>PowerPoint 加载项

你可以使用 PowerPoint 加载项，为用户的演示文稿跨平台创建引人入胜的解决方案，包括 Windows、iOS、Office Online 和 Mac。 你可以创建两种类型的 PowerPoint 加载项：

- 使用**内容加载项**向演示文稿添加动态 HTML5 内容。有关示例，请参阅可用于将交互关系图从 LucidChart 插入面板的 [PowerPoint 的 LucidChart 关系图](https://store.office.com/app.aspx?assetid=WA104380117&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Productivity&homapppos=3&homchv=2&appredirect=false)加载项。

- 使用**任务窗格加载项**带来参考信息，或通过服务将数据插入演示文稿。 例如，请参阅 [Shutterstock 图像](https://store.office.com/app.aspx?assetid=WA104380169&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Editor%2527s%2BPicks&homapppos=0&homchv=1&appredirect=false)加载项，其中可以将专业照片添加到你的演示文稿。 

## <a name="powerpoint-add-in-scenarios"></a>PowerPoint 加载项方案

本文中的代码示例展示为 PowerPoint 开发加载项的一些基本任务。 请注意以下内容：

- 若要显示信息，这些示例使用 `app.showNotification` 函数，这包含在 Visual Studio Office 加载项项目模板中。 如果不使用 Visual Studio 来开发加载项，则需要使用自己的代码替换 `showNotification` 函数。 

- 其中的几个示例还使用超出这些函数范围声明的 `Globals` 对象：   `var Globals = {activeViewHandler:0, firstSlideId:0};`

- 若要使用这些示例，你的加载项项目必须[引用 Office.js v1.1 库或更高版本](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md)。

## <a name="detect-the-presentations-active-view-and-handle-the-activeviewchanged-event"></a>检测演示文稿的活动视图并处理 ActiveViewChanged 事件

若要生成内容加载项，则需要获取演示文稿的活动视图，并在 `Office.Initialize`  处理程序期间处理 `ActiveViewChanged`  事件。 

> [!NOTE]
> 在 PowerPoint Online 中，从不会触发 [Document.ActiveViewChanged](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js) 事件，因为幻灯片放映将被视为新会话。 在这种情况下，加载项必须在加载时获取活动视图，如以下代码示例所示。

在下面的代码示例中：

- `getActiveFileView`  函数将调用 [ Document.getActiveViewAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getactiveviewasync-options--callback-)  方法，以返回演示文稿的当前视图是“编辑”（你可在其中编辑幻灯片的任何视图，如**普通**或**大纲视图**）还是“阅读”（**幻灯片放映**或**阅读视图**）视图。

- `registerActiveViewChanged` 函数调用 [addHandlerAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#addhandlerasync-eventtype--handler--options--callback-) 方法，以注册 [Document.ActiveViewChanged](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js) 事件的事件处理程序。 


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

在下面的代码示例中，`getSelectedRange` 函数调用 [Document.getSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getselecteddataasync-coerciontype--options--callback-) 方法以获取 `asyncResult.value` 返回的 JSON 对象，其中包含名为**幻灯片**的数组。 该**幻灯片**数组包含所选幻灯片范围（或当前幻灯片，如果未选择多个幻灯片）的 id、标题和索引。 它还将保存所选范围的第一张幻灯片的 id 到全局变量。

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

在下面的代码示例中，`goToFirstSlide` 函数调用 [Document.goToByIdAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#gotobyidasync-id--gototype--options--callback-) 方法，以导航到之前所述 `getSelectedRange` 函数确定的第一张幻灯片。

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

在下面的代码示例中，`goToSlideByIndex` 函数调用 **Document.goToByIdAsync** 方法，以转到演示文稿的下一张幻灯片。

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

在下面的代码示例中， `getFileUrl` 函数调用 [Document.getFileProperties](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getfilepropertiesasync-options--callback-) 方法以获取演示文稿文件的 URL。

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



## <a name="see-also"></a>另请参阅
- [PowerPoint 代码示例](https://developer.microsoft.com/en-us/office/gallery/?filterBy=Samples,PowerPoint)
- [如何每文档保存内容和任务窗格加载项的加载项状态和设置](../develop/persisting-add-in-state-and-settings.md#how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins)
- [对文档或电子表格中的活动选择执行数据读取和写入操作](../develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)
- [通过 PowerPoint 或 Word 加载项获取整个文档](../powerpoint/get-the-whole-document-from-an-add-in-for-powerpoint.md)
- [在 PowerPoint 加载项中使用文档主题](use-document-themes-in-your-powerpoint-add-ins.md)
    
