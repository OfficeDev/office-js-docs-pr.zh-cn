---
title: PowerPoint 加载项
description: ''
ms.date: 09/03/2019
ms.topic: overview
scenarios: getting-started
localization_priority: Priority
ms.openlocfilehash: 9a5b9746acab2916288a3c63d8ff14fa3b9f46f8
ms.sourcegitcommit: 78998a9f0ebb81c4dd2b77574148b16fe6725cfc
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/03/2019
ms.locfileid: "36715597"
---
# <a name="powerpoint-add-ins"></a>PowerPoint 加载项

使用 PowerPoint 加载项，可以跨平台（包括 Windows、iOS、Mac 和浏览器）生成极具吸引力的解决方案，从而有效展示用户的演示文稿。 可以创建以下两种类型的 PowerPoint 加载项：

- 使用**内容外接程序**向演示文稿添加动态 HTML5 内容。有关示例，请参阅可用于将交互关系图从 LucidChart 插入面板的 [PowerPoint 的 LucidChart 关系图](https://appsource.microsoft.com/product/office/WA104380117)外接程序。

- 使用**任务窗格加载项**引入参考信息或通过服务将数据插入演示文稿。 有关示例，请参阅可用于在演示文稿中添加专业照片的 [Pixton 漫画人物](https://appsource.microsoft.com/product/office/WA104380907)加载项。 

## <a name="powerpoint-add-in-scenarios"></a>PowerPoint 加载项方案

本文中的代码示例展示了开发 PowerPoint 加载项涉及的一些基本任务。 请注意以下几点：

- 这些示例使用 `app.showNotification` 函数来显示信息，该函数包含在 Visual Studio Office 加载项项目模板中。 如果你没打算使用 Visual Studio 开发加载项，则需要将 `showNotification` 函数替换为你自己的代码。 

- 其中一些示例还使用在这些函数的作用域外声明的 `Globals` 对象：`var Globals = {activeViewHandler:0, firstSlideId:0};`

- 若要使用这些示例，您的加载项项目必须[引用 Office.js v1.1 库或更高版本](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md)。

## <a name="detect-the-presentations-active-view-and-handle-the-activeviewchanged-event"></a>检测演示文稿的活动视图并处理 ActiveViewChanged 事件

若要生成内容外接程序，则需要获取演示文稿的活动视图，并在 `ActiveViewChanged` 处理程序期间处理 `Office.Initialize` 事件。

> [!NOTE]
> 在 PowerPoint 网页版中，[Document.ActiveViewChanged](/javascript/api/office/office.document) 事件永远不会触发，因为幻灯片放映模式被视为新会话。 在这种情况下，加载项必须在加载时提取活动视图，如下面的代码示例所述。

在以下代码示例中：

- `getActiveFileView` 函数将调用 [Document.getActiveViewAsync](/javascript/api/office/office.document#getactiveviewasync-options--callback-) 方法，以返回演示文稿的当前视图是“编辑”（你可在其中编辑幻灯片的任何视图，如**普通**或**大纲视图**）还是“阅读”（**幻灯片放映**或**阅读视图**）。

- `registerActiveViewChanged` 函数调用 [addHandlerAsync](/javascript/api/office/office.document#addhandlerasync-eventtype--handler--options--callback-) 方法，以注册 [Document.ActiveViewChanged](/javascript/api/office/office.document) 事件的处理程序。


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

在以下代码示例中，`getSelectedRange` 函数将调用 [Document.getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) 方法已获取 `asyncResult.value` 返回的 JSON 对象，其中包括一个名为 **slides** 的数组。 **slides** 数组包含所选范围内的幻灯片（或当前幻灯片，如果未选择多张幻灯片）的 ID、标题和索引。 此外，它会将所选范围内的第一张幻灯片的 ID 保存为全局变量。

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

在以下代码示例中，`goToFirstSlide` 函数将调用 [Document.goToByIdAsync](/javascript/api/office/office.document#gotobyidasync-id--gototype--options--callback-) 方法，以导航至由之前显示的 `getSelectedRange` 函数标识的第一张幻灯片。

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

在以下代码示例中，`goToSlideByIndex` 函数将调用 **Document.goToByIdAsync** 方法，以导航至演示文稿中的下一张幻灯片。

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

在以下代码实例中，`getFileUrl` 函数将调用 [Document.getFileProperties](/javascript/api/office/office.document#getfilepropertiesasync-options--callback-) 方法，以获取演示文稿文件的URL。

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

## <a name="create-a-presentation"></a>创建演示文稿

加载项可创建新的演示文稿，且与当前运行此加载项的 PowerPoint 实例分开。 PowerPoint 命名空间针对此目的提供了 `createPresentation` 方法。 调用此方法时，新的演示文稿将立即打开并在 PowerPoint 新实例中显示。 加载项保持打开状态，并随之前的演示文稿一起运行。

```js
PowerPoint.createPresentation();
```

此外，`createPresentation` 方法还可创建现有演示文稿的副本。 此方法接受 .pptx 文件的 base64 编码字符串表示形式作为可选参数。 若字符串参数为有效的 .pptx 文件，则生成的演示文稿是该文件的副本。 可以使用 [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) 类将文件转换为所需的 base64 编码字符串，如以下示例所示。

```js
var myFile = document.getElementById("file");
var reader = new FileReader();

reader.onload = function (event) {
    // strip off the metadata before the base64-encoded string
    var startIndex = reader.result.toString().indexOf("base64,");
    var copyBase64 = reader.result.toString().substr(startIndex + 7);

    PowerPoint.createPresentation(copyBase64);
};

// read in the file as a data URL so we can parse the base64-encoded string
reader.readAsDataURL(myFile.files[0]);
```

## <a name="see-also"></a>另请参阅

- [PowerPoint 加载项文档](index.md)
- [PowerPoint 代码示例](https://developer.microsoft.com/office/gallery/?filterBy=Samples,PowerPoint)
- [如何每文档保存内容和任务窗格加载项的加载项状态和设置](../develop/persisting-add-in-state-and-settings.md#how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins)
- [对文档或电子表格中的活动选择执行数据读取和写入操作](../develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)
- [通过 PowerPoint 或 Word 加载项获取整个文档](../powerpoint/get-the-whole-document-from-an-add-in-for-powerpoint.md)
- [在 PowerPoint 加载项中使用文档主题](use-document-themes-in-your-powerpoint-add-ins.md)
