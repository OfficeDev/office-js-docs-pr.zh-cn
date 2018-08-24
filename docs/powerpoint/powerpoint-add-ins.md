---
title: PowerPoint 加载项
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: e5c605410601d711e28ca04ff6e26387019cbb41
ms.sourcegitcommit: 4de2a1b62ccaa8e51982e95537fc9f52c0c5e687
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/10/2018
ms.locfileid: "22925316"
---
# <a name="powerpoint-add-ins"></a><span data-ttu-id="79ca5-102">PowerPoint 加载项</span><span class="sxs-lookup"><span data-stu-id="79ca5-102">PowerPoint add-ins</span></span>

<span data-ttu-id="79ca5-p101">你可以使用 PowerPoint 外接程序为用户演示文稿构建跨平台（包括 Windows、iOS、Office Online 和 Mac）出色解决方案。可以创建以下两种类型的外接程序：</span><span class="sxs-lookup"><span data-stu-id="79ca5-p101">You can use PowerPoint add-ins to build engaging solutions for your users' presentations across platforms including Windows, iOS, Office Online, and Mac. You can create one of two types of add-ins:</span></span>

- <span data-ttu-id="79ca5-p102">使用**内容外接程序**向演示文稿添加动态 HTML5 内容。有关示例，请参阅可用于将交互关系图从 LucidChart 插入面板的 [PowerPoint 的 LucidChart 关系图](https://store.office.com/app.aspx?assetid=WA104380117&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Productivity&homapppos=3&homchv=2&appredirect=false)外接程序。</span><span class="sxs-lookup"><span data-stu-id="79ca5-p102">Use **content add-ins** to add dynamic HTML5 content to your presentations. For example, see the [LucidChart Diagrams for PowerPoint](https://store.office.com/app.aspx?assetid=WA104380117&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Productivity&homapppos=3&homchv=2&appredirect=false) add-in, which you can use to inject an interactive diagram from LucidChart into your deck.</span></span>
- <span data-ttu-id="79ca5-p103">使用**任务窗格外界程序**引入参考信息或通过服务将数据插入幻灯片。有关示例，请参阅可用于在演示文稿中添加专业照片的 [Shutterstock 图像](https://store.office.com/app.aspx?assetid=WA104380169&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Editor%2527s%2BPicks&homapppos=0&homchv=1&appredirect=false)外接程序。</span><span class="sxs-lookup"><span data-stu-id="79ca5-p103">Use **task pane add-ins** to bring in reference information or insert data into the slide via a service. For example, see the [Shutterstock Images](https://store.office.com/app.aspx?assetid=WA104380169&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Editor%2527s%2BPicks&homapppos=0&homchv=1&appredirect=false) add-in, which you can use to add professional photos to your presentation.</span></span> 

## <a name="powerpoint-add-in-scenarios"></a><span data-ttu-id="79ca5-109">PowerPoint 加载项方案</span><span class="sxs-lookup"><span data-stu-id="79ca5-109">PowerPoint add-in scenarios</span></span>

<span data-ttu-id="79ca5-110">本文中的代码示例展示了开发 PowerPoint 内容外接程序涉及的一些基本任务。</span><span class="sxs-lookup"><span data-stu-id="79ca5-110">The code examples in the article show you some basic tasks for developing content add-ins for PowerPoint.</span></span> 

<span data-ttu-id="79ca5-p104">这些示例依赖 `app.showNotification` 函数来显示信息，该函数包含在 Visual Studio Office 外接程序项目模板中。如果你没打算使用 Visual Studio 开发外接程序，则需要将 `showNotification` 函数替换为你自己的代码。其中一些示例还依赖在这些函数的作用域外声明的 `globals` 对象：`var globals = {activeViewHandler:0, firstSlideId:0};`</span><span class="sxs-lookup"><span data-stu-id="79ca5-p104">To display information, these examples depend on the `app.showNotification` function, which is included in the Visual Studio Office Add-ins project templates. If you aren't using Visual Studio to develop your add-in, you'll need replace the `showNotification` function with your own code. Several of these examples also depend on this `globals` object that is declared outside of the scope of these functions: `var globals = {activeViewHandler:0, firstSlideId:0};`</span></span>

<span data-ttu-id="79ca5-114">这些代码示例要求项目[引用 Office.js v1.1 库或更高版本](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md)。</span><span class="sxs-lookup"><span data-stu-id="79ca5-114">These code examples require your project to [reference Office.js v1.1 library or later](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).</span></span>


## <a name="detect-the-presentations-active-view-and-handle-the-activeviewchanged-event"></a><span data-ttu-id="79ca5-115">检测演示文稿的活动视图并处理 ActiveViewChanged 事件</span><span class="sxs-lookup"><span data-stu-id="79ca5-115">Detect the presentation's active view and handle the ActiveViewChanged event</span></span>

<span data-ttu-id="79ca5-116">若要生成内容外接程序，则需要获取演示文稿的活动视图，并在 Office.Initialize 处理程序期间处理 ActiveViewChanged 事件。</span><span class="sxs-lookup"><span data-stu-id="79ca5-116">If you are building a content add-in, you will need to get the presentation's active view and handle the ActiveViewChanged event, as part of your Office.Initialize handler.</span></span>


- <span data-ttu-id="79ca5-117">`getActiveFileView` 函数将调用 [Document.getActiveViewAsync](https://dev.office.com/reference/add-ins/shared/document.getactiveviewasync) 方法，以返回演示文稿的当前视图是“编辑”（你可在其中编辑幻灯片的任何视图，如**普通**或**大纲视图**）还是“阅读”（**幻灯片放映**或**阅读视图**）视图。</span><span class="sxs-lookup"><span data-stu-id="79ca5-117">The  `getActiveFileView` function calls the [Document.getActiveViewAsync](https://dev.office.com/reference/add-ins/shared/document.getactiveviewasync) method to return whether the presentation's current view is "edit" (any of the views in which you can edit slides, such as **Normal** or **Outline View**) or "read" ( **Slide Show** or **Reading View**) view.</span></span>


- <span data-ttu-id="79ca5-118">`registerActiveViewChanged` 函数调用 [addHandlerAsync](https://dev.office.com/reference/add-ins/shared/document.addhandlerasync) 方法，以注册 [Document.ActiveViewChanged](https://dev.office.com/reference/add-ins/shared/document.activeviewchanged) 事件的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="79ca5-118">The  `registerActiveViewChanged` function calls the [addHandlerAsync](https://dev.office.com/reference/add-ins/shared/document.addhandlerasync) method to register a handler for the [Document.ActiveViewChanged](https://dev.office.com/reference/add-ins/shared/document.activeviewchanged) event.</span></span> 

> [!NOTE]
> <span data-ttu-id="79ca5-p105">在 PowerPoint Online 中，[Document.ActiveViewChanged](https://dev.office.com/reference/add-ins/shared/document.activeviewchanged) 事件永远不会触发，因为幻灯片放映模式被视为新会话。在这种情况下，加载项必须在加载时提取活动视图，如下所述。</span><span class="sxs-lookup"><span data-stu-id="79ca5-p105">In PowerPoint Online, the [Document.ActiveViewChanged](https://dev.office.com/reference/add-ins/shared/document.activeviewchanged) event will never fire as Slide Show mode is treated as a new session. In this case, the add-in must fetch the active view on load, as noted below.</span></span>

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
    

## <a name="navigate-to-a-particular-slide-in-the-presentation"></a><span data-ttu-id="79ca5-121">转到演示文稿中的特定幻灯片</span><span class="sxs-lookup"><span data-stu-id="79ca5-121">Navigate to a particular slide in the presentation</span></span>

<span data-ttu-id="79ca5-p106">`getSelectedRange` 函数将调用 [Document.getSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.getselecteddataasync) 方法，以获取 `asyncResult.value` 返回的、包含名为“slides”的阵列的 JSON 对象，该阵列中包含所选幻灯片范围（或仅当前幻灯片）的 ID、标题和索引。它还会将所选范围内第一张幻灯片的 ID 保存到一个全局变量。</span><span class="sxs-lookup"><span data-stu-id="79ca5-p106">The  `getSelectedRange` function calls the [Document.getSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.getselecteddataasync) method to get a JSON object returned by `asyncResult.value`, which contains an array named "slides" that contains the ids, titles, and indexes of selected range of slides (or just the current slide). It also saves the id of the first slide in the selected range to a global variable.</span></span>


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

<span data-ttu-id="79ca5-124">`goToFirstSlide` 函数调用 [Document.goToByIdAsync](https://dev.office.com/reference/add-ins/shared/document.gotobyidasync) 方法，以转到上述 `getSelectedRange` 函数存储的第一张幻灯片的 ID。</span><span class="sxs-lookup"><span data-stu-id="79ca5-124">The  `goToFirstSlide` function calls the [Document.goToByIdAsync](https://dev.office.com/reference/add-ins/shared/document.gotobyidasync) method to go to the id of the first slide stored by the `getSelectedRange` function above.</span></span>




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


## <a name="navigate-between-slides-in-the-presentation"></a><span data-ttu-id="79ca5-125">在演示文稿的幻灯片之间导航</span><span class="sxs-lookup"><span data-stu-id="79ca5-125">Navigate between slides in the presentation</span></span>

<span data-ttu-id="79ca5-126">`goToSlideByIndex` 函数调用 **Document.goToByIdAsync** 方法，转到演示文稿的下一张幻灯片。</span><span class="sxs-lookup"><span data-stu-id="79ca5-126">The  `goToSlideByIndex` function calls the **Document.goToByIdAsync** method to navigate to the next slide in the presentation.</span></span>


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

## <a name="get-the-url-of-the-presentation"></a><span data-ttu-id="79ca5-127">获取演示文稿的 URL</span><span class="sxs-lookup"><span data-stu-id="79ca5-127">Get the URL of the presentation</span></span>

<span data-ttu-id="79ca5-128">`getFileUrl` 函数调用 [Document.getFileProperties](https://dev.office.com/reference/add-ins/shared/document.getfilepropertiesasync) 方法，以获取演示文稿文件 URL。</span><span class="sxs-lookup"><span data-stu-id="79ca5-128">The  `getFileUrl` function calls the [Document.getFileProperties](https://dev.office.com/reference/add-ins/shared/document.getfilepropertiesasync) method to get the URL of the presentation file.</span></span>


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



## <a name="see-also"></a><span data-ttu-id="79ca5-129">另请参阅</span><span class="sxs-lookup"><span data-stu-id="79ca5-129">See also</span></span>
- [<span data-ttu-id="79ca5-130">PowerPoint 代码示例</span><span class="sxs-lookup"><span data-stu-id="79ca5-130">PowerPoint Code Samples</span></span>](https://dev.office.com/code-samples#?filters=powerpoint)
- [<span data-ttu-id="79ca5-131">如何每文档保存内容和任务窗格加载项的加载项状态和设置</span><span class="sxs-lookup"><span data-stu-id="79ca5-131">How to save add-in state and settings per document for content and task pane add-ins</span></span>](../develop/persisting-add-in-state-and-settings.md#how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins)
- [<span data-ttu-id="79ca5-132">对文档或电子表格中的活动选择执行数据读取和写入操作</span><span class="sxs-lookup"><span data-stu-id="79ca5-132">Read and write data to the active selection in a document or spreadsheet</span></span>](../develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)
- [<span data-ttu-id="79ca5-133">通过 PowerPoint 或 Word 加载项获取整个文档</span><span class="sxs-lookup"><span data-stu-id="79ca5-133">Get the whole document from an add-in for PowerPoint or Word</span></span>](../powerpoint/get-the-whole-document-from-an-add-in-for-powerpoint.md)
- [<span data-ttu-id="79ca5-134">在 PowerPoint 加载项中使用文档主题</span><span class="sxs-lookup"><span data-stu-id="79ca5-134">Use document themes in your PowerPoint add-ins</span></span>](use-document-themes-in-your-powerpoint-add-ins.md)
    
