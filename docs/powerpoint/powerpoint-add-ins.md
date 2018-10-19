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
# <a name="powerpoint-add-ins"></a><span data-ttu-id="6cc6b-102">PowerPoint 加载项</span><span class="sxs-lookup"><span data-stu-id="6cc6b-102">PowerPoint add-ins</span></span>

<span data-ttu-id="6cc6b-103">你可以使用 PowerPoint 加载项，为用户的演示文稿跨平台创建引人入胜的解决方案，包括 Windows、iOS、Office Online 和 Mac。</span><span class="sxs-lookup"><span data-stu-id="6cc6b-103">You can use PowerPoint add-ins to build engaging solutions for your users' presentations across platforms including Windows, iOS, Office Online, and Mac. You can create one of two types of add-ins:</span></span> <span data-ttu-id="6cc6b-104">你可以创建两种类型的 PowerPoint 加载项：</span><span class="sxs-lookup"><span data-stu-id="6cc6b-104">You can create two types of PowerPoint add-ins:</span></span>

- <span data-ttu-id="6cc6b-p102">使用**内容加载项**向演示文稿添加动态 HTML5 内容。有关示例，请参阅可用于将交互关系图从 LucidChart 插入面板的 [PowerPoint 的 LucidChart 关系图](https://store.office.com/app.aspx?assetid=WA104380117&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Productivity&homapppos=3&homchv=2&appredirect=false)加载项。</span><span class="sxs-lookup"><span data-stu-id="6cc6b-p102">Use **content add-ins** to add dynamic HTML5 content to your presentations. For example, see the [LucidChart Diagrams for PowerPoint](https://store.office.com/app.aspx?assetid=WA104380117&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Productivity&homapppos=3&homchv=2&appredirect=false) add-in, which you can use to inject an interactive diagram from LucidChart into your deck.</span></span>

- <span data-ttu-id="6cc6b-107">使用**任务窗格加载项**带来参考信息，或通过服务将数据插入演示文稿。</span><span class="sxs-lookup"><span data-stu-id="6cc6b-107">Use **task pane add-ins** to bring in reference information or insert data into the slide via a service. For example, see the Shutterstock Images add-in, which you can use to add professional photos to your presentation.</span></span> <span data-ttu-id="6cc6b-108">例如，请参阅 [Shutterstock 图像](https://store.office.com/app.aspx?assetid=WA104380169&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Editor%2527s%2BPicks&homapppos=0&homchv=1&appredirect=false)加载项，其中可以将专业照片添加到你的演示文稿。</span><span class="sxs-lookup"><span data-stu-id="6cc6b-108">Use task pane add-ins to bring in reference information or insert data into the slide via a service. For example, see the [Shutterstock Images](https://store.office.com/app.aspx?assetid=WA104380169&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Editor%2527s%2BPicks&homapppos=0&homchv=1&appredirect=false) add-in, which you can use to add professional photos to your presentation.</span></span> 

## <a name="powerpoint-add-in-scenarios"></a><span data-ttu-id="6cc6b-109">PowerPoint 加载项方案</span><span class="sxs-lookup"><span data-stu-id="6cc6b-109">PowerPoint add-in scenarios</span></span>

<span data-ttu-id="6cc6b-110">本文中的代码示例展示为 PowerPoint 开发加载项的一些基本任务。</span><span class="sxs-lookup"><span data-stu-id="6cc6b-110">The code examples in the article show you some basic tasks for developing content add-ins for PowerPoint.</span></span> <span data-ttu-id="6cc6b-111">请注意以下内容：</span><span class="sxs-lookup"><span data-stu-id="6cc6b-111">Please note the following in 2nd_ProjServ_12 Beta 2:</span></span>

- <span data-ttu-id="6cc6b-112">若要显示信息，这些示例使用 `app.showNotification` 函数，这包含在 Visual Studio Office 加载项项目模板中。</span><span class="sxs-lookup"><span data-stu-id="6cc6b-112">To display information, these examples use the `app.showNotification` function, which is included in the Visual Studio Office Add-ins project templates.</span></span> <span data-ttu-id="6cc6b-113">如果不使用 Visual Studio 来开发加载项，则需要使用自己的代码替换 `showNotification` 函数。</span><span class="sxs-lookup"><span data-stu-id="6cc6b-113">If you aren't using Visual Studio to develop your add-in, you'll need replace the `showNotification` function with your own code.</span></span> 

- <span data-ttu-id="6cc6b-114">其中的几个示例还使用超出这些函数范围声明的 `Globals` 对象：   `var Globals = {activeViewHandler:0, firstSlideId:0};`</span><span class="sxs-lookup"><span data-stu-id="6cc6b-114">Several of these examples also use a `Globals` object that is declared beyond the scope of these functions as:   `var Globals = {activeViewHandler:0, firstSlideId:0};`</span></span>

- <span data-ttu-id="6cc6b-115">若要使用这些示例，你的加载项项目必须[引用 Office.js v1.1 库或更高版本](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md)。</span><span class="sxs-lookup"><span data-stu-id="6cc6b-115">To use these examples, your add-in project must [reference Office.js v1.1 library or later](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).</span></span>

## <a name="detect-the-presentations-active-view-and-handle-the-activeviewchanged-event"></a><span data-ttu-id="6cc6b-116">检测演示文稿的活动视图并处理 ActiveViewChanged 事件</span><span class="sxs-lookup"><span data-stu-id="6cc6b-116">Detect the presentation's active view and handle the ActiveViewChanged event</span></span>

<span data-ttu-id="6cc6b-117">若要生成内容加载项，则需要获取演示文稿的活动视图，并在 `Office.Initialize`  处理程序期间处理 `ActiveViewChanged`  事件。</span><span class="sxs-lookup"><span data-stu-id="6cc6b-117">If you are building a content add-in, you will need to get the presentation's active view and handle the ActiveViewChanged event, as part of your Office.Initialize handler.</span></span> 

> [!NOTE]
> <span data-ttu-id="6cc6b-118">在 PowerPoint Online 中，从不会触发 [Document.ActiveViewChanged](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js) 事件，因为幻灯片放映将被视为新会话。</span><span class="sxs-lookup"><span data-stu-id="6cc6b-118">In PowerPoint Online, the [Document.ActiveViewChanged](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js) event will never fire as Slide Show mode is treated as a new session. In this case, the add-in must fetch the active view on load, as noted below.</span></span> <span data-ttu-id="6cc6b-119">在这种情况下，加载项必须在加载时获取活动视图，如以下代码示例所示。</span><span class="sxs-lookup"><span data-stu-id="6cc6b-119">In this case, the add-in must fetch the active view on load, as shown in the following code sample.</span></span>

<span data-ttu-id="6cc6b-120">在下面的代码示例中：</span><span class="sxs-lookup"><span data-stu-id="6cc6b-120">In the following code sample:</span></span>

- <span data-ttu-id="6cc6b-121">`getActiveFileView`  函数将调用 [ Document.getActiveViewAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getactiveviewasync-options--callback-)  方法，以返回演示文稿的当前视图是“编辑”（你可在其中编辑幻灯片的任何视图，如**普通**或**大纲视图**）还是“阅读”（**幻灯片放映**或**阅读视图**）视图。</span><span class="sxs-lookup"><span data-stu-id="6cc6b-121">The getFileView function calls the Document.getActiveViewAsync method to return whether the presentation's current view is "edit" (any of the views in which you can edit slides, such as Normal or Outline View) or "read" (Slide Show or Reading View) view.</span></span>

- <span data-ttu-id="6cc6b-122">`registerActiveViewChanged` 函数调用 [addHandlerAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#addhandlerasync-eventtype--handler--options--callback-) 方法，以注册 [Document.ActiveViewChanged](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js) 事件的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="6cc6b-122">The  `registerActiveViewChanged` function calls the [addHandlerAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#addhandlerasync-eventtype--handler--options--callback-) method to register a handler for the [Document.ActiveViewChanged](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js) event.</span></span> 


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

## <a name="navigate-to-a-particular-slide-in-the-presentation"></a><span data-ttu-id="6cc6b-123">转到演示文稿中的特定幻灯片</span><span class="sxs-lookup"><span data-stu-id="6cc6b-123">Navigate to a particular slide in the presentation</span></span>

<span data-ttu-id="6cc6b-124">在下面的代码示例中，`getSelectedRange` 函数调用 [Document.getSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getselecteddataasync-coerciontype--options--callback-) 方法以获取 `asyncResult.value` 返回的 JSON 对象，其中包含名为**幻灯片**的数组。</span><span class="sxs-lookup"><span data-stu-id="6cc6b-124">In the following code sample, the `getSelectedRange` function calls the [Document.getSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getselecteddataasync-coerciontype--options--callback-) method to get the JSON object returned by `asyncResult.value`, which contains an array named **slides**.</span></span> <span data-ttu-id="6cc6b-125">该**幻灯片**数组包含所选幻灯片范围（或当前幻灯片，如果未选择多个幻灯片）的 id、标题和索引。</span><span class="sxs-lookup"><span data-stu-id="6cc6b-125">The **slides** array contains the ids, titles, and indexes of selected range of slides (or of the current slide, if multiple slides are not selected).</span></span> <span data-ttu-id="6cc6b-126">它还将保存所选范围的第一张幻灯片的 id 到全局变量。</span><span class="sxs-lookup"><span data-stu-id="6cc6b-126">It also saves the id of the first slide in the selected range to a global variable.</span></span>

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

<span data-ttu-id="6cc6b-127">在下面的代码示例中，`goToFirstSlide` 函数调用 [Document.goToByIdAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#gotobyidasync-id--gototype--options--callback-) 方法，以导航到之前所述 `getSelectedRange` 函数确定的第一张幻灯片。</span><span class="sxs-lookup"><span data-stu-id="6cc6b-127">In the following code sample, the `goToFirstSlide` function calls the [Document.goToByIdAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#gotobyidasync-id--gototype--options--callback-) method to navigate to the first slide that was identified by the `getSelectedRange` function shown previously.</span></span>

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

## <a name="navigate-between-slides-in-the-presentation"></a><span data-ttu-id="6cc6b-128">在演示文稿的幻灯片之间导航</span><span class="sxs-lookup"><span data-stu-id="6cc6b-128">Navigate between slides in the presentation</span></span>

<span data-ttu-id="6cc6b-129">在下面的代码示例中，`goToSlideByIndex` 函数调用 **Document.goToByIdAsync** 方法，以转到演示文稿的下一张幻灯片。</span><span class="sxs-lookup"><span data-stu-id="6cc6b-129">The  `goToSlideByIndex` function calls the **Document.goToByIdAsync** method to navigate to the next slide in the presentation.</span></span>

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

## <a name="get-the-url-of-the-presentation"></a><span data-ttu-id="6cc6b-130">获取演示文稿的 URL</span><span class="sxs-lookup"><span data-stu-id="6cc6b-130">Get the URL of the presentation</span></span>

<span data-ttu-id="6cc6b-131">在下面的代码示例中， `getFileUrl` 函数调用 [Document.getFileProperties](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getfilepropertiesasync-options--callback-) 方法以获取演示文稿文件的 URL。</span><span class="sxs-lookup"><span data-stu-id="6cc6b-131">The  `getFileUrl` function calls the [Document.getFileProperties](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getfilepropertiesasync-options--callback-) method to get the URL of the presentation file.</span></span>

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



## <a name="see-also"></a><span data-ttu-id="6cc6b-132">另请参阅</span><span class="sxs-lookup"><span data-stu-id="6cc6b-132">See also</span></span>
- [<span data-ttu-id="6cc6b-133">PowerPoint 代码示例</span><span class="sxs-lookup"><span data-stu-id="6cc6b-133">PowerPoint Code Samples</span></span>](https://developer.microsoft.com/en-us/office/gallery/?filterBy=Samples,PowerPoint)
- [<span data-ttu-id="6cc6b-134">如何每文档保存内容和任务窗格加载项的加载项状态和设置</span><span class="sxs-lookup"><span data-stu-id="6cc6b-134">How to save add-in state and settings per document for content and task pane add-ins</span></span>](../develop/persisting-add-in-state-and-settings.md#how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins)
- [<span data-ttu-id="6cc6b-135">对文档或电子表格中的活动选择执行数据读取和写入操作</span><span class="sxs-lookup"><span data-stu-id="6cc6b-135">Read and write data to the active selection in a document or spreadsheet</span></span>](../develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)
- [<span data-ttu-id="6cc6b-136">通过 PowerPoint 或 Word 加载项获取整个文档</span><span class="sxs-lookup"><span data-stu-id="6cc6b-136">Get the whole document from an add-in for PowerPoint or Word</span></span>](../powerpoint/get-the-whole-document-from-an-add-in-for-powerpoint.md)
- [<span data-ttu-id="6cc6b-137">在 PowerPoint 加载项中使用文档主题</span><span class="sxs-lookup"><span data-stu-id="6cc6b-137">Use document themes in your PowerPoint add-ins</span></span>](use-document-themes-in-your-powerpoint-add-ins.md)
    
