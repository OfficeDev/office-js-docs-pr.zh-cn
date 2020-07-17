---
title: PowerPoint 加载项
description: 了解如何使用 PowerPoint 加载项跨平台（包括 Windows、iPad、Mac 和浏览器）生成极具吸引力的解决方案，从而有效展示演示文稿。
ms.date: 06/29/2020
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 314b441f3d4b6d2188ed630fe2b254aec42a86bc
ms.sourcegitcommit: 065bf4f8e0d26194cee9689f7126702b391340cc
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/01/2020
ms.locfileid: "45006449"
---
# <a name="powerpoint-add-ins"></a><span data-ttu-id="14f00-103">PowerPoint 加载项</span><span class="sxs-lookup"><span data-stu-id="14f00-103">PowerPoint add-ins</span></span>

<span data-ttu-id="14f00-104">使用 PowerPoint 加载项，可以跨平台（包括 Windows、iPad、Mac 和浏览器）生成极具吸引力的解决方案，从而有效展示用户的演示文稿。</span><span class="sxs-lookup"><span data-stu-id="14f00-104">You can use PowerPoint add-ins to build engaging solutions for your users' presentations across platforms including Windows, iPad, Mac, and in a browser.</span></span> <span data-ttu-id="14f00-105">可以创建以下两种类型的 PowerPoint 加载项：</span><span class="sxs-lookup"><span data-stu-id="14f00-105">You can create two types of PowerPoint add-ins:</span></span>

- <span data-ttu-id="14f00-p102">使用**内容外接程序**向演示文稿添加动态 HTML5 内容。有关示例，请参阅可用于将交互关系图从 LucidChart 插入面板的 [PowerPoint 的 LucidChart 关系图](https://appsource.microsoft.com/product/office/wa104380117)外接程序。</span><span class="sxs-lookup"><span data-stu-id="14f00-p102">Use **content add-ins** to add dynamic HTML5 content to your presentations. For example, see the [LucidChart Diagrams for PowerPoint](https://appsource.microsoft.com/product/office/wa104380117) add-in, which you can use to inject an interactive diagram from LucidChart into your deck.</span></span>

- <span data-ttu-id="14f00-108">使用**任务窗格加载项**引入参考信息或通过服务将数据插入演示文稿。</span><span class="sxs-lookup"><span data-stu-id="14f00-108">Use **task pane add-ins** to bring in reference information or insert data into the presentation via a service.</span></span> <span data-ttu-id="14f00-109">有关示例，请参阅可用于在演示文稿中添加专业照片的 [Pexels - 免费素材图片](https://appsource.microsoft.com/product/office/wa104379997)加载项。</span><span class="sxs-lookup"><span data-stu-id="14f00-109">For example, see the [Pexels - Free Stock Photos](https://appsource.microsoft.com/product/office/wa104379997) add-in, which you can use to add professional photos to your presentation.</span></span>

## <a name="powerpoint-add-in-scenarios"></a><span data-ttu-id="14f00-110">PowerPoint 加载项方案</span><span class="sxs-lookup"><span data-stu-id="14f00-110">PowerPoint add-in scenarios</span></span>

<span data-ttu-id="14f00-111">本文中的代码示例展示了开发 PowerPoint 加载项涉及的一些基本任务。</span><span class="sxs-lookup"><span data-stu-id="14f00-111">The code examples in this article demonstrate some basic tasks for developing add-ins for PowerPoint.</span></span> <span data-ttu-id="14f00-112">请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="14f00-112">Please note the following:</span></span>

- <span data-ttu-id="14f00-113">这些示例使用 `app.showNotification` 函数来显示信息，该函数包含在 Visual Studio Office 加载项项目模板中。</span><span class="sxs-lookup"><span data-stu-id="14f00-113">To display information, these examples use the `app.showNotification` function, which is included in the Visual Studio Office Add-ins project templates.</span></span> <span data-ttu-id="14f00-114">如果你没打算使用 Visual Studio 开发加载项，则需要将 `showNotification` 函数替换为你自己的代码。</span><span class="sxs-lookup"><span data-stu-id="14f00-114">If you aren't using Visual Studio to develop your add-in, you'll need replace the `showNotification` function with your own code.</span></span>

- <span data-ttu-id="14f00-115">其中一些示例还使用在这些函数的作用域外声明的 `Globals` 对象：`var Globals = {activeViewHandler:0, firstSlideId:0};`</span><span class="sxs-lookup"><span data-stu-id="14f00-115">Several of these examples also use a `Globals` object that is declared beyond the scope of these functions as:   `var Globals = {activeViewHandler:0, firstSlideId:0};`</span></span>

- <span data-ttu-id="14f00-116">若要使用这些示例，您的加载项项目必须[引用 Office.js v1.1 库或更高版本](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md)。</span><span class="sxs-lookup"><span data-stu-id="14f00-116">To use these examples, your add-in project must [reference Office.js v1.1 library or later](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).</span></span>

## <a name="detect-the-presentations-active-view-and-handle-the-activeviewchanged-event"></a><span data-ttu-id="14f00-117">检测演示文稿的活动视图并处理 ActiveViewChanged 事件</span><span class="sxs-lookup"><span data-stu-id="14f00-117">Detect the presentation's active view and handle the ActiveViewChanged event</span></span>

<span data-ttu-id="14f00-118">若要生成内容外接程序，则需要获取演示文稿的活动视图，并在 `ActiveViewChanged` 处理程序期间处理 `Office.Initialize` 事件。</span><span class="sxs-lookup"><span data-stu-id="14f00-118">If you are building a content add-in, you will need to get the presentation's active view and handle the `ActiveViewChanged` event, as part of your `Office.Initialize` handler.</span></span>

> [!NOTE]
> <span data-ttu-id="14f00-119">在 PowerPoint 网页版中，[Document.ActiveViewChanged](/javascript/api/office/office.document) 事件永远不会触发，因为幻灯片放映模式被视为新会话。</span><span class="sxs-lookup"><span data-stu-id="14f00-119">In PowerPoint on the web, the [Document.ActiveViewChanged](/javascript/api/office/office.document) event will never fire as Slide Show mode is treated as a new session.</span></span> <span data-ttu-id="14f00-120">在这种情况下，加载项必须在加载时提取活动视图，如下面的代码示例所述。</span><span class="sxs-lookup"><span data-stu-id="14f00-120">In this case, the add-in must fetch the active view on load, as shown in the following code sample.</span></span>

<span data-ttu-id="14f00-121">在以下代码示例中：</span><span class="sxs-lookup"><span data-stu-id="14f00-121">In the following code sample:</span></span>

- <span data-ttu-id="14f00-122">`getActiveFileView` 函数将调用 [Document.getActiveViewAsync](/javascript/api/office/office.document#getactiveviewasync-options--callback-) 方法，以返回演示文稿的当前视图是“编辑”（你可在其中编辑幻灯片的任何视图，如**普通**或**大纲视图**）还是“阅读”（**幻灯片放映**或**阅读视图**）。</span><span class="sxs-lookup"><span data-stu-id="14f00-122">The  `getActiveFileView` function calls the [Document.getActiveViewAsync](/javascript/api/office/office.document#getactiveviewasync-options--callback-) method to return whether the presentation's current view is "edit" (any of the views in which you can edit slides, such as **Normal** or **Outline View**) or "read" (**Slide Show** or **Reading View**).</span></span>

- <span data-ttu-id="14f00-123">`registerActiveViewChanged` 函数调用 [addHandlerAsync](/javascript/api/office/office.document#addhandlerasync-eventtype--handler--options--callback-) 方法，以注册 [Document.ActiveViewChanged](/javascript/api/office/office.document) 事件的处理程序。</span><span class="sxs-lookup"><span data-stu-id="14f00-123">The  `registerActiveViewChanged` function calls the [addHandlerAsync](/javascript/api/office/office.document#addhandlerasync-eventtype--handler--options--callback-) method to register a handler for the [Document.ActiveViewChanged](/javascript/api/office/office.document) event.</span></span>


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

## <a name="navigate-to-a-particular-slide-in-the-presentation"></a><span data-ttu-id="14f00-124">转到演示文稿中的特定幻灯片</span><span class="sxs-lookup"><span data-stu-id="14f00-124">Navigate to a particular slide in the presentation</span></span>

<span data-ttu-id="14f00-125">在以下代码示例中，`getSelectedRange` 函数将调用 [Document.getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) 方法以获取 `asyncResult.value` 返回的 JSON 对象，其中包括一个名为 `slides` 的数组。</span><span class="sxs-lookup"><span data-stu-id="14f00-125">In the following code sample, the `getSelectedRange` function calls the [Document.getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) method to get the JSON object returned by `asyncResult.value`, which contains an array named `slides`.</span></span> <span data-ttu-id="14f00-126">`slides` 数组包含所选范围内的幻灯片（或当前幻灯片，如果未选择多张幻灯片）的 ID、标题和索引。</span><span class="sxs-lookup"><span data-stu-id="14f00-126">The `slides` array contains the ids, titles, and indexes of selected range of slides (or of the current slide, if multiple slides are not selected).</span></span> <span data-ttu-id="14f00-127">此外，它会将所选范围内的第一张幻灯片的 ID 保存为全局变量。</span><span class="sxs-lookup"><span data-stu-id="14f00-127">It also saves the id of the first slide in the selected range to a global variable.</span></span>

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

<span data-ttu-id="14f00-128">在以下代码示例中，`goToFirstSlide` 函数将调用 [Document.goToByIdAsync](/javascript/api/office/office.document#gotobyidasync-id--gototype--options--callback-) 方法，以导航至由之前显示的 `getSelectedRange` 函数标识的第一张幻灯片。</span><span class="sxs-lookup"><span data-stu-id="14f00-128">In the following code sample, the `goToFirstSlide` function calls the [Document.goToByIdAsync](/javascript/api/office/office.document#gotobyidasync-id--gototype--options--callback-) method to navigate to the first slide that was identified by the `getSelectedRange` function shown previously.</span></span>

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

## <a name="navigate-between-slides-in-the-presentation"></a><span data-ttu-id="14f00-129">在演示文稿的幻灯片之间导航</span><span class="sxs-lookup"><span data-stu-id="14f00-129">Navigate between slides in the presentation</span></span>

<span data-ttu-id="14f00-130">在以下代码示例中，`goToSlideByIndex` 函数将调用 `Document.goToByIdAsync` 方法，以导航至演示文稿中的下一张幻灯片。</span><span class="sxs-lookup"><span data-stu-id="14f00-130">In the following code sample, the `goToSlideByIndex` function calls the `Document.goToByIdAsync` method to navigate to the next slide in the presentation.</span></span>

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

## <a name="get-the-url-of-the-presentation"></a><span data-ttu-id="14f00-131">获取演示文稿的 URL</span><span class="sxs-lookup"><span data-stu-id="14f00-131">Get the URL of the presentation</span></span>

<span data-ttu-id="14f00-132">在以下代码实例中，`getFileUrl` 函数将调用 [Document.getFileProperties](/javascript/api/office/office.document#getfilepropertiesasync-options--callback-) 方法，以获取演示文稿文件的URL。</span><span class="sxs-lookup"><span data-stu-id="14f00-132">In the following code sample, the  `getFileUrl` function calls the [Document.getFileProperties](/javascript/api/office/office.document#getfilepropertiesasync-options--callback-) method to get the URL of the presentation file.</span></span>

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

## <a name="create-a-presentation"></a><span data-ttu-id="14f00-133">创建演示文稿</span><span class="sxs-lookup"><span data-stu-id="14f00-133">Create a presentation</span></span>

<span data-ttu-id="14f00-134">加载项可创建新的演示文稿，且与当前运行此加载项的 PowerPoint 实例分开。</span><span class="sxs-lookup"><span data-stu-id="14f00-134">Your add-in can create a new presentation, separate from the PowerPoint instance in which the add-in is currently running.</span></span> <span data-ttu-id="14f00-135">PowerPoint 命名空间针对此目的提供了 `createPresentation` 方法。</span><span class="sxs-lookup"><span data-stu-id="14f00-135">The PowerPoint namespace has the `createPresentation` method for this purpose.</span></span> <span data-ttu-id="14f00-136">调用此方法时，新的演示文稿将立即打开并在 PowerPoint 新实例中显示。</span><span class="sxs-lookup"><span data-stu-id="14f00-136">When this method is called, the new presentation is immediately opened and displayed in a new instance of PowerPoint.</span></span> <span data-ttu-id="14f00-137">加载项保持打开状态，并随之前的演示文稿一起运行。</span><span class="sxs-lookup"><span data-stu-id="14f00-137">Your add-in remains open and running with the previous presentation.</span></span>

```js
PowerPoint.createPresentation();
```

<span data-ttu-id="14f00-138">此外，`createPresentation` 方法还可创建现有演示文稿的副本。</span><span class="sxs-lookup"><span data-stu-id="14f00-138">The `createPresentation` method can also create a copy of an existing presentation.</span></span> <span data-ttu-id="14f00-139">此方法接受 .pptx 文件的 base64 编码字符串表示形式作为可选参数。</span><span class="sxs-lookup"><span data-stu-id="14f00-139">The method accepts a base64-encoded string representation of an .pptx file as an optional parameter.</span></span> <span data-ttu-id="14f00-140">若字符串参数为有效的 .pptx 文件，则生成的演示文稿是该文件的副本。</span><span class="sxs-lookup"><span data-stu-id="14f00-140">The resulting presentation will be a copy of that file, assuming the string argument is a valid .pptx file.</span></span> <span data-ttu-id="14f00-141">可以使用 [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) 类将文件转换为所需的 base64 编码字符串，如以下示例所示。</span><span class="sxs-lookup"><span data-stu-id="14f00-141">The [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) class can be used to convert a file into the required base64-encoded string, as demonstrated in the following example.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="14f00-142">另请参阅</span><span class="sxs-lookup"><span data-stu-id="14f00-142">See also</span></span>

- [<span data-ttu-id="14f00-143">构建 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="14f00-143">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
- [<span data-ttu-id="14f00-144">PowerPoint 代码示例</span><span class="sxs-lookup"><span data-stu-id="14f00-144">PowerPoint Code Samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,PowerPoint)
- [<span data-ttu-id="14f00-145">如何每文档保存内容和任务窗格加载项的加载项状态和设置</span><span class="sxs-lookup"><span data-stu-id="14f00-145">How to save add-in state and settings per document for content and task pane add-ins</span></span>](../develop/persisting-add-in-state-and-settings.md#how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins)
- [<span data-ttu-id="14f00-146">对文档或电子表格中的活动选择执行数据读取和写入操作</span><span class="sxs-lookup"><span data-stu-id="14f00-146">Read and write data to the active selection in a document or spreadsheet</span></span>](../develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)
- [<span data-ttu-id="14f00-147">通过 PowerPoint 或 Word 加载项获取整个文档</span><span class="sxs-lookup"><span data-stu-id="14f00-147">Get the whole document from an add-in for PowerPoint or Word</span></span>](../powerpoint/get-the-whole-document-from-an-add-in-for-powerpoint.md)
- [<span data-ttu-id="14f00-148">在 PowerPoint 加载项中使用文档主题</span><span class="sxs-lookup"><span data-stu-id="14f00-148">Use document themes in your PowerPoint add-ins</span></span>](use-document-themes-in-your-powerpoint-add-ins.md)
