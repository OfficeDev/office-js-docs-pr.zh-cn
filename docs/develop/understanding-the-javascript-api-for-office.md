---
title: 了解适用于 Office 的 JavaScript API
description: ''
ms.date: 01/17/2019
localization_priority: Priority
ms.openlocfilehash: e685985783b08b51725165b03863ff3b0fffeeaf
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/23/2019
ms.locfileid: "29388820"
---
# <a name="understanding-the-javascript-api-for-office"></a><span data-ttu-id="3eb4c-102">了解适用于 Office 的 JavaScript API</span><span class="sxs-lookup"><span data-stu-id="3eb4c-102">Understanding the JavaScript API for Office</span></span>

<span data-ttu-id="3eb4c-p101">本文提供了有关适用于 Office 的 JavaScript API 的信息以及使用方法。有关参考信息，请参阅 [适用于 Office 的 JavaScript API](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)。有关将 Visual Studio 项目文件更新到适用于 Office 的 JavaScript API 的最新当前版本的信息，请参阅 [更新适用于 Office 的 JavaScript API 版本和清单架构文件](update-your-javascript-api-for-office-and-manifest-schema-version.md)。</span><span class="sxs-lookup"><span data-stu-id="3eb4c-p101">This article provides information about the JavaScript API for Office and how to use it. For reference information, see [JavaScript API for Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office). For information about updating Visual Studio project files to the most current version of the JavaScript API for Office, see [Update the version of your JavaScript API for Office and manifest schema files](update-your-javascript-api-for-office-and-manifest-schema-version.md).</span></span>

> [!NOTE]
> <span data-ttu-id="3eb4c-p102">如果计划将加载项[发布](../publish/publish.md)到 AppSource 并适用于 Office 体验，请务必遵循 [AppSource 验证策略](https://docs.microsoft.com/office/dev/store/validation-policies)。例如，加载项必须适用于支持已定义方法的所有平台，才能通过验证（有关详细信息，请参阅[第 4.12 部分](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably)以及 [Office 加载项主机和可用性](../overview/office-add-in-availability.md)页面）。</span><span class="sxs-lookup"><span data-stu-id="3eb4c-p102">If you plan to [publish](../publish/publish.md) your add-in to AppSource and make it available within the Office experience, make sure that you conform to the [AppSource validation policies](https://docs.microsoft.com/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)).</span></span> 

## <a name="referencing-the-javascript-api-for-office-library-in-your-add-in"></a><span data-ttu-id="3eb4c-108">在加载项中引用适用于 Office 的 JavaScript API 库</span><span class="sxs-lookup"><span data-stu-id="3eb4c-108">Referencing the JavaScript API for Office library in your add-in</span></span>

<span data-ttu-id="3eb4c-p103">[适用于 Office 的 JavaScript](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office) 库包含 Office.js 文件和关联的特定于主机应用程序的 .js 文件，例如 Excel-15.js 和 Outlook-15.js。引用该 API 最简单的方法是通过添加以下 `<script>` 到你的页面的 `<head>` 标记来使用我们的 CDN：</span><span class="sxs-lookup"><span data-stu-id="3eb4c-p103">The [JavaScript API for Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office) library consists of the Office.js file and associated host application-specific .js files, such as Excel-15.js and Outlook-15.js. The simplest method of referencing the API is using our CDN by adding the following `<script>` to your page's `<head>` tag:</span></span>  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

<span data-ttu-id="3eb4c-111">这将在加载项首次加载时下载并缓存适用于 Office 的 JavaScript API 文件，以确保对特定版本使用 Office.js 及其关联文件的最新实现。</span><span class="sxs-lookup"><span data-stu-id="3eb4c-111">This will download and cache the JavaScript API for Office files the first time your add-in loads to make sure that it is using the most up-to-date implementation of Office.js and its associated files for the specified version.</span></span>

<span data-ttu-id="3eb4c-112">有关 Office.js CDN 的详细信息，包括如何处理版本控制和向后兼容性，请参阅[从适用于 Office 的 JavaScript API 的内容传送网络 (CDN) 引用适用于 Office 的 JavaScript API 库](referencing-the-javascript-api-for-office-library-from-its-cdn.md)。</span><span class="sxs-lookup"><span data-stu-id="3eb4c-112">For more details around the Office.js CDN, including how versioning and backward compatibility is handled, see [Referencing the JavaScript API for Office library from its content delivery network (CDN)](referencing-the-javascript-api-for-office-library-from-its-cdn.md).</span></span>

## <a name="initializing-your-add-in"></a><span data-ttu-id="3eb4c-113">初始化加载项</span><span class="sxs-lookup"><span data-stu-id="3eb4c-113">Initializing your add-in</span></span>

<span data-ttu-id="3eb4c-114">**适用于：** 所有加载项类型</span><span class="sxs-lookup"><span data-stu-id="3eb4c-114">**Applies to:** All add-in types</span></span>

<span data-ttu-id="3eb4c-115">Office 加载项通常使用启动逻辑执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="3eb4c-115">Office Add-ins often have start-up logic to do things such as:</span></span>

- <span data-ttu-id="3eb4c-116">检查用户的 Office 版本是否支持代码调用的所有 Office API。</span><span class="sxs-lookup"><span data-stu-id="3eb4c-116">Check that the user's version of Office will support all the Office APIs that your code calls.</span></span>

- <span data-ttu-id="3eb4c-117">确保存在某些项目，例如具有特定名称的工作表。</span><span class="sxs-lookup"><span data-stu-id="3eb4c-117">Ensure the existence of certain artifacts, such as worksheet with a specific name.</span></span>

- <span data-ttu-id="3eb4c-118">提示用户选择 Excel 中的某些单元格，然后插入使用这些所选值进行初始化的图表。</span><span class="sxs-lookup"><span data-stu-id="3eb4c-118">Prompting the user to select some cells in Excel, and then inserting a chart initialized with those selected values.</span></span>

- <span data-ttu-id="3eb4c-119">建立绑定。</span><span class="sxs-lookup"><span data-stu-id="3eb4c-119">Establish bindings.</span></span>

- <span data-ttu-id="3eb4c-120">使用 Office 对话框 API 提示用户使用默认的加载项设置值。</span><span class="sxs-lookup"><span data-stu-id="3eb4c-120">Use the Office dialog API to prompt the user for default add-in settings values.</span></span>

<span data-ttu-id="3eb4c-121">但在加载库前，启动代码不得调用任何 Office.js API。</span><span class="sxs-lookup"><span data-stu-id="3eb4c-121">But your start-up code must not call any Office.js APIs until the library is fully loaded.</span></span> <span data-ttu-id="3eb4c-122">有两种方法可让代码确保已加载库。</span><span class="sxs-lookup"><span data-stu-id="3eb4c-122">There are two ways that your code can ensure that the library is loaded.</span></span> <span data-ttu-id="3eb4c-123">以下各部分介绍了这两种方法：</span><span class="sxs-lookup"><span data-stu-id="3eb4c-123">They are described in the following sections:</span></span> 

- [<span data-ttu-id="3eb4c-124">使用 Office.onReady() 进行初始化</span><span class="sxs-lookup"><span data-stu-id="3eb4c-124">Initialize with Office.onReady()</span></span>](#initialize-with-officeonready)
- [<span data-ttu-id="3eb4c-125">使用 Office.initialize 进行初始化</span><span class="sxs-lookup"><span data-stu-id="3eb4c-125">Initialize with Office.initialize</span></span>](#initialize-with-officeinitialize)

> [!TIP]
> <span data-ttu-id="3eb4c-126">建议使用 `Office.onReady()` 取代 `Office.initialize`。</span><span class="sxs-lookup"><span data-stu-id="3eb4c-126">We recommend that you use `Office.onReady()` instead of `Office.initialize`.</span></span> <span data-ttu-id="3eb4c-127">虽然仍然支持 `Office.initialize`，但使用 `Office.onReady()` 可提供更大的灵活性。</span><span class="sxs-lookup"><span data-stu-id="3eb4c-127">Although `Office.initialize` is still supported, using `Office.onReady()` provides more flexibility.</span></span> <span data-ttu-id="3eb4c-128">可以仅将一个处理程序分配到 `Office.initialize` 并仅由 Office 基础结构调用一次，但可以在代码中的不同位置调用 `Office.onReady()` 并使用不同的回调。</span><span class="sxs-lookup"><span data-stu-id="3eb4c-128">You can assign only one handler to `Office.initialize` and it's called only once by the Office infrastructure, but you can call `Office.onReady()` in different places in your code and use different callbacks.</span></span>
> 
> <span data-ttu-id="3eb4c-129">有关这两种方法之间的差别信息，请参阅 [Office.initialize 和 Office.onReady() 之间的主要差别](#major-differences-between-officeinitialize-and-officeonready)。</span><span class="sxs-lookup"><span data-stu-id="3eb4c-129">For information about the differences in these techniques, see [Major differences between Office.initialize and Office.onReady()](#major-differences-between-officeinitialize-and-officeonready).</span></span>

<span data-ttu-id="3eb4c-130">有关初始化加载项时的事件顺序的更多详细信息，请参阅 [加载 DOM 和运行时环境](loading-the-dom-and-runtime-environment.md)。</span><span class="sxs-lookup"><span data-stu-id="3eb4c-130">For more details about the sequence of events when an add-in is initialized, see [Loading the DOM and runtime environment](loading-the-dom-and-runtime-environment.md).</span></span>

### <a name="initialize-with-officeonready"></a><span data-ttu-id="3eb4c-131">使用 Office.onReady() 进行初始化</span><span class="sxs-lookup"><span data-stu-id="3eb4c-131">Initialize with Office.onReady()</span></span>

<span data-ttu-id="3eb4c-132">`Office.onReady()` 是一个异步方法，它在查看是否加载 Office.js 库时返回 Promise 对象。</span><span class="sxs-lookup"><span data-stu-id="3eb4c-132">`Office.onReady()` is an asynchronous method that returns a Promise object while it checks to see if the Office.js library is loaded.</span></span> <span data-ttu-id="3eb4c-133">仅在加载库时，它才将 Promise 解析为一个对象，该对象指定具有 `Office.HostType` 枚举值（`Excel`、`Word` 等）的 Office 主机应用程序的对象和具有 `Office.PlatformType` 枚举值（`PC`、`Mac`、`OfficeOnline` 等）的平台。</span><span class="sxs-lookup"><span data-stu-id="3eb4c-133">When the library is loaded, it resolves the Promise as an object that specifies the Office host application with an `Office.HostType` enum value (`Excel`, `Word`, etc.) and the platform with an `Office.PlatformType` enum value (`PC`, `Mac`, `OfficeOnline`, etc.).</span></span> <span data-ttu-id="3eb4c-134">如果在调用 `Office.onReady()` 时已加载库，则 Promise 将立即解析。</span><span class="sxs-lookup"><span data-stu-id="3eb4c-134">The Promise resolves immediately if the library is already loaded when `Office.onReady()` is called.</span></span>

<span data-ttu-id="3eb4c-135">调用 `Office.onReady()` 的一种方法是向其传递一个回调方法。</span><span class="sxs-lookup"><span data-stu-id="3eb4c-135">One way to call `Office.onReady()` is to pass it a callback method.</span></span> <span data-ttu-id="3eb4c-136">下面是一个示例：</span><span class="sxs-lookup"><span data-stu-id="3eb4c-136">Here's an example:</span></span>

```js
Office.onReady(function(info) {
    if (info.host === Office.HostType.Excel) {
        // Do Excel-specific initialization (for example, make add-in task pane's
        // appearance compatible with Excel "green").
    }
    if (info.platform === Office.PlatformType.PC) {
        // Make minor layout changes in the task pane.
    }
    console.log(`Office.js is now ready in ${info.host} on ${info.platform}`);
});
```

<span data-ttu-id="3eb4c-137">或者，可以将 `then()` 方法链接到 `Office.onReady()` 的调用，而不是传递回调。</span><span class="sxs-lookup"><span data-stu-id="3eb4c-137">Alternatively, you can chain a `then()` method to the call of `Office.onReady()`, instead of passing a callback.</span></span> <span data-ttu-id="3eb4c-138">例如，以下代码检查用户的 Excel 版本是否支持加载项可能调用的所有 API。</span><span class="sxs-lookup"><span data-stu-id="3eb4c-138">For example, the following code checks to see that the user's version of Excel supports all the APIs that the add-in might call.</span></span>

```js
Office.onReady()
    .then(function() {
        if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
            console.log("Sorry, this add-in only works with newer versions of Excel.");
        }
    });
```

<span data-ttu-id="3eb4c-139">以下是在 TypeScript 中使用 `async` 和 `await` 关键字的同一示例：</span><span class="sxs-lookup"><span data-stu-id="3eb4c-139">Here is the same example using the `async` and `await` keywords in TypeScript:</span></span>

```typescript
(async () => {
    await Office.onReady();
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
})();
```

<span data-ttu-id="3eb4c-140">如果使用的是其他 JavaScript 框架，其中包括它们自己的初始化处理程序或测试，那么它们*通常*应放置在 `Office.onReady()` 的响应内。</span><span class="sxs-lookup"><span data-stu-id="3eb4c-140">If you are using additional JavaScript frameworks that include their own initialization handler or tests, these should *usually* be placed within the response to `Office.onReady()`.</span></span> <span data-ttu-id="3eb4c-141">例如，会对 [JQuery 的](https://jquery.com) `$(document).ready()` 函数进行以下引用：</span><span class="sxs-lookup"><span data-stu-id="3eb4c-141">For example, [JQuery's](https://jquery.com) `$(document).ready()` function would be referenced as follows:</span></span>

```js
Office.onReady(function() {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
});
```

<span data-ttu-id="3eb4c-142">但是，此做法存在例外情况。</span><span class="sxs-lookup"><span data-stu-id="3eb4c-142">However, there are exceptions to this practice.</span></span> <span data-ttu-id="3eb4c-143">例如，假设想要在浏览器中打开加载项（而不是在 Office 主机中旁加载它）以使用浏览器工具调试 UI。</span><span class="sxs-lookup"><span data-stu-id="3eb4c-143">For example, suppose you want to open your add-in in a browser (instead of sideload it in an Office host) in order to debug your UI with browser tools.</span></span> <span data-ttu-id="3eb4c-144">由于 Office.js 将不会在浏览器中加载，所以，`onReady` 将不会运行，且如果在 Office `$(document).ready` 内调用它，则 `onReady` 将不会运行。</span><span class="sxs-lookup"><span data-stu-id="3eb4c-144">Since Office.js won't load in the browser, `onReady` won't run and the `$(document).ready` won't run if it's called inside the Office `onReady`.</span></span> <span data-ttu-id="3eb4c-145">另一个例外情况：希望在加载项加载时在任务窗格中显示进度指示器。</span><span class="sxs-lookup"><span data-stu-id="3eb4c-145">Another exception: you want a progress indicator to appear in the task pane while the add-in is loading.</span></span> <span data-ttu-id="3eb4c-146">在此方案中，代码应调用 jQuery `ready` 并使用其回调来呈现进度指示器。</span><span class="sxs-lookup"><span data-stu-id="3eb4c-146">In this scenario, your code should call the jQuery `ready` and use it's callback to render the progress indicator.</span></span> <span data-ttu-id="3eb4c-147">然后，Office `onReady` 的回调可将进度指示器替换为最终 UI。</span><span class="sxs-lookup"><span data-stu-id="3eb4c-147">Then the Office `onReady`'s callback can replace the progress indicator with the final UI.</span></span> 

### <a name="initialize-with-officeinitialize"></a><span data-ttu-id="3eb4c-148">使用 Office.initialize 进行初始化</span><span class="sxs-lookup"><span data-stu-id="3eb4c-148">Initialize with Office.initialize</span></span>

<span data-ttu-id="3eb4c-149">当 Office.js 库加载并准备好用于用户交互时将触发初始化事件。</span><span class="sxs-lookup"><span data-stu-id="3eb4c-149">An initialize event fires when the Office.js library is fully loaded and ready for user interaction.</span></span> <span data-ttu-id="3eb4c-150">可将处理程序分配到实现初始化逻辑的 `Office.initialize`。</span><span class="sxs-lookup"><span data-stu-id="3eb4c-150">You can assign a handler to `Office.initialize` that implements your initialization logic.</span></span> <span data-ttu-id="3eb4c-151">以下是检查用户的 Excel 版本是否支持加载项可能调用的所有 API 的示例。</span><span class="sxs-lookup"><span data-stu-id="3eb4c-151">The following is an example that checks to see that the user's version of Excel supports all the APIs that the add-in might call.</span></span>

```js
Office.initialize = function () {
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
};
```

<span data-ttu-id="3eb4c-152">如果使用的是其他 JavaScript 框架，其中包含它们自己的初始化处理程序或测试，那么它们*通常*应放置在 `Office.initialize` 事件内。</span><span class="sxs-lookup"><span data-stu-id="3eb4c-152">If you are using additional JavaScript frameworks that include their own initialization handler or tests, these should *usually* be placed within the `Office.initialize` event.</span></span> <span data-ttu-id="3eb4c-153">（但在之前**使用 Office.onReady() 进行初始化**部分中所述的例外情况也适用于此情况。）例如，会对 [JQuery 的](https://jquery.com) `$(document).ready()` 函数进行以下引用：</span><span class="sxs-lookup"><span data-stu-id="3eb4c-153">(But the exceptions described in the **Initialize with Office.onReady()** section earlier apply in this case also.) For example, [JQuery's](https://jquery.com) `$(document).ready()` function would be referenced as follows:</span></span>

```js
Office.initialize = function () {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
  };
```

<span data-ttu-id="3eb4c-154">对于任务窗格和内容加载项，`Office.initialize` 提供了其他 _reason_ 参数。</span><span class="sxs-lookup"><span data-stu-id="3eb4c-154">For task pane and content add-ins, `Office.initialize` provides an additional _reason_ parameter.</span></span> <span data-ttu-id="3eb4c-155">此参数指定如何将加载项添加到当前文档。</span><span class="sxs-lookup"><span data-stu-id="3eb4c-155">This parameter specifies how an add-in was added to the current document.</span></span> <span data-ttu-id="3eb4c-156">可以使用此参数针对首次插入加载项时和加载项已存在于文档中时实施不同的逻辑。</span><span class="sxs-lookup"><span data-stu-id="3eb4c-156">You can use this to provide different logic for when an add-in is first inserted versus when it already existed within the document.</span></span>

```js
Office.initialize = function (reason) {
    $(document).ready(function () {
        switch (reason) {
            case 'inserted': console.log('The add-in was just inserted.');
            case 'documentOpened': console.log('The add-in is already part of the document.');
        }
    });
 };
```

<span data-ttu-id="3eb4c-157">有关详细信息，请参阅 [Office.initialize 事件](https://docs.microsoft.com/javascript/api/office)和 [InitializationReason 枚举](https://docs.microsoft.com/javascript/api/office/office.initializationreason)。</span><span class="sxs-lookup"><span data-stu-id="3eb4c-157">For more information, see [Office.initialize Event](https://docs.microsoft.com/javascript/api/office) and [InitializationReason Enumeration](https://docs.microsoft.com/javascript/api/office/office.initializationreason).</span></span>

### <a name="major-differences-between-officeinitialize-and-officeonready"></a><span data-ttu-id="3eb4c-158">Office.initialize 和 Office.onReady 之间的主要差别</span><span class="sxs-lookup"><span data-stu-id="3eb4c-158">Major differences between Office.initialize and Office.onReady</span></span>

- <span data-ttu-id="3eb4c-159">可以仅将一个处理程序分配到 `Office.initialize` 并仅由 Office 基础结构调用一次，但可以在代码中的不同位置调用 `Office.onReady()` 并使用不同的回调。</span><span class="sxs-lookup"><span data-stu-id="3eb4c-159">You can assign only one handler to `Office.initialize` and it is called, only once, by the Office infrastructure; but you can call `Office.onReady()` in different places in your code and use different callbacks.</span></span> <span data-ttu-id="3eb4c-160">例如，只要自定义脚本使用运行初始化逻辑的回调进行加载，代码就可以调用 `Office.onReady()`。代码还可以在任务窗格中设置一个按钮，其脚本会使用不同的回调调用 `Office.onReady()`。</span><span class="sxs-lookup"><span data-stu-id="3eb4c-160">For example, your code could call `Office.onReady()` as soon as your custom script loads with a callback that runs initialization logic; and your code could also have a button in the task pane, whose script calls `Office.onReady()` with a different callback.</span></span> <span data-ttu-id="3eb4c-161">如果是这样，则会在单击该按钮后运行第二个回调。</span><span class="sxs-lookup"><span data-stu-id="3eb4c-161">If so, the second callback runs when the button is clicked.</span></span>

- <span data-ttu-id="3eb4c-162">`Office.initialize` 事件将在 Office.js 初始化其本身的内部过程的末尾处触发。</span><span class="sxs-lookup"><span data-stu-id="3eb4c-162">The `Office.initialize` event fires at the end of the internal process in which Office.js initializes itself.</span></span> <span data-ttu-id="3eb4c-163">并且它会在内部过程结束后*立即*触发。</span><span class="sxs-lookup"><span data-stu-id="3eb4c-163">And it fires *immediately* after the internal process ends.</span></span> <span data-ttu-id="3eb4c-164">如果将处理程序分配到事件所使用的代码在事件触发后执行的时间过长，则处理程序将不会运行。</span><span class="sxs-lookup"><span data-stu-id="3eb4c-164">If the code in which you assign a handler to the event executes too long after the event fires, then your handler doesn't run.</span></span> <span data-ttu-id="3eb4c-165">例如，如果使用的是 WebPack 任务管理器，则在加载 Office.js 后但在加载自定义 JavaScript 前，它会配置加载项的主页以加载填充代码文件。</span><span class="sxs-lookup"><span data-stu-id="3eb4c-165">For example, if you are using the WebPack task manager, it might configure the add-in's home page to load polyfill files after it loads Office.js but before it loads your custom JavaScript.</span></span> <span data-ttu-id="3eb4c-166">在脚本加载和分配处理程序时，初始化事件已经发生。</span><span class="sxs-lookup"><span data-stu-id="3eb4c-166">By the time your script loads and assigns the handler, the initialize event has already happened.</span></span> <span data-ttu-id="3eb4c-167">但调用 `Office.onReady()` 永远不会“太迟”。</span><span class="sxs-lookup"><span data-stu-id="3eb4c-167">But it is never "too late" to call `Office.onReady()`.</span></span> <span data-ttu-id="3eb4c-168">如果初始化事件已经发生，则回调将立即运行。</span><span class="sxs-lookup"><span data-stu-id="3eb4c-168">If the initialize event has already happened, the callback runs immediately.</span></span>

> [!NOTE]
> <span data-ttu-id="3eb4c-169">即使没有启动逻辑，也应在加载项 JavaScript 加载时调用 `Office.onReady()` 或将空函数分配到 `Office.initialize`。</span><span class="sxs-lookup"><span data-stu-id="3eb4c-169">Even if you have no start-up logic, you should either call `Office.onReady()` or assign an empty function to `Office.initialize` when your add-in JavaScript loads.</span></span> <span data-ttu-id="3eb4c-170">有些 Office 主机和平台组合只有在发生这些情况之一后才会加载任务窗格。</span><span class="sxs-lookup"><span data-stu-id="3eb4c-170">Some Office host and platform combinations won't load the task pane until one of these happens.</span></span> <span data-ttu-id="3eb4c-171">以下示例显示了这两种方法。</span><span class="sxs-lookup"><span data-stu-id="3eb4c-171">The following examples show these two approaches.</span></span>
>
>```js  
>Office.onReady();  
>```    
>
> 
>```js
>Office.initialize = function () {};
>```

## <a name="office-javascript-api-object-model"></a><span data-ttu-id="3eb4c-172">Office JavaScript API 对象模型</span><span class="sxs-lookup"><span data-stu-id="3eb4c-172">Office JavaScript API object model</span></span>

<span data-ttu-id="3eb4c-173">初始化后，加载项可与主机进行交互（例如，Excel、Outlook）。</span><span class="sxs-lookup"><span data-stu-id="3eb4c-173">Once initialized, the add-in can interact with the host (e.g. Excel, Outlook).</span></span> <span data-ttu-id="3eb4c-174">[Office JavaScript API 对象模型](office-javascript-api-object-model.md)页提供了有关特定使用模式的更为详细的信息。</span><span class="sxs-lookup"><span data-stu-id="3eb4c-174">The [Office JavaScript API object model](office-javascript-api-object-model.md) page has more details on specific usage patterns.</span></span> <span data-ttu-id="3eb4c-175">还提供了有关[通用 API](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office) 和主机特定 API 的详细参考文档。</span><span class="sxs-lookup"><span data-stu-id="3eb4c-175">There is also detailed reference documentation for both [Common APIs](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office) and host-specific APIs.</span></span>

## <a name="api-support-matrix"></a><span data-ttu-id="3eb4c-176">API 支持矩阵</span><span class="sxs-lookup"><span data-stu-id="3eb4c-176">API support matrix</span></span>

<span data-ttu-id="3eb4c-177">下表总结了各种类型的加载项（内容、任务窗格和 Outlook）支持的 API 和功能，以及使用[适用于 Office 的 JavaScript API v1.1 支持的 1.1 加载项清单架构和功能](update-your-javascript-api-for-office-and-manifest-schema-version.md)指定加载项支持的 Office 主机应用时，可以托管它们的 Office 应用。</span><span class="sxs-lookup"><span data-stu-id="3eb4c-177">This table summarizes the API and features supported across add-in types (content, task pane, and Outlook) and the Office applications that can host them when you specify the Office host applications your add-in supports by using the [1.1 add-in manifest schema and features supported by v1.1 JavaScript API for Office](update-your-javascript-api-for-office-and-manifest-schema-version.md).</span></span>


|||||||||
|:-----|:-----|:-----:|:-----:|:-----:|:-----:|:-----:|:-----:|
||<span data-ttu-id="3eb4c-178">**主机名**</span><span class="sxs-lookup"><span data-stu-id="3eb4c-178">**Host name**</span></span>|<span data-ttu-id="3eb4c-179">数据库</span><span class="sxs-lookup"><span data-stu-id="3eb4c-179">Database</span></span>|<span data-ttu-id="3eb4c-180">工作簿</span><span class="sxs-lookup"><span data-stu-id="3eb4c-180">Workbook</span></span>|<span data-ttu-id="3eb4c-181">邮箱</span><span class="sxs-lookup"><span data-stu-id="3eb4c-181">Mailbox</span></span>|<span data-ttu-id="3eb4c-182">演示文稿</span><span class="sxs-lookup"><span data-stu-id="3eb4c-182">Presentation</span></span>|<span data-ttu-id="3eb4c-183">文档</span><span class="sxs-lookup"><span data-stu-id="3eb4c-183">Document</span></span>|<span data-ttu-id="3eb4c-184">Project</span><span class="sxs-lookup"><span data-stu-id="3eb4c-184">Project</span></span>|
||<span data-ttu-id="3eb4c-185">**支持的\*\*\*\*主机应用程序**</span><span class="sxs-lookup"><span data-stu-id="3eb4c-185">**Supported** **Host applications**</span></span>|<span data-ttu-id="3eb4c-186">Access Web App</span><span class="sxs-lookup"><span data-stu-id="3eb4c-186">Access web apps</span></span>|<span data-ttu-id="3eb4c-187">Excel、</span><span class="sxs-lookup"><span data-stu-id="3eb4c-187">Excel,</span></span><br/><span data-ttu-id="3eb4c-188">Excel Online</span><span class="sxs-lookup"><span data-stu-id="3eb4c-188">Excel Online</span></span>|<span data-ttu-id="3eb4c-189">Outlook、</span><span class="sxs-lookup"><span data-stu-id="3eb4c-189">Outlook,</span></span><br/><span data-ttu-id="3eb4c-190">Outlook Web App、</span><span class="sxs-lookup"><span data-stu-id="3eb4c-190">Outlook Web App,</span></span><br/><span data-ttu-id="3eb4c-191">适用于设备的 OWA</span><span class="sxs-lookup"><span data-stu-id="3eb4c-191">OWA for Devices</span></span>|<span data-ttu-id="3eb4c-192">PowerPoint、</span><span class="sxs-lookup"><span data-stu-id="3eb4c-192">PowerPoint,</span></span><br/><span data-ttu-id="3eb4c-193">PowerPoint Online</span><span class="sxs-lookup"><span data-stu-id="3eb4c-193">PowerPoint Online</span></span>|<span data-ttu-id="3eb4c-194">Word</span><span class="sxs-lookup"><span data-stu-id="3eb4c-194">Word</span></span>|<span data-ttu-id="3eb4c-195">项目</span><span class="sxs-lookup"><span data-stu-id="3eb4c-195">Project</span></span>|
|<span data-ttu-id="3eb4c-196">**支持的外接程序类型**</span><span class="sxs-lookup"><span data-stu-id="3eb4c-196">**Supported add-in types**</span></span>|<span data-ttu-id="3eb4c-197">内容</span><span class="sxs-lookup"><span data-stu-id="3eb4c-197">Content</span></span>|<span data-ttu-id="3eb4c-198">Y</span><span class="sxs-lookup"><span data-stu-id="3eb4c-198">Y</span></span>|<span data-ttu-id="3eb4c-199">Y</span><span class="sxs-lookup"><span data-stu-id="3eb4c-199">Y</span></span>||<span data-ttu-id="3eb4c-200">Y</span><span class="sxs-lookup"><span data-stu-id="3eb4c-200">Y</span></span>|||
||<span data-ttu-id="3eb4c-201">任务窗格</span><span class="sxs-lookup"><span data-stu-id="3eb4c-201">Task pane</span></span>||<span data-ttu-id="3eb4c-202">Y</span><span class="sxs-lookup"><span data-stu-id="3eb4c-202">Y</span></span>||<span data-ttu-id="3eb4c-203">Y</span><span class="sxs-lookup"><span data-stu-id="3eb4c-203">Y</span></span>|<span data-ttu-id="3eb4c-204">Y</span><span class="sxs-lookup"><span data-stu-id="3eb4c-204">Y</span></span>|<span data-ttu-id="3eb4c-205">Y</span><span class="sxs-lookup"><span data-stu-id="3eb4c-205">Y</span></span>|
||<span data-ttu-id="3eb4c-206">Outlook</span><span class="sxs-lookup"><span data-stu-id="3eb4c-206">Outlook</span></span>|||<span data-ttu-id="3eb4c-207">Y</span><span class="sxs-lookup"><span data-stu-id="3eb4c-207">Y</span></span>||||
|<span data-ttu-id="3eb4c-208">**支持的 API 功能**</span><span class="sxs-lookup"><span data-stu-id="3eb4c-208">**Supported API features**</span></span>|<span data-ttu-id="3eb4c-209">读/写文本</span><span class="sxs-lookup"><span data-stu-id="3eb4c-209">Read/Write Text</span></span>||<span data-ttu-id="3eb4c-210">Y</span><span class="sxs-lookup"><span data-stu-id="3eb4c-210">Y</span></span>||<span data-ttu-id="3eb4c-211">Y</span><span class="sxs-lookup"><span data-stu-id="3eb4c-211">Y</span></span>|<span data-ttu-id="3eb4c-212">Y</span><span class="sxs-lookup"><span data-stu-id="3eb4c-212">Y</span></span>|<span data-ttu-id="3eb4c-213">是</span><span class="sxs-lookup"><span data-stu-id="3eb4c-213">Y</span></span><br/><span data-ttu-id="3eb4c-214">（只读）</span><span class="sxs-lookup"><span data-stu-id="3eb4c-214">(Read only)</span></span>|
||<span data-ttu-id="3eb4c-215">读/写矩阵</span><span class="sxs-lookup"><span data-stu-id="3eb4c-215">Read/Write Matrix</span></span>||<span data-ttu-id="3eb4c-216">Y</span><span class="sxs-lookup"><span data-stu-id="3eb4c-216">Y</span></span>|||<span data-ttu-id="3eb4c-217">Y</span><span class="sxs-lookup"><span data-stu-id="3eb4c-217">Y</span></span>||
||<span data-ttu-id="3eb4c-218">读/写表</span><span class="sxs-lookup"><span data-stu-id="3eb4c-218">Read/Write Table</span></span>||<span data-ttu-id="3eb4c-219">Y</span><span class="sxs-lookup"><span data-stu-id="3eb4c-219">Y</span></span>|||<span data-ttu-id="3eb4c-220">Y</span><span class="sxs-lookup"><span data-stu-id="3eb4c-220">Y</span></span>||
||<span data-ttu-id="3eb4c-221">读/写 HTML</span><span class="sxs-lookup"><span data-stu-id="3eb4c-221">Read/Write HTML</span></span>|||||<span data-ttu-id="3eb4c-222">是</span><span class="sxs-lookup"><span data-stu-id="3eb4c-222">Y</span></span>||
||<span data-ttu-id="3eb4c-223">读/写</span><span class="sxs-lookup"><span data-stu-id="3eb4c-223">Read/Write</span></span><br/><span data-ttu-id="3eb4c-224">Office Open XML</span><span class="sxs-lookup"><span data-stu-id="3eb4c-224">Office Open XML</span></span>|||||<span data-ttu-id="3eb4c-225">是</span><span class="sxs-lookup"><span data-stu-id="3eb4c-225">Y</span></span>||
||<span data-ttu-id="3eb4c-226">读取任务、资源、视图和字段属性</span><span class="sxs-lookup"><span data-stu-id="3eb4c-226">Read task, resource, view, and field properties</span></span>||||||<span data-ttu-id="3eb4c-227">Y</span><span class="sxs-lookup"><span data-stu-id="3eb4c-227">Y</span></span>|
||<span data-ttu-id="3eb4c-228">选择已更改事件</span><span class="sxs-lookup"><span data-stu-id="3eb4c-228">Selection changed events</span></span>||<span data-ttu-id="3eb4c-229">Y</span><span class="sxs-lookup"><span data-stu-id="3eb4c-229">Y</span></span>|||<span data-ttu-id="3eb4c-230">Y</span><span class="sxs-lookup"><span data-stu-id="3eb4c-230">Y</span></span>||
||<span data-ttu-id="3eb4c-231">获取整个文档</span><span class="sxs-lookup"><span data-stu-id="3eb4c-231">Get whole document</span></span>||||<span data-ttu-id="3eb4c-232">Y</span><span class="sxs-lookup"><span data-stu-id="3eb4c-232">Y</span></span>|<span data-ttu-id="3eb4c-233">是</span><span class="sxs-lookup"><span data-stu-id="3eb4c-233">Y</span></span>||
||<span data-ttu-id="3eb4c-234">绑定和绑定事件</span><span class="sxs-lookup"><span data-stu-id="3eb4c-234">Bindings and binding events</span></span>|<span data-ttu-id="3eb4c-235">是</span><span class="sxs-lookup"><span data-stu-id="3eb4c-235">Y</span></span><br/><span data-ttu-id="3eb4c-236">（仅限完全和部分表格绑定）</span><span class="sxs-lookup"><span data-stu-id="3eb4c-236">(Only full and partial table bindings)</span></span>|<span data-ttu-id="3eb4c-237">是</span><span class="sxs-lookup"><span data-stu-id="3eb4c-237">Y</span></span>|||<span data-ttu-id="3eb4c-238">是</span><span class="sxs-lookup"><span data-stu-id="3eb4c-238">Y</span></span>||
||<span data-ttu-id="3eb4c-239">读/写自定义 XML 部分</span><span class="sxs-lookup"><span data-stu-id="3eb4c-239">Read/Write Custom XML Parts</span></span>|||||<span data-ttu-id="3eb4c-240">是</span><span class="sxs-lookup"><span data-stu-id="3eb4c-240">Y</span></span>||
||<span data-ttu-id="3eb4c-241">暂留加载项状态数据（设置）</span><span class="sxs-lookup"><span data-stu-id="3eb4c-241">Persist add-in state data (settings)</span></span>|<span data-ttu-id="3eb4c-242">是</span><span class="sxs-lookup"><span data-stu-id="3eb4c-242">Y</span></span><br/><span data-ttu-id="3eb4c-243">（每主机加载项）</span><span class="sxs-lookup"><span data-stu-id="3eb4c-243">(Per host add-in)</span></span>|<span data-ttu-id="3eb4c-244">是</span><span class="sxs-lookup"><span data-stu-id="3eb4c-244">Y</span></span><br/><span data-ttu-id="3eb4c-245">（每文档）</span><span class="sxs-lookup"><span data-stu-id="3eb4c-245">(Per document)</span></span>|<span data-ttu-id="3eb4c-246">是</span><span class="sxs-lookup"><span data-stu-id="3eb4c-246">Y</span></span><br/><span data-ttu-id="3eb4c-247">（每邮箱）</span><span class="sxs-lookup"><span data-stu-id="3eb4c-247">(Per mailbox)</span></span>|<span data-ttu-id="3eb4c-248">是</span><span class="sxs-lookup"><span data-stu-id="3eb4c-248">Y</span></span><br/><span data-ttu-id="3eb4c-249">（每文档）</span><span class="sxs-lookup"><span data-stu-id="3eb4c-249">(Per document)</span></span>|<span data-ttu-id="3eb4c-250">是</span><span class="sxs-lookup"><span data-stu-id="3eb4c-250">Y</span></span><br/><span data-ttu-id="3eb4c-251">（每文档）</span><span class="sxs-lookup"><span data-stu-id="3eb4c-251">(Per document)</span></span>||
||<span data-ttu-id="3eb4c-252">设置更改事件</span><span class="sxs-lookup"><span data-stu-id="3eb4c-252">Settings changed events</span></span>|<span data-ttu-id="3eb4c-253">Y</span><span class="sxs-lookup"><span data-stu-id="3eb4c-253">Y</span></span>|<span data-ttu-id="3eb4c-254">Y</span><span class="sxs-lookup"><span data-stu-id="3eb4c-254">Y</span></span>||<span data-ttu-id="3eb4c-255">Y</span><span class="sxs-lookup"><span data-stu-id="3eb4c-255">Y</span></span>|<span data-ttu-id="3eb4c-256">是</span><span class="sxs-lookup"><span data-stu-id="3eb4c-256">Y</span></span>||
||<span data-ttu-id="3eb4c-257">获取活动视图模式</span><span class="sxs-lookup"><span data-stu-id="3eb4c-257">Get active view mode</span></span><br/><span data-ttu-id="3eb4c-258">和视图更改事件</span><span class="sxs-lookup"><span data-stu-id="3eb4c-258">and view changed events</span></span>||||<span data-ttu-id="3eb4c-259">是</span><span class="sxs-lookup"><span data-stu-id="3eb4c-259">Y</span></span>|||
||<span data-ttu-id="3eb4c-260">转到文档中</span><span class="sxs-lookup"><span data-stu-id="3eb4c-260">Navigate to locations</span></span><br/><span data-ttu-id="3eb4c-261">的相应位置</span><span class="sxs-lookup"><span data-stu-id="3eb4c-261">in the document</span></span>||<span data-ttu-id="3eb4c-262">是</span><span class="sxs-lookup"><span data-stu-id="3eb4c-262">Y</span></span>||<span data-ttu-id="3eb4c-263">Y</span><span class="sxs-lookup"><span data-stu-id="3eb4c-263">Y</span></span>|<span data-ttu-id="3eb4c-264">是</span><span class="sxs-lookup"><span data-stu-id="3eb4c-264">Y</span></span>||
||<span data-ttu-id="3eb4c-265">使用规则和 RegEx </span><span class="sxs-lookup"><span data-stu-id="3eb4c-265">Activate contextually</span></span><br/><span data-ttu-id="3eb4c-266">执行上下文式激活</span><span class="sxs-lookup"><span data-stu-id="3eb4c-266">using rules and RegEx</span></span>|||<span data-ttu-id="3eb4c-267">是</span><span class="sxs-lookup"><span data-stu-id="3eb4c-267">Y</span></span>||||
||<span data-ttu-id="3eb4c-268">读取项目属性</span><span class="sxs-lookup"><span data-stu-id="3eb4c-268">Read Item properties</span></span>|||<span data-ttu-id="3eb4c-269">Y</span><span class="sxs-lookup"><span data-stu-id="3eb4c-269">Y</span></span>||||
||<span data-ttu-id="3eb4c-270">读取用户配置文件</span><span class="sxs-lookup"><span data-stu-id="3eb4c-270">Read User profile</span></span>|||<span data-ttu-id="3eb4c-271">Y</span><span class="sxs-lookup"><span data-stu-id="3eb4c-271">Y</span></span>||||
||<span data-ttu-id="3eb4c-272">获取附件</span><span class="sxs-lookup"><span data-stu-id="3eb4c-272">Get attachments</span></span>|||<span data-ttu-id="3eb4c-273">Y</span><span class="sxs-lookup"><span data-stu-id="3eb4c-273">Y</span></span>||||
||<span data-ttu-id="3eb4c-274">获取用户标识令牌</span><span class="sxs-lookup"><span data-stu-id="3eb4c-274">Get User identity token</span></span>|||<span data-ttu-id="3eb4c-275">Y</span><span class="sxs-lookup"><span data-stu-id="3eb4c-275">Y</span></span>||||
||<span data-ttu-id="3eb4c-276">调用 Exchange Web 服务</span><span class="sxs-lookup"><span data-stu-id="3eb4c-276">Call Exchange Web Services</span></span>|||<span data-ttu-id="3eb4c-277">Y</span><span class="sxs-lookup"><span data-stu-id="3eb4c-277">Y</span></span>||||
