---
title: 了解适用于 Office 的 JavaScript API
description: ''
ms.date: 10/17/2018
ms.openlocfilehash: 14de5d8bab791d0954179c21163ba0a08824b834
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/28/2018
ms.locfileid: "27458102"
---
# <a name="understanding-the-javascript-api-for-office"></a><span data-ttu-id="581a9-102">了解适用于 Office 的 JavaScript API</span><span class="sxs-lookup"><span data-stu-id="581a9-102">Understanding the JavaScript API for Office</span></span>

<span data-ttu-id="581a9-p101">本文提供了有关适用于 Office 的 JavaScript API 的信息以及使用方法。有关参考信息，请参阅 [适用于 Office 的 JavaScript API](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)。有关将 Visual Studio 项目文件更新到适用于 Office 的 JavaScript API 的最新当前版本的信息，请参阅 [更新适用于 Office 的 JavaScript API 版本和清单架构文件](update-your-javascript-api-for-office-and-manifest-schema-version.md)。</span><span class="sxs-lookup"><span data-stu-id="581a9-p101">This article provides information about the JavaScript API for Office and how to use it. For reference information, see [JavaScript API for Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office). For information about updating Visual Studio project files to the most current version of the JavaScript API for Office, see [Update the version of your JavaScript API for Office and manifest schema files](update-your-javascript-api-for-office-and-manifest-schema-version.md).</span></span>

> [!NOTE]
> <span data-ttu-id="581a9-p102">如果计划将加载项[发布](../publish/publish.md)到 AppSource 并适用于 Office 体验，请务必遵循 [AppSource 验证策略](https://docs.microsoft.com/office/dev/store/validation-policies)。例如，加载项必须适用于支持已定义方法的所有平台，才能通过验证（有关详细信息，请参阅[第 4.12 部分](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably)以及 [Office 加载项主机和可用性](../overview/office-add-in-availability.md)页面）。</span><span class="sxs-lookup"><span data-stu-id="581a9-p102">If you plan to [publish](../publish/publish.md) your add-in to AppSource and make it available within the Office experience, make sure that you conform to the [AppSource validation policies](https://docs.microsoft.com/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)).</span></span> 

## <a name="referencing-the-javascript-api-for-office-library-in-your-add-in"></a><span data-ttu-id="581a9-108">在加载项中引用适用于 Office 的 JavaScript API 库</span><span class="sxs-lookup"><span data-stu-id="581a9-108">Referencing the JavaScript API for Office library in your add-in</span></span>

<span data-ttu-id="581a9-p103">[适用于 Office 的 JavaScript](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office) 库包含 Office.js 文件和关联的特定于主机应用程序的 .js 文件，例如 Excel-15.js 和 Outlook-15.js。引用该 API 最简单的方法是通过添加以下 `<script>` 到你的页面的 `<head>` 标记来使用我们的 CDN：</span><span class="sxs-lookup"><span data-stu-id="581a9-p103">The [JavaScript API for Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office) library consists of the Office.js file and associated host application-specific .js files, such as Excel-15.js and Outlook-15.js. The simplest method of referencing the API is using our CDN by adding the following `<script>` to your page's `<head>` tag:</span></span>  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

<span data-ttu-id="581a9-111">这将在加载项首次加载时下载并缓存适用于 Office 的 JavaScript API 文件，以确保对特定版本使用 Office.js 及其关联文件的最新实现。</span><span class="sxs-lookup"><span data-stu-id="581a9-111">This will download and cache the JavaScript API for Office files the first time your add-in loads to make sure that it is using the most up-to-date implementation of Office.js and its associated files for the specified version.</span></span>

<span data-ttu-id="581a9-112">有关 Office.js CDN 的详细信息，包括如何处理版本控制和向后兼容性，请参阅[从适用于 Office 的 JavaScript API 的内容传送网络 (CDN) 引用适用于 Office 的 JavaScript API 库](referencing-the-javascript-api-for-office-library-from-its-cdn.md)。</span><span class="sxs-lookup"><span data-stu-id="581a9-112">For more details around the Office.js CDN, including how versioning and backward compatibility is handled, see [Referencing the JavaScript API for Office library from its content delivery network (CDN)](referencing-the-javascript-api-for-office-library-from-its-cdn.md).</span></span>

## <a name="initializing-your-add-in"></a><span data-ttu-id="581a9-113">初始化加载项</span><span class="sxs-lookup"><span data-stu-id="581a9-113">Initializing your add-in</span></span>

<span data-ttu-id="581a9-114">**适用于：** 所有加载项类型</span><span class="sxs-lookup"><span data-stu-id="581a9-114">**Applies to:** All add-in types</span></span>

<span data-ttu-id="581a9-115">Office 加载项通常使用启动逻辑执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="581a9-115">Office Add-ins often have start-up logic to do things such as:</span></span>

- <span data-ttu-id="581a9-116">检查用户的 Office 版本是否支持代码调用的所有 Office API。</span><span class="sxs-lookup"><span data-stu-id="581a9-116">Check that the user's version of Office will support all the Office APIs that your code calls.</span></span>

- <span data-ttu-id="581a9-117">确保存在某些项目，例如具有特定名称的工作表。</span><span class="sxs-lookup"><span data-stu-id="581a9-117">Ensure the existence of certain artifacts, such as worksheet with a specific name.</span></span>

- <span data-ttu-id="581a9-118">提示用户选择 Excel 中的某些单元格，然后插入使用这些所选值进行初始化的图表。</span><span class="sxs-lookup"><span data-stu-id="581a9-118">Prompting the user to select some cells in Excel, and then inserting a chart initialized with those selected values.</span></span>

- <span data-ttu-id="581a9-119">建立绑定。</span><span class="sxs-lookup"><span data-stu-id="581a9-119">Establish bindings.</span></span>

- <span data-ttu-id="581a9-120">使用 Office 对话框 API 提示用户使用默认的加载项设置值。</span><span class="sxs-lookup"><span data-stu-id="581a9-120">Use the Office dialog API to prompt the user for default add-in settings values.</span></span>

<span data-ttu-id="581a9-121">但在完全加载库前，启动代码不得调用任何 Office.js API。</span><span class="sxs-lookup"><span data-stu-id="581a9-121">But your start-up code must not call any Office.js APIs until the library is fully loaded.</span></span> <span data-ttu-id="581a9-122">有两种方法可让代码确保已加载库。</span><span class="sxs-lookup"><span data-stu-id="581a9-122">There are two ways that your code can ensure that the library is loaded.</span></span> <span data-ttu-id="581a9-123">以下各部分介绍了这两种方法：</span><span class="sxs-lookup"><span data-stu-id="581a9-123">They are described in the following sections:</span></span> 

- [<span data-ttu-id="581a9-124">使用 Office.onReady() 进行初始化</span><span class="sxs-lookup"><span data-stu-id="581a9-124">Initialize with Office.onReady()</span></span>](#initialize-with-officeonready)
- [<span data-ttu-id="581a9-125">使用 Office.initialize 进行初始化</span><span class="sxs-lookup"><span data-stu-id="581a9-125">Initialize with Office.initialize</span></span>](#initialize-with-officeinitialize)

<span data-ttu-id="581a9-126">有关这两种方法之间的差别信息，请参阅 [Office.initialize 和 Office.onReady() 之间的主要差别](#major-differences-between-officeinitialize-and-officeonready)。</span><span class="sxs-lookup"><span data-stu-id="581a9-126">For information about the differences in these techniques, see [Major differences between Office.initialize and Office.onReady()](#major-differences-between-officeinitialize-and-officeonready).</span></span> <span data-ttu-id="581a9-127">有关初始化加载项时的事件顺序的更多详细信息，请参阅 [加载 DOM 和运行时环境](loading-the-dom-and-runtime-environment.md)。</span><span class="sxs-lookup"><span data-stu-id="581a9-127">For more details about the sequence of events when an add-in is initialized, see [Loading the DOM and runtime environment](loading-the-dom-and-runtime-environment.md).</span></span>

### <a name="initialize-with-officeonready"></a><span data-ttu-id="581a9-128">使用 Office.onReady() 进行初始化</span><span class="sxs-lookup"><span data-stu-id="581a9-128">Initialize with Office.onReady()</span></span>

<span data-ttu-id="581a9-129">`Office.onReady()` 是一个异步方法，它在查看是否完全加载 Office.js 库时返回 Promise 对象。</span><span class="sxs-lookup"><span data-stu-id="581a9-129">`Office.onReady()` is an asynchronous method that returns a Promise object while it checks to see if the Office.js library is fully loaded.</span></span> <span data-ttu-id="581a9-130">仅在加载库时，它才将 Promise 解析为一个对象，该对象指定具有 `Office.HostType` 枚举值（`Excel`、`Word` 等）的 Office 主机应用程序的对象和具有 `Office.PlatformType` 枚举值（`PC`、`Mac`、`OfficeOnline` 等）的平台。</span><span class="sxs-lookup"><span data-stu-id="581a9-130">When, and only when, the library is loaded, it resolves the Promise as an object that specifies the Office host application with an `Office.HostType` enum value (`Excel`, `Word`, etc.) and the platform with an `Office.PlatformType` enum value (`PC`, `Mac`, `OfficeOnline`, etc.).</span></span> <span data-ttu-id="581a9-131">如果在调用 `Office.onReady()` 时已加载库，则 Promise 将立即解析。</span><span class="sxs-lookup"><span data-stu-id="581a9-131">If the library is already loaded when `Office.onReady()` is called, then the Promise resolves immediately.</span></span>

<span data-ttu-id="581a9-132">调用 `Office.onReady()` 的一种方法是向其传递一个回调方法。</span><span class="sxs-lookup"><span data-stu-id="581a9-132">One way to call `Office.onReady()` is to pass it a callback method.</span></span> <span data-ttu-id="581a9-133">下面是一个示例：</span><span class="sxs-lookup"><span data-stu-id="581a9-133">Here's an example:</span></span>

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

<span data-ttu-id="581a9-134">或者，可以将 `then()` 方法链接到 `Office.onReady()` 的调用，而不是传递回调。</span><span class="sxs-lookup"><span data-stu-id="581a9-134">Alternatively, you can chain a `then()` method to the call of `Office.onReady()`, instead of passing a callback.</span></span> <span data-ttu-id="581a9-135">例如，以下代码检查用户的 Excel 版本是否支持加载项可能调用的所有 API。</span><span class="sxs-lookup"><span data-stu-id="581a9-135">For example, the following code checks to see that the user's version of Excel supports all the APIs that the add-in might call.</span></span>

```js
Office.onReady()
    .then(function() {
        if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
            console.log("Sorry, this add-in only works with newer versions of Excel.");
        }
    });
```

<span data-ttu-id="581a9-136">以下是在 TypeScript 中使用 `async` 和 `await` 关键字的同一示例：</span><span class="sxs-lookup"><span data-stu-id="581a9-136">Here is the same example using the `async` and `await` keywords in TypeScript:</span></span>

```typescript
(async () => {
    await Office.onReady();
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
})();
```

<span data-ttu-id="581a9-137">如果使用的是其他 JavaScript 框架，其中包括它们自己的初始化处理程序或测试，那么它们*通常*应放置在 `Office.onReady()` 的响应内。</span><span class="sxs-lookup"><span data-stu-id="581a9-137">If you are using additional JavaScript frameworks that include their own initialization handler or tests, these should *usually* be placed within the response to `Office.onReady()`.</span></span> <span data-ttu-id="581a9-138">例如，会对 [JQuery 的](https://jquery.com) `$(document).ready()` 函数进行以下引用：</span><span class="sxs-lookup"><span data-stu-id="581a9-138">For example, [JQuery's](https://jquery.com) `$(document).ready()` function would be referenced as follows:</span></span>

```js
Office.onReady(function() {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
});
```

<span data-ttu-id="581a9-139">但是，此做法存在例外情况。</span><span class="sxs-lookup"><span data-stu-id="581a9-139">However, there are exceptions to this practice.</span></span> <span data-ttu-id="581a9-140">例如，假设想要在浏览器中打开加载项（而不是在 Office 主机中旁加载它）以使用浏览器工具调试 UI。</span><span class="sxs-lookup"><span data-stu-id="581a9-140">For example, suppose you want to open your add-in in a browser (instead of sideload it in an Office host) in order to debug your UI with browser tools.</span></span> <span data-ttu-id="581a9-141">由于 Office.js 将不会在浏览器中加载，所以，`onReady` 将不会运行，且如果在 Office `$(document).ready` 内调用它，则 `onReady` 将不会运行。</span><span class="sxs-lookup"><span data-stu-id="581a9-141">Since Office.js won't load in the browser, `onReady` won't run and the `$(document).ready` won't run if it's called inside the Office `onReady`.</span></span> <span data-ttu-id="581a9-142">另一个例外情况：希望在加载项加载时在任务窗格中显示进度指示器。</span><span class="sxs-lookup"><span data-stu-id="581a9-142">Another exception: you want a progress indicator to appear in the task pane while the add-in is loading.</span></span> <span data-ttu-id="581a9-143">在此方案中，代码应调用 jQuery `ready` 并使用其回调来呈现进度指示器。</span><span class="sxs-lookup"><span data-stu-id="581a9-143">In this scenario, your code should call the jQuery `ready` and use it's callback to render the progress indicator.</span></span> <span data-ttu-id="581a9-144">然后，Office `onReady` 的回调可将进度指示器替换为最终 UI。</span><span class="sxs-lookup"><span data-stu-id="581a9-144">Then the Office `onReady`'s callback can replace the progress indicator with the final UI.</span></span> 

### <a name="initialize-with-officeinitialize"></a><span data-ttu-id="581a9-145">使用 Office.initialize 进行初始化</span><span class="sxs-lookup"><span data-stu-id="581a9-145">Initialize with Office.initialize</span></span>

<span data-ttu-id="581a9-146">当 Office.js 库完全加载并准备好用于用户交互时将触发初始化事件。</span><span class="sxs-lookup"><span data-stu-id="581a9-146">An initialize event fires when the Office.js library is fully loaded and ready for user interaction.</span></span> <span data-ttu-id="581a9-147">可将处理程序分配到实现初始化逻辑的 `Office.initialize`。</span><span class="sxs-lookup"><span data-stu-id="581a9-147">You can assign a handler to `Office.initialize` that implements your initialization logic.</span></span> <span data-ttu-id="581a9-148">以下是检查用户的 Excel 版本是否支持加载项可能调用的所有 API 的示例。</span><span class="sxs-lookup"><span data-stu-id="581a9-148">The following is an example that checks to see that the user's version of Excel supports all the APIs that the add-in might call.</span></span>

```js
Office.initialize = function () {
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
};
```

<span data-ttu-id="581a9-149">如果使用的是其他 JavaScript 框架，其中包含它们自己的初始化处理程序或测试，那么它们*通常*应放置在 `Office.initialize` 事件内。</span><span class="sxs-lookup"><span data-stu-id="581a9-149">If you are using additional JavaScript frameworks that include their own initialization handler or tests, these should *usually* be placed within the `Office.initialize` event.</span></span> <span data-ttu-id="581a9-150">（但在之前**使用 Office.onReady() 进行初始化**部分中所述的例外情况也适用于此情况。）例如，会对 [JQuery 的](https://jquery.com) `$(document).ready()` 函数进行以下引用：</span><span class="sxs-lookup"><span data-stu-id="581a9-150">(But the exceptions described in the **Initialize with Office.onReady()** section earlier apply in this case also.) For example, [JQuery's](https://jquery.com) `$(document).ready()` function would be referenced as follows:</span></span>

```js
Office.initialize = function () {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
  };
```

<span data-ttu-id="581a9-151">对于任务窗格和内容加载项，`Office.initialize` 提供了其他 _reason_ 参数。</span><span class="sxs-lookup"><span data-stu-id="581a9-151">For task pane and content add-ins, `Office.initialize` provides an additional _reason_ parameter.</span></span> <span data-ttu-id="581a9-152">此参数指定如何将加载项添加到当前文档。</span><span class="sxs-lookup"><span data-stu-id="581a9-152">This parameter specifies how an add-in was added to the current document.</span></span> <span data-ttu-id="581a9-153">可以使用此参数针对首次插入加载项时和加载项已存在于文档中时实施不同的逻辑。</span><span class="sxs-lookup"><span data-stu-id="581a9-153">You can use this to provide different logic for when an add-in is first inserted versus when it already existed within the document.</span></span>

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

<span data-ttu-id="581a9-154">有关详细信息，请参阅 [Office.initialize 事件](https://docs.microsoft.com/javascript/api/office)和 [InitializationReason 枚举](https://docs.microsoft.com/javascript/api/office/office.initializationreason)。</span><span class="sxs-lookup"><span data-stu-id="581a9-154">For more information, see [Office.initialize Event](https://docs.microsoft.com/javascript/api/office) and [InitializationReason Enumeration](https://docs.microsoft.com/javascript/api/office/office.initializationreason).</span></span>

> [!NOTE]
> <span data-ttu-id="581a9-155">目前，无论是否同时调用 `Office.onReady()`均必须设置 `Office.Initialize`。</span><span class="sxs-lookup"><span data-stu-id="581a9-155">Currently, you must set `Office.Initialize`, regardless of whether `Office.onReady()` is also called.</span></span> <span data-ttu-id="581a9-156">如果无需使用 `Office.Initialize`，可以将其设置为一个空函数，如以下示例中所示。</span><span class="sxs-lookup"><span data-stu-id="581a9-156">If you have no use for `Office.Initialize`, you can set it to an empty function as shown in the following example.</span></span>
> 
>```js
>Office.initialize = function () {};
>```

### <a name="major-differences-between-officeinitialize-and-officeonready"></a><span data-ttu-id="581a9-157">Office.initialize 和 Office.onReady 之间的主要差别</span><span class="sxs-lookup"><span data-stu-id="581a9-157">Major differences between Office.initialize and Office.onReady</span></span>

- <span data-ttu-id="581a9-158">可以仅将一个处理程序分配到 `Office.initialize` 并仅由 Office 基础结构调用一次，但可以在代码中的不同位置调用 `Office.onReady()` 并使用不同的回调。</span><span class="sxs-lookup"><span data-stu-id="581a9-158">You can assign only one handler to `Office.initialize` and it is called, only once, by the Office infrastructure; but you can call `Office.onReady()` in different places in your code and use different callbacks.</span></span> <span data-ttu-id="581a9-159">例如，只要自定义脚本使用运行初始化逻辑的回调进行加载，代码就可以调用 `Office.onReady()`。代码还可以在任务窗格中设置一个按钮，其脚本会使用不同的回调调用 `Office.onReady()`。</span><span class="sxs-lookup"><span data-stu-id="581a9-159">For example, your code could call `Office.onReady()` as soon as your custom script loads with a callback that runs initialization logic; and your code could also have a button in the task pane, whose script calls `Office.onReady()` with a different callback.</span></span> <span data-ttu-id="581a9-160">如果是这样，则会在单击该按钮后运行第二个回调。</span><span class="sxs-lookup"><span data-stu-id="581a9-160">If so, the second callback runs when the button is clicked.</span></span>

- <span data-ttu-id="581a9-161">`Office.initialize` 事件将在 Office.js 初始化其本身的内部过程的末尾处触发。</span><span class="sxs-lookup"><span data-stu-id="581a9-161">The `Office.initialize` event fires at the end of the internal process in which Office.js initializes itself.</span></span> <span data-ttu-id="581a9-162">并且它会在内部过程结束后*立即*触发。</span><span class="sxs-lookup"><span data-stu-id="581a9-162">And it fires *immediately* after the internal process ends.</span></span> <span data-ttu-id="581a9-163">如果将处理程序分配到事件所使用的代码在事件触发后执行的时间过长，则处理程序将不会运行。</span><span class="sxs-lookup"><span data-stu-id="581a9-163">If the code in which you assign a handler to the event executes too long after the event fires, then your handler doesn't run.</span></span> <span data-ttu-id="581a9-164">例如，如果使用的是 WebPack 任务管理器，则在加载 Office.js 后但在加载自定义 JavaScript 前，它会配置加载项的主页以加载填充代码文件。</span><span class="sxs-lookup"><span data-stu-id="581a9-164">For example, if you are using the WebPack task manager, it might configure the add-in's home page to load polyfill files after it loads Office.js but before it loads your custom JavaScript.</span></span> <span data-ttu-id="581a9-165">在脚本加载和分配处理程序时，初始化事件已经发生。</span><span class="sxs-lookup"><span data-stu-id="581a9-165">By the time your script loads and assigns the handler, the initialize event has already happened.</span></span> <span data-ttu-id="581a9-166">但调用 `Office.onReady()` 永远不会“太迟”。</span><span class="sxs-lookup"><span data-stu-id="581a9-166">But it is never "too late" to call `Office.onReady()`.</span></span> <span data-ttu-id="581a9-167">如果初始化事件已经发生，则回调将立即运行。</span><span class="sxs-lookup"><span data-stu-id="581a9-167">If the initialize event has already happened, the callback runs immediately.</span></span>

> [!NOTE]
> <span data-ttu-id="581a9-168">即使没有启动逻辑，也应在加载项 JavaScript 加载时将空函数分配到 `Office.initialize`，如以下示例中所示。</span><span class="sxs-lookup"><span data-stu-id="581a9-168">Even if you have no start-up logic, you should assign an empty function to `Office.initialize` when your add-in JavaScript loads, as shown in the following example.</span></span> <span data-ttu-id="581a9-169">在初始化事件触发且指定的事件处理程序函数运行前，某些 Office 主机和平台组合将不会加载任务窗格。</span><span class="sxs-lookup"><span data-stu-id="581a9-169">Some Office host and platform combinations won't load the task pane until the initialize event fires and the specified event handler function runs.</span></span>
> 
>```js
>Office.initialize = function () {};
>```

## <a name="office-javascript-api-object-model"></a><span data-ttu-id="581a9-170">Office JavaScript API 对象模型</span><span class="sxs-lookup"><span data-stu-id="581a9-170">Office JavaScript API object model</span></span>

<span data-ttu-id="581a9-171">初始化后，加载项可与主机进行交互（例如，Excel、Outlook）。</span><span class="sxs-lookup"><span data-stu-id="581a9-171">Once initialized, the add-in can interact with the host (e.g. Excel, Outlook).</span></span> <span data-ttu-id="581a9-172">[Office JavaScript API 对象模型](office-javascript-api-object-model.md)页提供了有关特定使用模式的更为详细的信息。</span><span class="sxs-lookup"><span data-stu-id="581a9-172">The [Office JavaScript API object model](office-javascript-api-object-model.md) page has more details on specific usage patterns.</span></span> <span data-ttu-id="581a9-173">还提供了有关[通用 API](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office) 和主机特定 API 的详细参考文档。</span><span class="sxs-lookup"><span data-stu-id="581a9-173">There is also detailed reference documentation for both [shared APIs](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office) and specific hosts.</span></span>

## <a name="api-support-matrix"></a><span data-ttu-id="581a9-174">API 支持矩阵</span><span class="sxs-lookup"><span data-stu-id="581a9-174">API support matrix</span></span>

<span data-ttu-id="581a9-175">下表总结了各种类型的加载项（内容、任务窗格和 Outlook）支持的 API 和功能，以及使用[适用于 Office 的 JavaScript API v1.1 支持的 1.1 加载项清单架构和功能](update-your-javascript-api-for-office-and-manifest-schema-version.md)指定加载项支持的 Office 主机应用时，可以托管它们的 Office 应用。</span><span class="sxs-lookup"><span data-stu-id="581a9-175">This table summarizes the API and features supported across add-in types (content, task pane, and Outlook) and the Office applications that can host them when you specify the Office host applications your add-in supports by using the [1.1 add-in manifest schema and features supported by v1.1 JavaScript API for Office](update-your-javascript-api-for-office-and-manifest-schema-version.md).</span></span>


|||||||||
|:-----|:-----|:-----:|:-----:|:-----:|:-----:|:-----:|:-----:|
||<span data-ttu-id="581a9-176">**主机名**</span><span class="sxs-lookup"><span data-stu-id="581a9-176">**Host name**</span></span>|<span data-ttu-id="581a9-177">数据库</span><span class="sxs-lookup"><span data-stu-id="581a9-177">Database</span></span>|<span data-ttu-id="581a9-178">工作簿</span><span class="sxs-lookup"><span data-stu-id="581a9-178">Workbook</span></span>|<span data-ttu-id="581a9-179">邮箱</span><span class="sxs-lookup"><span data-stu-id="581a9-179">Mailbox</span></span>|<span data-ttu-id="581a9-180">演示文稿</span><span class="sxs-lookup"><span data-stu-id="581a9-180">Presentation</span></span>|<span data-ttu-id="581a9-181">文档</span><span class="sxs-lookup"><span data-stu-id="581a9-181">Document</span></span>|<span data-ttu-id="581a9-182">Project</span><span class="sxs-lookup"><span data-stu-id="581a9-182">Project</span></span>|
||<span data-ttu-id="581a9-183">**支持的\*\*\*\*主机应用程序**</span><span class="sxs-lookup"><span data-stu-id="581a9-183">**Supported** **Host applications**</span></span>|<span data-ttu-id="581a9-184">Access Web App</span><span class="sxs-lookup"><span data-stu-id="581a9-184">Access web apps</span></span>|<span data-ttu-id="581a9-185">Excel、</span><span class="sxs-lookup"><span data-stu-id="581a9-185">Excel,</span></span><br/><span data-ttu-id="581a9-186">Excel Online</span><span class="sxs-lookup"><span data-stu-id="581a9-186">Excel Online</span></span>|<span data-ttu-id="581a9-187">Outlook、</span><span class="sxs-lookup"><span data-stu-id="581a9-187">Outlook,</span></span><br/><span data-ttu-id="581a9-188">Outlook Web App、</span><span class="sxs-lookup"><span data-stu-id="581a9-188">Outlook Web App,</span></span><br/><span data-ttu-id="581a9-189">适用于设备的 OWA</span><span class="sxs-lookup"><span data-stu-id="581a9-189">OWA for Devices</span></span>|<span data-ttu-id="581a9-190">PowerPoint、</span><span class="sxs-lookup"><span data-stu-id="581a9-190">PowerPoint,</span></span><br/><span data-ttu-id="581a9-191">PowerPoint Online</span><span class="sxs-lookup"><span data-stu-id="581a9-191">PowerPoint Online</span></span>|<span data-ttu-id="581a9-192">Word</span><span class="sxs-lookup"><span data-stu-id="581a9-192">Word</span></span>|<span data-ttu-id="581a9-193">项目</span><span class="sxs-lookup"><span data-stu-id="581a9-193">Project</span></span>|
|<span data-ttu-id="581a9-194">**支持的外接程序类型**</span><span class="sxs-lookup"><span data-stu-id="581a9-194">**Supported add-in types**</span></span>|<span data-ttu-id="581a9-195">内容</span><span class="sxs-lookup"><span data-stu-id="581a9-195">Content</span></span>|<span data-ttu-id="581a9-196">Y</span><span class="sxs-lookup"><span data-stu-id="581a9-196">Y</span></span>|<span data-ttu-id="581a9-197">Y</span><span class="sxs-lookup"><span data-stu-id="581a9-197">Y</span></span>||<span data-ttu-id="581a9-198">Y</span><span class="sxs-lookup"><span data-stu-id="581a9-198">Y</span></span>|||
||<span data-ttu-id="581a9-199">任务窗格</span><span class="sxs-lookup"><span data-stu-id="581a9-199">Task pane</span></span>||<span data-ttu-id="581a9-200">Y</span><span class="sxs-lookup"><span data-stu-id="581a9-200">Y</span></span>||<span data-ttu-id="581a9-201">Y</span><span class="sxs-lookup"><span data-stu-id="581a9-201">Y</span></span>|<span data-ttu-id="581a9-202">Y</span><span class="sxs-lookup"><span data-stu-id="581a9-202">Y</span></span>|<span data-ttu-id="581a9-203">Y</span><span class="sxs-lookup"><span data-stu-id="581a9-203">Y</span></span>|
||<span data-ttu-id="581a9-204">Outlook</span><span class="sxs-lookup"><span data-stu-id="581a9-204">Outlook</span></span>|||<span data-ttu-id="581a9-205">Y</span><span class="sxs-lookup"><span data-stu-id="581a9-205">Y</span></span>||||
|<span data-ttu-id="581a9-206">**支持的 API 功能**</span><span class="sxs-lookup"><span data-stu-id="581a9-206">**Supported API features**</span></span>|<span data-ttu-id="581a9-207">读/写文本</span><span class="sxs-lookup"><span data-stu-id="581a9-207">Read/Write Text</span></span>||<span data-ttu-id="581a9-208">Y</span><span class="sxs-lookup"><span data-stu-id="581a9-208">Y</span></span>||<span data-ttu-id="581a9-209">Y</span><span class="sxs-lookup"><span data-stu-id="581a9-209">Y</span></span>|<span data-ttu-id="581a9-210">Y</span><span class="sxs-lookup"><span data-stu-id="581a9-210">Y</span></span>|<span data-ttu-id="581a9-211">是</span><span class="sxs-lookup"><span data-stu-id="581a9-211">Y</span></span><br/><span data-ttu-id="581a9-212">（只读）</span><span class="sxs-lookup"><span data-stu-id="581a9-212">(Read only)</span></span>|
||<span data-ttu-id="581a9-213">读/写矩阵</span><span class="sxs-lookup"><span data-stu-id="581a9-213">Read/Write Matrix</span></span>||<span data-ttu-id="581a9-214">Y</span><span class="sxs-lookup"><span data-stu-id="581a9-214">Y</span></span>|||<span data-ttu-id="581a9-215">Y</span><span class="sxs-lookup"><span data-stu-id="581a9-215">Y</span></span>||
||<span data-ttu-id="581a9-216">读/写表</span><span class="sxs-lookup"><span data-stu-id="581a9-216">Read/Write Table</span></span>||<span data-ttu-id="581a9-217">Y</span><span class="sxs-lookup"><span data-stu-id="581a9-217">Y</span></span>|||<span data-ttu-id="581a9-218">Y</span><span class="sxs-lookup"><span data-stu-id="581a9-218">Y</span></span>||
||<span data-ttu-id="581a9-219">读/写 HTML</span><span class="sxs-lookup"><span data-stu-id="581a9-219">Read/Write HTML</span></span>|||||<span data-ttu-id="581a9-220">是</span><span class="sxs-lookup"><span data-stu-id="581a9-220">Y</span></span>||
||<span data-ttu-id="581a9-221">读/写</span><span class="sxs-lookup"><span data-stu-id="581a9-221">Read/Write</span></span><br/><span data-ttu-id="581a9-222">Office Open XML</span><span class="sxs-lookup"><span data-stu-id="581a9-222">Office Open XML</span></span>|||||<span data-ttu-id="581a9-223">是</span><span class="sxs-lookup"><span data-stu-id="581a9-223">Y</span></span>||
||<span data-ttu-id="581a9-224">读取任务、资源、视图和字段属性</span><span class="sxs-lookup"><span data-stu-id="581a9-224">Read task, resource, view, and field properties</span></span>||||||<span data-ttu-id="581a9-225">Y</span><span class="sxs-lookup"><span data-stu-id="581a9-225">Y</span></span>|
||<span data-ttu-id="581a9-226">选择已更改事件</span><span class="sxs-lookup"><span data-stu-id="581a9-226">Selection changed events</span></span>||<span data-ttu-id="581a9-227">Y</span><span class="sxs-lookup"><span data-stu-id="581a9-227">Y</span></span>|||<span data-ttu-id="581a9-228">Y</span><span class="sxs-lookup"><span data-stu-id="581a9-228">Y</span></span>||
||<span data-ttu-id="581a9-229">获取整个文档</span><span class="sxs-lookup"><span data-stu-id="581a9-229">Get whole document</span></span>||||<span data-ttu-id="581a9-230">Y</span><span class="sxs-lookup"><span data-stu-id="581a9-230">Y</span></span>|<span data-ttu-id="581a9-231">是</span><span class="sxs-lookup"><span data-stu-id="581a9-231">Y</span></span>||
||<span data-ttu-id="581a9-232">绑定和绑定事件</span><span class="sxs-lookup"><span data-stu-id="581a9-232">Bindings and binding events</span></span>|<span data-ttu-id="581a9-233">是</span><span class="sxs-lookup"><span data-stu-id="581a9-233">Y</span></span><br/><span data-ttu-id="581a9-234">（仅限完全和部分表格绑定）</span><span class="sxs-lookup"><span data-stu-id="581a9-234">(Only full and partial table bindings)</span></span>|<span data-ttu-id="581a9-235">是</span><span class="sxs-lookup"><span data-stu-id="581a9-235">Y</span></span>|||<span data-ttu-id="581a9-236">是</span><span class="sxs-lookup"><span data-stu-id="581a9-236">Y</span></span>||
||<span data-ttu-id="581a9-237">读/写自定义 XML 部分</span><span class="sxs-lookup"><span data-stu-id="581a9-237">Read/Write Custom XML Parts</span></span>|||||<span data-ttu-id="581a9-238">是</span><span class="sxs-lookup"><span data-stu-id="581a9-238">Y</span></span>||
||<span data-ttu-id="581a9-239">暂留加载项状态数据（设置）</span><span class="sxs-lookup"><span data-stu-id="581a9-239">Persist add-in state data (settings)</span></span>|<span data-ttu-id="581a9-240">是</span><span class="sxs-lookup"><span data-stu-id="581a9-240">Y</span></span><br/><span data-ttu-id="581a9-241">（每主机加载项）</span><span class="sxs-lookup"><span data-stu-id="581a9-241">(Per host add-in)</span></span>|<span data-ttu-id="581a9-242">是</span><span class="sxs-lookup"><span data-stu-id="581a9-242">Y</span></span><br/><span data-ttu-id="581a9-243">（每文档）</span><span class="sxs-lookup"><span data-stu-id="581a9-243">(Per document)</span></span>|<span data-ttu-id="581a9-244">是</span><span class="sxs-lookup"><span data-stu-id="581a9-244">Y</span></span><br/><span data-ttu-id="581a9-245">（每邮箱）</span><span class="sxs-lookup"><span data-stu-id="581a9-245">(Per mailbox)</span></span>|<span data-ttu-id="581a9-246">是</span><span class="sxs-lookup"><span data-stu-id="581a9-246">Y</span></span><br/><span data-ttu-id="581a9-247">（每文档）</span><span class="sxs-lookup"><span data-stu-id="581a9-247">(Per document)</span></span>|<span data-ttu-id="581a9-248">是</span><span class="sxs-lookup"><span data-stu-id="581a9-248">Y</span></span><br/><span data-ttu-id="581a9-249">（每文档）</span><span class="sxs-lookup"><span data-stu-id="581a9-249">(Per document)</span></span>||
||<span data-ttu-id="581a9-250">设置更改事件</span><span class="sxs-lookup"><span data-stu-id="581a9-250">Settings changed events</span></span>|<span data-ttu-id="581a9-251">Y</span><span class="sxs-lookup"><span data-stu-id="581a9-251">Y</span></span>|<span data-ttu-id="581a9-252">Y</span><span class="sxs-lookup"><span data-stu-id="581a9-252">Y</span></span>||<span data-ttu-id="581a9-253">Y</span><span class="sxs-lookup"><span data-stu-id="581a9-253">Y</span></span>|<span data-ttu-id="581a9-254">是</span><span class="sxs-lookup"><span data-stu-id="581a9-254">Y</span></span>||
||<span data-ttu-id="581a9-255">获取活动视图模式</span><span class="sxs-lookup"><span data-stu-id="581a9-255">Get active view mode</span></span><br/><span data-ttu-id="581a9-256">和视图更改事件</span><span class="sxs-lookup"><span data-stu-id="581a9-256">and view changed events</span></span>||||<span data-ttu-id="581a9-257">是</span><span class="sxs-lookup"><span data-stu-id="581a9-257">Y</span></span>|||
||<span data-ttu-id="581a9-258">转到文档中</span><span class="sxs-lookup"><span data-stu-id="581a9-258">Navigate to locations</span></span><br/><span data-ttu-id="581a9-259">的相应位置</span><span class="sxs-lookup"><span data-stu-id="581a9-259">in the document</span></span>||<span data-ttu-id="581a9-260">是</span><span class="sxs-lookup"><span data-stu-id="581a9-260">Y</span></span>||<span data-ttu-id="581a9-261">Y</span><span class="sxs-lookup"><span data-stu-id="581a9-261">Y</span></span>|<span data-ttu-id="581a9-262">是</span><span class="sxs-lookup"><span data-stu-id="581a9-262">Y</span></span>||
||<span data-ttu-id="581a9-263">使用规则和 RegEx </span><span class="sxs-lookup"><span data-stu-id="581a9-263">Activate contextually</span></span><br/><span data-ttu-id="581a9-264">执行上下文式激活</span><span class="sxs-lookup"><span data-stu-id="581a9-264">using rules and RegEx</span></span>|||<span data-ttu-id="581a9-265">是</span><span class="sxs-lookup"><span data-stu-id="581a9-265">Y</span></span>||||
||<span data-ttu-id="581a9-266">读取项目属性</span><span class="sxs-lookup"><span data-stu-id="581a9-266">Read Item properties</span></span>|||<span data-ttu-id="581a9-267">Y</span><span class="sxs-lookup"><span data-stu-id="581a9-267">Y</span></span>||||
||<span data-ttu-id="581a9-268">读取用户配置文件</span><span class="sxs-lookup"><span data-stu-id="581a9-268">Read User profile</span></span>|||<span data-ttu-id="581a9-269">Y</span><span class="sxs-lookup"><span data-stu-id="581a9-269">Y</span></span>||||
||<span data-ttu-id="581a9-270">获取附件</span><span class="sxs-lookup"><span data-stu-id="581a9-270">Get attachments</span></span>|||<span data-ttu-id="581a9-271">Y</span><span class="sxs-lookup"><span data-stu-id="581a9-271">Y</span></span>||||
||<span data-ttu-id="581a9-272">获取用户标识令牌</span><span class="sxs-lookup"><span data-stu-id="581a9-272">Get User identity token</span></span>|||<span data-ttu-id="581a9-273">Y</span><span class="sxs-lookup"><span data-stu-id="581a9-273">Y</span></span>||||
||<span data-ttu-id="581a9-274">调用 Exchange Web 服务</span><span class="sxs-lookup"><span data-stu-id="581a9-274">Call Exchange Web Services</span></span>|||<span data-ttu-id="581a9-275">Y</span><span class="sxs-lookup"><span data-stu-id="581a9-275">Y</span></span>||||
