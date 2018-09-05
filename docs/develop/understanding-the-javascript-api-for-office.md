---
title: 了解适用于 Office 的 JavaScript API
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 12e7d9030ec37746f84e3fc725cddda2a5675761
ms.sourcegitcommit: 5bef9828f047da03ecf2f43c6eb5b8514eff28ce
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/31/2018
ms.locfileid: "23782792"
---
# <a name="understanding-the-javascript-api-for-office"></a><span data-ttu-id="e8967-102">了解适用于 Office 的 JavaScript API</span><span class="sxs-lookup"><span data-stu-id="e8967-102">Understanding the JavaScript API for Office</span></span>

<span data-ttu-id="e8967-p101">本文提供了有关适用于 Office 的 JavaScript API 的信息以及使用方法。有关参考信息，请参阅 [适用于 Office 的 JavaScript API](https://dev.office.com/reference/add-ins/javascript-api-for-office)。有关将 Visual Studio 项目文件更新到适用于 Office 的 JavaScript API 的最新当前版本的信息，请参阅 [更新适用于 Office 的 JavaScript API 版本和清单架构文件](update-your-javascript-api-for-office-and-manifest-schema-version.md)。</span><span class="sxs-lookup"><span data-stu-id="e8967-p101">This article provides information about the JavaScript API for Office and how to use it. For reference information, see [JavaScript API for Office](https://dev.office.com/reference/add-ins/javascript-api-for-office). For information about updating Visual Studio project files to the most current version of the JavaScript API for Office, see [Update the version of your JavaScript API for Office and manifest schema files](update-your-javascript-api-for-office-and-manifest-schema-version.md).</span></span>

> [!NOTE]
> <span data-ttu-id="e8967-p102">如果计划将加载项[发布](../publish/publish.md)到 AppSource 并适用于 Office 体验，请务必遵循 [AppSource 验证策略](https://docs.microsoft.com/office/dev/store/validation-policies)。例如，加载项必须适用于支持已定义方法的所有平台，才能通过验证（有关详细信息，请参阅[第 4.12 部分](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably)以及 [Office 加载项主机和可用性](../overview/office-add-in-availability.md)页面）。</span><span class="sxs-lookup"><span data-stu-id="e8967-p102">If you plan to [publish](../publish/publish.md) your add-in to AppSource and make it available within the Office experience, make sure that you conform to the [AppSource validation policies](https://docs.microsoft.com/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)).</span></span> 

## <a name="referencing-the-javascript-api-for-office-library-in-your-add-in"></a><span data-ttu-id="e8967-108">在加载项中引用适用于 Office 的 JavaScript API 库</span><span class="sxs-lookup"><span data-stu-id="e8967-108">Referencing the JavaScript API for Office library in your add-in</span></span>

<span data-ttu-id="e8967-p103">[适用于 Office 的 JavaScript](https://dev.office.com/reference/add-ins/javascript-api-for-office) 库包含 Office.js 文件和关联的特定于主机应用程序的 .js 文件，例如 Excel-15.js 和 Outlook-15.js。引用该 API 最简单的方法是通过添加以下 `<script>` 到你的页面的 `<head>` 标记来使用我们的 CDN：</span><span class="sxs-lookup"><span data-stu-id="e8967-p103">The [JavaScript API for Office](https://dev.office.com/reference/add-ins/javascript-api-for-office) library consists of the Office.js file and associated host application-specific .js files, such as Excel-15.js and Outlook-15.js. The simplest method of referencing the API is using our CDN by adding the following `<script>` to your page's `<head>` tag:</span></span>  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

<span data-ttu-id="e8967-111">这将在加载项首次加载时下载并缓存适用于 Office 的 JavaScript API 文件，以确保对特定版本使用 Office.js 及其关联文件的最新实现。</span><span class="sxs-lookup"><span data-stu-id="e8967-111">This will download and cache the JavaScript API for Office files the first time your add-in loads to make sure that it is using the most up-to-date implementation of Office.js and its associated files for the specified version.</span></span>

<span data-ttu-id="e8967-112">有关 Office.js CDN 的更多详细信息（包括如何处理版本控制和向后兼容性），请参阅[从内容分发网络 (CDN) 引用适用于 Office 的 JavaScript API 库](referencing-the-javascript-api-for-office-library-from-its-cdn.md)。</span><span class="sxs-lookup"><span data-stu-id="e8967-112">For more details around the Office.js CDN, including how versioning and backward compatability is handled, see [Referencing the JavaScript API for Office library from its content delivery network (CDN)](referencing-the-javascript-api-for-office-library-from-its-cdn.md).</span></span>

## <a name="initializing-your-add-in"></a><span data-ttu-id="e8967-113">初始化加载项</span><span class="sxs-lookup"><span data-stu-id="e8967-113">Initializing your add-in</span></span>

<span data-ttu-id="e8967-114">**适用于：** 所有加载项类型</span><span class="sxs-lookup"><span data-stu-id="e8967-114">**Applies to:** All add-in types</span></span>

<span data-ttu-id="e8967-115">Office 加载项通常有启动逻辑，以执行以下事项：</span><span class="sxs-lookup"><span data-stu-id="e8967-115">Office add-ins often have start-up logic to do things such as:</span></span>

- <span data-ttu-id="e8967-116">检查用户的 Office 版本是否支持您的代码调用的所有 Office Api。</span><span class="sxs-lookup"><span data-stu-id="e8967-116">Check that the user's version of Office will support all the Office APIs that your code calls.</span></span>

- <span data-ttu-id="e8967-117">确保存在某些工件，如具有特定名称的工作表。</span><span class="sxs-lookup"><span data-stu-id="e8967-117">Ensure the existence of certain artifacts, such as worksheet with a specific name.</span></span>

- <span data-ttu-id="e8967-118">提示用户选择 Excel 中的一些单元格，然后插入使用这些选定值初始化的图表。</span><span class="sxs-lookup"><span data-stu-id="e8967-118">You can use the initialize event handler to implement common add-in initialization scenarios, such as prompting the user to select some cells in Excel, and then inserting a chart initialized with those selected values.</span></span>

- <span data-ttu-id="e8967-119">建立绑定。</span><span class="sxs-lookup"><span data-stu-id="e8967-119">Establish bindings.</span></span>

- <span data-ttu-id="e8967-120">使用 Office 对话框 API 提示用户输入默认加载项设置值。</span><span class="sxs-lookup"><span data-stu-id="e8967-120">Use the Office dialog API to prompt the user for default add-in settings values.</span></span>

<span data-ttu-id="e8967-121">但是，在完全加载完库之前，您启动代码不得调用任何 Office.js Api。</span><span class="sxs-lookup"><span data-stu-id="e8967-121">But your start-up code must not call any Office.js APIs until the library is fully loaded.</span></span> <span data-ttu-id="e8967-122">有两种方法让您的代码可以确保加载库。</span><span class="sxs-lookup"><span data-stu-id="e8967-122">There are two ways that your code can ensure that the library is loaded.</span></span> <span data-ttu-id="e8967-123">这将在以下各节介绍。</span><span class="sxs-lookup"><span data-stu-id="e8967-123">They are described in the sections below.</span></span> <span data-ttu-id="e8967-124">我们建议您使用名为 `Office.onReady()` 的较新、 更灵活的技术。</span><span class="sxs-lookup"><span data-stu-id="e8967-124">We recommend that you use the newer, more flexible, technique, calling `Office.onReady()`.</span></span> <span data-ttu-id="e8967-125">仍然支持分配处理程序 `Office.initialize` 的旧技术。</span><span class="sxs-lookup"><span data-stu-id="e8967-125">The older technique, assigning a handler to `Office.initialize`, is still supported.</span></span> <span data-ttu-id="e8967-126">请参阅 [Office.initialize 和 Office.onReady() 的主要区别](#major-differences-between-office-initialize-and-office-onready)。</span><span class="sxs-lookup"><span data-stu-id="e8967-126">See also [Major differences between Office.initialize and Office.onReady()](#major-differences-between-office-initialize-and-office-onready).</span></span>

<span data-ttu-id="e8967-127">若要详细了解加载项初始化时的事件发生顺序，请参阅[加载 DOM 和运行时环境](loading-the-dom-and-runtime-environment.md)。</span><span class="sxs-lookup"><span data-stu-id="e8967-127">For more detail about the sequence of events when an add-in is initialized, see [Loading the DOM and runtime environment](loading-the-dom-and-runtime-environment.md).</span></span>

### <a name="initialize-with-officeonready"></a><span data-ttu-id="e8967-128">使用 Office.onReady() 初始化</span><span class="sxs-lookup"><span data-stu-id="e8967-128">Initialize with Office.onReady()</span></span>

<span data-ttu-id="e8967-129">`Office.onReady()` 是返回承诺对象，同时检查 Office.js 库是否完全加载的异步方法。</span><span class="sxs-lookup"><span data-stu-id="e8967-129">`Office.onReady()` is an asynchronous method that returns a Promise object while it checks to see if the Office.js library is fully loaded.</span></span> <span data-ttu-id="e8967-130">只有在加载库后，它才会将承诺解析为对象，这将使用`Office.HostType` 枚举值 (`Excel`， `Word`等) 和与平台 `Office.PlatformType` 枚举值 (`PC`， `Mac`， `OfficeOnline`，等）指定 Office 主机应用程序。</span><span class="sxs-lookup"><span data-stu-id="e8967-130">When, and only when, the library is loaded, it resolves the Promise as an object that specifies the Office host application with an `Office.HostType` enum value (`Excel`, `Word`, etc.) and the platform with an `Office.PlatformType` enum value (`PC`, `Mac`, `OfficeOnline`, etc.).</span></span> <span data-ttu-id="e8967-131">如果在调用 `Office.onReady()` 时已加载库，则承诺立即解析。</span><span class="sxs-lookup"><span data-stu-id="e8967-131">If the library is already loaded when `Office.onReady()` is called, then the Promise resolves immediately.</span></span>

<span data-ttu-id="e8967-132">调用的一种方法 `Office.onReady()` 是，将其传递给回调方法。</span><span class="sxs-lookup"><span data-stu-id="e8967-132">One way to call `Office.onReady()` is to pass it a callback method.</span></span> <span data-ttu-id="e8967-133">下面是一个示例：</span><span class="sxs-lookup"><span data-stu-id="e8967-133">Here's an example:</span></span>

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

<span data-ttu-id="e8967-134">或者，您可以将 `then()` 方法与 `Office.onReady()` 的调用链接而不是传递回调。</span><span class="sxs-lookup"><span data-stu-id="e8967-134">Alternatively, you can chain a `then()` method to the call of `Office.onReady()`, instead of passing a callback.</span></span> <span data-ttu-id="e8967-135">例如，下面的代码将检查用户的 Excel 版本是否支持加载项可能调用的所有 Api。</span><span class="sxs-lookup"><span data-stu-id="e8967-135">For example, the following code checks to see that the user's version of Excel supports all the APIs that the add-in might call.</span></span>

```js
Office.onReady()
    .then(function() {
        if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
            console.log("Sorry, this add-in only works with newer versions of Excel.");
        }
    });
```

<span data-ttu-id="e8967-136">以下是在 TypeScript 中使用 `async` 和 `await` 关键字的相同示例：</span><span class="sxs-lookup"><span data-stu-id="e8967-136">Here is the same example using the `async` and `await` keywords in TypeScript:</span></span>

```typescript
(async () => {
    await Office.onReady();
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
})();
```

<span data-ttu-id="e8967-137">如果您正使用包括自己的初始化处理程序或测试的其他 JavaScript 框架，则这些*通常应*放在 `Office.onReady()` 的响应内。</span><span class="sxs-lookup"><span data-stu-id="e8967-137">If you are using additional JavaScript frameworks that include their own initialization handler or tests, these should be placed within the Office.initialize event.</span></span> <span data-ttu-id="e8967-138">例如，会对 [JQuery](https://jquery.com) `$(document).ready()` 函数进行以下引用：</span><span class="sxs-lookup"><span data-stu-id="e8967-138">For example, [JQuery's](https://jquery.com) `$(document).ready()` function would be referenced as follows:</span></span>

```js
Office.onReady(function() {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
});
```

<span data-ttu-id="e8967-139">但是，这一做法存在一些例外。</span><span class="sxs-lookup"><span data-stu-id="e8967-139">However, there are exceptions to this practice.</span></span> <span data-ttu-id="e8967-140">例如，假设您想要在浏览器中打开您的加载项（而不是 侧加载到 Office 主机），从而使用浏览器工具调试您的 UI。</span><span class="sxs-lookup"><span data-stu-id="e8967-140">For example, suppose you want to open your add-in in a browser (instead of sideload it in an Office host) in order to debug your UI with browser tools.</span></span> <span data-ttu-id="e8967-141">由于 Office.js 无法在浏览器中加载，`onReady` 将无法运行，同时如果在 Office `onReady` 内调用，`$(document).ready` 将无法运行。</span><span class="sxs-lookup"><span data-stu-id="e8967-141">Since Office.js won't load in the browser, `onReady` won't run and the `$(document).ready` won't run if it's called inside the Office `onReady`.</span></span> <span data-ttu-id="e8967-142">另一个异常：加载加载项期间，您希望在任务窗格中显示进度指示器。</span><span class="sxs-lookup"><span data-stu-id="e8967-142">Another exception: you want a progress indicator to appear in the task pane while the add-in is loading.</span></span> <span data-ttu-id="e8967-143">在此方案中，您的代码应调用 jQuery `ready`，并使用它的回调以呈现进度指示器。</span><span class="sxs-lookup"><span data-stu-id="e8967-143">In this scenario, your code should call the jQuery `ready` and use it's callback to render the progress indicator.</span></span> <span data-ttu-id="e8967-144">然后，Office `onReady`的回调可以替换最终用户界面的进度指示器。</span><span class="sxs-lookup"><span data-stu-id="e8967-144">Then the Office `onReady`'s callback can replace the progress indicator with the final UI.</span></span> 

### <a name="initialize-with-officeinitialize"></a><span data-ttu-id="e8967-145">使用 Office.initialize 初始化</span><span class="sxs-lookup"><span data-stu-id="e8967-145">Initialize with Office.initialize</span></span>

<span data-ttu-id="e8967-146">当 Office.js 库完全加载并可供用户交互时，初始化事件触发。</span><span class="sxs-lookup"><span data-stu-id="e8967-146">An initialize event fires when the Office.js library is fully loaded and ready for user interaction.</span></span> <span data-ttu-id="e8967-147">您可以分配一个处理程序给实施初始化逻辑的 `Office.initialize`。</span><span class="sxs-lookup"><span data-stu-id="e8967-147">You can assign a handler to `Office.initialize` that implements your initialization logic.</span></span> <span data-ttu-id="e8967-148">以下是检查用户的 Excel 版本是否支持所有可能调用加载项的 Api 示例。</span><span class="sxs-lookup"><span data-stu-id="e8967-148">The following is an example that checks to see that the user's version of Excel supports all the APIs that the add-in might call.</span></span>

```js
Office.initialize = function () {
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
};
```

<span data-ttu-id="e8967-149">如果您正使用包括自己的初始化处理程序或测试的其他 JavaScript 框架，则这应*通常*放在 `Office.initialize` 事件内。</span><span class="sxs-lookup"><span data-stu-id="e8967-149">If you are using additional JavaScript frameworks that include their own initialization handler or tests, these should be placed within the Office.initialize event.</span></span> <span data-ttu-id="e8967-150">（但是，更早版本 **与 Office.onReady() 初始化** 一节描述的异常也适用于这种情况。）例如， [JQuery](https://jquery.com) `$(document).ready()`函数会被引用为：</span><span class="sxs-lookup"><span data-stu-id="e8967-150">(But the exceptions described in the **Initialize with Office.onReady()** section earlier apply in this case also.) For example, [JQuery's](https://jquery.com) `$(document).ready()` function would be referenced as follows:</span></span>

```js
Office.initialize = function () {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
  };
```

<span data-ttu-id="e8967-151">对于任务窗格和内容加载项，`Office.initialize` 提供其他_原因_参数。</span><span class="sxs-lookup"><span data-stu-id="e8967-151">For task pane and content add-ins, `Office.initialize` provides an additional _reason_ parameter.</span></span> <span data-ttu-id="e8967-152">此参数指定如何添加加载项到当前文档。</span><span class="sxs-lookup"><span data-stu-id="e8967-152">This parameter can be used to determine how an add-in was added to the current document.</span></span> <span data-ttu-id="e8967-153">您可以使用此参数提供首次插入加载项时和加载项已存在于文档中时的不同逻辑。</span><span class="sxs-lookup"><span data-stu-id="e8967-153">You can use this to provide different logic for when an add-in is first inserted versus when it already existed within the document.</span></span>

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

<span data-ttu-id="e8967-154">有关详细信息，请参阅 [Office.initialize 事件](https://dev.office.com/reference/add-ins/shared/office.initialize)和 [InitializationReason 枚举](https://dev.office.com/reference/add-ins/shared/initializationreason-enumeration)。</span><span class="sxs-lookup"><span data-stu-id="e8967-154">For more information, see [Office.initialize Event](https://dev.office.com/reference/add-ins/shared/office.initialize) and [InitializationReason Enumeration](https://dev.office.com/reference/add-ins/shared/initializationreason-enumeration).</span></span>

### <a name="major-differences-between-officeinitialize-and-officeonready"></a><span data-ttu-id="e8967-155">Office.initialize 和 Office.onReady 的主要区别</span><span class="sxs-lookup"><span data-stu-id="e8967-155">Major differences between Office.initialize and Office.onReady</span></span>

- <span data-ttu-id="e8967-156">您仅可分配一个处理程序到 `Office.initialize`，同时它由由 Office 基础架构仅调用一次；但是，您可以在代码中的不同位置调用 `Office.onReady()` 并可使用不同的回调。</span><span class="sxs-lookup"><span data-stu-id="e8967-156">You can assign only one handler to `Office.initialize` and it is called, only once, by the Office infrastructure; but you can call `Office.onReady()` in different places in your code and use different callbacks.</span></span> <span data-ttu-id="e8967-157">例如，一旦使用运行初始化逻辑的调用加载您的自定义脚本，您的代码即可调用 `Office.onReady()`；同时，您的代码还可在任务窗格中有一个按钮，其脚本以不同的回调来调用 `Office.onReady()`。</span><span class="sxs-lookup"><span data-stu-id="e8967-157">For example, your code could call `Office.onReady()` as soon as your custom script loads with a callback that runs initialization logic; and your code could also have a button in the task pane, whose script calls `Office.onReady()` with a different callback.</span></span> <span data-ttu-id="e8967-158">如果是这样，单击按钮时将运行第二个回调。</span><span class="sxs-lookup"><span data-stu-id="e8967-158">If so, the second callback runs when the button is clicked.</span></span>

- <span data-ttu-id="e8967-159"> `Office.initialize` 事件在 Office.js 初始化本身的内部过程末尾触发。</span><span class="sxs-lookup"><span data-stu-id="e8967-159">The `Office.initialize` event fires at the end of the internal process in which Office.js initializes itself.</span></span> <span data-ttu-id="e8967-160">这在内部过程结束后*立即*触发。</span><span class="sxs-lookup"><span data-stu-id="e8967-160">And it fires *immediately* after the internal process ends.</span></span> <span data-ttu-id="e8967-161">如果事件触发后指定处理程序给事件的代码执行时间过长，则不运行您的处理程序。</span><span class="sxs-lookup"><span data-stu-id="e8967-161">If the code in which you assign a handler to the event executes too long after the event fires, then your handler doesn't run.</span></span> <span data-ttu-id="e8967-162">例如，如果您使用 WebPack 任务管理器，它可能在加载 Office.js 后，但在加载您的自定义 JavaScript 之前配置加载项主页以加载 polyfill 文件。</span><span class="sxs-lookup"><span data-stu-id="e8967-162">For example, if you are using the WebPack task manager, it might configure the add-in's home page to load polyfill files after it loads Office.js but before it loads your custom JavaScript.</span></span> <span data-ttu-id="e8967-163">脚本加载并分配该处理程序时，初始化事件已经发生。</span><span class="sxs-lookup"><span data-stu-id="e8967-163">By the time your script loads and assigns the handler, the initialize event has already happened.</span></span> <span data-ttu-id="e8967-164">但是，调用 `Office.onReady()` 不会"过晚"。</span><span class="sxs-lookup"><span data-stu-id="e8967-164">But it is never "too late" to call `Office.onReady()`.</span></span> <span data-ttu-id="e8967-165">如果初始化事件已经发生，则回调立即运行。</span><span class="sxs-lookup"><span data-stu-id="e8967-165">If the initialize event has already happened, the callback runs immediately.</span></span>

> [!NOTE]
> <span data-ttu-id="e8967-166">即使未启动逻辑，当加载加载项 JavaScript 时，调用 `Office.onReady()` 或分配到一个空函数给 `Office.initialize` 是一个好的做法，因为在发生下列任一情况之前，某些 Office 主机和平台组合不会加载任务窗格。</span><span class="sxs-lookup"><span data-stu-id="e8967-166">Even if you have no start-up logic, it is a good practice to either call `Office.onReady()` or assign an empty function to `Office.initialize` when your add-in JavaScript loads, because some Office host and platform combinations won't load the task pane until one of these happens.</span></span> <span data-ttu-id="e8967-167">以下各行显示可以完成这个的两种方式：</span><span class="sxs-lookup"><span data-stu-id="e8967-167">The following lines show the two ways this can be done:</span></span>
>
>```js
>Office.onReady();
>```
>
>```js
>Office.initialize = function () {};
>```

## <a name="office-javascript-api-object-model"></a><span data-ttu-id="e8967-168">Office JavaScript API 对象模型</span><span class="sxs-lookup"><span data-stu-id="e8967-168">Office JavaScript API object model</span></span>

<span data-ttu-id="e8967-169">初始化后，加载项可以与主机（例如 Excel、Outlook）交互。</span><span class="sxs-lookup"><span data-stu-id="e8967-169">Once initialized, the add-in can interact with the host (e.g. Excel, Outlook).</span></span> <span data-ttu-id="e8967-170">[Office JavaScript API 对象模型](office-javascript-api-object-model.md)页面有关于特定使用模式的更多详细信息。</span><span class="sxs-lookup"><span data-stu-id="e8967-170">The [Office JavaScript API object model](office-javascript-api-object-model.md)) page has more details on specific usage patterns.</span></span> <span data-ttu-id="e8967-171">[共享 API](https://dev.office.com/reference/add-ins/javascript-api-for-office) 和特定主机都有详细的参考文档。</span><span class="sxs-lookup"><span data-stu-id="e8967-171">There is also detailed reference documentation for both [shared APIs](https://dev.office.com/reference/add-ins/javascript-api-for-office) and specific hosts.</span></span>

## <a name="api-support-matrix"></a><span data-ttu-id="e8967-172">API 支持矩阵</span><span class="sxs-lookup"><span data-stu-id="e8967-172">API support matrix</span></span>

<span data-ttu-id="e8967-173">下表总结了各种类型的加载项（内容、任务窗格和 Outlook）支持的 API 和功能，以及使用[适用于 Office 的 JavaScript API v1.1 支持的 1.1 加载项清单架构和功能](update-your-javascript-api-for-office-and-manifest-schema-version.md)指定加载项支持的 Office 主机应用时，可以托管它们的 Office 应用。</span><span class="sxs-lookup"><span data-stu-id="e8967-173">This table summarizes the API and features supported across add-in types (content, task pane, and Outlook) and the Office applications that can host them when you specify the Office host applications your add-in supports by using the [1.1 add-in manifest schema and features supported by v1.1 JavaScript API for Office](update-your-javascript-api-for-office-and-manifest-schema-version.md).</span></span>


|||||||||
|:-----|:-----|:-----:|:-----:|:-----:|:-----:|:-----:|:-----:|
||<span data-ttu-id="e8967-174">**主机名**</span><span class="sxs-lookup"><span data-stu-id="e8967-174">**Host name**</span></span>|<span data-ttu-id="e8967-175">数据库</span><span class="sxs-lookup"><span data-stu-id="e8967-175">Database</span></span>|<span data-ttu-id="e8967-176">工作簿</span><span class="sxs-lookup"><span data-stu-id="e8967-176">Workbook</span></span>|<span data-ttu-id="e8967-177">邮箱</span><span class="sxs-lookup"><span data-stu-id="e8967-177">Mailbox</span></span>|<span data-ttu-id="e8967-178">演示文稿</span><span class="sxs-lookup"><span data-stu-id="e8967-178">Presentation</span></span>|<span data-ttu-id="e8967-179">文档</span><span class="sxs-lookup"><span data-stu-id="e8967-179">Document</span></span>|<span data-ttu-id="e8967-180">项目</span><span class="sxs-lookup"><span data-stu-id="e8967-180">Project</span></span>|
||<span data-ttu-id="e8967-181">**支持的****主机应用程序**</span><span class="sxs-lookup"><span data-stu-id="e8967-181">**Supported** **Host applications**</span></span>|<span data-ttu-id="e8967-182">Access Web App</span><span class="sxs-lookup"><span data-stu-id="e8967-182">Access web apps</span></span>|<span data-ttu-id="e8967-183">Excel、</span><span class="sxs-lookup"><span data-stu-id="e8967-183">Excel,</span></span><br/><span data-ttu-id="e8967-184">Excel 在线</span><span class="sxs-lookup"><span data-stu-id="e8967-184">Excel Online</span></span>|<span data-ttu-id="e8967-185">Outlook、</span><span class="sxs-lookup"><span data-stu-id="e8967-185">Outlook,</span></span><br/><span data-ttu-id="e8967-186">Outlook Web App、</span><span class="sxs-lookup"><span data-stu-id="e8967-186">Outlook Web App,</span></span><br/><span data-ttu-id="e8967-187">适用于设备的 OWA</span><span class="sxs-lookup"><span data-stu-id="e8967-187">OWA for Devices</span></span>|<span data-ttu-id="e8967-188">PowerPoint、</span><span class="sxs-lookup"><span data-stu-id="e8967-188">PowerPoint,</span></span><br/><span data-ttu-id="e8967-189">PowerPoint 联机</span><span class="sxs-lookup"><span data-stu-id="e8967-189">PowerPoint Online</span></span>|<span data-ttu-id="e8967-190">Word</span><span class="sxs-lookup"><span data-stu-id="e8967-190">Word</span></span>|<span data-ttu-id="e8967-191">项目</span><span class="sxs-lookup"><span data-stu-id="e8967-191">Project</span></span>|
|<span data-ttu-id="e8967-192">**支持的外接程序类型**</span><span class="sxs-lookup"><span data-stu-id="e8967-192">**Supported add-in types**</span></span>|<span data-ttu-id="e8967-193">内容</span><span class="sxs-lookup"><span data-stu-id="e8967-193">Content</span></span>|<span data-ttu-id="e8967-194">是</span><span class="sxs-lookup"><span data-stu-id="e8967-194">Y</span></span>|<span data-ttu-id="e8967-195">是</span><span class="sxs-lookup"><span data-stu-id="e8967-195">Y</span></span>||<span data-ttu-id="e8967-196">是</span><span class="sxs-lookup"><span data-stu-id="e8967-196">Y</span></span>|||
||<span data-ttu-id="e8967-197">任务窗格</span><span class="sxs-lookup"><span data-stu-id="e8967-197">Task pane</span></span>||<span data-ttu-id="e8967-198">是</span><span class="sxs-lookup"><span data-stu-id="e8967-198">Y</span></span>||<span data-ttu-id="e8967-199">是</span><span class="sxs-lookup"><span data-stu-id="e8967-199">Y</span></span>|<span data-ttu-id="e8967-200">是</span><span class="sxs-lookup"><span data-stu-id="e8967-200">Y</span></span>|<span data-ttu-id="e8967-201">是</span><span class="sxs-lookup"><span data-stu-id="e8967-201">Y</span></span>|
||<span data-ttu-id="e8967-202">Outlook</span><span class="sxs-lookup"><span data-stu-id="e8967-202">Outlook</span></span>|||<span data-ttu-id="e8967-203">是</span><span class="sxs-lookup"><span data-stu-id="e8967-203">Y</span></span>||||
|<span data-ttu-id="e8967-204">**支持的 API 功能**</span><span class="sxs-lookup"><span data-stu-id="e8967-204">**Supported API features**</span></span>|<span data-ttu-id="e8967-205">读/写文本</span><span class="sxs-lookup"><span data-stu-id="e8967-205">Read/Write Text</span></span>||<span data-ttu-id="e8967-206">是</span><span class="sxs-lookup"><span data-stu-id="e8967-206">Y</span></span>||<span data-ttu-id="e8967-207">是</span><span class="sxs-lookup"><span data-stu-id="e8967-207">Y</span></span>|<span data-ttu-id="e8967-208">是</span><span class="sxs-lookup"><span data-stu-id="e8967-208">Y</span></span>|<span data-ttu-id="e8967-209">是</span><span class="sxs-lookup"><span data-stu-id="e8967-209">Y</span></span><br/><span data-ttu-id="e8967-210">（只读）</span><span class="sxs-lookup"><span data-stu-id="e8967-210">(Read only)</span></span>|
||<span data-ttu-id="e8967-211">读/写矩阵</span><span class="sxs-lookup"><span data-stu-id="e8967-211">Read/Write Matrix</span></span>||<span data-ttu-id="e8967-212">是</span><span class="sxs-lookup"><span data-stu-id="e8967-212">Y</span></span>|||<span data-ttu-id="e8967-213">是</span><span class="sxs-lookup"><span data-stu-id="e8967-213">Y</span></span>||
||<span data-ttu-id="e8967-214">读/写表</span><span class="sxs-lookup"><span data-stu-id="e8967-214">Read/Write Table</span></span>||<span data-ttu-id="e8967-215">是</span><span class="sxs-lookup"><span data-stu-id="e8967-215">Y</span></span>|||<span data-ttu-id="e8967-216">是</span><span class="sxs-lookup"><span data-stu-id="e8967-216">Y</span></span>||
||<span data-ttu-id="e8967-217">读/写 HTML</span><span class="sxs-lookup"><span data-stu-id="e8967-217">Read/Write HTML</span></span>|||||<span data-ttu-id="e8967-218">是</span><span class="sxs-lookup"><span data-stu-id="e8967-218">Y</span></span>||
||<span data-ttu-id="e8967-219">读/写</span><span class="sxs-lookup"><span data-stu-id="e8967-219">Read/Write</span></span><br/><span data-ttu-id="e8967-220">Office Open XML</span><span class="sxs-lookup"><span data-stu-id="e8967-220">Office Open XML</span></span>|||||<span data-ttu-id="e8967-221">是</span><span class="sxs-lookup"><span data-stu-id="e8967-221">Y</span></span>||
||<span data-ttu-id="e8967-222">读取任务、资源、视图和字段属性</span><span class="sxs-lookup"><span data-stu-id="e8967-222">Read task, resource, view, and field properties</span></span>||||||<span data-ttu-id="e8967-223">是</span><span class="sxs-lookup"><span data-stu-id="e8967-223">Y</span></span>|
||<span data-ttu-id="e8967-224">选择已更改事件</span><span class="sxs-lookup"><span data-stu-id="e8967-224">Selection changed events</span></span>||<span data-ttu-id="e8967-225">是</span><span class="sxs-lookup"><span data-stu-id="e8967-225">Y</span></span>|||<span data-ttu-id="e8967-226">是</span><span class="sxs-lookup"><span data-stu-id="e8967-226">Y</span></span>||
||<span data-ttu-id="e8967-227">获取整个文档</span><span class="sxs-lookup"><span data-stu-id="e8967-227">Get whole document</span></span>||||<span data-ttu-id="e8967-228">是</span><span class="sxs-lookup"><span data-stu-id="e8967-228">Y</span></span>|<span data-ttu-id="e8967-229">是</span><span class="sxs-lookup"><span data-stu-id="e8967-229">Y</span></span>||
||<span data-ttu-id="e8967-230">绑定和绑定事件</span><span class="sxs-lookup"><span data-stu-id="e8967-230">Bindings and binding events</span></span>|<span data-ttu-id="e8967-231">是</span><span class="sxs-lookup"><span data-stu-id="e8967-231">Y</span></span><br/><span data-ttu-id="e8967-232">（仅限完全和部分表格绑定）</span><span class="sxs-lookup"><span data-stu-id="e8967-232">(Only full and partial table bindings)</span></span>|<span data-ttu-id="e8967-233">是</span><span class="sxs-lookup"><span data-stu-id="e8967-233">Y</span></span>|||<span data-ttu-id="e8967-234">是</span><span class="sxs-lookup"><span data-stu-id="e8967-234">Y</span></span>||
||<span data-ttu-id="e8967-235">读/写自定义 XML 部分</span><span class="sxs-lookup"><span data-stu-id="e8967-235">Read/Write Custom XML Parts</span></span>|||||<span data-ttu-id="e8967-236">是</span><span class="sxs-lookup"><span data-stu-id="e8967-236">Y</span></span>||
||<span data-ttu-id="e8967-237">暂留加载项状态数据（设置）</span><span class="sxs-lookup"><span data-stu-id="e8967-237">Persist add-in state data (settings)</span></span>|<span data-ttu-id="e8967-238">是</span><span class="sxs-lookup"><span data-stu-id="e8967-238">Y</span></span><br/><span data-ttu-id="e8967-239">（每主机加载项）</span><span class="sxs-lookup"><span data-stu-id="e8967-239">(Per host add-in)</span></span>|<span data-ttu-id="e8967-240">是</span><span class="sxs-lookup"><span data-stu-id="e8967-240">Y</span></span><br/><span data-ttu-id="e8967-241">（每文档）</span><span class="sxs-lookup"><span data-stu-id="e8967-241">(Per document)</span></span>|<span data-ttu-id="e8967-242">是</span><span class="sxs-lookup"><span data-stu-id="e8967-242">Y</span></span><br/><span data-ttu-id="e8967-243">（每邮箱）</span><span class="sxs-lookup"><span data-stu-id="e8967-243">(Per mailbox)</span></span>|<span data-ttu-id="e8967-244">是</span><span class="sxs-lookup"><span data-stu-id="e8967-244">Y</span></span><br/><span data-ttu-id="e8967-245">（每文档）</span><span class="sxs-lookup"><span data-stu-id="e8967-245">(Per document)</span></span>|<span data-ttu-id="e8967-246">是</span><span class="sxs-lookup"><span data-stu-id="e8967-246">Y</span></span><br/><span data-ttu-id="e8967-247">（每文档）</span><span class="sxs-lookup"><span data-stu-id="e8967-247">(Per document)</span></span>||
||<span data-ttu-id="e8967-248">设置更改事件</span><span class="sxs-lookup"><span data-stu-id="e8967-248">Settings changed events</span></span>|<span data-ttu-id="e8967-249">是</span><span class="sxs-lookup"><span data-stu-id="e8967-249">Y</span></span>|<span data-ttu-id="e8967-250">是</span><span class="sxs-lookup"><span data-stu-id="e8967-250">Y</span></span>||<span data-ttu-id="e8967-251">是</span><span class="sxs-lookup"><span data-stu-id="e8967-251">Y</span></span>|<span data-ttu-id="e8967-252">是</span><span class="sxs-lookup"><span data-stu-id="e8967-252">Y</span></span>||
||<span data-ttu-id="e8967-253">获取活动视图模式</span><span class="sxs-lookup"><span data-stu-id="e8967-253">Get active view mode</span></span><br/><span data-ttu-id="e8967-254">和视图更改事件</span><span class="sxs-lookup"><span data-stu-id="e8967-254">and view changed events</span></span>||||<span data-ttu-id="e8967-255">是</span><span class="sxs-lookup"><span data-stu-id="e8967-255">Y</span></span>|||
||<span data-ttu-id="e8967-256">转到文档中</span><span class="sxs-lookup"><span data-stu-id="e8967-256">Navigate to locations</span></span><br/><span data-ttu-id="e8967-257">的相应位置</span><span class="sxs-lookup"><span data-stu-id="e8967-257">in the document</span></span>||<span data-ttu-id="e8967-258">是</span><span class="sxs-lookup"><span data-stu-id="e8967-258">Y</span></span>||<span data-ttu-id="e8967-259">是</span><span class="sxs-lookup"><span data-stu-id="e8967-259">Y</span></span>|<span data-ttu-id="e8967-260">是</span><span class="sxs-lookup"><span data-stu-id="e8967-260">Y</span></span>||
||<span data-ttu-id="e8967-261">使用规则和 RegEx</span><span class="sxs-lookup"><span data-stu-id="e8967-261">Activate contextually</span></span><br/><span data-ttu-id="e8967-262">执行上下文式激活</span><span class="sxs-lookup"><span data-stu-id="e8967-262">using rules and RegEx</span></span>|||<span data-ttu-id="e8967-263">是</span><span class="sxs-lookup"><span data-stu-id="e8967-263">Y</span></span>||||
||<span data-ttu-id="e8967-264">读取项目属性</span><span class="sxs-lookup"><span data-stu-id="e8967-264">Read Item properties</span></span>|||<span data-ttu-id="e8967-265">是</span><span class="sxs-lookup"><span data-stu-id="e8967-265">Y</span></span>||||
||<span data-ttu-id="e8967-266">读取用户配置文件</span><span class="sxs-lookup"><span data-stu-id="e8967-266">Read User profile</span></span>|||<span data-ttu-id="e8967-267">是</span><span class="sxs-lookup"><span data-stu-id="e8967-267">Y</span></span>||||
||<span data-ttu-id="e8967-268">获取附件</span><span class="sxs-lookup"><span data-stu-id="e8967-268">Get attachments</span></span>|||<span data-ttu-id="e8967-269">是</span><span class="sxs-lookup"><span data-stu-id="e8967-269">Y</span></span>||||
||<span data-ttu-id="e8967-270">获取用户标识令牌</span><span class="sxs-lookup"><span data-stu-id="e8967-270">Get User identity token</span></span>|||<span data-ttu-id="e8967-271">是</span><span class="sxs-lookup"><span data-stu-id="e8967-271">Y</span></span>||||
||<span data-ttu-id="e8967-272">调用 Exchange Web 服务</span><span class="sxs-lookup"><span data-stu-id="e8967-272">Call Exchange Web Services</span></span>|||<span data-ttu-id="e8967-273">是</span><span class="sxs-lookup"><span data-stu-id="e8967-273">Y</span></span>||||
