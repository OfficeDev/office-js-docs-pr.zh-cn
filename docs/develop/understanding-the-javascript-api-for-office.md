---
title: 了解适用于 Office 的 JavaScript API
description: ''
ms.date: 10/17/2018
ms.openlocfilehash: 58829c623c06225bcc7d15925fb02a082df039c6
ms.sourcegitcommit: a6d6348075c1abed76d2146ddfc099b0151fe403
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/19/2018
ms.locfileid: "25640090"
---
# <a name="understanding-the-javascript-api-for-office"></a><span data-ttu-id="6f7d6-102">了解适用于 Office 的 JavaScript API</span><span class="sxs-lookup"><span data-stu-id="6f7d6-102">Understanding the JavaScript API for Office</span></span>

<span data-ttu-id="6f7d6-p101">本文提供了有关适用于 Office 的 JavaScript API 的信息以及使用方法。有关参考信息，请参阅 [适用于 Office 的 JavaScript API](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js)。有关将 Visual Studio 项目文件更新到适用于 Office 的 JavaScript API 的最新当前版本的信息，请参阅 [更新适用于 Office 的 JavaScript API 版本和清单架构文件](update-your-javascript-api-for-office-and-manifest-schema-version.md)。</span><span class="sxs-lookup"><span data-stu-id="6f7d6-p101">This article provides information about the JavaScript API for Office and how to use it. For reference information, see [JavaScript API for Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js). For information about updating Visual Studio project files to the most current version of the JavaScript API for Office, see [Update the version of your JavaScript API for Office and manifest schema files](update-your-javascript-api-for-office-and-manifest-schema-version.md).</span></span>

> [!NOTE]
> <span data-ttu-id="6f7d6-p102">如果计划将加载项[发布](../publish/publish.md)到 AppSource 并适用于 Office 体验，请务必遵循 [AppSource 验证策略](https://docs.microsoft.com/office/dev/store/validation-policies)。例如，加载项必须适用于支持已定义方法的所有平台，才能通过验证（欲知详请，请参阅[第 4.12 部分](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably)以及 [Office 加载项主机和可用性页面](../overview/office-add-in-availability.md)）。</span><span class="sxs-lookup"><span data-stu-id="6f7d6-p102">If you plan to [publish](../publish/publish.md) your add-in to AppSource and make it available within the Office experience, make sure that you conform to the [AppSource validation policies](https://docs.microsoft.com/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)).</span></span> 

## <a name="referencing-the-javascript-api-for-office-library-in-your-add-in"></a><span data-ttu-id="6f7d6-108">在加载项中引用适用于 Office 的 JavaScript API 库</span><span class="sxs-lookup"><span data-stu-id="6f7d6-108">Referencing the JavaScript API for Office library in your add-in</span></span>

<span data-ttu-id="6f7d6-p103">[适用于 Office 的 JavaScript](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js) 库包含 Office.js 文件和关联的特定于主机应用程序的 .js 文件，例如 Excel-15.js 和 Outlook-15.js。引用该 API 最简单的方法是通过添加以下 `<script>` 到你的页面的 `<head>` 标记来使用我们的 CDN：</span><span class="sxs-lookup"><span data-stu-id="6f7d6-p103">The [JavaScript API for Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js) library consists of the Office.js file and associated host application-specific .js files, such as Excel-15.js and Outlook-15.js. The simplest method of referencing the API is using our CDN by adding the following `<script>` to your page's `<head>` tag:</span></span>  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

<span data-ttu-id="6f7d6-111">这将在加载项首次加载时下载并缓存适用于 Office 的 JavaScript API 文件，以确保对特定版本使用 Office.js 及其关联文件的最新实现。</span><span class="sxs-lookup"><span data-stu-id="6f7d6-111">This will download and cache the JavaScript API for Office files the first time your add-in loads to make sure that it is using the most up-to-date implementation of Office.js and its associated files for the specified version.</span></span>

<span data-ttu-id="6f7d6-112">有关 Office.js CDN 的更多详细信息（包括如何处理版本控制和向后兼容性），请参阅[从内容分发网络 (CDN) 引用适用于 Office 的 JavaScript API 库](referencing-the-javascript-api-for-office-library-from-its-cdn.md)。</span><span class="sxs-lookup"><span data-stu-id="6f7d6-112">For more details around the Office.js CDN, including how versioning and backward compatability is handled, see [Referencing the JavaScript API for Office library from its content delivery network (CDN)](referencing-the-javascript-api-for-office-library-from-its-cdn.md).</span></span>

## <a name="initializing-your-add-in"></a><span data-ttu-id="6f7d6-113">初始化加载项</span><span class="sxs-lookup"><span data-stu-id="6f7d6-113">Initializing your add-in</span></span>

<span data-ttu-id="6f7d6-114">**适用于：** 所有加载项类型</span><span class="sxs-lookup"><span data-stu-id="6f7d6-114">**Applies to:** All add-in types</span></span>

<span data-ttu-id="6f7d6-115">Office 加载项通常有启动逻辑，以执行以下事项：</span><span class="sxs-lookup"><span data-stu-id="6f7d6-115">Office add-ins often have start-up logic to do things such as:</span></span>

- <span data-ttu-id="6f7d6-116">检查用户的 Office 版本是否支持你的代码调用的所有 Office Api。</span><span class="sxs-lookup"><span data-stu-id="6f7d6-116">Check that the user's version of Office will support all the Office APIs that your code calls.</span></span>

- <span data-ttu-id="6f7d6-117">确保存在某些工件，如具有特定名称的工作表。</span><span class="sxs-lookup"><span data-stu-id="6f7d6-117">Ensure the existence of certain artifacts, such as worksheet with a specific name.</span></span>

- <span data-ttu-id="6f7d6-118">提示用户选择 Excel 中的一些单元格，然后插入使用这些选定值初始化的图表。</span><span class="sxs-lookup"><span data-stu-id="6f7d6-118">You can use the initialize event handler to implement common add-in initialization scenarios, such as prompting the user to select some cells in Excel, and then inserting a chart initialized with those selected values.</span></span>

- <span data-ttu-id="6f7d6-119">建立绑定。</span><span class="sxs-lookup"><span data-stu-id="6f7d6-119">Establish bindings.</span></span>

- <span data-ttu-id="6f7d6-120">使用 Office 对话框 API 提示用户输入默认加载项设置值。</span><span class="sxs-lookup"><span data-stu-id="6f7d6-120">Use the Office dialog API to prompt the user for default add-in settings values.</span></span>

<span data-ttu-id="6f7d6-p104">但是，在完全加载完库之前，您启动代码不得调用任何 Office.js Api。有两种方法让您的代码可以确保加载库。这些将在以下章节加以说明：</span><span class="sxs-lookup"><span data-stu-id="6f7d6-p104">But your start-up code must not call any Office.js APIs until the library is fully loaded. There are two ways that your code can ensure that the library is loaded. They are described in the sections below. We recommend that you use the newer, more flexible, technique, calling . The older technique, assigning a handler to , is still supported. See also Major differences between Office.initialize and Office.onReady().</span></span> 

- [<span data-ttu-id="6f7d6-124">使用 Office.onReady() 初始化</span><span class="sxs-lookup"><span data-stu-id="6f7d6-124">Initialize with Office.onReady()</span></span>](#initialize-with-officeonready)
- [<span data-ttu-id="6f7d6-125">使用 Office.initialize 初始化</span><span class="sxs-lookup"><span data-stu-id="6f7d6-125">Initialize with Office.initialize</span></span>](#initialize-with-officeinitialize)

<span data-ttu-id="6f7d6-126">有关这些技术中的差异的信息，请参阅 [Office.initialize 和 Office.onReady() 之间的主要区别](#major-differences-between-officeinitialize-and-officeonready)。</span><span class="sxs-lookup"><span data-stu-id="6f7d6-126">For information about the differences in these techniques, see [Major differences between Office.initialize and Office.onReady()](#major-differences-between-officeinitialize-and-officeonready).</span></span> <span data-ttu-id="6f7d6-127">若要详细了解加载项初始化时的事件发生顺序，请参阅[加载 DOM 和运行时环境](loading-the-dom-and-runtime-environment.md)。</span><span class="sxs-lookup"><span data-stu-id="6f7d6-127">For more detail about the sequence of events when an add-in is initialized, see [Loading the DOM and runtime environment](loading-the-dom-and-runtime-environment.md).</span></span>

### <a name="initialize-with-officeonready"></a><span data-ttu-id="6f7d6-128">使用 Office.onReady() 初始化</span><span class="sxs-lookup"><span data-stu-id="6f7d6-128">Initialize with Office.onReady()</span></span>

<span data-ttu-id="6f7d6-p106">`Office.onReady()` 是返回承诺对象，同时检查 Office.js 库是否完全加载的异步方法。只有在加载库后，它才会将承诺解析为对象，这将使用 `Office.HostType` 枚举值 (`Excel`，`Word`等) 和与平台 `Office.PlatformType` 枚举值 (`PC`，`Mac`，`OfficeOnline`，等）指定 Office 主机应用程序。如果在调用 `Office.onReady()` 时已加载库，则承诺立即解析。</span><span class="sxs-lookup"><span data-stu-id="6f7d6-p106">`Office.onReady()` is an asynchronous method that returns a Promise object while it checks to see if the Office.js library is fully loaded. When, and only when, the library is loaded, it resolves the Promise as an object that specifies the Office host application with an `Office.HostType` enum value (`Excel`, `Word`, etc.) and the platform with an `Office.PlatformType` enum value (`PC`, `Mac`, `OfficeOnline`, etc.). If the library is already loaded when `Office.onReady()` is called, then the Promise resolves immediately.</span></span>

<span data-ttu-id="6f7d6-p107">调用 `Office.onReady()` 的一种方法是，将其传递给回调方法。下面是一个示例：</span><span class="sxs-lookup"><span data-stu-id="6f7d6-p107">One way to call `Office.onReady()` is to pass it a callback method. Here's an example:</span></span>

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

<span data-ttu-id="6f7d6-p108">或者，您可以将 `then()` 方法与 `Office.onReady()` 的调用链接而不是传递回调。例如，下面的代码将检查用户的 Excel 版本是否支持加载项可能调用的所有 Api。</span><span class="sxs-lookup"><span data-stu-id="6f7d6-p108">Alternatively, you can chain a `then()` method to the call of `Office.onReady()`, instead of passing a callback. For example, the following code checks to see that the user's version of Excel supports all the APIs that the add-in might call.</span></span>

```js
Office.onReady()
    .then(function() {
        if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
            console.log("Sorry, this add-in only works with newer versions of Excel.");
        }
    });
```

<span data-ttu-id="6f7d6-136">以下是在 TypeScript 中使用 `async` 和 `await` 关键字的相同示例：</span><span class="sxs-lookup"><span data-stu-id="6f7d6-136">Here is the same example using the `async` and `await` keywords in TypeScript:</span></span>

```typescript
(async () => {
    await Office.onReady();
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
})();
```

<span data-ttu-id="6f7d6-p109">如果你使用其他 JavaScript 框架，其中包括它们自己的初始化处理程序或测试，那么它们*通常*应放置在对 `Office.onReady()` 的响应内。例如，会对[JQuery](https://jquery.com) 的 `$(document).ready()` 函数进行以下引用：</span><span class="sxs-lookup"><span data-stu-id="6f7d6-p109">If you are using additional JavaScript frameworks that include their own initialization handler or tests, these should be *usually* be placed within the response to `Office.onReady()`. For example, [JQuery's](https://jquery.com) `$(document).ready()` function would be referenced as follows:</span></span>

```js
Office.onReady(function() {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
});
```

<span data-ttu-id="6f7d6-p110">但是，这一做法存在一些例外。例如，假设您想要在浏览器中打开您的加载项（而不是 侧加载到 Office 主机），从而使用浏览器工具调试您的 UI。由于 Office.js 无法在浏览器中加载，`onReady` 将无法运行，同时如果在 Office `onReady` 内调用，`$(document).ready` 将无法运行。另一个异常：加载加载项期间，您希望在任务窗格中显示进度指示器。在此方案中，您的代码应调用 jQuery `ready`，并使用它的回调以呈现进度指示器。然后，Office `onReady` 的回调可以替换最终用户界面的进度指示器。</span><span class="sxs-lookup"><span data-stu-id="6f7d6-p110">However, there are exceptions to this practice. For example, suppose you want to open your add-in in a browser (instead of sideload it in an Office host) in order to debug your UI with browser tools. Since Office.js won't load in the browser, `onReady` won't run and the `$(document).ready` won't run if it's called inside the Office `onReady`. Another exception: you want a progress indicator to appear in the task pane while the add-in is loading. In this scenario, your code should call the jQuery `ready` and use it's callback to render the progress indicator. Then the Office `onReady`'s callback can replace the progress indicator with the final UI.</span></span> 

### <a name="initialize-with-officeinitialize"></a><span data-ttu-id="6f7d6-145">使用 Office.initialize 初始化</span><span class="sxs-lookup"><span data-stu-id="6f7d6-145">Initialize with Office.initialize</span></span>

<span data-ttu-id="6f7d6-p111">当 Office.js 库完全加载并可供用户交互时，初始化事件触发。您可以分配一个处理程序给实施初始化逻辑的 `Office.initialize`。以下是检查用户的 Excel 版本是否支持所有可能调用加载项的 Api 示例。</span><span class="sxs-lookup"><span data-stu-id="6f7d6-p111">An initialize event fires when the Office.js library is fully loaded and ready for user interaction. You can assign a handler to `Office.initialize` that implements your initialization logic. The following is an example that checks to see that the user's version of Excel supports all the APIs that the add-in might call.</span></span>

```js
Office.initialize = function () {
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
};
```

<span data-ttu-id="6f7d6-p112">如果你使用其他 JavaScript 框架，其中包括它们自己的初始化处理程序或测试，那么它们*通常*应放置在 `Office.initialize` 事件内。（但是，更早版本**与 Office.onReady() 初始化**一节描述的异常也适用于这种情况。）例如，[JQuery](https://jquery.com) 的 `$(document).ready()` 函数会被引用为：</span><span class="sxs-lookup"><span data-stu-id="6f7d6-p112">If you are using additional JavaScript frameworks that include their own initialization handler or tests, these should *usually* be placed within the `Office.initialize` event. (But the exceptions described in the **Initialize with Office.onReady()** section earlier apply in this case also.) For example, [JQuery's](https://jquery.com) `$(document).ready()` function would be referenced as follows:</span></span>

```js
Office.initialize = function () {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
  };
```

<span data-ttu-id="6f7d6-p113">对于任务窗格和内容加载项，提供其他 `Office.initialize` _  _ 参数。此参数指定如何添加加载项到当前文档。你可以使用此参数针对首次插入加载项时和加载项已存在于文档中时实施不同的逻辑。</span><span class="sxs-lookup"><span data-stu-id="6f7d6-p113">For task pane and content add-ins, `Office.initialize` provides an additional _reason_ parameter. This parameter specifies how an add-in was added to the current document. You can use this to provide different logic for when an add-in is first inserted versus when it already existed within the document.</span></span>

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

<span data-ttu-id="6f7d6-154">有关详细信息，请参阅 [Office.initialize 事件](https://docs.microsoft.com/javascript/api/office?view=office-js)和 [InitializationReason 枚举](https://docs.microsoft.com/javascript/api/office/office.initializationreason?view=office-js)。</span><span class="sxs-lookup"><span data-stu-id="6f7d6-154">For more information, see [Office.initialize Event](https://docs.microsoft.com/javascript/api/office?view=office-js) and [InitializationReason Enumeration](https://docs.microsoft.com/javascript/api/office/office.initializationreason?view=office-js).</span></span>

> [!NOTE]
> <span data-ttu-id="6f7d6-155">目前，无论是否还调用 `Office.onReady()`，你都必须设置 `Office.Initialize`。</span><span class="sxs-lookup"><span data-stu-id="6f7d6-155">Currently, you must set `Office.Initialize`, regardless of whether `Office.onReady()` is also called.</span></span> <span data-ttu-id="6f7d6-156">如果没有使用 `Office.Initialize`，则可以将其设置为空函数，如以下示例所示。</span><span class="sxs-lookup"><span data-stu-id="6f7d6-156">If you have no use for `Office.Initialize`, you can set it to an empty function as shown in the following example.</span></span>
> 
>```js
>Office.initialize = function () {};
>```

### <a name="major-differences-between-officeinitialize-and-officeonready"></a><span data-ttu-id="6f7d6-157">Office.initialize 和 Office.onReady 的主要区别</span><span class="sxs-lookup"><span data-stu-id="6f7d6-157">Major differences between Office.initialize and Office.onReady</span></span>

- <span data-ttu-id="6f7d6-p115">您仅可分配一个处理程序到 `Office.initialize` ，同时它由 Office 基础架构仅调用一次；但是，你可以在代码中的不同位置调用 `Office.onReady()` 并可使用不同的回调。例如，一旦使用运行初始化逻辑的调用加载你的自定义脚本，你的代码即可调用 `Office.onReady()` ；同时，你的代码还可在任务窗格中有一个按钮，其脚本以不同的回调来调用 `Office.onReady()`。如果是这样，单击按钮时将运行第二个回调。</span><span class="sxs-lookup"><span data-stu-id="6f7d6-p115">You can assign only one handler to `Office.initialize` and it is called, only once, by the Office infrastructure; but you can call `Office.onReady()` in different places in your code and use different callbacks. For example, your code could call `Office.onReady()` as soon as your custom script loads with a callback that runs initialization logic; and your code could also have a button in the task pane, whose script calls `Office.onReady()` with a different callback. If so, the second callback runs when the button is clicked.</span></span>

- <span data-ttu-id="6f7d6-p116">`Office.initialize` 事件在 Office.js 初始化本身的内部过程末尾触发。这在内部过程结束后*立即*触发。如果事件触发后指定处理程序给事件的代码执行时间过长，则不运行你的处理程序。例如，如果你使用 WebPack 任务管理器，它可能在加载 Office.js 后，但在加载你的自定义 JavaScript 之前配置加载项主页以加载 polyfill 文件。脚本加载并分配该处理程序时，初始化事件已经发生。但是，调用 `Office.onReady()` 不会“过晚”。如果初始化事件已经发生，则回调立即运行。</span><span class="sxs-lookup"><span data-stu-id="6f7d6-p116">The `Office.initialize` event fires at the end of the internal process in which Office.js initializes itself. And it fires *immediately* after the internal process ends. If the code in which you assign a handler to the event executes too long after the event fires, then your handler doesn't run. For example, if you are using the WebPack task manager, it might configure the add-in's home page to load polyfill files after it loads Office.js but before it loads your custom JavaScript. By the time your script loads and assigns the handler, the initialize event has already happened. But it is never "too late" to call `Office.onReady()`. If the initialize event has already happened, the callback runs immediately.</span></span>

> [!NOTE]
> <span data-ttu-id="6f7d6-168">即使没有启动逻辑，也应该在加载加载项 JavaScript 时为 `Office.initialize` 分配一个空函数，如下例所示。</span><span class="sxs-lookup"><span data-stu-id="6f7d6-168">Even if you have no start-up logic, you should assign an empty function to `Office.initialize` when your add-in JavaScript loads, as shown in the following example.</span></span> <span data-ttu-id="6f7d6-169">某些 Office 主机和平台组合将不会加载任务窗格，直到触发初始化事件并运行指定的事件处理函数。</span><span class="sxs-lookup"><span data-stu-id="6f7d6-169">Some Office host and platform combinations won't load the task pane until the initialize event fires and the specified event handler function runs.</span></span>
> 
>```js
>Office.initialize = function () {};
>```

## <a name="office-javascript-api-object-model"></a><span data-ttu-id="6f7d6-170">Office JavaScript API 对象模型</span><span class="sxs-lookup"><span data-stu-id="6f7d6-170">Office JavaScript API object model</span></span>

<span data-ttu-id="6f7d6-p118">初始化后，加载项可以与主机 （例如 Excel、 Outlook）交互。[Office JavaScript API 对象模型](office-javascript-api-object-model.md)页上提供特定的使用模式的详细信息。此外，还有 [共享 API](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js) 及特定主机详细的参考文档。</span><span class="sxs-lookup"><span data-stu-id="6f7d6-p118">Once initialized, the add-in can interact with the host (e.g. Excel, Outlook). The [Office JavaScript API object model](office-javascript-api-object-model.md) page has more details on specific usage patterns. There is also detailed reference documentation for both [shared APIs](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js) and specific hosts.</span></span>

## <a name="api-support-matrix"></a><span data-ttu-id="6f7d6-174">API 支持矩阵</span><span class="sxs-lookup"><span data-stu-id="6f7d6-174">API support matrix</span></span>

<span data-ttu-id="6f7d6-175">下表总结了各种类型的加载项（内容、任务窗格和 Outlook）支持的 API 和功能，以及使用[适用于 Office 的 JavaScript API v1.1 支持的 1.1 加载项清单架构和功能](update-your-javascript-api-for-office-and-manifest-schema-version.md)指定加载项支持的 Office 主机应用时，可以托管它们的 Office 应用。</span><span class="sxs-lookup"><span data-stu-id="6f7d6-175">This table summarizes the API and features supported across add-in types (content, task pane, and Outlook) and the Office applications that can host them when you specify the Office host applications your add-in supports by using the [1.1 add-in manifest schema and features supported by v1.1 JavaScript API for Office](update-your-javascript-api-for-office-and-manifest-schema-version.md).</span></span>


|||||||||
|:-----|:-----|:-----:|:-----:|:-----:|:-----:|:-----:|:-----:|
||<span data-ttu-id="6f7d6-176">**主机名**</span><span class="sxs-lookup"><span data-stu-id="6f7d6-176">**Host name**</span></span>|<span data-ttu-id="6f7d6-177">数据库</span><span class="sxs-lookup"><span data-stu-id="6f7d6-177">Database</span></span>|<span data-ttu-id="6f7d6-178">工作簿</span><span class="sxs-lookup"><span data-stu-id="6f7d6-178">Workbook</span></span>|<span data-ttu-id="6f7d6-179">邮箱</span><span class="sxs-lookup"><span data-stu-id="6f7d6-179">Mailbox</span></span>|<span data-ttu-id="6f7d6-180">演示文稿</span><span class="sxs-lookup"><span data-stu-id="6f7d6-180">Presentation</span></span>|<span data-ttu-id="6f7d6-181">文档</span><span class="sxs-lookup"><span data-stu-id="6f7d6-181">Document</span></span>|<span data-ttu-id="6f7d6-182">项目</span><span class="sxs-lookup"><span data-stu-id="6f7d6-182">Project</span></span>|
||<span data-ttu-id="6f7d6-183">**支持的\*\*\*\*主机应用程序**</span><span class="sxs-lookup"><span data-stu-id="6f7d6-183">**Supported** **Host applications**</span></span>|<span data-ttu-id="6f7d6-184">Access Web 应用</span><span class="sxs-lookup"><span data-stu-id="6f7d6-184">Access web apps</span></span>|<span data-ttu-id="6f7d6-185">Excel、</span><span class="sxs-lookup"><span data-stu-id="6f7d6-185">Excel,</span></span><br/><span data-ttu-id="6f7d6-186">Excel Online</span><span class="sxs-lookup"><span data-stu-id="6f7d6-186">Excel Online</span></span>|<span data-ttu-id="6f7d6-187">Outlook、</span><span class="sxs-lookup"><span data-stu-id="6f7d6-187">Outlook,</span></span><br/><span data-ttu-id="6f7d6-188">Outlook Web App、</span><span class="sxs-lookup"><span data-stu-id="6f7d6-188">Outlook Web App,</span></span><br/><span data-ttu-id="6f7d6-189">适用于设备的 OWA</span><span class="sxs-lookup"><span data-stu-id="6f7d6-189">OWA for Devices</span></span>|<span data-ttu-id="6f7d6-190">PowerPoint、</span><span class="sxs-lookup"><span data-stu-id="6f7d6-190">PowerPoint,</span></span><br/><span data-ttu-id="6f7d6-191">PowerPoint Online</span><span class="sxs-lookup"><span data-stu-id="6f7d6-191">PowerPoint Online</span></span>|<span data-ttu-id="6f7d6-192">Word</span><span class="sxs-lookup"><span data-stu-id="6f7d6-192">Word</span></span>|<span data-ttu-id="6f7d6-193">项目</span><span class="sxs-lookup"><span data-stu-id="6f7d6-193">Project</span></span>|
|<span data-ttu-id="6f7d6-194">**支持的加载项类型**</span><span class="sxs-lookup"><span data-stu-id="6f7d6-194">**Supported add-in types**</span></span>|<span data-ttu-id="6f7d6-195">内容</span><span class="sxs-lookup"><span data-stu-id="6f7d6-195">Content</span></span>|<span data-ttu-id="6f7d6-196">是</span><span class="sxs-lookup"><span data-stu-id="6f7d6-196">Y</span></span>|<span data-ttu-id="6f7d6-197">是</span><span class="sxs-lookup"><span data-stu-id="6f7d6-197">Y</span></span>||<span data-ttu-id="6f7d6-198">是</span><span class="sxs-lookup"><span data-stu-id="6f7d6-198">Y</span></span>|||
||<span data-ttu-id="6f7d6-199">任务窗格</span><span class="sxs-lookup"><span data-stu-id="6f7d6-199">Task pane</span></span>||<span data-ttu-id="6f7d6-200">是</span><span class="sxs-lookup"><span data-stu-id="6f7d6-200">Y</span></span>||<span data-ttu-id="6f7d6-201">是</span><span class="sxs-lookup"><span data-stu-id="6f7d6-201">Y</span></span>|<span data-ttu-id="6f7d6-202">是</span><span class="sxs-lookup"><span data-stu-id="6f7d6-202">Y</span></span>|<span data-ttu-id="6f7d6-203">是</span><span class="sxs-lookup"><span data-stu-id="6f7d6-203">Y</span></span>|
||<span data-ttu-id="6f7d6-204">Outlook</span><span class="sxs-lookup"><span data-stu-id="6f7d6-204">Outlook</span></span>|||<span data-ttu-id="6f7d6-205">是</span><span class="sxs-lookup"><span data-stu-id="6f7d6-205">Y</span></span>||||
|<span data-ttu-id="6f7d6-206">**支持的 API 功能**</span><span class="sxs-lookup"><span data-stu-id="6f7d6-206">**Supported API features**</span></span>|<span data-ttu-id="6f7d6-207">读/写文本</span><span class="sxs-lookup"><span data-stu-id="6f7d6-207">Read/Write Text</span></span>||<span data-ttu-id="6f7d6-208">是</span><span class="sxs-lookup"><span data-stu-id="6f7d6-208">Y</span></span>||<span data-ttu-id="6f7d6-209">是</span><span class="sxs-lookup"><span data-stu-id="6f7d6-209">Y</span></span>|<span data-ttu-id="6f7d6-210">是</span><span class="sxs-lookup"><span data-stu-id="6f7d6-210">Y</span></span>|<span data-ttu-id="6f7d6-211">是</span><span class="sxs-lookup"><span data-stu-id="6f7d6-211">Y</span></span><br/><span data-ttu-id="6f7d6-212">（只读）</span><span class="sxs-lookup"><span data-stu-id="6f7d6-212">(Read only)</span></span>|
||<span data-ttu-id="6f7d6-213">读/写矩阵</span><span class="sxs-lookup"><span data-stu-id="6f7d6-213">Read/Write Matrix</span></span>||<span data-ttu-id="6f7d6-214">是</span><span class="sxs-lookup"><span data-stu-id="6f7d6-214">Y</span></span>|||<span data-ttu-id="6f7d6-215">是</span><span class="sxs-lookup"><span data-stu-id="6f7d6-215">Y</span></span>||
||<span data-ttu-id="6f7d6-216">读/写表</span><span class="sxs-lookup"><span data-stu-id="6f7d6-216">Read/Write Table</span></span>||<span data-ttu-id="6f7d6-217">是</span><span class="sxs-lookup"><span data-stu-id="6f7d6-217">Y</span></span>|||<span data-ttu-id="6f7d6-218">是</span><span class="sxs-lookup"><span data-stu-id="6f7d6-218">Y</span></span>||
||<span data-ttu-id="6f7d6-219">读/写 HTML</span><span class="sxs-lookup"><span data-stu-id="6f7d6-219">Read/Write HTML</span></span>|||||<span data-ttu-id="6f7d6-220">是</span><span class="sxs-lookup"><span data-stu-id="6f7d6-220">Y</span></span>||
||<span data-ttu-id="6f7d6-221">读/写</span><span class="sxs-lookup"><span data-stu-id="6f7d6-221">Read/Write</span></span><br/><span data-ttu-id="6f7d6-222">Office Open XML</span><span class="sxs-lookup"><span data-stu-id="6f7d6-222">Office Open XML</span></span>|||||<span data-ttu-id="6f7d6-223">是</span><span class="sxs-lookup"><span data-stu-id="6f7d6-223">Y</span></span>||
||<span data-ttu-id="6f7d6-224">读取任务、资源、视图和字段属性</span><span class="sxs-lookup"><span data-stu-id="6f7d6-224">Read task, resource, view, and field properties</span></span>||||||<span data-ttu-id="6f7d6-225">是</span><span class="sxs-lookup"><span data-stu-id="6f7d6-225">Y</span></span>|
||<span data-ttu-id="6f7d6-226">选择已更改事件</span><span class="sxs-lookup"><span data-stu-id="6f7d6-226">Selection changed events</span></span>||<span data-ttu-id="6f7d6-227">是</span><span class="sxs-lookup"><span data-stu-id="6f7d6-227">Y</span></span>|||<span data-ttu-id="6f7d6-228">是</span><span class="sxs-lookup"><span data-stu-id="6f7d6-228">Y</span></span>||
||<span data-ttu-id="6f7d6-229">获取整个文档</span><span class="sxs-lookup"><span data-stu-id="6f7d6-229">Get whole document</span></span>||||<span data-ttu-id="6f7d6-230">是</span><span class="sxs-lookup"><span data-stu-id="6f7d6-230">Y</span></span>|<span data-ttu-id="6f7d6-231">是</span><span class="sxs-lookup"><span data-stu-id="6f7d6-231">Y</span></span>||
||<span data-ttu-id="6f7d6-232">绑定和绑定事件</span><span class="sxs-lookup"><span data-stu-id="6f7d6-232">Bindings and binding events</span></span>|<span data-ttu-id="6f7d6-233">是</span><span class="sxs-lookup"><span data-stu-id="6f7d6-233">Y</span></span><br/><span data-ttu-id="6f7d6-234">（仅限完全和部分表格绑定）</span><span class="sxs-lookup"><span data-stu-id="6f7d6-234">(Only full and partial table bindings)</span></span>|<span data-ttu-id="6f7d6-235">是</span><span class="sxs-lookup"><span data-stu-id="6f7d6-235">Y</span></span>|||<span data-ttu-id="6f7d6-236">是</span><span class="sxs-lookup"><span data-stu-id="6f7d6-236">Y</span></span>||
||<span data-ttu-id="6f7d6-237">读/写自定义 XML 部分</span><span class="sxs-lookup"><span data-stu-id="6f7d6-237">Read/Write Custom XML Parts</span></span>|||||<span data-ttu-id="6f7d6-238">是</span><span class="sxs-lookup"><span data-stu-id="6f7d6-238">Y</span></span>||
||<span data-ttu-id="6f7d6-239">暂留加载项状态数据（设置）</span><span class="sxs-lookup"><span data-stu-id="6f7d6-239">Persist add-in state data (settings)</span></span>|<span data-ttu-id="6f7d6-240">是</span><span class="sxs-lookup"><span data-stu-id="6f7d6-240">Y</span></span><br/><span data-ttu-id="6f7d6-241">（每主机加载项）</span><span class="sxs-lookup"><span data-stu-id="6f7d6-241">(Per host add-in)</span></span>|<span data-ttu-id="6f7d6-242">是</span><span class="sxs-lookup"><span data-stu-id="6f7d6-242">Y</span></span><br/><span data-ttu-id="6f7d6-243">（每文档）</span><span class="sxs-lookup"><span data-stu-id="6f7d6-243">(Per document)</span></span>|<span data-ttu-id="6f7d6-244">是</span><span class="sxs-lookup"><span data-stu-id="6f7d6-244">Y</span></span><br/><span data-ttu-id="6f7d6-245">（每邮箱）</span><span class="sxs-lookup"><span data-stu-id="6f7d6-245">(Per mailbox)</span></span>|<span data-ttu-id="6f7d6-246">是</span><span class="sxs-lookup"><span data-stu-id="6f7d6-246">Y</span></span><br/><span data-ttu-id="6f7d6-247">（每文档）</span><span class="sxs-lookup"><span data-stu-id="6f7d6-247">(Per document)</span></span>|<span data-ttu-id="6f7d6-248">是</span><span class="sxs-lookup"><span data-stu-id="6f7d6-248">Y</span></span><br/><span data-ttu-id="6f7d6-249">（每文档）</span><span class="sxs-lookup"><span data-stu-id="6f7d6-249">(Per document)</span></span>||
||<span data-ttu-id="6f7d6-250">设置更改事件</span><span class="sxs-lookup"><span data-stu-id="6f7d6-250">Settings changed events</span></span>|<span data-ttu-id="6f7d6-251">是</span><span class="sxs-lookup"><span data-stu-id="6f7d6-251">Y</span></span>|<span data-ttu-id="6f7d6-252">是</span><span class="sxs-lookup"><span data-stu-id="6f7d6-252">Y</span></span>||<span data-ttu-id="6f7d6-253">是</span><span class="sxs-lookup"><span data-stu-id="6f7d6-253">Y</span></span>|<span data-ttu-id="6f7d6-254">是</span><span class="sxs-lookup"><span data-stu-id="6f7d6-254">Y</span></span>||
||<span data-ttu-id="6f7d6-255">获取活动视图模式</span><span class="sxs-lookup"><span data-stu-id="6f7d6-255">Get active view mode</span></span><br/><span data-ttu-id="6f7d6-256">和视图更改事件</span><span class="sxs-lookup"><span data-stu-id="6f7d6-256">and view changed events</span></span>||||<span data-ttu-id="6f7d6-257">是</span><span class="sxs-lookup"><span data-stu-id="6f7d6-257">Y</span></span>|||
||<span data-ttu-id="6f7d6-258">转到文档中</span><span class="sxs-lookup"><span data-stu-id="6f7d6-258">Navigate to locations</span></span><br/><span data-ttu-id="6f7d6-259">的相应位置</span><span class="sxs-lookup"><span data-stu-id="6f7d6-259">in the document</span></span>||<span data-ttu-id="6f7d6-260">是</span><span class="sxs-lookup"><span data-stu-id="6f7d6-260">Y</span></span>||<span data-ttu-id="6f7d6-261">是</span><span class="sxs-lookup"><span data-stu-id="6f7d6-261">Y</span></span>|<span data-ttu-id="6f7d6-262">是</span><span class="sxs-lookup"><span data-stu-id="6f7d6-262">Y</span></span>||
||<span data-ttu-id="6f7d6-263">使用规则和 RegEx</span><span class="sxs-lookup"><span data-stu-id="6f7d6-263">Activate contextually</span></span><br/><span data-ttu-id="6f7d6-264">执行上下文式激活</span><span class="sxs-lookup"><span data-stu-id="6f7d6-264">using rules and RegEx</span></span>|||<span data-ttu-id="6f7d6-265">是</span><span class="sxs-lookup"><span data-stu-id="6f7d6-265">Y</span></span>||||
||<span data-ttu-id="6f7d6-266">读取项目属性</span><span class="sxs-lookup"><span data-stu-id="6f7d6-266">Read Item properties</span></span>|||<span data-ttu-id="6f7d6-267">是</span><span class="sxs-lookup"><span data-stu-id="6f7d6-267">Y</span></span>||||
||<span data-ttu-id="6f7d6-268">读取用户配置文件</span><span class="sxs-lookup"><span data-stu-id="6f7d6-268">Read User profile</span></span>|||<span data-ttu-id="6f7d6-269">是</span><span class="sxs-lookup"><span data-stu-id="6f7d6-269">Y</span></span>||||
||<span data-ttu-id="6f7d6-270">获取附件</span><span class="sxs-lookup"><span data-stu-id="6f7d6-270">Get attachments</span></span>|||<span data-ttu-id="6f7d6-271">是</span><span class="sxs-lookup"><span data-stu-id="6f7d6-271">Y</span></span>||||
||<span data-ttu-id="6f7d6-272">获取用户标识令牌</span><span class="sxs-lookup"><span data-stu-id="6f7d6-272">Get User identity token</span></span>|||<span data-ttu-id="6f7d6-273">是</span><span class="sxs-lookup"><span data-stu-id="6f7d6-273">Y</span></span>||||
||<span data-ttu-id="6f7d6-274">调用 Exchange Web 服务</span><span class="sxs-lookup"><span data-stu-id="6f7d6-274">Call Exchange Web Services</span></span>|||<span data-ttu-id="6f7d6-275">是</span><span class="sxs-lookup"><span data-stu-id="6f7d6-275">Y</span></span>||||
