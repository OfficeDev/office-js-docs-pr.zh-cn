---
title: 初始化 Office 加载项
description: 了解如何初始化 Office 外接程序。
ms.date: 02/27/2020
localization_priority: Normal
ms.openlocfilehash: 5dc9d0143ac9eaab18625e280891bd601fa9f899
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293322"
---
# <a name="initialize-your-office-add-in"></a><span data-ttu-id="109b9-103">初始化 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="109b9-103">Initialize your Office Add-in</span></span>

<span data-ttu-id="109b9-104">Office 加载项通常使用启动逻辑执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="109b9-104">Office Add-ins often have start-up logic to do things such as:</span></span>

- <span data-ttu-id="109b9-105">检查用户的 Office 版本是否支持您的代码调用的所有 Office Api。</span><span class="sxs-lookup"><span data-stu-id="109b9-105">Check that the user's version of Office supports all the Office APIs that your code calls.</span></span>

- <span data-ttu-id="109b9-106">确保存在某些项目，如具有特定名称的工作表。</span><span class="sxs-lookup"><span data-stu-id="109b9-106">Ensure the existence of certain artifacts, such as a worksheet with a specific name.</span></span>

- <span data-ttu-id="109b9-107">提示用户选择 Excel 中的某些单元格，然后插入用这些选定值初始化的图表。</span><span class="sxs-lookup"><span data-stu-id="109b9-107">Prompt the user to select some cells in Excel, and then insert a chart initialized with those selected values.</span></span>

- <span data-ttu-id="109b9-108">建立绑定。</span><span class="sxs-lookup"><span data-stu-id="109b9-108">Establish bindings.</span></span>

- <span data-ttu-id="109b9-109">使用 Office 对话框 API 提示用户输入默认的外接程序设置值。</span><span class="sxs-lookup"><span data-stu-id="109b9-109">Use the Office Dialog API to prompt the user for default add-in settings values.</span></span>

<span data-ttu-id="109b9-110">但是，在加载库之前，Office 外接程序无法成功调用任何 Office JavaScript Api。</span><span class="sxs-lookup"><span data-stu-id="109b9-110">However, an Office Add-in cannot successfully call any Office JavaScript APIs until the library has been loaded.</span></span> <span data-ttu-id="109b9-111">本文介绍了您的代码可确保库已加载的两种方法：</span><span class="sxs-lookup"><span data-stu-id="109b9-111">This article describes the two ways your code can ensure that the library has been loaded:</span></span>

- <span data-ttu-id="109b9-112">使用初始化 `Office.onReady()` 。</span><span class="sxs-lookup"><span data-stu-id="109b9-112">Initialize with `Office.onReady()`.</span></span>
- <span data-ttu-id="109b9-113">使用初始化 `Office.initialize` 。</span><span class="sxs-lookup"><span data-stu-id="109b9-113">Initialize with `Office.initialize`.</span></span>

> [!TIP]
> <span data-ttu-id="109b9-114">建议使用 `Office.onReady()` 取代 `Office.initialize`。</span><span class="sxs-lookup"><span data-stu-id="109b9-114">We recommend that you use `Office.onReady()` instead of `Office.initialize`.</span></span> <span data-ttu-id="109b9-115">尽管 `Office.initialize` 仍受支持，但 `Office.onReady()` 提供了更大的灵活性。</span><span class="sxs-lookup"><span data-stu-id="109b9-115">Although `Office.initialize` is still supported, `Office.onReady()` provides more flexibility.</span></span> <span data-ttu-id="109b9-116">只能将一个处理程序分配给 `Office.initialize` ，而只是通过 Office 基础结构调用一次。</span><span class="sxs-lookup"><span data-stu-id="109b9-116">You can assign only one handler to `Office.initialize` and it's called only once by the Office infrastructure.</span></span> <span data-ttu-id="109b9-117">您可以 `Office.onReady()` 在代码中的不同位置调用，并使用不同的回调。</span><span class="sxs-lookup"><span data-stu-id="109b9-117">You can call `Office.onReady()` in different places in your code and use different callbacks.</span></span>
> 
> <span data-ttu-id="109b9-118">有关这两种方法之间的差别信息，请参阅 [Office.initialize 和 Office.onReady() 之间的主要差别](#major-differences-between-officeinitialize-and-officeonready)。</span><span class="sxs-lookup"><span data-stu-id="109b9-118">For information about the differences in these techniques, see [Major differences between Office.initialize and Office.onReady()](#major-differences-between-officeinitialize-and-officeonready).</span></span>

<span data-ttu-id="109b9-119">有关初始化加载项时的事件顺序的更多详细信息，请参阅 [加载 DOM 和运行时环境](loading-the-dom-and-runtime-environment.md)。</span><span class="sxs-lookup"><span data-stu-id="109b9-119">For more details about the sequence of events when an add-in is initialized, see [Loading the DOM and runtime environment](loading-the-dom-and-runtime-environment.md).</span></span>

## <a name="initialize-with-officeonready"></a><span data-ttu-id="109b9-120">使用 Office.onReady() 进行初始化</span><span class="sxs-lookup"><span data-stu-id="109b9-120">Initialize with Office.onReady()</span></span>

<span data-ttu-id="109b9-121">`Office.onReady()` 是一种异步方法，它在检查是否已加载 Office.js 库时返回一个 [承诺](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) 对象。</span><span class="sxs-lookup"><span data-stu-id="109b9-121">`Office.onReady()` is an asynchronous method that returns a [Promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) object while it checks to see if the Office.js library is loaded.</span></span> <span data-ttu-id="109b9-122">加载库时，它会将承诺解析为一个对象，该对象指定具有 enum 值的 Office 客户端应用程序 `Office.HostType` (`Excel` 、、 `Word` 等 ) 以及具有 `Office.PlatformType` enum 值的平台 (`PC` 、 `Mac` 、、 `OfficeOnline` 等 ) 。</span><span class="sxs-lookup"><span data-stu-id="109b9-122">When the library is loaded, it resolves the Promise as an object that specifies the Office client application with an `Office.HostType` enum value (`Excel`, `Word`, etc.) and the platform with an `Office.PlatformType` enum value (`PC`, `Mac`, `OfficeOnline`, etc.).</span></span> <span data-ttu-id="109b9-123">如果在调用 `Office.onReady()` 时已加载库，则 Promise 将立即解析。</span><span class="sxs-lookup"><span data-stu-id="109b9-123">The Promise resolves immediately if the library is already loaded when `Office.onReady()` is called.</span></span>

<span data-ttu-id="109b9-124">调用 `Office.onReady()` 的一种方法是向其传递一个回调方法。</span><span class="sxs-lookup"><span data-stu-id="109b9-124">One way to call `Office.onReady()` is to pass it a callback method.</span></span> <span data-ttu-id="109b9-125">下面是一个示例：</span><span class="sxs-lookup"><span data-stu-id="109b9-125">Here's an example:</span></span>

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

<span data-ttu-id="109b9-126">或者，可以将 `then()` 方法链接到 `Office.onReady()` 的调用，而不是传递回调。</span><span class="sxs-lookup"><span data-stu-id="109b9-126">Alternatively, you can chain a `then()` method to the call of `Office.onReady()`, instead of passing a callback.</span></span> <span data-ttu-id="109b9-127">例如，以下代码检查用户的 Excel 版本是否支持加载项可能调用的所有 API。</span><span class="sxs-lookup"><span data-stu-id="109b9-127">For example, the following code checks to see that the user's version of Excel supports all the APIs that the add-in might call.</span></span>

```js
Office.onReady()
    .then(function() {
        if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
            console.log("Sorry, this add-in only works with newer versions of Excel.");
        }
    });
```

<span data-ttu-id="109b9-128">以下是在 TypeScript 中使用 `async` 和 `await` 关键字的同一示例：</span><span class="sxs-lookup"><span data-stu-id="109b9-128">Here is the same example using the `async` and `await` keywords in TypeScript:</span></span>

```typescript
(async () => {
    await Office.onReady();
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
})();
```

<span data-ttu-id="109b9-129">如果使用的是其他 JavaScript 框架，其中包括它们自己的初始化处理程序或测试，那么它们*通常*应放置在 `Office.onReady()` 的响应内。</span><span class="sxs-lookup"><span data-stu-id="109b9-129">If you are using additional JavaScript frameworks that include their own initialization handler or tests, these should *usually* be placed within the response to `Office.onReady()`.</span></span> <span data-ttu-id="109b9-130">例如，会对 [JQuery 的](https://jquery.com) `$(document).ready()` 函数进行以下引用：</span><span class="sxs-lookup"><span data-stu-id="109b9-130">For example, [JQuery's](https://jquery.com) `$(document).ready()` function would be referenced as follows:</span></span>

```js
Office.onReady(function() {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
});
```

<span data-ttu-id="109b9-131">但是，此做法存在例外情况。</span><span class="sxs-lookup"><span data-stu-id="109b9-131">However, there are exceptions to this practice.</span></span> <span data-ttu-id="109b9-132">例如，假设您想要在浏览器中打开加载项 (而不是将其旁加载在 Office 应用程序中) ，以便使用浏览器工具调试您的 UI。</span><span class="sxs-lookup"><span data-stu-id="109b9-132">For example, suppose you want to open your add-in in a browser (instead of sideload it in an Office application) in order to debug your UI with browser tools.</span></span> <span data-ttu-id="109b9-133">由于 Office.js 将不会在浏览器中加载，所以，`onReady` 将不会运行，且如果在 Office `$(document).ready` 内调用它，则 `onReady` 将不会运行。</span><span class="sxs-lookup"><span data-stu-id="109b9-133">Since Office.js won't load in the browser, `onReady` won't run and the `$(document).ready` won't run if it's called inside the Office `onReady`.</span></span> 

<span data-ttu-id="109b9-134">如果您希望在加载加载项时，任务窗格中显示进度指示器，另一个例外。</span><span class="sxs-lookup"><span data-stu-id="109b9-134">Another exception would be if you want a progress indicator to appear in the task pane while the add-in is loading.</span></span> <span data-ttu-id="109b9-135">在这种情况下，代码应调用 jQuery `ready` 并使用其回调来呈现进度指示器。</span><span class="sxs-lookup"><span data-stu-id="109b9-135">In this scenario, your code should call the jQuery `ready` and use its callback to render the progress indicator.</span></span> <span data-ttu-id="109b9-136">然后，Office `onReady` 的回调可将进度指示器替换为最终 UI。</span><span class="sxs-lookup"><span data-stu-id="109b9-136">Then the Office `onReady`'s callback can replace the progress indicator with the final UI.</span></span> 

## <a name="initialize-with-officeinitialize"></a><span data-ttu-id="109b9-137">使用 Office.initialize 进行初始化</span><span class="sxs-lookup"><span data-stu-id="109b9-137">Initialize with Office.initialize</span></span>

<span data-ttu-id="109b9-138">当 Office.js 库加载并准备好用于用户交互时将触发初始化事件。</span><span class="sxs-lookup"><span data-stu-id="109b9-138">An initialize event fires when the Office.js library is loaded and ready for user interaction.</span></span> <span data-ttu-id="109b9-139">可将处理程序分配到实现初始化逻辑的 `Office.initialize`。</span><span class="sxs-lookup"><span data-stu-id="109b9-139">You can assign a handler to `Office.initialize` that implements your initialization logic.</span></span> <span data-ttu-id="109b9-140">以下是检查用户的 Excel 版本是否支持加载项可能调用的所有 API 的示例。</span><span class="sxs-lookup"><span data-stu-id="109b9-140">The following is an example that checks to see that the user's version of Excel supports all the APIs that the add-in might call.</span></span>

```js
Office.initialize = function () {
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
};
```

<span data-ttu-id="109b9-141">如果使用的是包含其自己的初始化处理程序或测试的其他 JavaScript 框架，则 *通常* 应将这些框架放在 `Office.initialize` 事件 (在 OnReady 中的 \*\*Initialize ( # B2 \*\* 部分中所述的异常也会在此情况下应用) 。</span><span class="sxs-lookup"><span data-stu-id="109b9-141">If you are using additional JavaScript frameworks that include their own initialization handler or tests, these should *usually* be placed within the `Office.initialize` event (the exceptions described in the **Initialize with Office.onReady()** section earlier apply in this case also).</span></span> <span data-ttu-id="109b9-142">例如，会对 [JQuery 的](https://jquery.com) `$(document).ready()` 函数进行以下引用：</span><span class="sxs-lookup"><span data-stu-id="109b9-142">For example, [JQuery's](https://jquery.com) `$(document).ready()` function would be referenced as follows:</span></span>

```js
Office.initialize = function () {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
  };
```

<span data-ttu-id="109b9-143">对于任务窗格和内容加载项，`Office.initialize` 提供了其他 _reason_ 参数。</span><span class="sxs-lookup"><span data-stu-id="109b9-143">For task pane and content add-ins, `Office.initialize` provides an additional _reason_ parameter.</span></span> <span data-ttu-id="109b9-144">此参数指定如何将加载项添加到当前文档。</span><span class="sxs-lookup"><span data-stu-id="109b9-144">This parameter specifies how an add-in was added to the current document.</span></span> <span data-ttu-id="109b9-145">可以使用此参数针对首次插入加载项时和加载项已存在于文档中时实施不同的逻辑。</span><span class="sxs-lookup"><span data-stu-id="109b9-145">You can use this to provide different logic for when an add-in is first inserted versus when it already existed within the document.</span></span>

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

<span data-ttu-id="109b9-146">有关详细信息，请参阅 [Office.initialize 事件](/javascript/api/office)和 [InitializationReason 枚举](/javascript/api/office/office.initializationreason)。</span><span class="sxs-lookup"><span data-stu-id="109b9-146">For more information, see [Office.initialize Event](/javascript/api/office) and [InitializationReason Enumeration](/javascript/api/office/office.initializationreason).</span></span>

## <a name="major-differences-between-officeinitialize-and-officeonready"></a><span data-ttu-id="109b9-147">Office.initialize 和 Office.onReady 之间的主要差别</span><span class="sxs-lookup"><span data-stu-id="109b9-147">Major differences between Office.initialize and Office.onReady</span></span>

- <span data-ttu-id="109b9-148">可以仅将一个处理程序分配到 `Office.initialize` 并仅由 Office 基础结构调用一次，但可以在代码中的不同位置调用 `Office.onReady()` 并使用不同的回调。</span><span class="sxs-lookup"><span data-stu-id="109b9-148">You can assign only one handler to `Office.initialize` and it's called only once by the Office infrastructure; but you can call `Office.onReady()` in different places in your code and use different callbacks.</span></span> <span data-ttu-id="109b9-149">例如，只要自定义脚本使用运行初始化逻辑的回调进行加载，代码就可以调用 `Office.onReady()`。代码还可以在任务窗格中设置一个按钮，其脚本会使用不同的回调调用 `Office.onReady()`。</span><span class="sxs-lookup"><span data-stu-id="109b9-149">For example, your code could call `Office.onReady()` as soon as your custom script loads with a callback that runs initialization logic; and your code could also have a button in the task pane, whose script calls `Office.onReady()` with a different callback.</span></span> <span data-ttu-id="109b9-150">如果是这样，则会在单击该按钮后运行第二个回调。</span><span class="sxs-lookup"><span data-stu-id="109b9-150">If so, the second callback runs when the button is clicked.</span></span>

- <span data-ttu-id="109b9-151">`Office.initialize` 事件将在 Office.js 初始化其本身的内部过程的末尾处触发。</span><span class="sxs-lookup"><span data-stu-id="109b9-151">The `Office.initialize` event fires at the end of the internal process in which Office.js initializes itself.</span></span> <span data-ttu-id="109b9-152">并且它会在内部过程结束后*立即*触发。</span><span class="sxs-lookup"><span data-stu-id="109b9-152">And it fires *immediately* after the internal process ends.</span></span> <span data-ttu-id="109b9-153">如果将处理程序分配到事件所使用的代码在事件触发后执行的时间过长，则处理程序将不会运行。</span><span class="sxs-lookup"><span data-stu-id="109b9-153">If the code in which you assign a handler to the event executes too long after the event fires, then your handler doesn't run.</span></span> <span data-ttu-id="109b9-154">例如，如果使用的是 WebPack 任务管理器，则在加载 Office.js 后但在加载自定义 JavaScript 前，它会配置加载项的主页以加载填充代码文件。</span><span class="sxs-lookup"><span data-stu-id="109b9-154">For example, if you are using the WebPack task manager, it might configure the add-in's home page to load polyfill files after it loads Office.js but before it loads your custom JavaScript.</span></span> <span data-ttu-id="109b9-155">在脚本加载和分配处理程序时，初始化事件已经发生。</span><span class="sxs-lookup"><span data-stu-id="109b9-155">By the time your script loads and assigns the handler, the initialize event has already happened.</span></span> <span data-ttu-id="109b9-156">但调用 `Office.onReady()` 永远不会“太迟”。</span><span class="sxs-lookup"><span data-stu-id="109b9-156">But it is never "too late" to call `Office.onReady()`.</span></span> <span data-ttu-id="109b9-157">如果初始化事件已经发生，则回调将立即运行。</span><span class="sxs-lookup"><span data-stu-id="109b9-157">If the initialize event has already happened, the callback runs immediately.</span></span>

> [!NOTE]
> <span data-ttu-id="109b9-158">即使没有启动逻辑，也应在加载项 JavaScript 加载时调用 `Office.onReady()` 或将空函数分配到 `Office.initialize`。</span><span class="sxs-lookup"><span data-stu-id="109b9-158">Even if you have no start-up logic, you should either call `Office.onReady()` or assign an empty function to `Office.initialize` when your add-in JavaScript loads.</span></span> <span data-ttu-id="109b9-159">某些 Office 应用程序和平台组合不会加载任务窗格，除非发生其中一种情况。</span><span class="sxs-lookup"><span data-stu-id="109b9-159">Some Office application and platform combinations won't load the task pane until one of these happens.</span></span> <span data-ttu-id="109b9-160">以下示例显示了这两种方法。</span><span class="sxs-lookup"><span data-stu-id="109b9-160">The following examples show these two approaches.</span></span>
>
>```js    
>Office.onReady();
>```
>
>
>```js
>Office.initialize = function () {};
>```

## <a name="see-also"></a><span data-ttu-id="109b9-161">另请参阅</span><span class="sxs-lookup"><span data-stu-id="109b9-161">See also</span></span>

- [<span data-ttu-id="109b9-162">了解 Office JavaScript API</span><span class="sxs-lookup"><span data-stu-id="109b9-162">Understanding the Office JavaScript API</span></span>](understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="109b9-163">加载 DOM 和运行时环境</span><span class="sxs-lookup"><span data-stu-id="109b9-163">Loading the DOM and runtime environment</span></span>](loading-the-dom-and-runtime-environment.md)