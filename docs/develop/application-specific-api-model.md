---
title: 使用特定于应用程序的 API 模型
description: 了解 Excel、OneNote 和 Word 外接程序的基于承诺的 API 模型。
ms.date: 07/29/2020
localization_priority: Normal
ms.openlocfilehash: cabd1ea0076b672a1dbda3079a767b0e8a1a62b7
ms.sourcegitcommit: 4adfc368a366f00c3f3d7ed387f34aaecb47f17c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/01/2020
ms.locfileid: "47326280"
---
# <a name="using-the-application-specific-api-model"></a><span data-ttu-id="65c8d-103">使用特定于应用程序的 API 模型</span><span class="sxs-lookup"><span data-stu-id="65c8d-103">Using the application-specific API model</span></span>

<span data-ttu-id="65c8d-104">本文介绍如何使用 API 模型在 Excel、Word 和 OneNote 中构建外接程序。</span><span class="sxs-lookup"><span data-stu-id="65c8d-104">This article describes how to use the API model for building add-ins in Excel, Word, and OneNote.</span></span> <span data-ttu-id="65c8d-105">它介绍了使用基于承诺的 Api 的基础的核心概念。</span><span class="sxs-lookup"><span data-stu-id="65c8d-105">It introduces core concepts that are fundamental to using the promise-based APIs.</span></span>

> [!NOTE]
> <span data-ttu-id="65c8d-106">Office 2013 客户端不支持此模型。</span><span class="sxs-lookup"><span data-stu-id="65c8d-106">This model is not supported by Office 2013 clients.</span></span> <span data-ttu-id="65c8d-107">使用 [通用 API 模型](office-javascript-api-object-model.md) 来处理这些 Office 版本。</span><span class="sxs-lookup"><span data-stu-id="65c8d-107">Use the [Common API model](office-javascript-api-object-model.md) to work with those Office versions.</span></span> <span data-ttu-id="65c8d-108">有关完整的平台可用性说明，请参阅 [适用于 office 的 Office 外接程序的 office 客户端应用程序和平台可用性](../overview/office-add-in-availability.md)。</span><span class="sxs-lookup"><span data-stu-id="65c8d-108">For full platform availability notes, see [Office client application and platform availability for Office Add-ins](../overview/office-add-in-availability.md).</span></span>

> [!TIP]
> <span data-ttu-id="65c8d-109">此页面中的示例使用 Excel JavaScript Api，但这些概念也适用于 OneNote、Visio 和 Word JavaScript Api。</span><span class="sxs-lookup"><span data-stu-id="65c8d-109">The examples in this page use the Excel JavaScript APIs, but the concepts also apply to OneNote, Visio, and Word JavaScript APIs.</span></span>

## <a name="asynchronous-nature-of-the-promise-based-apis"></a><span data-ttu-id="65c8d-110">基于承诺的 Api 的异步特性</span><span class="sxs-lookup"><span data-stu-id="65c8d-110">Asynchronous nature of the promise-based APIs</span></span>

<span data-ttu-id="65c8d-111">Office 外接程序是在 Office 应用程序（如 Excel）内的浏览器容器中显示的网站。</span><span class="sxs-lookup"><span data-stu-id="65c8d-111">Office Add-ins are websites which appear inside a browser container within Office applications, such as Excel.</span></span> <span data-ttu-id="65c8d-112">此容器嵌入在基于桌面的平台（如 Windows 上的 Office）中的 Office 应用程序中，并在 web 上的 Office 中的 HTML iFrame 内运行。</span><span class="sxs-lookup"><span data-stu-id="65c8d-112">This container is embedded within the Office application on desktop-based platforms, such as Office on Windows, and runs inside an HTML iFrame in Office on the web.</span></span> <span data-ttu-id="65c8d-113">由于性能方面的考虑，Office.js Api 无法跨所有平台与 Office 应用程序同步交互。</span><span class="sxs-lookup"><span data-stu-id="65c8d-113">Due to performance considerations, the Office.js APIs cannot interact synchronously with the Office applications across all platforms.</span></span> <span data-ttu-id="65c8d-114">因此， `sync()` Office.js 中的 API 调用返回在 Office 应用程序完成请求的读取或写入操作时解决的 [承诺](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) 。</span><span class="sxs-lookup"><span data-stu-id="65c8d-114">Therefore, the `sync()` API call in Office.js returns a [Promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) that is resolved when the Office application completes the requested read or write actions.</span></span> <span data-ttu-id="65c8d-115">此外，还可以对多个操作（如设置属性或调用方法）进行排队，并将它们作为一批命令运行 `sync()` ，而不是为每个操作发送单独的请求。</span><span class="sxs-lookup"><span data-stu-id="65c8d-115">Also, you can queue up multiple actions, such as setting properties or invoking methods, and run them as a batch of commands with a single call to `sync()`, rather than sending a separate request for each action.</span></span> <span data-ttu-id="65c8d-116">以下各节介绍如何使用和 api 来完成此操作 `run()` `sync()` 。</span><span class="sxs-lookup"><span data-stu-id="65c8d-116">The following sections describe how to accomplish this using the `run()` and `sync()` APIs.</span></span>

## <a name="run-function"></a><span data-ttu-id="65c8d-117">\*. run 函数</span><span class="sxs-lookup"><span data-stu-id="65c8d-117">\*.run function</span></span>

<span data-ttu-id="65c8d-118">`Excel.run`、 `Word.run` 和 `OneNote.run` 执行一个函数，该函数指定要对 Excel、Word 和 OneNote 执行的操作。</span><span class="sxs-lookup"><span data-stu-id="65c8d-118">`Excel.run`, `Word.run`, and `OneNote.run` execute a function that specifies the actions to perform against Excel, Word, and OneNote.</span></span> <span data-ttu-id="65c8d-119">`*.run` 自动创建可用于与 Office 对象进行交互的请求上下文。</span><span class="sxs-lookup"><span data-stu-id="65c8d-119">`*.run` automatically creates a request context that you can use to interact with Office objects.</span></span> <span data-ttu-id="65c8d-120">`*.run`完成后，将会解决承诺，并且会自动释放在运行时分配的任何对象。</span><span class="sxs-lookup"><span data-stu-id="65c8d-120">When `*.run` completes, a promise is resolved, and any objects that were allocated at runtime are automatically released.</span></span>

<span data-ttu-id="65c8d-121">下面的示例演示如何使用 `Excel.run` 。</span><span class="sxs-lookup"><span data-stu-id="65c8d-121">The following example shows how to use `Excel.run`.</span></span> <span data-ttu-id="65c8d-122">Word 和 OneNote 也使用相同的模式。</span><span class="sxs-lookup"><span data-stu-id="65c8d-122">The same pattern is also used with Word and OneNote.</span></span>

```js
Excel.run(function (context) {
    // Add your Excel JS API calls here that will be batched and sent to the workbook.
    console.log('Your code goes here.');
}).catch(function (error) {
    // Catch and log any errors that occur within `Excel.run`.
    console.log('error: ' + error);
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

## <a name="request-context"></a><span data-ttu-id="65c8d-123">请求上下文</span><span class="sxs-lookup"><span data-stu-id="65c8d-123">Request context</span></span>

<span data-ttu-id="65c8d-124">Office 应用程序和外接程序在两个不同的进程中运行。</span><span class="sxs-lookup"><span data-stu-id="65c8d-124">The Office application and your add-in run in two different processes.</span></span> <span data-ttu-id="65c8d-125">由于它们使用不同的运行时环境，因此外接程序需要对象才能将 `RequestContext` 外接程序连接到 Office 中的对象，如工作表、区域、段落和表。</span><span class="sxs-lookup"><span data-stu-id="65c8d-125">Since they use different runtime environments, add-ins require a `RequestContext` object in order to connect your add-in to objects in Office such as worksheets, ranges, paragraphs, and tables.</span></span> <span data-ttu-id="65c8d-126">`RequestContext`调用时，此对象作为参数提供 `*.run` 。</span><span class="sxs-lookup"><span data-stu-id="65c8d-126">This `RequestContext` object is provided as an argument when calling `*.run`.</span></span>

## <a name="proxy-objects"></a><span data-ttu-id="65c8d-127">代理对象</span><span class="sxs-lookup"><span data-stu-id="65c8d-127">Proxy objects</span></span>

<span data-ttu-id="65c8d-128">您声明并与基于承诺的 Api 一起使用的 Office JavaScript 对象是代理对象。</span><span class="sxs-lookup"><span data-stu-id="65c8d-128">The Office JavaScript objects that you declare and use with the promise-based APIs are proxy objects.</span></span> <span data-ttu-id="65c8d-129">调用的任何方法或在代理对象上设置或加载的属性都只是添加到挂起命令的队列中。</span><span class="sxs-lookup"><span data-stu-id="65c8d-129">Any methods that you invoke or properties that you set or load on proxy objects are simply added to a queue of pending commands.</span></span> <span data-ttu-id="65c8d-130">在 `sync()` 请求上下文上调用方法时 (例如， `context.sync()`) ，队列命令将被调度到 Office 应用程序并运行。</span><span class="sxs-lookup"><span data-stu-id="65c8d-130">When you call the `sync()` method on the request context (for example, `context.sync()`), the queued commands are dispatched to the Office application and run.</span></span> <span data-ttu-id="65c8d-131">这些 Api 从根本上以批处理为中心。</span><span class="sxs-lookup"><span data-stu-id="65c8d-131">These APIs are fundamentally batch-centric.</span></span> <span data-ttu-id="65c8d-132">您可以根据需要在请求上下文中排列任意数量的更改，然后调用 `sync()` 方法以运行队列中的命令批次。</span><span class="sxs-lookup"><span data-stu-id="65c8d-132">You can queue up as many changes as you wish on the request context, and then call the `sync()` method to run the batch of queued commands.</span></span>

<span data-ttu-id="65c8d-133">例如，以下代码段声明了本地 JavaScript [Excel Range](/javascript/api/excel/excel.range) 对象， `selectedRange` 以引用 Excel 工作簿中的选定区域，然后设置该对象的一些属性。</span><span class="sxs-lookup"><span data-stu-id="65c8d-133">For example, the following code snippet declares the local JavaScript [Excel.Range](/javascript/api/excel/excel.range) object, `selectedRange`, to reference a selected range in the Excel workbook, and then sets some properties on that object.</span></span> <span data-ttu-id="65c8d-134">该 `selectedRange` 对象是一个代理对象，因此在您的外接程序调用之前，设置的属性和在该对象上调用的方法将不会反映在 Excel 文档中 `context.sync()` 。</span><span class="sxs-lookup"><span data-stu-id="65c8d-134">The `selectedRange` object is a proxy object, so the properties that are set and the method that is invoked on that object will not be reflected in the Excel document until your add-in calls `context.sync()`.</span></span>

```js
var selectedRange = context.workbook.getSelectedRange();
selectedRange.format.fill.color = "#4472C4";
selectedRange.format.font.color = "white";
selectedRange.format.autofitColumns();
```

### <a name="performance-tip-minimize-the-number-of-proxy-objects-created"></a><span data-ttu-id="65c8d-135">性能提示：最大限度地减少创建的代理对象数</span><span class="sxs-lookup"><span data-stu-id="65c8d-135">Performance tip: Minimize the number of proxy objects created</span></span>

<span data-ttu-id="65c8d-136">避免重复创建同一个代理对象。</span><span class="sxs-lookup"><span data-stu-id="65c8d-136">Avoid repeatedly creating the same proxy object.</span></span> <span data-ttu-id="65c8d-137">如果多个操作需要同一个代理对象，则改为创建一次并将其分配给一个变量，然后在代码中使用该变量。</span><span class="sxs-lookup"><span data-stu-id="65c8d-137">Instead, if you need the same proxy object for more than one operation, create it once and assign it to a variable, then use that variable in your code.</span></span>

```js
// BAD: Repeated calls to .getRange() to create the same proxy object.
worksheet.getRange("A1").format.fill.color = "red";
worksheet.getRange("A1").numberFormat = "0.00%";
worksheet.getRange("A1").values = [[1]];

// GOOD: Create the range proxy object once and assign to a variable.
var range = worksheet.getRange("A1")
range.format.fill.color = "red";
range.numberFormat = "0.00%";
range.values = [[1]];

// ALSO GOOD: Use a "set" method to immediately set all the properties without even needing to create a variable!
worksheet.getRange("A1").set({
    numberFormat: [["0.00%"]],
    values: [[1]],
    format: {
        fill: {
            color: "red"
        }
    }
});
```

### <a name="sync"></a><span data-ttu-id="65c8d-138">sync()</span><span class="sxs-lookup"><span data-stu-id="65c8d-138">sync()</span></span>

<span data-ttu-id="65c8d-139">`sync()`对请求上下文调用方法将同步 Office 文档中的代理对象和对象之间的状态。</span><span class="sxs-lookup"><span data-stu-id="65c8d-139">Calling the `sync()` method on the request context synchronizes the state between proxy objects and objects in the Office document.</span></span> <span data-ttu-id="65c8d-140">该 `sync()` 方法运行在请求上下文中排队的任何命令，并检索应在代理对象上加载的任何属性的值。</span><span class="sxs-lookup"><span data-stu-id="65c8d-140">The `sync()` method runs any commands that are queued on the request context and retrieves values for any properties that should be loaded on the proxy objects.</span></span> <span data-ttu-id="65c8d-141">`sync()`方法以异步方式执行，并返回一个[承诺](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)，该方法在 `sync()` 方法完成时得到解决。</span><span class="sxs-lookup"><span data-stu-id="65c8d-141">The `sync()` method executes asynchronously and returns a [Promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise), which is resolved when the `sync()` method completes.</span></span>

<span data-ttu-id="65c8d-142">下面的示例演示了一个批处理函数，该函数定义本地 JavaScript 代理对象 (`selectedRange`) ，加载该对象的属性，然后使用 JavaScript 承诺模式来调用， `context.sync()` 以同步 Excel 文档中的代理对象和对象之间的状态。</span><span class="sxs-lookup"><span data-stu-id="65c8d-142">The following example shows a batch function that defines a local JavaScript proxy object (`selectedRange`), loads a property of that object, and then uses the JavaScript promises pattern to call `context.sync()` to synchronize the state between proxy objects and objects in the Excel document.</span></span>

```js
Excel.run(function (context) {
    var selectedRange = context.workbook.getSelectedRange();
    selectedRange.load('address');
    return context.sync()
      .then(function () {
        console.log('The selected range is: ' + selectedRange.address);
    });
}).catch(function (error) {
    console.log('error: ' + error);
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

<span data-ttu-id="65c8d-143">在上一示例中，已设置 `selectedRange`，并且将在调用 `context.sync()` 时加载其 `address` 属性。</span><span class="sxs-lookup"><span data-stu-id="65c8d-143">In the previous example, `selectedRange` is set and its `address` property is loaded when `context.sync()` is called.</span></span>

<span data-ttu-id="65c8d-144">由于 `sync()` 是异步操作，因此应始终返回 `Promise` 对象以确保 `sync()` 操作在脚本继续运行之前完成。</span><span class="sxs-lookup"><span data-stu-id="65c8d-144">Since `sync()` is an asynchronous operation, you should always return the `Promise` object to ensure the `sync()` operation completes before the script continues to run.</span></span> <span data-ttu-id="65c8d-145">如果使用的是 TypeScript 或 ES6 + JavaScript，则可以 `await` `context.sync()` 调用，而不是返回承诺。</span><span class="sxs-lookup"><span data-stu-id="65c8d-145">If you're using TypeScript or ES6+ JavaScript, you can `await` the `context.sync()` call instead of returning the promise.</span></span>

#### <a name="performance-tip-minimize-the-number-of-sync-calls"></a><span data-ttu-id="65c8d-146">性能提示：最大限度地减少同步调用数</span><span class="sxs-lookup"><span data-stu-id="65c8d-146">Performance tip: Minimize the number of sync calls</span></span>

<span data-ttu-id="65c8d-147">在 Excel JavaScript API 中，`sync()` 是唯一的异步操作，在某些情况下可能会很慢，尤其是对于 Excel 网页版。</span><span class="sxs-lookup"><span data-stu-id="65c8d-147">In the Excel JavaScript API, `sync()` is the only asynchronous operation, and it can be slow under some circumstances, especially for Excel on the web.</span></span> <span data-ttu-id="65c8d-148">若要优化性能，在调用之前，通过尽可能多地将更改加入队列来最大程度减少调用 `sync()` 的次数。</span><span class="sxs-lookup"><span data-stu-id="65c8d-148">To optimize performance, minimize the number of calls to `sync()` by queueing up as many changes as possible before calling it.</span></span> <span data-ttu-id="65c8d-149">有关优化性能的详细信息 `sync()` ，请参阅 [避免在循环中使用 context. sync 方法](../concepts/correlated-objects-pattern.md)。</span><span class="sxs-lookup"><span data-stu-id="65c8d-149">For more information about optimizing performance with `sync()`, see [Avoid using the context.sync method in loops](../concepts/correlated-objects-pattern.md).</span></span>

### <a name="load"></a><span data-ttu-id="65c8d-150">load()</span><span class="sxs-lookup"><span data-stu-id="65c8d-150">load()</span></span>

<span data-ttu-id="65c8d-151">在可以读取代理对象的属性之前，必须显式加载属性以使用 Office 文档中的数据填充代理对象，然后再调用 `context.sync()` 。</span><span class="sxs-lookup"><span data-stu-id="65c8d-151">Before you can read the properties of a proxy object, you must explicitly load the properties to populate the proxy object with data from the Office document, and then call `context.sync()`.</span></span> <span data-ttu-id="65c8d-152">例如，如果创建代理对象以引用选定区域，然后想要读取选定区域的 `address` 属性，则需要先加载该属性，然后才能 `address` 阅读该属性。</span><span class="sxs-lookup"><span data-stu-id="65c8d-152">For example, if you create a proxy object to reference a selected range, and then want to read the selected range's `address` property, you need to load the `address` property before you can read it.</span></span> <span data-ttu-id="65c8d-153">若要加载代理对象的属性，请对 `load()` 该对象调用方法，并指定要加载的属性。</span><span class="sxs-lookup"><span data-stu-id="65c8d-153">To request properties of a proxy object be loaded, call the `load()` method on the object and specify the properties to load.</span></span> <span data-ttu-id="65c8d-154">下面的示例展示了 `Range.address` 要加载的属性 `myRange` 。</span><span class="sxs-lookup"><span data-stu-id="65c8d-154">The following example shows the `Range.address` property being loaded for `myRange`.</span></span>

```js
Excel.run(function (context) {
    var sheetName = 'Sheet1';
    var rangeAddress = 'A1:B2';
    var myRange = context.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);

    myRange.load('address');

    return context.sync()
      .then(function () {
        console.log (myRange.address);   // ok
        //console.log (myRange.values);  // not ok as it was not loaded
        });
    }).then(function () {
        console.log('done');
}).catch(function (error) {
    console.log('Error: ' + error);
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

> [!NOTE]
> <span data-ttu-id="65c8d-155">如果只调用方法或设置代理对象的属性，则不需要调用该 `load()` 方法。</span><span class="sxs-lookup"><span data-stu-id="65c8d-155">If you are only calling methods or setting properties on a proxy object, you don't need to call the `load()` method.</span></span> <span data-ttu-id="65c8d-156">`load()`仅当您想要读取代理对象的属性时，才需要使用此方法。</span><span class="sxs-lookup"><span data-stu-id="65c8d-156">The `load()` method is only required when you want to read properties on a proxy object.</span></span>

<span data-ttu-id="65c8d-p115">类似于对代理对象设置属性或调用方法的请求，加载代理对象属性的请求会被添加到请求上下文的挂起命令队列中，将在下一次调用 `sync()` 方法时运行。必要时，可以将请求上下文中尽可能多的 `load()` 调用排入队列。</span><span class="sxs-lookup"><span data-stu-id="65c8d-p115">Just like requests to set properties or invoke methods on proxy objects, requests to load properties on proxy objects get added to the queue of pending commands on the request context, which will run the next time you call the `sync()` method. You can queue up as many `load()` calls on the request context as necessary.</span></span>

#### <a name="scalar-and-navigation-properties"></a><span data-ttu-id="65c8d-159">标量和导航属性</span><span class="sxs-lookup"><span data-stu-id="65c8d-159">Scalar and navigation properties</span></span>

<span data-ttu-id="65c8d-160">属性分为两种类别：**标量**和**导航**。</span><span class="sxs-lookup"><span data-stu-id="65c8d-160">There are two categories of properties: **scalar** and **navigational**.</span></span> <span data-ttu-id="65c8d-161">标量属性是可分配的类型，如字符串、整数和 JSON 结构。</span><span class="sxs-lookup"><span data-stu-id="65c8d-161">Scalar properties are assignable types such as strings, integers, and JSON structs.</span></span> <span data-ttu-id="65c8d-162">导航属性是只读对象和已分配字段的对象集合，而不是直接分配属性。</span><span class="sxs-lookup"><span data-stu-id="65c8d-162">Navigation properties are read-only objects and collections of objects that have their fields assigned, instead of directly assigning the property.</span></span> <span data-ttu-id="65c8d-163">例如， `name` 和的 `position` 成员在 [Excel 中。工作表](/javascript/api/excel/excel.worksheet) 对象是标量属性，而 `protection` 并 `tables` 是导航属性。</span><span class="sxs-lookup"><span data-stu-id="65c8d-163">For example, `name` and `position` members on the [Excel.Worksheet](/javascript/api/excel/excel.worksheet) object are scalar properties, whereas `protection` and `tables` are navigation properties.</span></span>

<span data-ttu-id="65c8d-164">您的外接程序可以将导航属性用作加载特定标量属性的路径。</span><span class="sxs-lookup"><span data-stu-id="65c8d-164">Your add-in can use navigational properties as a path to load specific scalar properties.</span></span> <span data-ttu-id="65c8d-165">下面的代码 `load` 对对象使用的字体名称的命令进行排队 `Excel.Range` ，而不加载任何其他信息。</span><span class="sxs-lookup"><span data-stu-id="65c8d-165">The following code queues up a `load` command for the name of the font used by an `Excel.Range` object, without loading any other information.</span></span>

```js
someRange.load("format/font/name")
```

<span data-ttu-id="65c8d-166">您还可以通过遍历路径来设置导航属性的标量属性。</span><span class="sxs-lookup"><span data-stu-id="65c8d-166">You can also set the scalar properties of a navigation property by traversing the path.</span></span> <span data-ttu-id="65c8d-167">例如，可以使用设置的字体大小 `Excel.Range` `someRange.format.font.size = 10;` 。</span><span class="sxs-lookup"><span data-stu-id="65c8d-167">For example, you could set the font size for an `Excel.Range` by using `someRange.format.font.size = 10;`.</span></span> <span data-ttu-id="65c8d-168">在设置属性之前，不需要加载该属性。</span><span class="sxs-lookup"><span data-stu-id="65c8d-168">You don't need to load the property before you set it.</span></span>

<span data-ttu-id="65c8d-169">请注意，一个对象下的一些属性可能与另一个对象的名称相同。</span><span class="sxs-lookup"><span data-stu-id="65c8d-169">Please be aware that some of the properties under an object may have the same name as another object.</span></span> <span data-ttu-id="65c8d-170">例如， `format` 是对象下的属性 `Excel.Range` ，但本身也 `format` 是对象。</span><span class="sxs-lookup"><span data-stu-id="65c8d-170">For example, `format` is a property under the `Excel.Range` object, but `format` itself is an object as well.</span></span> <span data-ttu-id="65c8d-171">因此，如果发出类似的调用，则 `range.load("format")` 等效于 `range.format.load()` (不需要的空 `load()` 语句) 。</span><span class="sxs-lookup"><span data-stu-id="65c8d-171">So, if you make a call such as `range.load("format")`, this is equivalent to `range.format.load()` (an undesirable empty `load()` statement).</span></span> <span data-ttu-id="65c8d-172">若要避免这种情况，代码应仅加载对象树中的 "叶节点"。</span><span class="sxs-lookup"><span data-stu-id="65c8d-172">To avoid this, your code should only load the "leaf nodes" in an object tree.</span></span>

#### <a name="calling-load-without-parameters-not-recommended"></a><span data-ttu-id="65c8d-173">`load`不建议调用不带参数的 () </span><span class="sxs-lookup"><span data-stu-id="65c8d-173">Calling `load` without parameters (not recommended)</span></span>

<span data-ttu-id="65c8d-174">如果在 `load()` 未指定任何参数的情况下对对象 (或集合) 调用方法，则将加载该对象的所有标量属性或该集合的对象。</span><span class="sxs-lookup"><span data-stu-id="65c8d-174">If you call the `load()` method on an object (or collection) without specifying any parameters, all scalar properties of the object or the collection's objects will be loaded.</span></span> <span data-ttu-id="65c8d-175">加载不需要的数据会降低外接程序的速度。</span><span class="sxs-lookup"><span data-stu-id="65c8d-175">Loading unneeded data will slow down your add-in.</span></span> <span data-ttu-id="65c8d-176">应始终显式指定要加载的属性。</span><span class="sxs-lookup"><span data-stu-id="65c8d-176">You should always explicitly specify which properties to load.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="65c8d-177">无参数 `load` 语句返回的数据量可能超过该服务的大小限制。</span><span class="sxs-lookup"><span data-stu-id="65c8d-177">The amount of data returned by a parameter-less `load` statement can exceed the size limits of the service.</span></span> <span data-ttu-id="65c8d-178">为了降低较旧加载项的风险，`load` 不会在明确请求它们之前返回某些属性。</span><span class="sxs-lookup"><span data-stu-id="65c8d-178">To reduce the risks to older add-ins, some properties are not returned by `load` without explicitly requesting them.</span></span> <span data-ttu-id="65c8d-179">此类加载操作中排除了以下属性：</span><span class="sxs-lookup"><span data-stu-id="65c8d-179">The following properties are excluded from such load operations:</span></span>
>
> * `Excel.Range.numberFormatCategories`

### <a name="clientresult"></a><span data-ttu-id="65c8d-180">ClientResult</span><span class="sxs-lookup"><span data-stu-id="65c8d-180">ClientResult</span></span>

<span data-ttu-id="65c8d-181">返回基元类型的基于承诺的 api 中的方法具有与范例类似的模式 `load` / `sync` 。</span><span class="sxs-lookup"><span data-stu-id="65c8d-181">Methods in the promise-based APIs that return primitive types have a similar pattern to the `load`/`sync` paradigm.</span></span> <span data-ttu-id="65c8d-182">举个例子，`Excel.TableCollection.getCount`获取集合中的表的数量。</span><span class="sxs-lookup"><span data-stu-id="65c8d-182">As an example, `Excel.TableCollection.getCount` gets the number of tables in the collection.</span></span> <span data-ttu-id="65c8d-183">`getCount` 返回 a `ClientResult<number>` ，表示 `value` 返回的属性 [`ClientResult`](/javascript/api/office/officeextension.clientresult) 为数字。</span><span class="sxs-lookup"><span data-stu-id="65c8d-183">`getCount` returns a `ClientResult<number>`, meaning the `value` property in the returned [`ClientResult`](/javascript/api/office/officeextension.clientresult) is a number.</span></span> <span data-ttu-id="65c8d-184">在调用 `context.sync()` 之前，脚本无法访问此值。</span><span class="sxs-lookup"><span data-stu-id="65c8d-184">Your script can't access that value until `context.sync()` is called.</span></span>

<span data-ttu-id="65c8d-185">下面的代码获取 Excel 工作簿中的总表数，并将该数目的日志记录到控制台。</span><span class="sxs-lookup"><span data-stu-id="65c8d-185">The following code gets the total number of tables in an Excel workbook and logs that number to the console.</span></span>

```js
var tableCount = context.workbook.tables.getCount();

// This sync call implicitly loads tableCount.value.
// Any other ClientResult values are loaded too.
return context.sync()
    .then(function () {
        // Trying to log the value before calling sync would throw an error.
        console.log (tableCount.value);
    });
```

### <a name="set"></a><span data-ttu-id="65c8d-186">set()</span><span class="sxs-lookup"><span data-stu-id="65c8d-186">set()</span></span>

<span data-ttu-id="65c8d-187">在具有嵌套导航属性的对象上设置属性可能很麻烦。</span><span class="sxs-lookup"><span data-stu-id="65c8d-187">Setting properties on an object with nested navigation properties can be cumbersome.</span></span> <span data-ttu-id="65c8d-188">除了以上所述使用导航路径设置各个属性之外，您还可以使用 `object.set()` 基于承诺的 JavaScript api 中的对象上提供的方法。</span><span class="sxs-lookup"><span data-stu-id="65c8d-188">As an alternative to setting individual properties using navigation paths as described above, you can use the `object.set()` method that is available on objects in the promise-based JavaScript APIs.</span></span> <span data-ttu-id="65c8d-189">使用此方法，可以通过传递相同 Office.js 类型的另一个对象或 JavaScript 对象（其属性结构类似于调用该方法的对象的属性）一次设置对象的多个属性。</span><span class="sxs-lookup"><span data-stu-id="65c8d-189">With this method, you can set multiple properties of an object at once by passing either another object of the same Office.js type or a JavaScript object with properties that are structured like the properties of the object on which the method is called.</span></span>

<span data-ttu-id="65c8d-p124">下面的代码示例设置区域的多个格式属性，具体方法是调用 `set()` 方法，并传入 JavaScript 对象，其中包含可反映 `Range` 对象中属性结构的属性名称和类型。此示例假定区域 **B2:E2** 中有数据。</span><span class="sxs-lookup"><span data-stu-id="65c8d-p124">The following code sample sets several format properties of a range by calling the `set()` method and passing in a JavaScript object with property names and types that mirror the structure of properties in the `Range` object. This example assumes that there is data in range **B2:E2**.</span></span>

```js
Excel.run(function (ctx) {
    var sheet = ctx.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:E2");
    range.set({
        format: {
            fill: {
                color: '#4472C4'
            },
            font: {
                name: 'Verdana',
                color: 'white'
            }
        }
    });
    range.format.autofitColumns();

    return ctx.sync();
}).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

## <a name="42ornullobject-methods-and-properties"></a><span data-ttu-id="65c8d-192">&#42;OrNullObject 方法和属性</span><span class="sxs-lookup"><span data-stu-id="65c8d-192">&#42;OrNullObject methods and properties</span></span>

<span data-ttu-id="65c8d-193">当所需的对象不存在时，某些访问器方法和属性将引发异常。</span><span class="sxs-lookup"><span data-stu-id="65c8d-193">Some accessor methods and properties throw an exception when the desired object doesn't exist.</span></span> <span data-ttu-id="65c8d-194">例如，如果尝试通过指定不在工作簿中的工作表名称来获取 Excel 工作表，则该 `getItem()` 方法将引发 `ItemNotFound` 异常。</span><span class="sxs-lookup"><span data-stu-id="65c8d-194">For example, if you attempt to get an Excel worksheet by specifying a worksheet name that isn't in the workbook, the `getItem()` method throws an `ItemNotFound` exception.</span></span> <span data-ttu-id="65c8d-195">特定于应用程序的库为代码提供了一种测试文档实体是否存在的方法，而不需要异常处理代码。</span><span class="sxs-lookup"><span data-stu-id="65c8d-195">The application-specific libraries provide a way for your code to test for the existence of document entities without requiring exception handling code.</span></span> <span data-ttu-id="65c8d-196">这是通过使用 `*OrNullObject` 方法和属性的变体来实现的。</span><span class="sxs-lookup"><span data-stu-id="65c8d-196">This is accomplished by using the `*OrNullObject` variations of methods and properties.</span></span> <span data-ttu-id="65c8d-197">`isNullObject` `true` 如果指定的项不存在，而不是引发异常，则这些变体返回其属性设置为的对象。</span><span class="sxs-lookup"><span data-stu-id="65c8d-197">These variations return an object whose `isNullObject` property is set to `true`, if the specified item doesn't exist, rather than throwing an exception.</span></span>

<span data-ttu-id="65c8d-198">例如，可以对 `getItemOrNullObject()` 集合（如 **工作表** ）调用方法，以从集合中检索项。</span><span class="sxs-lookup"><span data-stu-id="65c8d-198">For example, you can call the `getItemOrNullObject()` method on a collection such as **Worksheets** to retrieve an item from the collection.</span></span> <span data-ttu-id="65c8d-199">`getItemOrNullObject()`如果指定的项存在，则该方法将返回它; 否则，将返回其 `isNullObject` 属性设置为的对象 `true` 。</span><span class="sxs-lookup"><span data-stu-id="65c8d-199">The `getItemOrNullObject()` method returns the specified item if it exists; otherwise, it returns an object whose `isNullObject` property is set to `true`.</span></span> <span data-ttu-id="65c8d-200">然后，您的代码可以对此属性进行评估，以确定该对象是否存在。</span><span class="sxs-lookup"><span data-stu-id="65c8d-200">Your code can then evaluate this property to determine whether the object exists.</span></span>

> [!NOTE]
> <span data-ttu-id="65c8d-201">`*OrNullObject`变体从不返回 JavaScript 值 `null` 。</span><span class="sxs-lookup"><span data-stu-id="65c8d-201">The `*OrNullObject` variations do not ever return the JavaScript value `null`.</span></span> <span data-ttu-id="65c8d-202">它们返回普通的 Office 代理对象。</span><span class="sxs-lookup"><span data-stu-id="65c8d-202">They return ordinary Office proxy objects.</span></span> <span data-ttu-id="65c8d-203">如果该对象所代表的实体不存在，则 `isNullObject` 将该对象的属性设置为 `true` 。</span><span class="sxs-lookup"><span data-stu-id="65c8d-203">If the the entity that the object represents does not exist then the `isNullObject` property of the object is set to `true`.</span></span> <span data-ttu-id="65c8d-204">请勿为 null 或 falsity 测试返回的对象。</span><span class="sxs-lookup"><span data-stu-id="65c8d-204">Do not test the returned object for nullity or falsity.</span></span> <span data-ttu-id="65c8d-205">它永远不会是、 `null` `false` 或 `undefined` 。</span><span class="sxs-lookup"><span data-stu-id="65c8d-205">It is never `null`, `false`, or `undefined`.</span></span>

<span data-ttu-id="65c8d-206">下面的代码示例尝试使用方法检索名为 "Data" 的 Excel 工作表 `getItemOrNullObject()` 。</span><span class="sxs-lookup"><span data-stu-id="65c8d-206">The following code sample attempts to retrieve an Excel worksheet named "Data" by using the `getItemOrNullObject()` method.</span></span> <span data-ttu-id="65c8d-207">如果具有该名称的工作表不存在，则创建一个新工作表。</span><span class="sxs-lookup"><span data-stu-id="65c8d-207">If a worksheet with that name does not exist, a new sheet is created.</span></span> <span data-ttu-id="65c8d-208">请注意，该代码不会加载该 `isNullObject` 属性。</span><span class="sxs-lookup"><span data-stu-id="65c8d-208">Note that the code does not load the `isNullObject` property.</span></span> <span data-ttu-id="65c8d-209">Office 将在调用时自动加载此属性 `context.sync` ，因此无需使用类似的内容显式加载它 `datasheet.load('isNullObject')` 。</span><span class="sxs-lookup"><span data-stu-id="65c8d-209">Office automatically loads this property when `context.sync` is called, so you do not need to explicitly load it with something like `datasheet.load('isNullObject')`.</span></span>

```js
var dataSheet = context.workbook.worksheets.getItemOrNullObject("Data");

return context.sync()
    .then(function () {
        if (dataSheet.isNullObject) {
            dataSheet = context.workbook.worksheets.add("Data");
        }

        // Set `dataSheet` to be the second worksheet in the workbook.
        dataSheet.position = 1;
    });
```

## <a name="see-also"></a><span data-ttu-id="65c8d-210">另请参阅</span><span class="sxs-lookup"><span data-stu-id="65c8d-210">See also</span></span>

* [<span data-ttu-id="65c8d-211">常见 JavaScript API 对象模型</span><span class="sxs-lookup"><span data-stu-id="65c8d-211">Common JavaScript API object model</span></span>](office-javascript-api-object-model.md)
* <span data-ttu-id="65c8d-212">[常见的编码问题和意外的平台行为](common-coding-issues.md)。</span><span class="sxs-lookup"><span data-stu-id="65c8d-212">[Common coding issues and unexpected platform behaviors](common-coding-issues.md).</span></span>
* [<span data-ttu-id="65c8d-213">Office 外接程序的资源限制和性能优化</span><span class="sxs-lookup"><span data-stu-id="65c8d-213">Resource limits and performance optimization for Office Add-ins</span></span>](../concepts/resource-limits-and-performance-optimization.md)
