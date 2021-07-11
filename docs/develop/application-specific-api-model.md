---
title: 使用应用程序专用 API 模型
description: 了解 Excel、OneNote 和 Word 加载项基于承诺的 API 模型。
ms.date: 09/08/2020
localization_priority: Normal
ms.openlocfilehash: 5cf1d088dfa883e5df9eaba25e395857cfce9f5c
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/09/2021
ms.locfileid: "53350062"
---
# <a name="using-the-application-specific-api-model"></a><span data-ttu-id="938ca-103">使用应用程序专用 API 模型</span><span class="sxs-lookup"><span data-stu-id="938ca-103">Using the application-specific API model</span></span>

<span data-ttu-id="938ca-104">本文介绍如何使用 API 模型在 Excel、Word 和 OneNote 中构建加载项。</span><span class="sxs-lookup"><span data-stu-id="938ca-104">This article describes how to use the API model for building add-ins in Excel, Word, and OneNote.</span></span> <span data-ttu-id="938ca-105">本文介绍核心概念，这些概念是使用基于承诺的 API 的基础。</span><span class="sxs-lookup"><span data-stu-id="938ca-105">It introduces core concepts that are fundamental to using the promise-based APIs.</span></span>

> [!NOTE]
> <span data-ttu-id="938ca-106">Office 2013 客户端不支持此模型。</span><span class="sxs-lookup"><span data-stu-id="938ca-106">This model is not supported by Office 2013 clients.</span></span> <span data-ttu-id="938ca-107">使用 [API 模型](office-javascript-api-object-model.md) 这些 Office 版本。</span><span class="sxs-lookup"><span data-stu-id="938ca-107">Use the [Common API model](office-javascript-api-object-model.md) to work with those Office versions.</span></span> <span data-ttu-id="938ca-108">有关完整的平台可用性说明，请参阅 [Office 客户端应用程序和平台可用性的 Office 加载项组](../overview/office-add-in-availability.md)。</span><span class="sxs-lookup"><span data-stu-id="938ca-108">For full platform availability notes, see [Office client application and platform availability for Office Add-ins](../overview/office-add-in-availability.md).</span></span>

> [!TIP]
> <span data-ttu-id="938ca-109">本页中的示例使用 Excel JavaScript API，但概念也适用于 OneNote、Visio 和 Word JavaScript API。</span><span class="sxs-lookup"><span data-stu-id="938ca-109">The examples in this page use the Excel JavaScript APIs, but the concepts also apply to OneNote, Visio, and Word JavaScript APIs.</span></span>

## <a name="asynchronous-nature-of-the-promise-based-apis"></a><span data-ttu-id="938ca-110">基于承诺的 API 的异步性质</span><span class="sxs-lookup"><span data-stu-id="938ca-110">Asynchronous nature of the promise-based APIs</span></span>

<span data-ttu-id="938ca-111">Office 加载项是显示在 Office 应用程序（如 Excel）中的浏览器容器内的网站。</span><span class="sxs-lookup"><span data-stu-id="938ca-111">Office Add-ins are websites which appear inside a browser container within Office applications, such as Excel.</span></span> <span data-ttu-id="938ca-112">此容器嵌入在基于桌面的 Office 应用程序（如 Windows 上的 Office）上的 Office 应用程序中，在 Office 网页版中的 HTML iFrame 内运行。</span><span class="sxs-lookup"><span data-stu-id="938ca-112">This container is embedded within the Office application on desktop-based platforms, such as Office on Windows, and runs inside an HTML iFrame in Office on the web.</span></span> <span data-ttu-id="938ca-113">出于性能方面的考虑，Office.js API 无法跨所有平台与 Office 应用程序同步交互。</span><span class="sxs-lookup"><span data-stu-id="938ca-113">Due to performance considerations, the Office.js APIs cannot interact synchronously with the Office applications across all platforms.</span></span> <span data-ttu-id="938ca-114">因此，Office.js `sync()` API 调用返回 [Office](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) 请求的读取或写入操作时要解决的"承诺"问题。</span><span class="sxs-lookup"><span data-stu-id="938ca-114">Therefore, the `sync()` API call in Office.js returns a [Promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) that is resolved when the Office application completes the requested read or write actions.</span></span> <span data-ttu-id="938ca-115">此外，还可以将多个操作（例如设置属性或调用方法）排队，并且只要一次呼叫 `sync()`，就可以作为一批命令运行这些操作，而不是针对每个操作发送单独的请求。</span><span class="sxs-lookup"><span data-stu-id="938ca-115">Also, you can queue up multiple actions, such as setting properties or invoking methods, and run them as a batch of commands with a single call to `sync()`, rather than sending a separate request for each action.</span></span> <span data-ttu-id="938ca-116">以下各节介绍如何使用 API 和 `run()``sync()`实现此操作。</span><span class="sxs-lookup"><span data-stu-id="938ca-116">The following sections describe how to accomplish this using the `run()` and `sync()` APIs.</span></span>

## <a name="run-function"></a><span data-ttu-id="938ca-117">\*.run 函数</span><span class="sxs-lookup"><span data-stu-id="938ca-117">\*.run function</span></span>

<span data-ttu-id="938ca-118">`Excel.run`、 `Word.run`和 `OneNote.run` 执行一个函数，指定针对 Excel、Word 和 OneNote 要执行的操作。</span><span class="sxs-lookup"><span data-stu-id="938ca-118">`Excel.run`, `Word.run`, and `OneNote.run` execute a function that specifies the actions to perform against Excel, Word, and OneNote.</span></span> <span data-ttu-id="938ca-119">`*.run` 会自动创建可用于与 Office 对象交互的请求上下文。</span><span class="sxs-lookup"><span data-stu-id="938ca-119">`*.run` automatically creates a request context that you can use to interact with Office objects.</span></span> <span data-ttu-id="938ca-120">当 `*.run` ，将做出承诺，并自动发布运行时分配的任何对象。</span><span class="sxs-lookup"><span data-stu-id="938ca-120">When `*.run` completes, a promise is resolved, and any objects that were allocated at runtime are automatically released.</span></span>

<span data-ttu-id="938ca-121">以下示例显示了如何使用 `Excel.run`。</span><span class="sxs-lookup"><span data-stu-id="938ca-121">The following example shows how to use `Excel.run`.</span></span> <span data-ttu-id="938ca-122">Word 和 OneNote 也使用同一模式。</span><span class="sxs-lookup"><span data-stu-id="938ca-122">The same pattern is also used with Word and OneNote.</span></span>

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

## <a name="request-context"></a><span data-ttu-id="938ca-123">请求上下文</span><span class="sxs-lookup"><span data-stu-id="938ca-123">Request context</span></span>

<span data-ttu-id="938ca-124">Office 应用程序和加载项在两个不同过程中运行。</span><span class="sxs-lookup"><span data-stu-id="938ca-124">The Office application and your add-in run in two different processes.</span></span> <span data-ttu-id="938ca-125">由于加载项使用不同的运行时环境，因此需要一个 `RequestContext` 对象才能将加载项连接到 Office 中的对象，例如工作表、区域、段落和表。</span><span class="sxs-lookup"><span data-stu-id="938ca-125">Since they use different runtime environments, add-ins require a `RequestContext` object in order to connect your add-in to objects in Office such as worksheets, ranges, paragraphs, and tables.</span></span> <span data-ttu-id="938ca-126">调用 `RequestContext` 时，此对象作为 `*.run`。</span><span class="sxs-lookup"><span data-stu-id="938ca-126">This `RequestContext` object is provided as an argument when calling `*.run`.</span></span>

## <a name="proxy-objects"></a><span data-ttu-id="938ca-127">代理对象</span><span class="sxs-lookup"><span data-stu-id="938ca-127">Proxy objects</span></span>

<span data-ttu-id="938ca-128">声明并用于基于承诺的 API 的 Office JavaScript 对象是代理对象。</span><span class="sxs-lookup"><span data-stu-id="938ca-128">The Office JavaScript objects that you declare and use with the promise-based APIs are proxy objects.</span></span> <span data-ttu-id="938ca-129">调用的任何方法或在代理对象上设置或加载的属性都只是添加到挂起命令的队列中。</span><span class="sxs-lookup"><span data-stu-id="938ca-129">Any methods that you invoke or properties that you set or load on proxy objects are simply added to a queue of pending commands.</span></span> <span data-ttu-id="938ca-130">调用请求 `sync()` （例如 `context.sync()`）上的方法时，排队的命令将调用 Office 应用程序并运行。</span><span class="sxs-lookup"><span data-stu-id="938ca-130">When you call the `sync()` method on the request context (for example, `context.sync()`), the queued commands are dispatched to the Office application and run.</span></span> <span data-ttu-id="938ca-131">这些 API 在根本上以批处理为中心。</span><span class="sxs-lookup"><span data-stu-id="938ca-131">These APIs are fundamentally batch-centric.</span></span> <span data-ttu-id="938ca-132">您可以根据对请求上下文希望排入多达数个更改，然后调用 `sync()` 方法，以运行排队命令的批处理。</span><span class="sxs-lookup"><span data-stu-id="938ca-132">You can queue up as many changes as you wish on the request context, and then call the `sync()` method to run the batch of queued commands.</span></span>

<span data-ttu-id="938ca-133">例如，以下代码片段声明本地 JavaScript [Excel.Range](/javascript/api/excel/excel.range) 对象 `selectedRange`引用 Excel 工作簿中的选定区域，然后针对该对象设置一些属性。</span><span class="sxs-lookup"><span data-stu-id="938ca-133">For example, the following code snippet declares the local JavaScript [Excel.Range](/javascript/api/excel/excel.range) object, `selectedRange`, to reference a selected range in the Excel workbook, and then sets some properties on that object.</span></span> <span data-ttu-id="938ca-134">对象 `selectedRange` 代理对象，因此在调用加载项之前，不会在 Excel 文档中反映已设置的属性和在该对象上调用 `context.sync()`。</span><span class="sxs-lookup"><span data-stu-id="938ca-134">The `selectedRange` object is a proxy object, so the properties that are set and the method that is invoked on that object will not be reflected in the Excel document until your add-in calls `context.sync()`.</span></span>

```js
var selectedRange = context.workbook.getSelectedRange();
selectedRange.format.fill.color = "#4472C4";
selectedRange.format.font.color = "white";
selectedRange.format.autofitColumns();
```

### <a name="performance-tip-minimize-the-number-of-proxy-objects-created"></a><span data-ttu-id="938ca-135">性能提示：最小化创建代理对象的数量</span><span class="sxs-lookup"><span data-stu-id="938ca-135">Performance tip: Minimize the number of proxy objects created</span></span>

<span data-ttu-id="938ca-136">避免重复创建同一个代理对象。</span><span class="sxs-lookup"><span data-stu-id="938ca-136">Avoid repeatedly creating the same proxy object.</span></span> <span data-ttu-id="938ca-137">如果多个操作需要同一个代理对象，则改为创建一次并将其分配给一个变量，然后在代码中使用该变量。</span><span class="sxs-lookup"><span data-stu-id="938ca-137">Instead, if you need the same proxy object for more than one operation, create it once and assign it to a variable, then use that variable in your code.</span></span>

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

### <a name="sync"></a><span data-ttu-id="938ca-138">sync()</span><span class="sxs-lookup"><span data-stu-id="938ca-138">sync()</span></span>

<span data-ttu-id="938ca-139">调用 `sync()` 上下文的方法可同步 Office 文档中代理对象和对象之间的状态。</span><span class="sxs-lookup"><span data-stu-id="938ca-139">Calling the `sync()` method on the request context synchronizes the state between proxy objects and objects in the Office document.</span></span> <span data-ttu-id="938ca-140">该 `sync()` 在请求上下文中排入队列的任何命令，并检索应在代理对象上加载的任何属性的值。</span><span class="sxs-lookup"><span data-stu-id="938ca-140">The `sync()` method runs any commands that are queued on the request context and retrieves values for any properties that should be loaded on the proxy objects.</span></span> <span data-ttu-id="938ca-141">方法 `sync()` 异步执行，并返回 [承诺](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)，该 `sync()` 完成时完成。</span><span class="sxs-lookup"><span data-stu-id="938ca-141">The `sync()` method executes asynchronously and returns a [Promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise), which is resolved when the `sync()` method completes.</span></span>

<span data-ttu-id="938ca-142">下面的示例显示一个批处理函数，定义本地 JavaScript 代理对象 （`selectedRange`），加载该对象的属性，然后使用 JavaScript 形式调用 `context.sync()` 以在 Excel 文档中的代理对象和对象间同步状态。</span><span class="sxs-lookup"><span data-stu-id="938ca-142">The following example shows a batch function that defines a local JavaScript proxy object (`selectedRange`), loads a property of that object, and then uses the JavaScript promises pattern to call `context.sync()` to synchronize the state between proxy objects and objects in the Excel document.</span></span>

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

<span data-ttu-id="938ca-143">在上一示例中，已设置 `selectedRange`，并且将在调用 `context.sync()` 时加载其 `address` 属性。</span><span class="sxs-lookup"><span data-stu-id="938ca-143">In the previous example, `selectedRange` is set and its `address` property is loaded when `context.sync()` is called.</span></span>

<span data-ttu-id="938ca-144">由于 `sync()` 是异步操作，因此在脚本继续运行之前，应始终 `Promise` 同步对象，以确保 `sync()` 操作完成。</span><span class="sxs-lookup"><span data-stu-id="938ca-144">Since `sync()` is an asynchronous operation, you should always return the `Promise` object to ensure the `sync()` operation completes before the script continues to run.</span></span> <span data-ttu-id="938ca-145">如果使用 TypeScript 或 ES6+ JavaScript， `await` 调用 `context.sync()` ，而不是返回承诺。</span><span class="sxs-lookup"><span data-stu-id="938ca-145">If you're using TypeScript or ES6+ JavaScript, you can `await` the `context.sync()` call instead of returning the promise.</span></span>

#### <a name="performance-tip-minimize-the-number-of-sync-calls"></a><span data-ttu-id="938ca-146">性能提示：减少同步呼叫数</span><span class="sxs-lookup"><span data-stu-id="938ca-146">Performance tip: Minimize the number of sync calls</span></span>

<span data-ttu-id="938ca-147">在 Excel JavaScript API 中，`sync()` 是唯一的异步操作，在某些情况下可能会很慢，尤其是对于 Excel 网页版。</span><span class="sxs-lookup"><span data-stu-id="938ca-147">In the Excel JavaScript API, `sync()` is the only asynchronous operation, and it can be slow under some circumstances, especially for Excel on the web.</span></span> <span data-ttu-id="938ca-148">若要优化性能，在调用之前，通过尽可能多地将更改加入队列来最大程度减少调用 `sync()` 的次数。</span><span class="sxs-lookup"><span data-stu-id="938ca-148">To optimize performance, minimize the number of calls to `sync()` by queueing up as many changes as possible before calling it.</span></span> <span data-ttu-id="938ca-149">有关使用 `sync()`优化性能，请参阅 [循环使用 context.sync 方法](../concepts/correlated-objects-pattern.md)。</span><span class="sxs-lookup"><span data-stu-id="938ca-149">For more information about optimizing performance with `sync()`, see [Avoid using the context.sync method in loops](../concepts/correlated-objects-pattern.md).</span></span>

### <a name="load"></a><span data-ttu-id="938ca-150">load()</span><span class="sxs-lookup"><span data-stu-id="938ca-150">load()</span></span>

<span data-ttu-id="938ca-151">必须显式加载属性才能读取代理对象的属性，才能使用 Office 文档中的数据填充代理对象，然后调用 `context.sync()`。</span><span class="sxs-lookup"><span data-stu-id="938ca-151">Before you can read the properties of a proxy object, you must explicitly load the properties to populate the proxy object with data from the Office document, and then call `context.sync()`.</span></span> <span data-ttu-id="938ca-152">例如，如果创建代理对象来引用选定的区域，然后希望读取所选区域的 `address` 属性，需要首先加载 `address` 属性，然后才可以读取它。</span><span class="sxs-lookup"><span data-stu-id="938ca-152">For example, if you create a proxy object to reference a selected range, and then want to read the selected range's `address` property, you need to load the `address` property before you can read it.</span></span> <span data-ttu-id="938ca-153">若要请求加载代理对象的属性，调用 `load()` 的方法并指定要加载的属性。</span><span class="sxs-lookup"><span data-stu-id="938ca-153">To request properties of a proxy object be loaded, call the `load()` method on the object and specify the properties to load.</span></span> <span data-ttu-id="938ca-154">以下示例显示了为 `Range.address` 加载的 `myRange`。</span><span class="sxs-lookup"><span data-stu-id="938ca-154">The following example shows the `Range.address` property being loaded for `myRange`.</span></span>

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
> <span data-ttu-id="938ca-155">如果只是调用代理对象或设置属性，则无需调用代理 `load()` 方法。</span><span class="sxs-lookup"><span data-stu-id="938ca-155">If you are only calling methods or setting properties on a proxy object, you don't need to call the `load()` method.</span></span> <span data-ttu-id="938ca-156">只有在想要读取代理对象上的属性时 `load()` 代理方法才必需。</span><span class="sxs-lookup"><span data-stu-id="938ca-156">The `load()` method is only required when you want to read properties on a proxy object.</span></span>

<span data-ttu-id="938ca-p115">类似于对代理对象设置属性或调用方法的请求，加载代理对象属性的请求会被添加到请求上下文的挂起命令队列中，将在下一次调用 `sync()` 方法时运行。必要时，可以将请求上下文中尽可能多的 `load()` 调用排入队列。</span><span class="sxs-lookup"><span data-stu-id="938ca-p115">Just like requests to set properties or invoke methods on proxy objects, requests to load properties on proxy objects get added to the queue of pending commands on the request context, which will run the next time you call the `sync()` method. You can queue up as many `load()` calls on the request context as necessary.</span></span>

#### <a name="scalar-and-navigation-properties"></a><span data-ttu-id="938ca-159">标量和导航属性</span><span class="sxs-lookup"><span data-stu-id="938ca-159">Scalar and navigation properties</span></span>

<span data-ttu-id="938ca-160">属性分为两种类别：**标量** 和 **导航**。</span><span class="sxs-lookup"><span data-stu-id="938ca-160">There are two categories of properties: **scalar** and **navigational**.</span></span> <span data-ttu-id="938ca-161">标量属性是可分配的类型，如字符串、整数和 JSON 结构。</span><span class="sxs-lookup"><span data-stu-id="938ca-161">Scalar properties are assignable types such as strings, integers, and JSON structs.</span></span> <span data-ttu-id="938ca-162">导航属性是分配了字段的只读对象和对象集合，而不是直接分配属性。</span><span class="sxs-lookup"><span data-stu-id="938ca-162">Navigation properties are read-only objects and collections of objects that have their fields assigned, instead of directly assigning the property.</span></span> <span data-ttu-id="938ca-163">例如，[Worksheet](/javascript/api/excel/excel.worksheet) 对象上的 `name`和 `position` 成员是标量属性，而 `protection` 和 `tables` 是导航属性。</span><span class="sxs-lookup"><span data-stu-id="938ca-163">For example, `name` and `position` members on the [Excel.Worksheet](/javascript/api/excel/excel.worksheet) object are scalar properties, whereas `protection` and `tables` are navigation properties.</span></span>

<span data-ttu-id="938ca-164">加载项可使用导航属性作为加载特定标量属性的路径。</span><span class="sxs-lookup"><span data-stu-id="938ca-164">Your add-in can use navigational properties as a path to load specific scalar properties.</span></span> <span data-ttu-id="938ca-165">以下代码会按照 `load` 对象使用的字体的名称将向上排队 `Excel.Range` 命令，而无需加载任何其他信息。</span><span class="sxs-lookup"><span data-stu-id="938ca-165">The following code queues up a `load` command for the name of the font used by an `Excel.Range` object, without loading any other information.</span></span>

```js
someRange.load("format/font/name")
```

<span data-ttu-id="938ca-166">还可通过遍历路径来设置导航属性的标量属性。</span><span class="sxs-lookup"><span data-stu-id="938ca-166">You can also set the scalar properties of a navigation property by traversing the path.</span></span> <span data-ttu-id="938ca-167">例如，通过使用"另一种" `Excel.Range` ， `someRange.format.font.size = 10;`。</span><span class="sxs-lookup"><span data-stu-id="938ca-167">For example, you could set the font size for an `Excel.Range` by using `someRange.format.font.size = 10;`.</span></span> <span data-ttu-id="938ca-168">设置属性前无需加载属性。</span><span class="sxs-lookup"><span data-stu-id="938ca-168">You don't need to load the property before you set it.</span></span>

<span data-ttu-id="938ca-169">请注意，一个对象下的某些“属性”可能与另一个对象同名。</span><span class="sxs-lookup"><span data-stu-id="938ca-169">Please be aware that some of the properties under an object may have the same name as another object.</span></span> <span data-ttu-id="938ca-170">例如， `format` 是对象下 `Excel.Range` 属性， `format` 值本身也是一个对象。</span><span class="sxs-lookup"><span data-stu-id="938ca-170">For example, `format` is a property under the `Excel.Range` object, but `format` itself is an object as well.</span></span> <span data-ttu-id="938ca-171">因此，如果你进行 `range.load("format")`等呼叫，这相当于 `range.format.load()` （一个空的 `load()` 语句）。</span><span class="sxs-lookup"><span data-stu-id="938ca-171">So, if you make a call such as `range.load("format")`, this is equivalent to `range.format.load()` (an undesirable empty `load()` statement).</span></span> <span data-ttu-id="938ca-172">若要避免这种情况，代码应仅加载对象树中的“叶节点”。</span><span class="sxs-lookup"><span data-stu-id="938ca-172">To avoid this, your code should only load the "leaf nodes" in an object tree.</span></span>

#### <a name="calling-load-without-parameters-not-recommended"></a><span data-ttu-id="938ca-173">不带 `load` （不推荐）的呼叫方</span><span class="sxs-lookup"><span data-stu-id="938ca-173">Calling `load` without parameters (not recommended)</span></span>

<span data-ttu-id="938ca-174">如果在不 `load()` 参数的情况下调用对象（或集合）上的标量方法，将加载该对象或集合对象的所有标量属性。</span><span class="sxs-lookup"><span data-stu-id="938ca-174">If you call the `load()` method on an object (or collection) without specifying any parameters, all scalar properties of the object or the collection's objects will be loaded.</span></span> <span data-ttu-id="938ca-175">加载不需要的数据会降低加载项的加载速度。</span><span class="sxs-lookup"><span data-stu-id="938ca-175">Loading unneeded data will slow down your add-in.</span></span> <span data-ttu-id="938ca-176">应始终显式指定要加载的属性。</span><span class="sxs-lookup"><span data-stu-id="938ca-176">You should always explicitly specify which properties to load.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="938ca-177">无参数 `load` 语句返回的数据量可能超过该服务的大小限制。</span><span class="sxs-lookup"><span data-stu-id="938ca-177">The amount of data returned by a parameter-less `load` statement can exceed the size limits of the service.</span></span> <span data-ttu-id="938ca-178">为了降低较旧加载项的风险，`load` 不会在明确请求它们之前返回某些属性。</span><span class="sxs-lookup"><span data-stu-id="938ca-178">To reduce the risks to older add-ins, some properties are not returned by `load` without explicitly requesting them.</span></span> <span data-ttu-id="938ca-179">以下属性从此类加载操作中排除。</span><span class="sxs-lookup"><span data-stu-id="938ca-179">The following properties are excluded from such load operations.</span></span>
>
> * `Excel.Range.numberFormatCategories`

### <a name="clientresult"></a><span data-ttu-id="938ca-180">ClientResult</span><span class="sxs-lookup"><span data-stu-id="938ca-180">ClientResult</span></span>

<span data-ttu-id="938ca-181">基于承诺的 API 中返回类型 API 的方法与现代方法`load`/`sync`模式。</span><span class="sxs-lookup"><span data-stu-id="938ca-181">Methods in the promise-based APIs that return primitive types have a similar pattern to the `load`/`sync` paradigm.</span></span> <span data-ttu-id="938ca-182">举个例子，`Excel.TableCollection.getCount`获取集合中的表的数量。</span><span class="sxs-lookup"><span data-stu-id="938ca-182">As an example, `Excel.TableCollection.getCount` gets the number of tables in the collection.</span></span> <span data-ttu-id="938ca-183">`getCount` 返回一`ClientResult<number>`，这意味着返回的`value`中的 [`ClientResult`](/javascript/api/office/officeextension.clientresult) 属性是一个数字。</span><span class="sxs-lookup"><span data-stu-id="938ca-183">`getCount` returns a `ClientResult<number>`, meaning the `value` property in the returned [`ClientResult`](/javascript/api/office/officeextension.clientresult) is a number.</span></span> <span data-ttu-id="938ca-184">在调用 `context.sync()` 之前，脚本无法访问此值。</span><span class="sxs-lookup"><span data-stu-id="938ca-184">Your script can't access that value until `context.sync()` is called.</span></span>

<span data-ttu-id="938ca-185">以下代码获取 Excel 工作簿中的表总数，并对此数目的日志记录到控制台。</span><span class="sxs-lookup"><span data-stu-id="938ca-185">The following code gets the total number of tables in an Excel workbook and logs that number to the console.</span></span>

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

### <a name="set"></a><span data-ttu-id="938ca-186">set()</span><span class="sxs-lookup"><span data-stu-id="938ca-186">set()</span></span>

<span data-ttu-id="938ca-187">在具有嵌套导航属性的对象上设置属性可能很麻烦。</span><span class="sxs-lookup"><span data-stu-id="938ca-187">Setting properties on an object with nested navigation properties can be cumbersome.</span></span> <span data-ttu-id="938ca-188">除了使用上述导航路径设置单个属性， `object.set()` 基于承诺的 JavaScript API 中的对象上可用的另一种方法。</span><span class="sxs-lookup"><span data-stu-id="938ca-188">As an alternative to setting individual properties using navigation paths as described above, you can use the `object.set()` method that is available on objects in the promise-based JavaScript APIs.</span></span> <span data-ttu-id="938ca-189">使用此方法，可以通过传递相同 Office.js 类型的另一个对象或 JavaScript 对象（其属性结构类似于调用该方法的对象的属性）一次设置对象的多个属性。</span><span class="sxs-lookup"><span data-stu-id="938ca-189">With this method, you can set multiple properties of an object at once by passing either another object of the same Office.js type or a JavaScript object with properties that are structured like the properties of the object on which the method is called.</span></span>

<span data-ttu-id="938ca-p124">下面的代码示例设置区域的多个格式属性，具体方法是调用 `set()` 方法，并传入 JavaScript 对象，其中包含可反映 `Range` 对象中属性结构的属性名称和类型。此示例假定区域 **B2:E2** 中有数据。</span><span class="sxs-lookup"><span data-stu-id="938ca-p124">The following code sample sets several format properties of a range by calling the `set()` method and passing in a JavaScript object with property names and types that mirror the structure of properties in the `Range` object. This example assumes that there is data in range **B2:E2**.</span></span>

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

### <a name="some-properties-cannot-be-set-directly"></a><span data-ttu-id="938ca-192">某些属性不能直接设置</span><span class="sxs-lookup"><span data-stu-id="938ca-192">Some properties cannot be set directly</span></span>

<span data-ttu-id="938ca-193">尽管可写的属性，但某些属性不能设置。</span><span class="sxs-lookup"><span data-stu-id="938ca-193">Some properties cannot be set, despite being writable.</span></span> <span data-ttu-id="938ca-194">这些属性是必须将设置为单个对象的父属性的一部分。</span><span class="sxs-lookup"><span data-stu-id="938ca-194">These properties are part of a parent property that must be set as a single object.</span></span> <span data-ttu-id="938ca-195">这是因为父属性依赖于具有特定逻辑关系的子属性。</span><span class="sxs-lookup"><span data-stu-id="938ca-195">This is because that parent property relies on the subproperties having specific, logical relationships.</span></span> <span data-ttu-id="938ca-196">必须使用对象文字表示法设置这些父属性来设置整个对象，而不是设置该对象的单个子问题。</span><span class="sxs-lookup"><span data-stu-id="938ca-196">These parent properties must be set using object literal notation to set the entire object, instead of setting that object's individual subproperties.</span></span> <span data-ttu-id="938ca-197">PageLayout [中可找到此示例](/javascript/api/excel/excel.pagelayout)。</span><span class="sxs-lookup"><span data-stu-id="938ca-197">One example of this is found in [PageLayout](/javascript/api/excel/excel.pagelayout).</span></span> <span data-ttu-id="938ca-198">必须 `zoom` 单个 PageLayoutZoomOptions [每个对象设置](/javascript/api/excel/excel.pagelayoutzoomoptions) ，如下所示：</span><span class="sxs-lookup"><span data-stu-id="938ca-198">The `zoom` property must be set with a single [PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions) object, as shown here:</span></span>

```js
// PageLayout.zoom.scale must be set by assigning PageLayout.zoom to a PageLayoutZoomOptions object.
sheet.pageLayout.zoom = { scale: 200 };
```

<span data-ttu-id="938ca-199">在上一示例中，***无法*** 直接分配为值`zoom`：`sheet.pageLayout.zoom.scale = 200;`。</span><span class="sxs-lookup"><span data-stu-id="938ca-199">In the previous example, you would ***not*** be able to directly assign `zoom` a value: `sheet.pageLayout.zoom.scale = 200;`.</span></span> <span data-ttu-id="938ca-200">该语句会引发错误， `zoom` 加载错误。</span><span class="sxs-lookup"><span data-stu-id="938ca-200">That statement throws an error because `zoom` is not loaded.</span></span> <span data-ttu-id="938ca-201">即使 `zoom` ，该比例也会生效。</span><span class="sxs-lookup"><span data-stu-id="938ca-201">Even if `zoom` were to be loaded, the set of scale will not take effect.</span></span> <span data-ttu-id="938ca-202">所有上下文操作 `zoom`、刷新加载项中的代理对象并覆盖本地设置的值。</span><span class="sxs-lookup"><span data-stu-id="938ca-202">All context operations happen on `zoom`, refreshing the proxy object in the add-in and overwriting locally set values.</span></span>

<span data-ttu-id="938ca-203">此行为与 [Range.format](application-specific-api-model.md#scalar-and-navigation-properties) 等 [导航属性](/javascript/api/excel/excel.range#format)。</span><span class="sxs-lookup"><span data-stu-id="938ca-203">This behavior differs from [navigational properties](application-specific-api-model.md#scalar-and-navigation-properties) like [Range.format](/javascript/api/excel/excel.range#format).</span></span> <span data-ttu-id="938ca-204">您可以使用对象 `format` 设置对象属性，如下所示：</span><span class="sxs-lookup"><span data-stu-id="938ca-204">Properties of `format` can be set using object navigation, as shown here:</span></span>

```js
// This will set the font size on the range during the next `content.sync()`.
range.format.font.size = 10;
```

<span data-ttu-id="938ca-205">可通过检查其只读修改者，识别不能直接设置其子问题的属性。</span><span class="sxs-lookup"><span data-stu-id="938ca-205">You can identify a property that cannot have its subproperties directly set by checking its read-only modifier.</span></span> <span data-ttu-id="938ca-206">所有只读属性可直接设置其非只读子问题。</span><span class="sxs-lookup"><span data-stu-id="938ca-206">All read-only properties can have their non-read-only subproperties directly set.</span></span> <span data-ttu-id="938ca-207">必须在该级别 `PageLayout.zoom` 可编写属性，如属性。</span><span class="sxs-lookup"><span data-stu-id="938ca-207">Writeable properties like `PageLayout.zoom` must be set with an object at that level.</span></span> <span data-ttu-id="938ca-208">摘要：</span><span class="sxs-lookup"><span data-stu-id="938ca-208">In summary:</span></span>

- <span data-ttu-id="938ca-209">只读属性：可通过导航设置子项目。</span><span class="sxs-lookup"><span data-stu-id="938ca-209">Read-only property: Subproperties can be set through navigation.</span></span>
- <span data-ttu-id="938ca-210">可写的属性：无法通过导航设置子项目（必须设置为初始父对象分配的一部分）。</span><span class="sxs-lookup"><span data-stu-id="938ca-210">Writable property: Subproperties cannot be set through navigation (must be set as part of the initial parent object assignment).</span></span>



## <a name="42ornullobject-methods-and-properties"></a><span data-ttu-id="938ca-211">&#42;OrNullObject 方法与属性</span><span class="sxs-lookup"><span data-stu-id="938ca-211">&#42;OrNullObject methods and properties</span></span>

<span data-ttu-id="938ca-212">当所需对象不存在时，某些配件方法和属性将引发异常。</span><span class="sxs-lookup"><span data-stu-id="938ca-212">Some accessor methods and properties throw an exception when the desired object doesn't exist.</span></span> <span data-ttu-id="938ca-213">例如，如果尝试通过指定工作簿未包含的工作表名称获取 Excel 工作表，则 `getItem()` 会引发 `ItemNotFound` 异常。</span><span class="sxs-lookup"><span data-stu-id="938ca-213">For example, if you attempt to get an Excel worksheet by specifying a worksheet name that isn't in the workbook, the `getItem()` method throws an `ItemNotFound` exception.</span></span> <span data-ttu-id="938ca-214">特定于应用程序的库为代码提供了一种方法，用于测试文档实体是否存在，而无需异常处理代码。</span><span class="sxs-lookup"><span data-stu-id="938ca-214">The application-specific libraries provide a way for your code to test for the existence of document entities without requiring exception handling code.</span></span> <span data-ttu-id="938ca-215">此操作是通过使用多种 `*OrNullObject` 和属性实现的。</span><span class="sxs-lookup"><span data-stu-id="938ca-215">This is accomplished by using the `*OrNullObject` variations of methods and properties.</span></span> <span data-ttu-id="938ca-216">如果指定的项目不存在，这些变体将返回其 `isNullObject` 被设置为 `true`值的对象，而不是引发异常。</span><span class="sxs-lookup"><span data-stu-id="938ca-216">These variations return an object whose `isNullObject` property is set to `true`, if the specified item doesn't exist, rather than throwing an exception.</span></span>

<span data-ttu-id="938ca-217">例如，可以在集合（如 **Worksheets**）上调用 `getItemOrNullObject()` 方法，尝试从集合中检索某个项。</span><span class="sxs-lookup"><span data-stu-id="938ca-217">For example, you can call the `getItemOrNullObject()` method on a collection such as **Worksheets** to retrieve an item from the collection.</span></span> <span data-ttu-id="938ca-218">方法 `getItemOrNullObject()` 返回指定项目（如果存在）;否则，将返回其属性 `isNullObject` 为 <a0/ `true`。</span><span class="sxs-lookup"><span data-stu-id="938ca-218">The `getItemOrNullObject()` method returns the specified item if it exists; otherwise, it returns an object whose `isNullObject` property is set to `true`.</span></span> <span data-ttu-id="938ca-219">然后，代码可评估此属性，以确定该对象是否存在。</span><span class="sxs-lookup"><span data-stu-id="938ca-219">Your code can then evaluate this property to determine whether the object exists.</span></span>

> [!NOTE]
> <span data-ttu-id="938ca-220">这些 `*OrNullObject` 变体永远不会返回值 JavaScript `null`。</span><span class="sxs-lookup"><span data-stu-id="938ca-220">The `*OrNullObject` variations do not ever return the JavaScript value `null`.</span></span> <span data-ttu-id="938ca-221">它们返回普通 Office 代理对象。</span><span class="sxs-lookup"><span data-stu-id="938ca-221">They return ordinary Office proxy objects.</span></span> <span data-ttu-id="938ca-222">如果对象表示的实体不存在，则对象的 `isNullObject` 属性设置为 `true`。</span><span class="sxs-lookup"><span data-stu-id="938ca-222">If the the entity that the object represents does not exist then the `isNullObject` property of the object is set to `true`.</span></span> <span data-ttu-id="938ca-223">不要测试返回的对象为 nullity 或 fality。</span><span class="sxs-lookup"><span data-stu-id="938ca-223">Do not test the returned object for nullity or falsity.</span></span> <span data-ttu-id="938ca-224">它从 `null`、 `false`或 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="938ca-224">It is never `null`, `false`, or `undefined`.</span></span>

<span data-ttu-id="938ca-225">以下代码示例尝试使用以下方法检索名为"数据"的 Excel `getItemOrNullObject()`。</span><span class="sxs-lookup"><span data-stu-id="938ca-225">The following code sample attempts to retrieve an Excel worksheet named "Data" by using the `getItemOrNullObject()` method.</span></span> <span data-ttu-id="938ca-226">如果具有该名称的工作表不存在，将创建一个新工作表。</span><span class="sxs-lookup"><span data-stu-id="938ca-226">If a worksheet with that name does not exist, a new sheet is created.</span></span> <span data-ttu-id="938ca-227">请注意，该代码不会加载 `isNullObject` 属性。</span><span class="sxs-lookup"><span data-stu-id="938ca-227">Note that the code does not load the `isNullObject` property.</span></span> <span data-ttu-id="938ca-228">当调用此属性时，Office `context.sync` 加载，因此不需要使用 `datasheet.load('isNullObject')`等内容显式加载。</span><span class="sxs-lookup"><span data-stu-id="938ca-228">Office automatically loads this property when `context.sync` is called, so you do not need to explicitly load it with something like `datasheet.load('isNullObject')`.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="938ca-229">另请参阅</span><span class="sxs-lookup"><span data-stu-id="938ca-229">See also</span></span>

* [<span data-ttu-id="938ca-230">常见的 JavaScript API 对象模型</span><span class="sxs-lookup"><span data-stu-id="938ca-230">Common JavaScript API object model</span></span>](office-javascript-api-object-model.md)
* [<span data-ttu-id="938ca-231">Office 外接程序的资源限制和性能优化</span><span class="sxs-lookup"><span data-stu-id="938ca-231">Resource limits and performance optimization for Office Add-ins</span></span>](../concepts/resource-limits-and-performance-optimization.md)
