---
title: 使用 Excel JavaScript API 的基本编程概念
description: 使用 Excel JavaScript API 构建适用于 Excel 的加载项。
ms.date: 10/03/2018
ms.openlocfilehash: f93ec7b5e34f90f2d61f29d861b7e0c19f66f6e3
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/12/2018
ms.locfileid: "25505984"
---
# <a name="fundamental-programming-concepts-with-the-excel-javascript-api"></a><span data-ttu-id="2a0d2-103">使用 Excel JavaScript API 的基本编程概念</span><span class="sxs-lookup"><span data-stu-id="2a0d2-103">Fundamental programming concepts with the Excel JavaScript API</span></span>
 
<span data-ttu-id="2a0d2-p101">本文介绍如何使用 [Excel JavaScript API](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview?view=office-js) 生成适用于 Excel 2016 或更高版本的外接程序。它引入了一些核心概念，这些概念是使用 API 的基础，并为执行特定任务提供指导，如读取或写入较大区域、更新区域内的所有单元格等等。</span><span class="sxs-lookup"><span data-stu-id="2a0d2-p101">This article describes how to use the [Excel JavaScript API](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview?view=office-js) to build add-ins for Excel 2016 or later. It introduces core concepts that are fundamental to using the API and provides guidance for performing specific tasks such as reading or writing to a large range, updating all cells in range, and more.</span></span>

## <a name="asynchronous-nature-of-excel-apis"></a><span data-ttu-id="2a0d2-106">Excel API 的异步特性</span><span class="sxs-lookup"><span data-stu-id="2a0d2-106">Asynchronous nature of Excel APIs</span></span>

<span data-ttu-id="2a0d2-p102">基于 Web 的 Excel 加载项在浏览器容器内运行，该容器内嵌在基于桌面平台（如 Office for Windows）上的 Office 应用程序中，并在 Office Online 中的 HTML iFrame 内运行。出于性能考虑，启用 Office.js API 以与所有支持平台上的 Excel 主机进行同步交互是不可行的。因此，Office.js 中的 **sync()** API 调用返回 Excel 应用程序完成请求的读取或写入操作时解析的 [promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) 。此外，可以将多个操作加入队列，例如设置属性或调用方法，并通过对 **sync()** 的单一调用将它们作为一批命令运行，而不是为每个操作发送单独的请求。以下部分介绍了如何使用 **Excel.run()** 和 **sync()** API 来完成此操作。</span><span class="sxs-lookup"><span data-stu-id="2a0d2-p102">The web-based Excel add-ins run inside a browser container that is embedded within the Office application on desktop-based platforms such as Office for Windows and runs inside an HTML iFrame in Office Online. Enabling the Office.js API to interact synchronously with the Excel host across all supported platforms is not feasible due to performance considerations. Therefore, the **sync()** API call in Office.js returns a [promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) that is resolved when the Excel application completes the requested read or write actions. Also, you can queue up multiple actions, such as setting properties or invoking methods, and run them as a batch of commands with a single call to **sync()**, rather than sending a separate request for each action. The following sections describe how to accomplish this using the **Excel.run()** and **sync()** APIs.</span></span>
 
## <a name="excelrun"></a><span data-ttu-id="2a0d2-112">Excel.run</span><span class="sxs-lookup"><span data-stu-id="2a0d2-112">Excel.run</span></span>
 
<span data-ttu-id="2a0d2-p103">**Excel.run** 执行一个函数，可以在其中指定要对 Excel 对象模型执行的操作。**Excel.run** 自动创建可用于与 Excel 对象进行交互的请求上下文。完成 **Excel.run** 时，将实现承诺，并自动释放在运行时分配的任何对象。</span><span class="sxs-lookup"><span data-stu-id="2a0d2-p103">**Excel.run** executes a function where you specify the actions to perform against the Excel object model. **Excel.run** automatically creates a request context that you can use to interact with Excel objects. When **Excel.run** completes, a promise is resolved, and any objects that were allocated at runtime are automatically released.</span></span>
 
<span data-ttu-id="2a0d2-p104">以下示例演示如何使用 **Excel.run**。Catch 语句捕获并记录 **Excel.run** 中发生的错误。</span><span class="sxs-lookup"><span data-stu-id="2a0d2-p104">The following example shows how to use **Excel.run**. The catch statement catches and logs errors that occur within the **Excel.run**.</span></span>
 
```js
Excel.run(function (context) {
  // You can use the Excel JavaScript API here in the batch function
  // to execute actions on the Excel object model.
  console.log('Your code goes here.');
}).catch(function (error) {
  console.log('error: ' + error);
  if (error instanceof OfficeExtension.Error) {
    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
  }
});
```

## <a name="request-context"></a><span data-ttu-id="2a0d2-118">请求上下文</span><span class="sxs-lookup"><span data-stu-id="2a0d2-118">Request context</span></span>
 
<span data-ttu-id="2a0d2-p105">Excel 和加载项在两个不同的进程中运行。由于它们使用不同的运行时环境，因此 Excel 加载项需要使用 **RequestContext** 对象，将加载项连接到 Excel 中的对象，如工作表、区域、图表和表格。</span><span class="sxs-lookup"><span data-stu-id="2a0d2-p105">Excel and your add-in run in two different processes. Since they use different runtime environments, Excel add-ins require a **RequestContext** object in order to connect your add-in to objects in Excel such as worksheets, ranges, charts, and tables.</span></span>
 
## <a name="proxy-objects"></a><span data-ttu-id="2a0d2-121">代理对象</span><span class="sxs-lookup"><span data-stu-id="2a0d2-121">Proxy objects</span></span>
 
<span data-ttu-id="2a0d2-p106">在加载项中声明和使用的 Excel JavaScript 对象为代理对象。调用的任何方法或在代理对象上设置或加载的属性都只是添加到挂起命令的队列中。在请求上下文（例如 `context.sync()`）上调用 **sync()** 方法时，已加入队列的命令将被发送到 Excel 并运行。从根本上来说，Excel JavaScript API 是以批处理为中心的。可以在请求上下文中将任意数量的更改加入队列，然后调用 **sync()** 方法来运行此批已加入队列的命令。</span><span class="sxs-lookup"><span data-stu-id="2a0d2-p106">The Excel JavaScript objects that you declare and use in an add-in are proxy objects. Any methods that you invoke or properties that you set or load on proxy objects are simply added to a queue of pending commands. When you call the **sync()** method on the request context (for example, `context.sync()`), the queued commands are dispatched to Excel and run. The Excel JavaScript API is fundamentally batch-centric. You can queue up as many changes as you wish on the request context, and then call the **sync()** method to run the batch of queued commands.</span></span>
 
<span data-ttu-id="2a0d2-p107">例如，下面的代码段声明本地 JavaScript 对象 **selectedRange** 以引用 Excel 文档中选定的区域，然后在该对象上设置某些属性。**selectedRange**  对象是一个代理对象，因此在该对象上所设置的属性以及调用的方法将不会反映在 Excel 文档中，直到加载项调用 **context.sync()**。</span><span class="sxs-lookup"><span data-stu-id="2a0d2-p107">For example, the following code snippet declares the local JavaScript object **selectedRange** to reference a selected range in the Excel document, and then sets some properties on that object. The **selectedRange** object is a proxy object, so the properties that are set and method that is invoked on that object will not be reflected in the Excel document until your add-in calls **context.sync()**.</span></span>
 
```js
const selectedRange = context.workbook.getSelectedRange();
selectedRange.format.fill.color = "#4472C4";
selectedRange.format.font.color = "white";
selectedRange.format.autofitColumns();
```
 
### <a name="sync"></a><span data-ttu-id="2a0d2-129">sync()</span><span class="sxs-lookup"><span data-stu-id="2a0d2-129">sync()</span></span>
 
<span data-ttu-id="2a0d2-p108">在请求上下文中调用 **sync()**  方法将在 Excel 文档中同步代理对象与对象之间的状态。**Sync()** 方法运行在请求上下文中加入队列的所有命令，并检索应该在代理对象上加载的任何属性的值。**Sync()** 方法以异步方式执行并返回 [promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)（在**sync()** 方法完成后解析）。</span><span class="sxs-lookup"><span data-stu-id="2a0d2-p108">Calling the **sync()** method on the request context synchronizes the state between proxy objects and objects in the Excel document. The **sync()** method runs any commands that are queued on the request context and retrieves values for any properties that should be loaded on the proxy objects. The **sync()** method executes asynchronously and returns a [promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise), which is resolved when the **sync()** method completes.</span></span>
 
<span data-ttu-id="2a0d2-133">下面的示例演示了一个批处理函数，它定义本地 JavaScript 代理对象 (**selectedRange**)，加载该对象的属性，然后使用 JavaScript Promises 模式调用 **context.sync()** 以同步 Excel 文档中代理对象与对象之间的状态。</span><span class="sxs-lookup"><span data-stu-id="2a0d2-133">The following example shows a batch function that defines a local JavaScript proxy object (**selectedRange**), loads a property of that object, and then uses the JavaScript Promises pattern to call **context.sync()** to synchronize the state between proxy objects and objects in the Excel document.</span></span>
 
```js
Excel.run(function (context) {
  const selectedRange = context.workbook.getSelectedRange();
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
 
<span data-ttu-id="2a0d2-134">在前面的示例中设置了 **selectedRange**，并在调用 **context.sync()** 时加载其 **address** 属性。</span><span class="sxs-lookup"><span data-stu-id="2a0d2-134">In the previous example, **selectedRange** is set and its **address** property is loaded when **context.sync()** is called.</span></span>
 
<span data-ttu-id="2a0d2-p109">由于 **sync()** 是一个返回 promise 的异步操作，因此，（在 JavaScript 中）应始终**返回** promise。这样做可确保在脚本继续运行之前完成 **sync()** 操作。有关使用 **sync()** 优化性能的详细信息，请参阅 [Excel JavaScript API 性能优化](https://docs.microsoft.com/office/dev/add-ins/excel/performance)。</span><span class="sxs-lookup"><span data-stu-id="2a0d2-p109">Because **sync()** is an asynchronous operation that returns a promise, you should always **return** the promise (in JavaScript). Doing so ensures that the **sync()** operation completes before the script continues to run. For more information about optimizing performance with **sync()**, see [Excel JavaScript API performance optimization](https://docs.microsoft.com/office/dev/add-ins/excel/performance).</span></span>
 
### <a name="load"></a><span data-ttu-id="2a0d2-138">load()</span><span class="sxs-lookup"><span data-stu-id="2a0d2-138">load()</span></span>
 
<span data-ttu-id="2a0d2-p110">在可以读取代理对象的属性之前，必须显式加载这些属性，以便使用 Excel 文档中的数据填充代理对象，然后调用 **context.sync()**。例如，如果创建代理对象来引用选定的区域，然后希望读取所选区域的 **address** 属性，需要首先加载 **address** 属性，然后才可以读取它。若要请求获取加载的代理对象的属性，请对对象调用 **load()** 方法，并指定要加载的属性。</span><span class="sxs-lookup"><span data-stu-id="2a0d2-p110">Before you can read the properties of a proxy object, you must explicitly load the properties to populate the proxy object with data from the Excel document, and then call **context.sync()**. For example, if you create a proxy object to reference a selected range, and then want to read the selected range's **address** property, you need to load the **address** property before you can read it. To request properties of a proxy object be loaded, call the **load()** method on the object and specify the properties to load.</span></span> 

> [!NOTE]
> <span data-ttu-id="2a0d2-p111">如果只在代理对象上调用方法或设置属性，无需调用 **load()** 方法。只在要读取代理对象属性时，才需要调用 **Load()** 方法。</span><span class="sxs-lookup"><span data-stu-id="2a0d2-p111">If you are only calling methods or setting properties on a proxy object, you do not need to call the **load()** method. The **load()** method is only required when you want to read properties on a proxy object.</span></span>
 
<span data-ttu-id="2a0d2-p112">类似于对代理对象设置属性或调用方法的请求，加载代理对象属性的请求会被添加到请求上下文的挂起命令队列中，将在下一次调用 **sync()** 方法时运行。必要时，可以将请求上下文中尽可能多的 **load()** 调用排入队列。</span><span class="sxs-lookup"><span data-stu-id="2a0d2-p112">Just like requests to set properties or invoke methods on proxy objects, requests to load properties on proxy objects get added to the queue of pending commands on the request context, which will run the next time you call the **sync()** method. You can queue up as many **load()** calls on the request context as necessary.</span></span>
 
<span data-ttu-id="2a0d2-146">下面的示例仅加载区域的特定属性。</span><span class="sxs-lookup"><span data-stu-id="2a0d2-146">In the following example, only specific properties of the range are loaded.</span></span>
 
```js
Excel.run(function (context) {
  const sheetName = 'Sheet1';
  const rangeAddress = 'A1:B2';
  const myRange = context.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
 
  myRange.load(['address', 'format/*', 'format/fill', 'entireRow' ]);
 
  return context.sync()
    .then(function () {
      console.log (myRange.address);              // ok
      console.log (myRange.format.wrapText);      // ok
      console.log (myRange.format.fill.color);    // ok
      //console.log (myRange.format.font.color);  // not ok as it was not loaded
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
 
<span data-ttu-id="2a0d2-147">在前面的示例中，由于在调用 **myRange.load()** 时未指定 `format/font`，因此无法读取 `format.font.color` 属性。</span><span class="sxs-lookup"><span data-stu-id="2a0d2-147">In the previous example, because `format/font` is not specified in the call to **myRange.load()**, the `format.font.color` property cannot be read.</span></span>

<span data-ttu-id="2a0d2-148">为了优化性能，应该在对象上使用 **load()** 方法时明确指定要加载的属性和关系，如 [Excel JavaScript API 性能优化](performance.md)所述。</span><span class="sxs-lookup"><span data-stu-id="2a0d2-148">To optimize performance, you should explicitly specify the properties and relationships to load when using the **load()** method on an object, as covered in [Excel JavaScript API performance optimizations](performance.md).</span></span> <span data-ttu-id="2a0d2-149">有关 **load()** 方法的详细信息，请参阅 [Excel JavaScript API 的高级编程概念](excel-add-ins-advanced-concepts.md)。</span><span class="sxs-lookup"><span data-stu-id="2a0d2-149">For more information about the **load()** method, see [Advanced programming concepts with the Excel JavaScript API](excel-add-ins-advanced-concepts.md).</span></span>

## <a name="null-or-blank-property-values"></a><span data-ttu-id="2a0d2-150">null 或空属性值</span><span class="sxs-lookup"><span data-stu-id="2a0d2-150">null or blank property values</span></span>
 
### <a name="null-input-in-2-d-array"></a><span data-ttu-id="2a0d2-151">二维数组中的 null 输入</span><span class="sxs-lookup"><span data-stu-id="2a0d2-151">null input in 2-D Array</span></span>
 
<span data-ttu-id="2a0d2-p114">在 Excel 中，一个区域由一个二维数组表示，其中第一个维度是行，第二个维度是列。若要仅为某个区域内的特定单元格设置值、数字格式或公式，请指定二维数组中这些单元格的值、数字格式或公式，并为二维数组中的所有其他单元格指定 `null`。</span><span class="sxs-lookup"><span data-stu-id="2a0d2-p114">In Excel, a range is represented by a 2-D array, where the first dimension is rows and the second dimension is columns. To set values, number format, or formula for only specific cells within a range, specify the values, number format, or formula for those cells in the 2-D array, and specify `null` for all other cells in the 2-D array.</span></span>
 
<span data-ttu-id="2a0d2-p115">例如，要更新一个区域内某一个单元格的数字格式，并保留该区域内所有其他单元格的现有数字格式，可指定要更新的单元格的新数字格式，并为所有其他单元格指定 `null`。下面的代码段为该区域内的第四个单元格设置了一个新的数字格式，并保留该区域内前三个单元格的数字格式不变。</span><span class="sxs-lookup"><span data-stu-id="2a0d2-p115">For example, to update the number format for only one cell within a range, and retain the existing number format for all other cells in the range, specify the new number format for the cell to update, and specify `null` for all other cells. The following code snippet sets a new number format for the fourth cell in the range, and leaves the number format unchanged for the first three cells in the range.</span></span>
 
```js
range.values = [['Eurasia', '29.96', '0.25', '15-Feb' ]];
range.numberFormat = [[null, null, null, 'm/d/yyyy;@']];
```
 
### <a name="null-input-for-a-property"></a><span data-ttu-id="2a0d2-156">属性的 null 输入</span><span class="sxs-lookup"><span data-stu-id="2a0d2-156">null input for a property</span></span>
 
<span data-ttu-id="2a0d2-p116">`null` 不是单个属性的有效输入。例如，下面的代码片段无效，因为区域的 **values** 属性不能设置为 `null`。</span><span class="sxs-lookup"><span data-stu-id="2a0d2-p116">`null` is not a valid input for single property. For example, the following code snippet is not valid, as the **values** property of the range cannot be set to `null`.</span></span>
 
```js
range.values = null;
```
 
<span data-ttu-id="2a0d2-159">同样，下面的代码片段也无效，因为 `null` 不是 **color** 属性的有效值。</span><span class="sxs-lookup"><span data-stu-id="2a0d2-159">Likewise, the following code snippet is not valid, as `null` is not a valid value for the **color** property.</span></span>
 
```js
range.format.fill.color =  null;
```
 
### <a name="null-property-values-in-the-response"></a><span data-ttu-id="2a0d2-160">响应中的 null 属性值</span><span class="sxs-lookup"><span data-stu-id="2a0d2-160">null property values in the response</span></span>
 
<span data-ttu-id="2a0d2-p117">如果指定区域内存在不同的值，诸如 `size` 和 `color` 等格式化属性将在响应中包含 `null` 值。例如，如果你检索某个区域并加载其 `format.font.color` 属性：</span><span class="sxs-lookup"><span data-stu-id="2a0d2-p117">Formatting properties such as `size` and `color` will contain `null` values in the response when different values exist in the specified range. For example, if you retrieve a range and load its `format.font.color` property:</span></span>
 
* <span data-ttu-id="2a0d2-163">如果区域中的所有单元格都具有相同的字体颜色，则 `range.format.font.color` 会指定该颜色。</span><span class="sxs-lookup"><span data-stu-id="2a0d2-163">If all cells in the range have the same font color, `range.format.font.color` specifies that color.</span></span>
* <span data-ttu-id="2a0d2-164">如果该区域内存在多种字体颜色，则 `range.format.font.color` 为 `null`。</span><span class="sxs-lookup"><span data-stu-id="2a0d2-164">If multiple font colors are present within the range, `range.format.font.color` is `null`.</span></span>
 
### <a name="blank-input-for-a-property"></a><span data-ttu-id="2a0d2-165">属性的空白输入</span><span class="sxs-lookup"><span data-stu-id="2a0d2-165">Blank input for a property</span></span>
 
<span data-ttu-id="2a0d2-p118">如果为属性指定空白值（即两个引号之间没有空格 `''`），它会被解释为属性清除或重置指令。例如：</span><span class="sxs-lookup"><span data-stu-id="2a0d2-p118">When you specify a blank value for a property (i.e., two quotation marks with no space in-between `''`), it will be interpreted as an instruction to clear or reset the property. For example:</span></span>
 
* <span data-ttu-id="2a0d2-168">如果为区域的 `values` 属性指定空白值，此区域的内容会被清除。</span><span class="sxs-lookup"><span data-stu-id="2a0d2-168">If you specify a blank value for the `values` property of a range, the content of the range is cleared.</span></span>
 
* <span data-ttu-id="2a0d2-169">如果为 `numberFormat` 属性指定一个空值，则数字格式会重置为 `General`。</span><span class="sxs-lookup"><span data-stu-id="2a0d2-169">If you specify a blank value for the `numberFormat` property, the number format is reset to `General`.</span></span>
 
* <span data-ttu-id="2a0d2-170">如果为 `formula` 属性和 `formulaLocale` 属性指定一个空值，则公式值将被清除。</span><span class="sxs-lookup"><span data-stu-id="2a0d2-170">If you specify a blank value for the `formula` property and `formulaLocale` property, the formula values are cleared.</span></span>
 
### <a name="blank-property-values-in-the-response"></a><span data-ttu-id="2a0d2-171">响应中的空属性值</span><span class="sxs-lookup"><span data-stu-id="2a0d2-171">Blank property values in the response</span></span>
 
<span data-ttu-id="2a0d2-p119">对于读取操作，响应中的空属性值（即两个引号之间没有空格 `''`）指示该单元格不包含任何数据或值。在下面第一个示例中，区域中的第一个和最后一个单元格不包含任何数据。在第二个示例中，区域中的前两个单元格不包含公式。</span><span class="sxs-lookup"><span data-stu-id="2a0d2-p119">For read operations, a blank property value in the response (i.e., two quotation marks with no space in-between `''`) indicates that cell contains no data or value. In the first example below, the first and last cell in the range contain no data. In the second example, the first two cells in the range do not contain a formula.</span></span>
 
```js
range.values = [['', 'some', 'data', 'in', 'other', 'cells', '']];
```
 
```js
range.formula = [['', '', '=Rand()']];
```
 
## <a name="read-or-write-to-an-unbounded-range"></a><span data-ttu-id="2a0d2-175">读取或写入无限区域</span><span class="sxs-lookup"><span data-stu-id="2a0d2-175">Read or write to an unbounded range</span></span>
 
### <a name="read-an-unbounded-range"></a><span data-ttu-id="2a0d2-176">读取无限区域</span><span class="sxs-lookup"><span data-stu-id="2a0d2-176">Read an unbounded range</span></span>
 
<span data-ttu-id="2a0d2-p120">无限区域地址是指定整个列（一列或多列）或整个行（一行或多行）的区域地址。例如：</span><span class="sxs-lookup"><span data-stu-id="2a0d2-p120">An unbounded range address is a range address that specifies either entire column(s) or entire row(s). For example:</span></span>
 
* <span data-ttu-id="2a0d2-179">包含整个列（一列或多列）的区域地址：</span><span class="sxs-lookup"><span data-stu-id="2a0d2-179">Range addresses comprised of entire column(s):</span></span><ul><li>`C:C`</li><li>`A:F`</li></ul>
* <span data-ttu-id="2a0d2-180">包含整个行（一行或多行）的区域地址：</span><span class="sxs-lookup"><span data-stu-id="2a0d2-180">Range addresses comprised of entire row(s):</span></span><ul><li>`2:2`</li><li>`1:4`</li></ul>
 
<span data-ttu-id="2a0d2-p121">API 发出请求以检索无限区域时（例如，`getRange('C:C')`），该响应将包含单元格级别属性（如 `values`、`text`、`numberFormat` 和 `formula`）的 `null` 值。其他区域属性（如 `address` 和 `cellCount`）将包含无限区域的有效值。</span><span class="sxs-lookup"><span data-stu-id="2a0d2-p121">When the API makes a request to retrieve an unbounded range (for example, `getRange('C:C')`), the response will contain `null` values for cell-level properties such as `values`, `text`, `numberFormat`, and `formula`. Other properties of the range, such as `address` and `cellCount`, will contain valid values for the unbounded range.</span></span>
 
### <a name="write-to-an-unbounded-range"></a><span data-ttu-id="2a0d2-183">写入一个无限区域</span><span class="sxs-lookup"><span data-stu-id="2a0d2-183">Write to an unbounded range</span></span>
 
<span data-ttu-id="2a0d2-p122">由于输入请求过大，因此不能在无限区域中设置单元格级别的属性，如 `values`、`numberFormat` 和 `formula`。例如，下面的代码段无效，因为它尝试为无限区域指定 `values`。如果尝试为无限区域设置单元格级别的属性，API 将返回一个错误。</span><span class="sxs-lookup"><span data-stu-id="2a0d2-p122">You cannot set cell-level properties such as `values`, `numberFormat`, and `formula` on unbounded range because the input request is too large. For example, the following code snippet is not valid because it attempts to specify `values` for an unbounded range. The API will return an error if you attempt to set cell-level properties for an unbounded range.</span></span>
 
```js
const range = context.workbook.worksheets.getActiveWorksheet().getRange('A:B');
range.values = 'Due Date';
```
 
## <a name="read-or-write-to-a-large-range"></a><span data-ttu-id="2a0d2-187">读取或写入较大区域</span><span class="sxs-lookup"><span data-stu-id="2a0d2-187">Read or write to a large range</span></span>
 
<span data-ttu-id="2a0d2-p123">如果区域中包含大量单元格、值、数字格式和/或公式，它可能无法在该区域运行 API 操作。API 将始终尽量尝试在区域内运行所请求的操作（即检索或写入指定的数据），但尝试对较大区域执行读取或写入操作可能会因资源利用率过高而导致 API 错误。为避免此类错误，建议为较大区域的较小子集运行单独的读取或写入操作，而不是尝试在较大区域内运行单个读取或写入操作。</span><span class="sxs-lookup"><span data-stu-id="2a0d2-p123">If a range contains a large number of cells, values, number formats, and/or formulas, it may not be possible to run API operations on that range. The API will always make a best attempt to run the requested operation on a range (i.e., to retrieve or write the specified data), but attempting to perform read or write operations for a large range may result in an API error due to excessive resource utilization. To avoid such errors, we recommend that you run separate read or write operations for smaller subsets of a large range, instead of attempting to run a single read or write operation on a large range.</span></span>
 
## <a name="update-all-cells-in-a-range"></a><span data-ttu-id="2a0d2-191">更新区域中的所有单元格</span><span class="sxs-lookup"><span data-stu-id="2a0d2-191">Update all cells in a range</span></span>
 
<span data-ttu-id="2a0d2-192">要对一个区域内的所有单元格应用相同更新（例如，用相同的值填充所有单元格、设置相同的数字格式，或者用相同的公式填充所有单元格），可以将 **range** 对象的相应属性设置为所需的（单个）值。</span><span class="sxs-lookup"><span data-stu-id="2a0d2-192">To apply the same update to all cells in a range, (for example, to populate all cells with the same value, set the same number format, or populate all cells with the same formula), set the corresponding property on the **range** object to the desired (single) value.</span></span>
 
<span data-ttu-id="2a0d2-193">下面的示例获取一个包含 20 个单元格的区域，然后设置数字格式，并使用值 **3/11/2015** 填充区域内的所有单元格。</span><span class="sxs-lookup"><span data-stu-id="2a0d2-193">The following example gets a range that contains 20 cells, and then sets the number format and populates all cells in the range with the value **3/11/2015**.</span></span>
 
```js
Excel.run(function (context) {
  const sheetName = 'Sheet1';
  const rangeAddress = 'A1:A20';
  const worksheet = context.workbook.worksheets.getItem(sheetName);
 
  const range = worksheet.getRange(rangeAddress);
  range.numberFormat = 'm/d/yyyy';
  range.values = '3/11/2015';
  range.load('text');
 
  return context.sync()
    .then(function () {
      console.log(range.text);
  });
}).catch(function (error) {
  console.log('Error: ' + error);
  if (error instanceof OfficeExtension.Error) {
    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
  }
});
```
 
## <a name="error-messages"></a><span data-ttu-id="2a0d2-194">错误消息</span><span class="sxs-lookup"><span data-stu-id="2a0d2-194">Error messages</span></span>
 
<span data-ttu-id="2a0d2-p124">出现 API 错误时，API 将返回包含代码和消息的 **error** 对象。下表定义了 API 可能返回的错误列表。</span><span class="sxs-lookup"><span data-stu-id="2a0d2-p124">When an API error occurs, the API will return an **error** object that contains a code and a message. The following table defines a list of errors that the API may return.</span></span>
 
|<span data-ttu-id="2a0d2-197">error.code</span><span class="sxs-lookup"><span data-stu-id="2a0d2-197">error.code</span></span> | <span data-ttu-id="2a0d2-198">error.message</span><span class="sxs-lookup"><span data-stu-id="2a0d2-198">error.message</span></span> |
|:----------|:--------------|
|<span data-ttu-id="2a0d2-199">InvalidArgument</span><span class="sxs-lookup"><span data-stu-id="2a0d2-199">InvalidArgument</span></span> |<span data-ttu-id="2a0d2-200">自变量无效、缺少或格式不正确。</span><span class="sxs-lookup"><span data-stu-id="2a0d2-200">The argument is invalid or missing or has an incorrect format.</span></span>|
|<span data-ttu-id="2a0d2-201">InvalidRequest</span><span class="sxs-lookup"><span data-stu-id="2a0d2-201">InvalidRequest</span></span>  |<span data-ttu-id="2a0d2-202">无法处理此请求。</span><span class="sxs-lookup"><span data-stu-id="2a0d2-202">Cannot process the request.</span></span>|
|<span data-ttu-id="2a0d2-203">InvalidReference</span><span class="sxs-lookup"><span data-stu-id="2a0d2-203">InvalidReference</span></span>|<span data-ttu-id="2a0d2-204">此引用对于当前操作无效。</span><span class="sxs-lookup"><span data-stu-id="2a0d2-204">This reference is not valid for the current operation.</span></span>|
|<span data-ttu-id="2a0d2-205">InvalidBinding</span><span class="sxs-lookup"><span data-stu-id="2a0d2-205">InvalidBinding</span></span>  |<span data-ttu-id="2a0d2-206">由于之前的更新，此对象绑定不再有效。</span><span class="sxs-lookup"><span data-stu-id="2a0d2-206">This object binding is no longer valid due to previous updates.</span></span>|
|<span data-ttu-id="2a0d2-207">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="2a0d2-207">InvalidSelection</span></span>|<span data-ttu-id="2a0d2-208">当前选定内容对于此操作无效。</span><span class="sxs-lookup"><span data-stu-id="2a0d2-208">The current selection is invalid for this operation.</span></span>|
|<span data-ttu-id="2a0d2-209">Unauthenticated</span><span class="sxs-lookup"><span data-stu-id="2a0d2-209">Unauthenticated</span></span> |<span data-ttu-id="2a0d2-210">所需的身份验证信息缺少或无效。</span><span class="sxs-lookup"><span data-stu-id="2a0d2-210">Required authentication information is either missing or invalid.</span></span>|
|<span data-ttu-id="2a0d2-211">AccessDenied</span><span class="sxs-lookup"><span data-stu-id="2a0d2-211">AccessDenied</span></span> |<span data-ttu-id="2a0d2-212">无法执行所请求的操作。</span><span class="sxs-lookup"><span data-stu-id="2a0d2-212">You cannot perform the requested operation.</span></span>|
|<span data-ttu-id="2a0d2-213">ItemNotFound</span><span class="sxs-lookup"><span data-stu-id="2a0d2-213">ItemNotFound</span></span> |<span data-ttu-id="2a0d2-214">所请求的资源不存在。</span><span class="sxs-lookup"><span data-stu-id="2a0d2-214">The requested resource doesn't exist.</span></span>|
|<span data-ttu-id="2a0d2-215">ActivityLimitReached</span><span class="sxs-lookup"><span data-stu-id="2a0d2-215">ActivityLimitReached</span></span>|<span data-ttu-id="2a0d2-216">已达到活动限制。</span><span class="sxs-lookup"><span data-stu-id="2a0d2-216">Activity limit has been reached.</span></span>|
|<span data-ttu-id="2a0d2-217">GeneralException</span><span class="sxs-lookup"><span data-stu-id="2a0d2-217">GeneralException</span></span>|<span data-ttu-id="2a0d2-218">处理请求时出现内部错误。</span><span class="sxs-lookup"><span data-stu-id="2a0d2-218">There was an internal error while processing the request.</span></span>|
|<span data-ttu-id="2a0d2-219">NotImplemented</span><span class="sxs-lookup"><span data-stu-id="2a0d2-219">NotImplemented</span></span>  |<span data-ttu-id="2a0d2-220">所请求的功能未实现。</span><span class="sxs-lookup"><span data-stu-id="2a0d2-220">The requested feature isn't implemented.</span></span>|
|<span data-ttu-id="2a0d2-221">ServiceNotAvailable</span><span class="sxs-lookup"><span data-stu-id="2a0d2-221">ServiceNotAvailable</span></span>|<span data-ttu-id="2a0d2-222">服务不可用。</span><span class="sxs-lookup"><span data-stu-id="2a0d2-222">The service is unavailable.</span></span>|
|<span data-ttu-id="2a0d2-223">冲突</span><span class="sxs-lookup"><span data-stu-id="2a0d2-223">Conflict</span></span>              |<span data-ttu-id="2a0d2-224">由于冲突，无法处理请求。</span><span class="sxs-lookup"><span data-stu-id="2a0d2-224">Request could not be processed because of a conflict.</span></span>|
|<span data-ttu-id="2a0d2-225">ItemAlreadyExists</span><span class="sxs-lookup"><span data-stu-id="2a0d2-225">ItemAlreadyExists</span></span>|<span data-ttu-id="2a0d2-226">所创建的资源已存在。</span><span class="sxs-lookup"><span data-stu-id="2a0d2-226">The resource being created already exists.</span></span>|
|<span data-ttu-id="2a0d2-227">UnsupportedOperation</span><span class="sxs-lookup"><span data-stu-id="2a0d2-227">UnsupportedOperation</span></span>|<span data-ttu-id="2a0d2-228">不支持正在尝试的操作。</span><span class="sxs-lookup"><span data-stu-id="2a0d2-228">The operation being attempted is not supported.</span></span>|
|<span data-ttu-id="2a0d2-229">RequestAborted</span><span class="sxs-lookup"><span data-stu-id="2a0d2-229">RequestAborted</span></span>|<span data-ttu-id="2a0d2-230">请求在运行时已中止。</span><span class="sxs-lookup"><span data-stu-id="2a0d2-230">The request was aborted during run time.</span></span>|
|<span data-ttu-id="2a0d2-231">wApiNotAvailable</span><span class="sxs-lookup"><span data-stu-id="2a0d2-231">ApiNotAvailable</span></span>|<span data-ttu-id="2a0d2-232">请求的 API 不可用。</span><span class="sxs-lookup"><span data-stu-id="2a0d2-232">The requested API is not available.</span></span>|
|<span data-ttu-id="2a0d2-233">InsertDeleteConflict</span><span class="sxs-lookup"><span data-stu-id="2a0d2-233">InsertDeleteConflict</span></span>|<span data-ttu-id="2a0d2-234">尝试的插入或删除操作导致冲突。</span><span class="sxs-lookup"><span data-stu-id="2a0d2-234">The insert or delete operation attempted resulted in a conflict.</span></span>|
|<span data-ttu-id="2a0d2-235">InvalidOperation</span><span class="sxs-lookup"><span data-stu-id="2a0d2-235">InvalidOperation</span></span>|<span data-ttu-id="2a0d2-236">尝试的操作对于对象无效。</span><span class="sxs-lookup"><span data-stu-id="2a0d2-236">The operation attempted is invalid on the object.</span></span>|
 
## <a name="see-also"></a><span data-ttu-id="2a0d2-237">另请参阅</span><span class="sxs-lookup"><span data-stu-id="2a0d2-237">See also</span></span>
 
* [<span data-ttu-id="2a0d2-238">开始使用 Excel 加载项</span><span class="sxs-lookup"><span data-stu-id="2a0d2-238">Get started with Excel add-ins</span></span>](excel-add-ins-get-started-overview.md)
* [<span data-ttu-id="2a0d2-239">Excel 加载项代码示例</span><span class="sxs-lookup"><span data-stu-id="2a0d2-239">Excel add-ins code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples)
* [<span data-ttu-id="2a0d2-240">使用 Excel JavaScript API 的高级编程概念</span><span class="sxs-lookup"><span data-stu-id="2a0d2-240">Advanced programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-advanced-concepts.md)
* [<span data-ttu-id="2a0d2-241">Excel JavaScript API 性能优化</span><span class="sxs-lookup"><span data-stu-id="2a0d2-241">Excel JavaScript API performance optimization</span></span>](https://docs.microsoft.com/office/dev/add-ins/excel/performance)
* [<span data-ttu-id="2a0d2-242">Excel JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="2a0d2-242">Excel JavaScript API reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview?view=office-js)
