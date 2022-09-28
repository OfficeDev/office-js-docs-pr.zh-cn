---
title: 使用应用程序专用 API 模型
description: 了解 Excel、OneNote 和 Word 加载项基于承诺的 API 模型。
ms.date: 09/23/2022
ms.localizationpriority: medium
ms.openlocfilehash: d7cb6f1f47c853d5c6e389c2c81ec2d36d21eb43
ms.sourcegitcommit: 05be1086deb2527c6c6ff3eafcef9d7ed90922ec
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/28/2022
ms.locfileid: "68092887"
---
# <a name="application-specific-api-model"></a>特定于应用程序的 API 模型

本文介绍如何使用 API 模型在 Excel、Word、PowerPoint 和 OneNote 中生成加载项。 本文介绍核心概念，这些概念是使用基于承诺的 API 的基础。

> [!NOTE]
> Office 2013 客户端不支持此模型。 使用 [API 模型](office-javascript-api-object-model.md) 这些 Office 版本。 有关完整的平台可用性说明，请参阅 [Office 客户端应用程序和平台可用性的 Office 加载项组](/javascript/api/requirement-sets)。

> [!TIP]
> 本页中的示例使用 Excel JavaScript API，但这些概念也适用于 OneNote、PowerPoint、Visio 和 Word JavaScript API。

## <a name="asynchronous-nature-of-the-promise-based-apis"></a>基于承诺的 API 的异步性质

Office 加载项是显示在 Office 应用程序（如 Excel）中的浏览器容器内的网站。 此容器嵌入在基于桌面的 Office 应用程序（如 Windows 上的 Office）上的 Office 应用程序中，在 Office 网页版中的 HTML iFrame 内运行。 出于性能方面的考虑，Office.js API 无法跨所有平台与 Office 应用程序同步交互。 因此，Office.js `sync()` API 调用返回 [Office](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) 请求的读取或写入操作时要解决的"承诺"问题。 此外，还可以将多个操作（例如设置属性或调用方法）排队，并且只要一次呼叫 `sync()`，就可以作为一批命令运行这些操作，而不是针对每个操作发送单独的请求。 以下各节介绍如何使用 API 和 `run()``sync()`实现此操作。

## <a name="run-function"></a>*.run 函数

`Excel.run``PowerPoint.run`，`OneNote.run`并`Word.run`执行一个函数，该函数指定要针对 Excel、Word 和 OneNote 执行的操作。 `*.run` 会自动创建可用于与 Office 对象交互的请求上下文。 当 `*.run` ，将做出承诺，并自动发布运行时分配的任何对象。

以下示例显示了如何使用 `Excel.run`。 OneNote、PowerPoint 和 Word 也使用相同的模式。

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

## <a name="request-context"></a>请求上下文

Office 应用程序和外接程序在不同的进程中运行。 由于加载项使用不同的运行时环境，因此需要一个 `RequestContext` 对象才能将加载项连接到 Office 中的对象，例如工作表、区域、段落和表。 调用 `RequestContext` 时，此对象作为 `*.run`。

## <a name="proxy-objects"></a>代理对象

声明并用于基于承诺的 API 的 Office JavaScript 对象是代理对象。 调用的任何方法或在代理对象上设置或加载的属性都只是添加到挂起命令的队列中。 调用请求 `sync()` （例如 `context.sync()`）上的方法时，排队的命令将调用 Office 应用程序并运行。 这些 API 在根本上以批处理为中心。 您可以根据对请求上下文希望排入多达数个更改，然后调用 `sync()` 方法，以运行排队命令的批处理。

例如，以下代码片段声明本地 JavaScript [Excel.Range](/javascript/api/excel/excel.range) 对象 `selectedRange`引用 Excel 工作簿中的选定区域，然后针对该对象设置一些属性。 对象 `selectedRange` 代理对象，因此在调用加载项之前，不会在 Excel 文档中反映已设置的属性和在该对象上调用 `context.sync()`。

```js
const selectedRange = context.workbook.getSelectedRange();
selectedRange.format.fill.color = "#4472C4";
selectedRange.format.font.color = "white";
selectedRange.format.autofitColumns();
```

### <a name="performance-tip-minimize-the-number-of-proxy-objects-created"></a>性能提示：最小化创建代理对象的数量

避免重复创建同一个代理对象。 如果多个操作需要同一个代理对象，则改为创建一次并将其分配给一个变量，然后在代码中使用该变量。

```js
// BAD: Repeated calls to .getRange() to create the same proxy object.
worksheet.getRange("A1").format.fill.color = "red";
worksheet.getRange("A1").numberFormat = "0.00%";
worksheet.getRange("A1").values = [[1]];

// GOOD: Create the range proxy object once and assign to a variable.
const range = worksheet.getRange("A1");
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

### <a name="sync"></a>sync()

调用 `sync()` 上下文的方法可同步 Office 文档中代理对象和对象之间的状态。 该 `sync()` 在请求上下文中排入队列的任何命令，并检索应在代理对象上加载的任何属性的值。 方法 `sync()` 异步执行，并返回 [承诺](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)，该 `sync()` 完成时完成。

下面的示例显示一个批处理函数，定义本地 JavaScript 代理对象 （`selectedRange`），加载该对象的属性，然后使用 JavaScript 形式调用 `context.sync()` 以在 Excel 文档中的代理对象和对象间同步状态。

```js
await Excel.run(async (context) => {
    const selectedRange = context.workbook.getSelectedRange();
    selectedRange.load('address');
    await context.sync();
    console.log('The selected range is: ' + selectedRange.address);
});
```

在上一示例中，已设置 `selectedRange`，并且将在调用 `context.sync()` 时加载其 `address` 属性。

由于 `sync()` 是异步操作，因此在脚本继续运行之前，应始终 `Promise` 同步对象，以确保 `sync()` 操作完成。 如果使用 TypeScript 或 ES6+ JavaScript， `await` 调用 `context.sync()` ，而不是返回承诺。

#### <a name="performance-tip-minimize-the-number-of-sync-calls"></a>性能提示：减少同步呼叫数

在 Excel JavaScript API 中，`sync()` 是唯一的异步操作，在某些情况下可能会很慢，尤其是对于 Excel 网页版。 若要优化性能，在调用之前，通过尽可能多地将更改加入队列来最大程度减少调用 `sync()` 的次数。 有关使用 `sync()`优化性能，请参阅 [循环使用 context.sync 方法](../concepts/correlated-objects-pattern.md)。

### <a name="load"></a>load()

必须显式加载属性才能读取代理对象的属性，才能使用 Office 文档中的数据填充代理对象，然后调用 `context.sync()`。 例如，如果创建代理对象来引用选定的区域，然后希望读取所选区域的 `address` 属性，需要首先加载 `address` 属性，然后才可以读取它。 若要请求加载代理对象的属性，调用 `load()` 的方法并指定要加载的属性。 以下示例显示了为 `Range.address` 加载的 `myRange`。

```js
await Excel.run(async (context) => {
    const sheetName = 'Sheet1';
    const rangeAddress = 'A1:B2';
    const myRange = context.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);

    myRange.load('address');
    await context.sync();
      
    console.log (myRange.address);   // ok
    //console.log (myRange.values);  // not ok as it was not loaded

    console.log('done');
});
```

> [!NOTE]
> 如果只是调用代理对象或设置属性，则无需调用代理 `load()` 方法。 只有在想要读取代理对象上的属性时 `load()` 代理方法才必需。

Just like requests to set properties or invoke methods on proxy objects, requests to load properties on proxy objects get added to the queue of pending commands on the request context, which will run the next time you call the `sync()` method. You can queue up as many `load()` calls on the request context as necessary.

#### <a name="scalar-and-navigation-properties"></a>标量和导航属性

属性分为两种类别：**标量** 和 **导航**。 标量属性是可分配的类型，如字符串、整数和 JSON 结构。 导航属性是分配了字段的只读对象和对象集合，而不是直接分配属性。 例如，[Worksheet](/javascript/api/excel/excel.worksheet) 对象上的 `name`和 `position` 成员是标量属性，而 `protection` 和 `tables` 是导航属性。

加载项可使用导航属性作为加载特定标量属性的路径。 以下代码会按照 `load` 对象使用的字体的名称将向上排队 `Excel.Range` 命令，而无需加载任何其他信息。

```js
someRange.load("format/font/name")
```

还可通过遍历路径来设置导航属性的标量属性。 例如，通过使用"另一种" `Excel.Range` ， `someRange.format.font.size = 10;`。 设置属性前无需加载属性。

请注意，一个对象下的某些“属性”可能与另一个对象同名。 例如， `format` 是对象下 `Excel.Range` 属性， `format` 值本身也是一个对象。 因此，如果你进行 `range.load("format")`等呼叫，这相当于 `range.format.load()` （一个空的 `load()` 语句）。 若要避免这种情况，代码应仅加载对象树中的“叶节点”。

#### <a name="calling-load-without-parameters-not-recommended"></a>不带 `load` （不推荐）的呼叫方

如果在不 `load()` 参数的情况下调用对象（或集合）上的标量方法，将加载该对象或集合对象的所有标量属性。 加载不需要的数据会降低加载项的加载速度。 应始终显式指定要加载的属性。

> [!IMPORTANT]
> 无参数 `load` 语句返回的数据量可能超过该服务的大小限制。 为了降低较旧加载项的风险，`load` 不会在明确请求它们之前返回某些属性。 从此类负载操作中排除以下属性。
>
> * `Excel.Range.numberFormatCategories`

### <a name="clientresult"></a>ClientResult

基于承诺的 API 中返回类型 API 的方法与现代方法`load`/`sync`模式。 举个例子，`Excel.TableCollection.getCount`获取集合中的表的数量。 `getCount` 返回一`ClientResult<number>`，这意味着返回的`value`中的 [`ClientResult`](/javascript/api/office/officeextension.clientresult) 属性是一个数字。 在调用 `context.sync()` 之前，脚本无法访问此值。

以下代码获取 Excel 工作簿中的表总数，并对此数目的日志记录到控制台。

```js
const tableCount = context.workbook.tables.getCount();

// This sync call implicitly loads tableCount.value.
// Any other ClientResult values are loaded too.
await context.sync();

// Trying to log the value before calling sync would throw an error.
console.log (tableCount.value);
```

### <a name="set"></a>set()

在具有嵌套导航属性的对象上设置属性可能很麻烦。 除了使用上述导航路径设置单个属性， `object.set()` 基于承诺的 JavaScript API 中的对象上可用的另一种方法。 使用此方法，可以通过传递相同 Office.js 类型的另一个对象或 JavaScript 对象（其属性结构类似于调用该方法的对象的属性）一次设置对象的多个属性。

The following code sample sets several format properties of a range by calling the `set()` method and passing in a JavaScript object with property names and types that mirror the structure of properties in the `Range` object. This example assumes that there is data in range **B2:E2**.

```js
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Sample");
    const range = sheet.getRange("B2:E2");
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

    await context.sync();
});
```

### <a name="some-properties-cannot-be-set-directly"></a>某些属性不能直接设置

尽管可写的属性，但某些属性不能设置。 这些属性是必须将设置为单个对象的父属性的一部分。 这是因为父属性依赖于具有特定逻辑关系的子属性。 必须使用对象文字表示法设置这些父属性来设置整个对象，而不是设置该对象的单个子问题。 PageLayout [中可找到此示例](/javascript/api/excel/excel.pagelayout)。 `zoom`必须使用单个 [PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions) 对象设置该属性，如下所示。

```js
// PageLayout.zoom.scale must be set by assigning PageLayout.zoom to a PageLayoutZoomOptions object.
sheet.pageLayout.zoom = { scale: 200 };
```

在上一示例中，***无法*** 直接分配为值`zoom`：`sheet.pageLayout.zoom.scale = 200;`。 该语句会引发错误， `zoom` 加载错误。 即使 `zoom` ，该比例也会生效。 所有上下文操作 `zoom`、刷新加载项中的代理对象并覆盖本地设置的值。

此行为与 [Range.format](application-specific-api-model.md#scalar-and-navigation-properties) 等 [导航属性](/javascript/api/excel/excel.range#excel-excel-range-format-member)。 `format`可以使用对象导航设置属性，如下所示。

```js
// This will set the font size on the range during the next `content.sync()`.
range.format.font.size = 10;
```

可通过检查其只读修改者，识别不能直接设置其子问题的属性。 所有只读属性可直接设置其非只读子问题。 必须在该级别 `PageLayout.zoom` 可编写属性，如属性。 摘要：

- 只读属性：可通过导航设置子项目。
- 可写的属性：无法通过导航设置子项目（必须设置为初始父对象分配的一部分）。

## <a name="42ornullobject-methods-and-properties"></a>&#42;OrNullObject 方法与属性

当所需对象不存在时，某些配件方法和属性将引发异常。 例如，如果尝试通过指定工作簿未包含的工作表名称获取 Excel 工作表，则 `getItem()` 会引发 `ItemNotFound` 异常。 特定于应用程序的库为代码提供了一种方法，用于测试文档实体是否存在，而无需异常处理代码。 此操作是通过使用多种 `*OrNullObject` 和属性实现的。 如果指定的项目不存在，这些变体将返回其 `isNullObject` 被设置为 `true`值的对象，而不是引发异常。

例如，可以在集合（如 **Worksheets**）上调用 `getItemOrNullObject()` 方法，尝试从集合中检索某个项。 方法 `getItemOrNullObject()` 返回指定项目（如果存在）;否则，将返回其属性 `isNullObject` 为 <a0/ `true`。 然后，代码可评估此属性，以确定该对象是否存在。

> [!NOTE]
> 这些 `*OrNullObject` 变体永远不会返回值 JavaScript `null`。 它们返回普通 Office 代理对象。 如果对象表示的实体不存在，则对象的 `isNullObject` 属性设置为 `true`。 不要测试返回的对象为 nullity 或 fality。 它从 `null`、 `false`或 `undefined`。

以下代码示例尝试使用以下方法检索名为"数据"的 Excel `getItemOrNullObject()`。 如果具有该名称的工作表不存在，将创建一个新工作表。 请注意，该代码不会加载 `isNullObject` 属性。 当调用此属性时，Office `context.sync` 加载，因此不需要使用 `dataSheet.load('isNullObject')`等内容显式加载。

```js
await Excel.run(async (context) => {
    let dataSheet = context.workbook.worksheets.getItemOrNullObject("Data");
    
    await context.sync();
    
    if (dataSheet.isNullObject) {
        dataSheet = context.workbook.worksheets.add("Data");
    }
    
    // Set `dataSheet` to be the second worksheet in the workbook.
    dataSheet.position = 1;
});
```

## <a name="see-also"></a>另请参阅

- [常见的 JavaScript API 对象模型](office-javascript-api-object-model.md)
- [Office 外接程序的资源限制和性能优化](../concepts/resource-limits-and-performance-optimization.md)
