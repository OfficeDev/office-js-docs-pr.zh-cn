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
# <a name="using-the-application-specific-api-model"></a>使用特定于应用程序的 API 模型

本文介绍如何使用 API 模型在 Excel、Word 和 OneNote 中构建外接程序。 它介绍了使用基于承诺的 Api 的基础的核心概念。

> [!NOTE]
> Office 2013 客户端不支持此模型。 使用 [通用 API 模型](office-javascript-api-object-model.md) 来处理这些 Office 版本。 有关完整的平台可用性说明，请参阅 [适用于 office 的 Office 外接程序的 office 客户端应用程序和平台可用性](../overview/office-add-in-availability.md)。

> [!TIP]
> 此页面中的示例使用 Excel JavaScript Api，但这些概念也适用于 OneNote、Visio 和 Word JavaScript Api。

## <a name="asynchronous-nature-of-the-promise-based-apis"></a>基于承诺的 Api 的异步特性

Office 外接程序是在 Office 应用程序（如 Excel）内的浏览器容器中显示的网站。 此容器嵌入在基于桌面的平台（如 Windows 上的 Office）中的 Office 应用程序中，并在 web 上的 Office 中的 HTML iFrame 内运行。 由于性能方面的考虑，Office.js Api 无法跨所有平台与 Office 应用程序同步交互。 因此， `sync()` Office.js 中的 API 调用返回在 Office 应用程序完成请求的读取或写入操作时解决的 [承诺](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) 。 此外，还可以对多个操作（如设置属性或调用方法）进行排队，并将它们作为一批命令运行 `sync()` ，而不是为每个操作发送单独的请求。 以下各节介绍如何使用和 api 来完成此操作 `run()` `sync()` 。

## <a name="run-function"></a>*. run 函数

`Excel.run`、 `Word.run` 和 `OneNote.run` 执行一个函数，该函数指定要对 Excel、Word 和 OneNote 执行的操作。 `*.run` 自动创建可用于与 Office 对象进行交互的请求上下文。 `*.run`完成后，将会解决承诺，并且会自动释放在运行时分配的任何对象。

下面的示例演示如何使用 `Excel.run` 。 Word 和 OneNote 也使用相同的模式。

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

Office 应用程序和外接程序在两个不同的进程中运行。 由于它们使用不同的运行时环境，因此外接程序需要对象才能将 `RequestContext` 外接程序连接到 Office 中的对象，如工作表、区域、段落和表。 `RequestContext`调用时，此对象作为参数提供 `*.run` 。

## <a name="proxy-objects"></a>代理对象

您声明并与基于承诺的 Api 一起使用的 Office JavaScript 对象是代理对象。 调用的任何方法或在代理对象上设置或加载的属性都只是添加到挂起命令的队列中。 在 `sync()` 请求上下文上调用方法时 (例如， `context.sync()`) ，队列命令将被调度到 Office 应用程序并运行。 这些 Api 从根本上以批处理为中心。 您可以根据需要在请求上下文中排列任意数量的更改，然后调用 `sync()` 方法以运行队列中的命令批次。

例如，以下代码段声明了本地 JavaScript [Excel Range](/javascript/api/excel/excel.range) 对象， `selectedRange` 以引用 Excel 工作簿中的选定区域，然后设置该对象的一些属性。 该 `selectedRange` 对象是一个代理对象，因此在您的外接程序调用之前，设置的属性和在该对象上调用的方法将不会反映在 Excel 文档中 `context.sync()` 。

```js
var selectedRange = context.workbook.getSelectedRange();
selectedRange.format.fill.color = "#4472C4";
selectedRange.format.font.color = "white";
selectedRange.format.autofitColumns();
```

### <a name="performance-tip-minimize-the-number-of-proxy-objects-created"></a>性能提示：最大限度地减少创建的代理对象数

避免重复创建同一个代理对象。 如果多个操作需要同一个代理对象，则改为创建一次并将其分配给一个变量，然后在代码中使用该变量。

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

### <a name="sync"></a>sync()

`sync()`对请求上下文调用方法将同步 Office 文档中的代理对象和对象之间的状态。 该 `sync()` 方法运行在请求上下文中排队的任何命令，并检索应在代理对象上加载的任何属性的值。 `sync()`方法以异步方式执行，并返回一个[承诺](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)，该方法在 `sync()` 方法完成时得到解决。

下面的示例演示了一个批处理函数，该函数定义本地 JavaScript 代理对象 (`selectedRange`) ，加载该对象的属性，然后使用 JavaScript 承诺模式来调用， `context.sync()` 以同步 Excel 文档中的代理对象和对象之间的状态。

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

在上一示例中，已设置 `selectedRange`，并且将在调用 `context.sync()` 时加载其 `address` 属性。

由于 `sync()` 是异步操作，因此应始终返回 `Promise` 对象以确保 `sync()` 操作在脚本继续运行之前完成。 如果使用的是 TypeScript 或 ES6 + JavaScript，则可以 `await` `context.sync()` 调用，而不是返回承诺。

#### <a name="performance-tip-minimize-the-number-of-sync-calls"></a>性能提示：最大限度地减少同步调用数

在 Excel JavaScript API 中，`sync()` 是唯一的异步操作，在某些情况下可能会很慢，尤其是对于 Excel 网页版。 若要优化性能，在调用之前，通过尽可能多地将更改加入队列来最大程度减少调用 `sync()` 的次数。 有关优化性能的详细信息 `sync()` ，请参阅 [避免在循环中使用 context. sync 方法](../concepts/correlated-objects-pattern.md)。

### <a name="load"></a>load()

在可以读取代理对象的属性之前，必须显式加载属性以使用 Office 文档中的数据填充代理对象，然后再调用 `context.sync()` 。 例如，如果创建代理对象以引用选定区域，然后想要读取选定区域的 `address` 属性，则需要先加载该属性，然后才能 `address` 阅读该属性。 若要加载代理对象的属性，请对 `load()` 该对象调用方法，并指定要加载的属性。 下面的示例展示了 `Range.address` 要加载的属性 `myRange` 。

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
> 如果只调用方法或设置代理对象的属性，则不需要调用该 `load()` 方法。 `load()`仅当您想要读取代理对象的属性时，才需要使用此方法。

类似于对代理对象设置属性或调用方法的请求，加载代理对象属性的请求会被添加到请求上下文的挂起命令队列中，将在下一次调用 `sync()` 方法时运行。必要时，可以将请求上下文中尽可能多的 `load()` 调用排入队列。

#### <a name="scalar-and-navigation-properties"></a>标量和导航属性

属性分为两种类别：**标量**和**导航**。 标量属性是可分配的类型，如字符串、整数和 JSON 结构。 导航属性是只读对象和已分配字段的对象集合，而不是直接分配属性。 例如， `name` 和的 `position` 成员在 [Excel 中。工作表](/javascript/api/excel/excel.worksheet) 对象是标量属性，而 `protection` 并 `tables` 是导航属性。

您的外接程序可以将导航属性用作加载特定标量属性的路径。 下面的代码 `load` 对对象使用的字体名称的命令进行排队 `Excel.Range` ，而不加载任何其他信息。

```js
someRange.load("format/font/name")
```

您还可以通过遍历路径来设置导航属性的标量属性。 例如，可以使用设置的字体大小 `Excel.Range` `someRange.format.font.size = 10;` 。 在设置属性之前，不需要加载该属性。

请注意，一个对象下的一些属性可能与另一个对象的名称相同。 例如， `format` 是对象下的属性 `Excel.Range` ，但本身也 `format` 是对象。 因此，如果发出类似的调用，则 `range.load("format")` 等效于 `range.format.load()` (不需要的空 `load()` 语句) 。 若要避免这种情况，代码应仅加载对象树中的 "叶节点"。

#### <a name="calling-load-without-parameters-not-recommended"></a>`load`不建议调用不带参数的 () 

如果在 `load()` 未指定任何参数的情况下对对象 (或集合) 调用方法，则将加载该对象的所有标量属性或该集合的对象。 加载不需要的数据会降低外接程序的速度。 应始终显式指定要加载的属性。

> [!IMPORTANT]
> 无参数 `load` 语句返回的数据量可能超过该服务的大小限制。 为了降低较旧加载项的风险，`load` 不会在明确请求它们之前返回某些属性。 此类加载操作中排除了以下属性：
>
> * `Excel.Range.numberFormatCategories`

### <a name="clientresult"></a>ClientResult

返回基元类型的基于承诺的 api 中的方法具有与范例类似的模式 `load` / `sync` 。 举个例子，`Excel.TableCollection.getCount`获取集合中的表的数量。 `getCount` 返回 a `ClientResult<number>` ，表示 `value` 返回的属性 [`ClientResult`](/javascript/api/office/officeextension.clientresult) 为数字。 在调用 `context.sync()` 之前，脚本无法访问此值。

下面的代码获取 Excel 工作簿中的总表数，并将该数目的日志记录到控制台。

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

### <a name="set"></a>set()

在具有嵌套导航属性的对象上设置属性可能很麻烦。 除了以上所述使用导航路径设置各个属性之外，您还可以使用 `object.set()` 基于承诺的 JavaScript api 中的对象上提供的方法。 使用此方法，可以通过传递相同 Office.js 类型的另一个对象或 JavaScript 对象（其属性结构类似于调用该方法的对象的属性）一次设置对象的多个属性。

下面的代码示例设置区域的多个格式属性，具体方法是调用 `set()` 方法，并传入 JavaScript 对象，其中包含可反映 `Range` 对象中属性结构的属性名称和类型。此示例假定区域 **B2:E2** 中有数据。

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

## <a name="42ornullobject-methods-and-properties"></a>&#42;OrNullObject 方法和属性

当所需的对象不存在时，某些访问器方法和属性将引发异常。 例如，如果尝试通过指定不在工作簿中的工作表名称来获取 Excel 工作表，则该 `getItem()` 方法将引发 `ItemNotFound` 异常。 特定于应用程序的库为代码提供了一种测试文档实体是否存在的方法，而不需要异常处理代码。 这是通过使用 `*OrNullObject` 方法和属性的变体来实现的。 `isNullObject` `true` 如果指定的项不存在，而不是引发异常，则这些变体返回其属性设置为的对象。

例如，可以对 `getItemOrNullObject()` 集合（如 **工作表** ）调用方法，以从集合中检索项。 `getItemOrNullObject()`如果指定的项存在，则该方法将返回它; 否则，将返回其 `isNullObject` 属性设置为的对象 `true` 。 然后，您的代码可以对此属性进行评估，以确定该对象是否存在。

> [!NOTE]
> `*OrNullObject`变体从不返回 JavaScript 值 `null` 。 它们返回普通的 Office 代理对象。 如果该对象所代表的实体不存在，则 `isNullObject` 将该对象的属性设置为 `true` 。 请勿为 null 或 falsity 测试返回的对象。 它永远不会是、 `null` `false` 或 `undefined` 。

下面的代码示例尝试使用方法检索名为 "Data" 的 Excel 工作表 `getItemOrNullObject()` 。 如果具有该名称的工作表不存在，则创建一个新工作表。 请注意，该代码不会加载该 `isNullObject` 属性。 Office 将在调用时自动加载此属性 `context.sync` ，因此无需使用类似的内容显式加载它 `datasheet.load('isNullObject')` 。

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

## <a name="see-also"></a>另请参阅

* [常见 JavaScript API 对象模型](office-javascript-api-object-model.md)
* [常见的编码问题和意外的平台行为](common-coding-issues.md)。
* [Office 外接程序的资源限制和性能优化](../concepts/resource-limits-and-performance-optimization.md)
