---
title: Excel JavaScript API 基本编程概念
description: 使用 Excel JavaScript API 生成 Excel 加载项。
ms.date: 06/20/2019
localization_priority: Priority
ms.openlocfilehash: eed6a7a4dcc480d93e15bbb75432a2345364a5dc
ms.sourcegitcommit: 88d81aa2d707105cf0eb55d9774b2e7cf468b03a
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/13/2019
ms.locfileid: "38301916"
---
# <a name="fundamental-programming-concepts-with-the-excel-javascript-api"></a>Excel JavaScript API 基本编程概念

本文介绍了如何使用 [Excel JavaScript API](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview) 生成 Excel 2016 或更高版本的加载项。 它引入了一些核心概念，这些概念是使用 API 的基础，并为执行特定任务提供指导，如读取或写入较大区域、更新区域内的所有单元格等等。

## <a name="asynchronous-nature-of-excel-apis"></a>Excel API 的异步特性

基于 Web 的 Excel 加载项在浏览器容器内运行，此容器内嵌在基于桌面的平台版 Office 应用程序（如 Windows 版 Office）中，并在 Office 网页版中的 HTML iFrame 内运行。出于性能考虑，启用 Office.js API 以跨所有受支持的平台与 Excel 主机进行同步交互是不可行的。因此，Office.js 中的 **sync()** API 调用返回 [promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)，它在 Excel 应用程序完成请求的读取或写入操作时进行解析。此外，还可以将多个操作排入队列（如设置属性或调用方法），并通过一次调用 **sync()** 将它们作为一批命令运行，而不是为每个操作单独发送请求。以下几个部分介绍了如何使用 **Excel.run()** 和 **sync()** API 来实现此目的。

## <a name="excelrun"></a>Excel.run

**Excel.run** 执行一个函数，可以在其中指定要对 Excel 对象模型执行的操作。 **Excel.run** 自动创建可用于与 Excel 对象进行交互的请求上下文。 完成 **Excel.run** 时，将实现承诺，并自动释放在运行时分配的任何对象。

以下示例演示如何使用 **Excel.run**。 catch 语句捕获并记录 **Excel.run** 中发生的错误。

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

### <a name="run-options"></a>运行选项

**Excel.run** 包含需要使用 [RunOptions](/javascript/api/excel/excel.runoptions) 对象的重载。 这包含一组影响函数运行时平台行为的属性。 目前，支持以下属性：

- `delayForCellEdit`：确定 Excel 是否将批处理请求延迟到用户退出单元格编辑模式时执行。 若为 **true**，批处理请求延迟到用户退出单元格编辑模式时执行。 若为 **false**，批处理请求会在用户处于单元格编辑模式时（导致无法访问用户的错误出现）自动失败。 未指定 `delayForCellEdit` 属性的默认行为等同于此属性为 **false**。

```js
Excel.run({ delayForCellEdit: true }, function (context) { ... })
```

## <a name="request-context"></a>请求上下文

Excel 和加载项在两个不同的进程中运行。由于它们使用不同的运行时环境，因此 Excel 加载项需要使用 **RequestContext** 对象，将加载项连接到 Excel 中的对象，如工作表、区域、图表和表格。

## <a name="proxy-objects"></a>代理对象

在加载项中声明和使用的 Excel JavaScript 对象为代理对象。 调用的任何方法或在代理对象上设置或加载的属性都只是添加到挂起命令的队列中。 如果在请求上下文（例如 ****）时调用 `context.sync()` 方法，已加入队列的命令将被发送到 Excel 并运行。 从根本上来说，Excel JavaScript API 是以批处理为中心的。 可以在请求上下文中将任意数量的更改加入队列，然后调用 **sync()** 方法来运行此批已加入队列的命令。

例如，下面的代码段声明本地 JavaScript 对象 **selectedRange** 以引用 Excel 文档中选定的区域，然后在该对象上设置某些属性。 **SelectedRange** 对象是一个代理对象，因此在该对象上所设置的属性以及调用的方法将不会反映在 Excel 文档中，直到加载项调用 **context.sync()**。

```js
var selectedRange = context.workbook.getSelectedRange();
selectedRange.format.fill.color = "#4472C4";
selectedRange.format.font.color = "white";
selectedRange.format.autofitColumns();
```

### <a name="sync"></a>sync()

在请求上下文中调用 **sync()** 方法将在 Excel 文档中同步代理对象与对象之间的状态。 **Sync()** 方法运行在请求上下文中加入队列的所有命令，并检索应该在代理对象上加载的任何属性的值。 **sync()** 方法以异步方式执行并返回 [promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)（在 **sync()** 方法完成后解析）。

下面的示例演示了一个批处理函数，它定义本地 JavaScript 代理对象 (**selectedRange**)，加载该对象的属性，然后使用 JavaScript Promises 模式调用 **context.sync()** 以同步 Excel 文档中代理对象与对象之间的状态。

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

在前面的示例中设置了 **selectedRange**，并在调用 **context.sync()** 时加载其 **address** 属性。

由于 **sync()** 是一个返回 promise 的异步操作，因此，（在 JavaScript 中）应始终**返回** promise。 这样做可确保在脚本继续运行之前完成 **sync()** 操作。 若要详细了解如何优化使用 **sync()** 时的性能，请参阅 [Excel JavaScript API 性能优化](/office/dev/add-ins/excel/performance)。

### <a name="load"></a>load()

在可以读取代理对象的属性之前，必须显式加载这些属性，以便使用 Excel 文档中的数据填充代理对象，然后调用 **context.sync()**。 例如，如果创建代理对象来引用选定的区域，然后希望读取所选区域的 **address** 属性，需要首先加载 **address** 属性，然后才可以读取它。 若要请求获取加载的代理对象的属性，请对对象调用 **load()** 方法，并指定要加载的属性。 

> [!NOTE]
> 如果只要对代理对象调用方法或设置属性，无需调用 **load()** 方法。 只在要读取代理对象属性时，才需要调用 **load()** 方法。

类似于对代理对象设置属性或调用方法的请求，加载代理对象属性的请求会被添加到请求上下文的挂起命令队列中，将在下一次调用 **sync()** 方法时运行。必要时，可以将请求上下文中尽可能多的 **load()** 调用排入队列。

下面的示例仅加载区域的特定属性。

```js
Excel.run(function (context) {
    var sheetName = 'Sheet1';
    var rangeAddress = 'A1:B2';
    var myRange = context.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);

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

在上一示例中，由于在调用 **myRange.load()** 时未指定 `format/font`，因此无法读取 `format.font.color` 属性。

为了优化性能，应在对对象使用 **load()** 方法时，显式指定要加载的属性，如 [Excel JavaScript API 性能优化](performance.md)中所述。 若要详细了解 **load()** 方法，请参阅 [Excel JavaScript API 高级编程概念](excel-add-ins-advanced-concepts.md)。

## <a name="null-or-blank-property-values"></a>null 或空属性值

### <a name="null-input-in-2-d-array"></a>二维数组中的 null 输入

在 Excel 中，一个区域由一个二维数组表示，其中第一个维度是行，第二个维度是列。 若要仅为某个区域内的特定单元格设置值、数字格式或公式，请指定二维数组中这些单元格的值、数字格式或公式，并为二维数组中的所有其他单元格指定 `null`。

例如，要更新一个区域内某一个单元格的数字格式，并保留该区域内所有其他单元格的现有数字格式，可指定要更新的单元格的新数字格式，并为所有其他单元格指定 `null`。 下面的代码段为该区域内的第四个单元格设置了一个新的数字格式，并保留该区域内前三个单元格的数字格式不变。

```js
range.values = [['Eurasia', '29.96', '0.25', '15-Feb' ]];
range.numberFormat = [[null, null, null, 'm/d/yyyy;@']];
```

### <a name="null-input-for-a-property"></a>属性的 null 输入

`null` 不是单个属性的有效输入。例如，下面的代码片段无效，因为区域的 **values** 属性不能设置为 `null`。

```js
range.values = null;
```

同样，下面的代码片段也无效，因为 `null` 不是 **color** 属性的有效值。

```js
range.format.fill.color =  null;
```

### <a name="null-property-values-in-the-response"></a>响应中的 null 属性值

如果指定区域内存在不同的值，诸如 `size` 和 `color` 等格式化属性将在响应中包含 `null` 值。 例如，如果你检索某个区域并加载其 `format.font.color` 属性：

- 如果区域中的所有单元格都具有相同的字体颜色，则 `range.format.font.color` 会指定该颜色。
- 如果该区域内存在多种字体颜色，则 `range.format.font.color` 为 `null`。

### <a name="blank-input-for-a-property"></a>属性的空白输入

如果为属性指定空白值（即两个引号之间没有空格 `''`），它会被解释为属性清除或重置指令。例如：

- 如果为区域的 `values` 属性指定空白值，此区域的内容会被清除。

- 如果为 `numberFormat` 属性指定一个空值，则数字格式会重置为 `General`。

- 如果为 `formula` 属性和 `formulaLocale` 属性指定一个空值，则公式值将被清除。

### <a name="blank-property-values-in-the-response"></a>响应中的空属性值

对于读取操作，响应中的空属性值（即两个引号之间没有空格 `''`）指示该单元格不包含任何数据或值。 在下面第一个示例中，区域中的第一个和最后一个单元格不包含任何数据。 在第二个示例中，区域中的前两个单元格不包含公式。

```js
range.values = [['', 'some', 'data', 'in', 'other', 'cells', '']];
```

```js
range.formula = [['', '', '=Rand()']];
```

## <a name="read-or-write-to-an-unbounded-range"></a>读取或写入无限区域

### <a name="read-an-unbounded-range"></a>读取无限区域

无限区域地址是指定整个列（一列或多列）或整个行（一行或多行）的区域地址。例如：

- 包含整个列（一列或多列）的区域地址：<ul><li>`C:C`</li><li>`A:F`</li></ul>
- 包含整个行的区域地址：<ul><li>`2:2`</li><li>`1:4`</li></ul>

API 发出请求以检索无限区域时（例如，`getRange('C:C')`），该响应将包含单元格级别属性（如 `null`、`values`、`text` 和 `numberFormat`）的 `formula` 值。 其他区域属性（如 `address` 和 `cellCount`）将包含无限区域的有效值。

### <a name="write-to-an-unbounded-range"></a>写入一个无限区域

由于输入请求过大，因此不能在无限区域中设置单元格级别的属性，如 `values`、`numberFormat` 和 `formula`。 例如，下面的代码段无效，因为它尝试为无限区域指定 `values`。 如果尝试为无限区域设置单元格级别的属性，API 将返回一个错误。

```js
var range = context.workbook.worksheets.getActiveWorksheet().getRange('A:B');
range.values = 'Due Date';
```

## <a name="read-or-write-to-a-large-range"></a>读取或写入较大区域

如果区域中包含大量单元格、值、数字格式和/或公式，它可能无法在该区域运行 API 操作。 API 将始终尽量尝试在区域内运行所请求的操作（即检索或写入指定的数据），但尝试对较大区域执行读取或写入操作可能会因资源利用率过高而导致 API 错误。 为避免此类错误，建议为较大区域的较小子集运行单独的读取或写入操作，而不是尝试在较大区域内运行单个读取或写入操作。

有关系统限制的详细信息，请参阅 [Excel 数据传输限制](../develop/common-coding-issues.md#excel-data-transfer-limits)。

## <a name="update-all-cells-in-a-range"></a>更新区域中的所有单元格

要对一个区域内的所有单元格应用相同更新（例如，用相同的值填充所有单元格、设置相同的数字格式，或者用相同的公式填充所有单元格），可以将 **range** 对象的相应属性设置为所需的（单个）值。

下面的示例获取一个包含 20 个单元格的区域，然后设置数字格式，并使用值 **3/11/2015** 填充区域内的所有单元格。

```js
Excel.run(function (context) {
    var sheetName = 'Sheet1';
    var rangeAddress = 'A1:A20';
    var worksheet = context.workbook.worksheets.getItem(sheetName);

    var range = worksheet.getRange(rangeAddress);
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

## <a name="handle-errors"></a>处理错误

当 API 错误出现时，API 返回包含代码和消息的 **error** 对象。 若要详细了解错误处理（包括 API 错误列表），请参阅[错误处理](excel-add-ins-error-handling.md)。

## <a name="see-also"></a>另请参阅

- [生成首个 Excel 加载项](../quickstarts/excel-quickstart-jquery.md)
- [Excel 加载项代码示例](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
- [Excel JavaScript API 高级编程概念](excel-add-ins-advanced-concepts.md)
- [Excel JavaScript API 性能优化](/office/dev/add-ins/excel/performance)
- [Excel JavaScript API 参考](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
- [常见的编码问题和意外的平台行为](../develop/common-coding-issues.md)。
