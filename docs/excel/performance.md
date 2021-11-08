---
title: Excel JavaScript API 性能优化
description: 使用 javaScript API Excel优化加载项性能。
ms.date: 08/24/2021
ms.localizationpriority: medium
ms.openlocfilehash: ade2ac02f22c93d920174f54e6fc2efed349e3d5
ms.sourcegitcommit: e4b83d43c117225898a60391ea06465ba490f895
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/08/2021
ms.locfileid: "60809061"
---
# <a name="performance-optimization-using-the-excel-javascript-api"></a>使用 Excel JavaScript API 优化性能

有多种方法可以使用 Excel JavaScript API 执行常见任务。 你将发现不同方法之间的显著性能差异。 本文提供指导和代码示例，展示如何使用 Excel JavaScript API 来高效执行常见任务。

> [!IMPORTANT]
> 可以通过建议的用法和调用解决许多 `load` `sync` 性能问题。 有关有效使用特定于应用程序的 API 的建议，请参阅[Office](../concepts/resource-limits-and-performance-optimization.md#performance-improvements-with-the-application-specific-apis)外接程序的资源限制和性能优化的"使用特定于应用程序的 API 的性能改进"部分。

## <a name="suspend-excel-processes-temporarily"></a>暂时挂起 Excel 进程

Excel 中的多个后台任务将反应来自用户和外接程序的输入。 可以控制其中的部分 Excel 进程以提高性能。 这在外接程序处理大型数据集时尤其有用。

### <a name="suspend-calculation-temporarily"></a>暂停计算

如果你试图在大量单元格上执行操作（例如，设置一个大范围对象的值），而且不介意在操作完成时暂停 Excel 中的计算，建议暂停计算，直到调用下一个 `context.sync()`。

有关如何使用 `suspendApiCalculationUntilNextSync()` API 以便捷的方式暂停和重新激活计算的信息，请参阅[应用程序对象](/javascript/api/excel/excel.application)参考文档。 以下代码演示了如何暂时暂停计算。

```js
Excel.run(async function(ctx) {
    var app = ctx.workbook.application;
    var sheet = ctx.workbook.worksheets.getItem("sheet1");
    var rangeToSet: Excel.Range;
    var rangeToGet: Excel.Range;
    app.load("calculationMode");
    await ctx.sync();
    // Calculation mode should be "Automatic" by default
    console.log(app.calculationMode);

    rangeToSet = sheet.getRange("A1:C1");
    rangeToSet.values = [[1, 2, "=SUM(A1:B1)"]];
    rangeToGet = sheet.getRange("A1:C1");
    rangeToGet.load("values");
    await ctx.sync();
    // Range value should be [1, 2, 3] now
    console.log(rangeToGet.values);

    // Suspending recalculation
    app.suspendApiCalculationUntilNextSync();
    rangeToSet = sheet.getRange("A1:B1");
    rangeToSet.values = [[10, 20]];
    rangeToGet = sheet.getRange("A1:C1");
    rangeToGet.load("values");
    app.load("calculationMode");
    await ctx.sync();
    // Range value should be [10, 20, 3] when we load the property, because calculation is suspended at that point
    console.log(rangeToGet.values);
    // Calculation mode should still be "Automatic" even with suspend recalculation
    console.log(app.calculationMode);

    rangeToGet.load("values");
    await ctx.sync();
    // Range value should be [10, 20, 30] when we load the property, because calculation is resumed after last sync
    console.log(rangeToGet.values);
})
```

请注意，仅暂停公式计算。 仍将重新生成任何更改的引用。 例如，重命名工作表仍将更新公式中对此工作表的任何引用。

### <a name="suspend-screen-updating"></a>暂停屏幕更新

Excel 大约会在代码发生更改时显示外接程序所进行的这些更改。 对于大型迭代数据集，你无需实时在屏幕上查看此进度。 在外接程序调用 `context.sync()` 或者在 `Excel.run` 结束（隐式调用 `context.sync`）之前，`Application.suspendScreenUpdatingUntilNextSync()` 将暂停对 Excel 的可视化更新。 请注意，在下次同步之前，Excel 不会显示任何活动迹象。你的外接程序应为用户提供相关指南，以便为此延迟做好准备，或者提供一个状态栏，以演示相关活动。

> [!NOTE]
> 不要重复调用 `suspendScreenUpdatingUntilNextSync` (，如在循环) 。 重复调用将导致Excel闪烁。

### <a name="enable-and-disable-events"></a>启用和禁用事件

可以通过禁用事件来改进加载项性能。 [使用事件](excel-add-ins-events.md#enable-and-disable-events)文章中的代码示例展示了如何启用和禁用事件。

## <a name="importing-data-into-tables"></a>将数据导入表

当试图将大量数据直接导入到 [Table](/javascript/api/excel/excel.table) 对象中时（例如，通过使用 `TableRowCollection.add()`），可能会遇到性能缓慢的问题。 如果尝试添加一个新表，应首先通过设置 `range.values` 来填充数据，然后调用 `worksheet.tables.add()` 在该区域内创建一个表。 如果尝试将数据写入现有表，请通过 `table.getDataBodyRange()` 将数据写入一个 range 对象，表将自动展开。

下面是此方法的一个示例：

```js
Excel.run(async (ctx) => {
    var sheet = ctx.workbook.worksheets.getItem("Sheet1");
    // Write the data into the range first.
    var range = sheet.getRange("A1:B3");
    range.values = [["Key", "Value"], ["A", 1], ["B", 2]];

    // Create the table over the range
    var table = sheet.tables.add('A1:B3', true);
    table.name = "Example";
    await ctx.sync();


    // Insert a new row to the table
    table.getDataBodyRange().getRowsBelow(1).values = [["C", 3]];
    // Change a existing row value
    table.getDataBodyRange().getRow(1).values = [["D", 4]];
    await ctx.sync();
})
```

> [!NOTE]
> 可以使用 [Table.convertToRange()](/javascript/api/excel/excel.table#convertToRange__) 方法将 Table 对象转换为 Range 对象，此做法非常方便。

## <a name="payload-size-limit-best-practices"></a>有效负载大小限制最佳实践

JavaScript API Excel API 调用的大小限制。 Excel web 版请求和响应的有效负载大小限制为 5MB，如果超出此限制，API `RichAPI.Error` 将返回错误。 在所有平台上，一个范围限制为五百万个单元格，用于获取操作。 较大区域通常超过这两个限制。

请求的有效负载大小是以下三个组件的组合。

* API 调用数
* 对象的数量，例如 `Range` 对象
* 要设置或获取的值的长度

如果 API 返回错误，请使用本文中介绍的最佳实践策略来优化 `RequestPayloadSizeLimitExceeded` 脚本并避免错误。

### <a name="strategy-1-move-unchanged-values-out-of-loops"></a>策略 1：将未更改的值移出循环

限制循环内发生的进程数以提高性能。 在下面的代码示例中 `context.workbook.worksheets.getActiveWorksheet()` ，可以移出 `for` 循环，因为它不会在此循环中更改。

```js
// DO NOT USE THIS CODE SAMPLE. This sample shows a poor performance strategy. 
async function run() {
  await Excel.run(async (context) => {
    var ranges = [];
    
    // This sample retrieves the worksheet every time the loop runs, which is bad for performance.
    for (let i = 0; i < 7500; i++) {
      var rangeByIndex = context.workbook.worksheets.getActiveWorksheet().getRangeByIndexes(i, 1, 1, 1);
    }    
    await context.sync();
  });
}
```

下面的代码示例演示了与前面的代码示例类似的逻辑，但具有改进的性能策略。 该值在循环之前检索，因为每次循环运行时都不需要检索 `context.workbook.worksheets.getActiveWorksheet()` `for` `for` 此值。 仅应在该循环中检索在循环上下文中更改的值。

```js
// This code sample shows a good performance strategy.
async function run() {
  await Excel.run(async (context) => {
    var ranges = [];
    // Retrieve the worksheet outside the loop.
    var worksheet = context.workbook.worksheets.getActiveWorksheet(); 

    // Only process the necessary values inside the loop.
    for (let i = 0; i < 7500; i++) {
      var rangeByIndex = worksheet.getRangeByIndexes(i, 1, 1, 1);
    }    
    await context.sync();
  });
}
```

### <a name="strategy-2-create-fewer-range-objects"></a>策略 2：创建更少的 range 对象

创建更少的 range 对象以提高性能并最小化有效负载大小。 以下文章部分和代码示例介绍了两种创建较少的 range 对象的方法。

#### <a name="split-each-range-array-into-multiple-arrays"></a>将每个区域数组拆分为多个数组

创建较少的 range 对象的一种方式是，将每个区域数组拆分为多个数组，然后使用循环和新调用处理每个新 `context.sync()` 数组。

> [!IMPORTANT]
> 仅在首次确定超出有效负载请求大小限制时使用此策略。 使用多个循环可以减少每个有效负载请求的大小，以避免超出 5MB 的限制，但使用多个循环和多个调用 `context.sync()` 也会对性能产生负面影响。

下面的代码示例尝试在一个循环中处理一个大型区域数组，然后处理一个 `context.sync()` 调用。 在一次调用中处理过多的范围 `context.sync()` 值会导致有效负载请求大小超过 5MB 限制。

```js
// This code sample does not show a recommended strategy.
// Calling 10,000 rows would likely exceed the 5MB payload size limit in a real-world situation.
async function run() {
  await Excel.run(async (context) => {
    var worksheet = context.workbook.worksheets.getActiveWorksheet();
    
    // This sample attempts to process too many ranges at once. 
    for (let row = 1; row < 10000; row++) {
      var range = sheet.getRangeByIndexes(row, 1, 1, 1);
      range.values = [["1"]];
    }
    await context.sync(); 
  });
}
```

下面的代码示例演示了类似于前面的代码示例的逻辑，但具有避免超过 5 MB 有效负载请求大小限制的策略。 在下面的代码示例中，范围在两个单独的循环中进行处理，每个循环后跟一个 `context.sync()` 调用。

```js
// This code sample shows a strategy for reducing payload request size.
// However, using multiple loops and `context.sync()` calls negatively impacts performance.
// Only use this strategy if you've determined that you're exceeding the payload request limit.
async function run() {
  await Excel.run(async (context) => {
    var worksheet = context.workbook.worksheets.getActiveWorksheet();

    // Split the ranges into two loops, rows 1-5000 and then 5001-10000.
    for (let row = 1; row < 5000; row++) {
      var range = worksheet.getRangeByIndexes(row, 1, 1, 1);
      range.values = [["1"]];
    }
    // Sync after each loop. 
    await context.sync(); 
    
    for (let row = 5001; row < 10000; row++) {
      var range = worksheet.getRangeByIndexes(row, 1, 1, 1);
      range.values = [["1"]];
    }
    await context.sync(); 
  });
}
```

#### <a name="set-range-values-in-an-array"></a>设置数组中的区域值

创建较少的 range 对象的另一种方式是创建一个数组，使用循环设置该数组中的所有数据，然后将数组值传递到一个范围。 这有利于性能和有效负载大小。 不是在 `range.values` 循环中调用每个区域， `range.values` 而是在循环外调用一次。

下面的代码示例演示如何创建数组、在循环中设置该数组的值，然后将数组值传递到循环 `for` 外部的范围。

```js
// This code sample shows a good performance strategy.
async function run() {
  await Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getActiveWorksheet();    
    // Create an array.
    const array = new Array(10000);

    // Set the values of the array inside the loop.
    for (var i = 0; i < 10000; i++) {
      array[i] = [1];
    }

    // Pass the array values to a range outside the loop. 
    var range = worksheet.getRange("A1:A10000");
    range.values = array;
    await context.sync();
  });
}
```

## <a name="see-also"></a>另请参阅

* [Excel 加载项中的 Word JavaScript 对象模型](excel-add-ins-core-concepts.md)
* [JavaScript API Excel错误处理](excel-add-ins-error-handling.md)
* [Office 外接程序的资源限制和性能优化](../concepts/resource-limits-and-performance-optimization.md)
* [工作表函数对象（适用于 Excel 的 JavaScript API）](/javascript/api/excel/excel.functions)
