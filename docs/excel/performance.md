---
title: Excel JavaScript API 性能优化
description: 使用 JavaScript API 优化 Excel 加载项性能。
ms.date: 02/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: bad5d35ec1cc3f99cd37b3571dee78d3432102e6
ms.sourcegitcommit: d8ea4b761f44d3227b7f2c73e52f0d2233bf22e2
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/11/2022
ms.locfileid: "66712725"
---
# <a name="performance-optimization-using-the-excel-javascript-api"></a>使用 Excel JavaScript API 优化性能

有多种方法可以使用 Excel JavaScript API 执行常见任务。 你将发现不同方法之间的显著性能差异。 本文提供指导和代码示例，展示如何使用 Excel JavaScript API 来高效执行常见任务。

> [!IMPORTANT]
> 可以通过建议使用和`sync`调用`load`来解决许多性能问题。 请参阅 [Office 外接程序的资源限制和性能优化](../concepts/resource-limits-and-performance-optimization.md#performance-improvements-with-the-application-specific-apis) 的“使用特定于应用程序的 API 的性能改进”部分，以获取有关以有效方式使用特定于应用程序的 API 的建议。

## <a name="suspend-excel-processes-temporarily"></a>暂时挂起 Excel 进程

Excel 中的多个后台任务将反应来自用户和外接程序的输入。 可以控制其中的部分 Excel 进程以提高性能。 这在外接程序处理大型数据集时尤其有用。

### <a name="suspend-calculation-temporarily"></a>暂停计算

如果你试图在大量单元格上执行操作（例如，设置一个大范围对象的值），而且不介意在操作完成时暂停 Excel 中的计算，建议暂停计算，直到调用下一个 `context.sync()`。

有关如何使用 `suspendApiCalculationUntilNextSync()` API 以便捷的方式暂停和重新激活计算的信息，请参阅[应用程序对象](/javascript/api/excel/excel.application)参考文档。 以下代码演示如何暂时暂停计算。

```js
await Excel.run(async (context) => {
    let app = context.workbook.application;
    let sheet = context.workbook.worksheets.getItem("sheet1");
    let rangeToSet: Excel.Range;
    let rangeToGet: Excel.Range;
    app.load("calculationMode");
    await context.sync();
    // Calculation mode should be "Automatic" by default
    console.log(app.calculationMode);

    rangeToSet = sheet.getRange("A1:C1");
    rangeToSet.values = [[1, 2, "=SUM(A1:B1)"]];
    rangeToGet = sheet.getRange("A1:C1");
    rangeToGet.load("values");
    await context.sync();
    // Range value should be [1, 2, 3] now
    console.log(rangeToGet.values);

    // Suspending recalculation
    app.suspendApiCalculationUntilNextSync();
    rangeToSet = sheet.getRange("A1:B1");
    rangeToSet.values = [[10, 20]];
    rangeToGet = sheet.getRange("A1:C1");
    rangeToGet.load("values");
    app.load("calculationMode");
    await context.sync();
    // Range value should be [10, 20, 3] when we load the property, because calculation is suspended at that point
    console.log(rangeToGet.values);
    // Calculation mode should still be "Automatic" even with suspend recalculation
    console.log(app.calculationMode);

    rangeToGet.load("values");
    await context.sync();
    // Range value should be [10, 20, 30] when we load the property, because calculation is resumed after last sync
    console.log(rangeToGet.values);
});
```

请注意，仅暂停公式计算。 仍会重新生成任何更改的引用。 例如，重命名工作表仍会更新公式中对该工作表的任何引用。

### <a name="suspend-screen-updating"></a>暂停屏幕更新

Excel 大约会在代码发生更改时显示外接程序所进行的这些更改。 对于大型迭代数据集，你无需实时在屏幕上查看此进度。 在外接程序调用 `context.sync()` 或者在 `Excel.run` 结束（隐式调用 `context.sync`）之前，`Application.suspendScreenUpdatingUntilNextSync()` 将暂停对 Excel 的可视化更新。 请注意，在下次同步之前，Excel 不会显示任何活动迹象。你的外接程序应为用户提供相关指南，以便为此延迟做好准备，或者提供一个状态栏，以演示相关活动。

> [!NOTE]
> 请勿重复调用 `suspendScreenUpdatingUntilNextSync` (，例如循环) 。 重复调用将导致 Excel 窗口闪烁。

### <a name="enable-and-disable-events"></a>启用和禁用事件

可以通过禁用事件来改进加载项性能。 [使用事件](excel-add-ins-events.md#enable-and-disable-events)文章中的代码示例展示了如何启用和禁用事件。

## <a name="importing-data-into-tables"></a>将数据导入表

当试图将大量数据直接导入到 [Table](/javascript/api/excel/excel.table) 对象中时（例如，通过使用 `TableRowCollection.add()`），可能会遇到性能缓慢的问题。 如果尝试添加一个新表，应首先通过设置 `range.values` 来填充数据，然后调用 `worksheet.tables.add()` 在该区域内创建一个表。 如果尝试将数据写入现有表，请通过 `table.getDataBodyRange()` 将数据写入一个 range 对象，表将自动展开。

下面是此方法的一个示例：

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sheet1");
    // Write the data into the range first.
    let range = sheet.getRange("A1:B3");
    range.values = [["Key", "Value"], ["A", 1], ["B", 2]];

    // Create the table over the range
    let table = sheet.tables.add('A1:B3', true);
    table.name = "Example";
    await context.sync();


    // Insert a new row to the table
    table.getDataBodyRange().getRowsBelow(1).values = [["C", 3]];
    // Change a existing row value
    table.getDataBodyRange().getRow(1).values = [["D", 4]];
    await context.sync();
});
```

> [!NOTE]
> 可以使用 [Table.convertToRange()](/javascript/api/excel/excel.table#excel-excel-table-converttorange-member(1)) 方法将 Table 对象转换为 Range 对象，此做法非常方便。

## <a name="payload-size-limit-best-practices"></a>有效负载大小限制最佳做法

Excel JavaScript API 对 API 调用有大小限制。 Excel web 版请求和响应的有效负载大小限制为 5MB，如果超出此限制，API 将返回`RichAPI.Error`错误。 在所有平台上，范围限制为 500 万个单元格以进行获取操作。 大范围通常超过这两个限制。

请求的有效负载大小是以下三个组件的组合。

* API 调用数
* 对象的数量，例如 `Range` 对象
* 要设置或获取的值的长度

如果 API 返回 `RequestPayloadSizeLimitExceeded` 错误，请使用本文中记录的最佳做法策略来优化脚本并避免错误。

### <a name="strategy-1-move-unchanged-values-out-of-loops"></a>策略 1：将未更改的值移出循环

限制循环中发生的进程数，以提高性能。 在下面的代码示例中`for`，`context.workbook.worksheets.getActiveWorksheet()`可以移出循环，因为它不会在该循环中更改。

```js
// DO NOT USE THIS CODE SAMPLE. This sample shows a poor performance strategy. 
async function run() {
  await Excel.run(async (context) => {
    let ranges = [];
    
    // This sample retrieves the worksheet every time the loop runs, which is bad for performance.
    for (let i = 0; i < 7500; i++) {
      let rangeByIndex = context.workbook.worksheets.getActiveWorksheet().getRangeByIndexes(i, 1, 1, 1);
    }    
    await context.sync();
  });
}
```

下面的代码示例显示了与前面的代码示例类似的逻辑，但改进了性能策略。 在循环之前`for`检索该值`context.workbook.worksheets.getActiveWorksheet()`，因为每次`for`循环运行时都不需要检索此值。 只有在循环上下文中发生更改的值才应在该循环中检索。

```js
// This code sample shows a good performance strategy.
async function run() {
  await Excel.run(async (context) => {
    let ranges = [];
    // Retrieve the worksheet outside the loop.
    let worksheet = context.workbook.worksheets.getActiveWorksheet(); 

    // Only process the necessary values inside the loop.
    for (let i = 0; i < 7500; i++) {
      let rangeByIndex = worksheet.getRangeByIndexes(i, 1, 1, 1);
    }    
    await context.sync();
  });
}
```

### <a name="strategy-2-create-fewer-range-objects"></a>策略 2：创建更少的范围对象

创建更少的范围对象来提高性能并最大程度地减少有效负载大小。 以下文章部分和代码示例介绍了创建较少范围对象的两种方法。

#### <a name="split-each-range-array-into-multiple-arrays"></a>将每个范围数组拆分为多个数组

创建更少范围对象的一种方法是将每个范围数组拆分为多个数组，然后使用循环和新调用处理每个新 `context.sync()` 数组。

> [!IMPORTANT]
> 仅当首次确定超出有效负载请求大小限制时，才使用此策略。 使用多个循环可以减小每个有效负载请求的大小以避免超过 5MB 限制，但使用多个循环和多个 `context.sync()` 调用也会对性能产生负面影响。

下面的代码示例尝试在单个循环和单 `context.sync()` 个调用中处理大型范围数组。 在一次 `context.sync()` 调用中处理过多的范围值会导致有效负载请求大小超过 5MB 限制。

```js
// This code sample does not show a recommended strategy.
// Calling 10,000 rows would likely exceed the 5MB payload size limit in a real-world situation.
async function run() {
  await Excel.run(async (context) => {
    let worksheet = context.workbook.worksheets.getActiveWorksheet();
    
    // This sample attempts to process too many ranges at once. 
    for (let row = 1; row < 10000; row++) {
      let range = sheet.getRangeByIndexes(row, 1, 1, 1);
      range.values = [["1"]];
    }
    await context.sync(); 
  });
}
```

下面的代码示例显示了与前面的代码示例类似的逻辑，但具有避免超过 5MB 有效负载请求大小限制的策略。 在下面的代码示例中，范围以两个单独的循环进行处理，每个循环后跟一个 `context.sync()` 调用。

```js
// This code sample shows a strategy for reducing payload request size.
// However, using multiple loops and `context.sync()` calls negatively impacts performance.
// Only use this strategy if you've determined that you're exceeding the payload request limit.
async function run() {
  await Excel.run(async (context) => {
    let worksheet = context.workbook.worksheets.getActiveWorksheet();

    // Split the ranges into two loops, rows 1-5000 and then 5001-10000.
    for (let row = 1; row < 5000; row++) {
      let range = worksheet.getRangeByIndexes(row, 1, 1, 1);
      range.values = [["1"]];
    }
    // Sync after each loop. 
    await context.sync(); 
    
    for (let row = 5001; row < 10000; row++) {
      let range = worksheet.getRangeByIndexes(row, 1, 1, 1);
      range.values = [["1"]];
    }
    await context.sync(); 
  });
}
```

#### <a name="set-range-values-in-an-array"></a>在数组中设置范围值

创建更少范围对象的另一种方法是创建数组，使用循环设置该数组中的所有数据，然后将数组值传递给某个区域。 这有利于性能和有效负载大小。 不是在循环中调用 `range.values` 每个范围， `range.values` 而是在循环外部调用一次。

下面的代码示例演示如何创建数组、在循环中 `for` 设置该数组的值，然后将数组值传递到循环外部的区域。

```js
// This code sample shows a good performance strategy.
async function run() {
  await Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getActiveWorksheet();    
    // Create an array.
    const array = new Array(10000);

    // Set the values of the array inside the loop.
    for (let i = 0; i < 10000; i++) {
      array[i] = [1];
    }

    // Pass the array values to a range outside the loop. 
    let range = worksheet.getRange("A1:A10000");
    range.values = array;
    await context.sync();
  });
}
```

## <a name="see-also"></a>另请参阅

* [Excel 加载项中的 Word JavaScript 对象模型](excel-add-ins-core-concepts.md)
* [使用特定于应用程序的 JavaScript API 进行错误处理](../testing/application-specific-api-error-handling.md)
* [Office 外接程序的资源限制和性能优化](../concepts/resource-limits-and-performance-optimization.md)
* [工作表函数对象（适用于 Excel 的 JavaScript API）](/javascript/api/excel/excel.functions)
