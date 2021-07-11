---
title: Excel JavaScript API 性能优化
description: 使用 javaScript API Excel优化加载项性能。
ms.date: 07/29/2020
localization_priority: Normal
ms.openlocfilehash: 5313bb3fe25d165e49cc0508e81d58294db48798
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349383"
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
> 可以使用 [Table.convertToRange()](/javascript/api/excel/excel.table#converttorange--) 方法将 Table 对象转换为 Range 对象，此做法非常方便。

## <a name="see-also"></a>另请参阅

* [Excel 加载项中的 Word JavaScript 对象模型](excel-add-ins-core-concepts.md)
* [Office 外接程序的资源限制和性能优化](../concepts/resource-limits-and-performance-optimization.md)
* [工作表函数对象（适用于 Excel 的 JavaScript API）](/javascript/api/excel/excel.functions)
