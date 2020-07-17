---
title: Excel JavaScript API 性能优化
description: 使用 Excel JavaScript API 优化性能
ms.date: 07/14/2020
localization_priority: Normal
ms.openlocfilehash: 193cbe8c8cd1a432c6567401ed645990cb93e5e9
ms.sourcegitcommit: 472b81642e9eb5fb2a55cd98a7b0826d37eb7f73
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/17/2020
ms.locfileid: "45159092"
---
# <a name="performance-optimization-using-the-excel-javascript-api"></a>使用 Excel JavaScript API 优化性能

有多种方法可以使用 Excel JavaScript API 执行常见任务。 你将发现不同方法之间的显著性能差异。 本文提供指导和代码示例，展示如何使用 Excel JavaScript API 来高效执行常见任务。

## <a name="minimize-the-number-of-sync-calls"></a>减少 sync() 调用次数

在 Excel JavaScript API 中，`sync()` 是唯一的异步操作，在某些情况下可能会很慢，尤其是对于 Excel 网页版。 若要优化性能，在调用之前，通过尽可能多地将更改加入队列来最大程度减少调用 `sync()` 的次数。

有关按照此做法操作的代码示例，请参阅[核心概念 - sync()](excel-add-ins-core-concepts.md#sync)。

## <a name="minimize-the-number-of-proxy-objects-created"></a>最大程度减少创建的代理对象数目

避免重复创建同一个代理对象。 如果多个操作需要同一个代理对象，则改为创建一次并将其分配给一个变量，然后在代码中使用该变量。

```js
// BAD: repeated calls to .getRange() to create the same proxy object
worksheet.getRange("A1").format.fill.color = "red";
worksheet.getRange("A1").numberFormat = "0.00%";
worksheet.getRange("A1").values = [[1]];

// GOOD: create the range proxy object once and assign to a variable
var range = worksheet.getRange("A1")
range.format.fill.color = "red";
range.numberFormat = "0.00%";
range.values = [[1]];

// ALSO GOOD: use a "set" method to immediately set all the properties without even needing to create a variable!
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

## <a name="load-necessary-properties-only"></a>仅加载必要属性

在 Excel JavaScript API 中，需要显式加载代理对象的属性。 虽然可以使用空的 `load()` 调用一次性加载所有属性，但这种方法可能会产生大量的性能开销。 我们转为建议只加载必要的属性，特别是对于那些具有大量属性的对象。

例如，如果您只想读取 `address` range 对象的属性，请在调用方法时仅指定该属性 `load()` ：

```js
range.load('address');
```

您可以 `load()` 通过以下任一方式调用方法：

_语法：_

```js
object.load(string: properties);
// or
object.load(array: properties);
// or
object.load({ loadOption });
```

_其中：_

* `properties` 列出了要加载的属性，指定为逗号分隔的字符串或名称数组。 有关详细信息，请参阅 `load()` 为[EXCEL JavaScript API 参考](../reference/overview/excel-add-ins-reference-overview.md)中的对象定义的方法。
* `loadOption` 指定的对象描述了选择、展开、置顶和跳过选项。有关详细信息，请参阅对象加载[选项](/javascript/api/office/officeextension.loadoption)。

请注意，对象下的某些 "属性" 可能与另一个对象具有相同的名称。 例如，`format` 是区域对象下的一个属性，但 `format` 本身也是一个对象。 因此，如果发出 `range.load("format")` 之类的调用，这就相当于 `range.format.load()`，后者是一个空 load() 调用，它可能会导致前面所述的性能问题。 若要避免这种情况，代码应仅加载对象树中的 "叶节点"。

## <a name="suspend-excel-processes-temporarily"></a>暂时挂起 Excel 进程

Excel 中的多个后台任务将反应来自用户和外接程序的输入。 可以控制其中的部分 Excel 进程以提高性能。 这在外接程序处理大型数据集时尤其有用。

### <a name="suspend-calculation-temporarily"></a>暂停计算

如果你试图在大量单元格上执行操作（例如，设置一个大范围对象的值），而且不介意在操作完成时暂停 Excel 中的计算，建议暂停计算，直到调用下一个 `context.sync()`。

有关如何使用 `suspendApiCalculationUntilNextSync()` API 以便捷的方式暂停和重新激活计算的信息，请参阅[应用程序对象](/javascript/api/excel/excel.application)参考文档。 下面的代码演示了如何暂停计算：

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

请注意，只有公式计算才会被挂起。 仍将重新生成任何已更改的引用。 例如，重命名工作表仍会将公式中的任何引用更新到该工作表。

### <a name="suspend-screen-updating"></a>暂停屏幕更新

Excel 大约会在代码发生更改时显示外接程序所进行的这些更改。 对于大型迭代数据集，你无需实时在屏幕上查看此进度。 在外接程序调用 `context.sync()` 或者在 `Excel.run` 结束（隐式调用 `context.sync`）之前，`Application.suspendScreenUpdatingUntilNextSync()` 将暂停对 Excel 的可视化更新。 请注意，在下次同步之前，Excel 不会显示任何活动迹象。你的外接程序应为用户提供相关指南，以便为此延迟做好准备，或者提供一个状态栏，以演示相关活动。

> [!NOTE]
> 请勿 `suspendScreenUpdatingUntilNextSync` 重复调用（如在循环中）。 重复调用将导致 Excel 窗口闪烁。

### <a name="enable-and-disable-events"></a>启用和禁用事件

可以通过禁用事件来改进加载项性能。 [使用事件](excel-add-ins-events.md#enable-and-disable-events)文章中的代码示例展示了如何启用和禁用事件。

## <a name="importing-data-into-tables"></a>将数据导入表

当试图将大量数据直接导入到 [Table](/javascript/api/excel/excel.table) 对象中时（例如，通过使用 `TableRowCollection.add()`），可能会遇到性能缓慢的问题。 如果尝试添加一个新表，应首先通过设置 `range.values` 来填充数据，然后调用 `worksheet.tables.add()` 在该区域内创建一个表。 如果尝试将数据写入现有表，请通过 `table.getDataBodyRange()` 将数据写入一个 range 对象，表将自动展开。 

下面是此方法的一个示例：

```js
Excel.run(async (ctx) => {
    var sheet = ctx.workbook.worksheets.getItem("Sheet1");
    // Write the data into the range first 
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

## <a name="untrack-unneeded-ranges"></a>取消跟踪不需要的区域

JavaScript 层为加载项创建代理对象，以便与 Excel 工作簿和基础区域交互。 这些对象将一直保存在内存中，直到调用 `context.sync()`。 大型批处理操作可能会生成许多代理对象，加载项只需用到这些对象一次，并且可以在批处理执行之前从内存中释放。

[Range.untrack()](/javascript/api/excel/excel.range#untrack--) 方法从内存中释放 Excel Range 对象。 在加载项处理完区域后调用此方法，应会在使用大量 Range 对象时产生明显的性能优势。

> [!NOTE]
> `Range.untrack()` 是 [ClientRequestContext.trackedObjects.remove(thisRange)](/javascript/api/office/officeextension.trackedobjects#remove-object-) 的快捷方式。 任何代理对象都可以通过从上下文中的跟踪对象列表中删除它来取消跟踪。 通常情况下，Range 对象是数量充足的用来证明取消跟踪合理性的惟一 Excel 对象。

下面的代码示例用数据填充选定区域，每次填充一个单元格。 将值添加到单元格后，表示该单元格的区域将被取消跟踪。 在选定的 10,000 到 20,000 个单元格区域运行此代码，首先使用 `cell.untrack()` 行，然后取消使用。 应会注意到，使用 `cell.untrack()` 行的代码比不使用的代码运行速度要快。 此外，可能还会注意到之后的响应时间更快，因为清理步骤花费的时间更少。

```js
Excel.run(async (context) => {
    var largeRange = context.workbook.getSelectedRange();
    largeRange.load(["rowCount", "columnCount"]);
    await context.sync();

    for (var i = 0; i < largeRange.rowCount; i++) {
        for (var j = 0; j < largeRange.columnCount; j++) {
            var cell = largeRange.getCell(i, j);
            cell.values = [[i *j]];

            // call untrack() to release the range from memory
            cell.untrack();
        }
    }

    await context.sync();
});
```

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 基本编程概念](excel-add-ins-core-concepts.md)
- [Excel JavaScript API 高级编程概念](excel-add-ins-advanced-concepts.md)
- [Office 外接程序的资源限制和性能优化](../concepts/resource-limits-and-performance-optimization.md)
- [工作表函数对象（适用于 Excel 的 JavaScript API）](/javascript/api/excel/excel.functions)
