---
title: Excel JavaScript API 性能优化
description: 使用 JavaScript API 优化 Excel 加载项性能。
ms.date: 07/29/2020
localization_priority: Normal
ms.openlocfilehash: 42ab5f28717f0f7dcd06461840de692a5daf60ce
ms.sourcegitcommit: c6308cf245ac1bc66a876eaa0a7bb4a2492991ac
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/08/2020
ms.locfileid: "47408612"
---
# <a name="performance-optimization-using-the-excel-javascript-api"></a><span data-ttu-id="31ca6-103">使用 Excel JavaScript API 优化性能</span><span class="sxs-lookup"><span data-stu-id="31ca6-103">Performance optimization using the Excel JavaScript API</span></span>

<span data-ttu-id="31ca6-104">有多种方法可以使用 Excel JavaScript API 执行常见任务。</span><span class="sxs-lookup"><span data-stu-id="31ca6-104">There are multiple ways that you can perform common tasks with the Excel JavaScript API.</span></span> <span data-ttu-id="31ca6-105">你将发现不同方法之间的显著性能差异。</span><span class="sxs-lookup"><span data-stu-id="31ca6-105">You'll find significant performance differences between various approaches.</span></span> <span data-ttu-id="31ca6-106">本文提供指导和代码示例，展示如何使用 Excel JavaScript API 来高效执行常见任务。</span><span class="sxs-lookup"><span data-stu-id="31ca6-106">This article provides guidance and code samples to show you how to perform common tasks efficiently using Excel JavaScript API.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="31ca6-107">可以通过推荐使用和呼叫解决许多性能问题 `load` `sync` 。</span><span class="sxs-lookup"><span data-stu-id="31ca6-107">Many performance issues can be addressed through recommended usage of `load` and `sync` calls.</span></span> <span data-ttu-id="31ca6-108">请参阅 [Office 外接程序的资源限制和性能优化](../concepts/resource-limits-and-performance-optimization.md#performance-improvements-with-the-application-specific-apis) 一节中的 "特定于应用程序的 Api 的性能改进" 一节，以高效的方式使用应用程序特定的 api 的建议。</span><span class="sxs-lookup"><span data-stu-id="31ca6-108">See the "Performance improvements with the application-specific APIs" section of [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md#performance-improvements-with-the-application-specific-apis) for advice on working with the application-specific APIs in an efficient way.</span></span>

## <a name="suspend-excel-processes-temporarily"></a><span data-ttu-id="31ca6-109">暂时挂起 Excel 进程</span><span class="sxs-lookup"><span data-stu-id="31ca6-109">Suspend Excel processes temporarily</span></span>

<span data-ttu-id="31ca6-110">Excel 中的多个后台任务将反应来自用户和外接程序的输入。</span><span class="sxs-lookup"><span data-stu-id="31ca6-110">Excel has a number of background tasks reacting to input from both users and your add-in.</span></span> <span data-ttu-id="31ca6-111">可以控制其中的部分 Excel 进程以提高性能。</span><span class="sxs-lookup"><span data-stu-id="31ca6-111">Some of these Excel processes can be controlled to yield a performance benefit.</span></span> <span data-ttu-id="31ca6-112">这在外接程序处理大型数据集时尤其有用。</span><span class="sxs-lookup"><span data-stu-id="31ca6-112">This is especially helpful when your add-in deals with large data sets.</span></span>

### <a name="suspend-calculation-temporarily"></a><span data-ttu-id="31ca6-113">暂停计算</span><span class="sxs-lookup"><span data-stu-id="31ca6-113">Suspend calculation temporarily</span></span>

<span data-ttu-id="31ca6-114">如果你试图在大量单元格上执行操作（例如，设置一个大范围对象的值），而且不介意在操作完成时暂停 Excel 中的计算，建议暂停计算，直到调用下一个 `context.sync()`。</span><span class="sxs-lookup"><span data-stu-id="31ca6-114">If you are trying to perform an operation on a large number of cells (for example, setting the value of a huge range object) and you don't mind suspending the calculation in Excel temporarily while your operation finishes, we recommend that you suspend calculation until the next `context.sync()` is called.</span></span>

<span data-ttu-id="31ca6-115">有关如何使用 `suspendApiCalculationUntilNextSync()` API 以便捷的方式暂停和重新激活计算的信息，请参阅[应用程序对象](/javascript/api/excel/excel.application)参考文档。</span><span class="sxs-lookup"><span data-stu-id="31ca6-115">See the [Application Object](/javascript/api/excel/excel.application) reference documentation for information about how to use the `suspendApiCalculationUntilNextSync()` API to suspend and reactivate calculations in a very convenient way.</span></span> <span data-ttu-id="31ca6-116">下面的代码演示了如何暂停计算：</span><span class="sxs-lookup"><span data-stu-id="31ca6-116">The following code demonstrates how to suspend calculation temporarily:</span></span>

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

<span data-ttu-id="31ca6-117">请注意，只有公式计算才会被挂起。</span><span class="sxs-lookup"><span data-stu-id="31ca6-117">Please note that only formula calculations are suspended.</span></span> <span data-ttu-id="31ca6-118">仍将重新生成任何已更改的引用。</span><span class="sxs-lookup"><span data-stu-id="31ca6-118">Any altered references are still rebuilt.</span></span> <span data-ttu-id="31ca6-119">例如，重命名工作表仍会将公式中的任何引用更新到该工作表。</span><span class="sxs-lookup"><span data-stu-id="31ca6-119">For example, renaming a worksheet still updates any references in formulas to that worksheet.</span></span>

### <a name="suspend-screen-updating"></a><span data-ttu-id="31ca6-120">暂停屏幕更新</span><span class="sxs-lookup"><span data-stu-id="31ca6-120">Suspend screen updating</span></span>

<span data-ttu-id="31ca6-121">Excel 大约会在代码发生更改时显示外接程序所进行的这些更改。</span><span class="sxs-lookup"><span data-stu-id="31ca6-121">Excel displays changes your add-in makes approximately as they happen in the code.</span></span> <span data-ttu-id="31ca6-122">对于大型迭代数据集，你无需实时在屏幕上查看此进度。</span><span class="sxs-lookup"><span data-stu-id="31ca6-122">For large, iterative data sets, you may not need to see this progress on the screen in real-time.</span></span> <span data-ttu-id="31ca6-123">在外接程序调用 `context.sync()` 或者在 `Excel.run` 结束（隐式调用 `context.sync`）之前，`Application.suspendScreenUpdatingUntilNextSync()` 将暂停对 Excel 的可视化更新。</span><span class="sxs-lookup"><span data-stu-id="31ca6-123">`Application.suspendScreenUpdatingUntilNextSync()` pauses visual updates to Excel until the add-in calls `context.sync()`, or until `Excel.run` ends (implicitly calling `context.sync`).</span></span> <span data-ttu-id="31ca6-124">请注意，在下次同步之前，Excel 不会显示任何活动迹象。你的外接程序应为用户提供相关指南，以便为此延迟做好准备，或者提供一个状态栏，以演示相关活动。</span><span class="sxs-lookup"><span data-stu-id="31ca6-124">Be aware, Excel will not show any signs of activity until the next sync. Your add-in should either give users guidance to prepare them for this delay or provide a status bar to demonstrate activity.</span></span>

> [!NOTE]
> <span data-ttu-id="31ca6-125">请勿 `suspendScreenUpdatingUntilNextSync` 反复调用 (如在循环) 中。</span><span class="sxs-lookup"><span data-stu-id="31ca6-125">Don't call `suspendScreenUpdatingUntilNextSync` repeatedly (such as in a loop).</span></span> <span data-ttu-id="31ca6-126">重复调用将导致 Excel 窗口闪烁。</span><span class="sxs-lookup"><span data-stu-id="31ca6-126">Repeated calls will cause the Excel window to flicker.</span></span>

### <a name="enable-and-disable-events"></a><span data-ttu-id="31ca6-127">启用和禁用事件</span><span class="sxs-lookup"><span data-stu-id="31ca6-127">Enable and disable events</span></span>

<span data-ttu-id="31ca6-128">可以通过禁用事件来改进加载项性能。</span><span class="sxs-lookup"><span data-stu-id="31ca6-128">Performance of an add-in may be improved by disabling events.</span></span> <span data-ttu-id="31ca6-129">[使用事件](excel-add-ins-events.md#enable-and-disable-events)文章中的代码示例展示了如何启用和禁用事件。</span><span class="sxs-lookup"><span data-stu-id="31ca6-129">A code sample showing how to enable and disable events is in the [Work with Events](excel-add-ins-events.md#enable-and-disable-events) article.</span></span>

## <a name="importing-data-into-tables"></a><span data-ttu-id="31ca6-130">将数据导入表</span><span class="sxs-lookup"><span data-stu-id="31ca6-130">Importing data into tables</span></span>

<span data-ttu-id="31ca6-131">当试图将大量数据直接导入到 [Table](/javascript/api/excel/excel.table) 对象中时（例如，通过使用 `TableRowCollection.add()`），可能会遇到性能缓慢的问题。</span><span class="sxs-lookup"><span data-stu-id="31ca6-131">When trying to import a huge amount of data directly into a [Table](/javascript/api/excel/excel.table) object directly (for example, by using `TableRowCollection.add()`), you might experience slow performance.</span></span> <span data-ttu-id="31ca6-132">如果尝试添加一个新表，应首先通过设置 `range.values` 来填充数据，然后调用 `worksheet.tables.add()` 在该区域内创建一个表。</span><span class="sxs-lookup"><span data-stu-id="31ca6-132">If you are trying to add a new table, you should fill in the data first by setting `range.values`, and then call `worksheet.tables.add()` to create a table over the range.</span></span> <span data-ttu-id="31ca6-133">如果尝试将数据写入现有表，请通过 `table.getDataBodyRange()` 将数据写入一个 range 对象，表将自动展开。</span><span class="sxs-lookup"><span data-stu-id="31ca6-133">If you are trying to write data into an existing table, write the data into a range object via `table.getDataBodyRange()`, and the table will expand automatically.</span></span>

<span data-ttu-id="31ca6-134">下面是此方法的一个示例：</span><span class="sxs-lookup"><span data-stu-id="31ca6-134">Here is an example of this approach:</span></span>

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
> <span data-ttu-id="31ca6-135">可以使用 [Table.convertToRange()](/javascript/api/excel/excel.table#converttorange--) 方法将 Table 对象转换为 Range 对象，此做法非常方便。</span><span class="sxs-lookup"><span data-stu-id="31ca6-135">You can conveniently convert a Table object to a Range object by using the [Table.convertToRange()](/javascript/api/excel/excel.table#converttorange--) method.</span></span>

## <a name="see-also"></a><span data-ttu-id="31ca6-136">另请参阅</span><span class="sxs-lookup"><span data-stu-id="31ca6-136">See also</span></span>

* [<span data-ttu-id="31ca6-137">Office 外接程序中的 Excel JavaScript 对象模型</span><span class="sxs-lookup"><span data-stu-id="31ca6-137">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
* [<span data-ttu-id="31ca6-138">Office 外接程序的资源限制和性能优化</span><span class="sxs-lookup"><span data-stu-id="31ca6-138">Resource limits and performance optimization for Office Add-ins</span></span>](../concepts/resource-limits-and-performance-optimization.md)
* [<span data-ttu-id="31ca6-139">工作表函数对象（适用于 Excel 的 JavaScript API）</span><span class="sxs-lookup"><span data-stu-id="31ca6-139">Worksheet Functions Object (JavaScript API for Excel)</span></span>](/javascript/api/excel/excel.functions)
