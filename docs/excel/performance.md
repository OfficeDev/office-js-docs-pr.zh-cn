---
title: Excel JavaScript API 性能优化
description: 使用 Excel JavaScript API 优化性能
ms.date: 06/20/2019
localization_priority: Priority
ms.openlocfilehash: 7ffca45d1d7f6619fa1d942befed1856f77934d9
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/30/2019
ms.locfileid: "36695845"
---
# <a name="performance-optimization-using-the-excel-javascript-api"></a><span data-ttu-id="4eb1b-103">使用 Excel JavaScript API 优化性能</span><span class="sxs-lookup"><span data-stu-id="4eb1b-103">Performance optimization using the Excel JavaScript API</span></span>

<span data-ttu-id="4eb1b-104">有多种方法可以使用 Excel JavaScript API 执行常见任务。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-104">There are multiple ways that you can perform common tasks with the Excel JavaScript API.</span></span> <span data-ttu-id="4eb1b-105">你将发现不同方法之间的显著性能差异。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-105">You'll find significant performance differences between various approaches.</span></span> <span data-ttu-id="4eb1b-106">本文提供指导和代码示例，展示如何使用 Excel JavaScript API 来高效执行常见任务。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-106">This article provides guidance and code samples to show you how to perform common tasks efficiently using Excel JavaScript API.</span></span>

## <a name="minimize-the-number-of-sync-calls"></a><span data-ttu-id="4eb1b-107">减少 sync() 调用次数</span><span class="sxs-lookup"><span data-stu-id="4eb1b-107">Minimize the number of sync() calls</span></span>

<span data-ttu-id="4eb1b-108">在 Excel JavaScript API 中，```sync()``` 是唯一的异步操作，在某些情况下可能会很慢，尤其是对于 Excel 网页版。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-108">In the Excel JavaScript API, ```sync()``` is the only asynchronous operation, and it can be slow under some circumstances, especially for Excel Online.</span></span> <span data-ttu-id="4eb1b-109">若要优化性能，在调用之前，通过尽可能多地将更改加入队列来最大程度减少调用 ```sync()``` 的次数。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-109">To optimize performance, minimize the number of calls to ```sync()``` by queueing up as many changes as possible before calling it.</span></span>

<span data-ttu-id="4eb1b-110">有关按照此做法操作的代码示例，请参阅[核心概念 - sync()](excel-add-ins-core-concepts.md#sync)。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-110">See [Core Concepts - sync()](excel-add-ins-core-concepts.md#sync) for code samples that follow this practice.</span></span>

## <a name="minimize-the-number-of-proxy-objects-created"></a><span data-ttu-id="4eb1b-111">最大程度减少创建的代理对象数目</span><span class="sxs-lookup"><span data-stu-id="4eb1b-111">Minimize the number of proxy objects created</span></span>

<span data-ttu-id="4eb1b-112">避免重复创建同一个代理对象。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-112">Avoid repeatedly creating the same proxy object.</span></span> <span data-ttu-id="4eb1b-113">如果多个操作需要同一个代理对象，则改为创建一次并将其分配给一个变量，然后在代码中使用该变量。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-113">Instead, if you need the same proxy object for more than one operation, create it once and assign it to a variable, then use that variable in your code.</span></span>

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

## <a name="load-necessary-properties-only"></a><span data-ttu-id="4eb1b-114">仅加载必要属性</span><span class="sxs-lookup"><span data-stu-id="4eb1b-114">Load necessary properties only</span></span>

<span data-ttu-id="4eb1b-115">在 Excel JavaScript API 中，需要显式加载代理对象的属性。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-115">In the Excel JavaScript API, you need to explicitly load the properties of a proxy object.</span></span> <span data-ttu-id="4eb1b-116">虽然可以使用空的 ```load()``` 调用一次性加载所有属性，但这种方法可能会产生大量的性能开销。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-116">Although you're able to load all the properties at once with an empty ```load()``` call, that approach can have significant performance overhead.</span></span> <span data-ttu-id="4eb1b-117">我们转为建议只加载必要的属性，特别是对于那些具有大量属性的对象。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-117">Instead, we suggest that you only load the necessary properties, especially for those objects which have a large number of properties.</span></span>

<span data-ttu-id="4eb1b-118">例如，如果你只想读取区域对象的 **address** 属性，则在调用 **load()** 方法时仅指定该属性：</span><span class="sxs-lookup"><span data-stu-id="4eb1b-118">For example, if you only intend to read the **address** property of a range object, specify only that property when you call the **load()** method:</span></span>

```js
range.load('address');
```

<span data-ttu-id="4eb1b-119">可以通过以下任意方式调用 **load()** 方法：</span><span class="sxs-lookup"><span data-stu-id="4eb1b-119">You can call **load()** method in any of the following ways:</span></span>

<span data-ttu-id="4eb1b-120">_语法：_</span><span class="sxs-lookup"><span data-stu-id="4eb1b-120">_Syntax:_</span></span>

```js
object.load(string: properties);
// or
object.load(array: properties);
// or
object.load({ loadOption });
```

<span data-ttu-id="4eb1b-121">_其中：_</span><span class="sxs-lookup"><span data-stu-id="4eb1b-121">_Where:_</span></span>

* <span data-ttu-id="4eb1b-122">`properties` 列出了要加载的属性，指定为逗号分隔的字符串或名称数组。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-122">`properties` is the list of properties to load, specified as comma-delimited strings or as an array of names.</span></span> <span data-ttu-id="4eb1b-123">有关详细信息，请参阅 [Excel JavaScript API 参考](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)中为对象定义的 **load()** 方法。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-123">For more information, see the **load()** methods defined for objects in [Excel JavaScript API reference](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview).</span></span>
* <span data-ttu-id="4eb1b-p106">`loadOption` 指定的对象描述了选择、展开、置顶和跳过选项。有关详细信息，请参阅对象加载[选项](/javascript/api/office/officeextension.loadoption)。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-p106">`loadOption` specifies an object that describes the selection, expansion, top, and skip options. See object load [options](/javascript/api/office/officeextension.loadoption) for details.</span></span>

<span data-ttu-id="4eb1b-126">请注意，一个对象下的某些“属性”可能与另一个对象同名。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-126">Please be aware that some of the “properties” under an object may have the same name as another object.</span></span> <span data-ttu-id="4eb1b-127">例如，`format` 是区域对象下的一个属性，但 `format` 本身也是一个对象。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-127">For example, `format` is a property under range object, but `format` itself is an object as well.</span></span> <span data-ttu-id="4eb1b-128">因此，如果发出 `range.load("format")` 之类的调用，这就相当于 `range.format.load()`，后者是一个空 load() 调用，它可能会导致前面所述的性能问题。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-128">So, if you make a call such as `range.load("format")`, this is equivalent to `range.format.load()`, which is an empty load() call that can cause performance problems as outlined previously.</span></span> <span data-ttu-id="4eb1b-129">若要避免这种情况，代码应仅加载对象树中的“叶节点”。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-129">To avoid this, your code should only load the “leaf nodes” in an object tree.</span></span> 

## <a name="suspend-excel-processes-temporarily"></a><span data-ttu-id="4eb1b-130">暂时挂起 Excel 进程</span><span class="sxs-lookup"><span data-stu-id="4eb1b-130">Suspend Excel processes temporarily</span></span>

<span data-ttu-id="4eb1b-131">Excel 中的多个后台任务将反应来自用户和外接程序的输入。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-131">Excel has a number of background tasks reacting to input from both users and your add-in.</span></span> <span data-ttu-id="4eb1b-132">可以控制其中的部分 Excel 进程以提高性能。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-132">Some of these Excel processes can be controlled to yield a performance benefit.</span></span> <span data-ttu-id="4eb1b-133">这在外接程序处理大型数据集时尤其有用。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-133">This is especially helpful when your add-in deals with large data sets.</span></span>

### <a name="suspend-calculation-temporarily"></a><span data-ttu-id="4eb1b-134">暂停计算</span><span class="sxs-lookup"><span data-stu-id="4eb1b-134">Suspend calculation temporarily</span></span>

<span data-ttu-id="4eb1b-135">如果你试图在大量单元格上执行操作（例如，设置一个大范围对象的值），而且不介意在操作完成时暂停 Excel 中的计算，建议暂停计算，直到调用下一个 `context.sync()`。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-135">If you are trying to perform an operation on a large number of cells (for example, setting the value of a huge range object) and you don't mind suspending the calculation in Excel temporarily while your operation finishes, we recommend that you suspend calculation until the next `context.sync()` is called.</span></span>

<span data-ttu-id="4eb1b-136">有关如何使用 `suspendApiCalculationUntilNextSync()` API 以便捷的方式暂停和重新激活计算的信息，请参阅[应用程序对象](/javascript/api/excel/excel.application)参考文档。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-136">See the [Application Object](/javascript/api/excel/excel.application) reference documentation for information about how to use the `suspendApiCalculationUntilNextSync()` API to suspend and reactivate calculations in a very convenient way.</span></span> <span data-ttu-id="4eb1b-137">下面的代码演示了如何暂停计算：</span><span class="sxs-lookup"><span data-stu-id="4eb1b-137">The following code demonstrates how to suspend calculation temporarily:</span></span>

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

### <a name="suspend-screen-updating"></a><span data-ttu-id="4eb1b-138">暂停屏幕更新</span><span class="sxs-lookup"><span data-stu-id="4eb1b-138">Suspend screen updating</span></span>

<span data-ttu-id="4eb1b-139">Excel 大约会在代码发生更改时显示外接程序所进行的这些更改。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-139">Excel displays changes your add-in makes approximately as they happen in the code.</span></span> <span data-ttu-id="4eb1b-140">对于大型迭代数据集，你无需实时在屏幕上查看此进度。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-140">For large, iterative data sets, you may not need to see this progress on the screen in real-time.</span></span> <span data-ttu-id="4eb1b-141">在外接程序调用 `context.sync()` 或者在 `Excel.run` 结束（隐式调用 `context.sync`）之前，`Application.suspendScreenUpdatingUntilNextSync()` 将暂停对 Excel 的可视化更新。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-141">`Application.suspendScreenUpdatingUntilNextSync()` pauses visual updates to Excel until the add-in calls `context.sync()`, or until `Excel.run` ends (implicitly calling `context.sync`).</span></span> <span data-ttu-id="4eb1b-142">请注意，在下次同步之前，Excel 不会显示任何活动迹象。你的外接程序应为用户提供相关指南，以便为此延迟做好准备，或者提供一个状态栏，以演示相关活动。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-142">Be aware, Excel will not show any signs of activity until the next sync. Your add-in should either give users guidance to prepare them for this delay or provide a status bar to demonstrate activity.</span></span>

### <a name="enable-and-disable-events"></a><span data-ttu-id="4eb1b-143">启用和禁用事件</span><span class="sxs-lookup"><span data-stu-id="4eb1b-143">Enable and disable events</span></span>

<span data-ttu-id="4eb1b-144">可以通过禁用事件来改进加载项性能。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-144">Performance of an add-in may be improved by disabling events.</span></span> <span data-ttu-id="4eb1b-145">[使用事件](excel-add-ins-events.md#enable-and-disable-events)文章中的代码示例展示了如何启用和禁用事件。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-145">A code sample showing how to enable and disable events is in the [Work with Events](excel-add-ins-events.md#enable-and-disable-events) article.</span></span>

## <a name="update-all-cells-in-a-range"></a><span data-ttu-id="4eb1b-146">更新区域中的所有单元格</span><span class="sxs-lookup"><span data-stu-id="4eb1b-146">Update all cells in a range</span></span>

<span data-ttu-id="4eb1b-147">当你需要更新区域中具有相同值或属性的所有单元格，通过重复指定相同值的二维数组来实现此操作可能会比较慢，因为此方法需要 Excel 遍历区域内的所有单元格，以分别设置每个单元格。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-147">When you need to update all cells in a range with the same value or property, it can be slow to do this via a 2-dimensional array that repeatedly specifies the same value, since that approach requires Excel to iterate over all of the cells in the range to set each one separately.</span></span> <span data-ttu-id="4eb1b-148">Excel 有一种更有效的方法来更新区域内具有相同值或属性的所有单元格。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-148">Excel has a more efficient way to update all the cells in a range with the same value or property.</span></span>

<span data-ttu-id="4eb1b-149">如果需要对一个区域内的单元格应用相同值、相同数字格式或相同公式，指定单个值比指定一组值更高效。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-149">If you need to apply the same value, the same number format, or the same formula to a range of cells, it's more efficient to specify a single value instead of an array of values.</span></span> <span data-ttu-id="4eb1b-150">此操作能够显著提高性能。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-150">Doing so will significantly improve performance.</span></span> <span data-ttu-id="4eb1b-151">有关显示此方法实际运行的代码示例，请参阅[核心概念 - 更新区域内的所有单元格](excel-add-ins-core-concepts.md#update-all-cells-in-a-range)。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-151">For a code sample that shows this approach in action, see [Core concepts - Update all cells in a range](excel-add-ins-core-concepts.md#update-all-cells-in-a-range).</span></span>

<span data-ttu-id="4eb1b-152">可以应用此方法的一个常见场景是，在工作表的不同列上设置不同数字格式。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-152">A common scenario where you can apply this approach is when setting different number formats on different columns in a worksheet.</span></span> <span data-ttu-id="4eb1b-153">在此情况下，只需遍历列，并在每个列上用单个值设置数字格式。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-153">In this case, you can simply iterate through the columns and set the number format on each column with a single value.</span></span> <span data-ttu-id="4eb1b-154">将每一列作为一个区域处理，如[更新区域中的所有单元格](excel-add-ins-core-concepts.md#update-all-cells-in-a-range)代码示例中所示。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-154">Handle each column as a range, as shown in the [Update all cells in a range](excel-add-ins-core-concepts.md#update-all-cells-in-a-range) code sample.</span></span>

> [!NOTE]
> <span data-ttu-id="4eb1b-155">如果使用 TypeScript，你会注意到一个编译错误，指示不能将单个值设置为二维数组。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-155">If you're using TypeScript, you will notice a compile error saying that a single value cannot be set to a 2D array.</span></span>  <span data-ttu-id="4eb1b-156">这是不可避免的，因为在检索属性时，这些值*是*一个二维数组，且 TypeScript 不允许不同的 setter 和 getter 类型。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-156">This is unavoidable since the values *are* a 2D array when retrieving the properties, and TypeScript does not allow different setter vs getter types.</span></span>  <span data-ttu-id="4eb1b-157">但是，一个简单的解决方法是使用 `as any` 后缀设置值，例如 `range.values = "hello world" as any`。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-157">However, a simple workaround is to set the values with a `as any` suffix, e.g., `range.values = "hello world" as any`.</span></span>

## <a name="importing-data-into-tables"></a><span data-ttu-id="4eb1b-158">将数据导入表</span><span class="sxs-lookup"><span data-stu-id="4eb1b-158">Importing data into tables</span></span>

<span data-ttu-id="4eb1b-159">当试图将大量数据直接导入到 [Table](/javascript/api/excel/excel.table) 对象中时（例如，通过使用 `TableRowCollection.add()`），可能会遇到性能缓慢的问题。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-159">When trying to import a huge amount of data directly into a [Table](/javascript/api/excel/excel.table) object directly (for example, by using `TableRowCollection.add()`), you might experience slow performance.</span></span> <span data-ttu-id="4eb1b-160">如果尝试添加一个新表，应首先通过设置 `range.values` 来填充数据，然后调用 `worksheet.tables.add()` 在该区域内创建一个表。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-160">If you are trying to add a new table, you should fill in the data first by setting `range.values`, and then call `worksheet.tables.add()` to create a table over the range.</span></span> <span data-ttu-id="4eb1b-161">如果尝试将数据写入现有表，请通过 `table.getDataBodyRange()` 将数据写入一个 range 对象，表将自动展开。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-161">If you are trying to write data into an existing table, write the data into a range object via `table.getDataBodyRange()`, and the table will expand automatically.</span></span> 

<span data-ttu-id="4eb1b-162">下面是此方法的一个示例：</span><span class="sxs-lookup"><span data-stu-id="4eb1b-162">Here is an example of this approach:</span></span>

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
> <span data-ttu-id="4eb1b-163">可以使用 [Table.convertToRange()](/javascript/api/excel/excel.table#converttorange--) 方法将 Table 对象转换为 Range 对象，此做法非常方便。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-163">You can conveniently convert a Table object to a Range object by using the [Table.convertToRange()](/javascript/api/excel/excel.table#converttorange--) method.</span></span>

## <a name="untrack-unneeded-ranges"></a><span data-ttu-id="4eb1b-164">取消跟踪不需要的区域</span><span class="sxs-lookup"><span data-stu-id="4eb1b-164">Untrack unneeded ranges</span></span>

<span data-ttu-id="4eb1b-165">JavaScript 层为加载项创建代理对象，以便与 Excel 工作簿和基础区域交互。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-165">The JavaScript layer creates proxy objects for your add-in to interact with the Excel workbook and underlying ranges.</span></span> <span data-ttu-id="4eb1b-166">这些对象将一直保存在内存中，直到调用 `context.sync()`。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-166">These objects persist in memory until `context.sync()` is called.</span></span> <span data-ttu-id="4eb1b-167">大型批处理操作可能会生成许多代理对象，加载项只需用到这些对象一次，并且可以在批处理执行之前从内存中释放。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-167">Large batch operations may generate a lot of proxy objects that are only needed once by the add-in and can be released from memory before the batch executes.</span></span>

<span data-ttu-id="4eb1b-168">[Range.untrack()](/javascript/api/excel/excel.range#untrack--) 方法从内存中释放 Excel Range 对象。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-168">The [Range.untrack()](/javascript/api/excel/excel.range#untrack--) method releases an Excel Range object from memory.</span></span> <span data-ttu-id="4eb1b-169">在加载项处理完区域后调用此方法，应会在使用大量 Range 对象时产生明显的性能优势。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-169">Calling this method after your add-in is done with the range should yield a noticeable performance benefit when using large numbers of Range objects.</span></span>

> [!NOTE]
> <span data-ttu-id="4eb1b-170">`Range.untrack()` 是 [ClientRequestContext.trackedObjects.remove(thisRange)](/javascript/api/office/officeextension.trackedobjects#remove-object-) 的快捷方式。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-170">`Range.untrack()` is a shortcut for [ClientRequestContext.trackedObjects.remove(thisRange)](/javascript/api/office/officeextension.trackedobjects#remove-object-).</span></span> <span data-ttu-id="4eb1b-171">任何代理对象都可以通过从上下文中的跟踪对象列表中删除它来取消跟踪。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-171">Any proxy object can be untracked by removing it from the tracked objects list in the context.</span></span> <span data-ttu-id="4eb1b-172">通常情况下，Range 对象是数量充足的用来证明取消跟踪合理性的惟一 Excel 对象。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-172">Typically, Range objects are the only Excel objects used in sufficient quantity to justify untracking.</span></span>

<span data-ttu-id="4eb1b-173">下面的代码示例用数据填充选定区域，每次填充一个单元格。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-173">The following code sample fills a selected range with data, one cell at a time.</span></span> <span data-ttu-id="4eb1b-174">将值添加到单元格后，表示该单元格的区域将被取消跟踪。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-174">After the value is added to the cell, the range representing that cell is untracked.</span></span> <span data-ttu-id="4eb1b-175">在选定的 10,000 到 20,000 个单元格区域运行此代码，首先使用 `cell.untrack()` 行，然后取消使用。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-175">Run this code with a selected range of 10,000 to 20,000 cells, first with the `cell.untrack()` line, and then without it.</span></span> <span data-ttu-id="4eb1b-176">应会注意到，使用 `cell.untrack()` 行的代码比不使用的代码运行速度要快。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-176">You should notice the code runs faster with the `cell.untrack()` line than without it.</span></span> <span data-ttu-id="4eb1b-177">此外，可能还会注意到之后的响应时间更快，因为清理步骤花费的时间更少。</span><span class="sxs-lookup"><span data-stu-id="4eb1b-177">You may also notice a quicker response time afterwards, since the cleanup step takes less time.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="4eb1b-178">另请参阅</span><span class="sxs-lookup"><span data-stu-id="4eb1b-178">See also</span></span>

- [<span data-ttu-id="4eb1b-179">Excel JavaScript API 基本编程概念</span><span class="sxs-lookup"><span data-stu-id="4eb1b-179">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="4eb1b-180">Excel JavaScript API 高级编程概念</span><span class="sxs-lookup"><span data-stu-id="4eb1b-180">Advanced programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-advanced-concepts.md)
- [<span data-ttu-id="4eb1b-181">Office 外接程序的资源限制和性能优化</span><span class="sxs-lookup"><span data-stu-id="4eb1b-181">Resource limits and performance optimization for Office Add-ins</span></span>](../concepts/resource-limits-and-performance-optimization.md)
- [<span data-ttu-id="4eb1b-182">Excel JavaScript API 开放性规范</span><span class="sxs-lookup"><span data-stu-id="4eb1b-182">Excel JavaScript API Open Specification</span></span>](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)
- [<span data-ttu-id="4eb1b-183">工作表函数对象（适用于 Excel 的 JavaScript API）</span><span class="sxs-lookup"><span data-stu-id="4eb1b-183">Worksheet Functions Object (JavaScript API for Excel)</span></span>](/javascript/api/excel/excel.functions)
