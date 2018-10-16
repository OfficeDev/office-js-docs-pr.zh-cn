---
title: Excel JavaScript API 性能优化
description: 使用 Excel JavaScript API 优化性能
ms.date: 03/28/2018
ms.openlocfilehash: ee1687fcb1a5db74e65f5e73994653df235b4823
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/12/2018
ms.locfileid: "25505375"
---
# <a name="performance-optimization-using-the-excel-javascript-api"></a><span data-ttu-id="f68c7-103">使用 Excel JavaScript API 进行的性能优化</span><span class="sxs-lookup"><span data-stu-id="f68c7-103">Performance optimization using the Excel JavaScript API</span></span>

<span data-ttu-id="f68c7-p101">有多种你可以用 Excel JavaScript API 执行常见任务的方法。你将发现各个方法之间的显著的性能差异。本文提供了指导和代码示例，以向你展示如何使用 Excel JavaScript API 有效地执行常见任务。</span><span class="sxs-lookup"><span data-stu-id="f68c7-p101">There are multiple ways that you can perform common tasks with the Excel JavaScript API. You'll find significant performance differences between various approaches. This article provides guidance and code samples to show you how to perform common tasks efficiently using Excel JavaScript API.</span></span>

## <a name="minimize-the-number-of-sync-calls"></a><span data-ttu-id="f68c7-107">最小化 sync() 调用的数量</span><span class="sxs-lookup"><span data-stu-id="f68c7-107">Minimize the number of sync() calls</span></span>

<span data-ttu-id="f68c7-p102">在 Excel JavaScript API 中，```sync()``` 是唯一的异步操作，并且在某些情况下可能会很慢，尤其是对于 Excel Online。为了优化性能，通过尽可能先将多个更改排入队列再调用来最小化 ```sync()``` 的调用数量。</span><span class="sxs-lookup"><span data-stu-id="f68c7-p102">In the Excel JavaScript API, ```sync()``` is the only asynchronous operation, and it can be slow under some circumstances, especially for Excel Online. To optimize performance, minimize the number of calls to ```sync()``` by queueing up as many changes as possible before calling it.</span></span>

<span data-ttu-id="f68c7-110">参阅 [核心概念 - sync()](excel-add-ins-core-concepts.md#sync) 了解遵循这种做法的代码示例。</span><span class="sxs-lookup"><span data-stu-id="f68c7-110">See [Core Concepts - sync()](excel-add-ins-core-concepts.md#sync) for code samples that follow this practice.</span></span>

## <a name="minimize-the-number-of-proxy-objects-created"></a><span data-ttu-id="f68c7-111">最小化创建的代理对象的数量</span><span class="sxs-lookup"><span data-stu-id="f68c7-111">Minimize the number of proxy objects created</span></span>

<span data-ttu-id="f68c7-p103">请避免重复创建相同的代理对象。相反，如果需要对多个操作使用相同的代理对象，请仅创建一次并将其分配给一个变量，然后在你的代码中使用此变量。</span><span class="sxs-lookup"><span data-stu-id="f68c7-p103">Avoid repeatedly creating the same proxy object. Instead, if you need the same proxy object for more than one operation, create it once and assign it to a variable, then use that variable in your code.</span></span>

```javascript
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

## <a name="load-necessary-properties-only"></a><span data-ttu-id="f68c7-114">仅加载必要的属性</span><span class="sxs-lookup"><span data-stu-id="f68c7-114">Load necessary properties only</span></span>

<span data-ttu-id="f68c7-p104">在 Excel JavaScript API 中，你需要显式加载代理对象的属性。尽管用一个空的 ```load()``` 调用可以同时加载所有属性，但是此方法可能具有严重的性能开销。相反，我们建议你仅加载必要的属性，尤其是对于那些具有大量的属性的对象。</span><span class="sxs-lookup"><span data-stu-id="f68c7-p104">In the Excel JavaScript API, you need to explicitly load the properties of a proxy object. Although you're able to load all the properties at once with an empty ```load()``` call, that approach can have significant performance overhead. Instead, we suggest that you only load the necessary properties, especially for those objects which have a large number of properties.</span></span>

<span data-ttu-id="f68c7-118">例如，如果只想读取范围对象的 **address** 属性，请在调用 **load()** 方法时仅指定此属性：</span><span class="sxs-lookup"><span data-stu-id="f68c7-118">For example, if you only intend to read back the **address** property of a range object, specify only that property when you call the **load()** method:</span></span>
 
```js
range.load('address');
```
 
<span data-ttu-id="f68c7-119">可以通过以下任意方式调用 **load()** 方法：</span><span class="sxs-lookup"><span data-stu-id="f68c7-119">You can call **load()** method in any of the following ways:</span></span>
 
<span data-ttu-id="f68c7-120">_句法：_</span><span class="sxs-lookup"><span data-stu-id="f68c7-120">_Syntax:_</span></span>
 
```js
object.load(string: properties);
// or
object.load(array: properties);
// or
object.load({ loadOption });
```
 
<span data-ttu-id="f68c7-121">_其中：_</span><span class="sxs-lookup"><span data-stu-id="f68c7-121">_Where:_</span></span>
 
* <span data-ttu-id="f68c7-p105">`properties` 是要加载的属性列表，指定为以逗号分隔的字符串或为名称的数组。欲知详细信息，请参阅 [Excel JavaScript API 参考](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)中为对象定义的 **load()** 方法。</span><span class="sxs-lookup"><span data-stu-id="f68c7-p105">`properties` is the list of properties to load, specified as comma-delimited strings or as an array of names. For more information, see the **load()** methods defined for objects in [Excel JavaScript API reference](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview).</span></span>
* <span data-ttu-id="f68c7-p106">`loadOption` 指定描述选择、展开、置顶和跳过选项的对象。请参阅对象加载[选项](https://docs.microsoft.com/javascript/api/office/officeextension.loadoption)了解详细信息。</span><span class="sxs-lookup"><span data-stu-id="f68c7-p106">`loadOption` specifies an object that describes the selection, expansion, top, and skip options. See object load [options](https://docs.microsoft.com/javascript/api/office/officeextension.loadoption) for details.</span></span>

<span data-ttu-id="f68c7-p107">请注意，一个对象下的一些“属性”可能与另一个对象有相同的名称。例如，`format` 是 range 对象下的一个属性，但 `format` 本身也是一个对象。因此，如果进行像 `range.load("format")` 这样的调用，则此属性等同于 `range.format.load()`，即可能导致如前所述的性能问题的空的 load() 调用。要避免此问题，你的代码应仅加载对象树中的“叶节点”。</span><span class="sxs-lookup"><span data-stu-id="f68c7-p107">Please be aware that some of the “properties” under an object may have the same name as another object. For example, `format` is a property under range object, but `format` itself is an object as well. So, if you make a call such as `range.load("format")`, this is equivalent to `range.format.load()`, which is an empty load() call that can cause performance problems as outlined previously. To avoid this, your code should only load the “leaf nodes” in an object tree.</span></span> 

## <a name="suspend-calculation-temporarily"></a><span data-ttu-id="f68c7-130">临时暂停计算</span><span class="sxs-lookup"><span data-stu-id="f68c7-130">Suspend calculation temporarily</span></span>

<span data-ttu-id="f68c7-131">如果试图对大量单元格执行操作（例如，设置大 range 对象的值），并且不介意在操作完成时临时暂停 Excel 中的计算，则我们建议你暂停计算直到调用下一个 ```context.sync()```。</span><span class="sxs-lookup"><span data-stu-id="f68c7-131">If you are trying to perform an operation on a large number of cells (for example, setting the value of a huge range object) and you don't mind suspending the calculation in Excel temporarily while your operation finishes, we recommend that you suspend calculation until the next ```context.sync()``` is called.</span></span>

<span data-ttu-id="f68c7-p108">请参阅 [Application 对象](https://docs.microsoft.com/javascript/api/excel/excel.application)参考文档了解有关如何以非常便捷的方式使用 ```suspendApiCalculationUntilNextSync()``` API 暂停和重新激活计算等信息。下面的代码演示如何暂时暂停计算：</span><span class="sxs-lookup"><span data-stu-id="f68c7-p108">See [Application Object](https://docs.microsoft.com/javascript/api/excel/excel.application) reference documentation for information about how to use the ```suspendApiCalculationUntilNextSync()``` API to suspend and reactivate calculations in a very convenient way. The following code demonstrates how to suspend calculation temporarily:</span></span>

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

    // Suspending recalc
    app.suspendApiCalculationUntilNextSync();
    rangeToSet = sheet.getRange("A1:B1");
    rangeToSet.values = [[10, 20]];
    rangeToGet = sheet.getRange("A1:C1");
    rangeToGet.load("values");
    app.load("calculationMode");
    await ctx.sync();
    // Range value should be [10, 20, 3] when we load the property, because calculation is suspended at that point
    console.log(rangeToGet.values);
    // Calculation mode should still be "Automatic" even with supend recalc
    console.log(app.calculationMode);

    rangeToGet.load("values");
    await ctx.sync();
    // Range value should be [10, 20, 30] when we load the property, because calculation is resumed after last sync
    console.log(rangeToGet.values);
})
```

## <a name="update-all-cells-in-a-range"></a><span data-ttu-id="f68c7-134">更新区域中的所有单元格</span><span class="sxs-lookup"><span data-stu-id="f68c7-134">Update all cells in a range</span></span> 

<span data-ttu-id="f68c7-p109">当需要更新具有相同值或属性的区域中的所有单元格时，通过重复指定相同值的二维数组这样做可能很慢，因为该方法要求 Excel 循环访问区域中的所有单元格以独立设置每一个单元格。Excel 有更高效的方法来更新具有相同值或属性的区域中的所有单元格。</span><span class="sxs-lookup"><span data-stu-id="f68c7-p109">When you need to update all cells in a range with the same value or property, it can be slow to do this via a 2-dimensional array that repeatedly specifies the same value, since that approach requires Excel to iterate over all of the cells in the range to set each one separately. Excel has a more efficient way to update all the cells in a range with the same value or property.</span></span>

<span data-ttu-id="f68c7-p110">如果需要将相同的值、相同的数字格式或同一公式应用到单元格区域，则指定单个值更高效，而不是值的数组。这样做会显著提高性能。欲知在操作中显示此方法的代码示例，请参阅[核心概念 - 更新区域中所有单元格](excel-add-ins-core-concepts.md#update-all-cells-in-a-range)。</span><span class="sxs-lookup"><span data-stu-id="f68c7-p110">If you need to apply the same value, the same number format, or the same formula to a range of cells, it's more efficient to specify a single value instead of an array of values. Doing so will significantly improve performance. For a code sample that shows this approach in action, see [Core concepts - Update all cells in a range](excel-add-ins-core-concepts.md#update-all-cells-in-a-range).</span></span>

<span data-ttu-id="f68c7-p111">可以应用此方法的常见方案是当在工作表中的不同列上设置不同的数值格式时。在这种情况下，可以只循环访问列并用单个值在每个列上设置数值格式。处理作为区域的每个列，如[更新区域中的所有单元格](excel-add-ins-core-concepts.md#update-all-cells-in-a-range)代码示例中所示。</span><span class="sxs-lookup"><span data-stu-id="f68c7-p111">A common scenario where you can apply this approach is when setting different number formats on different columns in a worksheet. In this case, you can simply iterate through the columns and set the number format on each column with a single value. Handle each column as a range, as shown in the [Update all cells in a range](excel-add-ins-core-concepts.md#update-all-cells-in-a-range) code sample.</span></span>

> [!NOTE]
> <span data-ttu-id="f68c7-p112">如果你正在使用 TypeScript，你会注意到一个编译错误称单个值无法设置为二维数组。这是不可避免的，因为当检索属性时，值 *是* 二维数组，并且 TypeScript 不允许不同 setter vs getter 类型。但是，简单的替代方法是设置值用 `as any` 后缀，例如，`range.values = "hello world" as any`。</span><span class="sxs-lookup"><span data-stu-id="f68c7-p112">If you're using TypeScript, you will notice a compile error saying that a single value cannot be set to a 2D array.  This is unavoidable since the values *are* a 2D array when retrieving the properties, and TypeScript does not allow different setter vs getter types.  However, a simple workaround is to set the values with a `as any` suffix, e.g., `range.values = "hello world" as any`.</span></span>

## <a name="importing-data-into-tables"></a><span data-ttu-id="f68c7-146">将数据导入表</span><span class="sxs-lookup"><span data-stu-id="f68c7-146">Importing data into tables</span></span>

<span data-ttu-id="f68c7-p113">当尝试将大量数据直接导入到一个 [Table](https://docs.microsoft.com/javascript/api/excel/excel.table) 对象（例如，通过使用 `TableRowCollection.add()`），可能会遇到性能缓慢。如果试图添加一个新表，则应首先通过设置 `range.values` 填入数据，然后调用 `worksheet.tables.add()` 以在区域内创建表。如果试图将数据写入现有表，请通过 `table.getDataBodyRange()` 将数据写入一个 range 对象，该表就会自动展开。</span><span class="sxs-lookup"><span data-stu-id="f68c7-p113">When trying to import a huge amount of data directly into a [Table](https://docs.microsoft.com/javascript/api/excel/excel.table) object directly (for example, by using `TableRowCollection.add()`), you might experience slow performance. If you are trying to add a new table, you should fill in the data first by setting `range.values`, and then call `worksheet.tables.add()` to create a table over the range. If you are trying to write data into an existing table, write the data into a range object via `table.getDataBodyRange()`, and the table will expand automatically.</span></span> 

<span data-ttu-id="f68c7-150">以下是此方法的示例：</span><span class="sxs-lookup"><span data-stu-id="f68c7-150">Here is an example in JavaScript of this operation.</span></span>

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
> <span data-ttu-id="f68c7-151">您可以使用 [Table.convertToRange()](https://docs.microsoft.com/javascript/api/excel/excel.table#converttorange--) 方法，方便地将表对象转换为范围对象。</span><span class="sxs-lookup"><span data-stu-id="f68c7-151">You can conveniently convert a Table object to a Range object by using the [Table.convertToRange()](https://docs.microsoft.com/javascript/api/excel/excel.table#converttorange--) method.</span></span>

## <a name="enable-and-disable-events"></a><span data-ttu-id="f68c7-152">启用和禁用事件</span><span class="sxs-lookup"><span data-stu-id="f68c7-152">Enable and disable events</span></span>

<span data-ttu-id="f68c7-p114">通过禁用事件可以提高加载项的性能。显示如何启用和禁用事件的代码示例在[使用事件](excel-add-ins-events.md#enable-and-disable-events)一文中。</span><span class="sxs-lookup"><span data-stu-id="f68c7-p114">Performance of an add-in may be improved by disabling events. A code sample showing how to enable and disable events is in the [Work with Events](excel-add-ins-events.md#enable-and-disable-events) article.</span></span>

## <a name="see-also"></a><span data-ttu-id="f68c7-155">另请参阅</span><span class="sxs-lookup"><span data-stu-id="f68c7-155">See also</span></span>

- [<span data-ttu-id="f68c7-156">使用 Excel JavaScript API 的基本编程概念</span><span class="sxs-lookup"><span data-stu-id="f68c7-156">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="f68c7-157">使用 Excel JavaScript API 的高级编程概念</span><span class="sxs-lookup"><span data-stu-id="f68c7-157">Advanced programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-advanced-concepts.md)
- [<span data-ttu-id="f68c7-158">Excel JavaScript API 开放性规范</span><span class="sxs-lookup"><span data-stu-id="f68c7-158">Excel JavaScript API Open Specification</span></span>](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)
- [<span data-ttu-id="f68c7-159">工作表函数对象（适用于 Excel 的 JavaScript API）</span><span class="sxs-lookup"><span data-stu-id="f68c7-159">Worksheet Functions Object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.functions)
