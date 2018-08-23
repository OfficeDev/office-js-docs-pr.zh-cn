---
title: Excel JavaScript API 性能优化
description: Excel JavaScript API 性能优化
ms.date: 03/28/2018
ms.openlocfilehash: dabbb69f8dee0df782a265edcfdfb1c89894e915
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/23/2018
ms.locfileid: "19437407"
---
# <a name="performance-optimization-using-the-excel-javascript-api"></a><span data-ttu-id="ea2b7-103">使用Excel JavaScript API进行的性能优化</span><span class="sxs-lookup"><span data-stu-id="ea2b7-103">Performance optimization using the Excel JavaScript API</span></span>

<span data-ttu-id="ea2b7-104">您可以通过多种方式使用Excel JavaScript API执行常见任务。</span><span class="sxs-lookup"><span data-stu-id="ea2b7-104">There are multiple ways that you can perform common tasks with the Excel JavaScript API.</span></span> <span data-ttu-id="ea2b7-105">你会发现各种方法之间的明显性能差异。</span><span class="sxs-lookup"><span data-stu-id="ea2b7-105">You'll find significant performance differences between various approaches.</span></span> <span data-ttu-id="ea2b7-106">本文提供指导和代码示例，向您展示如何使用Excel JavaScript API高效地执行常见任务。</span><span class="sxs-lookup"><span data-stu-id="ea2b7-106">This article provides code samples that show how to perform common tasks with worksheets using the Excel JavaScript API.</span></span>

## <a name="minimize-the-number-of-sync-calls"></a><span data-ttu-id="ea2b7-107">最小化同步()调用的数量</span><span class="sxs-lookup"><span data-stu-id="ea2b7-107">Minimize the number of sync() calls</span></span>

<span data-ttu-id="ea2b7-108">在Excel JavaScript API中，```sync()``` 是唯一的异步操作，在某些情况下可能会很慢，特别是对于Excel Online。</span><span class="sxs-lookup"><span data-stu-id="ea2b7-108">In the Excel JavaScript API, ```sync()``` is the only asynchronous operation, and it can be slow under some circumstances, especially for Excel Online.</span></span> <span data-ttu-id="ea2b7-109">为了优化性能，请在调用之前尽可能多的队列更改，以此最小化调用 ```sync()```的次数。</span><span class="sxs-lookup"><span data-stu-id="ea2b7-109">To optimize performance, minimize the number of calls to ```sync()``` by queueing up as many changes as possible before calling it.</span></span>

<span data-ttu-id="ea2b7-110">参阅遵循这种做法的代码示例的 [核心概念 - 同步()](excel-add-ins-core-concepts.md#sync)。</span><span class="sxs-lookup"><span data-stu-id="ea2b7-110">See [Core Concepts - sync()](excel-add-ins-core-concepts.md#sync) for code samples that follow this practice.</span></span>

## <a name="minimize-the-number-of-proxy-objects-created"></a><span data-ttu-id="ea2b7-111">最小化创建的代理对象的数量</span><span class="sxs-lookup"><span data-stu-id="ea2b7-111">Minimize the number of proxy objects created</span></span>

<span data-ttu-id="ea2b7-112">避免重复创建相同的代理对象。</span><span class="sxs-lookup"><span data-stu-id="ea2b7-112">Avoid repeatedly creating the same proxy object.</span></span> <span data-ttu-id="ea2b7-113">相反，如果您需要对多个操作使用相同的代理对象，请仅创建一个对象并将其分配给一个变量，然后在您的代码中使用该变量。</span><span class="sxs-lookup"><span data-stu-id="ea2b7-113">Instead, if you need the same proxy object for more than one operation, create it once and assign it to a variable, then use that variable in your code.</span></span>

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

## <a name="load-necessary-properties-only"></a><span data-ttu-id="ea2b7-114">仅加载必要的属性</span><span class="sxs-lookup"><span data-stu-id="ea2b7-114">Load necessary properties only</span></span>

<span data-ttu-id="ea2b7-115">在Excel JavaScript API中，您需要显式加载代理对象的属性。</span><span class="sxs-lookup"><span data-stu-id="ea2b7-115">In the Excel JavaScript API, you need to explicitly load the properties of a proxy object.</span></span> <span data-ttu-id="ea2b7-116">尽管您可以使用一个空白的 ```load()``` 调用一次性加载所有属性，但该方法可能会导致显着的性能负担。</span><span class="sxs-lookup"><span data-stu-id="ea2b7-116">Although you're able to load all the properties at once with an empty ```load()``` call, that approach can have significant performance overhead.</span></span> <span data-ttu-id="ea2b7-117">相反，我们建议您只加载必要的属性，特别是对于具有大量属性的对象。</span><span class="sxs-lookup"><span data-stu-id="ea2b7-117">Instead, we suggest that you only load the necessary properties, especially for those objects which have a large number of properties.</span></span>

<span data-ttu-id="ea2b7-118">相反，我们建议您只加载必要的属性，特别是对于具有大量属性的对象。**** ** **</span><span class="sxs-lookup"><span data-stu-id="ea2b7-118">For example, if you only intend to read back the **address** property of a range object, specify only that property when you call the **load()** method:</span></span>
 
```js
range.load('address');
```
 
<span data-ttu-id="ea2b7-119">您可以通过以下任意方式调用 **load()** 方法：</span><span class="sxs-lookup"><span data-stu-id="ea2b7-119">You can call **load()** method in any of the following ways:</span></span>
 
<span data-ttu-id="ea2b7-120">_语法：_</span><span class="sxs-lookup"><span data-stu-id="ea2b7-120">_Syntax:_</span></span>
 
```js
object.load(string: properties);
// or
object.load(array: properties);
// or
object.load({ loadOption });
```
 
<span data-ttu-id="ea2b7-121">_条件：_</span><span class="sxs-lookup"><span data-stu-id="ea2b7-121">_Where:_</span></span>
 
* <span data-ttu-id="ea2b7-122">`properties` 是要加载的属性列表，指定为以逗号分隔的字符串或名称数组。</span><span class="sxs-lookup"><span data-stu-id="ea2b7-122">`properties` is the list of properties and/or relationship names to be loaded specified as comma-delimited strings, or an array of names.</span></span> <span data-ttu-id="ea2b7-123">有关详细信息，请参阅 ** Excel JavaScript API 参考**  中为对象定义的 [ load()](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview) 方法。</span><span class="sxs-lookup"><span data-stu-id="ea2b7-123">For more information, see the **load()** methods defined for objects in [Excel JavaScript API reference](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview).</span></span>
* <span data-ttu-id="ea2b7-p106">`loadOption` 指定的对象描述了选择、展开、置顶和跳过选项。有关详细信息，请参阅对象加载 [选项](https://dev.office.com/reference/add-ins/excel/loadoption)。</span><span class="sxs-lookup"><span data-stu-id="ea2b7-p106">`loadOption` specifies an object that describes the selection, expansion, top, and skip options. See object load [options](https://dev.office.com/reference/add-ins/excel/loadoption) for details.</span></span>

<span data-ttu-id="ea2b7-126">请注意，一个对象下的某些“属性”可能与另一个对象具有相同的名称。</span><span class="sxs-lookup"><span data-stu-id="ea2b7-126">Please be aware that some of the “properties” under an object may have the same name as another object.</span></span> <span data-ttu-id="ea2b7-127">例如，`format` 是范围对象下的属性，但是`format` 本身也是一个对象。</span><span class="sxs-lookup"><span data-stu-id="ea2b7-127">For example, `format` is a property under range object, but `format` itself is an object as well.</span></span> <span data-ttu-id="ea2b7-128">所以，如果您进行了诸如 `range.load("format")`的调用，这相当于 `range.format.load()`，这是一个空的load()调用，可能会导致性能问题，如前所述。</span><span class="sxs-lookup"><span data-stu-id="ea2b7-128">So, if you make a call such as `range.load("format")`, this is equivalent to `range.format.load()`, which is an empty load() call that can cause performance problems as outlined previously.</span></span> <span data-ttu-id="ea2b7-129">为了避免这种情况，您的代码应该只加载对象树中的“叶节点”。</span><span class="sxs-lookup"><span data-stu-id="ea2b7-129">To avoid this, your code should only load the “leaf nodes” in an object tree.</span></span> 

## <a name="suspend-calculation-temporarily"></a><span data-ttu-id="ea2b7-130">临时暂停计算</span><span class="sxs-lookup"><span data-stu-id="ea2b7-130">Suspend calculation temporarily</span></span>

<span data-ttu-id="ea2b7-131">如果您试图对大量单元格执行操作（例如，设置大范围对象的值），并且您不介意在操作完成时临时暂停Excel中的计算，那么我们建议您暂停计算直到下一个 ```context.sync()``` 被调用。</span><span class="sxs-lookup"><span data-stu-id="ea2b7-131">If you are trying to perform an operation on a large number of cells (for example, setting the value of a huge range object) and you don't mind suspending the calculation in Excel temporarily while your operation finishes, we recommend that you suspend calculation until the next ```context.sync()``` is called.</span></span>

<span data-ttu-id="ea2b7-132">参阅 [应用对象](https://dev.office.com/reference/add-ins/excel/application) 参考文档以获取有关如何使用 ```suspendApiCalculationUntilNextSync()``` API以非常方便的方式暂停和重新激活计算。</span><span class="sxs-lookup"><span data-stu-id="ea2b7-132">See [Application Object](https://dev.office.com/reference/add-ins/excel/application) reference documentation for information about how to use the ```suspendApiCalculationUntilNextSync()``` API to suspend and reactivate calculations in a very convenient way.</span></span> <span data-ttu-id="ea2b7-133">以下代码演示了如何临时暂停计算：</span><span class="sxs-lookup"><span data-stu-id="ea2b7-133">The following code demonstrates how to suspend calculation temporarily:</span></span>

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

## <a name="update-all-cells-in-a-range"></a><span data-ttu-id="ea2b7-134">更新范围中的所有单元格</span><span class="sxs-lookup"><span data-stu-id="ea2b7-134">Update all cells in a range</span></span> 

<span data-ttu-id="ea2b7-135">当您需要更新具有相同值或属性的范围中的所有单元格时，通过重复指定相同值的2维数组可能会很慢，因为该方法需要Excel遍历所有单元格范围以分别设置每一个。</span><span class="sxs-lookup"><span data-stu-id="ea2b7-135">When you need to update all cells in a range with the same value or property, it can be slow to do this via a 2-dimensional array that repeatedly specifies the same value, since that approach requires Excel to iterate over all of the cells in the range to set each one separately.</span></span> <span data-ttu-id="ea2b7-136">Excel有一种更高效的方法来更新具有相同值或属性的范围内的所有单元格。</span><span class="sxs-lookup"><span data-stu-id="ea2b7-136">Excel has a more efficient way to update all the cells in a range with the same value or property.</span></span>

<span data-ttu-id="ea2b7-137">如果您需要将相同值，相同的数字格式或相同的公式应用于一个范围的单元格，则指定单个值而非值数组时效果更佳。</span><span class="sxs-lookup"><span data-stu-id="ea2b7-137">If you need to apply the same value, the same number format, or the same formula to a range of cells, it's more efficient to specify a single value instead of an array of values.</span></span> <span data-ttu-id="ea2b7-138">这样做会显著提高性能。</span><span class="sxs-lookup"><span data-stu-id="ea2b7-138">Doing so will significantly improve performance.</span></span> <span data-ttu-id="ea2b7-139">有关显示此方法的代码示例，请参阅 [核心概念 - 更新范围内的所有单元格](excel-add-ins-core-concepts.md#update-all-cells-in-a-range)。</span><span class="sxs-lookup"><span data-stu-id="ea2b7-139">For a code sample that shows this approach in action, see [Core concepts - Update all cells in a range](excel-add-ins-core-concepts.md#update-all-cells-in-a-range).</span></span>

<span data-ttu-id="ea2b7-140">您可以应用此方法的一种常见情况是在工作表中的不同列上设置不同的数字格式。</span><span class="sxs-lookup"><span data-stu-id="ea2b7-140">A common scenario where you can apply this approach is when setting different number formats on different columns in a worksheet.</span></span> <span data-ttu-id="ea2b7-141">在这种情况下，您可以简单地遍历列并使用单个值设置每列的数字格式。</span><span class="sxs-lookup"><span data-stu-id="ea2b7-141">In this case, you can simply iterate through the columns and set the number format on each column with a single value.</span></span> <span data-ttu-id="ea2b7-142">将每个列作为一个范围处理，如 [更新范围内的所有单元格](excel-add-ins-core-concepts.md#update-all-cells-in-a-range) 代码示例所示。</span><span class="sxs-lookup"><span data-stu-id="ea2b7-142">Handle each column as a range, as shown in the [Update all cells in a range](excel-add-ins-core-concepts.md#update-all-cells-in-a-range) code sample.</span></span>

> [!NOTE]
> <span data-ttu-id="ea2b7-143">如果您使用的是TypeScript，则会发现编译错误，指出无法将单个值设置为二维数组。</span><span class="sxs-lookup"><span data-stu-id="ea2b7-143">If you're using TypeScript, you will notice a compile error saying that a single value cannot be set to a 2D array.</span></span>  <span data-ttu-id="ea2b7-144">这是不可避免的，因为值 *是* 检索属性时使用二维数组，而TypeScript不允许使用不同的setter和getter类型。</span><span class="sxs-lookup"><span data-stu-id="ea2b7-144">This is unavoidable since the values *are* a 2D array when retrieving the properties, and TypeScript does not allow different setter vs getter types.</span></span>  <span data-ttu-id="ea2b7-145">但是，一个简单的解决方法是使用 `as any` 后缀来设置值，例如， `range.values = "hello world" as any`。</span><span class="sxs-lookup"><span data-stu-id="ea2b7-145">However, a simple workaround is to set the values with a `as any` suffix, e.g., `range.values = "hello world" as any`.</span></span>

## <a name="importing-data-into-tables"></a><span data-ttu-id="ea2b7-146">将数据导入表格</span><span class="sxs-lookup"><span data-stu-id="ea2b7-146">Importing data into tables</span></span>

<span data-ttu-id="ea2b7-147">当您试图将大量数据直接导入到 [表](https://dev.office.com/reference/add-ins/excel/table) 对象时（例如，通过使用 `TableRowCollection.add()`），您可能会遇到性能下降。</span><span class="sxs-lookup"><span data-stu-id="ea2b7-147">When trying to import a huge amount of data directly into a [Table](https://dev.office.com/reference/add-ins/excel/table) object directly (for example, by using `TableRowCollection.add()`), you might experience slow performance.</span></span> <span data-ttu-id="ea2b7-148">如果您尝试添加一个新表，则应先通过设置 `range.values`填入数据，然后调用 `worksheet.tables.add()` 在范围内创建一个表格。</span><span class="sxs-lookup"><span data-stu-id="ea2b7-148">If you are trying to add a new table, you should fill in the data first by setting `range.values`, and then call `worksheet.tables.add()` to create a table over the range.</span></span> <span data-ttu-id="ea2b7-149">如果您尝试将数据写入现有表，请通过 `table.getDataBodyRange()`将数据写入范围对象，表会自动扩展。</span><span class="sxs-lookup"><span data-stu-id="ea2b7-149">If you are trying to write data into an existing table, write the data into a range object via `table.getDataBodyRange()`, and the table will expand automatically.</span></span> 

<span data-ttu-id="ea2b7-150">这是此方法的一个例子：</span><span class="sxs-lookup"><span data-stu-id="ea2b7-150">Here is an example in JavaScript of this operation.</span></span>

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
> <span data-ttu-id="ea2b7-151">您可以使用 [Table.convertToRange()](https://dev.office.com/reference/add-ins/excel/table#converttorange) 方法，方便地将表对象转换为范围对象。</span><span class="sxs-lookup"><span data-stu-id="ea2b7-151">You can conveniently convert a Table object to a Range object by using the [Table.convertToRange()](https://dev.office.com/reference/add-ins/excel/table#converttorange) method.</span></span>

## <a name="see-also"></a><span data-ttu-id="ea2b7-152">另请参阅</span><span class="sxs-lookup"><span data-stu-id="ea2b7-152">See also</span></span>

- [<span data-ttu-id="ea2b7-153">Excel JavaScript API 核心概念</span><span class="sxs-lookup"><span data-stu-id="ea2b7-153">Excel JavaScript API core concepts</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="ea2b7-154">Excel JavaScript API 高级概念</span><span class="sxs-lookup"><span data-stu-id="ea2b7-154">Excel JavaScript API advanced concepts</span></span>](excel-add-ins-advanced-concepts.md)
- [<span data-ttu-id="ea2b7-155">Excel JavaScript API 开放性规范</span><span class="sxs-lookup"><span data-stu-id="ea2b7-155">Excel JavaScript API Open Specification</span></span>](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)
- [<span data-ttu-id="ea2b7-156">工作表函数对象（适用于 Excel 的 JavaScript API）</span><span class="sxs-lookup"><span data-stu-id="ea2b7-156">Worksheet Functions Object (JavaScript API for Excel)</span></span>](https://dev.office.com/reference/add-ins/excel/functions)
