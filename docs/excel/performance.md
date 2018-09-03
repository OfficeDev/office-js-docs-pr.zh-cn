---
title: Excel JavaScript API 性能优化
description: Excel JavaScript API 性能优化
ms.date: 03/28/2018
ms.openlocfilehash: 50fac999093abb3fbfe1bd5be1cd6a77dc930399
ms.sourcegitcommit: 78b28ae88d53bfef3134c09cc4336a5a8722c70b
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/01/2018
ms.locfileid: "23797313"
---
# <a name="performance-optimization-using-the-excel-javascript-api"></a>使用Excel JavaScript API进行的性能优化

您可以通过多种方式使用Excel JavaScript API执行常见任务。 你会发现各种方法之间的明显性能差异。 本文提供指导和代码示例，向您展示如何使用 Excel JavaScript API 高效地执行常见任务。

## <a name="minimize-the-number-of-sync-calls"></a>最小化同步()调用的数量

在Excel JavaScript API中，```sync()``` 是唯一的异步操作，在某些情况下可能会很慢，特别是对于Excel Online。 为了优化性能，请在调用之前尽可能多的队列更改，以此最小化调用 ```sync()```的次数。

参阅遵循这种做法的代码示例的 [核心概念 - 同步()](excel-add-ins-core-concepts.md#sync)。

## <a name="minimize-the-number-of-proxy-objects-created"></a>最小化创建的代理对象的数量

避免重复创建相同的代理对象。 相反，如果您需要对多个操作使用相同的代理对象，请仅创建一个对象并将其分配给一个变量，然后在您的代码中使用该变量。

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

## <a name="load-necessary-properties-only"></a>仅加载必要的属性

在Excel JavaScript API中，您需要显式加载代理对象的属性。 尽管您可以使用一个空白的 ```load()``` 调用一次性加载所有属性，但该方法可能会导致显着的性能负担。 相反，我们建议您只加载必要的属性，特别是对于具有大量属性的对象。

例如，如果您只想读取范围对象的**地址**属性，请在调用 **load()** 方法时仅指定该属性：
 
```js
range.load('address');
```
 
可以通过以下任意方式调用 **load()** 方法：
 
_语法：_
 
```js
object.load(string: properties);
// or
object.load(array: properties);
// or
object.load({ loadOption });
```
 
_其中：_
 
* `properties` 是要加载的属性列表，指定为以逗号分隔的字符串或名称数组。 有关详细信息，请参阅 [Excel JavaScript API 引用](https://docs.microsoft.com/javascript/office/overview/excel-add-ins-reference-overview)中为对象定义的 **load()** 方法。
* `loadOption` 指定描述选择、展开、置顶和跳过选项的对象。有关详细信息，请参阅对象加载[选项](https://docs.microsoft.com/javascript/api/office/officeextension.loadoption)。

请注意，一个对象下的某些“属性”可能与另一个对象具有相同的名称。 例如，`format` 是范围对象下的属性，但是`format` 本身也是一个对象。 所以，如果您进行了诸如 `range.load("format")`的调用，这相当于 `range.format.load()`，这是一个空的load()调用，可能会导致性能问题，如前所述。 为了避免这种情况，您的代码应该只加载对象树中的“叶节点”。 

## <a name="suspend-calculation-temporarily"></a>临时暂停计算

如果您试图对大量单元格执行操作（例如，设置大范围对象的值），并且您不介意在操作完成时临时暂停Excel中的计算，那么我们建议您暂停计算直到下一个 ```context.sync()``` 被调用。

参阅 [应用对象](https://docs.microsoft.com/javascript/api/excel/excel.application) 参考文档以获取有关如何使用 ```suspendApiCalculationUntilNextSync()``` API以非常方便的方式暂停和重新激活计算。 以下代码演示了如何临时暂停计算：

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

## <a name="update-all-cells-in-a-range"></a>更新区域中的所有单元格 

当您需要更新具有相同值或属性的范围中的所有单元格时，通过重复指定相同值的2维数组可能会很慢，因为该方法需要Excel遍历所有单元格范围以分别设置每一个。 Excel有一种更高效的方法来更新具有相同值或属性的范围内的所有单元格。

如果您需要将相同值，相同的数字格式或相同的公式应用于一个范围的单元格，则指定单个值而非值数组时效果更佳。 这样做会显著提高性能。 有关显示此方法的代码示例，请参阅 [核心概念 - 更新范围内的所有单元格](excel-add-ins-core-concepts.md#update-all-cells-in-a-range)。

您可以应用此方法的一种常见情况是在工作表中的不同列上设置不同的数字格式。 在这种情况下，您可以简单地遍历列并使用单个值设置每列的数字格式。 将每个列作为一个范围处理，如 [更新范围内的所有单元格](excel-add-ins-core-concepts.md#update-all-cells-in-a-range) 代码示例所示。

> [!NOTE]
> 如果您使用的是TypeScript，则会发现编译错误，指出无法将单个值设置为二维数组。  这是不可避免的，因为值 *是* 检索属性时使用二维数组，而TypeScript不允许使用不同的setter和getter类型。  但是，一个简单的解决方法是使用 `as any` 后缀来设置值，例如， `range.values = "hello world" as any`。

## <a name="importing-data-into-tables"></a>将数据导入表格

当您试图将大量数据直接导入到 [表](https://docs.microsoft.com/javascript/api/excel/excel.table) 对象时（例如，通过使用 `TableRowCollection.add()`），您可能会遇到性能下降。 如果您尝试添加一个新表，则应先通过设置 `range.values`填入数据，然后调用 `worksheet.tables.add()` 在范围内创建一个表格。 如果您尝试将数据写入现有表，请通过 `table.getDataBodyRange()`将数据写入范围对象，表会自动扩展。 

以下是此方法的示例：

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
> 您可以使用 [Table.convertToRange()](https://docs.microsoft.com/javascript/api/excel/excel.table#converttorange--) 方法，方便地将表对象转换为范围对象。

## <a name="enable-and-disable-events"></a>启用和禁用事件

可以通过禁用事件来提高加载项的性能。 在[使用事件](excel-add-ins-events.md#enable-and-disable-events)一文中有显示如何启用和禁用事件的代码示例。

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 核心概念](excel-add-ins-core-concepts.md)
- [Excel JavaScript API 高级概念](excel-add-ins-advanced-concepts.md)
- [Excel JavaScript API 开放性规范](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)
- [工作表函数对象（适用于 Excel 的 JavaScript API）](https://docs.microsoft.com/javascript/api/excel/excel.functions)
