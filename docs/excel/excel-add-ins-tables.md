---
title: 使用 Excel JavaScript API 处理表格
description: 显示如何使用 JavaScript API 对表执行常见Excel示例。
ms.date: 10/22/2021
ms.localizationpriority: medium
ms.openlocfilehash: f5ea4e12b4662c890259e29c52b98f1b16b9e5f6
ms.sourcegitcommit: 23ce57b2702aca19054e31fcb2d2f015b4183ba1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/02/2021
ms.locfileid: "60681158"
---
# <a name="work-with-tables-using-the-excel-javascript-api"></a>使用 Excel JavaScript API 处理表格

本文中的代码示例展示了如何使用 Excel JavaScript API 对表格执行常见任务。 有关 和 对象支持的属性和方法的完整列表，请参阅 Table Object (JavaScript API for Excel) 和 `Table` `TableCollection` [TableCollection Object (JavaScript API for Excel) 。 ](/javascript/api/excel/excel.tablecollection) [](/javascript/api/excel/excel.table)

## <a name="create-a-table"></a>创建表

下面的代码示例在名为 **Sample** 的工作表中创建一个表。 此表包含标题，并且包含四列和七行数据。 如果运行Excel的应用程序支持要求集 [](../reference/requirement-sets/excel-api-requirement-sets.md)**ExcelApi 1.2，** 则列宽和行高将设置为最适合表格中的当前数据。

> [!NOTE]
> 若要指定表的名称，必须先创建表，然后设置其属性 `name` ，如以下示例所示。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var expensesTable = sheet.tables.add("A1:D1", true /*hasHeaders*/);
    expensesTable.name = "ExpensesTable";

    expensesTable.getHeaderRowRange().values = [["Date", "Merchant", "Category", "Amount"]];

    expensesTable.rows.add(null /*add rows to the end of the table*/, [
        ["1/1/2017", "The Phone Company", "Communications", "$120"],
        ["1/2/2017", "Northwind Electric Cars", "Transportation", "$142"],
        ["1/5/2017", "Best For You Organics Company", "Groceries", "$27"],
        ["1/10/2017", "Coho Vineyard", "Restaurant", "$33"],
        ["1/11/2017", "Bellows College", "Education", "$350"],
        ["1/15/2017", "Trey Research", "Other", "$135"],
        ["1/15/2017", "Best For You Organics Company", "Groceries", "$97"]
    ]);

    if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
    }

    sheet.activate();

    return context.sync();
}).catch(errorHandlerFunction);
```

**新建表**

![Excel 中的新表。](../images/excel-tables-create.png)

## <a name="add-rows-to-a-table"></a>向表添加行

下面的代码示例将七个新行添加到名为 **Sample** 的工作表内的 **ExpensesTable** 表中。 新行被添加到表的末尾。 如果运行Excel的应用程序支持要求集 [](../reference/requirement-sets/excel-api-requirement-sets.md)**ExcelApi 1.2，** 则列宽和行高将设置为最适合表格中的当前数据。

> [!NOTE]
> `index` [TableRow 对象的 属性](/javascript/api/excel/excel.tablerow)指示表的行集合中行的索引号。 `TableRow`对象不包含可用于标识行的唯一 `id` 键的属性。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var expensesTable = sheet.tables.getItem("ExpensesTable");

    expensesTable.rows.add(null /*add rows to the end of the table*/, [
        ["1/16/2017", "THE PHONE COMPANY", "Communications", "$120"],
        ["1/20/2017", "NORTHWIND ELECTRIC CARS", "Transportation", "$142"],
        ["1/20/2017", "BEST FOR YOU ORGANICS COMPANY", "Groceries", "$27"],
        ["1/21/2017", "COHO VINEYARD", "Restaurant", "$33"],
        ["1/25/2017", "BELLOWS COLLEGE", "Education", "$350"],
        ["1/28/2017", "TREY RESEARCH", "Other", "$135"],
        ["1/31/2017", "BEST FOR YOU ORGANICS COMPANY", "Groceries", "$97"]
    ]);

    if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
    }

    return context.sync();
}).catch(errorHandlerFunction);
```

**包含新行的表**

![包含新行的Excel。](../images/excel-tables-add-rows.png)

## <a name="add-a-column-to-a-table"></a>向表添加列

下面的示例演示如何向表添加列。 第一个示例使用静态值填充新列；第二个示例使用公式填充新列。

> [!NOTE]
> **TableColumn** 对象的 [index](/javascript/api/excel/excel.tablecolumn) 属性表示表格列集合内列的索引编号。 **TableColumn** 对象的 **id** 属性包含用于标识列的唯一键。

### <a name="add-a-column-that-contains-static-values"></a>添加包含静态值的列

下面的代码示例将一个新列添加到名为 **Sample** 的工作表内的 **ExpensesTable** 表中。 新列添加到表中所有现有列后面，并且包含一个标题（“星期几”），以及用于填充列中单元格的数据。 如果运行Excel的应用程序支持要求集 [](../reference/requirement-sets/excel-api-requirement-sets.md)**ExcelApi 1.2，** 则列宽和行高将设置为最适合表格中的当前数据。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var expensesTable = sheet.tables.getItem("ExpensesTable");

    expensesTable.columns.add(null /*add columns to the end of the table*/, [
        ["Day of the Week"],
        ["Saturday"],
        ["Friday"],
        ["Monday"],
        ["Thursday"],
        ["Sunday"],
        ["Saturday"],
        ["Monday"]
    ]);

    if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
    }

    return context.sync();
}).catch(errorHandlerFunction);
```

**包含新列的表**

![包含新列的Excel。](../images/excel-tables-add-column.png)

### <a name="add-a-column-that-contains-formulas"></a>添加包含公式的列

下面的代码示例将一个新列添加到名为 **Sample** 的工作表内的 **ExpensesTable** 表中。 新列添加到表的末尾，包含标题（“日期类型”），并使用一个公式来填充列中的每个数据单元格。 如果运行Excel的应用程序支持要求集 [](../reference/requirement-sets/excel-api-requirement-sets.md)**ExcelApi 1.2，** 则列宽和行高将设置为最适合表格中的当前数据。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var expensesTable = sheet.tables.getItem("ExpensesTable");

    expensesTable.columns.add(null /*add columns to the end of the table*/, [
        ["Type of the Day"],
        ['=IF(OR((TEXT([DATE], "dddd") = "Saturday"), (TEXT([DATE], "dddd") = "Sunday")), "Weekend", "Weekday")'],
        ['=IF(OR((TEXT([DATE], "dddd") = "Saturday"), (TEXT([DATE], "dddd") = "Sunday")), "Weekend", "Weekday")'],
        ['=IF(OR((TEXT([DATE], "dddd") = "Saturday"), (TEXT([DATE], "dddd") = "Sunday")), "Weekend", "Weekday")'],
        ['=IF(OR((TEXT([DATE], "dddd") = "Saturday"), (TEXT([DATE], "dddd") = "Sunday")), "Weekend", "Weekday")'],
        ['=IF(OR((TEXT([DATE], "dddd") = "Saturday"), (TEXT([DATE], "dddd") = "Sunday")), "Weekend", "Weekday")'],
        ['=IF(OR((TEXT([DATE], "dddd") = "Saturday"), (TEXT([DATE], "dddd") = "Sunday")), "Weekend", "Weekday")'],
        ['=IF(OR((TEXT([DATE], "dddd") = "Saturday"), (TEXT([DATE], "dddd") = "Sunday")), "Weekend", "Weekday")']
    ]);

    if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
    }

    return context.sync();
}).catch(errorHandlerFunction);
```

**包含新的计算列的表**

![包含新计算列的表Excel。](../images/excel-tables-add-calculated-column.png)

## <a name="resize-a-table"></a>调整表格大小

加载项可以调整表格大小，而无需向表格添加数据或更改单元格值。 若要调整表格的大小，请使用 [Table.resize](/javascript/api/excel/excel.table#resize_newRange_) 方法。 下面的代码示例演示如何调整表的大小。 此代码示例使用本文前面创建表格部分 [](#create-a-table)中的 **ExpensesTable，** 将表的新范围设置 **为 A1：D20**。

```js
Excel.run(function (context) {
    // Retrieve the worksheet and a table on that worksheet.
    var sheet = context.workbook.worksheets.getItem("Sample");
    var expensesTable = sheet.tables.getItem("ExpensesTable");

    // Resize the table.
    expensesTable.resize("A1:D20");

    return context.sync();
}).catch(errorHandlerFunction);
```

> [!IMPORTANT]
> 表格的新区域必须与原始区域重叠，并且表格标题 (或表格顶部) 必须在同一行中。

**调整大小后的表格** 

![表中有多个空行的Excel。](../images/excel-tables-resize.png)

## <a name="update-column-name"></a>更新列名称

下面的代码示例将表中第一列的名称更新为 **"购买日期"。** 如果运行Excel的应用程序支持要求集 [](../reference/requirement-sets/excel-api-requirement-sets.md)**ExcelApi 1.2，** 则列宽和行高将设置为最适合表格中的当前数据。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var expensesTable = sheet.tables.getItem("ExpensesTable");
    expensesTable.columns.load("items");

    return context.sync()
        .then(function () {
            expensesTable.columns.items[0].name = "Purchase date";

            if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
                sheet.getUsedRange().format.autofitColumns();
                sheet.getUsedRange().format.autofitRows();
            }

            return context.sync();
        });
}).catch(errorHandlerFunction);
```

**包含新列名称的表格**

![包含新列名称的表Excel。](../images/excel-tables-update-column-name.png)

## <a name="get-data-from-a-table"></a>从表中获取数据

下面的代码示例从名为 **Sample** 的工作表内的 **ExpensesTable** 表中读取数据，然后在同一工作表中的表下输出该数据。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var expensesTable = sheet.tables.getItem("ExpensesTable");

    // Get data from the header row
    var headerRange = expensesTable.getHeaderRowRange().load("values");

    // Get data from the table
    var bodyRange = expensesTable.getDataBodyRange().load("values");

    // Get data from a single column
    var columnRange = expensesTable.columns.getItem("Merchant").getDataBodyRange().load("values");

    // Get data from a single row
    var rowRange = expensesTable.rows.getItemAt(1).load("values");

    // Sync to populate proxy objects with data from Excel
    return context.sync()
        .then(function () {
            var headerValues = headerRange.values;
            var bodyValues = bodyRange.values;
            var merchantColumnValues = columnRange.values;
            var secondRowValues = rowRange.values;

            // Write data from table back to the sheet
            sheet.getRange("A11:A11").values = [["Results"]];
            sheet.getRange("A13:D13").values = headerValues;
            sheet.getRange("A14:D20").values = bodyValues;
            sheet.getRange("B23:B29").values = merchantColumnValues;
            sheet.getRange("A32:D32").values = secondRowValues;

            // Sync to update the sheet in Excel
            return context.sync();
        });
}).catch(errorHandlerFunction);
```

**表和数据输出**

![数据表中的Excel。](../images/excel-tables-get-data.png)

## <a name="detect-data-changes"></a>检测数据更改

外接程序可能需要回应对表中的数据进行更改的用户。 若要检测这些更改，你可以为表的 `onChanged` 事件[注册事件处理程序](excel-add-ins-events.md#register-an-event-handler)。 当事件触发时，`onChanged` 事件的事件处理程序将收到 [TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs) 对象。

`TableChangedEventArgs` 对象提供有关更改和来源的信息。 由于 `onChanged` 会在数据的格式或值发生变化时触发，因此让加载项检查值是否已实际更改可能很有用。 `details` 属性以 [ChangedEventDetail](/javascript/api/excel/excel.changedeventdetail) 的形式封装此信息。 以下代码示例演示如何显示已更改的单元格的之前和之后的值及类型。

```js
// This function would be used as an event handler for the Table.onChanged event.
function onTableChanged(eventArgs) {
    Excel.run(function (context) {
        var details = eventArgs.details;
        var address = eventArgs.address;

        // Print the before and after types and values to the console.
        console.log(`Change at ${address}: was ${details.valueBefore}(${details.valueTypeBefore}),`
            + ` now is ${details.valueAfter}(${details.valueTypeAfter})`);
        return context.sync();
    });
}
```

## <a name="sort-data-in-a-table"></a>对表格中的数据进行排序

下面的代码示例根据表中第四列的值，对表数据按降序进行排序。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var expensesTable = sheet.tables.getItem("ExpensesTable");

    // Queue a command to sort data by the fourth column of the table (descending)
    var sortRange = expensesTable.getDataBodyRange();
    sortRange.sort.apply([
        {
            key: 3,
            ascending: false,
        },
    ]);

    // Sync to run the queued command in Excel
    return context.sync();
}).catch(errorHandlerFunction);
```

**按金额排序的表数据（降序）**

![排序表中的Excel。](../images/excel-tables-sort.png)

在工作表中对数据进行排序时，会触发事件通知。 要详细了解有关排序的事件以及加载项如何注册事件处理程序来响应此类事件，请参阅[处理排序事件](excel-add-ins-worksheets.md#handle-sorting-events)。

## <a name="apply-filters-to-a-table"></a>将筛选器应用于表

下面的代码示例将筛选器应用到表中的 **金额** 列和 **类别** 列。 筛选器筛选的结果是，仅显示符合以下条件的行：**类别** 为其中一个指定值且 **金额** 低于所有行的平均值。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var expensesTable = sheet.tables.getItem("ExpensesTable");

    // Queue a command to apply a filter on the Category column
    filter = expensesTable.columns.getItem("Category").filter;
    filter.apply({
        filterOn: Excel.FilterOn.values,
        values: ["Restaurant", "Groceries"]
    });

    // Queue a command to apply a filter on the Amount column
    var filter = expensesTable.columns.getItem("Amount").filter;
    filter.apply({
        filterOn: Excel.FilterOn.dynamic,
        dynamicCriteria: Excel.DynamicFilterCriteria.belowAverage
    });

    // Sync to run the queued commands in Excel
    return context.sync();
}).catch(errorHandlerFunction);
```

**将筛选器应用于类别和金额的表数据**

![在数据记录中筛选的Excel。](../images/excel-tables-filters-apply.png)

## <a name="clear-table-filters"></a>清除表筛选器

下面的代码示例清除当前应用于表的所有筛选器。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var expensesTable = sheet.tables.getItem("ExpensesTable");

    expensesTable.clearFilters();

    return context.sync();
}).catch(errorHandlerFunction);
```

**没有应用任何筛选器的表数据**

![表中未筛选的表Excel。](../images/excel-tables-filters-clear.png)

## <a name="get-the-visible-range-from-a-filtered-table"></a>从筛选表中获取可见区域

下面的代码示例获取一个区域，其中只包含当前在指定表中可见的单元格数据，然后将该区域的值写入控制台。 可以使用如下所示的方法，在应用列筛选器时获取表 `getVisibleView()` 的可见内容。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var expensesTable = sheet.tables.getItem("ExpensesTable");

    var visibleRange = expensesTable.getDataBodyRange().getVisibleView();
    visibleRange.load("values");

    return context.sync()
        .then(function() {
            console.log(visibleRange.values);
        });
}).catch(errorHandlerFunction);
```

## <a name="autofilter"></a>AutoFilter

加载项可使用表的 [AutoFilter](/javascript/api/excel/excel.autofilter) 对象筛选数据。 `AutoFilter` 对象是表或范围的整个筛选结构。 本文之前讨论的所有筛选操作均与 auto-filter 兼容。 通过单一访问点可以轻松访问和管理多个筛选器。

以下代码示例显示与[之前的代码示例相同的数据筛选](#apply-filters-to-a-table)，但完全通过 auto-filter 完成。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var expensesTable = sheet.tables.getItem("ExpensesTable");

    expensesTable.autoFilter.apply(expensesTable.getRange(), 2, {
        filterOn: Excel.FilterOn.values,
        values: ["Restaurant", "Groceries"]
    });
    expensesTable.autoFilter.apply(expensesTable.getRange(), 3, {
        filterOn: Excel.FilterOn.dynamic,
        dynamicCriteria: Excel.DynamicFilterCriteria.belowAverage
    });

    return context.sync();
}).catch(errorHandlerFunction);
```

`AutoFilter` 也可应用于工作表级别的范围。 有关详细信息，请参阅[使用 Excel JavaScript API 处理工作表](excel-add-ins-worksheets.md#filter-data)。

## <a name="format-a-table"></a>设置表格式

下面的代码示例将格式应用于表。 它为表的标题行、正文、第二行以及第一列指定不同的填充颜色。 有关可以用来指定格式的属性的信息，请参阅 [RangeFormat 对象 (Excel JavaScript API)](/javascript/api/excel/excel.rangeformat)。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var expensesTable = sheet.tables.getItem("ExpensesTable");

    expensesTable.getHeaderRowRange().format.fill.color = "#C70039";
    expensesTable.getDataBodyRange().format.fill.color = "#DAF7A6";
    expensesTable.rows.getItemAt(1).getRange().format.fill.color = "#FFC300";
    expensesTable.columns.getItemAt(0).getDataBodyRange().format.fill.color = "#FFA07A";

    return context.sync();
}).catch(errorHandlerFunction);
```

**应用格式设置的表**

![在应用格式设置后的表Excel。](../images/excel-tables-formatting-after.png)

## <a name="convert-a-range-to-a-table"></a>将区域转换为表

下面的代码示例创建一个数据区域，然后将该区域转换为表。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    // Define values for the range
    var values = [["Product", "Qtr1", "Qtr2", "Qtr3", "Qtr4"],
    ["Frames", 5000, 7000, 6544, 4377],
    ["Saddles", 400, 323, 276, 651],
    ["Brake levers", 12000, 8766, 8456, 9812],
    ["Chains", 1550, 1088, 692, 853],
    ["Mirrors", 225, 600, 923, 544],
    ["Spokes", 6005, 7634, 4589, 8765]];

    // Create the range
    var range = sheet.getRange("A1:E7");
    range.values = values;

    if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
    }

    sheet.activate();

    // Convert the range to a table
    var expensesTable = sheet.tables.add('A1:E7', true);
    expensesTable.name = "ExpensesTable";

    return context.sync();
}).catch(errorHandlerFunction);
```

**内的数据（在区域转换为表之前）**

![Excel 中Excel。](../images/excel-ranges.png)

**表中的数据（在区域转换为表之后）**

![数据表中的Excel。](../images/excel-tables-from-range.png)

## <a name="import-json-data-into-a-table"></a>将 JSON 数据导入表

下面的代码示例在名为 **Sample** 的工作表中创建一个表，然后使用定义了两行数据的 JSON 对象来填充表。 如果运行Excel的应用程序支持要求集 [](../reference/requirement-sets/excel-api-requirement-sets.md)**ExcelApi 1.2，** 则列宽和行高将设置为最适合表格中的当前数据。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var expensesTable = sheet.tables.add("A1:D1", true /*hasHeaders*/);
    expensesTable.name = "ExpensesTable";
    expensesTable.getHeaderRowRange().values = [["Date", "Merchant", "Category", "Amount"]];

    var transactions = [
      {
        "DATE": "1/1/2017",
        "MERCHANT": "The Phone Company",
        "CATEGORY": "Communications",
        "AMOUNT": "$120"
      },
      {
        "DATE": "1/1/2017",
        "MERCHANT": "Southridge Video",
        "CATEGORY": "Entertainment",
        "AMOUNT": "$40"
      }
    ];

    var newData = transactions.map(item =>
        [item.DATE, item.MERCHANT, item.CATEGORY, item.AMOUNT]);

    expensesTable.rows.add(null, newData);

    if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
    }

    sheet.activate();

    return context.sync();
}).catch(errorHandlerFunction);
```

**新建表**

![从导入的 JSON 数据的新表Excel。](../images/excel-tables-create-from-json.png)

## <a name="see-also"></a>另请参阅

- [Excel 加载项中的 Word JavaScript 对象模型](excel-add-ins-core-concepts.md)
