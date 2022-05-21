---
title: 使用 Excel JavaScript API 处理表格
description: 代码示例演示如何使用 Excel JavaScript API 通过表执行常见任务。
ms.date: 05/19/2022
ms.localizationpriority: medium
ms.openlocfilehash: f4cbed134c8ca9f53e89fa97bd4c7ccaa35e45c7
ms.sourcegitcommit: 4ca3334f3cefa34e6b391eb92a429a308229fe89
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/21/2022
ms.locfileid: "65628108"
---
# <a name="work-with-tables-using-the-excel-javascript-api"></a>使用 Excel JavaScript API 处理表格

本文中的代码示例展示了如何使用 Excel JavaScript API 对表格执行常见任务。 有关这些属性和对象支持的属性和方法的完整列表，请参阅适用于 [Excel) 的表对象 (JavaScript API](/javascript/api/excel/excel.table)，以及适用于[Excel) 的 JavaScript API (TableCollection 对象](/javascript/api/excel/excel.tablecollection)。`Table` `TableCollection`

## <a name="create-a-table"></a>创建表

下面的代码示例在名为 **Sample** 的工作表中创建一个表。 此表包含标题，并且包含四列和七行数据。 如果运行代码的Excel应用程序支持 [要求集](/javascript/api/requirement-sets/excel/excel-api-requirement-sets) **ExcelApi 1.2**，则将列的宽度和行的高度设置为最适合表中的当前数据。

> [!NOTE]
> 若要指定表的名称，必须先创建该表，然后设置其 `name` 属性，如以下示例所示。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let expensesTable = sheet.tables.add("A1:D1", true /*hasHeaders*/);
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

    await context.sync();
});
```

### <a name="new-table"></a>新表

![Excel中的新表。](../images/excel-tables-create.png)

## <a name="add-rows-to-a-table"></a>向表添加行

下面的代码示例将七个新行添加到名为 **Sample** 的工作表内的 **ExpensesTable** 表中。 方法 `index` 的 [`add`](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-add-member(1)) 参数设置为 `null`，该参数指定在表中的现有行之后添加行。 参数 `alwaysInsert` 设置为 `true`，指示将新行插入到表中，而不是在表下方。 然后，将列的宽度和行的高度设置为最适合表中的当前数据。

> [!NOTE]
> `index` [TableRow](/javascript/api/excel/excel.tablerow) 对象的属性指示表的行集合中的行的索引号。 对象 `TableRow` 不包含 `id` 可用作唯一键来标识行的属性。

```js
// This code sample shows how to add rows to a table that already exists 
// on a worksheet named Sample.
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let expensesTable = sheet.tables.getItem("ExpensesTable");

    expensesTable.rows.add(
        null, // index, Adds rows to the end of the table.
        [
            ["1/16/2017", "THE PHONE COMPANY", "Communications", "$120"],
            ["1/20/2017", "NORTHWIND ELECTRIC CARS", "Transportation", "$142"],
            ["1/20/2017", "BEST FOR YOU ORGANICS COMPANY", "Groceries", "$27"],
            ["1/21/2017", "COHO VINEYARD", "Restaurant", "$33"],
            ["1/25/2017", "BELLOWS COLLEGE", "Education", "$350"],
            ["1/28/2017", "TREY RESEARCH", "Other", "$135"],
            ["1/31/2017", "BEST FOR YOU ORGANICS COMPANY", "Groceries", "$97"]
        ], 
        true, // alwaysInsert, Specifies that the new rows be inserted into the table.
    );

    sheet.getUsedRange().format.autofitColumns();
    sheet.getUsedRange().format.autofitRows();

    await context.sync();
});
```

### <a name="table-with-new-rows"></a>包含新行的表

![Excel中包含新行的表。](../images/excel-tables-add-rows.png)

## <a name="add-a-column-to-a-table"></a>向表添加列

下面的示例演示如何向表添加列。 第一个示例使用静态值填充新列；第二个示例使用公式填充新列。

> [!NOTE]
> **TableColumn** 对象的 [index](/javascript/api/excel/excel.tablecolumn) 属性表示表格列集合内列的索引编号。 **TableColumn** 对象的 **id** 属性包含用于标识列的唯一键。

### <a name="add-a-column-that-contains-static-values"></a>添加包含静态值的列

下面的代码示例将一个新列添加到名为 **Sample** 的工作表内的 **ExpensesTable** 表中。 新列添加到表中所有现有列后面，并且包含一个标题（“星期几”），以及用于填充列中单元格的数据。 然后，将列的宽度和行的高度设置为最适合表中的当前数据。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let expensesTable = sheet.tables.getItem("ExpensesTable");

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

    sheet.getUsedRange().format.autofitColumns();
    sheet.getUsedRange().format.autofitRows();

    await context.sync();
});
```

#### <a name="table-with-new-column"></a>包含新列的表

![Excel中包含新列的表。](../images/excel-tables-add-column.png)

### <a name="add-a-column-that-contains-formulas"></a>添加包含公式的列

下面的代码示例将一个新列添加到名为 **Sample** 的工作表内的 **ExpensesTable** 表中。 新列添加到表的末尾，包含标题（“日期类型”），并使用一个公式来填充列中的每个数据单元格。 然后，将列的宽度和行的高度设置为最适合表中的当前数据。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let expensesTable = sheet.tables.getItem("ExpensesTable");

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

    sheet.getUsedRange().format.autofitColumns();
    sheet.getUsedRange().format.autofitRows();

    await context.sync();
});
```

#### <a name="table-with-new-calculated-column"></a>包含新的计算列的表

![Excel中包含新计算列的表。](../images/excel-tables-add-calculated-column.png)

## <a name="resize-a-table"></a>调整表的大小

外接程序可以调整表的大小，而无需将数据添加到表或更改单元格值。 若要调整表大小，请使用 [Table.resize](/javascript/api/excel/excel.table#excel-excel-table-resize-member(1)) 方法。 下面的代码示例演示如何调整表的大小。 此代码示例使用本文前面的 [“创建表](#create-a-table)”部分中的 **ExpensesTable**，并将表的新范围设置为 **A1：D20**。

```js
await Excel.run(async (context) => {
    // Retrieve the worksheet and a table on that worksheet.
    let sheet = context.workbook.worksheets.getItem("Sample");
    let expensesTable = sheet.tables.getItem("ExpensesTable");

    // Resize the table.
    expensesTable.resize("A1:D20");

    await context.sync();
});
```

> [!IMPORTANT]
> 表的新范围必须与原始范围重叠，表 (或表顶部) 必须在同一行中。

### <a name="table-after-resize"></a>调整大小后的表

![Excel中包含多个空行的表。](../images/excel-tables-resize.png)

## <a name="update-column-name"></a>更新列名称

下面的代码示例将表中第一列的名称更新为 **购买日期**。 然后，将列的宽度和行的高度设置为最适合表中的当前数据。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let expensesTable = sheet.tables.getItem("ExpensesTable");
    expensesTable.columns.load("items");

    await context.sync();
        
    expensesTable.columns.items[0].name = "Purchase date";

    sheet.getUsedRange().format.autofitColumns();
    sheet.getUsedRange().format.autofitRows();

    await context.sync();
});
```

### <a name="table-with-new-column-name"></a>包含新列名称的表

![Excel中包含新列名的表。](../images/excel-tables-update-column-name.png)

## <a name="get-data-from-a-table"></a>从表中获取数据

下面的代码示例从名为 **Sample** 的工作表内的 **ExpensesTable** 表中读取数据，然后在同一工作表中的表下输出该数据。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let expensesTable = sheet.tables.getItem("ExpensesTable");

    // Get data from the header row.
    let headerRange = expensesTable.getHeaderRowRange().load("values");

    // Get data from the table.
    let bodyRange = expensesTable.getDataBodyRange().load("values");

    // Get data from a single column.
    let columnRange = expensesTable.columns.getItem("Merchant").getDataBodyRange().load("values");

    // Get data from a single row.
    let rowRange = expensesTable.rows.getItemAt(1).load("values");

    // Sync to populate proxy objects with data from Excel.
    await context.sync();

    let headerValues = headerRange.values;
    let bodyValues = bodyRange.values;
    let merchantColumnValues = columnRange.values;
    let secondRowValues = rowRange.values;

    // Write data from table back to the sheet
    sheet.getRange("A11:A11").values = [["Results"]];
    sheet.getRange("A13:D13").values = headerValues;
    sheet.getRange("A14:D20").values = bodyValues;
    sheet.getRange("B23:B29").values = merchantColumnValues;
    sheet.getRange("A32:D32").values = secondRowValues;

    // Sync to update the sheet in Excel.
    await context.sync();
});
```

### <a name="table-and-data-output"></a>表和数据输出

![Excel中的表数据。](../images/excel-tables-get-data.png)

## <a name="detect-data-changes"></a>检测数据更改

外接程序可能需要回应对表中的数据进行更改的用户。 若要检测这些更改，你可以为表的 `onChanged` 事件[注册事件处理程序](excel-add-ins-events.md#register-an-event-handler)。 当事件触发时，`onChanged` 事件的事件处理程序将收到 [TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs) 对象。

`TableChangedEventArgs` 对象提供有关更改和来源的信息。 由于 `onChanged` 会在数据的格式或值发生变化时触发，因此让加载项检查值是否已实际更改可能很有用。 `details` 属性以 [ChangedEventDetail](/javascript/api/excel/excel.changedeventdetail) 的形式封装此信息。 以下代码示例演示如何显示已更改的单元格的之前和之后的值及类型。

```js
// This function would be used as an event handler for the Table.onChanged event.
async function onTableChanged(eventArgs) {
    await Excel.run(async (context) => {
        let details = eventArgs.details;
        let address = eventArgs.address;

        // Print the before and after types and values to the console.
        console.log(`Change at ${address}: was ${details.valueBefore}(${details.valueTypeBefore}),`
            + ` now is ${details.valueAfter}(${details.valueTypeAfter})`);
        await context.sync();
    });
}
```

## <a name="sort-data-in-a-table"></a>对表格中的数据进行排序

下面的代码示例根据表中第四列的值，对表数据按降序进行排序。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let expensesTable = sheet.tables.getItem("ExpensesTable");

    // Queue a command to sort data by the fourth column of the table (descending).
    let sortRange = expensesTable.getDataBodyRange();
    sortRange.sort.apply([
        {
            key: 3,
            ascending: false,
        },
    ]);

    // Sync to run the queued command in Excel.
    await context.sync();
});
```

### <a name="table-data-sorted-by-amount-descending"></a>按金额排序的表数据（降序）

![Excel中的表数据排序。](../images/excel-tables-sort.png)

在工作表中对数据进行排序时，会触发事件通知。 要详细了解有关排序的事件以及加载项如何注册事件处理程序来响应此类事件，请参阅[处理排序事件](excel-add-ins-worksheets.md#handle-sorting-events)。

## <a name="apply-filters-to-a-table"></a>将筛选器应用于表

下面的代码示例将筛选器应用到表中的 **金额** 列和 **类别** 列。 筛选器筛选的结果是，仅显示符合以下条件的行：**类别** 为其中一个指定值且 **金额** 低于所有行的平均值。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let expensesTable = sheet.tables.getItem("ExpensesTable");

    // Queue a command to apply a filter on the Category column.
    let categoryFilter = expensesTable.columns.getItem("Category").filter;
    categoryFilter.apply({
      filterOn: Excel.FilterOn.values,
      values: ["Restaurant", "Groceries"]
    });

    // Queue a command to apply a filter on the Amount column.
    let amountFilter = expensesTable.columns.getItem("Amount").filter;
    amountFilter.apply({
      filterOn: Excel.FilterOn.dynamic,
      dynamicCriteria: Excel.DynamicFilterCriteria.belowAverage
    });

    // Sync to run the queued commands in Excel.
    await context.sync();
});
```

### <a name="table-data-with-filters-applied-for-category-and-amount"></a>将筛选器应用于类别和金额的表数据

![在Excel中筛选的表数据。](../images/excel-tables-filters-apply.png)

## <a name="clear-table-filters"></a>清除表筛选器

下面的代码示例清除当前应用于表的所有筛选器。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let expensesTable = sheet.tables.getItem("ExpensesTable");

    expensesTable.clearFilters();

    await context.sync();
});
```

### <a name="table-data-with-no-filters-applied"></a>没有应用任何筛选器的表数据

![Excel中未筛选的表数据。](../images/excel-tables-filters-clear.png)

## <a name="get-the-visible-range-from-a-filtered-table"></a>从筛选表中获取可见区域

下面的代码示例获取一个区域，其中只包含当前在指定表中可见的单元格数据，然后将该区域的值写入控制台。 只要应用列筛选器，就可以使用 `getVisibleView()` 如下所示的方法获取表的可见内容。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let expensesTable = sheet.tables.getItem("ExpensesTable");

    let visibleRange = expensesTable.getDataBodyRange().getVisibleView();
    visibleRange.load("values");

    await context.sync();
    console.log(visibleRange.values);
});
```

## <a name="autofilter"></a>AutoFilter

加载项可使用表的 [AutoFilter](/javascript/api/excel/excel.autofilter) 对象筛选数据。 `AutoFilter` 对象是表或范围的整个筛选结构。 本文之前讨论的所有筛选操作均与 auto-filter 兼容。 通过单一访问点可以轻松访问和管理多个筛选器。

以下代码示例显示与[之前的代码示例相同的数据筛选](#apply-filters-to-a-table)，但完全通过 auto-filter 完成。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let expensesTable = sheet.tables.getItem("ExpensesTable");

    expensesTable.autoFilter.apply(expensesTable.getRange(), 2, {
        filterOn: Excel.FilterOn.values,
        values: ["Restaurant", "Groceries"]
    });
    expensesTable.autoFilter.apply(expensesTable.getRange(), 3, {
        filterOn: Excel.FilterOn.dynamic,
        dynamicCriteria: Excel.DynamicFilterCriteria.belowAverage
    });

    await context.sync();
});
```

`AutoFilter` 也可应用于工作表级别的范围。 有关详细信息，请参阅[使用 Excel JavaScript API 处理工作表](excel-add-ins-worksheets.md#filter-data)。

## <a name="format-a-table"></a>设置表格式

下面的代码示例将格式应用于表。 它为表的标题行、正文、第二行以及第一列指定不同的填充颜色。 有关可以用来指定格式的属性的信息，请参阅 [RangeFormat 对象 (Excel JavaScript API)](/javascript/api/excel/excel.rangeformat)。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let expensesTable = sheet.tables.getItem("ExpensesTable");

    expensesTable.getHeaderRowRange().format.fill.color = "#C70039";
    expensesTable.getDataBodyRange().format.fill.color = "#DAF7A6";
    expensesTable.rows.getItemAt(1).getRange().format.fill.color = "#FFC300";
    expensesTable.columns.getItemAt(0).getDataBodyRange().format.fill.color = "#FFA07A";

    await context.sync();
});
```

### <a name="table-after-formatting-is-applied"></a>应用了格式设置的表

![在Excel中应用格式设置后的表。](../images/excel-tables-formatting-after.png)

## <a name="convert-a-range-to-a-table"></a>将区域转换为表

下面的代码示例创建一个数据区域，然后将该区域转换为表。 然后，将列的宽度和行的高度设置为最适合表中的当前数据。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    // Define values for the range.
    let values = [["Product", "Qtr1", "Qtr2", "Qtr3", "Qtr4"],
    ["Frames", 5000, 7000, 6544, 4377],
    ["Saddles", 400, 323, 276, 651],
    ["Brake levers", 12000, 8766, 8456, 9812],
    ["Chains", 1550, 1088, 692, 853],
    ["Mirrors", 225, 600, 923, 544],
    ["Spokes", 6005, 7634, 4589, 8765]];

    // Create the range.
    let range = sheet.getRange("A1:E7");
    range.values = values;

    sheet.getUsedRange().format.autofitColumns();
    sheet.getUsedRange().format.autofitRows();

    sheet.activate();

    // Convert the range to a table.
    let expensesTable = sheet.tables.add('A1:E7', true);
    expensesTable.name = "ExpensesTable";

    await context.sync();
});
```

### <a name="data-in-the-range-before-the-range-is-converted-to-a-table"></a>区域内的数据（在区域转换为表之前）

![Excel范围内的数据。](../images/excel-ranges.png)

### <a name="data-in-the-table-after-the-range-is-converted-to-a-table"></a>表中的数据（在区域转换为表之后）

![Excel中的表中的数据。](../images/excel-tables-from-range.png)

## <a name="import-json-data-into-a-table"></a>将 JSON 数据导入表

下面的代码示例在名为 **Sample** 的工作表中创建一个表，然后使用定义了两行数据的 JSON 对象来填充表。 然后，将列的宽度和行的高度设置为最适合表中的当前数据。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let expensesTable = sheet.tables.add("A1:D1", true /*hasHeaders*/);
    expensesTable.name = "ExpensesTable";
    expensesTable.getHeaderRowRange().values = [["Date", "Merchant", "Category", "Amount"]];

    let transactions = [
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

    let newData = transactions.map(item =>
        [item.DATE, item.MERCHANT, item.CATEGORY, item.AMOUNT]);

    expensesTable.rows.add(null, newData);

    sheet.getUsedRange().format.autofitColumns();
    sheet.getUsedRange().format.autofitRows();

    sheet.activate();

    await context.sync();
});
```

### <a name="new-table"></a>新表

![Excel中导入的 JSON 数据中的新表。](../images/excel-tables-create-from-json.png)

## <a name="see-also"></a>另请参阅

- [Excel 加载项中的 Word JavaScript 对象模型](excel-add-ins-core-concepts.md)
