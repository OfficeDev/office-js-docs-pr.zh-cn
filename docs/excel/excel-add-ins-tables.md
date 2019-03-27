---
title: 使用 Excel JavaScript API 处理表格
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: a628c182ccb570fcda3db813f7debb237682b915
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/27/2019
ms.locfileid: "30869972"
---
# <a name="work-with-tables-using-the-excel-javascript-api"></a><span data-ttu-id="a3e7e-102">使用 Excel JavaScript API 处理表格</span><span class="sxs-lookup"><span data-stu-id="a3e7e-102">Work with tables using the Excel JavaScript API</span></span>

<span data-ttu-id="a3e7e-p101">本文中的代码示例展示了如何使用 Excel JavaScript API 对表格执行常见任务。 有关 **Table** 和 **TableCollection** 对象支持的属性和方法的完整列表，请参阅 [Table 对象 (Excel JavaScript API)](/javascript/api/excel/excel.table) 和 [TableCollection 对象 (Excel JavaScript API)](/javascript/api/excel/excel.tablecollection)。</span><span class="sxs-lookup"><span data-stu-id="a3e7e-p101">This article provides code samples that show how to perform common tasks with tables using the Excel JavaScript API. For the complete list of properties and methods that the **Table** and **TableCollection** objects support, see [Table Object (JavaScript API for Excel)](/javascript/api/excel/excel.table) and [TableCollection Object (JavaScript API for Excel)](/javascript/api/excel/excel.tablecollection).</span></span>

## <a name="create-a-table"></a><span data-ttu-id="a3e7e-105">创建表</span><span class="sxs-lookup"><span data-stu-id="a3e7e-105">Create a table</span></span>

<span data-ttu-id="a3e7e-p102">下面的代码示例在名为 **Sample** 的工作表中创建一个表。 此表包含标题，并且包含四列和七行数据。 如果在其中运行代码的 Excel 主机应用程序支持[要求集](/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets) **ExcelApi 1.2**，则为表中的当前数据设置最佳列宽和行高。</span><span class="sxs-lookup"><span data-stu-id="a3e7e-p102">The following code sample creates a table in the worksheet named **Sample**. The table has headers and contains four columns and seven rows of data. If the Excel host application where the code is running supports [requirement set](/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets) **ExcelApi 1.2**, the width of the columns and height of the rows are set to best fit the current data in the table.</span></span>

> [!NOTE]
> <span data-ttu-id="a3e7e-109">若要指定表格名称，必须先创建表格，再设置它的 **name** 属性，如下面的示例所示。</span><span class="sxs-lookup"><span data-stu-id="a3e7e-109">To specify a name for a table, you must first create the table and then set its **name** property, as shown in the example below.</span></span>

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

    if (Office.context.requirements.isSetSupported("ExcelApi", 1.2)) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
    }

    sheet.activate();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="a3e7e-110">**新建表格**</span><span class="sxs-lookup"><span data-stu-id="a3e7e-110">**New table**</span></span>

![Excel 中的新表](../images/excel-tables-create.png)

## <a name="add-rows-to-a-table"></a><span data-ttu-id="a3e7e-112">向表添加行</span><span class="sxs-lookup"><span data-stu-id="a3e7e-112">Add rows to a table</span></span>

<span data-ttu-id="a3e7e-p103">下面的代码示例将七个新行添加到名为 **Sample** 的工作表内的 **ExpensesTable** 表中。 新行被添加到表的末尾。 如果在其中运行代码的 Excel 主机应用程序支持[要求集](/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets) **ExcelApi 1.2**，则为表中的当前数据设置最佳列宽和行高。</span><span class="sxs-lookup"><span data-stu-id="a3e7e-p103">The following code sample adds seven new rows to the table named **ExpensesTable** within the worksheet named **Sample**. The new rows are added to the end of the table. If the Excel host application where the code is running supports [requirement set](/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets) **ExcelApi 1.2**, the width of the columns and height of the rows are set to best fit the current data in the table.</span></span>

> [!NOTE]
> <span data-ttu-id="a3e7e-p104">**TableRow** 对象的 [index](/javascript/api/excel/excel.tablerow) 属性表示表格行集合内行的索引编号。 **TableRow** 对象不包含可用作标识行的唯一键的 **id** 属性。</span><span class="sxs-lookup"><span data-stu-id="a3e7e-p104">The **index** property of a [TableRow](/javascript/api/excel/excel.tablerow) object indicates the index number of the row within the rows collection of the table. A **TableRow** object does not contain an **id** property that can be used as a unique key to identify the row.</span></span>

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

    if (Office.context.requirements.isSetSupported("ExcelApi", 1.2)) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
    }

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="a3e7e-118">**包含新行的表**</span><span class="sxs-lookup"><span data-stu-id="a3e7e-118">**Table with new rows**</span></span>

![Excel 中包含新行的表](../images/excel-tables-add-rows.png)

## <a name="add-a-column-to-a-table"></a><span data-ttu-id="a3e7e-120">向表添加列</span><span class="sxs-lookup"><span data-stu-id="a3e7e-120">Add a column to a table</span></span>

<span data-ttu-id="a3e7e-p105">下面的示例演示如何向表添加列。 第一个示例使用静态值填充新列；第二个示例使用公式填充新列。</span><span class="sxs-lookup"><span data-stu-id="a3e7e-p105">These examples show how to add a column to a table. The first example populates the new column with static values; the second example populates the new column with formulas.</span></span>

> [!NOTE]
> <span data-ttu-id="a3e7e-p106">**TableColumn** 对象的 [index](/javascript/api/excel/excel.tablecolumn) 属性表示表格列集合内列的索引编号。 **TableColumn** 对象的 **id** 属性包含用于标识列的唯一键。</span><span class="sxs-lookup"><span data-stu-id="a3e7e-p106">The **index** property of a [TableColumn](/javascript/api/excel/excel.tablecolumn) object indicates the index number of the column within the columns collection of the table. The **id** property of a **TableColumn** object contains a unique key that identifies the column.</span></span>

### <a name="add-a-column-that-contains-static-values"></a><span data-ttu-id="a3e7e-125">添加包含静态值的列</span><span class="sxs-lookup"><span data-stu-id="a3e7e-125">Add a column that contains static values</span></span>

<span data-ttu-id="a3e7e-p107">下面的代码示例将一个新列添加到名为 **Sample** 的工作表内的 **ExpensesTable** 表中。 新列添加到表中所有现有列后面，并且包含一个标题（“星期几”），以及用于填充列中单元格的数据。 如果在其中运行代码的 Excel 主机应用程序支持[要求集](/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets) **ExcelApi 1.2**，则为表中的当前数据设置最佳列宽和行高。</span><span class="sxs-lookup"><span data-stu-id="a3e7e-p107">The following code sample adds a new column to the table named **ExpensesTable** within the worksheet named **Sample**. The new column is added after all existing columns in the table and contains a header ("Day of the Week") as well as data to populate the cells in the column. If the Excel host application where the code is running supports [requirement set](/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets) **ExcelApi 1.2**, the width of the columns and height of the rows are set to best fit the current data in the table.</span></span>

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

    if (Office.context.requirements.isSetSupported("ExcelApi", 1.2)) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
    }

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="a3e7e-129">**包含新列的表**</span><span class="sxs-lookup"><span data-stu-id="a3e7e-129">**Table with new column**</span></span>

![Excel 中包含新列的表](../images/excel-tables-add-column.png)

### <a name="add-a-column-that-contains-formulas"></a><span data-ttu-id="a3e7e-131">添加包含公式的列</span><span class="sxs-lookup"><span data-stu-id="a3e7e-131">Add a column that contains formulas</span></span>

<span data-ttu-id="a3e7e-p108">下面的代码示例将一个新列添加到名为 **Sample** 的工作表内的 **ExpensesTable** 表中。 新列添加到表的末尾，包含标题（“日期类型”），并使用一个公式来填充列中的每个数据单元格。 如果在其中运行代码的 Excel 主机应用程序支持[要求集](/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets) **ExcelApi 1.2**，则为表中的当前数据设置最佳列宽和行高。</span><span class="sxs-lookup"><span data-stu-id="a3e7e-p108">The following code sample adds a new column to the table named **ExpensesTable** within the worksheet named **Sample**. The new column is added to the end of the table, contains a header ("Type of the Day"), and uses a formula to populate each data cell in the column. If the Excel host application where the code is running supports [requirement set](/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets) **ExcelApi 1.2**, the width of the columns and height of the rows are set to best fit the current data in the table.</span></span>

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

    if (Office.context.requirements.isSetSupported("ExcelApi", 1.2)) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
    }

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="a3e7e-135">**包含新的计算列的表**</span><span class="sxs-lookup"><span data-stu-id="a3e7e-135">**Table with new calculated column**</span></span>

![Excel 中包含新的计算列的表](../images/excel-tables-add-calculated-column.png)

## <a name="update-column-name"></a><span data-ttu-id="a3e7e-137">更新列名称</span><span class="sxs-lookup"><span data-stu-id="a3e7e-137">Update column name</span></span>

<span data-ttu-id="a3e7e-p109">下面的代码示例将表格中第一列的名称更新为“购买日期”\*\*\*\*。如果运行代码的 Excel 主机应用支持[要求集](/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets) **ExcelApi 1.2**，那么列宽和行高会设置为最适应表格中的当前数据。</span><span class="sxs-lookup"><span data-stu-id="a3e7e-p109">The following code sample updates the name of the first column in the table to **Purchase date**. If the Excel host application where the code is running supports [requirement set](/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets) **ExcelApi 1.2**, the width of the columns and height of the rows are set to best fit the current data in the table.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var expensesTable = sheet.tables.getItem("ExpensesTable");
    expensesTable.columns.load("items");

    return context.sync()
        .then(function () {
            expensesTable.columns.items[0].name = "Purchase date";

            if (Office.context.requirements.isSetSupported("ExcelApi", 1.2)) {
                sheet.getUsedRange().format.autofitColumns();
                sheet.getUsedRange().format.autofitRows();
            }

            return context.sync();
        });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="a3e7e-140">**包含新列名称的表格**</span><span class="sxs-lookup"><span data-stu-id="a3e7e-140">**Table with new column name**</span></span>

![Excel 中包含新的列名称的表](../images/excel-tables-update-column-name.png)

## <a name="get-data-from-a-table"></a><span data-ttu-id="a3e7e-142">从表中获取数据</span><span class="sxs-lookup"><span data-stu-id="a3e7e-142">Get data from a table</span></span>

<span data-ttu-id="a3e7e-143">下面的代码示例从名为 **Sample** 的工作表内的 **ExpensesTable** 表中读取数据，然后在同一工作表中的表下输出该数据。</span><span class="sxs-lookup"><span data-stu-id="a3e7e-143">The following code sample reads data from a table named **ExpensesTable** in the worksheet named **Sample** and then outputs that data below the table in the same worksheet.</span></span>

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

<span data-ttu-id="a3e7e-144">**表和数据输出**</span><span class="sxs-lookup"><span data-stu-id="a3e7e-144">**Table and data output**</span></span>

![Excel 中的表数据](../images/excel-tables-get-data.png)

## <a name="sort-data-in-a-table"></a><span data-ttu-id="a3e7e-146">在表中对数据进行排序</span><span class="sxs-lookup"><span data-stu-id="a3e7e-146">Sort data in a table</span></span>

<span data-ttu-id="a3e7e-147">下面的代码示例根据表中第四列的值，对表数据按降序进行排序。</span><span class="sxs-lookup"><span data-stu-id="a3e7e-147">The following code sample sorts table data in descending order according to the values in the fourth column of the table.</span></span>

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

<span data-ttu-id="a3e7e-148">**按金额排序的表数据（降序）**</span><span class="sxs-lookup"><span data-stu-id="a3e7e-148">**Table data sorted by Amount (descending)**</span></span>

![Excel 中的表数据](../images/excel-tables-sort.png)

## <a name="apply-filters-to-a-table"></a><span data-ttu-id="a3e7e-150">将筛选器应用于表</span><span class="sxs-lookup"><span data-stu-id="a3e7e-150">Apply filters to a table</span></span>

<span data-ttu-id="a3e7e-p110">下面的代码示例将筛选器应用到表中的**金额**列和**类别**列。 筛选器筛选的结果是，仅显示符合以下条件的行：**类别**为其中一个指定值且**金额**低于所有行的平均值。</span><span class="sxs-lookup"><span data-stu-id="a3e7e-p110">The following code sample applies filters to the **Amount** column and the **Category** column within a table. As a result of the filters, only rows where **Category** is one of the specified values and **Amount** is below the average value for all rows is shown.</span></span>

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

<span data-ttu-id="a3e7e-153">**将筛选器应用于类别和金额的表数据**</span><span class="sxs-lookup"><span data-stu-id="a3e7e-153">**Table data with filters applied for Category and Amount**</span></span>

![Excel 中经过筛选的表数据](../images/excel-tables-filters-apply.png)

## <a name="clear-table-filters"></a><span data-ttu-id="a3e7e-155">清除表筛选器</span><span class="sxs-lookup"><span data-stu-id="a3e7e-155">Clear table filters</span></span>

<span data-ttu-id="a3e7e-156">下面的代码示例清除当前应用于表的所有筛选器。</span><span class="sxs-lookup"><span data-stu-id="a3e7e-156">The following code sample clears any filters currently applied on the table.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var expensesTable = sheet.tables.getItem("ExpensesTable");

    expensesTable.clearFilters();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="a3e7e-157">**没有应用任何筛选器的表数据**</span><span class="sxs-lookup"><span data-stu-id="a3e7e-157">**Table data with no filters applied**</span></span>

![Excel 中未经筛选的表数据](../images/excel-tables-filters-clear.png)

## <a name="get-the-visible-range-from-a-filtered-table"></a><span data-ttu-id="a3e7e-159">从筛选表中获取可见区域</span><span class="sxs-lookup"><span data-stu-id="a3e7e-159">Get the visible range from a filtered table</span></span>

<span data-ttu-id="a3e7e-p111">下面的代码示例获取一个区域，其中只包含当前在指定表中可见的单元格数据，然后将该区域的值写入控制台。 可以使用如下所示的 **getVisibleView()** 方法，在应用列筛选器时，都能获取表的可见内容。</span><span class="sxs-lookup"><span data-stu-id="a3e7e-p111">The following code sample gets a range that contains data only for cells that are currently visible within the specified table, and then writes the values of that range to the console. You can use the **getVisibleView()** method as shown below to get the visible contents of a table whenever column filters have been applied.</span></span>

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

## <a name="format-a-table"></a><span data-ttu-id="a3e7e-162">设置表格式</span><span class="sxs-lookup"><span data-stu-id="a3e7e-162">Format a table</span></span>

<span data-ttu-id="a3e7e-p112">下面的代码示例将格式应用于表。 它为表的标题行、正文、第二行以及第一列指定不同的填充颜色。 有关可以用来指定格式的属性的信息，请参阅 [RangeFormat 对象 (Excel JavaScript API)](/javascript/api/excel/excel.rangeformat)。</span><span class="sxs-lookup"><span data-stu-id="a3e7e-p112">The following code sample applies formatting to a table. It specifies different fill colors for the header row of the table, the body of the table, the second row of the table, and the first column of the table. For information about the properties you can use to specify format, see [RangeFormat Object (JavaScript API for Excel)](/javascript/api/excel/excel.rangeformat).</span></span>

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

<span data-ttu-id="a3e7e-166">**应用格式设置的表**</span><span class="sxs-lookup"><span data-stu-id="a3e7e-166">**Table after formatting is applied**</span></span>

![Excel 中应用了格式设置的表](../images/excel-tables-formatting-after.png)

## <a name="convert-a-range-to-a-table"></a><span data-ttu-id="a3e7e-168">将区域转换为表</span><span class="sxs-lookup"><span data-stu-id="a3e7e-168">Convert a range to a table</span></span>

<span data-ttu-id="a3e7e-169">下面的代码示例创建一个数据区域，然后将该区域转换为表。</span><span class="sxs-lookup"><span data-stu-id="a3e7e-169">The following code sample creates a range of data and then converts that range to a table.</span></span>

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

    if (Office.context.requirements.isSetSupported("ExcelApi", 1.2)) {
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

<span data-ttu-id="a3e7e-170">**内的数据（在区域转换为表之前）**</span><span class="sxs-lookup"><span data-stu-id="a3e7e-170">**Data in the range (before the range is converted to a table)**</span></span>

![Excel 中区域内的数据](../images/excel-ranges.png)

<span data-ttu-id="a3e7e-172">**表中的数据（在区域转换为表之后）**</span><span class="sxs-lookup"><span data-stu-id="a3e7e-172">**Data in the table (after the range is converted to a table)**</span></span>

![Excel 中表的数据](../images/excel-tables-from-range.png)

## <a name="import-json-data-into-a-table"></a><span data-ttu-id="a3e7e-174">将 JSON 数据导入表</span><span class="sxs-lookup"><span data-stu-id="a3e7e-174">Import JSON data into a table</span></span>

<span data-ttu-id="a3e7e-p113">下面的代码示例在名为 **Sample** 的工作表中创建一个表，然后使用定义了两行数据的 JSON 对象来填充表。 如果在其中运行代码的 Excel 主机应用程序支持[要求集](/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets) **ExcelApi 1.2**，则为表中的当前数据设置最佳列宽和行高。</span><span class="sxs-lookup"><span data-stu-id="a3e7e-p113">The following code sample creates a table in the worksheet named **Sample** and then populates the table by using a JSON object that defines two rows of data. If the Excel host application where the code is running supports [requirement set](/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets) **ExcelApi 1.2**, the width of the columns and height of the rows are set to best fit the current data in the table.</span></span>

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

    if (Office.context.requirements.isSetSupported("ExcelApi", 1.2)) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
    }

    sheet.activate();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="a3e7e-177">**新建表**</span><span class="sxs-lookup"><span data-stu-id="a3e7e-177">**New table**</span></span>

![Excel 中的新表格](../images/excel-tables-create-from-json.png)

## <a name="see-also"></a><span data-ttu-id="a3e7e-179">另请参阅</span><span class="sxs-lookup"><span data-stu-id="a3e7e-179">See also</span></span>

- [<span data-ttu-id="a3e7e-180">Excel JavaScript API 基本编程概念</span><span class="sxs-lookup"><span data-stu-id="a3e7e-180">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
