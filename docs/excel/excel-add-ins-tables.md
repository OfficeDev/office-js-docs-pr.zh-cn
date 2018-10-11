---
title: 使用 Excel JavaScript API 处理表格
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 1e8c71f34de7a295fcac8e5ea6a4fff5cae4fdcf
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459159"
---
# <a name="work-with-tables-using-the-excel-javascript-api"></a><span data-ttu-id="1d8ca-102">使用 Excel JavaScript API 处理表格</span><span class="sxs-lookup"><span data-stu-id="1d8ca-102">Work with tables using the Excel JavaScript API</span></span>

<span data-ttu-id="1d8ca-103">本文中的代码示例展示了如何使用 Excel JavaScript API 对表格执行常见任务。</span><span class="sxs-lookup"><span data-stu-id="1d8ca-103">This article provides code samples that show how to perform common tasks with tables using the Excel JavaScript API.</span></span> <span data-ttu-id="1d8ca-104">有关 **Table** 和 **TableCollection** 对象支持的属性和方法的完整列表，请参阅 [Table 对象 (Excel JavaScript API)](https://docs.microsoft.com/javascript/api/excel/excel.table?view=office-js) 和 [TableCollection 对象 (Excel JavaScript API)](https://docs.microsoft.com/javascript/api/excel/excel.tablecollection?view=office-js)。</span><span class="sxs-lookup"><span data-stu-id="1d8ca-104">For the complete list of properties and methods that the **Table** and **TableCollection** objects support, see [Table Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.table?view=office-js) and [TableCollection Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.tablecollection?view=office-js).</span></span>

## <a name="create-a-table"></a><span data-ttu-id="1d8ca-105">创建表</span><span class="sxs-lookup"><span data-stu-id="1d8ca-105">Create a table</span></span>

<span data-ttu-id="1d8ca-106">下面的代码示例在名为 **Sample** 的工作表中创建一个表。</span><span class="sxs-lookup"><span data-stu-id="1d8ca-106">The following code sample creates a table in the worksheet named **Sample**.</span></span> <span data-ttu-id="1d8ca-107">此表包含标题，并且包含四列和七行数据。</span><span class="sxs-lookup"><span data-stu-id="1d8ca-107">The table has headers and contains four columns and seven rows of data.</span></span> <span data-ttu-id="1d8ca-108">如果在其中运行代码的 Excel 主机应用程序支持[要求集](https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets?view=office-js) **ExcelApi 1.2**，则为表中的当前数据设置最佳列宽和行高。</span><span class="sxs-lookup"><span data-stu-id="1d8ca-108">If the Excel host application where the code is running supports [requirement set](https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets?view=office-js) **ExcelApi 1.2**, the width of the columns and height of the rows are set to best fit the current data in the table.</span></span>

> [!NOTE]
> <span data-ttu-id="1d8ca-109">若要指定表格名称，必须先创建表格，再设置它的 **name** 属性，如下面的示例所示。</span><span class="sxs-lookup"><span data-stu-id="1d8ca-109">To specify a name for a table, you must first create the table and then set its **name** property, as shown in the example below.</span></span>

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

<span data-ttu-id="1d8ca-110">**新建表**</span><span class="sxs-lookup"><span data-stu-id="1d8ca-110">**New table**</span></span>

![Excel 中的新表格](../images/excel-tables-create.png)

## <a name="add-rows-to-a-table"></a><span data-ttu-id="1d8ca-112">向表添加行</span><span class="sxs-lookup"><span data-stu-id="1d8ca-112">Add rows to a table</span></span>

<span data-ttu-id="1d8ca-113">下面的代码示例将七个新行添加到名为 **Sample** 的工作表内的 **ExpensesTable** 表中。</span><span class="sxs-lookup"><span data-stu-id="1d8ca-113">The following code sample adds seven new rows to the table named **ExpensesTable** within the worksheet named **Sample**.</span></span> <span data-ttu-id="1d8ca-114">新行被添加到表的末尾。</span><span class="sxs-lookup"><span data-stu-id="1d8ca-114">The new rows are added to the end of the table.</span></span> <span data-ttu-id="1d8ca-115">如果在其中运行代码的 Excel 主机应用程序支持[要求集](https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets?view=office-js) **ExcelApi 1.2**，则为表中的当前数据设置最佳列宽和行高。</span><span class="sxs-lookup"><span data-stu-id="1d8ca-115">If the Excel host application where the code is running supports [requirement set](https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets?view=office-js) **ExcelApi 1.2**, the width of the columns and height of the rows are set to best fit the current data in the table.</span></span>

> [!NOTE]
> <span data-ttu-id="1d8ca-116">[TableRow](https://docs.microsoft.com/javascript/api/excel/excel.tablerow?view=office-js) 对象的 **index** 属性表示表格行集合内行的索引编号。</span><span class="sxs-lookup"><span data-stu-id="1d8ca-116">The **index** property of a [TableRow](https://docs.microsoft.com/javascript/api/excel/excel.tablerow?view=office-js) object indicates the index number of the row within the rows collection of the table.</span></span> <span data-ttu-id="1d8ca-117">**TableRow** 对象不包含可用作标识行的唯一键的 **id** 属性。</span><span class="sxs-lookup"><span data-stu-id="1d8ca-117">A **TableRow** object does not contain an **id** property that can be used as a unique key to identify the row.</span></span>

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

<span data-ttu-id="1d8ca-118">**包含新行的表**</span><span class="sxs-lookup"><span data-stu-id="1d8ca-118">**Table with new rows**</span></span>

![Excel 中包含新行的表](../images/excel-tables-add-rows.png)

## <a name="add-a-column-to-a-table"></a><span data-ttu-id="1d8ca-120">向表添加列</span><span class="sxs-lookup"><span data-stu-id="1d8ca-120">Add a column to a table</span></span>

<span data-ttu-id="1d8ca-121">下面的示例演示如何向表添加列。</span><span class="sxs-lookup"><span data-stu-id="1d8ca-121">These examples show how to add a column to a table.</span></span> <span data-ttu-id="1d8ca-122">第一个示例使用静态值填充新列；第二个示例使用公式填充新列。</span><span class="sxs-lookup"><span data-stu-id="1d8ca-122">The first example populates the new column with static values; the second example populates the new column with formulas.</span></span>

> [!NOTE]
> <span data-ttu-id="1d8ca-123">[TableColumn](https://docs.microsoft.com/javascript/api/excel/excel.tablecolumn?view=office-js) 对象的 **index** 属性表示表格列集合内列的索引编号。</span><span class="sxs-lookup"><span data-stu-id="1d8ca-123">The **index** property of a [TableColumn](https://docs.microsoft.com/javascript/api/excel/excel.tablecolumn?view=office-js) object indicates the index number of the column within the columns collection of the table.</span></span> <span data-ttu-id="1d8ca-124">**TableColumn** 对象的 **id** 属性包含用于标识列的唯一键。</span><span class="sxs-lookup"><span data-stu-id="1d8ca-124">The **id** property of a **TableColumn** object contains a unique key that identifies the column.</span></span>

### <a name="add-a-column-that-contains-static-values"></a><span data-ttu-id="1d8ca-125">添加包含静态值的列</span><span class="sxs-lookup"><span data-stu-id="1d8ca-125">Add a column that contains static values</span></span>

<span data-ttu-id="1d8ca-126">下面的代码示例将一个新列添加到名为 **Sample** 的工作表内的 **ExpensesTable** 表中。</span><span class="sxs-lookup"><span data-stu-id="1d8ca-126">The following code sample adds a new column to the table named **ExpensesTable** within the worksheet named **Sample**.</span></span> <span data-ttu-id="1d8ca-127">新列添加到表中所有现有列后面，并且包含一个标题（“星期几”），以及用于填充列中单元格的数据。</span><span class="sxs-lookup"><span data-stu-id="1d8ca-127">The new column is added after all existing columns in the table and contains a header ("Day of the Week") as well as data to populate the cells in the column.</span></span> <span data-ttu-id="1d8ca-128">如果在其中运行代码的 Excel 主机应用程序支持[要求集](https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets?view=office-js) **ExcelApi 1.2**，则为表中的当前数据设置最佳列宽和行高。</span><span class="sxs-lookup"><span data-stu-id="1d8ca-128">If the Excel host application where the code is running supports [requirement set](https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets?view=office-js) **ExcelApi 1.2**, the width of the columns and height of the rows are set to best fit the current data in the table.</span></span>

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

<span data-ttu-id="1d8ca-129">**包含新列的表**</span><span class="sxs-lookup"><span data-stu-id="1d8ca-129">**Table with new column**</span></span>

![Excel 中包含新列的表](../images/excel-tables-add-column.png)

### <a name="add-a-column-that-contains-formulas"></a><span data-ttu-id="1d8ca-131">添加包含公式的列</span><span class="sxs-lookup"><span data-stu-id="1d8ca-131">Add a column that contains formulas</span></span>

<span data-ttu-id="1d8ca-132">下面的代码示例将一个新列添加到名为 **Sample** 的工作表内的 **ExpensesTable** 表中。</span><span class="sxs-lookup"><span data-stu-id="1d8ca-132">The following code sample adds a new column to the table named **ExpensesTable** within the worksheet named **Sample**.</span></span> <span data-ttu-id="1d8ca-133">新列添加到表的末尾，包含标题（“日期类型”），并使用一个公式来填充列中的每个数据单元格。</span><span class="sxs-lookup"><span data-stu-id="1d8ca-133">The new column is added to the end of the table, contains a header ("Type of the Day"), and uses a formula to populate each data cell in the column.</span></span> <span data-ttu-id="1d8ca-134">如果在其中运行代码的 Excel 主机应用程序支持[要求集](https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets?view=office-js) **ExcelApi 1.2**，则为表中的当前数据设置最佳列宽和行高。</span><span class="sxs-lookup"><span data-stu-id="1d8ca-134">If the Excel host application where the code is running supports [requirement set](https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets?view=office-js) **ExcelApi 1.2**, the width of the columns and height of the rows are set to best fit the current data in the table.</span></span>

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

<span data-ttu-id="1d8ca-135">**包含新的计算列的表**</span><span class="sxs-lookup"><span data-stu-id="1d8ca-135">**Table with new calculated column**</span></span>

![Excel 中包含新的计算列的表](../images/excel-tables-add-calculated-column.png)

## <a name="update-column-name"></a><span data-ttu-id="1d8ca-137">更新列名称</span><span class="sxs-lookup"><span data-stu-id="1d8ca-137">Update column name</span></span>

<span data-ttu-id="1d8ca-p109">下面的代码示例将表格中第一列的名称更新为**购买日期**。如果运行代码的 Excel 主机应用支持[要求集](https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets?view=office-js) **ExcelApi 1.2**，那么列宽和行高会设置为最适应表格中的当前数据。</span><span class="sxs-lookup"><span data-stu-id="1d8ca-p109">The following code sample updates the name of the first column in the table to **Purchase date**. If the Excel host application where the code is running supports [requirement set](https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets?view=office-js) **ExcelApi 1.2**, the width of the columns and height of the rows are set to best fit the current data in the table.</span></span>

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

<span data-ttu-id="1d8ca-140">**包含新列名称的表格**</span><span class="sxs-lookup"><span data-stu-id="1d8ca-140">**Table with new column name**</span></span>

![Excel 中包含新的列名称的表](../images/excel-tables-update-column-name.png)

## <a name="get-data-from-a-table"></a><span data-ttu-id="1d8ca-142">从表中获取数据</span><span class="sxs-lookup"><span data-stu-id="1d8ca-142">Get data from a table</span></span>

<span data-ttu-id="1d8ca-143">下面的代码示例从名为 **Sample** 的工作表内的 **ExpensesTable** 表中读取数据，然后在同一工作表中的表下输出该数据。</span><span class="sxs-lookup"><span data-stu-id="1d8ca-143">The following code sample reads data from a table named **ExpensesTable** in the worksheet named **Sample** and then outputs that data below the table in the same worksheet.</span></span>

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

<span data-ttu-id="1d8ca-144">**表和数据输出**</span><span class="sxs-lookup"><span data-stu-id="1d8ca-144">**Table and data output**</span></span>

![Excel 中的表数据](../images/excel-tables-get-data.png)

## <a name="sort-data-in-a-table"></a><span data-ttu-id="1d8ca-146">在表中对数据进行排序</span><span class="sxs-lookup"><span data-stu-id="1d8ca-146">Sort data in a table</span></span>

<span data-ttu-id="1d8ca-147">下面的代码示例根据表中第四列的值，对表数据按降序进行排序。</span><span class="sxs-lookup"><span data-stu-id="1d8ca-147">The following code sample sorts table data in descending order according to the values in the fourth column of the table.</span></span>

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

<span data-ttu-id="1d8ca-148">**按金额排序的表数据（降序）**</span><span class="sxs-lookup"><span data-stu-id="1d8ca-148">**Table data sorted by Amount (descending)**</span></span>

![Excel 中的表数据](../images/excel-tables-sort.png)

## <a name="apply-filters-to-a-table"></a><span data-ttu-id="1d8ca-150">将筛选器应用于表</span><span class="sxs-lookup"><span data-stu-id="1d8ca-150">Apply filters to a table</span></span>

<span data-ttu-id="1d8ca-151">下面的代码示例将筛选器应用到表中的**金额**列和**类别**列。</span><span class="sxs-lookup"><span data-stu-id="1d8ca-151">The following code sample applies filters to the **Amount** column and the **Category** column within a table.</span></span> <span data-ttu-id="1d8ca-152">筛选器筛选的结果是，仅显示符合以下条件的行：**类别**为其中一个指定值且**金额**低于所有行的平均值。</span><span class="sxs-lookup"><span data-stu-id="1d8ca-152">As a result of the filters, only rows where **Category** is one of the specified values and **Amount** is below the average value for all rows is shown.</span></span>

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

<span data-ttu-id="1d8ca-153">**将筛选器应用于类别和金额的表数据**</span><span class="sxs-lookup"><span data-stu-id="1d8ca-153">**Table data with filters applied for Category and Amount**</span></span>

![Excel 中经过筛选的表数据](../images/excel-tables-filters-apply.png)

## <a name="clear-table-filters"></a><span data-ttu-id="1d8ca-155">清除表筛选器</span><span class="sxs-lookup"><span data-stu-id="1d8ca-155">Clear table filters</span></span>

<span data-ttu-id="1d8ca-156">下面的代码示例清除当前应用于表的所有筛选器。</span><span class="sxs-lookup"><span data-stu-id="1d8ca-156">The following code sample clears any filters currently applied on the table.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var expensesTable = sheet.tables.getItem("ExpensesTable");

    expensesTable.clearFilters();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="1d8ca-157">**没有应用任何筛选器的表数据**</span><span class="sxs-lookup"><span data-stu-id="1d8ca-157">**Table data with no filters applied**</span></span>

![Excel 中未经筛选的表数据](../images/excel-tables-filters-clear.png)

## <a name="get-the-visible-range-from-a-filtered-table"></a><span data-ttu-id="1d8ca-159">从筛选表中获取可见区域</span><span class="sxs-lookup"><span data-stu-id="1d8ca-159">Get the visible range from a filtered table</span></span>

<span data-ttu-id="1d8ca-160">下面的代码示例获取一个区域，其中只包含当前在指定表中可见的单元格数据，然后将该区域的值写入控制台。</span><span class="sxs-lookup"><span data-stu-id="1d8ca-160">The following code sample gets a range that contains data only for cells that are currently visible within the specified table, and then writes the values of that range to the console.</span></span> <span data-ttu-id="1d8ca-161">可以使用如下所示的 **getVisibleView()** 方法，在应用列筛选器时，都能获取表的可见内容。</span><span class="sxs-lookup"><span data-stu-id="1d8ca-161">You can use the **getVisibleView()** method as shown below to get the visible contents of a table whenever column filters have been applied.</span></span>

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

## <a name="format-a-table"></a><span data-ttu-id="1d8ca-162">设置表格式</span><span class="sxs-lookup"><span data-stu-id="1d8ca-162">Format a table</span></span>

<span data-ttu-id="1d8ca-163">下面的代码示例将格式应用于表。</span><span class="sxs-lookup"><span data-stu-id="1d8ca-163">The following code sample applies formatting to a table.</span></span> <span data-ttu-id="1d8ca-164">它为表的标题行、正文、第二行以及第一列指定不同的填充颜色。</span><span class="sxs-lookup"><span data-stu-id="1d8ca-164">It specifies different fill colors for the header row of the table, the body of the table, the second row of the table, and the first column of the table.</span></span> <span data-ttu-id="1d8ca-165">有关可以用来指定格式的属性的信息，请参阅 [RangeFormat 对象 (Excel JavaScript API)](https://docs.microsoft.com/javascript/api/excel/excel.rangeformat?view=office-js)。</span><span class="sxs-lookup"><span data-stu-id="1d8ca-165">For information about the properties you can use to specify format, see [RangeFormat Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.rangeformat?view=office-js).</span></span>

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

<span data-ttu-id="1d8ca-166">**应用格式设置的表**</span><span class="sxs-lookup"><span data-stu-id="1d8ca-166">**Table after formatting is applied**</span></span>

![Excel 中应用了格式设置的表](../images/excel-tables-formatting-after.png)

## <a name="convert-a-range-to-a-table"></a><span data-ttu-id="1d8ca-168">将区域转换为表</span><span class="sxs-lookup"><span data-stu-id="1d8ca-168">Convert a range to a table</span></span>

<span data-ttu-id="1d8ca-169">下面的代码示例创建一个数据区域，然后将该区域转换为表。</span><span class="sxs-lookup"><span data-stu-id="1d8ca-169">The following code sample creates a range of data and then converts that range to a table.</span></span>

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

<span data-ttu-id="1d8ca-170">**区域内的数据（在区域转换为表之前）**</span><span class="sxs-lookup"><span data-stu-id="1d8ca-170">**Data in the range (before the range is converted to a table)**</span></span>

![Excel 中区域内的数据](../images/excel-ranges.png)

<span data-ttu-id="1d8ca-172">**表中的数据（在区域转换为表之后）**</span><span class="sxs-lookup"><span data-stu-id="1d8ca-172">**Data in the table (after the range is converted to a table)**</span></span>

![Excel 中表的数据](../images/excel-tables-from-range.png)

## <a name="import-json-data-into-a-table"></a><span data-ttu-id="1d8ca-174">将 JSON 数据导入表</span><span class="sxs-lookup"><span data-stu-id="1d8ca-174">Import JSON data into a table</span></span>

<span data-ttu-id="1d8ca-175">下面的代码示例在名为 **Sample** 的工作表中创建一个表，然后使用定义了两行数据的 JSON 对象来填充表。</span><span class="sxs-lookup"><span data-stu-id="1d8ca-175">The following code sample creates a table in the worksheet named **Sample** and then populates the table by using a JSON object that defines two rows of data.</span></span> <span data-ttu-id="1d8ca-176">如果在其中运行代码的 Excel 主机应用程序支持[要求集](https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets?view=office-js) **ExcelApi 1.2**，则为表中的当前数据设置最佳列宽和行高。</span><span class="sxs-lookup"><span data-stu-id="1d8ca-176">If the Excel host application where the code is running supports [requirement set](https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets?view=office-js) **ExcelApi 1.2**, the width of the columns and height of the rows are set to best fit the current data in the table.</span></span>

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

<span data-ttu-id="1d8ca-177">**新建表**</span><span class="sxs-lookup"><span data-stu-id="1d8ca-177">**New table**</span></span>

![Excel 中的新表格](../images/excel-tables-create-from-json.png)

## <a name="see-also"></a><span data-ttu-id="1d8ca-179">另请参阅</span><span class="sxs-lookup"><span data-stu-id="1d8ca-179">See also</span></span>

- [<span data-ttu-id="1d8ca-180">使用 Excel JavaScript API 的基本编程概念</span><span class="sxs-lookup"><span data-stu-id="1d8ca-180">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)

