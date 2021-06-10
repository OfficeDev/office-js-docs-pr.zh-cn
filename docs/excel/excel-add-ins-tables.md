---
title: 使用 Excel JavaScript API 处理表格
description: 显示如何使用 JavaScript API 对表执行常见Excel示例。
ms.date: 06/07/2021
localization_priority: Normal
ms.openlocfilehash: a44a99e0ddc612342b292fd6e9d203799cde7b53
ms.sourcegitcommit: 5a151d4df81e5640363774406d0f329d6a0d3db8
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/09/2021
ms.locfileid: "52853996"
---
# <a name="work-with-tables-using-the-excel-javascript-api"></a><span data-ttu-id="94238-103">使用 Excel JavaScript API 处理表格</span><span class="sxs-lookup"><span data-stu-id="94238-103">Work with tables using the Excel JavaScript API</span></span>

<span data-ttu-id="94238-104">本文中的代码示例展示了如何使用 Excel JavaScript API 对表格执行常见任务。</span><span class="sxs-lookup"><span data-stu-id="94238-104">This article provides code samples that show how to perform common tasks with tables using the Excel JavaScript API.</span></span> <span data-ttu-id="94238-105">有关 和 对象支持的属性和方法的完整列表，请参阅适用于 Excel) 的 Table 对象 (JavaScript API 和适用于 Excel) 的 `Table` `TableCollection` [TableCollection 对象 (JavaScript API。 ](/javascript/api/excel/excel.tablecollection) [](/javascript/api/excel/excel.table)</span><span class="sxs-lookup"><span data-stu-id="94238-105">For the complete list of properties and methods that the `Table` and `TableCollection` objects support, see [Table Object (JavaScript API for Excel)](/javascript/api/excel/excel.table) and [TableCollection Object (JavaScript API for Excel)](/javascript/api/excel/excel.tablecollection).</span></span>

## <a name="create-a-table"></a><span data-ttu-id="94238-106">创建表</span><span class="sxs-lookup"><span data-stu-id="94238-106">Create a table</span></span>

<span data-ttu-id="94238-107">下面的代码示例在名为 **Sample** 的工作表中创建一个表。</span><span class="sxs-lookup"><span data-stu-id="94238-107">The following code sample creates a table in the worksheet named **Sample**.</span></span> <span data-ttu-id="94238-108">此表包含标题，并且包含四列和七行数据。</span><span class="sxs-lookup"><span data-stu-id="94238-108">The table has headers and contains four columns and seven rows of data.</span></span> <span data-ttu-id="94238-109">如果运行Excel的应用程序支持要求集 [](../reference/requirement-sets/excel-api-requirement-sets.md)**ExcelApi 1.2，** 则列宽和行高将设置为最适合表格中的当前数据。</span><span class="sxs-lookup"><span data-stu-id="94238-109">If the Excel application where the code is running supports [requirement set](../reference/requirement-sets/excel-api-requirement-sets.md) **ExcelApi 1.2**, the width of the columns and height of the rows are set to best fit the current data in the table.</span></span>

> [!NOTE]
> <span data-ttu-id="94238-110">若要指定表的名称，必须先创建表，然后设置其属性 `name` ，如以下示例所示。</span><span class="sxs-lookup"><span data-stu-id="94238-110">To specify a name for a table, you must first create the table and then set its `name` property, as shown in the following example.</span></span>

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

<span data-ttu-id="94238-111">**新建表**</span><span class="sxs-lookup"><span data-stu-id="94238-111">**New table**</span></span>

![Excel 中的新表](../images/excel-tables-create.png)

## <a name="add-rows-to-a-table"></a><span data-ttu-id="94238-113">向表添加行</span><span class="sxs-lookup"><span data-stu-id="94238-113">Add rows to a table</span></span>

<span data-ttu-id="94238-114">下面的代码示例将七个新行添加到名为 **Sample** 的工作表内的 **ExpensesTable** 表中。</span><span class="sxs-lookup"><span data-stu-id="94238-114">The following code sample adds seven new rows to the table named **ExpensesTable** within the worksheet named **Sample**.</span></span> <span data-ttu-id="94238-115">新行被添加到表的末尾。</span><span class="sxs-lookup"><span data-stu-id="94238-115">The new rows are added to the end of the table.</span></span> <span data-ttu-id="94238-116">如果运行Excel的应用程序支持要求集 [](../reference/requirement-sets/excel-api-requirement-sets.md)**ExcelApi 1.2，** 则列宽和行高将设置为最适合表格中的当前数据。</span><span class="sxs-lookup"><span data-stu-id="94238-116">If the Excel application where the code is running supports [requirement set](../reference/requirement-sets/excel-api-requirement-sets.md) **ExcelApi 1.2**, the width of the columns and height of the rows are set to best fit the current data in the table.</span></span>

> [!NOTE]
> <span data-ttu-id="94238-117">`index` [TableRow 对象的 属性](/javascript/api/excel/excel.tablerow)指示表的行集合中行的索引号。</span><span class="sxs-lookup"><span data-stu-id="94238-117">The `index` property of a [TableRow](/javascript/api/excel/excel.tablerow) object indicates the index number of the row within the rows collection of the table.</span></span> <span data-ttu-id="94238-118">`TableRow`对象不包含可用于标识行的唯一 `id` 键的属性。</span><span class="sxs-lookup"><span data-stu-id="94238-118">A `TableRow` object does not contain an `id` property that can be used as a unique key to identify the row.</span></span>

> [!WARNING]
> <span data-ttu-id="94238-119">从内容外接程序向表中添加行将导致内存泄漏。</span><span class="sxs-lookup"><span data-stu-id="94238-119">Adding rows to a table from a content add-in will result in a memory leak.</span></span> <span data-ttu-id="94238-120">有关[GitHub和其他#1415，](https://github.com/OfficeDev/office-js/issues/1415)请参阅"问题"报告。</span><span class="sxs-lookup"><span data-stu-id="94238-120">See [GitHub Issue #1415](https://github.com/OfficeDev/office-js/issues/1415) for current status and additional information.</span></span> 

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

<span data-ttu-id="94238-121">**包含新行的表**</span><span class="sxs-lookup"><span data-stu-id="94238-121">**Table with new rows**</span></span>

![Excel 中包含新行的表](../images/excel-tables-add-rows.png)

## <a name="add-a-column-to-a-table"></a><span data-ttu-id="94238-123">向表添加列</span><span class="sxs-lookup"><span data-stu-id="94238-123">Add a column to a table</span></span>

<span data-ttu-id="94238-p106">下面的示例演示如何向表添加列。 第一个示例使用静态值填充新列；第二个示例使用公式填充新列。</span><span class="sxs-lookup"><span data-stu-id="94238-p106">These examples show how to add a column to a table. The first example populates the new column with static values; the second example populates the new column with formulas.</span></span>

> [!NOTE]
> <span data-ttu-id="94238-p107">**TableColumn** 对象的 [index](/javascript/api/excel/excel.tablecolumn) 属性表示表格列集合内列的索引编号。 **TableColumn** 对象的 **id** 属性包含用于标识列的唯一键。</span><span class="sxs-lookup"><span data-stu-id="94238-p107">The **index** property of a [TableColumn](/javascript/api/excel/excel.tablecolumn) object indicates the index number of the column within the columns collection of the table. The **id** property of a **TableColumn** object contains a unique key that identifies the column.</span></span>

### <a name="add-a-column-that-contains-static-values"></a><span data-ttu-id="94238-128">添加包含静态值的列</span><span class="sxs-lookup"><span data-stu-id="94238-128">Add a column that contains static values</span></span>

<span data-ttu-id="94238-129">下面的代码示例将一个新列添加到名为 **Sample** 的工作表内的 **ExpensesTable** 表中。</span><span class="sxs-lookup"><span data-stu-id="94238-129">The following code sample adds a new column to the table named **ExpensesTable** within the worksheet named **Sample**.</span></span> <span data-ttu-id="94238-130">新列添加到表中所有现有列后面，并且包含一个标题（“星期几”），以及用于填充列中单元格的数据。</span><span class="sxs-lookup"><span data-stu-id="94238-130">The new column is added after all existing columns in the table and contains a header ("Day of the Week") as well as data to populate the cells in the column.</span></span> <span data-ttu-id="94238-131">如果运行Excel的应用程序支持要求集 [](../reference/requirement-sets/excel-api-requirement-sets.md)**ExcelApi 1.2，** 则列宽和行高将设置为最适合表格中的当前数据。</span><span class="sxs-lookup"><span data-stu-id="94238-131">If the Excel application where the code is running supports [requirement set](../reference/requirement-sets/excel-api-requirement-sets.md) **ExcelApi 1.2**, the width of the columns and height of the rows are set to best fit the current data in the table.</span></span>

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

<span data-ttu-id="94238-132">**包含新列的表**</span><span class="sxs-lookup"><span data-stu-id="94238-132">**Table with new column**</span></span>

![Excel 中包含新列的表](../images/excel-tables-add-column.png)

### <a name="add-a-column-that-contains-formulas&quot;></a><span data-ttu-id=&quot;94238-134&quot;>添加包含公式的列</span><span class=&quot;sxs-lookup&quot;><span data-stu-id=&quot;94238-134&quot;>Add a column that contains formulas</span></span>

<span data-ttu-id=&quot;94238-135&quot;>下面的代码示例将一个新列添加到名为 **Sample** 的工作表内的 **ExpensesTable** 表中。</span><span class=&quot;sxs-lookup&quot;><span data-stu-id=&quot;94238-135&quot;>The following code sample adds a new column to the table named **ExpensesTable** within the worksheet named **Sample**.</span></span> <span data-ttu-id=&quot;94238-136&quot;>新列添加到表的末尾，包含标题（“日期类型”），并使用一个公式来填充列中的每个数据单元格。</span><span class=&quot;sxs-lookup&quot;><span data-stu-id=&quot;94238-136&quot;>The new column is added to the end of the table, contains a header (&quot;Type of the Day"), and uses a formula to populate each data cell in the column.</span></span> <span data-ttu-id="94238-137">如果运行Excel的应用程序支持要求集 [](../reference/requirement-sets/excel-api-requirement-sets.md)**ExcelApi 1.2，** 则列宽和行高将设置为最适合表格中的当前数据。</span><span class="sxs-lookup"><span data-stu-id="94238-137">If the Excel application where the code is running supports [requirement set](../reference/requirement-sets/excel-api-requirement-sets.md) **ExcelApi 1.2**, the width of the columns and height of the rows are set to best fit the current data in the table.</span></span>

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

<span data-ttu-id="94238-138">**包含新的计算列的表**</span><span class="sxs-lookup"><span data-stu-id="94238-138">**Table with new calculated column**</span></span>

![Excel 中包含新的计算列的表](../images/excel-tables-add-calculated-column.png)

## <a name="resize-a-table-online-only"></a><span data-ttu-id="94238-140">调整表格大小 (仅联机) </span><span class="sxs-lookup"><span data-stu-id="94238-140">Resize a table (online-only)</span></span>

> [!NOTE]
> <span data-ttu-id="94238-141">`Table.resize`该方法当前仅在 ExcelApiOnline 1.1 中可用。</span><span class="sxs-lookup"><span data-stu-id="94238-141">The `Table.resize` method is currently only available in ExcelApiOnline 1.1.</span></span> <span data-ttu-id="94238-142">若要了解更多信息，请参阅[Excel JavaScript API 仅联机要求集](../reference/requirement-sets/excel-api-online-requirement-set.md)。</span><span class="sxs-lookup"><span data-stu-id="94238-142">To learn more, see [Excel JavaScript API online-only requirement set](../reference/requirement-sets/excel-api-online-requirement-set.md).</span></span>

<span data-ttu-id="94238-143">加载项可以调整表格大小，而无需向表格添加数据或更改单元格值。</span><span class="sxs-lookup"><span data-stu-id="94238-143">Your add-in can resize a table without adding data to the table or changing cell values.</span></span> <span data-ttu-id="94238-144">若要调整表格的大小，请使用 [Table.resize](/javascript/api/excel/excel.table#resize_newRange_) 方法。</span><span class="sxs-lookup"><span data-stu-id="94238-144">To resize a table, use the [Table.resize](/javascript/api/excel/excel.table#resize_newRange_) method.</span></span> <span data-ttu-id="94238-145">下面的代码示例演示如何调整表的大小。</span><span class="sxs-lookup"><span data-stu-id="94238-145">The following code sample shows how to resize a table.</span></span> <span data-ttu-id="94238-146">此代码示例使用本文前面创建表格部分 [](#create-a-table)中的 **ExpensesTable，** 将表的新范围设置 **为 A1：D20**。</span><span class="sxs-lookup"><span data-stu-id="94238-146">This code sample uses the **ExpensesTable** from the [Create a table](#create-a-table) section earlier in this article and sets the new range of the table to **A1:D20**.</span></span>

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
> <span data-ttu-id="94238-147">表格的新区域必须与原始区域重叠，并且表格标题 (或表格顶部) 必须在同一行中。</span><span class="sxs-lookup"><span data-stu-id="94238-147">The new range of the table must overlap with the original range, and the headers (or the top of the table) must be in the same row.</span></span>

<span data-ttu-id="94238-148">**调整大小后的表格**</span><span class="sxs-lookup"><span data-stu-id="94238-148">**Table after resize**</span></span> 

![包含多个空行的Excel](../images/excel-tables-resize.png)

## <a name="update-column-name"></a><span data-ttu-id="94238-150">更新列名称</span><span class="sxs-lookup"><span data-stu-id="94238-150">Update column name</span></span>

<span data-ttu-id="94238-151">下面的代码示例将表中第一列的名称更新为 **"购买日期"。**</span><span class="sxs-lookup"><span data-stu-id="94238-151">The following code sample updates the name of the first column in the table to **Purchase date**.</span></span> <span data-ttu-id="94238-152">如果运行Excel的应用程序支持要求集 [](../reference/requirement-sets/excel-api-requirement-sets.md)**ExcelApi 1.2，** 则列宽和行高将设置为最适合表格中的当前数据。</span><span class="sxs-lookup"><span data-stu-id="94238-152">If the Excel application where the code is running supports [requirement set](../reference/requirement-sets/excel-api-requirement-sets.md) **ExcelApi 1.2**, the width of the columns and height of the rows are set to best fit the current data in the table.</span></span>

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

<span data-ttu-id="94238-153">**包含新列名称的表格**</span><span class="sxs-lookup"><span data-stu-id="94238-153">**Table with new column name**</span></span>

![Excel 中包含新的列名称的表](../images/excel-tables-update-column-name.png)

## <a name="get-data-from-a-table"></a><span data-ttu-id="94238-155">从表中获取数据</span><span class="sxs-lookup"><span data-stu-id="94238-155">Get data from a table</span></span>

<span data-ttu-id="94238-156">下面的代码示例从名为 **Sample** 的工作表内的 **ExpensesTable** 表中读取数据，然后在同一工作表中的表下输出该数据。</span><span class="sxs-lookup"><span data-stu-id="94238-156">The following code sample reads data from a table named **ExpensesTable** in the worksheet named **Sample** and then outputs that data below the table in the same worksheet.</span></span>

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

<span data-ttu-id="94238-157">**表和数据输出**</span><span class="sxs-lookup"><span data-stu-id="94238-157">**Table and data output**</span></span>

![Excel 中的表数据](../images/excel-tables-get-data.png)

## <a name="detect-data-changes"></a><span data-ttu-id="94238-159">检测数据更改</span><span class="sxs-lookup"><span data-stu-id="94238-159">Detect data changes</span></span>

<span data-ttu-id="94238-160">外接程序可能需要回应对表中的数据进行更改的用户。</span><span class="sxs-lookup"><span data-stu-id="94238-160">Your add-in may need to react to users changing the data in a table.</span></span> <span data-ttu-id="94238-161">若要检测这些更改，你可以为表的 `onChanged` 事件[注册事件处理程序](excel-add-ins-events.md#register-an-event-handler)。</span><span class="sxs-lookup"><span data-stu-id="94238-161">To detect these changes, you can [register an event handler](excel-add-ins-events.md#register-an-event-handler) for the `onChanged` event of a table.</span></span> <span data-ttu-id="94238-162">当事件触发时，`onChanged` 事件的事件处理程序将收到 [TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs) 对象。</span><span class="sxs-lookup"><span data-stu-id="94238-162">Event handlers for the `onChanged` event receive a [TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs) object when the event fires.</span></span>

<span data-ttu-id="94238-163">`TableChangedEventArgs` 对象提供有关更改和来源的信息。</span><span class="sxs-lookup"><span data-stu-id="94238-163">The `TableChangedEventArgs` object provides information about the changes and the source.</span></span> <span data-ttu-id="94238-164">由于 `onChanged` 会在数据的格式或值发生变化时触发，因此让加载项检查值是否已实际更改可能很有用。</span><span class="sxs-lookup"><span data-stu-id="94238-164">Since `onChanged` fires when either the format or value of the data changes, it can be useful to have your add-in check if the values have actually changed.</span></span> <span data-ttu-id="94238-165">`details` 属性以 [ChangedEventDetail](/javascript/api/excel/excel.changedeventdetail) 的形式封装此信息。</span><span class="sxs-lookup"><span data-stu-id="94238-165">The `details` property encapsulates this information as a [ChangedEventDetail](/javascript/api/excel/excel.changedeventdetail).</span></span> <span data-ttu-id="94238-166">以下代码示例演示如何显示已更改的单元格的之前和之后的值及类型。</span><span class="sxs-lookup"><span data-stu-id="94238-166">The following code sample shows how to display the before and after values and types of a cell that has been changed.</span></span>

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

## <a name="sort-data-in-a-table"></a><span data-ttu-id="94238-167">对表格中的数据进行排序</span><span class="sxs-lookup"><span data-stu-id="94238-167">Sort data in a table</span></span>

<span data-ttu-id="94238-168">下面的代码示例根据表中第四列的值，对表数据按降序进行排序。</span><span class="sxs-lookup"><span data-stu-id="94238-168">The following code sample sorts table data in descending order according to the values in the fourth column of the table.</span></span>

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

<span data-ttu-id="94238-169">**按金额排序的表数据（降序）**</span><span class="sxs-lookup"><span data-stu-id="94238-169">**Table data sorted by Amount (descending)**</span></span>

![排序表中的Excel](../images/excel-tables-sort.png)

<span data-ttu-id="94238-171">在工作表中对数据进行排序时，会触发事件通知。</span><span class="sxs-lookup"><span data-stu-id="94238-171">When data is sorted in a worksheet, an event notification fires.</span></span> <span data-ttu-id="94238-172">要详细了解有关排序的事件以及加载项如何注册事件处理程序来响应此类事件，请参阅[处理排序事件](excel-add-ins-worksheets.md#handle-sorting-events)。</span><span class="sxs-lookup"><span data-stu-id="94238-172">To learn more about sort-related events and how your add-in can register event handlers to respond to such events, see [Handle sorting events](excel-add-ins-worksheets.md#handle-sorting-events).</span></span>

## <a name="apply-filters-to-a-table"></a><span data-ttu-id="94238-173">将筛选器应用于表</span><span class="sxs-lookup"><span data-stu-id="94238-173">Apply filters to a table</span></span>

<span data-ttu-id="94238-p116">下面的代码示例将筛选器应用到表中的 **金额** 列和 **类别** 列。 筛选器筛选的结果是，仅显示符合以下条件的行：**类别** 为其中一个指定值且 **金额** 低于所有行的平均值。</span><span class="sxs-lookup"><span data-stu-id="94238-p116">The following code sample applies filters to the **Amount** column and the **Category** column within a table. As a result of the filters, only rows where **Category** is one of the specified values and **Amount** is below the average value for all rows is shown.</span></span>

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

<span data-ttu-id="94238-176">**将筛选器应用于类别和金额的表数据**</span><span class="sxs-lookup"><span data-stu-id="94238-176">**Table data with filters applied for Category and Amount**</span></span>

![Excel 中经过筛选的表数据](../images/excel-tables-filters-apply.png)

## <a name="clear-table-filters"></a><span data-ttu-id="94238-178">清除表筛选器</span><span class="sxs-lookup"><span data-stu-id="94238-178">Clear table filters</span></span>

<span data-ttu-id="94238-179">下面的代码示例清除当前应用于表的所有筛选器。</span><span class="sxs-lookup"><span data-stu-id="94238-179">The following code sample clears any filters currently applied on the table.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var expensesTable = sheet.tables.getItem("ExpensesTable");

    expensesTable.clearFilters();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="94238-180">**没有应用任何筛选器的表数据**</span><span class="sxs-lookup"><span data-stu-id="94238-180">**Table data with no filters applied**</span></span>

![Excel 中未经筛选的表数据](../images/excel-tables-filters-clear.png)

## <a name="get-the-visible-range-from-a-filtered-table"></a><span data-ttu-id="94238-182">从筛选表中获取可见区域</span><span class="sxs-lookup"><span data-stu-id="94238-182">Get the visible range from a filtered table</span></span>

<span data-ttu-id="94238-183">下面的代码示例获取一个区域，其中只包含当前在指定表中可见的单元格数据，然后将该区域的值写入控制台。</span><span class="sxs-lookup"><span data-stu-id="94238-183">The following code sample gets a range that contains data only for cells that are currently visible within the specified table, and then writes the values of that range to the console.</span></span> <span data-ttu-id="94238-184">可以使用如下所示的方法，在应用列筛选器后获取表 `getVisibleView()` 的可见内容。</span><span class="sxs-lookup"><span data-stu-id="94238-184">You can use the `getVisibleView()` method as shown below to get the visible contents of a table whenever column filters have been applied.</span></span>

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

## <a name="autofilter"></a><span data-ttu-id="94238-185">AutoFilter</span><span class="sxs-lookup"><span data-stu-id="94238-185">AutoFilter</span></span>

<span data-ttu-id="94238-186">加载项可使用表的 [AutoFilter](/javascript/api/excel/excel.autofilter) 对象筛选数据。</span><span class="sxs-lookup"><span data-stu-id="94238-186">An add-in can use the table's [AutoFilter](/javascript/api/excel/excel.autofilter) object to filter data.</span></span> <span data-ttu-id="94238-187">`AutoFilter` 对象是表或范围的整个筛选结构。</span><span class="sxs-lookup"><span data-stu-id="94238-187">An `AutoFilter` object is the entire filter structure of a table or range.</span></span> <span data-ttu-id="94238-188">本文之前讨论的所有筛选操作均与 auto-filter 兼容。</span><span class="sxs-lookup"><span data-stu-id="94238-188">All of the filter operations discussed earlier in this article are compatible with the auto-filter.</span></span> <span data-ttu-id="94238-189">通过单一访问点可以轻松访问和管理多个筛选器。</span><span class="sxs-lookup"><span data-stu-id="94238-189">The single access point does make it easier to access and manage multiple filters.</span></span>

<span data-ttu-id="94238-190">以下代码示例显示与[之前的代码示例相同的数据筛选](#apply-filters-to-a-table)，但完全通过 auto-filter 完成。</span><span class="sxs-lookup"><span data-stu-id="94238-190">The following code sample shows the same [data filtering as the earlier code sample](#apply-filters-to-a-table), but done entirely through the auto-filter.</span></span>

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

<span data-ttu-id="94238-191">`AutoFilter` 也可应用于工作表级别的范围。</span><span class="sxs-lookup"><span data-stu-id="94238-191">An `AutoFilter` can also be applied to a range at the worksheet level.</span></span> <span data-ttu-id="94238-192">有关详细信息，请参阅[使用 Excel JavaScript API 处理工作表](excel-add-ins-worksheets.md#filter-data)。</span><span class="sxs-lookup"><span data-stu-id="94238-192">See [Work with worksheets using the Excel JavaScript API](excel-add-ins-worksheets.md#filter-data) for more information.</span></span>

## <a name="format-a-table"></a><span data-ttu-id="94238-193">设置表格式</span><span class="sxs-lookup"><span data-stu-id="94238-193">Format a table</span></span>

<span data-ttu-id="94238-p120">下面的代码示例将格式应用于表。 它为表的标题行、正文、第二行以及第一列指定不同的填充颜色。 有关可以用来指定格式的属性的信息，请参阅 [RangeFormat 对象 (Excel JavaScript API)](/javascript/api/excel/excel.rangeformat)。</span><span class="sxs-lookup"><span data-stu-id="94238-p120">The following code sample applies formatting to a table. It specifies different fill colors for the header row of the table, the body of the table, the second row of the table, and the first column of the table. For information about the properties you can use to specify format, see [RangeFormat Object (JavaScript API for Excel)](/javascript/api/excel/excel.rangeformat).</span></span>

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

<span data-ttu-id="94238-197">**应用格式设置的表**</span><span class="sxs-lookup"><span data-stu-id="94238-197">**Table after formatting is applied**</span></span>

![Excel 中应用了格式设置的表](../images/excel-tables-formatting-after.png)

## <a name="convert-a-range-to-a-table"></a><span data-ttu-id="94238-199">将区域转换为表</span><span class="sxs-lookup"><span data-stu-id="94238-199">Convert a range to a table</span></span>

<span data-ttu-id="94238-200">下面的代码示例创建一个数据区域，然后将该区域转换为表。</span><span class="sxs-lookup"><span data-stu-id="94238-200">The following code sample creates a range of data and then converts that range to a table.</span></span>

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

<span data-ttu-id="94238-201">**内的数据（在区域转换为表之前）**</span><span class="sxs-lookup"><span data-stu-id="94238-201">**Data in the range (before the range is converted to a table)**</span></span>

![Excel 中区域内的数据](../images/excel-ranges.png)

<span data-ttu-id="94238-203">**表中的数据（在区域转换为表之后）**</span><span class="sxs-lookup"><span data-stu-id="94238-203">**Data in the table (after the range is converted to a table)**</span></span>

![Excel 中表的数据](../images/excel-tables-from-range.png)

## <a name="import-json-data-into-a-table"></a><span data-ttu-id="94238-205">将 JSON 数据导入表</span><span class="sxs-lookup"><span data-stu-id="94238-205">Import JSON data into a table</span></span>

<span data-ttu-id="94238-206">下面的代码示例在名为 **Sample** 的工作表中创建一个表，然后使用定义了两行数据的 JSON 对象来填充表。</span><span class="sxs-lookup"><span data-stu-id="94238-206">The following code sample creates a table in the worksheet named **Sample** and then populates the table by using a JSON object that defines two rows of data.</span></span> <span data-ttu-id="94238-207">如果运行Excel的应用程序支持要求集 [](../reference/requirement-sets/excel-api-requirement-sets.md)**ExcelApi 1.2，** 则列宽和行高将设置为最适合表格中的当前数据。</span><span class="sxs-lookup"><span data-stu-id="94238-207">If the Excel application where the code is running supports [requirement set](../reference/requirement-sets/excel-api-requirement-sets.md) **ExcelApi 1.2**, the width of the columns and height of the rows are set to best fit the current data in the table.</span></span>

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

<span data-ttu-id="94238-208">**新建表**</span><span class="sxs-lookup"><span data-stu-id="94238-208">**New table**</span></span>

![导入的 JSON 数据的新表Excel](../images/excel-tables-create-from-json.png)

## <a name="see-also"></a><span data-ttu-id="94238-210">另请参阅</span><span class="sxs-lookup"><span data-stu-id="94238-210">See also</span></span>

- [<span data-ttu-id="94238-211">Excel 加载项中的 Word JavaScript 对象模型</span><span class="sxs-lookup"><span data-stu-id="94238-211">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
