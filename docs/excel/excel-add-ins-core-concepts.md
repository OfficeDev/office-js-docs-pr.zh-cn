---
title: Excel JavaScript API 基本编程概念
description: 使用 Excel JavaScript API 生成 Excel 加载项。
ms.date: 07/28/2020
localization_priority: Priority
ms.openlocfilehash: dde7dc66e0746fc4d9cf91ed3df824fab05c109d
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292588"
---
# <a name="fundamental-programming-concepts-with-the-excel-javascript-api"></a><span data-ttu-id="60a8c-103">Excel JavaScript API 基本编程概念</span><span class="sxs-lookup"><span data-stu-id="60a8c-103">Fundamental programming concepts with the Excel JavaScript API</span></span>

<span data-ttu-id="60a8c-104">本文介绍了如何使用 [Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) 生成 Excel 2016 或更高版本的加载项。</span><span class="sxs-lookup"><span data-stu-id="60a8c-104">This article describes how to use the [Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) to build add-ins for Excel 2016 or later.</span></span> <span data-ttu-id="60a8c-105">它引入了一些核心概念，这些概念是使用 API 的基础，并为执行特定任务提供指导，如读取或写入较大区域、更新区域内的所有单元格等等。</span><span class="sxs-lookup"><span data-stu-id="60a8c-105">It introduces core concepts that are fundamental to using the API and provides guidance for performing specific tasks such as reading or writing to a large range, updating all cells in range, and more.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="60a8c-106">请参阅[使用特定于应用程序的 API 模型](../develop/application-specific-api-model.md)，以了解 Excel API 的异步性质以及它们如何与工作簿协同工作。</span><span class="sxs-lookup"><span data-stu-id="60a8c-106">See [Using the application-specific API model](../develop/application-specific-api-model.md) to learn about the asynchronous nature of the Excel APIs and how they work with the workbook.</span></span>  

## <a name="officejs-apis-for-excel"></a><span data-ttu-id="60a8c-107">适用于 Excel 的 Office.js API</span><span class="sxs-lookup"><span data-stu-id="60a8c-107">Office.js APIs for Excel</span></span>

<span data-ttu-id="60a8c-108">Excel 加载项通过使用适 Office JavaScript API 与 Excel 中的对象进行交互，JavaScript API包括两个 JavaScript 对象模型：</span><span class="sxs-lookup"><span data-stu-id="60a8c-108">An Excel add-in interacts with objects in Excel by using the Office JavaScript API, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="60a8c-109">**Excel JavaScript API**：[Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) 随 Office 2016 一起引入，提供了强类型的对象，可用于访问工作表、区域、表格、图表等。</span><span class="sxs-lookup"><span data-stu-id="60a8c-109">**Excel JavaScript API**: Introduced with Office 2016, the [Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) provides strongly-typed objects that you can use to access worksheets, ranges, tables, charts, and more.</span></span>

* <span data-ttu-id="60a8c-110">**通用 API**：[通用 API](/javascript/api/office) 随 Office 2013 引入，可用于访问多种类型的 Office 应用程序中常见的 UI、对话框和客户端设置等功能。</span><span class="sxs-lookup"><span data-stu-id="60a8c-110">**Common APIs**: Introduced with Office 2013, the [Common API](/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications.</span></span>

<span data-ttu-id="60a8c-111">你可能会使用 Excel JavaScript API 开发面向 Excel 2016 或更高版本的加载项中的大部分功能，同时还可以使用通用 API 中的对象。</span><span class="sxs-lookup"><span data-stu-id="60a8c-111">While you'll likely use the Excel JavaScript API to develop the majority of functionality in add-ins that target Excel 2016 or later, you'll also use objects in the Common API.</span></span> <span data-ttu-id="60a8c-112">例如：</span><span class="sxs-lookup"><span data-stu-id="60a8c-112">For example:</span></span>

* <span data-ttu-id="60a8c-p103">[Context](/javascript/api/office/office.context)：`Context` 对象表示加载项的运行时环境，并提供对 API 关键对象的访问权限。 它由工作簿配置详细信息（如 `contentLanguage` 和 `officeTheme`）组成，并提供有关加载项的运行时环境（如 `host` 和 `platform`）的信息。 此外，它还提供了 `requirements.isSetSupported()` 方法，可用于检查运行加载项的 Excel 应用程序是否支持指定的要求集。</span><span class="sxs-lookup"><span data-stu-id="60a8c-p103">[Context](/javascript/api/office/office.context): The `Context` object represents the runtime environment of the add-in and provides access to key objects of the API. It consists of workbook configuration details such as `contentLanguage` and `officeTheme` and also provides information about the add-in's runtime environment such as `host` and `platform`. Additionally, it provides the `requirements.isSetSupported()` method, which you can use to check whether the specified requirement set is supported by the Excel application where the add-in is running.</span></span>
* <span data-ttu-id="60a8c-116">[Document](/javascript/api/office/office.document)：`Document` 对象提供 `getFileAsync()` 方法，用于下载运行加载项的 Excel 文件。</span><span class="sxs-lookup"><span data-stu-id="60a8c-116">[Document](/javascript/api/office/office.document): The `Document` object provides the `getFileAsync()` method, which you can use to download the Excel file where the add-in is running.</span></span>

<span data-ttu-id="60a8c-117">下图说明了可能使用 Excel JavaScript API 或公共 API 的情况。</span><span class="sxs-lookup"><span data-stu-id="60a8c-117">The following image illustrates when you might use the Excel JavaScript API or the Common APIs.</span></span>

![Excel JS API 和公共 API 之间差异的图像](../images/excel-js-api-common-api.png)

## <a name="object-model"></a><span data-ttu-id="60a8c-119">对象模型</span><span class="sxs-lookup"><span data-stu-id="60a8c-119">Object model</span></span>

<span data-ttu-id="60a8c-120">若要了解 Excel API，则必须了解工作簿的各个组件之间如何相互关联。</span><span class="sxs-lookup"><span data-stu-id="60a8c-120">To understand the Excel APIs, you must understand how the components of a workbook are related to one another.</span></span>

* <span data-ttu-id="60a8c-121">一个 **Workbook** 包含一个或多个 **Worksheet**。</span><span class="sxs-lookup"><span data-stu-id="60a8c-121">A **Workbook** contains one or more **Worksheets**.</span></span>
* <span data-ttu-id="60a8c-122">**Worksheet** 可通过 **Range** 对象访问单元格。</span><span class="sxs-lookup"><span data-stu-id="60a8c-122">A **Worksheet** gives access to cells through **Range** objects.</span></span>
* <span data-ttu-id="60a8c-123">**Range** 代表一组连续的单元格。</span><span class="sxs-lookup"><span data-stu-id="60a8c-123">A **Range** represents a group of contiguous cells.</span></span>
* <span data-ttu-id="60a8c-124">**Range** 用于创建和放置 **Table**、**Chart** 和 **Shape** 以及其他数据可视化或组织对象。</span><span class="sxs-lookup"><span data-stu-id="60a8c-124">**Ranges** are used to create and place **Tables**, **Charts**, **Shapes**, and other data visualization or organization objects.</span></span>
* <span data-ttu-id="60a8c-125">**Worksheet** 包含单个工作表中存在的那些数据对象的集合。</span><span class="sxs-lookup"><span data-stu-id="60a8c-125">A **Worksheet** contains collections of those data objects that are present in the individual sheet.</span></span>
* <span data-ttu-id="60a8c-126">**Workbook** 包含整个 **Workbook** 的某些数据对象（例如，**Table**）的集合。</span><span class="sxs-lookup"><span data-stu-id="60a8c-126">**Workbooks** contain collections of some of those data objects (such as **Tables**) for the entire **Workbook**.</span></span>

### <a name="ranges"></a><span data-ttu-id="60a8c-127">Range</span><span class="sxs-lookup"><span data-stu-id="60a8c-127">Ranges</span></span>

<span data-ttu-id="60a8c-128">Range 是工作簿中的一组连续单元格。</span><span class="sxs-lookup"><span data-stu-id="60a8c-128">A range is a group of contiguous cells in the workbook.</span></span> <span data-ttu-id="60a8c-129">加载项通常使用 A1 样式表示法（例如，对于 **B** 列和第 **3** 行中单个单元格，即 **B3** 或从 **C** 列至 **F** 列和第 **2** 行至第 **4** 行的单元格，即 **C2:F4**）来定义范围。</span><span class="sxs-lookup"><span data-stu-id="60a8c-129">Add-ins typically use A1-style notation (e.g. **B3** for the single cell in column **B** and row **3** or **C2:F4** for the cells from columns **C** through **F** and rows **2** through **4**) to define ranges.</span></span>

<span data-ttu-id="60a8c-130">Range 具有三个核心属性：`values`、`formulas` 和 `format`。</span><span class="sxs-lookup"><span data-stu-id="60a8c-130">Ranges have three core properties: `values`, `formulas`, and `format`.</span></span> <span data-ttu-id="60a8c-131">这些属性获取或设置单元格值、要计算的公式以及单元格的视觉对象格式设置。</span><span class="sxs-lookup"><span data-stu-id="60a8c-131">These properties get or set the cell values, formulas to be evaluated, and the visual formatting of the cells.</span></span>

#### <a name="range-sample"></a><span data-ttu-id="60a8c-132">Range 示例</span><span class="sxs-lookup"><span data-stu-id="60a8c-132">Range sample</span></span>

<span data-ttu-id="60a8c-133">以下示例显示了如何创建销售记录。</span><span class="sxs-lookup"><span data-stu-id="60a8c-133">The following sample shows how to create sales records.</span></span> <span data-ttu-id="60a8c-134">此函数使用 `Range` 对象来设置值、公式和格式。</span><span class="sxs-lookup"><span data-stu-id="60a8c-134">This function uses `Range` objects to set the values, formulas, and formats.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();

    // Create the headers and format them to stand out.
    var headers = [
      ["Product", "Quantity", "Unit Price", "Totals"]
    ];
    var headerRange = sheet.getRange("B2:E2");
    headerRange.values = headers;
    headerRange.format.fill.color = "#4472C4";
    headerRange.format.font.color = "white";

    // Create the product data rows.
    var productData = [
      ["Almonds", 6, 7.5],
      ["Coffee", 20, 34.5],
      ["Chocolate", 10, 9.56],
    ];
    var dataRange = sheet.getRange("B3:D5");
    dataRange.values = productData;

    // Create the formulas to total the amounts sold.
    var totalFormulas = [
      ["=C3 * D3"],
      ["=C4 * D4"],
      ["=C5 * D5"],
      ["=SUM(E3:E5)"]
    ];
    var totalRange = sheet.getRange("E3:E6");
    totalRange.formulas = totalFormulas;
    totalRange.format.font.bold = true;

    // Display the totals as US dollar amounts.
    totalRange.numberFormat = [["$0.00"]];

    return context.sync();
});
```

<span data-ttu-id="60a8c-135">此示例将在当前工作表中创建以下数据：</span><span class="sxs-lookup"><span data-stu-id="60a8c-135">This sample creates the following data in the current worksheet:</span></span>

![显示值行、公式列和格式化标题的销售记录。](../images/excel-overview-range-sample.png)

### <a name="charts-tables-and-other-data-objects"></a><span data-ttu-id="60a8c-137">Chart、Table 和其他数据对象</span><span class="sxs-lookup"><span data-stu-id="60a8c-137">Charts, tables, and other data objects</span></span>

<span data-ttu-id="60a8c-138">Excel JavaScript API 可以在 Excel 中创建和设置数据结构和可视化效果。</span><span class="sxs-lookup"><span data-stu-id="60a8c-138">The Excel JavaScript APIs can create and manipulate the data structures and visualizations within Excel.</span></span> <span data-ttu-id="60a8c-139">Table 和 Chart 是最常用的两个对象，但是 API 支持数据透视表、形状和图像等。</span><span class="sxs-lookup"><span data-stu-id="60a8c-139">Tables and charts are two of the more commonly used objects, but the APIs support PivotTables, shapes, images, and more.</span></span>

#### <a name="creating-a-table"></a><span data-ttu-id="60a8c-140">创建表</span><span class="sxs-lookup"><span data-stu-id="60a8c-140">Creating a table</span></span>

<span data-ttu-id="60a8c-141">通过使用数据填充范围创建表。</span><span class="sxs-lookup"><span data-stu-id="60a8c-141">Create tables by using data-filled ranges.</span></span> <span data-ttu-id="60a8c-142">会将格式设置和表控件（如筛选器）自动应用到该范围。</span><span class="sxs-lookup"><span data-stu-id="60a8c-142">Formatting and table controls (such as filters) are automatically applied to the range.</span></span>

<span data-ttu-id="60a8c-143">以下示例使用上一个示例中的范围创建了一个表。</span><span class="sxs-lookup"><span data-stu-id="60a8c-143">The following sample creates a table using the ranges from the previous sample.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.tables.add("B2:E5", true);
    return context.sync();
});
```

<span data-ttu-id="60a8c-144">在包含之前数据的工作表上使用此示例代码将创建下表：</span><span class="sxs-lookup"><span data-stu-id="60a8c-144">Using this sample code on the worksheet with the previous data creates the following table:</span></span>

![使用之前的销售记录制成的表。](../images/excel-overview-table-sample.png)

#### <a name="creating-a-chart"></a><span data-ttu-id="60a8c-146">创建图表</span><span class="sxs-lookup"><span data-stu-id="60a8c-146">Creating a chart</span></span>

<span data-ttu-id="60a8c-147">创建图表以直观显示某个范围内的数据。</span><span class="sxs-lookup"><span data-stu-id="60a8c-147">Create charts to visualize the data in a range.</span></span> <span data-ttu-id="60a8c-148">该 API 支持数十种图表类型，每种都可以根据需要进行自定义。</span><span class="sxs-lookup"><span data-stu-id="60a8c-148">The APIs support dozens of chart varieties, each of which can be customized to suit your needs.</span></span>

<span data-ttu-id="60a8c-149">下面的示例为三个项目创建一个简单的柱形图，并将其置于工作表顶部下方 100 像素处。</span><span class="sxs-lookup"><span data-stu-id="60a8c-149">The following sample creates a simple column chart for three items and places it 100 pixels below the top of the worksheet.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var chart = sheet.charts.add(Excel.ChartType.columnStacked, sheet.getRange("B3:C5"));
    chart.top = 100;
    return context.sync();
});
```

<span data-ttu-id="60a8c-150">在工作表上使用上一个表运行此示例将创建以下图表：</span><span class="sxs-lookup"><span data-stu-id="60a8c-150">Running this sample on the worksheet with the previous table creates the following chart:</span></span>

![一个柱形图，显示上一个销售记录中三个项目的数量。](../images/excel-overview-chart-sample.png)

## <a name="run-options"></a><span data-ttu-id="60a8c-152">运行选项</span><span class="sxs-lookup"><span data-stu-id="60a8c-152">Run options</span></span>

<span data-ttu-id="60a8c-153">`Excel.run` 包含需要使用 [RunOptions](/javascript/api/excel/excel.runoptions) 对象的重载。</span><span class="sxs-lookup"><span data-stu-id="60a8c-153">`Excel.run` has an overload that takes in a [RunOptions](/javascript/api/excel/excel.runoptions) object.</span></span> <span data-ttu-id="60a8c-154">这包含一组影响函数运行时平台行为的属性。</span><span class="sxs-lookup"><span data-stu-id="60a8c-154">This contains a set of properties that affect platform behavior when the function runs.</span></span> <span data-ttu-id="60a8c-155">目前，支持以下属性：</span><span class="sxs-lookup"><span data-stu-id="60a8c-155">The following property is currently supported:</span></span>

* <span data-ttu-id="60a8c-156">`delayForCellEdit`：确定 Excel 是否将批处理请求延迟到用户退出单元格编辑模式时执行。</span><span class="sxs-lookup"><span data-stu-id="60a8c-156">`delayForCellEdit`: Determines whether Excel delays the batch request until the user exits cell edit mode.</span></span> <span data-ttu-id="60a8c-157">若为 **true**，批处理请求延迟到用户退出单元格编辑模式时执行。</span><span class="sxs-lookup"><span data-stu-id="60a8c-157">When **true**, the batch request is delayed and runs when the user exits cell edit mode.</span></span> <span data-ttu-id="60a8c-158">若为 **false**，批处理请求会在用户处于单元格编辑模式时（导致无法访问用户的错误出现）自动失败。</span><span class="sxs-lookup"><span data-stu-id="60a8c-158">When **false**, the batch request automatically fails if the user is in cell edit mode (causing an error to reach the user).</span></span> <span data-ttu-id="60a8c-159">未指定 `delayForCellEdit` 属性的默认行为等同于此属性为 **false**。</span><span class="sxs-lookup"><span data-stu-id="60a8c-159">The default behavior with no `delayForCellEdit` property specified is equivalent to when it is **false**.</span></span>

```js
Excel.run({ delayForCellEdit: true }, function (context) { ... })
```

## <a name="null-or-blank-property-values"></a><span data-ttu-id="60a8c-160">null 或空属性值</span><span class="sxs-lookup"><span data-stu-id="60a8c-160">null or blank property values</span></span>

<span data-ttu-id="60a8c-161">`null` 和空字符串在 Excel JavaScript API 中具有特殊含义。</span><span class="sxs-lookup"><span data-stu-id="60a8c-161">`null` and empty strings have special implications in the Excel JavaScript APIs.</span></span> <span data-ttu-id="60a8c-162">它们用于表示空单元格、无格式或默认值。</span><span class="sxs-lookup"><span data-stu-id="60a8c-162">They're used to represent empty cells, no formatting, or default values.</span></span> <span data-ttu-id="60a8c-163">本节详细介绍了在获取和设置属性时如何使用 `null` 和空字符串。</span><span class="sxs-lookup"><span data-stu-id="60a8c-163">This section details the use of `null` and empty string when getting and setting properties.</span></span>

### <a name="null-input-in-2-d-array"></a><span data-ttu-id="60a8c-164">二维数组中的 null 输入</span><span class="sxs-lookup"><span data-stu-id="60a8c-164">null input in 2-D Array</span></span>

<span data-ttu-id="60a8c-p113">在 Excel 中，一个区域由一个二维数组表示，其中第一个维度是行，第二个维度是列。 若要仅为某个区域内的特定单元格设置值、数字格式或公式，请指定二维数组中这些单元格的值、数字格式或公式，并为二维数组中的所有其他单元格指定 `null`。</span><span class="sxs-lookup"><span data-stu-id="60a8c-p113">In Excel, a range is represented by a 2-D array, where the first dimension is rows and the second dimension is columns. To set values, number format, or formula for only specific cells within a range, specify the values, number format, or formula for those cells in the 2-D array, and specify `null` for all other cells in the 2-D array.</span></span>

<span data-ttu-id="60a8c-p114">例如，要更新一个区域内某一个单元格的数字格式，并保留该区域内所有其他单元格的现有数字格式，可指定要更新的单元格的新数字格式，并为所有其他单元格指定 `null`。 下面的代码段为该区域内的第四个单元格设置了一个新的数字格式，并保留该区域内前三个单元格的数字格式不变。</span><span class="sxs-lookup"><span data-stu-id="60a8c-p114">For example, to update the number format for only one cell within a range, and retain the existing number format for all other cells in the range, specify the new number format for the cell to update, and specify `null` for all other cells. The following code snippet sets a new number format for the fourth cell in the range, and leaves the number format unchanged for the first three cells in the range.</span></span>

```js
range.values = [['Eurasia', '29.96', '0.25', '15-Feb' ]];
range.numberFormat = [[null, null, null, 'm/d/yyyy;@']];
```

### <a name="null-input-for-a-property"></a><span data-ttu-id="60a8c-169">属性的 null 输入</span><span class="sxs-lookup"><span data-stu-id="60a8c-169">null input for a property</span></span>

<span data-ttu-id="60a8c-p115">`null` 不是单个属性的有效输入。例如，下面的代码片段无效，因为区域的 `values` 属性不能设置为 `null`。</span><span class="sxs-lookup"><span data-stu-id="60a8c-p115">`null` is not a valid input for single property. For example, the following code snippet is not valid, as the `values` property of the range cannot be set to `null`.</span></span>

```js
range.values = null;
```

<span data-ttu-id="60a8c-172">同样，下面的代码片段也无效，因为 `null` 不是 `color` 属性的有效值。</span><span class="sxs-lookup"><span data-stu-id="60a8c-172">Likewise, the following code snippet is not valid, as `null` is not a valid value for the `color` property.</span></span>

```js
range.format.fill.color =  null;
```

### <a name="null-property-values-in-the-response"></a><span data-ttu-id="60a8c-173">响应中的 null 属性值</span><span class="sxs-lookup"><span data-stu-id="60a8c-173">null property values in the response</span></span>

<span data-ttu-id="60a8c-p116">如果指定区域内存在不同的值，诸如 `size` 和 `color` 等格式化属性将在响应中包含 `null` 值。 例如，如果你检索某个区域并加载其 `format.font.color` 属性：</span><span class="sxs-lookup"><span data-stu-id="60a8c-p116">Formatting properties such as `size` and `color` will contain `null` values in the response when different values exist in the specified range. For example, if you retrieve a range and load its `format.font.color` property:</span></span>

* <span data-ttu-id="60a8c-176">如果区域中的所有单元格都具有相同的字体颜色，则 `range.format.font.color` 会指定该颜色。</span><span class="sxs-lookup"><span data-stu-id="60a8c-176">If all cells in the range have the same font color, `range.format.font.color` specifies that color.</span></span>
* <span data-ttu-id="60a8c-177">如果该区域内存在多种字体颜色，则 `range.format.font.color` 为 `null`。</span><span class="sxs-lookup"><span data-stu-id="60a8c-177">If multiple font colors are present within the range, `range.format.font.color` is `null`.</span></span>

### <a name="blank-input-for-a-property"></a><span data-ttu-id="60a8c-178">属性的空白输入</span><span class="sxs-lookup"><span data-stu-id="60a8c-178">Blank input for a property</span></span>

<span data-ttu-id="60a8c-p117">如果为属性指定空白值（即两个引号之间没有空格 `''`），它会被解释为属性清除或重置指令。例如：</span><span class="sxs-lookup"><span data-stu-id="60a8c-p117">When you specify a blank value for a property (i.e., two quotation marks with no space in-between `''`), it will be interpreted as an instruction to clear or reset the property. For example:</span></span>

* <span data-ttu-id="60a8c-181">如果为区域的 `values` 属性指定空白值，此区域的内容会被清除。</span><span class="sxs-lookup"><span data-stu-id="60a8c-181">If you specify a blank value for the `values` property of a range, the content of the range is cleared.</span></span>
* <span data-ttu-id="60a8c-182">如果为 `numberFormat` 属性指定一个空值，则数字格式会重置为 `General`。</span><span class="sxs-lookup"><span data-stu-id="60a8c-182">If you specify a blank value for the `numberFormat` property, the number format is reset to `General`.</span></span>
* <span data-ttu-id="60a8c-183">如果为 `formula` 属性和 `formulaLocale` 属性指定一个空值，则公式值将被清除。</span><span class="sxs-lookup"><span data-stu-id="60a8c-183">If you specify a blank value for the `formula` property and `formulaLocale` property, the formula values are cleared.</span></span>

### <a name="blank-property-values-in-the-response"></a><span data-ttu-id="60a8c-184">响应中的空属性值</span><span class="sxs-lookup"><span data-stu-id="60a8c-184">Blank property values in the response</span></span>

<span data-ttu-id="60a8c-p118">对于读取操作，响应中的空属性值（即两个引号之间没有空格 `''`）指示该单元格不包含任何数据或值。 在下面第一个示例中，区域中的第一个和最后一个单元格不包含任何数据。 在第二个示例中，区域中的前两个单元格不包含公式。</span><span class="sxs-lookup"><span data-stu-id="60a8c-p118">For read operations, a blank property value in the response (i.e., two quotation marks with no space in-between `''`) indicates that cell contains no data or value. In the first example below, the first and last cell in the range contain no data. In the second example, the first two cells in the range do not contain a formula.</span></span>

```js
range.values = [['', 'some', 'data', 'in', 'other', 'cells', '']];
```

```js
range.formula = [['', '', '=Rand()']];
```

## <a name="requirement-sets"></a><span data-ttu-id="60a8c-188">要求集</span><span class="sxs-lookup"><span data-stu-id="60a8c-188">Requirement sets</span></span>

<span data-ttu-id="60a8c-189">要求集是指各组已命名的 API 成员。</span><span class="sxs-lookup"><span data-stu-id="60a8c-189">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="60a8c-190">Office 加载项可以执行运行时检查或使用清单中指定的要求集确定 Office 应用程序是否支持加载项所需的 API。</span><span class="sxs-lookup"><span data-stu-id="60a8c-190">An Office Add-in can perform a runtime check or use requirement sets specified in the manifest to determine whether an Office application supports the APIs that the add-in needs.</span></span> <span data-ttu-id="60a8c-191">要确定每个受支持平台上可用的具体要求集，请参阅 [Excel JavaScript API 要求集](../reference/requirement-sets/excel-api-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="60a8c-191">To identify the specific requirement sets that are available on each supported platform, see [Excel JavaScript API requirement sets](../reference/requirement-sets/excel-api-requirement-sets.md).</span></span>

### <a name="checking-for-requirement-set-support-at-runtime"></a><span data-ttu-id="60a8c-192">在运行时检查要求集支持</span><span class="sxs-lookup"><span data-stu-id="60a8c-192">Checking for requirement set support at runtime</span></span>

<span data-ttu-id="60a8c-193">以下代码示例显示如何确定运行加载项的 Office 应用程序是否支持指定的 API 要求集。</span><span class="sxs-lookup"><span data-stu-id="60a8c-193">The following code sample shows how to determine whether the Office application where the add-in is running supports the specified API requirement set.</span></span>

```js
if (Office.context.requirements.isSetSupported('ExcelApi', '1.3')) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

### <a name="defining-requirement-set-support-in-the-manifest"></a><span data-ttu-id="60a8c-194">在清单中定义要求集支持</span><span class="sxs-lookup"><span data-stu-id="60a8c-194">Defining requirement set support in the manifest</span></span>

<span data-ttu-id="60a8c-195">可以在加载项清单中使用[要求元素](../reference/manifest/requirements.md)指定加载项要求激活的最小要求集和/或 API 方法。</span><span class="sxs-lookup"><span data-stu-id="60a8c-195">You can use the [Requirements element](../reference/manifest/requirements.md) in the add-in manifest to specify the minimal requirement sets and/or API methods that your add-in requires to activate.</span></span> <span data-ttu-id="60a8c-196">如果 Office 应用程序或平台不支持清单的 `Requirements` 元素中指定的要求集或 API 方法，该加载项不会在该应用程序或平台中运行，而且不会显示在“**我的加载项**”中显示的加载项列表中。</span><span class="sxs-lookup"><span data-stu-id="60a8c-196">If the Office application or platform doesn't support the requirement sets or API methods that are specified in the `Requirements` element of the manifest, the add-in won't run in that application or platform, and it won't display in the list of add-ins that are shown in **My Add-ins**.</span></span>

<span data-ttu-id="60a8c-197">以下代码示例显示加载项清单中的 `Requirements` 元素，该元素指定应在支持 ExcelApi 要求集版本 1.3 或更高版本的所有 Office 客户端应用程序中加载该加载项。</span><span class="sxs-lookup"><span data-stu-id="60a8c-197">The following code sample shows the `Requirements` element in an add-in manifest which specifies that the add-in should load in all Office client applications that support ExcelApi requirement set version 1.3 or greater.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.3">
      <Set Name="ExcelApi" MinVersion="1.3"/>
   </Sets>
</Requirements>
```

> [!NOTE]
> <span data-ttu-id="60a8c-198">为了让加载项适用于 Office 应用程序的所有平台（如 Excel 网页版、Windows 版 Excel 和 iPad 版 Excel），建议在运行时检查是否有要求支持，而不是在清单中定义要求集支持。</span><span class="sxs-lookup"><span data-stu-id="60a8c-198">To make your add-in available on all platforms of an Office application, such as Excel on the web, Windows, and iPad, we recommend that you check for requirement support at runtime instead of defining requirement set support in the manifest.</span></span>

### <a name="requirement-sets-for-the-officejs-common-api"></a><span data-ttu-id="60a8c-199">Office.js 通用 API 的要求集</span><span class="sxs-lookup"><span data-stu-id="60a8c-199">Requirement sets for the Office.js Common API</span></span>

<span data-ttu-id="60a8c-200">若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](../reference/requirement-sets/office-add-in-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="60a8c-200">For information about Common API requirement sets, see [Office Common API requirement sets](../reference/requirement-sets/office-add-in-requirement-sets.md).</span></span>

## <a name="handle-errors"></a><span data-ttu-id="60a8c-201">处理错误</span><span class="sxs-lookup"><span data-stu-id="60a8c-201">Handle errors</span></span>

<span data-ttu-id="60a8c-202">当 API 错误出现时，API 返回包含代码和消息的 `error` 对象。</span><span class="sxs-lookup"><span data-stu-id="60a8c-202">When an API error occurs, the API returns an `error` object that contains a code and a message.</span></span> <span data-ttu-id="60a8c-203">若要详细了解错误处理（包括 API 错误列表），请参阅[错误处理](excel-add-ins-error-handling.md)。</span><span class="sxs-lookup"><span data-stu-id="60a8c-203">For detailed information about error handling, including a list of API errors, see [Error handling](excel-add-ins-error-handling.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="60a8c-204">另请参阅</span><span class="sxs-lookup"><span data-stu-id="60a8c-204">See also</span></span>

* [<span data-ttu-id="60a8c-205">生成首个 Excel 加载项</span><span class="sxs-lookup"><span data-stu-id="60a8c-205">Build your first Excel add-in</span></span>](../quickstarts/excel-quickstart-jquery.md)
* [<span data-ttu-id="60a8c-206">Excel 加载项代码示例</span><span class="sxs-lookup"><span data-stu-id="60a8c-206">Excel add-ins code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="60a8c-207">Excel JavaScript API 性能优化</span><span class="sxs-lookup"><span data-stu-id="60a8c-207">Excel JavaScript API performance optimization</span></span>](../excel/performance.md)
* [<span data-ttu-id="60a8c-208">Excel JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="60a8c-208">Excel JavaScript API reference</span></span>](../reference/overview/excel-add-ins-reference-overview.md)
