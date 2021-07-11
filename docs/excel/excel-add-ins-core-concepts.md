---
title: Excel 加载项中的 Excel JavaScript 对象模型
description: 了解 Excel JavaScript API 中的关键对象类型，以及如何使用它们为 Excel 构建加载项。
ms.date: 04/05/2021
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 6c88dc84796d9fd898bee880035ed964ab6cd7c8
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349558"
---
# <a name="excel-javascript-object-model-in-office-add-ins"></a><span data-ttu-id="bce5a-103">Excel 加载项中的 Excel JavaScript 对象模型</span><span class="sxs-lookup"><span data-stu-id="bce5a-103">Excel JavaScript object model in Office Add-ins</span></span>

<span data-ttu-id="bce5a-104">本文介绍了如何使用 [Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) 生成 Excel 2016 或更高版本的加载项。</span><span class="sxs-lookup"><span data-stu-id="bce5a-104">This article describes how to use the [Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) to build add-ins for Excel 2016 or later.</span></span> <span data-ttu-id="bce5a-105">它引入了一些核心概念，这些概念是使用 API 的基础，并为执行特定任务提供指导，如读取或写入较大区域、更新区域内的所有单元格等等。</span><span class="sxs-lookup"><span data-stu-id="bce5a-105">It introduces core concepts that are fundamental to using the API and provides guidance for performing specific tasks such as reading or writing to a large range, updating all cells in range, and more.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="bce5a-106">请参阅[使用特定于应用程序的 API 模型](../develop/application-specific-api-model.md)，以了解 Excel API 的异步性质以及它们如何与工作簿协同工作。</span><span class="sxs-lookup"><span data-stu-id="bce5a-106">See [Using the application-specific API model](../develop/application-specific-api-model.md) to learn about the asynchronous nature of the Excel APIs and how they work with the workbook.</span></span>  

## <a name="officejs-apis-for-excel"></a><span data-ttu-id="bce5a-107">适用于 Excel 的 Office.js API</span><span class="sxs-lookup"><span data-stu-id="bce5a-107">Office.js APIs for Excel</span></span>

<span data-ttu-id="bce5a-108">Excel 加载项通过使用适 Office JavaScript API 与 Excel 中的对象进行交互，JavaScript API包括两个 JavaScript 对象模型：</span><span class="sxs-lookup"><span data-stu-id="bce5a-108">An Excel add-in interacts with objects in Excel by using the Office JavaScript API, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="bce5a-109">**Excel JavaScript API**：[Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) 随 Office 2016 一起引入，提供了强类型的对象，可用于访问工作表、区域、表格、图表等。</span><span class="sxs-lookup"><span data-stu-id="bce5a-109">**Excel JavaScript API**: Introduced with Office 2016, the [Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) provides strongly-typed objects that you can use to access worksheets, ranges, tables, charts, and more.</span></span>

* <span data-ttu-id="bce5a-110">**通用 API**：[通用 API](/javascript/api/office) 随 Office 2013 引入，可用于访问多种类型的 Office 应用程序中常见的 UI、对话框和客户端设置等功能。</span><span class="sxs-lookup"><span data-stu-id="bce5a-110">**Common APIs**: Introduced with Office 2013, the [Common API](/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications.</span></span>

<span data-ttu-id="bce5a-p102">你可能会使用 Excel JavaScript API 开发面向 Excel 2016 或更高版本的加载项中的大部分功能，同时还可以使用通用 API 中的对象。例如：</span><span class="sxs-lookup"><span data-stu-id="bce5a-p102">While you'll likely use the Excel JavaScript API to develop the majority of functionality in add-ins that target Excel 2016 or later, you'll also use objects in the Common API. For example:</span></span>

* <span data-ttu-id="bce5a-p103">[Context](/javascript/api/office/office.context)：`Context` 对象表示加载项的运行时环境，并提供对 API 关键对象的访问权限。 它由工作簿配置详细信息（如 `contentLanguage` 和 `officeTheme`）组成，并提供有关加载项的运行时环境（如 `host` 和 `platform`）的信息。 此外，它还提供了 `requirements.isSetSupported()` 方法，可用于检查运行加载项的 Excel 应用程序是否支持指定的要求集。</span><span class="sxs-lookup"><span data-stu-id="bce5a-p103">[Context](/javascript/api/office/office.context): The `Context` object represents the runtime environment of the add-in and provides access to key objects of the API. It consists of workbook configuration details such as `contentLanguage` and `officeTheme` and also provides information about the add-in's runtime environment such as `host` and `platform`. Additionally, it provides the `requirements.isSetSupported()` method, which you can use to check whether the specified requirement set is supported by the Excel application where the add-in is running.</span></span>
* <span data-ttu-id="bce5a-116">[Document](/javascript/api/office/office.document)：`Document` 对象提供 `getFileAsync()` 方法，用于下载运行加载项的 Excel 文件。</span><span class="sxs-lookup"><span data-stu-id="bce5a-116">[Document](/javascript/api/office/office.document): The `Document` object provides the `getFileAsync()` method, which you can use to download the Excel file where the add-in is running.</span></span>

<span data-ttu-id="bce5a-117">下图说明了可能使用 Excel JavaScript API 或公共 API 的情况。</span><span class="sxs-lookup"><span data-stu-id="bce5a-117">The following image illustrates when you might use the Excel JavaScript API or the Common APIs.</span></span>

![Excel JS API 和公共 API 之间的差异。](../images/excel-js-api-common-api.png)

## <a name="excel-specific-object-model"></a><span data-ttu-id="bce5a-119">特定于 Excel 的对象模型</span><span class="sxs-lookup"><span data-stu-id="bce5a-119">Excel-specific object model</span></span>

<span data-ttu-id="bce5a-120">若要了解 Excel API，则必须了解工作簿的各个组件之间如何相互关联。</span><span class="sxs-lookup"><span data-stu-id="bce5a-120">To understand the Excel APIs, you must understand how the components of a workbook are related to one another.</span></span>

* <span data-ttu-id="bce5a-121">一个 **Workbook** 包含一个或多个 **Worksheet**。</span><span class="sxs-lookup"><span data-stu-id="bce5a-121">A **Workbook** contains one or more **Worksheets**.</span></span>
* <span data-ttu-id="bce5a-122">**Worksheet** 包含出现在单个工作表中的那些数据对象的集合，并通过 **Range** 对象访问单元格。</span><span class="sxs-lookup"><span data-stu-id="bce5a-122">A **Worksheet** contains collections of those data objects that are present in the individual sheet, and gives access to cells through **Range** objects.</span></span>
* <span data-ttu-id="bce5a-123">**Range** 代表一组连续的单元格。</span><span class="sxs-lookup"><span data-stu-id="bce5a-123">A **Range** represents a group of contiguous cells.</span></span>
* <span data-ttu-id="bce5a-124">**Range** 用于创建和放置 **Table**、**Chart** 和 **Shape** 以及其他数据可视化或组织对象。</span><span class="sxs-lookup"><span data-stu-id="bce5a-124">**Ranges** are used to create and place **Tables**, **Charts**, **Shapes**, and other data visualization or organization objects.</span></span>
* <span data-ttu-id="bce5a-125">**Workbook** 包含整个 **Workbook** 的某些数据对象（例如，**Table**）的集合。</span><span class="sxs-lookup"><span data-stu-id="bce5a-125">**Workbooks** contain collections of some of those data objects (such as **Tables**) for the entire **Workbook**.</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

### <a name="ranges"></a><span data-ttu-id="bce5a-126">Range</span><span class="sxs-lookup"><span data-stu-id="bce5a-126">Ranges</span></span>

<span data-ttu-id="bce5a-127">Range 是工作簿中的一组连续单元格。</span><span class="sxs-lookup"><span data-stu-id="bce5a-127">A range is a group of contiguous cells in the workbook.</span></span> <span data-ttu-id="bce5a-128">加载项通常使用 A1 样式表示法（例如，对于 **B** 列和第 **3** 行中单个单元格，即 **B3** 或从 **C** 列至 **F** 列和第 **2** 行至第 **4** 行的单元格，即 **C2:F4**）来定义范围。</span><span class="sxs-lookup"><span data-stu-id="bce5a-128">Add-ins typically use A1-style notation (e.g. **B3** for the single cell in column **B** and row **3** or **C2:F4** for the cells from columns **C** through **F** and rows **2** through **4**) to define ranges.</span></span>

<span data-ttu-id="bce5a-129">Range 具有三个核心属性：`values`、`formulas` 和 `format`。</span><span class="sxs-lookup"><span data-stu-id="bce5a-129">Ranges have three core properties: `values`, `formulas`, and `format`.</span></span> <span data-ttu-id="bce5a-130">这些属性获取或设置单元格值、要计算的公式以及单元格的视觉对象格式设置。</span><span class="sxs-lookup"><span data-stu-id="bce5a-130">These properties get or set the cell values, formulas to be evaluated, and the visual formatting of the cells.</span></span>

#### <a name="range-sample"></a><span data-ttu-id="bce5a-131">Range 示例</span><span class="sxs-lookup"><span data-stu-id="bce5a-131">Range sample</span></span>

<span data-ttu-id="bce5a-132">以下示例显示了如何创建销售记录。</span><span class="sxs-lookup"><span data-stu-id="bce5a-132">The following sample shows how to create sales records.</span></span> <span data-ttu-id="bce5a-133">此函数使用 `Range` 对象来设置值、公式和格式。</span><span class="sxs-lookup"><span data-stu-id="bce5a-133">This function uses `Range` objects to set the values, formulas, and formats.</span></span>

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

<span data-ttu-id="bce5a-134">此示例将在当前工作表中创建以下数据。</span><span class="sxs-lookup"><span data-stu-id="bce5a-134">This sample creates the following data in the current worksheet.</span></span>

![显示值行、公式列和格式化标题的销售记录。](../images/excel-overview-range-sample.png)

<span data-ttu-id="bce5a-136">有关详细信息，请参阅[使用 Excel JavaScript API 设置和获取范围值、文本或公式](excel-add-ins-ranges-set-get-values.md)。</span><span class="sxs-lookup"><span data-stu-id="bce5a-136">For more information, see [Set and get range values, text, or formulas using the Excel JavaScript API](excel-add-ins-ranges-set-get-values.md).</span></span>

### <a name="charts-tables-and-other-data-objects"></a><span data-ttu-id="bce5a-137">Chart、Table 和其他数据对象</span><span class="sxs-lookup"><span data-stu-id="bce5a-137">Charts, tables, and other data objects</span></span>

<span data-ttu-id="bce5a-138">Excel JavaScript API 可以在 Excel 中创建和设置数据结构和可视化效果。</span><span class="sxs-lookup"><span data-stu-id="bce5a-138">The Excel JavaScript APIs can create and manipulate the data structures and visualizations within Excel.</span></span> <span data-ttu-id="bce5a-139">Table 和 Chart 是最常用的两个对象，但是 API 支持数据透视表、形状和图像等。</span><span class="sxs-lookup"><span data-stu-id="bce5a-139">Tables and charts are two of the more commonly used objects, but the APIs support PivotTables, shapes, images, and more.</span></span>

#### <a name="creating-a-table"></a><span data-ttu-id="bce5a-140">创建表</span><span class="sxs-lookup"><span data-stu-id="bce5a-140">Creating a table</span></span>

<span data-ttu-id="bce5a-p108">通过使用数据填充区域创建表。自动将格式设置和表格控件（如筛选器）应用到区域。</span><span class="sxs-lookup"><span data-stu-id="bce5a-p108">Create tables by using data-filled ranges. Formatting and table controls (such as filters) are automatically applied to the range.</span></span>

<span data-ttu-id="bce5a-143">以下示例使用上一个示例中的范围创建了一个表。</span><span class="sxs-lookup"><span data-stu-id="bce5a-143">The following sample creates a table using the ranges from the previous sample.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.tables.add("B2:E5", true);
    return context.sync();
});
```

<span data-ttu-id="bce5a-144">在包含之前数据的工作表上使用此示例代码将创建下表。</span><span class="sxs-lookup"><span data-stu-id="bce5a-144">Using this sample code on the worksheet with the previous data creates the following table.</span></span>

![使用之前的销售记录制成的表。](../images/excel-overview-table-sample.png)

<span data-ttu-id="bce5a-146">有关详细信息，请参阅[使用 Excel JavaScript API 处理表格](excel-add-ins-tables.md)。</span><span class="sxs-lookup"><span data-stu-id="bce5a-146">For more information, see [Work with tables using the Excel JavaScript API](excel-add-ins-tables.md).</span></span>

#### <a name="creating-a-chart"></a><span data-ttu-id="bce5a-147">创建图表</span><span class="sxs-lookup"><span data-stu-id="bce5a-147">Creating a chart</span></span>

<span data-ttu-id="bce5a-148">创建图表以直观显示某个范围内的数据。</span><span class="sxs-lookup"><span data-stu-id="bce5a-148">Create charts to visualize the data in a range.</span></span> <span data-ttu-id="bce5a-149">该 API 支持数十种图表类型，每种都可以根据需要进行自定义。</span><span class="sxs-lookup"><span data-stu-id="bce5a-149">The APIs support dozens of chart varieties, each of which can be customized to suit your needs.</span></span>

<span data-ttu-id="bce5a-150">下面的示例为三个项目创建一个简单的柱形图，并将其置于工作表顶部下方 100 像素处。</span><span class="sxs-lookup"><span data-stu-id="bce5a-150">The following sample creates a simple column chart for three items and places it 100 pixels below the top of the worksheet.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var chart = sheet.charts.add(Excel.ChartType.columnStacked, sheet.getRange("B3:C5"));
    chart.top = 100;
    return context.sync();
});
```

<span data-ttu-id="bce5a-151">在工作表上使用上一个表运行此示例将创建以下图表。</span><span class="sxs-lookup"><span data-stu-id="bce5a-151">Running this sample on the worksheet with the previous table creates the following chart.</span></span>

![一个柱形图，显示上一个销售记录中三个项目的数量。](../images/excel-overview-chart-sample.png)

<span data-ttu-id="bce5a-153">有关详细信息，请参阅[使用 Excel JavaScript API 处理图表](excel-add-ins-charts.md)。</span><span class="sxs-lookup"><span data-stu-id="bce5a-153">For more information, see [Work with charts using the Excel JavaScript API](excel-add-ins-charts.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="bce5a-154">另请参阅</span><span class="sxs-lookup"><span data-stu-id="bce5a-154">See also</span></span>

* [<span data-ttu-id="bce5a-155">生成首个 Excel 加载项</span><span class="sxs-lookup"><span data-stu-id="bce5a-155">Build your first Excel add-in</span></span>](../quickstarts/excel-quickstart-jquery.md)
* [<span data-ttu-id="bce5a-156">Excel 加载项代码示例</span><span class="sxs-lookup"><span data-stu-id="bce5a-156">Excel add-ins code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="bce5a-157">Excel JavaScript API 性能优化</span><span class="sxs-lookup"><span data-stu-id="bce5a-157">Excel JavaScript API performance optimization</span></span>](../excel/performance.md)
* [<span data-ttu-id="bce5a-158">Excel JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="bce5a-158">Excel JavaScript API reference</span></span>](../reference/overview/excel-add-ins-reference-overview.md)
