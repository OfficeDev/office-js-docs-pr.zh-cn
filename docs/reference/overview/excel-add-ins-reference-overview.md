---
title: Excel JavaScript API 概述
description: ''
ms.date: 03/19/2019
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: bf1d4642a7ceeb34eab51722a398887bb5c03fec
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450168"
---
# <a name="excel-javascript-api-overview"></a><span data-ttu-id="19057-102">Excel JavaScript API 概述</span><span class="sxs-lookup"><span data-stu-id="19057-102">Excel JavaScript API overview</span></span>

<span data-ttu-id="19057-103">可以使用 Excel JavaScript API 构建适用于 Excel 2016 或更高版本的加载项。</span><span class="sxs-lookup"><span data-stu-id="19057-103">You can use the Excel JavaScript API to build add-ins for Excel 2016 or later.</span></span> <span data-ttu-id="19057-104">以下列表显示在 API 中可用的高级 Excel 对象。</span><span class="sxs-lookup"><span data-stu-id="19057-104">The following list shows the high-level Excel objects that are available in the API.</span></span> <span data-ttu-id="19057-105">每个对象页面链接包含对象可用的属性、事件和方法的描述。</span><span class="sxs-lookup"><span data-stu-id="19057-105">Each object page link contains a description of the properties, events, and methods that are available on the object.</span></span> <span data-ttu-id="19057-106">如需了解详细信息，请从菜单中浏览相应链接。</span><span class="sxs-lookup"><span data-stu-id="19057-106">Explore the links from the menu to learn more.</span></span>

<span data-ttu-id="19057-107">为了方便起见，下面列出了一些核心 Excel 对象：</span><span class="sxs-lookup"><span data-stu-id="19057-107">Some of the core Excel objects are listed below for convenience:</span></span> 

- <span data-ttu-id="19057-108">[工作簿](/javascript/api/excel/excel.workbook)：包含相关 workbook 对象的顶级对象，例如 worksheet、table、range 等。它还可以用于列出相关的参考。</span><span class="sxs-lookup"><span data-stu-id="19057-108">[Workbook](/javascript/api/excel/excel.workbook): The top-level object that contains related workbook objects such as worksheets, tables, ranges, etc. It also can be used to list related references.</span></span>

- <span data-ttu-id="19057-109">[Worksheet](/javascript/api/excel/excel.worksheet)：表示工作簿中的工作表。</span><span class="sxs-lookup"><span data-stu-id="19057-109">[Worksheet](/javascript/api/excel/excel.worksheet): Represents a worksheet in a workbook.</span></span> 
    - <span data-ttu-id="19057-110">[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)：工作簿中 **Worksheet** 对象的集合。</span><span class="sxs-lookup"><span data-stu-id="19057-110">[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection): A collection of the **Worksheet** objects in a workbook.</span></span>
    - <span data-ttu-id="19057-111">[WorksheetProtection](/javascript/api/excel/excel.worksheetprotection)：表示对 **Worksheet** 对象的保护。</span><span class="sxs-lookup"><span data-stu-id="19057-111">[WorksheetProtection](/javascript/api/excel/excel.worksheetprotection): Represents the protection of a **Worksheet** object.</span></span>

- <span data-ttu-id="19057-112">[Range](/javascript/api/excel/excel.range)：表示某一单元格、某一行、某一列、某一单元格选定区域（其中包含一个或多个相邻单元格块）。</span><span class="sxs-lookup"><span data-stu-id="19057-112">[Range](/javascript/api/excel/excel.range): Represents a cell, a row, a column, or a selection of cells containing one or more contiguous blocks of cells.</span></span>
    - <span data-ttu-id="19057-113">[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)：定义满足规则条件时应用到该区域的规则和格式的对象。</span><span class="sxs-lookup"><span data-stu-id="19057-113">[ConditionalFormat](/javascript/api/excel/excel.conditionalformat): An object defining a rule and a format applied to the range when the rule's condition is met.</span></span>
    - <span data-ttu-id="19057-114">[DataValidation](/javascript/api/excel/excel.datavalidation)：根据各种条件将用户输入限制在某个区域内的对象。</span><span class="sxs-lookup"><span data-stu-id="19057-114">[DataValidation](/javascript/api/excel/excel.datavalidation): An object that restricts user input to a range based on a variety of criteria.</span></span>
    - <span data-ttu-id="19057-115">[RangeSort](/javascript/api/excel/excel.rangesort)：表示管理区域中排序操作的对象。</span><span class="sxs-lookup"><span data-stu-id="19057-115">[RangeSort](/javascript/api/excel/excel.rangesort): Represents a object that manages sorting operations on a range.</span></span>

- <span data-ttu-id="19057-116">[Table](/javascript/api/excel/excel.table)：表示有组织的单元格的集合，设计用于简化数据管理。</span><span class="sxs-lookup"><span data-stu-id="19057-116">[Table](/javascript/api/excel/excel.table): Represents a collection of organized cells designed to make management of the data easy.</span></span>
    - <span data-ttu-id="19057-117">[TableCollection](/javascript/api/excel/excel.tablecollection)：工作簿或工作表中的表的集合。</span><span class="sxs-lookup"><span data-stu-id="19057-117">[TableCollection](/javascript/api/excel/excel.tablecollection): A collection of tables in a workbook or worksheet.</span></span>
    - <span data-ttu-id="19057-118">[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)：表中所有列的集合。</span><span class="sxs-lookup"><span data-stu-id="19057-118">[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection): A collection of all the columns in a table.</span></span>
    - <span data-ttu-id="19057-119">[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)：表中所有行的集合。</span><span class="sxs-lookup"><span data-stu-id="19057-119">[TableRowCollection](/javascript/api/excel/excel.tablerowcollection): A collection of all the rows in a table.</span></span>
    - <span data-ttu-id="19057-120">[TableSort](/javascript/api/excel/excel.tablesort)：表示管理区域中排序操作的对象。</span><span class="sxs-lookup"><span data-stu-id="19057-120">[TableSort](/javascript/api/excel/excel.tablesort): Represents an object that manages sorting operations on a table.</span></span>

- <span data-ttu-id="19057-121">[Chart](/javascript/api/excel/excel.chart)：表示工作表中的 chart 对象，它是基础数据的可视表示形式。</span><span class="sxs-lookup"><span data-stu-id="19057-121">[Chart](/javascript/api/excel/excel.chart): Represents a chart object in a worksheet, which is a visual representation of underlying data.</span></span>
    - <span data-ttu-id="19057-122">[ChartCollection](/javascript/api/excel/excel.chartcollection)：工作表中的图表的集合。</span><span class="sxs-lookup"><span data-stu-id="19057-122">[ChartCollection](/javascript/api/excel/excel.chartcollection): A collection of charts in a worksheet.</span></span>
    
- <span data-ttu-id="19057-123">[PivotTable](/javascript/api/excel/excel.pivottable)：表示 Excel 数据透视表，它是数据的分层分组表示。</span><span class="sxs-lookup"><span data-stu-id="19057-123">[PivotTable](/javascript/api/excel/excel.pivottable): Represents an Excel PivotTable, which is a hierarchical grouping and presentation of data.</span></span> 
    - <span data-ttu-id="19057-124">[TableCollection](/javascript/api/excel/excel.pivottablecollection)：工作表中的数据透视表的集合。</span><span class="sxs-lookup"><span data-stu-id="19057-124">[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection): A collection of PivotTables in a worksheet.</span></span>

- <span data-ttu-id="19057-125">[Filter](/javascript/api/excel/excel.filter)：表示管理表格列筛选的对象。</span><span class="sxs-lookup"><span data-stu-id="19057-125">[Filter](/javascript/api/excel/excel.filter): Represents an object that manages the filtering of a table's column.</span></span>

- <span data-ttu-id="19057-126">[NamedItem](/javascript/api/excel/excel.nameditem)：表示单元格区域或值的定义名称。</span><span class="sxs-lookup"><span data-stu-id="19057-126">[NamedItem](/javascript/api/excel/excel.nameditem): Represents a defined name for a range of cells or a value.</span></span> 
    - <span data-ttu-id="19057-127">[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection)：工作簿中 **NamedItem** 对象的集合。</span><span class="sxs-lookup"><span data-stu-id="19057-127">[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection): A collection of the **NamedItem** objects in a workbook.</span></span>

- <span data-ttu-id="19057-128">[Binding](/javascript/api/excel/excel.binding)：表示对工作簿的某一部分的绑定的抽象类。</span><span class="sxs-lookup"><span data-stu-id="19057-128">[Binding](/javascript/api/excel/excel.binding): An abstract class that represents a binding to a section of the workbook.</span></span>
    - <span data-ttu-id="19057-129">[BindingCollection](/javascript/api/excel/excel.bindingcollection)：工作簿中 **Binding** 对象的集合。</span><span class="sxs-lookup"><span data-stu-id="19057-129">[BindingCollection](/javascript/api/excel/excel.bindingcollection): A collection of the **Binding** objects in a workbook.</span></span>

## <a name="excel-javascript-api-open-specifications"></a><span data-ttu-id="19057-130">Excel JavaScript API 开放性规范</span><span class="sxs-lookup"><span data-stu-id="19057-130">Excel JavaScript API open specifications</span></span>

<span data-ttu-id="19057-131">在我们设计和开发用于 Excel 加载项的新 API 时，我们将使它们可在[开放 API 规范](../openspec.md)页面上接收反馈。</span><span class="sxs-lookup"><span data-stu-id="19057-131">As we design and develop new APIs for Excel add-ins, we'll make them available for your feedback on our [Open API specifications](../openspec.md) page.</span></span> <span data-ttu-id="19057-132">了解即将推出的面向 Excel JavaScript API 的新功能，并提供对我们的设计规范的宝贵意见。</span><span class="sxs-lookup"><span data-stu-id="19057-132">Find out what new features are in the pipeline for the Excel JavaScript APIs, and provide your input on our design specifications.</span></span>

## <a name="excel-javascript-api-requirement-sets"></a><span data-ttu-id="19057-133">Excel JavaScript API 要求集</span><span class="sxs-lookup"><span data-stu-id="19057-133">Excel JavaScript API requirement sets</span></span>

<span data-ttu-id="19057-134">要求集是指各组已命名的 API 成员。</span><span class="sxs-lookup"><span data-stu-id="19057-134">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="19057-135">Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 主机是否支持加载项所需的 API。</span><span class="sxs-lookup"><span data-stu-id="19057-135">Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs.</span></span> <span data-ttu-id="19057-136">有关 Excel JavaScript API 要求集的详细信息，请参阅 [Excel JavaScript API 要求集](../requirement-sets/excel-api-requirement-sets.md)文章。</span><span class="sxs-lookup"><span data-stu-id="19057-136">For detailed information about Excel JavaScript API requirement sets, see the [Excel JavaScript API requirement sets](../requirement-sets/excel-api-requirement-sets.md) article.</span></span>

## <a name="excel-javascript-api-reference"></a><span data-ttu-id="19057-137">Excel JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="19057-137">Excel JavaScript API reference</span></span>

<span data-ttu-id="19057-138">有关 Excel JavaScript API 的详细信息，请参阅 [Excel JavaScript API 参考文档](/javascript/api/excel)。</span><span class="sxs-lookup"><span data-stu-id="19057-138">For detailed information about the Excel JavaScript API, see the [Excel JavaScript API reference documentation](/javascript/api/excel).</span></span>

## <a name="see-also"></a><span data-ttu-id="19057-139">另请参阅</span><span class="sxs-lookup"><span data-stu-id="19057-139">See also</span></span>

- [<span data-ttu-id="19057-140">Excel 加载项概述</span><span class="sxs-lookup"><span data-stu-id="19057-140">Excel add-ins overview</span></span>](/office/dev/add-ins/excel/excel-add-ins-overview)
- [<span data-ttu-id="19057-141">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="19057-141">Office Add-ins platform overview</span></span>](/office/dev/add-ins/overview/office-add-ins)
- [<span data-ttu-id="19057-142">GitHub 上的 Excel 加载项示例</span><span class="sxs-lookup"><span data-stu-id="19057-142">Excel add-in samples on GitHub</span></span>](https://github.com/OfficeDev?utf8=%E2%9C%93&q=Excel)
