# <a name="excel-javascript-api-overview"></a><span data-ttu-id="92a68-101">Excel JavaScript API 概述</span><span class="sxs-lookup"><span data-stu-id="92a68-101">Excel JavaScript API programming overview</span></span>

<span data-ttu-id="92a68-102">你可以使用 Excel JavaScript API 构建适用于 Excel 2016 或后续版本的外接程序。</span><span class="sxs-lookup"><span data-stu-id="92a68-102">You can use the Excel JavaScript API to build add-ins for Excel 2016.</span></span> <span data-ttu-id="92a68-103">以下列表显示在 API 中可用的高级 Excel 对象。</span><span class="sxs-lookup"><span data-stu-id="92a68-103">The following list shows the high-level Excel objects that are available in the API.</span></span> <span data-ttu-id="92a68-104">每个对象页面链接包含对象可用的属性、关系和方法的描述。</span><span class="sxs-lookup"><span data-stu-id="92a68-104">Each object page link contains a description of the properties, relationships, and methods that are available on the object.</span></span> <span data-ttu-id="92a68-105">如需了解详细信息，请从菜单中浏览相应链接。</span><span class="sxs-lookup"><span data-stu-id="92a68-105">Explore the links from the menu to learn more.</span></span>

<span data-ttu-id="92a68-106">为了方便起见，下面列出了一些核心 Excel 对象：</span><span class="sxs-lookup"><span data-stu-id="92a68-106">Some of the core Excel objects are listed below for convenience:</span></span> 

- <span data-ttu-id="92a68-107">[工作簿](/javascript/api/excel/excel.workbook)：包含相关 workbook 对象的顶级对象，例如 worksheet、table、range 等。它还可以用于列出相关的参考。</span><span class="sxs-lookup"><span data-stu-id="92a68-107">[Workbook](/javascript/api/excel/excel.workbook): The top-level object that contains related workbook objects such as worksheets, tables, ranges, etc. It also can be used to list related references.</span></span>

- <span data-ttu-id="92a68-108">[Worksheet](/javascript/api/excel/excel.worksheet)：表示工作簿中的工作表。</span><span class="sxs-lookup"><span data-stu-id="92a68-108">[Worksheet](/javascript/api/excel/excel.worksheet): Represents a worksheet in a workbook.</span></span> 
    - <span data-ttu-id="92a68-109">[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)：工作簿中 **Worksheet** 对象的集合。</span><span class="sxs-lookup"><span data-stu-id="92a68-109">[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection): A collection of the **Worksheet** objects in a workbook.</span></span>

- <span data-ttu-id="92a68-110">[Range](/javascript/api/excel/excel.range)：表示某一单元格、某一行、某一列、某一单元格选定区域（其中包含一个或多个相邻单元格块）。</span><span class="sxs-lookup"><span data-stu-id="92a68-110">[Range](/javascript/api/excel/excel.range): Represents a cell, a row, a column, or a selection of cells containing one or more contiguous blocks of cells.</span></span>

- <span data-ttu-id="92a68-111">[Table](/javascript/api/excel/excel.table)：表示有组织的单元格集合，设计用于简化数据管理。</span><span class="sxs-lookup"><span data-stu-id="92a68-111">[Table](/javascript/api/excel/excel.table): Represents a collection of organized cells designed to make management of the data easy.</span></span>
    - <span data-ttu-id="92a68-112">[TableCollection](/javascript/api/excel/excel.tablecollection)：工作簿或工作表中的表集合。</span><span class="sxs-lookup"><span data-stu-id="92a68-112">[TableCollection](/javascript/api/excel/excel.tablecollection): A collection of tables in a workbook or worksheet.</span></span>
    - <span data-ttu-id="92a68-113">[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)：表中所有列的集合。</span><span class="sxs-lookup"><span data-stu-id="92a68-113">[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection): A collection of all the columns in a table.</span></span>
    - <span data-ttu-id="92a68-114">[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)：表中所有行的集合。</span><span class="sxs-lookup"><span data-stu-id="92a68-114">[TableRowCollection](/javascript/api/excel/excel.tablerowcollection): A collection of all the rows in a table.</span></span>

- <span data-ttu-id="92a68-115">[Chart](/javascript/api/excel/excel.chart)：表示工作表中的图表对象，它是基础数据的可视表示形式。</span><span class="sxs-lookup"><span data-stu-id="92a68-115">[Chart](/javascript/api/excel/excel.chart): Represents a chart object in a worksheet, which is a visual representation of underlying data.</span></span>
    - <span data-ttu-id="92a68-116">[ChartCollection](/javascript/api/excel/excel.chartcollection)：工作表中的图表的集合。</span><span class="sxs-lookup"><span data-stu-id="92a68-116">[ChartCollection](/javascript/api/excel/excel.chartcollection): A collection of charts in a worksheet.</span></span>

- <span data-ttu-id="92a68-117">[TableSort](/javascript/api/excel/excel.tablesort)：表示管理 **Table** 对象排序操作的对象。</span><span class="sxs-lookup"><span data-stu-id="92a68-117">[TableSort](/javascript/api/excel/excel.tablesort): Represents an object that manages sorting operations on **Table** objects.</span></span>

- <span data-ttu-id="92a68-118">[RangeSort](/javascript/api/excel/excel.rangesort)：表示管理 **Range** 对象排序操作的对象。</span><span class="sxs-lookup"><span data-stu-id="92a68-118">[RangeSort](/javascript/api/excel/excel.rangesort): Represents a object that manages sorting operations on **Range** objects.</span></span>

- <span data-ttu-id="92a68-119">[Filter](/javascript/api/excel/excel.filter)：表示管理表格列筛选的对象。</span><span class="sxs-lookup"><span data-stu-id="92a68-119">[Filter](/javascript/api/excel/excel.filter): Represents an object that manages the filtering of a table's column.</span></span>

- <span data-ttu-id="92a68-120">[Filter](/javascript/api/excel/excel.worksheetprotection): 表示管理表格列筛选的对象。\*\*\*\*</span><span class="sxs-lookup"><span data-stu-id="92a68-120">[WorksheetProtection](/javascript/api/excel/excel.worksheetprotection): Represents the protection of a **Worksheet** object.</span></span>

- <span data-ttu-id="92a68-121">[NamedItem](/javascript/api/excel/excel.nameditem)：表示单元格区域或值的定义名称。</span><span class="sxs-lookup"><span data-stu-id="92a68-121">[NamedItem](/javascript/api/excel/excel.nameditem): Represents a defined name for a range of cells or a value.</span></span> 
    - <span data-ttu-id="92a68-122">[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection)：工作簿中 **NamedItem** 对象的集合。</span><span class="sxs-lookup"><span data-stu-id="92a68-122">[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection): A collection of the **NamedItem** objects in a workbook.</span></span>

- <span data-ttu-id="92a68-123">[Binding](/javascript/api/excel/excel.binding)：表示对工作簿的某一部分的绑定抽象类。</span><span class="sxs-lookup"><span data-stu-id="92a68-123">[Binding](/javascript/api/excel/excel.binding): An abstract class that represents a binding to a section of the workbook.</span></span>
    - <span data-ttu-id="92a68-124">[BindingCollection](/javascript/api/excel/excel.bindingcollection)：工作簿中 **Binding** 对象的集合。</span><span class="sxs-lookup"><span data-stu-id="92a68-124">[BindingCollection](/javascript/api/excel/excel.bindingcollection): A collection of the **Binding** objects in a workbook.</span></span>

## <a name="excel-javascript-api-open-specifications"></a><span data-ttu-id="92a68-125">Excel 的 JavaScript API 开放性规范</span><span class="sxs-lookup"><span data-stu-id="92a68-125">Excel JavaScript API open specifications</span></span>

<span data-ttu-id="92a68-126">在设计和开发新的 Excel  外接应用程序 API 时，我们会提供“[开放性 API 规范](../openspec.md)”页面以便获取您的反馈。</span><span class="sxs-lookup"><span data-stu-id="92a68-126">As we design and develop new APIs, we'll make them available for your feedback on our [Open API specifications](../openspec.md) page.</span></span> <span data-ttu-id="92a68-127">了解管道中的新增功能，并提供你对我们设计规范的宝贵意见。</span><span class="sxs-lookup"><span data-stu-id="92a68-127">Find out what new features are in the pipeline, and provide your input on our design specifications.</span></span>

## <a name="excel-javascript-api-reference"></a><span data-ttu-id="92a68-128">Excel JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="92a68-128">Excel JavaScript API reference</span></span>

<span data-ttu-id="92a68-129">有关 Excel JavaScript API 的详细信息，请参阅 [Excel 的 JavaScript API 参考文档](/javascript/api/excel)。</span><span class="sxs-lookup"><span data-stu-id="92a68-129">For detailed information about Excel JavaScript API, see the [Excel JavaScript API reference documentation](/javascript/api/excel).</span></span>

## <a name="see-also"></a><span data-ttu-id="92a68-130">请参阅</span><span class="sxs-lookup"><span data-stu-id="92a68-130">See also</span></span>

- [<span data-ttu-id="92a68-131">Excel 外接程序概述</span><span class="sxs-lookup"><span data-stu-id="92a68-131">Excel add-ins overview</span></span>](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-overview)
- [<span data-ttu-id="92a68-132">Office 外接程序平台概述</span><span class="sxs-lookup"><span data-stu-id="92a68-132">Office Add-ins platform overview</span></span>](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
- [<span data-ttu-id="92a68-133">GitHub Excel 加载项示例</span><span class="sxs-lookup"><span data-stu-id="92a68-133">Excel add-in samples on GitHub</span></span>](https://github.com/OfficeDev?utf8=%E2%9C%93&q=Excel)
