# <a name="excel-javascript-api-overview"></a>Excel JavaScript API 概述

你可以使用 Excel JavaScript API 构建适用于 Excel 2016 或后续版本的外接程序。 以下列表显示在 API 中可用的高级 Excel 对象。 每个对象页面链接包含对象可用的属性、关系和方法的描述。 如需了解详细信息，请从菜单中浏览相应链接。

为了方便起见，下面列出了一些核心 Excel 对象： 

- [工作簿](/javascript/api/excel/excel.workbook)：包含相关 workbook 对象的顶级对象，例如 worksheet、table、range 等。它还可以用于列出相关的参考。

- [Worksheet](/javascript/api/excel/excel.worksheet)：表示工作簿中的工作表。 
    - [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)：工作簿中 **Worksheet** 对象的集合。

- [Range](/javascript/api/excel/excel.range)：表示某一单元格、某一行、某一列、某一单元格选定区域（其中包含一个或多个相邻单元格块）。

- [Table](/javascript/api/excel/excel.table)：表示有组织的单元格集合，设计用于简化数据管理。
    - [TableCollection](/javascript/api/excel/excel.tablecollection)：工作簿或工作表中的表集合。
    - [TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)：表中所有列的集合。
    - [TableRowCollection](/javascript/api/excel/excel.tablerowcollection)：表中所有行的集合。

- [Chart](/javascript/api/excel/excel.chart)：表示工作表中的图表对象，它是基础数据的可视表示形式。
    - [ChartCollection](/javascript/api/excel/excel.chartcollection)：工作表中的图表的集合。

- [TableSort](/javascript/api/excel/excel.tablesort)：表示管理 **Table** 对象排序操作的对象。

- [RangeSort](/javascript/api/excel/excel.rangesort)：表示管理 **Range** 对象排序操作的对象。

- [Filter](/javascript/api/excel/excel.filter)：表示管理表格列筛选的对象。

- [Filter](/javascript/api/excel/excel.worksheetprotection): 表示管理表格列筛选的对象。****

- [NamedItem](/javascript/api/excel/excel.nameditem)：表示单元格区域或值的定义名称。 
    - [NamedItemCollection](/javascript/api/excel/excel.nameditemcollection)：工作簿中 **NamedItem** 对象的集合。

- [Binding](/javascript/api/excel/excel.binding)：表示对工作簿的某一部分的绑定抽象类。
    - [BindingCollection](/javascript/api/excel/excel.bindingcollection)：工作簿中 **Binding** 对象的集合。

## <a name="excel-javascript-api-open-specifications"></a>Excel 的 JavaScript API 开放性规范

在设计和开发新的 Excel  外接应用程序 API 时，我们会提供“[开放性 API 规范](../openspec.md)”页面以便获取您的反馈。 了解管道中的新增功能，并提供你对我们设计规范的宝贵意见。

## <a name="excel-javascript-api-reference"></a>Excel JavaScript API 参考

有关 Excel JavaScript API 的详细信息，请参阅 [Excel 的 JavaScript API 参考文档](/javascript/api/excel)。

## <a name="see-also"></a>请参阅

- [Excel 外接程序概述](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-overview)
- [Office 外接程序平台概述](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
- [GitHub Excel 加载项示例](https://github.com/OfficeDev?utf8=%E2%9C%93&q=Excel)
