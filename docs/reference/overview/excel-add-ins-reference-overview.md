---
title: Excel JavaScript API 概述
description: ''
ms.date: 06/10/2019
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: aa9574a93252c0011b211c39e37cc013beb64432
ms.sourcegitcommit: 3f84b2caa73d7fe1eb0d15e32ea4dec459e2ff53
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/12/2019
ms.locfileid: "34910145"
---
# <a name="excel-javascript-api-overview"></a>Excel JavaScript API 概述

可以使用 Excel JavaScript API 构建适用于 Excel 2016 或更高版本的加载项。 以下列表显示在 API 中可用的高级 Excel 对象。 每个对象页面链接包含对象可用的属性、事件和方法的描述。 如需了解详细信息，请从菜单中浏览相应链接。

为了方便起见，下面列出了一些核心 Excel 对象：

- [工作簿](/javascript/api/excel/excel.workbook)：包含相关 workbook 对象的顶级对象，例如 worksheet、table、range 等。它还可以用于列出相关的参考。

- [Worksheet](/javascript/api/excel/excel.worksheet)：表示工作簿中的工作表。
  - [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)：工作簿中 **Worksheet** 对象的集合。
  - [WorksheetProtection](/javascript/api/excel/excel.worksheetprotection)：表示对 **Worksheet** 对象的保护。

- [Range](/javascript/api/excel/excel.range)：表示某一单元格、某一行、某一列、某一单元格选定区域（其中包含一个或多个相邻单元格块）。
  - [ConditionalFormat](/javascript/api/excel/excel.conditionalformat)：定义满足规则条件时应用到该区域的规则和格式的对象。
  - [DataValidation](/javascript/api/excel/excel.datavalidation)：根据各种条件将用户输入限制在某个区域内的对象。
  - [RangeSort](/javascript/api/excel/excel.rangesort)：表示管理区域中排序操作的对象。

- [Table](/javascript/api/excel/excel.table)：表示有组织的单元格的集合，设计用于简化数据管理。
  - [TableCollection](/javascript/api/excel/excel.tablecollection)：工作簿或工作表中的表的集合。
  - [TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)：表中所有列的集合。
  - [TableRowCollection](/javascript/api/excel/excel.tablerowcollection)：表中所有行的集合。
  - [TableSort](/javascript/api/excel/excel.tablesort)：表示管理区域中排序操作的对象。

- [Chart](/javascript/api/excel/excel.chart)：表示工作表中的 chart 对象，它是基础数据的可视表示形式。
  - [ChartCollection](/javascript/api/excel/excel.chartcollection)：工作表中的图表的集合。

- [PivotTable](/javascript/api/excel/excel.pivottable)：表示 Excel 数据透视表，它是数据的分层分组表示。
  - [TableCollection](/javascript/api/excel/excel.pivottablecollection)：工作表中的数据透视表的集合。

- [Filter](/javascript/api/excel/excel.filter)：表示管理表格列筛选的对象。

- [NamedItem](/javascript/api/excel/excel.nameditem)：表示单元格区域或值的定义名称。
  - [NamedItemCollection](/javascript/api/excel/excel.nameditemcollection)：工作簿中 **NamedItem** 对象的集合。

- [Binding](/javascript/api/excel/excel.binding)：表示对工作簿的某一部分的绑定的抽象类。
  - [BindingCollection](/javascript/api/excel/excel.bindingcollection)：工作簿中 **Binding** 对象的集合。

## <a name="excel-javascript-api-requirement-sets"></a>Excel JavaScript API 要求集

要求集是指各组已命名的 API 成员。 Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 主机是否支持加载项所需的 API。 有关 Excel JavaScript API 要求集的详细信息，请参阅 [Excel JavaScript API 要求集](../requirement-sets/excel-api-requirement-sets.md)文章。

## <a name="excel-javascript-api-reference"></a>Excel JavaScript API 参考

有关 Excel JavaScript API 的详细信息，请参阅 [Excel JavaScript API 参考文档](/javascript/api/excel)。

## <a name="see-also"></a>另请参阅

- [Excel 加载项概述](/office/dev/add-ins/excel/excel-add-ins-overview)
- [Office 加载项平台概述](/office/dev/add-ins/overview/office-add-ins)
- [GitHub 上的 Excel 加载项示例](https://github.com/OfficeDev?utf8=%E2%9C%93&q=Excel)
- [API 开放性规范](../openspec/openspec.md)
