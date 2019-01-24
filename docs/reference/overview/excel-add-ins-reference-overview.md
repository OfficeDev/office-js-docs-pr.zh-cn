---
title: Excel JavaScript API 概述
description: ''
ms.date: 11/01/2018
localization_priority: Priority
ms.openlocfilehash: 34183c561b3da3e01a996f08761c753f204c766b
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/23/2019
ms.locfileid: "29388750"
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

## <a name="excel-javascript-api-open-specifications"></a>Excel JavaScript API 开放性规范

在我们设计和开发用于 Excel 加载项的新 API 时，我们将使它们可在[开放 API 规范](../openspec.md)页面上接收反馈。 了解即将推出的面向 Excel JavaScript API 的新功能，并提供对我们的设计规范的宝贵意见。

## <a name="excel-javascript-api-requirement-sets"></a>Excel JavaScript API 要求集

要求集是指各组已命名的 API 成员。 Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 主机是否支持加载项所需的 API。 有关 Excel JavaScript API 要求集的详细信息，请参阅 [Excel JavaScript API 要求集](../requirement-sets/excel-api-requirement-sets.md)文章。

## <a name="excel-javascript-api-reference"></a>Excel JavaScript API 参考

有关 Excel JavaScript API 的详细信息，请参阅 [Excel JavaScript API 参考文档](/javascript/api/excel)。

## <a name="see-also"></a>另请参阅

- [Excel 加载项概述](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-overview)
- [Office 加载项平台概述](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
- [GitHub 上的 Excel 加载项示例](https://github.com/OfficeDev?utf8=%E2%9C%93&q=Excel)
