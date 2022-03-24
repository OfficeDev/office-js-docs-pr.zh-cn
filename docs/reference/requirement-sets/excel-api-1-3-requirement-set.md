---
title: Excel JavaScript API 要求集 1.3
description: 有关 ExcelApi 1.3 要求集的详细信息。
ms.date: 11/09/2020
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 1bf8bc604c2c770f517878193994c1ed32640da1
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/23/2022
ms.locfileid: "63745337"
---
# <a name="whats-new-in-excel-javascript-api-13"></a>Excel JavaScript API 1.3 的最近更新

ExcelApi 1.3 增加了对数据绑定和基本数据透视表访问的支持。

## <a name="api-list"></a>API 列表

下表列出了 JavaScript API 要求Excel集 1.3 中的 API。 若要查看受 Excel JavaScript API 要求集 1.3 或更早版本支持的所有 API 的 API 参考文档，请参阅[要求集 1.3](/javascript/api/excel?view=excel-js-1.3&preserve-view=true) 或更早中的 Excel API。

| 类 | 域 | 说明 |
|:---|:---|:---|
|[Binding](/javascript/api/excel/excel.binding)|[delete()](/javascript/api/excel/excel.binding#excel-excel-binding-delete-member(1))|删除 binding 对象。|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[add (range： Range \| string， bindingType： Excel.BindingType，id：string) ](/javascript/api/excel/excel.bindingcollection#excel-excel-bindingcollection-add-member(1))|将新的 binding 对象添加到特定区域。|
||[addFromNamedItem (name： string， bindingType： Excel。BindingType，id：string) ](/javascript/api/excel/excel.bindingcollection#excel-excel-bindingcollection-addfromnameditem-member(1))|根据工作簿中的命名项添加新的 binding 对象。|
||[addFromSelection (bindingType： Excel。BindingType，id：string) ](/javascript/api/excel/excel.bindingcollection#excel-excel-bindingcollection-addfromselection-member(1))|根据当前选择的内容添加新的 binding 对象。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[name](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-name-member)|PivotTable 对象的名称。|
||[refresh()](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-refresh-member(1))|刷新 PivotTable 对象。|
||[worksheet](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-worksheet-member)|包含当前 PivotTable 对象的工作表。|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[getItem(name: string)](/javascript/api/excel/excel.pivottablecollection#excel-excel-pivottablecollection-getitem-member(1))|按名称获取 PivotTable 对象。|
||[items](/javascript/api/excel/excel.pivottablecollection#excel-excel-pivottablecollection-items-member)|获取此集合中已加载的子项。|
||[refreshAll () ](/javascript/api/excel/excel.pivottablecollection#excel-excel-pivottablecollection-refreshall-member(1))|刷新集合中的所有数据透视表。|
|[范围](/javascript/api/excel/excel.range)|[getVisibleView () ](/javascript/api/excel/excel.range#excel-excel-range-getvisibleview-member(1))|表示当前 range 对象的可见行。|
|[RangeView](/javascript/api/excel/excel.rangeview)|[cellAddresses](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-celladdresses-member)|表示 的单元格地址 `RangeView`。|
||[columnCount](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-columncount-member)|可见列数。|
||[formulas](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-formulas-member)|表示采用 A1 表示法的公式。|
||[formulasLocal](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-formulaslocal-member)|表示采用 A1 样式表示法的公式，使用用户的语言和数字格式区域设置。|
||[formulasR1C1](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-formulasr1c1-member)|表示采用 R1C1 样式表示法的公式。|
||[getRange()](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-getrange-member(1))|获取与当前 关联的父区域 `RangeView`。|
||[index](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-index-member)|返回一个值，该值代表 的索引 `RangeView`。|
||[numberFormat](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-numberformat-member)|表示 Excel 中指定单元格的数字格式代码。|
||[rowCount](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-rowcount-member)|可见行数。|
||[rows](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-rows-member)|表示一组与 range 相关联的 RangeView。|
||[text](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-text-member)|指定区域的文本值。|
||[valueTypes](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-valuetypes-member)|表示每个单元格的数据类型。|
||[values](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-values-member)|表示指定的 RangeView 的原始值。|
|[RangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|[getItemAt(index: number)](/javascript/api/excel/excel.rangeviewcollection#excel-excel-rangeviewcollection-getitemat-member(1))|通过索引 `RangeView` 获取行。|
||[items](/javascript/api/excel/excel.rangeviewcollection#excel-excel-rangeviewcollection-items-member)|获取此集合中已加载的子项。|
|[Table](/javascript/api/excel/excel.table)|[highlightFirstColumn](/javascript/api/excel/excel.table#excel-excel-table-highlightfirstcolumn-member)|指定第一列是否包含特殊格式。|
||[highlightLastColumn](/javascript/api/excel/excel.table#excel-excel-table-highlightlastcolumn-member)|指定最后一列是否包含特殊格式。|
||[showBandedColumns](/javascript/api/excel/excel.table#excel-excel-table-showbandedcolumns-member)|指定列是否显示带格式，其中奇数列的突出显示方式与偶数列不同，以便更轻松地阅读表格。|
||[showBandedRows](/javascript/api/excel/excel.table#excel-excel-table-showbandedrows-member)|指定行是否显示带格式，其中奇数行的突出显示方式与偶数行不同，以便更轻松地读取表。|
||[showFilterButton](/javascript/api/excel/excel.table#excel-excel-table-showfilterbutton-member)|指定筛选按钮是否在每个列标题的顶部可见。|
|[Workbook](/javascript/api/excel/excel.workbook)|[pivotTables](/javascript/api/excel/excel.workbook#excel-excel-workbook-pivottables-member)|表示一组与 workbook 相关联的 PivotTable 对象。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[pivotTables](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-pivottables-member)|一组属于工作表的数据透视表对象。|

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 参考文档](/javascript/api/excel?view=excel-js-1.3&preserve-view=true)
- [Excel JavaScript API 要求集](excel-api-requirement-sets.md)
