---
title: Excel JavaScript API 要求集1。3
description: 有关 ExcelApi 1.3 要求集的详细信息。
ms.date: 11/09/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 520755fe4b77008da866098d851f47ae3833bf13
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996471"
---
# <a name="whats-new-in-excel-javascript-api-13"></a>Excel JavaScript API 1.3 的最近更新

ExcelApi 1.3 增加了对数据绑定和基本数据透视表访问的支持。

## <a name="api-list"></a>API 列表

下表列出了 Excel JavaScript API 要求集1.3 中的 Api。 若要查看 Excel JavaScript API 要求集1.3 或更早版本支持的所有 Api 的 API 参考文档，请参阅 [要求集1.3 或更早版本中的 Excel api](/javascript/api/excel?view=excel-js-1.3&preserve-view=true)。

| Class | 域 | 说明 |
|:---|:---|:---|
|[Binding](/javascript/api/excel/excel.binding)|[delete()](/javascript/api/excel/excel.binding#delete--)|删除 binding 对象。|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[add (range： Range \| string，bindingType： bindingType，id： string) ](/javascript/api/excel/excel.bindingcollection#add-range--bindingtype--id-)|将新的 binding 对象添加到特定区域。|
||[addFromNamedItem (name： string，bindingType： BindingType，id： string) ](/javascript/api/excel/excel.bindingcollection#addfromnameditem-name--bindingtype--id-)|根据工作簿中的命名项添加新的 binding 对象。|
||[addFromSelection (bindingType： BindingType，id： string) ](/javascript/api/excel/excel.bindingcollection#addfromselection-bindingtype--id-)|根据当前选择的内容添加新的 binding 对象。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[name](/javascript/api/excel/excel.pivottable#name)|PivotTable 对象的名称。|
||[worksheet](/javascript/api/excel/excel.pivottable#worksheet)|包含当前 PivotTable 对象的工作表。|
||[refresh()](/javascript/api/excel/excel.pivottable#refresh--)|刷新 PivotTable 对象。|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[getItem(name: string)](/javascript/api/excel/excel.pivottablecollection#getitem-name-)|按名称获取 PivotTable 对象。|
||[items](/javascript/api/excel/excel.pivottablecollection#items)|获取此集合中已加载的子项。|
||[refreshAll ( # B1 ](/javascript/api/excel/excel.pivottablecollection#refreshall--)|刷新集合中的所有数据透视表。|
|[Range](/javascript/api/excel/excel.range)|[getVisibleView ( # B1 ](/javascript/api/excel/excel.range#getvisibleview--)|表示当前 range 对象的可见行。|
|[RangeView](/javascript/api/excel/excel.rangeview)|[formulas](/javascript/api/excel/excel.rangeview#formulas)|表示采用 A1 表示法的公式。|
||[formulasLocal](/javascript/api/excel/excel.rangeview#formulaslocal)|表示采用 A1 样式表示法的公式，使用用户的语言和数字格式区域设置。|
||[formulasR1C1](/javascript/api/excel/excel.rangeview#formulasr1c1)|表示采用 R1C1 样式表示法的公式。|
||[getRange()](/javascript/api/excel/excel.rangeview#getrange--)|获取与当前 RangeView 相关联的父 range。|
||[numberFormat](/javascript/api/excel/excel.rangeview#numberformat)|表示 Excel 中指定单元格的数字格式代码。|
||[cellAddresses](/javascript/api/excel/excel.rangeview#celladdresses)|表示 RangeView 的单元格地址。|
||[columnCount](/javascript/api/excel/excel.rangeview#columncount)|可见列的数量。|
||[index](/javascript/api/excel/excel.rangeview#index)|返回表示 RangeView 的索引的值。|
||[rowCount](/javascript/api/excel/excel.rangeview#rowcount)|可见行的数目。|
||[rows](/javascript/api/excel/excel.rangeview#rows)|表示一组与 range 相关联的 RangeView。|
||[text](/javascript/api/excel/excel.rangeview#text)|指定区域的文本值。|
||[valueTypes](/javascript/api/excel/excel.rangeview#valuetypes)|表示每个单元格的数据类型。|
||[values](/javascript/api/excel/excel.rangeview#values)|表示指定的 RangeView 的原始值。|
|[RangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|[getItemAt(index: number)](/javascript/api/excel/excel.rangeviewcollection#getitemat-index-)|通过其索引获取 RangeView 行。|
||[items](/javascript/api/excel/excel.rangeviewcollection#items)|获取此集合中已加载的子项。|
|[Table](/javascript/api/excel/excel.table)|[highlightFirstColumn](/javascript/api/excel/excel.table#highlightfirstcolumn)|指定第一列是否包含特殊格式。|
||[highlightLastColumn](/javascript/api/excel/excel.table#highlightlastcolumn)|指定最后一列是否包含特殊格式。|
||[showBandedColumns](/javascript/api/excel/excel.table#showbandedcolumns)|指定列是否显示镶边格式，其中奇数列以不同的方式突出显示，即使是偶数，也可以使表更易于阅读。|
||[showBandedRows](/javascript/api/excel/excel.table#showbandedrows)|指定行是否显示镶边格式，其中奇数行以不同的方式突出显示，即使是偶数行，也可以使表更易于阅读。|
||[showFilterButton](/javascript/api/excel/excel.table#showfilterbutton)|指定筛选按钮是否显示在每个列标头的顶部。|
|[Workbook](/javascript/api/excel/excel.workbook)|[数据](/javascript/api/excel/excel.workbook#pivottables)|表示一组与 workbook 相关联的 PivotTable 对象。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[数据](/javascript/api/excel/excel.worksheet#pivottables)|一组属于工作表的数据透视表对象。|

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 参考文档](/javascript/api/excel?view=excel-js-1.3&preserve-view=true)
- [Excel JavaScript API 要求集](excel-api-requirement-sets.md)
