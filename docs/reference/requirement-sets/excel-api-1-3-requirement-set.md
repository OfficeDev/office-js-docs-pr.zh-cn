---
title: ExcelJavaScript API 要求集 1.3
description: 有关 ExcelApi 1.3 要求集的详细信息。
ms.date: 11/09/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: d3606b74e8a1099cd58631cc047a783f27a09a19
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/08/2021
ms.locfileid: "58938535"
---
# <a name="whats-new-in-excel-javascript-api-13"></a>Excel JavaScript API 1.3 的最近更新

ExcelApi 1.3 增加了对数据绑定和基本数据透视表访问的支持。

## <a name="api-list"></a>API 列表

下表列出了 JavaScript API 要求集 1.3 Excel中的 API。 若要查看受 Excel JavaScript API 要求集 1.3 或更早版本支持的所有 API 的 API 参考文档，请参阅要求集[1.3](/javascript/api/excel?view=excel-js-1.3&preserve-view=true)或更早中的 Excel API。

| 类 | 域 | 说明 |
|:---|:---|:---|
|[Binding](/javascript/api/excel/excel.binding)|[delete()](/javascript/api/excel/excel.binding#delete__)|删除 binding 对象。|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[add (range： Range \| string， bindingType： Excel.BindingType，id：string) ](/javascript/api/excel/excel.bindingcollection#add_range__bindingType__id_)|将新的 binding 对象添加到特定区域。|
||[addFromNamedItem (name： string， bindingType： Excel。BindingType，id：string) ](/javascript/api/excel/excel.bindingcollection#addFromNamedItem_name__bindingType__id_)|根据工作簿中的命名项添加新的 binding 对象。|
||[addFromSelection (bindingType： Excel。BindingType，id：string) ](/javascript/api/excel/excel.bindingcollection#addFromSelection_bindingType__id_)|根据当前选择的内容添加新的 binding 对象。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[name](/javascript/api/excel/excel.pivottable#name)|PivotTable 对象的名称。|
||[worksheet](/javascript/api/excel/excel.pivottable#worksheet)|包含当前 PivotTable 对象的工作表。|
||[refresh()](/javascript/api/excel/excel.pivottable#refresh__)|刷新 PivotTable 对象。|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[getItem(name: string)](/javascript/api/excel/excel.pivottablecollection#getItem_name_)|按名称获取 PivotTable 对象。|
||[items](/javascript/api/excel/excel.pivottablecollection#items)|获取此集合中已加载的子项。|
||[refreshAll () ](/javascript/api/excel/excel.pivottablecollection#refreshAll__)|刷新集合中的所有数据透视表。|
|[Range](/javascript/api/excel/excel.range)|[getVisibleView () ](/javascript/api/excel/excel.range#getVisibleView__)|表示当前 range 对象的可见行。|
|[RangeView](/javascript/api/excel/excel.rangeview)|[formulas](/javascript/api/excel/excel.rangeview#formulas)|表示采用 A1 表示法的公式。|
||[formulasLocal](/javascript/api/excel/excel.rangeview#formulasLocal)|表示采用 A1 样式表示法的公式，使用用户的语言和数字格式区域设置。|
||[formulasR1C1](/javascript/api/excel/excel.rangeview#formulasR1C1)|表示采用 R1C1 样式表示法的公式。|
||[getRange()](/javascript/api/excel/excel.rangeview#getRange__)|获取与当前 关联的父区域 `RangeView` 。|
||[numberFormat](/javascript/api/excel/excel.rangeview#numberFormat)|表示 Excel 中指定单元格的数字格式代码。|
||[cellAddresses](/javascript/api/excel/excel.rangeview#cellAddresses)|表示 的单元格地址 `RangeView` 。|
||[columnCount](/javascript/api/excel/excel.rangeview#columnCount)|可见列数。|
||[index](/javascript/api/excel/excel.rangeview#index)|返回一个值，该值代表 的索引 `RangeView` 。|
||[rowCount](/javascript/api/excel/excel.rangeview#rowCount)|可见行数。|
||[rows](/javascript/api/excel/excel.rangeview#rows)|表示一组与 range 相关联的 RangeView。|
||[text](/javascript/api/excel/excel.rangeview#text)|指定区域的文本值。|
||[valueTypes](/javascript/api/excel/excel.rangeview#valueTypes)|表示每个单元格的数据类型。|
||[values](/javascript/api/excel/excel.rangeview#values)|表示指定的 RangeView 的原始值。|
|[RangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|[getItemAt(index: number)](/javascript/api/excel/excel.rangeviewcollection#getItemAt_index_)|通过索引 `RangeView` 获取行。|
||[items](/javascript/api/excel/excel.rangeviewcollection#items)|获取此集合中已加载的子项。|
|[Table](/javascript/api/excel/excel.table)|[highlightFirstColumn](/javascript/api/excel/excel.table#highlightFirstColumn)|指定第一列是否包含特殊格式。|
||[highlightLastColumn](/javascript/api/excel/excel.table#highlightLastColumn)|指定最后一列是否包含特殊格式。|
||[showBandedColumns](/javascript/api/excel/excel.table#showBandedColumns)|指定列是否显示带格式，其中奇数列的突出显示方式与偶数列不同，以便更轻松地阅读表格。|
||[showBandedRows](/javascript/api/excel/excel.table#showBandedRows)|指定行是否显示带格式，其中奇数行的突出显示方式与偶数行不同，以便更轻松地读取表。|
||[showFilterButton](/javascript/api/excel/excel.table#showFilterButton)|指定筛选按钮是否在每个列标题的顶部可见。|
|[Workbook](/javascript/api/excel/excel.workbook)|[pivotTables](/javascript/api/excel/excel.workbook#pivotTables)|表示一组与 workbook 相关联的 PivotTable 对象。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[pivotTables](/javascript/api/excel/excel.worksheet#pivotTables)|一组属于工作表的数据透视表对象。|

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 参考文档](/javascript/api/excel?view=excel-js-1.3&preserve-view=true)
- [Excel JavaScript API 要求集](excel-api-requirement-sets.md)
