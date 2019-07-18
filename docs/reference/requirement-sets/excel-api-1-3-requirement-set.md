---
title: Excel JavaScript API 要求集1。3
description: 有关 ExcelApi 1.3 要求集的详细信息
ms.date: 07/11/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 4698b0fad3122c8ecf52117c35d4928305d812fc
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771993"
---
# <a name="whats-new-in-excel-javascript-api-13"></a>Excel JavaScript API 1.3 的最近更新

ExcelApi 1.3 增加了对数据绑定和基本数据透视表访问的支持。

## <a name="api-list"></a>API 列表

| Class | 域 | 说明 |
|:---|:---|:---|
|[Binding](/javascript/api/excel/excel.binding)|[delete()](/javascript/api/excel/excel.binding#delete--)|删除 binding 对象。|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[add (range: Range \| String, bindingType: "range" \| "Table" \| "Text", id: string)](/javascript/api/excel/excel.bindingcollection#add-range--bindingtype--id-)|将新的 binding 对象添加到特定区域。|
||[add (range: Range \| String, BindingType: bindingType, id: string)](/javascript/api/excel/excel.bindingcollection#add-range--bindingtype--id-)|将新的 binding 对象添加到特定区域。|
||[addFromNamedItem (name: string, bindingType: "Range" \| "Table" \| "Text", id: string)](/javascript/api/excel/excel.bindingcollection#addfromnameditem-name--bindingtype--id-)|根据工作簿中的命名项添加新的 binding 对象。|
||[addFromNamedItem (name: string, bindingType: BindingType, id: string)](/javascript/api/excel/excel.bindingcollection#addfromnameditem-name--bindingtype--id-)|根据工作簿中的命名项添加新的 binding 对象。|
||[addFromSelection (bindingType: "Range" \| "表" \| "Text", id: string)](/javascript/api/excel/excel.bindingcollection#addfromselection-bindingtype--id-)|根据当前选择的内容添加新的 binding 对象。|
||[addFromSelection (bindingType: BindingType, id: string)](/javascript/api/excel/excel.bindingcollection#addfromselection-bindingtype--id-)|根据当前选择的内容添加新的 binding 对象。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[name](/javascript/api/excel/excel.pivottable#name)|PivotTable 对象的名称。|
||[worksheet](/javascript/api/excel/excel.pivottable#worksheet)|包含当前 PivotTable 对象的工作表。|
||[refresh()](/javascript/api/excel/excel.pivottable#refresh--)|刷新 PivotTable 对象。|
||[set (properties: Excel. 数据透视表)](/javascript/api/excel/excel.pivottable#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: PivotTableUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.pivottable#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[getItem(name: string)](/javascript/api/excel/excel.pivottablecollection#getitem-name-)|按名称获取 PivotTable 对象。|
||[items](/javascript/api/excel/excel.pivottablecollection#items)|获取此集合中已加载的子项。|
||[refreshAll ()](/javascript/api/excel/excel.pivottablecollection#refreshall--)|刷新集合中的所有数据透视表。|
|[PivotTableCollectionLoadOptions](/javascript/api/excel/excel.pivottablecollectionloadoptions)|[$all](/javascript/api/excel/excel.pivottablecollectionloadoptions#$all)||
||[name](/javascript/api/excel/excel.pivottablecollectionloadoptions#name)|对于集合中的每一项: 数据透视表的名称。|
||[worksheet](/javascript/api/excel/excel.pivottablecollectionloadoptions#worksheet)|对于集合中的每一项: 包含当前数据透视表的工作表。|
|[PivotTableData](/javascript/api/excel/excel.pivottabledata)|[name](/javascript/api/excel/excel.pivottabledata#name)|PivotTable 对象的名称。|
|[PivotTableLoadOptions](/javascript/api/excel/excel.pivottableloadoptions)|[$all](/javascript/api/excel/excel.pivottableloadoptions#$all)||
||[name](/javascript/api/excel/excel.pivottableloadoptions#name)|PivotTable 对象的名称。|
||[worksheet](/javascript/api/excel/excel.pivottableloadoptions#worksheet)|包含当前 PivotTable 对象的工作表。|
|[PivotTableUpdateData](/javascript/api/excel/excel.pivottableupdatedata)|[name](/javascript/api/excel/excel.pivottableupdatedata#name)|PivotTable 对象的名称。|
|[Range](/javascript/api/excel/excel.range)|[getVisibleView ()](/javascript/api/excel/excel.range#getvisibleview--)|表示当前 range 对象的可见行。|
|[RangeView](/javascript/api/excel/excel.rangeview)|[formulas](/javascript/api/excel/excel.rangeview#formulas)|表示采用 A1 表示法的公式。|
||[formulasLocal](/javascript/api/excel/excel.rangeview#formulaslocal)|表示采用 A1 样式表示法的公式，使用用户的语言和数字格式区域设置。例如，英语中的公式 "=SUM(A1, 1.5)" 在德语中将变为 "=SUMME(A1; 1,5)"。|
||[formulasR1C1](/javascript/api/excel/excel.rangeview#formulasr1c1)|表示采用 R1C1 表示法的公式。|
||[getRange()](/javascript/api/excel/excel.rangeview#getrange--)|获取与当前 RangeView 相关联的父 range。|
||[numberFormat](/javascript/api/excel/excel.rangeview#numberformat)|表示 Excel 中指定单元格的数字格式代码。|
||[cellAddresses](/javascript/api/excel/excel.rangeview#celladdresses)|表示 RangeView 的单元格地址。 只读。|
||[columnCount](/javascript/api/excel/excel.rangeview#columncount)|返回可见列数。 只读。|
||[index](/javascript/api/excel/excel.rangeview#index)|返回表示 RangeView 的索引的值。 只读。|
||[rowCount](/javascript/api/excel/excel.rangeview#rowcount)|返回可见行数。 只读。|
||[rows](/javascript/api/excel/excel.rangeview#rows)|表示一组与 range 相关联的 RangeView。 只读。|
||[text](/javascript/api/excel/excel.rangeview#text)|指定区域的文本值。文本值与单元格宽度无关。在 Excel UI 中替代 # 符号不会影响 API 返回的文本值。只读。|
||[valueTypes](/javascript/api/excel/excel.rangeview#valuetypes)|表示每个单元格的数据类型。 只读。|
||[set (properties: RangeView)](/javascript/api/excel/excel.rangeview#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: RangeViewUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.rangeview#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
||[values](/javascript/api/excel/excel.rangeview#values)|表示指定的 RangeView 的原始值。 返回的数据可能是字符串、数字，也可能是布尔值。 包含错误的单元格将返回错误字符串。|
|[RangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|[getItemAt(index: number)](/javascript/api/excel/excel.rangeviewcollection#getitemat-index-)|通过其索引获取 RangeView 行。 从零开始编制索引。|
||[items](/javascript/api/excel/excel.rangeviewcollection#items)|获取此集合中已加载的子项。|
|[RangeViewCollectionLoadOptions](/javascript/api/excel/excel.rangeviewcollectionloadoptions)|[$all](/javascript/api/excel/excel.rangeviewcollectionloadoptions#$all)||
||[cellAddresses](/javascript/api/excel/excel.rangeviewcollectionloadoptions#celladdresses)|对于集合中的每一项: 代表 RangeView 的单元格地址。 只读。|
||[columnCount](/javascript/api/excel/excel.rangeviewcollectionloadoptions#columncount)|对于集合中的每一项: 返回可见列的数量。 只读。|
||[formulas](/javascript/api/excel/excel.rangeviewcollectionloadoptions#formulas)|对于集合中的每一项: 代表 A1 样式表示法中的公式。|
||[formulasLocal](/javascript/api/excel/excel.rangeviewcollectionloadoptions#formulaslocal)|对于集合中的每一项: 代表 A1 样式表示法中的公式, 位于用户的语言和数字格式设置区域中。  例如，英语中的公式 "=SUM(A1, 1.5)" 在德语中将变为 "=SUMME(A1; 1,5)"。|
||[formulasR1C1](/javascript/api/excel/excel.rangeviewcollectionloadoptions#formulasr1c1)|对于集合中的每一项: 以 R1C1 样式表示法表示的公式。|
||[index](/javascript/api/excel/excel.rangeviewcollectionloadoptions#index)|对于集合中的每一项: 返回一个值, 该值表示 RangeView 的索引。 只读。|
||[numberFormat](/javascript/api/excel/excel.rangeviewcollectionloadoptions#numberformat)|对于集合中的每一项: 代表给定单元格的 Excel 数字格式代码。|
||[rowCount](/javascript/api/excel/excel.rangeviewcollectionloadoptions#rowcount)|对于集合中的每一项: 返回可见行的数目。 只读。|
||[text](/javascript/api/excel/excel.rangeviewcollectionloadoptions#text)|对于集合中的每一项: 指定区域的文本值。 文本值与单元格宽度无关。 在 Excel UI 中替代 # 符号不会影响 API 返回的文本值。 只读。|
||[valueTypes](/javascript/api/excel/excel.rangeviewcollectionloadoptions#valuetypes)|对于集合中的每一项: 代表每个单元格的数据类型。 只读。|
||[values](/javascript/api/excel/excel.rangeviewcollectionloadoptions#values)|对于集合中的每一项: 代表指定区域视图的原始值。 返回的数据可能是字符串、数字，也可能是布尔值。 包含错误的单元格将返回错误字符串。|
|[RangeViewData](/javascript/api/excel/excel.rangeviewdata)|[cellAddresses](/javascript/api/excel/excel.rangeviewdata#celladdresses)|表示 RangeView 的单元格地址。 只读。|
||[columnCount](/javascript/api/excel/excel.rangeviewdata#columncount)|返回可见列数。 只读。|
||[formulas](/javascript/api/excel/excel.rangeviewdata#formulas)|表示采用 A1 表示法的公式。|
||[formulasLocal](/javascript/api/excel/excel.rangeviewdata#formulaslocal)|表示采用 A1 样式表示法的公式，使用用户的语言和数字格式区域设置。例如，英语中的公式 "=SUM(A1, 1.5)" 在德语中将变为 "=SUMME(A1; 1,5)"。|
||[formulasR1C1](/javascript/api/excel/excel.rangeviewdata#formulasr1c1)|表示采用 R1C1 表示法的公式。|
||[index](/javascript/api/excel/excel.rangeviewdata#index)|返回表示 RangeView 的索引的值。 只读。|
||[numberFormat](/javascript/api/excel/excel.rangeviewdata#numberformat)|表示 Excel 中指定单元格的数字格式代码。|
||[rowCount](/javascript/api/excel/excel.rangeviewdata#rowcount)|返回可见行数。 只读。|
||[rows](/javascript/api/excel/excel.rangeviewdata#rows)|表示一组与 range 相关联的 RangeView。 只读。|
||[text](/javascript/api/excel/excel.rangeviewdata#text)|指定区域的文本值。文本值与单元格宽度无关。在 Excel UI 中替代 # 符号不会影响 API 返回的文本值。只读。|
||[valueTypes](/javascript/api/excel/excel.rangeviewdata#valuetypes)|表示每个单元格的数据类型。 只读。|
||[values](/javascript/api/excel/excel.rangeviewdata#values)|表示指定的 RangeView 的原始值。 返回的数据可能是字符串、数字，也可能是布尔值。 包含错误的单元格将返回错误字符串。|
|[RangeViewLoadOptions](/javascript/api/excel/excel.rangeviewloadoptions)|[$all](/javascript/api/excel/excel.rangeviewloadoptions#$all)||
||[cellAddresses](/javascript/api/excel/excel.rangeviewloadoptions#celladdresses)|表示 RangeView 的单元格地址。 只读。|
||[columnCount](/javascript/api/excel/excel.rangeviewloadoptions#columncount)|返回可见列数。 只读。|
||[formulas](/javascript/api/excel/excel.rangeviewloadoptions#formulas)|表示采用 A1 表示法的公式。|
||[formulasLocal](/javascript/api/excel/excel.rangeviewloadoptions#formulaslocal)|表示采用 A1 样式表示法的公式，使用用户的语言和数字格式区域设置。例如，英语中的公式 "=SUM(A1, 1.5)" 在德语中将变为 "=SUMME(A1; 1,5)"。|
||[formulasR1C1](/javascript/api/excel/excel.rangeviewloadoptions#formulasr1c1)|表示采用 R1C1 表示法的公式。|
||[index](/javascript/api/excel/excel.rangeviewloadoptions#index)|返回表示 RangeView 的索引的值。 只读。|
||[numberFormat](/javascript/api/excel/excel.rangeviewloadoptions#numberformat)|表示 Excel 中指定单元格的数字格式代码。|
||[rowCount](/javascript/api/excel/excel.rangeviewloadoptions#rowcount)|返回可见行数。 只读。|
||[text](/javascript/api/excel/excel.rangeviewloadoptions#text)|指定区域的文本值。文本值与单元格宽度无关。在 Excel UI 中替代 # 符号不会影响 API 返回的文本值。只读。|
||[valueTypes](/javascript/api/excel/excel.rangeviewloadoptions#valuetypes)|表示每个单元格的数据类型。 只读。|
||[values](/javascript/api/excel/excel.rangeviewloadoptions#values)|表示指定的 RangeView 的原始值。 返回的数据可能是字符串、数字，也可能是布尔值。 包含错误的单元格将返回错误字符串。|
|[RangeViewUpdateData](/javascript/api/excel/excel.rangeviewupdatedata)|[formulas](/javascript/api/excel/excel.rangeviewupdatedata#formulas)|表示采用 A1 表示法的公式。|
||[formulasLocal](/javascript/api/excel/excel.rangeviewupdatedata#formulaslocal)|表示采用 A1 样式表示法的公式，使用用户的语言和数字格式区域设置。例如，英语中的公式 "=SUM(A1, 1.5)" 在德语中将变为 "=SUMME(A1; 1,5)"。|
||[formulasR1C1](/javascript/api/excel/excel.rangeviewupdatedata#formulasr1c1)|表示采用 R1C1 表示法的公式。|
||[numberFormat](/javascript/api/excel/excel.rangeviewupdatedata#numberformat)|表示 Excel 中指定单元格的数字格式代码。|
||[values](/javascript/api/excel/excel.rangeviewupdatedata#values)|表示指定的 RangeView 的原始值。 返回的数据可能是字符串、数字，也可能是布尔值。 包含错误的单元格将返回错误字符串。|
|[Table](/javascript/api/excel/excel.table)|[highlightFirstColumn](/javascript/api/excel/excel.table#highlightfirstcolumn)|指明第一列是否包含特殊格式。|
||[highlightLastColumn](/javascript/api/excel/excel.table#highlightlastcolumn)|指明最后一列是否包含特殊格式。|
||[showBandedColumns](/javascript/api/excel/excel.table#showbandedcolumns)|指明列是否采用镶边格式来以不同的方式突出显示奇数列与偶数列，让表更易于阅读。|
||[showBandedRows](/javascript/api/excel/excel.table#showbandedrows)|指明行是否采用镶边格式来以不同的方式突出显示奇数行与偶数行，让表更易于阅读。|
||[showFilterButton](/javascript/api/excel/excel.table#showfilterbutton)|指明是否在每个列标题的顶部显示筛选器按钮。仅当 table 中包含标题行时，才允许设定此设置。|
|[TableCollectionLoadOptions](/javascript/api/excel/excel.tablecollectionloadoptions)|[highlightFirstColumn](/javascript/api/excel/excel.tablecollectionloadoptions#highlightfirstcolumn)|对于集合中的每一项: 指示第一列是否包含特殊格式。|
||[highlightLastColumn](/javascript/api/excel/excel.tablecollectionloadoptions#highlightlastcolumn)|对于集合中的每一项: 指示最后一列是否包含特殊格式。|
||[showBandedColumns](/javascript/api/excel/excel.tablecollectionloadoptions#showbandedcolumns)|对于集合中的每一项: 指示列是否显示条带格式, 其中奇数列以不同的方式突出显示, 而不是为了使表更易于阅读。|
||[showBandedRows](/javascript/api/excel/excel.tablecollectionloadoptions#showbandedrows)|对于集合中的每一项: 指示行是否显示条带格式, 其中奇数行以不同的方式突出显示, 即使是偶数行, 也可以使表更易于阅读。|
||[showFilterButton](/javascript/api/excel/excel.tablecollectionloadoptions#showfilterbutton)|对于集合中的每一项: 指示筛选按钮是否显示在每个列标头的顶部。 仅当 table 中包含标题行时，才允许设定此设置。|
|[TableData](/javascript/api/excel/excel.tabledata)|[highlightFirstColumn](/javascript/api/excel/excel.tabledata#highlightfirstcolumn)|指明第一列是否包含特殊格式。|
||[highlightLastColumn](/javascript/api/excel/excel.tabledata#highlightlastcolumn)|指明最后一列是否包含特殊格式。|
||[showBandedColumns](/javascript/api/excel/excel.tabledata#showbandedcolumns)|指明列是否采用镶边格式来以不同的方式突出显示奇数列与偶数列，让表更易于阅读。|
||[showBandedRows](/javascript/api/excel/excel.tabledata#showbandedrows)|指明行是否采用镶边格式来以不同的方式突出显示奇数行与偶数行，让表更易于阅读。|
||[showFilterButton](/javascript/api/excel/excel.tabledata#showfilterbutton)|指明是否在每个列标题的顶部显示筛选器按钮。仅当 table 中包含标题行时，才允许设定此设置。|
|[TableLoadOptions](/javascript/api/excel/excel.tableloadoptions)|[highlightFirstColumn](/javascript/api/excel/excel.tableloadoptions#highlightfirstcolumn)|指明第一列是否包含特殊格式。|
||[highlightLastColumn](/javascript/api/excel/excel.tableloadoptions#highlightlastcolumn)|指明最后一列是否包含特殊格式。|
||[showBandedColumns](/javascript/api/excel/excel.tableloadoptions#showbandedcolumns)|指明列是否采用镶边格式来以不同的方式突出显示奇数列与偶数列，让表更易于阅读。|
||[showBandedRows](/javascript/api/excel/excel.tableloadoptions#showbandedrows)|指明行是否采用镶边格式来以不同的方式突出显示奇数行与偶数行，让表更易于阅读。|
||[showFilterButton](/javascript/api/excel/excel.tableloadoptions#showfilterbutton)|指明是否在每个列标题的顶部显示筛选器按钮。仅当 table 中包含标题行时，才允许设定此设置。|
|[TableUpdateData](/javascript/api/excel/excel.tableupdatedata)|[highlightFirstColumn](/javascript/api/excel/excel.tableupdatedata#highlightfirstcolumn)|指明第一列是否包含特殊格式。|
||[highlightLastColumn](/javascript/api/excel/excel.tableupdatedata#highlightlastcolumn)|指明最后一列是否包含特殊格式。|
||[showBandedColumns](/javascript/api/excel/excel.tableupdatedata#showbandedcolumns)|指明列是否采用镶边格式来以不同的方式突出显示奇数列与偶数列，让表更易于阅读。|
||[showBandedRows](/javascript/api/excel/excel.tableupdatedata#showbandedrows)|指明行是否采用镶边格式来以不同的方式突出显示奇数行与偶数行，让表更易于阅读。|
||[showFilterButton](/javascript/api/excel/excel.tableupdatedata#showfilterbutton)|指明是否在每个列标题的顶部显示筛选器按钮。仅当 table 中包含标题行时，才允许设定此设置。|
|[Workbook](/javascript/api/excel/excel.workbook)|[数据](/javascript/api/excel/excel.workbook#pivottables)|表示一组与 workbook 相关联的 PivotTable 对象。 只读。|
|[WorkbookData](/javascript/api/excel/excel.workbookdata)|[数据](/javascript/api/excel/excel.workbookdata#pivottables)|表示一组与 workbook 相关联的 PivotTable 对象。 只读。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[数据](/javascript/api/excel/excel.worksheet#pivottables)|一组属于 worksheet 的 PivotTable 对象。 只读。|
|[WorksheetData](/javascript/api/excel/excel.worksheetdata)|[数据](/javascript/api/excel/excel.worksheetdata#pivottables)|一组属于 worksheet 的 PivotTable 对象。 只读。|

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 参考文档](/javascript/api/excel)
- [Excel JavaScript API 要求集](./excel-api-requirement-sets.md)
