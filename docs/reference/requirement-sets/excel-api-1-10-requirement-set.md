---
title: ExcelJavaScript API 要求集 1.10
description: 有关 ExcelApi 1.10 要求集的详细信息。
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 7b620bb76f758bc2574e8bd99d2c45d3d4bfae39
ms.sourcegitcommit: 3fa8c754a47bab909e559ae3e5d4237ba27fdbe4
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/30/2021
ms.locfileid: "53671222"
---
# <a name="whats-new-in-excel-javascript-api-110"></a>JavaScript API 1.10 Excel的新增功能

ExcelApi 1.10 引入了关键功能，如注释、大纲和切片器。 它还添加了对工作表级单击和排序的事件支持。

| 功能区域 | 说明 | 相关对象 |
|:--- |:--- |:--- |
| [备注](../../excel/excel-add-ins-comments.md) | 添加、编辑和删除备注。 | [Comment](/javascript/api/excel/excel.comment)、[CommentCollection](/javascript/api/excel/excel.commentcollection) |
| [大纲](../../excel/excel-add-ins-ranges-group.md) | 对行和列进行分组以形成可折叠的分级显示。 | [Range、Worksheet](/javascript/api/excel/excel.range) [](/javascript/api/excel/excel.worksheet) |
| [Slicers](../../excel/excel-add-ins-pivottables.md#filter-with-slicers) | 在表格和数据透视表中插入和配置切片器。 | [Slicer](/javascript/api/excel/excel.slicer) |
| [更多工作表事件](../../excel/excel-add-ins-events.md) | 侦听工作表中的单击和排序事件。 | [工作表 (事件) ](/javascript/api/excel/excel.worksheet#events) |

## <a name="api-list"></a>API 列表

下表列出了 JavaScript API 要求集 1.10 Excel中的 API。 若要查看受 Excel JavaScript API 要求集 1.10 或更早版本支持的所有 API 的 API 参考文档，请参阅要求集[1.10](/javascript/api/excel?view=excel-js-1.10&preserve-view=true)或更早版本中的 Excel API。

| 类 | 域 | 说明 |
|:---|:---|:---|
|[Comment](/javascript/api/excel/excel.comment)|[content](/javascript/api/excel/excel.comment#content)|注释的内容。|
||[delete()](/javascript/api/excel/excel.comment#delete__)|删除注释以及所有连接的回复。|
||[getLocation()](/javascript/api/excel/excel.comment#getLocation__)|获取此批注所在的单元格。|
||[authorEmail](/javascript/api/excel/excel.comment#authorEmail)|获取批注作者的电子邮件。|
||[authorName](/javascript/api/excel/excel.comment#authorName)|获取批注作者的姓名。|
||[creationDate](/javascript/api/excel/excel.comment#creationDate)|获取批注的创建时间。|
||[id](/javascript/api/excel/excel.comment#id)|指定注释标识符。|
||[replies](/javascript/api/excel/excel.comment#replies)|表示与批注关联的回复对象的集合。|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[add (cellAddress： Range \| string， content： string， contentType？： Excel.ContentType) ](/javascript/api/excel/excel.commentcollection#add_cellAddress__content__contentType_)|使用给定单元格上的给定内容创建新批注。|
||[getCount()](/javascript/api/excel/excel.commentcollection#getCount__)|获取集合中的批注数量。|
||[getItem(commentId: string)](/javascript/api/excel/excel.commentcollection#getItem_commentId_)|根据其 ID 从集合中获取批注。|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentcollection#getItemAt_index_)|根据其位置从集合中获取批注。|
||[getItemByCell(cellAddress: Range \| string)](/javascript/api/excel/excel.commentcollection#getItemByCell_cellAddress_)|从指定单元格获取的批注。|
||[getItemByReplyId(replyId: string)](/javascript/api/excel/excel.commentcollection#getItemByReplyId_replyId_)|获取给定答复所连接到的注释。|
||[items](/javascript/api/excel/excel.commentcollection#items)|获取此集合中已加载的子项。|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[content](/javascript/api/excel/excel.commentreply#content)|批注回复的内容。|
||[delete()](/javascript/api/excel/excel.commentreply#delete__)|删除批注回复。|
||[getLocation()](/javascript/api/excel/excel.commentreply#getLocation__)|获取此批注回复所在的单元格。|
||[getParentComment()](/javascript/api/excel/excel.commentreply#getParentComment__)|获取此回复的父批注。|
||[authorEmail](/javascript/api/excel/excel.commentreply#authorEmail)|获取批注回复作者的电子邮件。|
||[authorName](/javascript/api/excel/excel.commentreply#authorName)|获取批注回复作者的姓名。|
||[creationDate](/javascript/api/excel/excel.commentreply#creationDate)|获取批注回复的创建时间。|
||[id](/javascript/api/excel/excel.commentreply#id)|指定批注回复标识符。|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[add(content: string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentreplycollection#add_content__contentType_)|为批注创建批注回复。|
||[getCount()](/javascript/api/excel/excel.commentreplycollection#getCount__)|获取集合中的批注回复数量。|
||[getItem(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#getItem_commentReplyId_)|返回由其 ID 标识的批注回复。|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentreplycollection#getItemAt_index_)|根据其在集合中的位置获取批注回复。|
||[items](/javascript/api/excel/excel.commentreplycollection#items)|获取此集合中已加载的子项。|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[enableFieldList](/javascript/api/excel/excel.pivotlayout#enableFieldList)|指定字段列表是否可在 UI 中显示。|
|[PivotTableStyle](/javascript/api/excel/excel.pivottablestyle)|[delete()](/javascript/api/excel/excel.pivottablestyle#delete__)|删除数据透视表样式。|
||[duplicate()](/javascript/api/excel/excel.pivottablestyle#duplicate__)|使用所有样式元素的副本创建此数据透视表样式的副本。|
||[名称](/javascript/api/excel/excel.pivottablestyle#name)|获取数据透视表样式的名称。|
||[readOnly](/javascript/api/excel/excel.pivottablestyle#readOnly)|指定此 `PivotTableStyle` 对象是否只读。|
|[PivotTableStyleCollection](/javascript/api/excel/excel.pivottablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.pivottablestylecollection#add_name__makeUniqueName_)|创建具有 `PivotTableStyle` 指定名称的空白。|
||[getCount()](/javascript/api/excel/excel.pivottablestylecollection#getCount__)|获取集合中 PivotTable 的数量。|
||[getDefault()](/javascript/api/excel/excel.pivottablestylecollection#getDefault__)|获取父对象范围的默认数据透视表样式。|
||[getItem(name: string)](/javascript/api/excel/excel.pivottablestylecollection#getItem_name_)|按 `PivotTableStyle` 名称获取 。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablestylecollection#getItemOrNullObject_name_)|按 `PivotTableStyle` 名称获取 。|
||[items](/javascript/api/excel/excel.pivottablestylecollection#items)|获取此集合中已加载的子项。|
||[setDefault(newDefaultStyle: PivotTableStyle \| string)](/javascript/api/excel/excel.pivottablestylecollection#setDefault_newDefaultStyle_)|设置在父对象范围内使用的默认数据透视表样式。|
|[区域](/javascript/api/excel/excel.range)|[group (groupOption： Excel.GroupOption) ](/javascript/api/excel/excel.range#group_groupOption_)|对大纲的列和行进行分组。|
||[hideGroupDetails (groupOption： Excel。GroupOption) ](/javascript/api/excel/excel.range#hideGroupDetails_groupOption_)|隐藏行或列组的详细信息。|
||[height](/javascript/api/excel/excel.range#height)|返回从区域上边缘到区域下边缘的距离（100% 缩放）。以点表示。|
||[left](/javascript/api/excel/excel.range#left)|返回从工作表左边缘到区域左边缘的距离（100% 缩放）。以点表示。|
||[top](/javascript/api/excel/excel.range#top)|返回从工作表的上边缘到区域上边缘的距离（100% 缩放）。以点表示。|
||[width](/javascript/api/excel/excel.range#width)|返回从区域左边缘到区域右边缘的距离（以 100% 缩放表示）。|
||[showGroupDetails (groupOption： Excel。GroupOption) ](/javascript/api/excel/excel.range#showGroupDetails_groupOption_)|显示行或列组的详细信息。|
||[取消分组 (组选项：Excel。GroupOption) ](/javascript/api/excel/excel.range#ungroup_groupOption_)|取消大纲的列和行的组合。|
|[Shape](/javascript/api/excel/excel.shape)|[copyTo(destinationSheet?: Worksheet \| string)](/javascript/api/excel/excel.shape#copyTo_destinationSheet_)|复制和粘贴 `Shape` 对象。|
||[placement](/javascript/api/excel/excel.shape#placement)|表示对象如何附加到其下方的单元格。|
|[Slicer](/javascript/api/excel/excel.slicer)|[caption](/javascript/api/excel/excel.slicer#caption)|表示切片器的标题。|
||[clearFilters()](/javascript/api/excel/excel.slicer#clearFilters__)|清除当前切片器上应用的所有筛选器。|
||[delete()](/javascript/api/excel/excel.slicer#delete__)|删除切片器。|
||[getSelectedItems()](/javascript/api/excel/excel.slicer#getSelectedItems__)|返回所选项目密钥的数组。|
||[height](/javascript/api/excel/excel.slicer#height)|表示切片器的高度（以磅为单位）。|
||[left](/javascript/api/excel/excel.slicer#left)|表示从切片器左侧到工作表左侧的距离（以磅为单位）。|
||[name](/javascript/api/excel/excel.slicer#name)|表示切片器的名称。|
||[id](/javascript/api/excel/excel.slicer#id)|表示切片器的唯一 ID。|
||[isFilterCleared](/javascript/api/excel/excel.slicer#isFilterCleared)|如果 `true` 当前在切片器上应用的所有筛选器已清除，则值为 。|
||[slicerItems](/javascript/api/excel/excel.slicer#slicerItems)|表示属于切片器一部分的切片器项的集合。|
||[worksheet](/javascript/api/excel/excel.slicer#worksheet)|表示包含切片器的工作表。|
||[selectItems(items?: string[])](/javascript/api/excel/excel.slicer#selectItems_items_)|根据切片器项的键选择切片器项。|
||[sortBy](/javascript/api/excel/excel.slicer#sortBy)|表示切片器中的项目的排序顺序。|
||[style](/javascript/api/excel/excel.slicer#style)|代表切片器样式的常量值。|
||[top](/javascript/api/excel/excel.slicer#top)|表示从切片器上边缘到工作表顶部的距离（以磅为单位）。|
||[width](/javascript/api/excel/excel.slicer#width)|表示切片器的宽度（以磅为单位）。|
|[SlicerCollection](/javascript/api/excel/excel.slicercollection)|[add(slicerSource: string \| PivotTable \| Table, sourceField: string \| PivotField \| number \| TableColumn, slicerDestination?: string \| Worksheet)](/javascript/api/excel/excel.slicercollection#add_slicerSource__sourceField__slicerDestination_)|将新切片器添加到工作簿。|
||[getCount()](/javascript/api/excel/excel.slicercollection#getCount__)|返回集合中的切片器数量。|
||[getItem(key: string)](/javascript/api/excel/excel.slicercollection#getItem_key_)|使用其名称或 ID 获取切片器对象。|
||[getItemAt(index: number)](/javascript/api/excel/excel.slicercollection#getItemAt_index_)|根据其在集合中的位置获取切片器。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.slicercollection#getItemOrNullObject_key_)|使用其名称或 ID 获取切片器。|
||[items](/javascript/api/excel/excel.slicercollection#items)|获取此集合中已加载的子项。|
|[SlicerItem](/javascript/api/excel/excel.sliceritem)|[isSelected](/javascript/api/excel/excel.sliceritem#isSelected)|如果选择了 `true` 切片器项，则值为 。|
||[hasData](/javascript/api/excel/excel.sliceritem#hasData)|如果 `true` 切片器项具有数据，则值为 。|
||[key](/javascript/api/excel/excel.sliceritem#key)|表示代表切片器项的唯一值。|
||[name](/javascript/api/excel/excel.sliceritem#name)|表示在用户界面中显示的Excel。|
|[SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection)|[getCount()](/javascript/api/excel/excel.sliceritemcollection#getCount__)|返回切片器中的切片器项的数量。|
||[getItem(key: string)](/javascript/api/excel/excel.sliceritemcollection#getItem_key_)|使用其键或名称获取切片器项对象。|
||[getItemAt(index: number)](/javascript/api/excel/excel.sliceritemcollection#getItemAt_index_)|根据其在集合中的位置获取切片器项。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.sliceritemcollection#getItemOrNullObject_key_)|使用其键或名称获取切片器项。|
||[items](/javascript/api/excel/excel.sliceritemcollection#items)|获取此集合中已加载的子项。|
|[SlicerStyle](/javascript/api/excel/excel.slicerstyle)|[delete()](/javascript/api/excel/excel.slicerstyle#delete__)|删除切片器样式。|
||[duplicate()](/javascript/api/excel/excel.slicerstyle#duplicate__)|使用所有样式元素的副本创建此切片器样式的副本。|
||[名称](/javascript/api/excel/excel.slicerstyle#name)|获取切片器样式的名称。|
||[readOnly](/javascript/api/excel/excel.slicerstyle#readOnly)|指定此 `SlicerStyle` 对象是否只读。|
|[SlicerStyleCollection](/javascript/api/excel/excel.slicerstylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.slicerstylecollection#add_name__makeUniqueName_)|创建具有指定名称的空白切片器样式。|
||[getCount()](/javascript/api/excel/excel.slicerstylecollection#getCount__)|获取集合中的切片器样式数量。|
||[getDefault()](/javascript/api/excel/excel.slicerstylecollection#getDefault__)|获取 `SlicerStyle` 父对象的作用域的默认值。|
||[getItem(name: string)](/javascript/api/excel/excel.slicerstylecollection#getItem_name_)|按 `SlicerStyle` 名称获取 。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.slicerstylecollection#getItemOrNullObject_name_)|按 `SlicerStyle` 名称获取 。|
||[items](/javascript/api/excel/excel.slicerstylecollection#items)|获取此集合中已加载的子项。|
||[setDefault(newDefaultStyle: SlicerStyle \| string)](/javascript/api/excel/excel.slicerstylecollection#setDefault_newDefaultStyle_)|设置在父对象范围内使用的默认切片器样式。|
|[TableStyle](/javascript/api/excel/excel.tablestyle)|[delete()](/javascript/api/excel/excel.tablestyle#delete__)|删除表格样式。|
||[duplicate()](/javascript/api/excel/excel.tablestyle#duplicate__)|使用所有样式元素的副本创建此表格样式的副本。|
||[名称](/javascript/api/excel/excel.tablestyle#name)|获取表格样式的名称。|
||[readOnly](/javascript/api/excel/excel.tablestyle#readOnly)|指定此 `TableStyle` 对象是否只读。|
|[TableStyleCollection](/javascript/api/excel/excel.tablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.tablestylecollection#add_name__makeUniqueName_)|创建具有 `TableStyle` 指定名称的空白。|
||[getCount()](/javascript/api/excel/excel.tablestylecollection#getCount__)|获取集合中表格样式的数量。|
||[getDefault()](/javascript/api/excel/excel.tablestylecollection#getDefault__)|获取父对象范围的默认表格样式。|
||[getItem(name: string)](/javascript/api/excel/excel.tablestylecollection#getItem_name_)|按 `TableStyle` 名称获取 。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.tablestylecollection#getItemOrNullObject_name_)|按 `TableStyle` 名称获取 。|
||[items](/javascript/api/excel/excel.tablestylecollection#items)|获取此集合中已加载的子项。|
||[setDefault(newDefaultStyle: TableStyle \| string)](/javascript/api/excel/excel.tablestylecollection#setDefault_newDefaultStyle_)|设置在父对象范围内使用的默认表格样式。|
|[TimelineStyle](/javascript/api/excel/excel.timelinestyle)|[delete()](/javascript/api/excel/excel.timelinestyle#delete__)|删除表格样式。|
||[duplicate()](/javascript/api/excel/excel.timelinestyle#duplicate__)|使用所有样式元素的副本创建此时间线样式的副本。|
||[名称](/javascript/api/excel/excel.timelinestyle#name)|获取日程表样式的名称。|
||[readOnly](/javascript/api/excel/excel.timelinestyle#readOnly)|指定此 `TimelineStyle` 对象是否只读。|
|[TimelineStyleCollection](/javascript/api/excel/excel.timelinestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.timelinestylecollection#add_name__makeUniqueName_)|创建具有 `TimelineStyle` 指定名称的空白。|
||[getCount()](/javascript/api/excel/excel.timelinestylecollection#getCount__)|获取集合中日程表样式的数量。|
||[getDefault()](/javascript/api/excel/excel.timelinestylecollection#getDefault__)|获取父对象范围的默认时间线样式。|
||[getItem(name: string)](/javascript/api/excel/excel.timelinestylecollection#getItem_name_)|按 `TimelineStyle` 名称获取 。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.timelinestylecollection#getItemOrNullObject_name_)|按 `TimelineStyle` 名称获取 。|
||[items](/javascript/api/excel/excel.timelinestylecollection#items)|获取此集合中已加载的子项。|
||[setDefault(newDefaultStyle: TimelineStyle \| string)](/javascript/api/excel/excel.timelinestylecollection#setDefault_newDefaultStyle_)|设置在父对象范围内使用的默认日程表样式。|
|[Workbook](/javascript/api/excel/excel.workbook)|[getActiveSlicer()](/javascript/api/excel/excel.workbook#getActiveSlicer__)|获取工作簿中当前处于活动状态的切片器。|
||[getActiveSlicerOrNullObject()](/javascript/api/excel/excel.workbook#getActiveSlicerOrNullObject__)|获取工作簿中当前处于活动状态的切片器。|
||[comments](/javascript/api/excel/excel.workbook#comments)|表示与工作簿关联的注释的集合。|
||[pivotTableStyles](/javascript/api/excel/excel.workbook#pivotTableStyles)|表示一组与工作簿相关联的 PivotTableStyles。|
||[slicerStyles](/javascript/api/excel/excel.workbook#slicerStyles)|表示一组与工作簿相关联的 SlicerStyles。|
||[slicers](/javascript/api/excel/excel.workbook#slicers)|表示与工作簿关联的切片器的集合。|
||[tableStyles](/javascript/api/excel/excel.workbook#tableStyles)|表示一组与工作簿相关联的 TableStyles。|
||[timelineStyles](/javascript/api/excel/excel.workbook#timelineStyles)|表示一组与工作簿相关联的 TimelineStyles。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[comments](/javascript/api/excel/excel.worksheet#comments)|返回工作表上的所有 Comments 对象的集合。|
||[onColumnSorted](/javascript/api/excel/excel.worksheet#onColumnSorted)|在已对一个或多个列进行排序时发生。|
||[onRowSorted](/javascript/api/excel/excel.worksheet#onRowSorted)|在已对一个或多个行进行排序时发生。|
||[onSingleClicked](/javascript/api/excel/excel.worksheet#onSingleClicked)|在工作表中发生左键单击/点击操作时发生。|
||[slicers](/javascript/api/excel/excel.worksheet#slicers)|返回属于工作表的切片器集合。|
||[showOutlineLevels (rowLevels： number， columnLevels： number) ](/javascript/api/excel/excel.worksheet#showOutlineLevels_rowLevels__columnLevels_)|按大纲级别显示行或列组。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onColumnSorted](/javascript/api/excel/excel.worksheetcollection#onColumnSorted)|在已对一个或多个列进行排序时发生。|
||[onRowSorted](/javascript/api/excel/excel.worksheetcollection#onRowSorted)|在已对一个或多个行进行排序时发生。|
||[onSingleClicked](/javascript/api/excel/excel.worksheetcollection#onSingleClicked)|在工作表集合中发生左键单击/点击操作时发生。|
|[WorksheetColumnSortedEventArgs](/javascript/api/excel/excel.worksheetcolumnsortedeventargs)|[address](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#address)|获取区域地址，该地址表示特定工作表上的选定区域。|
||[source](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#source)|获取事件源。|
||[type](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#worksheetId)|获取发生排序的工作表的 ID。|
|[WorksheetRowSortedEventArgs](/javascript/api/excel/excel.worksheetrowsortedeventargs)|[address](/javascript/api/excel/excel.worksheetrowsortedeventargs#address)|获取区域地址，该地址表示特定工作表上的选定区域。|
||[source](/javascript/api/excel/excel.worksheetrowsortedeventargs#source)|获取事件源。|
||[type](/javascript/api/excel/excel.worksheetrowsortedeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.worksheetrowsortedeventargs#worksheetId)|获取发生排序的工作表的 ID。|
|[WorksheetSingleClickedEventArgs](/javascript/api/excel/excel.worksheetsingleclickedeventargs)|[address](/javascript/api/excel/excel.worksheetsingleclickedeventargs#address)|获取特定工作表中表示被左键单击/点击的单元格的地址。|
||[OffsetX](/javascript/api/excel/excel.worksheetsingleclickedeventargs#offsetX)|对于从右到左的语言，从左键单击/点击点到左 (或右到左) 单击/点击单元格的网格线边缘的距离（以点表示）。|
||[OffsetY](/javascript/api/excel/excel.worksheetsingleclickedeventargs#offsetY)|从左键单击/点击的点到左键单击/点击的单元格的顶部网格线边缘的距离（以磅为单位）。|
||[type](/javascript/api/excel/excel.worksheetsingleclickedeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.worksheetsingleclickedeventargs#worksheetId)|获取其中单元格被左键单击/点击的工作表的 ID。|

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 参考文档](/javascript/api/excel?view=excel-js-1.10&preserve-view=true)
- [Excel JavaScript API 要求集](excel-api-requirement-sets.md)