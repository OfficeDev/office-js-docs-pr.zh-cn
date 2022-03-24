---
title: Excel JavaScript API 要求集 1.10
description: 有关 ExcelApi 1.10 要求集的详细信息。
ms.date: 04/02/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 53cf0ec55a26f02a615a3c5eee0b718b818790d0
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/23/2022
ms.locfileid: "63746338"
---
# <a name="whats-new-in-excel-javascript-api-110"></a>JavaScript API 1.10 Excel的新增功能

ExcelApi 1.10 引入了关键功能，如注释、大纲和切片器。 它还添加了对工作表级单击和排序的事件支持。

| 功能区域 | 说明 | 相关对象 |
|:--- |:--- |:--- |
| [备注](../../excel/excel-add-ins-comments.md) | 添加、编辑和删除备注。 | [Comment](/javascript/api/excel/excel.comment)、[CommentCollection](/javascript/api/excel/excel.commentcollection) |
| [大纲](../../excel/excel-add-ins-ranges-group.md) | 对行和列进行分组以形成可折叠的分级显示。 | [Range](/javascript/api/excel/excel.range)、 [Worksheet](/javascript/api/excel/excel.worksheet) |
| [Slicers](../../excel/excel-add-ins-pivottables.md#filter-with-slicers) | 在表格和数据透视表中插入和配置切片器。 | [Slicer](/javascript/api/excel/excel.slicer) |
| [更多工作表事件](../../excel/excel-add-ins-events.md) | 侦听工作表中的单击和排序事件。 | [工作表 (事件) ](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-events-member) |

## <a name="api-list"></a>API 列表

下表列出了 JavaScript API 要求集 1.10 Excel中的 API。 若要查看受 Excel JavaScript API 要求集 1.10 或更早版本支持的所有 API 的 API 参考文档，请参阅[要求集 1.10](/javascript/api/excel?view=excel-js-1.10&preserve-view=true) 或更早中的 Excel API。

| 类 | 域 | 说明 |
|:---|:---|:---|
|[Comment](/javascript/api/excel/excel.comment)|[authorEmail](/javascript/api/excel/excel.comment#excel-excel-comment-authoremail-member)|获取批注作者的电子邮件。|
||[authorName](/javascript/api/excel/excel.comment#excel-excel-comment-authorname-member)|获取批注作者的姓名。|
||[content](/javascript/api/excel/excel.comment#excel-excel-comment-content-member)|注释的内容。|
||[creationDate](/javascript/api/excel/excel.comment#excel-excel-comment-creationdate-member)|获取批注的创建时间。|
||[delete()](/javascript/api/excel/excel.comment#excel-excel-comment-delete-member(1))|删除注释以及所有连接的回复。|
||[getLocation()](/javascript/api/excel/excel.comment#excel-excel-comment-getlocation-member(1))|获取此批注所在的单元格。|
||[id](/javascript/api/excel/excel.comment#excel-excel-comment-id-member)|指定注释标识符。|
||[replies](/javascript/api/excel/excel.comment#excel-excel-comment-replies-member)|表示与批注关联的回复对象的集合。|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[add (cellAddress： Range \| string， content： string， contentType？： Excel.ContentType) ](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-add-member(1))|使用给定单元格上的给定内容创建新批注。|
||[getCount()](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-getcount-member(1))|获取集合中的批注数量。|
||[getItem(commentId: string)](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-getitem-member(1))|根据其 ID 从集合中获取批注。|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-getitemat-member(1))|根据其位置从集合中获取批注。|
||[getItemByCell(cellAddress: Range \| string)](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-getitembycell-member(1))|从指定单元格获取的批注。|
||[getItemByReplyId(replyId: string)](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-getitembyreplyid-member(1))|获取给定答复所连接到的批注。|
||[items](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-items-member)|获取此集合中已加载的子项。|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[authorEmail](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-authoremail-member)|获取批注回复作者的电子邮件。|
||[authorName](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-authorname-member)|获取批注回复作者的姓名。|
||[content](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-content-member)|批注回复的内容。|
||[creationDate](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-creationdate-member)|获取批注回复的创建时间。|
||[delete()](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-delete-member(1))|删除批注回复。|
||[getLocation()](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-getlocation-member(1))|获取此批注回复所在的单元格。|
||[getParentComment()](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-getparentcomment-member(1))|获取此回复的父批注。|
||[id](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-id-member)|指定批注回复标识符。|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[add(content: string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-add-member(1))|为批注创建批注回复。|
||[getCount()](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-getcount-member(1))|获取集合中的批注回复数量。|
||[getItem(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-getitem-member(1))|返回由其 ID 标识的批注回复。|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-getitemat-member(1))|根据其在集合中的位置获取批注回复。|
||[items](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-items-member)|获取此集合中已加载的子项。|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[enableFieldList](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-enablefieldlist-member)|指定字段列表是否可在 UI 中显示。|
|[PivotTableStyle](/javascript/api/excel/excel.pivottablestyle)|[delete()](/javascript/api/excel/excel.pivottablestyle#excel-excel-pivottablestyle-delete-member(1))|删除数据透视表样式。|
||[duplicate()](/javascript/api/excel/excel.pivottablestyle#excel-excel-pivottablestyle-duplicate-member(1))|使用所有样式元素的副本创建此数据透视表样式的副本。|
||[name](/javascript/api/excel/excel.pivottablestyle#excel-excel-pivottablestyle-name-member)|获取数据透视表样式的名称。|
||[readOnly](/javascript/api/excel/excel.pivottablestyle#excel-excel-pivottablestyle-readonly-member)|指定此对象 `PivotTableStyle` 是否只读。|
|[PivotTableStyleCollection](/javascript/api/excel/excel.pivottablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.pivottablestylecollection#excel-excel-pivottablestylecollection-add-member(1))|创建具有指定 `PivotTableStyle` 名称的空白。|
||[getCount()](/javascript/api/excel/excel.pivottablestylecollection#excel-excel-pivottablestylecollection-getcount-member(1))|获取集合中 PivotTable 的数量。|
||[getDefault()](/javascript/api/excel/excel.pivottablestylecollection#excel-excel-pivottablestylecollection-getdefault-member(1))|获取父对象范围的默认数据透视表样式。|
||[getItem(name: string)](/javascript/api/excel/excel.pivottablestylecollection#excel-excel-pivottablestylecollection-getitem-member(1))|按名称 `PivotTableStyle` 获取 。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablestylecollection#excel-excel-pivottablestylecollection-getitemornullobject-member(1))|按名称 `PivotTableStyle` 获取 。|
||[items](/javascript/api/excel/excel.pivottablestylecollection#excel-excel-pivottablestylecollection-items-member)|获取此集合中已加载的子项。|
||[setDefault(newDefaultStyle: PivotTableStyle \| string)](/javascript/api/excel/excel.pivottablestylecollection#excel-excel-pivottablestylecollection-setdefault-member(1))|设置在父对象范围内使用的默认数据透视表样式。|
|[范围](/javascript/api/excel/excel.range)|[group (groupOption： Excel.GroupOption) ](/javascript/api/excel/excel.range#excel-excel-range-group-member(1))|对大纲的列和行进行分组。|
||[height](/javascript/api/excel/excel.range#excel-excel-range-height-member)|返回从区域上边缘到区域下边缘的距离（100% 缩放）。以点表示。|
||[hideGroupDetails (groupOption： Excel。GroupOption) ](/javascript/api/excel/excel.range#excel-excel-range-hidegroupdetails-member(1))|隐藏行或列组的详细信息。|
||[left](/javascript/api/excel/excel.range#excel-excel-range-left-member)|返回从工作表左边缘到区域左边缘的距离（100% 缩放）。以点表示。|
||[showGroupDetails (groupOption： Excel。GroupOption) ](/javascript/api/excel/excel.range#excel-excel-range-showgroupdetails-member(1))|显示行或列组的详细信息。|
||[top](/javascript/api/excel/excel.range#excel-excel-range-top-member)|返回从工作表的上边缘到区域上边缘的距离（100% 缩放）。以点表示。|
||[取消分组 (组选项：Excel。GroupOption) ](/javascript/api/excel/excel.range#excel-excel-range-ungroup-member(1))|取消大纲的列和行的组合。|
||[width](/javascript/api/excel/excel.range#excel-excel-range-width-member)|返回从区域左边缘到区域右边缘的距离（以 100% 缩放表示）。|
|[Shape](/javascript/api/excel/excel.shape)|[copyTo(destinationSheet?: Worksheet \| string)](/javascript/api/excel/excel.shape#excel-excel-shape-copyto-member(1))|复制和粘贴 `Shape` 对象。|
||[placement](/javascript/api/excel/excel.shape#excel-excel-shape-placement-member)|表示对象如何附加到其下方的单元格。|
|[Slicer](/javascript/api/excel/excel.slicer)|[caption](/javascript/api/excel/excel.slicer#excel-excel-slicer-caption-member)|表示切片器的标题。|
||[clearFilters()](/javascript/api/excel/excel.slicer#excel-excel-slicer-clearfilters-member(1))|清除当前切片器上应用的所有筛选器。|
||[delete()](/javascript/api/excel/excel.slicer#excel-excel-slicer-delete-member(1))|删除切片器。|
||[getSelectedItems()](/javascript/api/excel/excel.slicer#excel-excel-slicer-getselecteditems-member(1))|返回所选项目密钥的数组。|
||[height](/javascript/api/excel/excel.slicer#excel-excel-slicer-height-member)|表示切片器的高度（以磅为单位）。|
||[id](/javascript/api/excel/excel.slicer#excel-excel-slicer-id-member)|表示切片器的唯一 ID。|
||[isFilterCleared](/javascript/api/excel/excel.slicer#excel-excel-slicer-isfiltercleared-member)|如果当前 `true` 在切片器上应用的所有筛选器已清除，则值为 。|
||[left](/javascript/api/excel/excel.slicer#excel-excel-slicer-left-member)|表示从切片器左侧到工作表左侧的距离（以磅为单位）。|
||[name](/javascript/api/excel/excel.slicer#excel-excel-slicer-name-member)|表示切片器的名称。|
||[selectItems(items?: string[])](/javascript/api/excel/excel.slicer#excel-excel-slicer-selectitems-member(1))|基于切片器项的键选择切片器项。|
||[slicerItems](/javascript/api/excel/excel.slicer#excel-excel-slicer-sliceritems-member)|表示属于切片器一部分的切片器项的集合。|
||[sortBy](/javascript/api/excel/excel.slicer#excel-excel-slicer-sortby-member)|表示切片器中的项目的排序顺序。|
||[style](/javascript/api/excel/excel.slicer#excel-excel-slicer-style-member)|代表切片器样式的常量值。|
||[top](/javascript/api/excel/excel.slicer#excel-excel-slicer-top-member)|表示从切片器上边缘到工作表顶部的距离（以磅为单位）。|
||[width](/javascript/api/excel/excel.slicer#excel-excel-slicer-width-member)|表示切片器的宽度（以磅为单位）。|
||[worksheet](/javascript/api/excel/excel.slicer#excel-excel-slicer-worksheet-member)|表示包含切片器的工作表。|
|[SlicerCollection](/javascript/api/excel/excel.slicercollection)|[add(slicerSource: string \| PivotTable \| Table, sourceField: string \| PivotField \| number \| TableColumn, slicerDestination?: string \| Worksheet)](/javascript/api/excel/excel.slicercollection#excel-excel-slicercollection-add-member(1))|将新切片器添加到工作簿。|
||[getCount()](/javascript/api/excel/excel.slicercollection#excel-excel-slicercollection-getcount-member(1))|返回集合中的切片器数量。|
||[getItem(key: string)](/javascript/api/excel/excel.slicercollection#excel-excel-slicercollection-getitem-member(1))|使用其名称或 ID 获取切片器对象。|
||[getItemAt(index: number)](/javascript/api/excel/excel.slicercollection#excel-excel-slicercollection-getitemat-member(1))|根据其在集合中的位置获取切片器。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.slicercollection#excel-excel-slicercollection-getitemornullobject-member(1))|使用其名称或 ID 获取切片器。|
||[items](/javascript/api/excel/excel.slicercollection#excel-excel-slicercollection-items-member)|获取此集合中已加载的子项。|
|[SlicerItem](/javascript/api/excel/excel.sliceritem)|[hasData](/javascript/api/excel/excel.sliceritem#excel-excel-sliceritem-hasdata-member)|如果切片 `true` 器项具有数据，则值为 。|
||[isSelected](/javascript/api/excel/excel.sliceritem#excel-excel-sliceritem-isselected-member)|如果选择了 `true` 切片器项，则值为 。|
||[key](/javascript/api/excel/excel.sliceritem#excel-excel-sliceritem-key-member)|表示代表切片器项的唯一值。|
||[name](/javascript/api/excel/excel.sliceritem#excel-excel-sliceritem-name-member)|表示在用户界面中显示的Excel。|
|[SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection)|[getCount()](/javascript/api/excel/excel.sliceritemcollection#excel-excel-sliceritemcollection-getcount-member(1))|返回切片器中的切片器项的数量。|
||[getItem(key: string)](/javascript/api/excel/excel.sliceritemcollection#excel-excel-sliceritemcollection-getitem-member(1))|使用其键或名称获取切片器项对象。|
||[getItemAt(index: number)](/javascript/api/excel/excel.sliceritemcollection#excel-excel-sliceritemcollection-getitemat-member(1))|根据其在集合中的位置获取切片器项。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.sliceritemcollection#excel-excel-sliceritemcollection-getitemornullobject-member(1))|使用其键或名称获取切片器项。|
||[items](/javascript/api/excel/excel.sliceritemcollection#excel-excel-sliceritemcollection-items-member)|获取此集合中已加载的子项。|
|[SlicerStyle](/javascript/api/excel/excel.slicerstyle)|[delete()](/javascript/api/excel/excel.slicerstyle#excel-excel-slicerstyle-delete-member(1))|删除切片器样式。|
||[duplicate()](/javascript/api/excel/excel.slicerstyle#excel-excel-slicerstyle-duplicate-member(1))|使用所有样式元素的副本创建此切片器样式的副本。|
||[name](/javascript/api/excel/excel.slicerstyle#excel-excel-slicerstyle-name-member)|获取切片器样式的名称。|
||[readOnly](/javascript/api/excel/excel.slicerstyle#excel-excel-slicerstyle-readonly-member)|指定此对象 `SlicerStyle` 是否只读。|
|[SlicerStyleCollection](/javascript/api/excel/excel.slicerstylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.slicerstylecollection#excel-excel-slicerstylecollection-add-member(1))|创建具有指定名称的空白切片器样式。|
||[getCount()](/javascript/api/excel/excel.slicerstylecollection#excel-excel-slicerstylecollection-getcount-member(1))|获取集合中的切片器样式数量。|
||[getDefault()](/javascript/api/excel/excel.slicerstylecollection#excel-excel-slicerstylecollection-getdefault-member(1))|获取父 `SlicerStyle` 对象的作用域的默认值。|
||[getItem(name: string)](/javascript/api/excel/excel.slicerstylecollection#excel-excel-slicerstylecollection-getitem-member(1))|按名称 `SlicerStyle` 获取 。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.slicerstylecollection#excel-excel-slicerstylecollection-getitemornullobject-member(1))|按名称 `SlicerStyle` 获取 。|
||[items](/javascript/api/excel/excel.slicerstylecollection#excel-excel-slicerstylecollection-items-member)|获取此集合中已加载的子项。|
||[setDefault(newDefaultStyle: SlicerStyle \| string)](/javascript/api/excel/excel.slicerstylecollection#excel-excel-slicerstylecollection-setdefault-member(1))|设置在父对象范围内使用的默认切片器样式。|
|[TableStyle](/javascript/api/excel/excel.tablestyle)|[delete()](/javascript/api/excel/excel.tablestyle#excel-excel-tablestyle-delete-member(1))|删除表格样式。|
||[duplicate()](/javascript/api/excel/excel.tablestyle#excel-excel-tablestyle-duplicate-member(1))|使用所有样式元素的副本创建此表格样式的副本。|
||[name](/javascript/api/excel/excel.tablestyle#excel-excel-tablestyle-name-member)|获取表格样式的名称。|
||[readOnly](/javascript/api/excel/excel.tablestyle#excel-excel-tablestyle-readonly-member)|指定此对象 `TableStyle` 是否只读。|
|[TableStyleCollection](/javascript/api/excel/excel.tablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.tablestylecollection#excel-excel-tablestylecollection-add-member(1))|创建具有指定 `TableStyle` 名称的空白。|
||[getCount()](/javascript/api/excel/excel.tablestylecollection#excel-excel-tablestylecollection-getcount-member(1))|获取集合中表格样式的数量。|
||[getDefault()](/javascript/api/excel/excel.tablestylecollection#excel-excel-tablestylecollection-getdefault-member(1))|获取父对象范围的默认表格样式。|
||[getItem(name: string)](/javascript/api/excel/excel.tablestylecollection#excel-excel-tablestylecollection-getitem-member(1))|按名称 `TableStyle` 获取 。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.tablestylecollection#excel-excel-tablestylecollection-getitemornullobject-member(1))|按名称 `TableStyle` 获取 。|
||[items](/javascript/api/excel/excel.tablestylecollection#excel-excel-tablestylecollection-items-member)|获取此集合中已加载的子项。|
||[setDefault(newDefaultStyle: TableStyle \| string)](/javascript/api/excel/excel.tablestylecollection#excel-excel-tablestylecollection-setdefault-member(1))|设置在父对象范围内使用的默认表格样式。|
|[TimelineStyle](/javascript/api/excel/excel.timelinestyle)|[delete()](/javascript/api/excel/excel.timelinestyle#excel-excel-timelinestyle-delete-member(1))|删除表格样式。|
||[duplicate()](/javascript/api/excel/excel.timelinestyle#excel-excel-timelinestyle-duplicate-member(1))|使用所有样式元素的副本创建此时间线样式的副本。|
||[name](/javascript/api/excel/excel.timelinestyle#excel-excel-timelinestyle-name-member)|获取日程表样式的名称。|
||[readOnly](/javascript/api/excel/excel.timelinestyle#excel-excel-timelinestyle-readonly-member)|指定此对象 `TimelineStyle` 是否只读。|
|[TimelineStyleCollection](/javascript/api/excel/excel.timelinestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.timelinestylecollection#excel-excel-timelinestylecollection-add-member(1))|创建具有指定 `TimelineStyle` 名称的空白。|
||[getCount()](/javascript/api/excel/excel.timelinestylecollection#excel-excel-timelinestylecollection-getcount-member(1))|获取集合中日程表样式的数量。|
||[getDefault()](/javascript/api/excel/excel.timelinestylecollection#excel-excel-timelinestylecollection-getdefault-member(1))|获取父对象范围的默认时间线样式。|
||[getItem(name: string)](/javascript/api/excel/excel.timelinestylecollection#excel-excel-timelinestylecollection-getitem-member(1))|按名称 `TimelineStyle` 获取 。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.timelinestylecollection#excel-excel-timelinestylecollection-getitemornullobject-member(1))|按名称 `TimelineStyle` 获取 。|
||[items](/javascript/api/excel/excel.timelinestylecollection#excel-excel-timelinestylecollection-items-member)|获取此集合中已加载的子项。|
||[setDefault(newDefaultStyle: TimelineStyle \| string)](/javascript/api/excel/excel.timelinestylecollection#excel-excel-timelinestylecollection-setdefault-member(1))|设置在父对象范围内使用的默认日程表样式。|
|[Workbook](/javascript/api/excel/excel.workbook)|[comments](/javascript/api/excel/excel.workbook#excel-excel-workbook-comments-member)|表示与工作簿关联的注释的集合。|
||[getActiveSlicer()](/javascript/api/excel/excel.workbook#excel-excel-workbook-getactiveslicer-member(1))|获取工作簿中当前处于活动状态的切片器。|
||[getActiveSlicerOrNullObject()](/javascript/api/excel/excel.workbook#excel-excel-workbook-getactiveslicerornullobject-member(1))|获取工作簿中当前处于活动状态的切片器。|
||[pivotTableStyles](/javascript/api/excel/excel.workbook#excel-excel-workbook-pivottablestyles-member)|表示一组与工作簿相关联的 PivotTableStyles。|
||[slicerStyles](/javascript/api/excel/excel.workbook#excel-excel-workbook-slicerstyles-member)|表示一组与工作簿相关联的 SlicerStyles。|
||[slicers](/javascript/api/excel/excel.workbook#excel-excel-workbook-slicers-member)|表示与工作簿关联的切片器的集合。|
||[tableStyles](/javascript/api/excel/excel.workbook#excel-excel-workbook-tablestyles-member)|表示一组与工作簿相关联的 TableStyles。|
||[timelineStyles](/javascript/api/excel/excel.workbook#excel-excel-workbook-timelinestyles-member)|表示一组与工作簿相关联的 TimelineStyles。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[comments](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-comments-member)|返回工作表上的所有 Comments 对象的集合。|
||[onColumnSorted](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-oncolumnsorted-member)|在已对一个或多个列进行排序时发生。|
||[onRowSorted](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onrowsorted-member)|在已对一个或多个行进行排序时发生。|
||[onSingleClicked](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onsingleclicked-member)|在工作表中发生左键单击/点击操作时发生。|
||[showOutlineLevels (rowLevels： number， columnLevels： number) ](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-showoutlinelevels-member(1))|按大纲级别显示行或列组。|
||[slicers](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-slicers-member)|返回属于工作表的切片器集合。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onColumnSorted](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-oncolumnsorted-member)|在已对一个或多个列进行排序时发生。|
||[onRowSorted](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onrowsorted-member)|在已对一个或多个行进行排序时发生。|
||[onSingleClicked](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onsingleclicked-member)|在工作表集合中发生左键单击/点击操作时发生。|
|[WorksheetColumnSortedEventArgs](/javascript/api/excel/excel.worksheetcolumnsortedeventargs)|[address](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#excel-excel-worksheetcolumnsortedeventargs-address-member)|获取区域地址，该地址表示特定工作表上的选定区域。|
||[source](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#excel-excel-worksheetcolumnsortedeventargs-source-member)|获取事件源。|
||[type](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#excel-excel-worksheetcolumnsortedeventargs-type-member)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#excel-excel-worksheetcolumnsortedeventargs-worksheetid-member)|获取发生排序的工作表的 ID。|
|[WorksheetRowSortedEventArgs](/javascript/api/excel/excel.worksheetrowsortedeventargs)|[address](/javascript/api/excel/excel.worksheetrowsortedeventargs#excel-excel-worksheetrowsortedeventargs-address-member)|获取区域地址，该地址表示特定工作表上的选定区域。|
||[source](/javascript/api/excel/excel.worksheetrowsortedeventargs#excel-excel-worksheetrowsortedeventargs-source-member)|获取事件源。|
||[type](/javascript/api/excel/excel.worksheetrowsortedeventargs#excel-excel-worksheetrowsortedeventargs-type-member)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.worksheetrowsortedeventargs#excel-excel-worksheetrowsortedeventargs-worksheetid-member)|获取发生排序的工作表的 ID。|
|[WorksheetSingleClickedEventArgs](/javascript/api/excel/excel.worksheetsingleclickedeventargs)|[address](/javascript/api/excel/excel.worksheetsingleclickedeventargs#excel-excel-worksheetsingleclickedeventargs-address-member)|获取特定工作表中表示被左键单击/点击的单元格的地址。|
||[OffsetX](/javascript/api/excel/excel.worksheetsingleclickedeventargs#excel-excel-worksheetsingleclickedeventargs-offsetx-member)|对于从右到左的语言，从左键单击/点击点到左 (或右键的距离（以) 单击/点击的单元格的网格线边缘）。|
||[OffsetY](/javascript/api/excel/excel.worksheetsingleclickedeventargs#excel-excel-worksheetsingleclickedeventargs-offsety-member)|从左键单击/点击的点到左键单击/点击的单元格的顶部网格线边缘的距离（以磅为单位）。|
||[type](/javascript/api/excel/excel.worksheetsingleclickedeventargs#excel-excel-worksheetsingleclickedeventargs-type-member)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.worksheetsingleclickedeventargs#excel-excel-worksheetsingleclickedeventargs-worksheetid-member)|获取其中单元格被左键单击/点击的工作表的 ID。|

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 参考文档](/javascript/api/excel?view=excel-js-1.10&preserve-view=true)
- [Excel JavaScript API 要求集](excel-api-requirement-sets.md)