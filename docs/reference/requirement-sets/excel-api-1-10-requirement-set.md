---
title: Excel JavaScript API 要求集 1.10
description: 有关 ExcelApi 1.10 要求集的详细信息。
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 1bafdd2064166019c5c3f22aa4da1a2d0ec73f08
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/09/2021
ms.locfileid: "51650819"
---
# <a name="whats-new-in-excel-javascript-api-110"></a>Excel JavaScript API 1.10 中的新增功能

ExcelApi 1.10 引入了关键功能，如注释、大纲和切片器。 它还添加了对工作表级单击和排序的事件支持。

| 功能区域 | 说明 | 相关对象 |
|:--- |:--- |:--- |
| [备注](../../excel/excel-add-ins-comments.md) | 添加、编辑和删除备注。 | [Comment](/javascript/api/excel/excel.comment)、[CommentCollection](/javascript/api/excel/excel.commentcollection) |
| [大纲](../../excel/excel-add-ins-ranges-group.md) | 对行和列进行分组以形成可折叠的分级显示。 | [Range、Worksheet](/javascript/api/excel/excel.range) [](/javascript/api/excel/excel.worksheet) |
| [Slicers](../../excel/excel-add-ins-pivottables.md#filter-with-slicers) | 在表格和数据透视表中插入和配置切片器。 | [Slicer](/javascript/api/excel/excel.slicer) |
| [更多工作表事件](../../excel/excel-add-ins-events.md) | 侦听工作表中的单击和排序事件。 | [工作表 (事件) ](/javascript/api/excel/excel.worksheet#events) |

## <a name="api-list"></a>API 列表

下表列出了 Excel JavaScript API 要求集 1.10 中的 API。 若要查看 Excel JavaScript API 要求集 1.10 或更早版本支持的所有 API 的 API 参考文档，请参阅要求集 [1.10](/javascript/api/excel?view=excel-js-1.10&preserve-view=true)或更早版本中的 Excel API。

| 类 | 域 | 说明 |
|:---|:---|:---|
|[Comment](/javascript/api/excel/excel.comment)|[content](/javascript/api/excel/excel.comment#content)|注释的内容。|
||[delete()](/javascript/api/excel/excel.comment#delete--)|删除注释以及所有连接的回复。|
||[getLocation()](/javascript/api/excel/excel.comment#getlocation--)|获取此批注所在的单元格。|
||[authorEmail](/javascript/api/excel/excel.comment#authoremail)|获取批注作者的电子邮件。|
||[authorName](/javascript/api/excel/excel.comment#authorname)|获取批注作者的姓名。|
||[creationDate](/javascript/api/excel/excel.comment#creationdate)|获取批注的创建时间。|
||[id](/javascript/api/excel/excel.comment#id)|指定注释标识符。|
||[replies](/javascript/api/excel/excel.comment#replies)|表示与批注关联的回复对象的集合。|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[add (cellAddress： Range \| string， content： string， contentType？： Excel.ContentType) ](/javascript/api/excel/excel.commentcollection#add-celladdress--content--contenttype-)|使用给定单元格上的给定内容创建新批注。|
||[getCount()](/javascript/api/excel/excel.commentcollection#getcount--)|获取集合中的批注数量。|
||[getItem(commentId: string)](/javascript/api/excel/excel.commentcollection#getitem-commentid-)|根据其 ID 从集合中获取批注。|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentcollection#getitemat-index-)|根据其位置从集合中获取批注。|
||[getItemByCell(cellAddress: Range \| string)](/javascript/api/excel/excel.commentcollection#getitembycell-celladdress-)|从指定单元格获取的批注。|
||[getItemByReplyId(replyId: string)](/javascript/api/excel/excel.commentcollection#getitembyreplyid-replyid-)|获取给定答复所连接到的注释。|
||[items](/javascript/api/excel/excel.commentcollection#items)|获取此集合中已加载的子项。|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[content](/javascript/api/excel/excel.commentreply#content)|批注回复的内容。|
||[delete()](/javascript/api/excel/excel.commentreply#delete--)|删除批注回复。|
||[getLocation()](/javascript/api/excel/excel.commentreply#getlocation--)|获取此批注回复所在的单元格。|
||[getParentComment()](/javascript/api/excel/excel.commentreply#getparentcomment--)|获取此回复的父批注。|
||[authorEmail](/javascript/api/excel/excel.commentreply#authoremail)|获取批注回复作者的电子邮件。|
||[authorName](/javascript/api/excel/excel.commentreply#authorname)|获取批注回复作者的姓名。|
||[creationDate](/javascript/api/excel/excel.commentreply#creationdate)|获取批注回复的创建时间。|
||[id](/javascript/api/excel/excel.commentreply#id)|指定批注回复标识符。|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[add(content: string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentreplycollection#add-content--contenttype-)|为批注创建批注回复。|
||[getCount()](/javascript/api/excel/excel.commentreplycollection#getcount--)|获取集合中的批注回复数量。|
||[getItem(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#getitem-commentreplyid-)|返回由其 ID 标识的批注回复。|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentreplycollection#getitemat-index-)|根据其在集合中的位置获取批注回复。|
||[items](/javascript/api/excel/excel.commentreplycollection#items)|获取此集合中已加载的子项。|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[enableFieldList](/javascript/api/excel/excel.pivotlayout#enablefieldlist)|指定字段列表是否可在 UI 中显示。|
|[PivotTableStyle](/javascript/api/excel/excel.pivottablestyle)|[delete()](/javascript/api/excel/excel.pivottablestyle#delete--)|删除 PivotTableStyle。|
||[duplicate()](/javascript/api/excel/excel.pivottablestyle#duplicate--)|使用所有样式元素的副本创建此 PivotTableStyle 的副本。|
||[name](/javascript/api/excel/excel.pivottablestyle#name)|获取 PivotTableStyle 的名称。|
||[readOnly](/javascript/api/excel/excel.pivottablestyle#readonly)|指定此 PivotTableStyle 对象是否为只读。|
|[PivotTableStyleCollection](/javascript/api/excel/excel.pivottablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.pivottablestylecollection#add-name--makeuniquename-)|使用指定名称创建空白 PivotTableStyle。|
||[getCount()](/javascript/api/excel/excel.pivottablestylecollection#getcount--)|获取集合中 PivotTable 的数量。|
||[getDefault()](/javascript/api/excel/excel.pivottablestylecollection#getdefault--)|获取父对象范围的默认 PivotTableStyle。|
||[getItem(name: string)](/javascript/api/excel/excel.pivottablestylecollection#getitem-name-)|按名称获取 PivotTableStyle。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablestylecollection#getitemornullobject-name-)|按名称获取 PivotTableStyle。|
||[items](/javascript/api/excel/excel.pivottablestylecollection#items)|获取此集合中已加载的子项。|
||[setDefault(newDefaultStyle: PivotTableStyle \| string)](/javascript/api/excel/excel.pivottablestylecollection#setdefault-newdefaultstyle-)|设置在父对象范围内使用的默认 PivotTableStyle。|
|[Range](/javascript/api/excel/excel.range)|[group (groupOption：Excel.GroupOption) ](/javascript/api/excel/excel.range#group-groupoption-)|对大纲的列和行进行分组。|
||[hideGroupDetails (groupOption：Excel.GroupOption) ](/javascript/api/excel/excel.range#hidegroupdetails-groupoption-)|隐藏行或列组的详细信息。|
||[height](/javascript/api/excel/excel.range#height)|返回从区域的上边缘到区域的下边缘的 100％ 缩放的距离（以磅为单位）。|
||[left](/javascript/api/excel/excel.range#left)|返回从工作表的左边缘到区域的左边缘的 100％ 缩放的距离（以磅为单位）。|
||[top](/javascript/api/excel/excel.range#top)|返回从工作表的上边缘到区域的上边缘的 100％ 缩放的距离（以磅为单位）。|
||[width](/javascript/api/excel/excel.range#width)|返回从区域的左边缘到区域的右边缘的 100％ 缩放的距离（以磅为单位）。|
||[showGroupDetails (groupOption：Excel.GroupOption) ](/javascript/api/excel/excel.range#showgroupdetails-groupoption-)|显示行或列组的详细信息。|
||[ungroup (groupOption： Excel.GroupOption) ](/javascript/api/excel/excel.range#ungroup-groupoption-)|取消大纲的列和行的组合。|
|[Shape](/javascript/api/excel/excel.shape)|[copyTo(destinationSheet?: Worksheet \| string)](/javascript/api/excel/excel.shape#copyto-destinationsheet-)|复制并粘贴 Shape 对象。|
||[placement](/javascript/api/excel/excel.shape#placement)|表示对象如何附加到其下方的单元格。|
|[Slicer](/javascript/api/excel/excel.slicer)|[caption](/javascript/api/excel/excel.slicer#caption)|表示切片器的标题。|
||[clearFilters()](/javascript/api/excel/excel.slicer#clearfilters--)|清除当前切片器上应用的所有筛选器。|
||[delete()](/javascript/api/excel/excel.slicer#delete--)|删除切片器。|
||[getSelectedItems()](/javascript/api/excel/excel.slicer#getselecteditems--)|返回所选项目密钥的数组。|
||[height](/javascript/api/excel/excel.slicer#height)|表示切片器的高度（以磅为单位）。|
||[left](/javascript/api/excel/excel.slicer#left)|表示从切片器左侧到工作表左侧的距离（以磅为单位）。|
||[name](/javascript/api/excel/excel.slicer#name)|表示切片器的名称。|
||[id](/javascript/api/excel/excel.slicer#id)|表示切片器的唯一 ID。|
||[isFilterCleared](/javascript/api/excel/excel.slicer#isfiltercleared)|如果已清除当前切片器上应用的所有筛选器，则为 True。|
||[slicerItems](/javascript/api/excel/excel.slicer#sliceritems)|表示作为切片器一部分的 SlicerItems 的集合。|
||[worksheet](/javascript/api/excel/excel.slicer#worksheet)|表示包含切片器的工作表。|
||[selectItems(items?: string[])](/javascript/api/excel/excel.slicer#selectitems-items-)|根据切片器项的键选择切片器项。|
||[sortBy](/javascript/api/excel/excel.slicer#sortby)|表示切片器中的项目的排序顺序。|
||[style](/javascript/api/excel/excel.slicer#style)|表示切片器样式的常量值。|
||[top](/javascript/api/excel/excel.slicer#top)|表示从切片器上边缘到工作表顶部的距离（以磅为单位）。|
||[width](/javascript/api/excel/excel.slicer#width)|表示切片器的宽度（以磅为单位）。|
|[SlicerCollection](/javascript/api/excel/excel.slicercollection)|[add(slicerSource: string \| PivotTable \| Table, sourceField: string \| PivotField \| number \| TableColumn, slicerDestination?: string \| Worksheet)](/javascript/api/excel/excel.slicercollection#add-slicersource--sourcefield--slicerdestination-)|将新切片器添加到工作簿。|
||[getCount()](/javascript/api/excel/excel.slicercollection#getcount--)|返回集合中的切片器数量。|
||[getItem(key: string)](/javascript/api/excel/excel.slicercollection#getitem-key-)|使用其名称或 ID 获取 Slicer 对象。|
||[getItemAt(index: number)](/javascript/api/excel/excel.slicercollection#getitemat-index-)|根据其在集合中的位置获取切片器。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.slicercollection#getitemornullobject-key-)|使用其名称或 ID 获取切片器。|
||[items](/javascript/api/excel/excel.slicercollection#items)|获取此集合中已加载的子项。|
|[SlicerItem](/javascript/api/excel/excel.sliceritem)|[isSelected](/javascript/api/excel/excel.sliceritem#isselected)|如果选择了切片器项，则为 True。|
||[hasData](/javascript/api/excel/excel.sliceritem#hasdata)|如果切片器项包含数据，则为 True。|
||[key](/javascript/api/excel/excel.sliceritem#key)|表示代表切片器项的唯一值。|
||[name](/javascript/api/excel/excel.sliceritem#name)|表示 UI 中显示的标题。|
|[SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection)|[getCount()](/javascript/api/excel/excel.sliceritemcollection#getcount--)|返回切片器中的切片器项的数量。|
||[getItem(key: string)](/javascript/api/excel/excel.sliceritemcollection#getitem-key-)|使用其键或名称获取切片器项对象。|
||[getItemAt(index: number)](/javascript/api/excel/excel.sliceritemcollection#getitemat-index-)|根据其在集合中的位置获取切片器项。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.sliceritemcollection#getitemornullobject-key-)|使用其键或名称获取切片器项。|
||[items](/javascript/api/excel/excel.sliceritemcollection#items)|获取此集合中已加载的子项。|
|[SlicerStyle](/javascript/api/excel/excel.slicerstyle)|[delete()](/javascript/api/excel/excel.slicerstyle#delete--)|删除 SlicerStyle。|
||[duplicate()](/javascript/api/excel/excel.slicerstyle#duplicate--)|使用所有样式元素的副本创建此 SlicerStyle 的副本。|
||[name](/javascript/api/excel/excel.slicerstyle#name)|获取 SlicerStyle 的名称。|
||[readOnly](/javascript/api/excel/excel.slicerstyle#readonly)|指定此 SlicerStyle 对象是否为只读。|
|[SlicerStyleCollection](/javascript/api/excel/excel.slicerstylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.slicerstylecollection#add-name--makeuniquename-)|使用指定名称创建空白 SlicerStyle。|
||[getCount()](/javascript/api/excel/excel.slicerstylecollection#getcount--)|获取集合中的切片器样式数量。|
||[getDefault()](/javascript/api/excel/excel.slicerstylecollection#getdefault--)|获取父对象范围的默认 SlicerStyle。|
||[getItem(name: string)](/javascript/api/excel/excel.slicerstylecollection#getitem-name-)|按名称获取 SlicerStyle。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.slicerstylecollection#getitemornullobject-name-)|按名称获取 SlicerStyle。|
||[items](/javascript/api/excel/excel.slicerstylecollection#items)|获取此集合中已加载的子项。|
||[setDefault(newDefaultStyle: SlicerStyle \| string)](/javascript/api/excel/excel.slicerstylecollection#setdefault-newdefaultstyle-)|设置在父对象范围内使用的默认 SlicerStyle。|
|[TableStyle](/javascript/api/excel/excel.tablestyle)|[delete()](/javascript/api/excel/excel.tablestyle#delete--)|删除 TableStyle。|
||[duplicate()](/javascript/api/excel/excel.tablestyle#duplicate--)|使用所有样式元素的副本创建此 TableStyle 的副本。|
||[name](/javascript/api/excel/excel.tablestyle#name)|获取 TableStyle 的名称。|
||[readOnly](/javascript/api/excel/excel.tablestyle#readonly)|指定此 TableStyle 对象是否为只读。|
|[TableStyleCollection](/javascript/api/excel/excel.tablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.tablestylecollection#add-name--makeuniquename-)|使用指定名称创建空白 TableStyle。|
||[getCount()](/javascript/api/excel/excel.tablestylecollection#getcount--)|获取集合中表格样式的数量。|
||[getDefault()](/javascript/api/excel/excel.tablestylecollection#getdefault--)|获取父对象范围的默认 TableStyle。|
||[getItem(name: string)](/javascript/api/excel/excel.tablestylecollection#getitem-name-)|按名称获取 TableStyle。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.tablestylecollection#getitemornullobject-name-)|按名称获取 TableStyle。|
||[items](/javascript/api/excel/excel.tablestylecollection#items)|获取此集合中已加载的子项。|
||[setDefault(newDefaultStyle: TableStyle \| string)](/javascript/api/excel/excel.tablestylecollection#setdefault-newdefaultstyle-)|设置在父对象范围内使用的默认 TableStyle。|
|[TimelineStyle](/javascript/api/excel/excel.timelinestyle)|[delete()](/javascript/api/excel/excel.timelinestyle#delete--)|删除 TableStyle。|
||[duplicate()](/javascript/api/excel/excel.timelinestyle#duplicate--)|使用所有样式元素的副本创建此 TimelineStyle 的副本。|
||[name](/javascript/api/excel/excel.timelinestyle#name)|获取 TimelineStyle 的名称。|
||[readOnly](/javascript/api/excel/excel.timelinestyle#readonly)|指定此 TimelineStyle 对象是否为只读。|
|[TimelineStyleCollection](/javascript/api/excel/excel.timelinestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.timelinestylecollection#add-name--makeuniquename-)|使用指定名称创建空白 TimelineStyle。|
||[getCount()](/javascript/api/excel/excel.timelinestylecollection#getcount--)|获取集合中日程表样式的数量。|
||[getDefault()](/javascript/api/excel/excel.timelinestylecollection#getdefault--)|获取父对象范围的默认 TimelineStyle。|
||[getItem(name: string)](/javascript/api/excel/excel.timelinestylecollection#getitem-name-)|按名称获取 TimelineStyle。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.timelinestylecollection#getitemornullobject-name-)|按名称获取 TimelineStyle。|
||[items](/javascript/api/excel/excel.timelinestylecollection#items)|获取此集合中已加载的子项。|
||[setDefault(newDefaultStyle: TimelineStyle \| string)](/javascript/api/excel/excel.timelinestylecollection#setdefault-newdefaultstyle-)|设置在父对象范围内使用的默认 TimelineStyle。|
|[Workbook](/javascript/api/excel/excel.workbook)|[getActiveSlicer()](/javascript/api/excel/excel.workbook#getactiveslicer--)|获取工作簿中当前处于活动状态的切片器。|
||[getActiveSlicerOrNullObject()](/javascript/api/excel/excel.workbook#getactiveslicerornullobject--)|获取工作簿中当前处于活动状态的切片器。|
||[comments](/javascript/api/excel/excel.workbook#comments)|表示与工作簿关联的批注集合。|
||[pivotTableStyles](/javascript/api/excel/excel.workbook#pivottablestyles)|表示一组与工作簿相关联的 PivotTableStyles。|
||[slicerStyles](/javascript/api/excel/excel.workbook#slicerstyles)|表示一组与工作簿相关联的 SlicerStyles。|
||[slicers](/javascript/api/excel/excel.workbook#slicers)|表示与工作簿关联的切片器集合。|
||[tableStyles](/javascript/api/excel/excel.workbook#tablestyles)|表示一组与工作簿相关联的 TableStyles。|
||[timelineStyles](/javascript/api/excel/excel.workbook#timelinestyles)|表示一组与工作簿相关联的 TimelineStyles。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[comments](/javascript/api/excel/excel.worksheet#comments)|返回工作表上的所有 Comments 对象的集合。|
||[onColumnSorted](/javascript/api/excel/excel.worksheet#oncolumnsorted)|在已对一个或多个列进行排序时发生。|
||[onRowSorted](/javascript/api/excel/excel.worksheet#onrowsorted)|在已对一个或多个行进行排序时发生。|
||[onSingleClicked](/javascript/api/excel/excel.worksheet#onsingleclicked)|在工作表中发生左键单击/点击操作时发生。|
||[slicers](/javascript/api/excel/excel.worksheet#slicers)|返回属于工作表的切片器集合。|
||[showOutlineLevels (rowLevels： number， columnLevels： number) ](/javascript/api/excel/excel.worksheet#showoutlinelevels-rowlevels--columnlevels-)|按大纲级别显示行或列组。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onColumnSorted](/javascript/api/excel/excel.worksheetcollection#oncolumnsorted)|在已对一个或多个列进行排序时发生。|
||[onRowSorted](/javascript/api/excel/excel.worksheetcollection#onrowsorted)|在已对一个或多个行进行排序时发生。|
||[onSingleClicked](/javascript/api/excel/excel.worksheetcollection#onsingleclicked)|在工作表集合中发生左键单击/点击操作时发生。|
|[WorksheetColumnSortedEventArgs](/javascript/api/excel/excel.worksheetcolumnsortedeventargs)|[address](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#address)|获取区域地址，该地址表示特定工作表上的选定区域。|
||[source](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#source)|获取事件源。|
||[type](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#worksheetid)|获取发生排序的工作表的 id。|
|[WorksheetRowSortedEventArgs](/javascript/api/excel/excel.worksheetrowsortedeventargs)|[address](/javascript/api/excel/excel.worksheetrowsortedeventargs#address)|获取区域地址，该地址表示特定工作表上的选定区域。|
||[source](/javascript/api/excel/excel.worksheetrowsortedeventargs#source)|获取事件源。|
||[type](/javascript/api/excel/excel.worksheetrowsortedeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.worksheetrowsortedeventargs#worksheetid)|获取发生排序的工作表的 id。|
|[WorksheetSingleClickedEventArgs](/javascript/api/excel/excel.worksheetsingleclickedeventargs)|[address](/javascript/api/excel/excel.worksheetsingleclickedeventargs#address)|获取特定工作表中表示被左键单击/点击的单元格的地址。|
||[OffsetX](/javascript/api/excel/excel.worksheetsingleclickedeventargs#offsetx)|对于从右到左的语言，从左键单击/点击点到左 (或右到左) 单击/点击单元格的网格线边缘的距离（以点表示）。|
||[OffsetY](/javascript/api/excel/excel.worksheetsingleclickedeventargs#offsety)|从左键单击/点击的点到左键单击/点击的单元格的顶部网格线边缘的距离（以磅为单位）。|
||[type](/javascript/api/excel/excel.worksheetsingleclickedeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.worksheetsingleclickedeventargs#worksheetid)|获取已在其中左键单击/点击单元格的工作表的 ID。|

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 参考文档](/javascript/api/excel?view=excel-js-1.10&preserve-view=true)
- [Excel JavaScript API 要求集](excel-api-requirement-sets.md)