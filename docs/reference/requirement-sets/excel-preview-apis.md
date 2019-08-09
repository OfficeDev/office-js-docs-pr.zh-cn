---
title: Excel JavaScript 预览 API
description: 有关即将推出的 Excel JavaScript Api 的详细信息
ms.date: 08/06/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 4362a5e13e0031408236f34c718f0fcb3c4527e2
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268731"
---
# <a name="excel-javascript-preview-apis"></a>Excel JavaScript 预览 API

新的 Excel JavaScript API 首先在“预览版”中引入，在进行充分测试并获得用户反馈后，它将成为编号的特定要求集的一部分。

第一个表提供了 API 的简明摘要，而后续表提供了详细列表。

> [!NOTE]
> 预览 API 可能会发生变更，不适合在生产环境中使用。 我们建议你仅在测试和开发环境中试用它们。 不要在生产环境或业务关键型文档中使用预览 API。
>
> 若要使用预览 API，你必须引用 CDN 上的 **beta** 库（https://appsforoffice.microsoft.com/lib/beta/hosted/office.js)并且你可能还需要加入 Office 预览体验成员计划才能获得新的 Office 版本。

| 功能区域 | 说明 | 相关对象 |
|:--- |:--- |:--- |
| [Slicer](../../excel/excel-add-ins-pivottables.md#slicers-preview) | 在表格和数据透视表中插入和配置切片器。 | [Slicer](/javascript/api/excel/excel.slicer) |
| [备注](../../excel/excel-add-ins-workbooks.md#comments-preview) | 添加、编辑和删除备注。 | [Comment](/javascript/api/excel/excel.comment)、[CommentCollection](/javascript/api/excel/excel.commentcollection) |
| 工作簿[保存](../../excel/excel-add-ins-workbooks.md#save-the-workbook-preview)和[关闭](../../excel/excel-add-ins-workbooks.md#close-the-workbook-preview) | 保存和关闭工作簿。  | [Workbook](/javascript/api/excel/excel.workbook) |
| [插入工作簿](../../excel/excel-add-ins-workbooks.md#insert-a-copy-of-an-existing-workbook-into-the-current-one-preview) | 将一个工作簿插入另一个工作簿。  | [Workbook](/javascript/api/excel/excel.worksheetcollection) |

## <a name="api-list"></a>API 列表

下表列出了当前预览中的 Excel JavaScript Api。 若要查看所有 Excel JavaScript Api (包括预览 Api 和之前发布的 Api) 的完整列表, 请参阅[所有 Excel Javascript api](/javascript/api/excel?view=excel-js-preview)。

| Class | 域 | 说明 |
|:---|:---|:---|
|[Comment](/javascript/api/excel/excel.comment)|[content](/javascript/api/excel/excel.comment#content)|获取或设置批注的内容。 字符串为纯文本。|
||[delete()](/javascript/api/excel/excel.comment#delete--)|删除批注线程。|
||[getLocation()](/javascript/api/excel/excel.comment#getlocation--)|获取此注释所在的单元格。|
||[authorEmail](/javascript/api/excel/excel.comment#authoremail)|获取批注作者的电子邮件。|
||[authorName](/javascript/api/excel/excel.comment#authorname)|获取批注作者的姓名。|
||[creationDate](/javascript/api/excel/excel.comment#creationdate)|获取批注的创建时间。 如果批注是从备注转换而来的，则返回 null，因为批注没有创建日期。|
||[id](/javascript/api/excel/excel.comment#id)|表示批注标识符。 只读。|
||[replies](/javascript/api/excel/excel.comment#replies)|表示与批注关联的回复对象的集合。 只读。|
||[经过](/javascript/api/excel/excel.comment#resolved)|获取或设置批注线程的状态。 值为 "true" 表示注释线程处于 "已解决" 状态。|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[add(content: string, cellAddress: Range \| string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentcollection#add-content--celladdress--contenttype-)|在给定单元格上创建具有给定内容的新注释 (注释线程)。 如果`InvalidArgument`提供的范围大于一个单元格, 则会引发错误。|
||[getCount()](/javascript/api/excel/excel.commentcollection#getcount--)|获取集合中的批注数量。|
||[getItem(commentId: string)](/javascript/api/excel/excel.commentcollection#getitem-commentid-)|根据其 ID 从集合中获取批注。 只读。|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentcollection#getitemat-index-)|根据其位置从集合中获取批注。|
||[getItemByCell(cellAddress: Range \| string)](/javascript/api/excel/excel.commentcollection#getitembycell-celladdress-)|从指定单元格获取的批注。|
||[getItemByReplyId(replyId: string)](/javascript/api/excel/excel.commentcollection#getitembyreplyid-replyid-)|获取与集合中的回复 ID 相关的批注。|
||[items](/javascript/api/excel/excel.commentcollection#items)|获取此集合中已加载的子项。|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[content](/javascript/api/excel/excel.commentreply#content)|获取或设置批注回复的内容。 字符串为纯文本。|
||[delete()](/javascript/api/excel/excel.commentreply#delete--)|删除批注回复。|
||[getLocation()](/javascript/api/excel/excel.commentreply#getlocation--)|获取此批注答复所在的单元格。|
||[getParentComment()](/javascript/api/excel/excel.commentreply#getparentcomment--)|获取此回复的父注释。|
||[authorEmail](/javascript/api/excel/excel.commentreply#authoremail)|获取批注回复作者的电子邮件。|
||[authorName](/javascript/api/excel/excel.commentreply#authorname)|获取批注回复作者的姓名。|
||[creationDate](/javascript/api/excel/excel.commentreply#creationdate)|获取批注回复的创建时间。|
||[id](/javascript/api/excel/excel.commentreply#id)|表示批注回复标识符。 只读。|
||[经过](/javascript/api/excel/excel.commentreply#resolved)|获取或设置批注答复状态。 值为 "true" 表示批注答复处于 "已解决" 状态。|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[add(content: string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentreplycollection#add-content--contenttype-)|为批注创建批注回复。|
||[getCount()](/javascript/api/excel/excel.commentreplycollection#getcount--)|获取集合中的批注回复数量。|
||[getItem(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#getitem-commentreplyid-)|返回由其 ID 标识的批注回复。 只读。|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentreplycollection#getitemat-index-)|根据其在集合中的位置获取批注回复。|
||[items](/javascript/api/excel/excel.commentreplycollection#items)|获取此集合中已加载的子项。|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[enableFieldList](/javascript/api/excel/excel.pivotlayout#enablefieldlist)|指定是否可以 UI 中显示字段列表。|
||[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|根据数据层次结构以及各自层次结构的行和列项，获取数据透视表中的唯一单元格。 返回的单元格是给定行和列的交集，其中包含来自给定层次结构的数据。 此方法与在特定单元格上调用 getPivotItems 和 getDataHierarchy 相反。|
|[PivotTableStyle](/javascript/api/excel/excel.pivottablestyle)|[delete()](/javascript/api/excel/excel.pivottablestyle#delete--)|删除 PivotTableStyle。|
||[duplicate()](/javascript/api/excel/excel.pivottablestyle#duplicate--)|使用所有样式元素的副本创建此 PivotTableStyle 的副本。|
||[name](/javascript/api/excel/excel.pivottablestyle#name)|获取 PivotTableStyle 的名称。|
||[readOnly](/javascript/api/excel/excel.pivottablestyle#readonly)|指定此 PivotTableStyle 对象是否为只读。 只读。|
|[PivotTableStyleCollection](/javascript/api/excel/excel.pivottablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.pivottablestylecollection#add-name--makeuniquename-)|使用指定名称创建空白 PivotTableStyle。|
||[getCount()](/javascript/api/excel/excel.pivottablestylecollection#getcount--)|获取集合中 PivotTable 的数量。|
||[getDefault()](/javascript/api/excel/excel.pivottablestylecollection#getdefault--)|获取父对象范围的默认 PivotTableStyle。|
||[getItem(name: string)](/javascript/api/excel/excel.pivottablestylecollection#getitem-name-)|按名称获取 PivotTableStyle。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablestylecollection#getitemornullobject-name-)|按名称获取 PivotTableStyle。 如果没有 PivotTableStyle，将返回 null 对象。|
||[items](/javascript/api/excel/excel.pivottablestylecollection#items)|获取此集合中已加载的子项。|
||[setDefault(newDefaultStyle: PivotTableStyle \| string)](/javascript/api/excel/excel.pivottablestylecollection#setdefault-newdefaultstyle-)|设置在父对象范围内使用的默认 PivotTableStyle。|
|[Range](/javascript/api/excel/excel.range)|[getSpillParent()](/javascript/api/excel/excel.range#getspillparent--)|获取 Range 对象，它包含要将某个单元格溢出到的定位单元格。 如果应用于具有多个单元格的区域，则会失败。 只读。|
||[getSpillParentOrNullObject()](/javascript/api/excel/excel.range#getspillparentornullobject--)|获取 Range 对象，它包含要将某个单元格溢出到的定位单元格。 只读。|
||[getSpillingToRange()](/javascript/api/excel/excel.range#getspillingtorange--)|获取 Range 对象，它在调用定位单元格时包含溢出区域。 如果应用于具有多个单元格的区域，则会失败。 只读。|
||[getSpillingToRangeOrNullObject()](/javascript/api/excel/excel.range#getspillingtorangeornullobject--)|获取 Range 对象，它在调用定位单元格时包含溢出区域。 只读。|
||[group (groupOption: GroupOption)](/javascript/api/excel/excel.range#group-groupoption-)|对列和行进行分组以进行分级显示。|
||[hideGroupDetails (groupOption: GroupOption)](/javascript/api/excel/excel.range#hidegroupdetails-groupoption-)|隐藏行或列组的详细信息。|
||[hasSpill](/javascript/api/excel/excel.range#hasspill)|表示所有单元格是否都具有溢出边框。|
||[height](/javascript/api/excel/excel.range#height)|返回从区域的上边缘到区域的下边缘的 100％ 缩放的距离（以磅为单位）。 只读。|
||[left](/javascript/api/excel/excel.range#left)|返回从工作表的左边缘到区域的左边缘的 100％ 缩放的距离（以磅为单位）。 只读。|
||[savedAsArray](/javascript/api/excel/excel.range#savedasarray)|表示是否将所有单元格都保存为数组公式。|
||[top](/javascript/api/excel/excel.range#top)|返回从工作表的上边缘到区域的上边缘的 100％ 缩放的距离（以磅为单位）。 只读。|
||[width](/javascript/api/excel/excel.range#width)|返回从区域的左边缘到区域的右边缘的 100％ 缩放的距离（以磅为单位）。 只读。|
||[showGroupDetails (groupOption: GroupOption)](/javascript/api/excel/excel.range#showgroupdetails-groupoption-)|显示行或列组的详细信息。|
||[取消分组 (groupOption: GroupOption)](/javascript/api/excel/excel.range#ungroup-groupoption-)|取消边框的列和行的组合。|
|[Shape](/javascript/api/excel/excel.shape)|[copyTo(destinationSheet?: Worksheet \| string)](/javascript/api/excel/excel.shape#copyto-destinationsheet-)|复制并粘贴 Shape 对象。|
||[placement](/javascript/api/excel/excel.shape#placement)|表示对象如何附加到其下方的单元格。|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|从 XML 字符串创建可缩放的矢量图形 (SVG) 并将其添加到工作表。 返回表示新图片的 Shape 对象。|
|[Slicer](/javascript/api/excel/excel.slicer)|[caption](/javascript/api/excel/excel.slicer#caption)|表示切片器的标题。|
||[clearFilters()](/javascript/api/excel/excel.slicer#clearfilters--)|清除当前切片器上应用的所有筛选器。|
||[delete()](/javascript/api/excel/excel.slicer#delete--)|删除切片器。|
||[getSelectedItems()](/javascript/api/excel/excel.slicer#getselecteditems--)|返回所选项目密钥的数组。 只读。|
||[height](/javascript/api/excel/excel.slicer#height)|表示切片器的高度（以磅为单位）。|
||[left](/javascript/api/excel/excel.slicer#left)|表示从切片器左侧到工作表左侧的距离（以磅为单位）。|
||[name](/javascript/api/excel/excel.slicer#name)|表示切片器的名称。|
||[nameInFormula](/javascript/api/excel/excel.slicer#nameinformula)|表示公式中使用切片器名称。|
||[id](/javascript/api/excel/excel.slicer#id)|表示切片器的唯一 ID。 只读。|
||[isFilterCleared](/javascript/api/excel/excel.slicer#isfiltercleared)|如果已清除当前切片器上应用的所有筛选器，则为 True。|
||[slicerItems](/javascript/api/excel/excel.slicer#sliceritems)|表示作为切片器一部分的 SlicerItems 的集合。 只读。|
||[worksheet](/javascript/api/excel/excel.slicer#worksheet)|表示包含切片器的工作表。 只读。|
||[selectItems(items?: string[])](/javascript/api/excel/excel.slicer#selectitems-items-)|根据密钥选择切片器项目。 之前的选择将被清除。|
||[sortBy](/javascript/api/excel/excel.slicer#sortby)|表示切片器中的项目的排序顺序。 可能的值为：DataSourceOrder、Ascending、Descending。|
||[style](/javascript/api/excel/excel.slicer#style)|表示切片器样式的常量值。 可能的值为: "SlicerStyleLight1" 到 "SlicerStyleLight6"、"TableStyleOther1" 通过 "TableStyleOther2"、"SlicerStyleDark1" 到 "SlicerStyleDark6"。 还可以指定工作簿中显示的用户定义的自定义样式。|
||[top](/javascript/api/excel/excel.slicer#top)|表示从切片器上边缘到工作表顶部的距离（以磅为单位）。|
||[width](/javascript/api/excel/excel.slicer#width)|表示切片器的宽度（以磅为单位）。|
|[SlicerCollection](/javascript/api/excel/excel.slicercollection)|[add(slicerSource: string \| PivotTable \| Table, sourceField: string \| PivotField \| number \| TableColumn, slicerDestination?: string \| Worksheet)](/javascript/api/excel/excel.slicercollection#add-slicersource--sourcefield--slicerdestination-)|将新切片器添加到工作簿。|
||[getCount()](/javascript/api/excel/excel.slicercollection#getcount--)|返回集合中的切片器数量。|
||[getItem(key: string)](/javascript/api/excel/excel.slicercollection#getitem-key-)|使用其名称或 ID 获取 Slicer 对象。|
||[getItemAt(index: number)](/javascript/api/excel/excel.slicercollection#getitemat-index-)|根据其在集合中的位置获取切片器。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.slicercollection#getitemornullobject-key-)|使用其名称或 ID 获取切片器。如果没有切片器项，将返回 null 对象。|
||[items](/javascript/api/excel/excel.slicercollection#items)|获取此集合中已加载的子项。|
|[SlicerItem](/javascript/api/excel/excel.sliceritem)|[isSelected](/javascript/api/excel/excel.sliceritem#isselected)|如果选择了切片器项，则为 True。|
||[hasData](/javascript/api/excel/excel.sliceritem#hasdata)|如果切片器项包含数据，则为 True。|
||[key](/javascript/api/excel/excel.sliceritem#key)|表示代表切片器项的唯一值。|
||[name](/javascript/api/excel/excel.sliceritem#name)|代表 UI 中显示的标题。|
|[SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection)|[getCount()](/javascript/api/excel/excel.sliceritemcollection#getcount--)|返回切片器中的切片器项的数量。|
||[getItem(key: string)](/javascript/api/excel/excel.sliceritemcollection#getitem-key-)|使用其键或名称获取切片器项对象。|
||[getItemAt(index: number)](/javascript/api/excel/excel.sliceritemcollection#getitemat-index-)|根据其在集合中的位置获取切片器项。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.sliceritemcollection#getitemornullobject-key-)|使用其键或名称获取切片器项。 如果没有切片器项，将返回 null 对象。|
||[items](/javascript/api/excel/excel.sliceritemcollection#items)|获取此集合中已加载的子项。|
|[SlicerStyle](/javascript/api/excel/excel.slicerstyle)|[delete()](/javascript/api/excel/excel.slicerstyle#delete--)|删除 SlicerStyle。|
||[duplicate()](/javascript/api/excel/excel.slicerstyle#duplicate--)|使用所有样式元素的副本创建此 SlicerStyle 的副本。|
||[name](/javascript/api/excel/excel.slicerstyle#name)|获取 SlicerStyle 的名称。|
||[readOnly](/javascript/api/excel/excel.slicerstyle#readonly)|指定此 SlicerStyle 对象是否为只读。 只读。|
|[SlicerStyleCollection](/javascript/api/excel/excel.slicerstylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.slicerstylecollection#add-name--makeuniquename-)|使用指定名称创建空白 SlicerStyle。|
||[getCount()](/javascript/api/excel/excel.slicerstylecollection#getcount--)|获取集合中的切片器样式数量。|
||[getDefault()](/javascript/api/excel/excel.slicerstylecollection#getdefault--)|获取父对象范围的默认 SlicerStyle。|
||[getItem(name: string)](/javascript/api/excel/excel.slicerstylecollection#getitem-name-)|按名称获取 SlicerStyle。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.slicerstylecollection#getitemornullobject-name-)|按名称获取 SlicerStyle。 如果没有 SlicerStyle，将返回 null 对象。|
||[items](/javascript/api/excel/excel.slicerstylecollection#items)|获取此集合中已加载的子项。|
||[setDefault(newDefaultStyle: SlicerStyle \| string)](/javascript/api/excel/excel.slicerstylecollection#setdefault-newdefaultstyle-)|设置在父对象范围内使用的默认 SlicerStyle。|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|将表格更改为使用默认表格样式。|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|在特定表格上应用筛选器时发生。|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|在工作簿或工作表中的任何表格上应用筛选器时发生。|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|表示应用筛选器的表格的 ID。|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#type)|表示事件的类型。 有关详细信息，请参阅 Excel.EventType。|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|表示包含表格的工作表的 ID。|
|[TableStyle](/javascript/api/excel/excel.tablestyle)|[delete()](/javascript/api/excel/excel.tablestyle#delete--)|删除 TableStyle。|
||[duplicate()](/javascript/api/excel/excel.tablestyle#duplicate--)|使用所有样式元素的副本创建此 TableStyle 的副本。|
||[name](/javascript/api/excel/excel.tablestyle#name)|获取 TableStyle 的名称。|
||[readOnly](/javascript/api/excel/excel.tablestyle#readonly)|指定此 TableStyle 对象是否为只读。 只读。|
|[TableStyleCollection](/javascript/api/excel/excel.tablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.tablestylecollection#add-name--makeuniquename-)|使用指定名称创建空白 TableStyle。|
||[getCount()](/javascript/api/excel/excel.tablestylecollection#getcount--)|获取集合中表格样式的数量。|
||[getDefault()](/javascript/api/excel/excel.tablestylecollection#getdefault--)|获取父对象范围的默认 TableStyle。|
||[getItem(name: string)](/javascript/api/excel/excel.tablestylecollection#getitem-name-)|按名称获取 TableStyle。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.tablestylecollection#getitemornullobject-name-)|按名称获取 TableStyle。 如果没有 TableStyle，将返回 null 对象。|
||[items](/javascript/api/excel/excel.tablestylecollection#items)|获取此集合中已加载的子项。|
||[setDefault(newDefaultStyle: TableStyle \| string)](/javascript/api/excel/excel.tablestylecollection#setdefault-newdefaultstyle-)|设置在父对象范围内使用的默认 TableStyle。|
|[TimelineStyle](/javascript/api/excel/excel.timelinestyle)|[delete()](/javascript/api/excel/excel.timelinestyle#delete--)|删除 TableStyle。|
||[duplicate()](/javascript/api/excel/excel.timelinestyle#duplicate--)|使用所有样式元素的副本创建此 TimelineStyle 的副本。|
||[name](/javascript/api/excel/excel.timelinestyle#name)|获取 TimelineStyle 的名称。|
||[readOnly](/javascript/api/excel/excel.timelinestyle#readonly)|指定此 TimelineStyle 对象是否为只读。 只读。|
|[TimelineStyleCollection](/javascript/api/excel/excel.timelinestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.timelinestylecollection#add-name--makeuniquename-)|使用指定名称创建空白 TimelineStyle。|
||[getCount()](/javascript/api/excel/excel.timelinestylecollection#getcount--)|获取集合中日程表样式的数量。|
||[getDefault()](/javascript/api/excel/excel.timelinestylecollection#getdefault--)|获取父对象范围的默认 TimelineStyle。|
||[getItem(name: string)](/javascript/api/excel/excel.timelinestylecollection#getitem-name-)|按名称获取 TimelineStyle。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.timelinestylecollection#getitemornullobject-name-)|按名称获取 TimelineStyle。 如果没有 TimelineStyle，将返回 null 对象。|
||[items](/javascript/api/excel/excel.timelinestylecollection#items)|获取此集合中已加载的子项。|
||[setDefault(newDefaultStyle: TimelineStyle \| string)](/javascript/api/excel/excel.timelinestylecollection#setdefault-newdefaultstyle-)|设置在父对象范围内使用的默认 TimelineStyle。|
|[Workbook](/javascript/api/excel/excel.workbook)|[close(closeBehavior?: Excel.CloseBehavior)](/javascript/api/excel/excel.workbook#close-closebehavior-)|关闭当前工作簿。|
||[getActiveSlicer()](/javascript/api/excel/excel.workbook#getactiveslicer--)|获取工作簿中当前处于活动状态的切片器。 如果没有活动切片器, 则会`ItemNotFound`引发异常。|
||[getActiveSlicerOrNullObject()](/javascript/api/excel/excel.workbook#getactiveslicerornullobject--)|获取工作簿中当前处于活动状态的切片器。 如果没有处于活动状态的切片器，则返回 null 对象。|
||[comments](/javascript/api/excel/excel.workbook#comments)|表示与工作簿关联的批注集合。 只读。|
||[pivotTableStyles](/javascript/api/excel/excel.workbook#pivottablestyles)|表示一组与工作簿相关联的 PivotTableStyles。 只读。|
||[slicerStyles](/javascript/api/excel/excel.workbook#slicerstyles)|表示一组与工作簿相关联的 SlicerStyles。 只读。|
||[slicers](/javascript/api/excel/excel.workbook#slicers)|表示与工作簿关联的切片器集合。 只读。|
||[tableStyles](/javascript/api/excel/excel.workbook#tablestyles)|表示一组与工作簿相关联的 TableStyles。 只读。|
||[timelineStyles](/javascript/api/excel/excel.workbook#timelinestyles)|表示一组与工作簿相关联的 TimelineStyles。 只读。|
||[save(saveBehavior?: Excel.SaveBehavior)](/javascript/api/excel/excel.workbook#save-savebehavior-)|保存当前工作簿。|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|如果工作簿使用 1904 日期系统，则为 True。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[comments](/javascript/api/excel/excel.worksheet#comments)|返回工作表上的所有 Comments 对象的集合。 只读。|
||[onColumnSorted](/javascript/api/excel/excel.worksheet#oncolumnsorted)|在列上排序时发生。|
||[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|在特定工作表上应用筛选器时发生。|
||[onRowHiddenChanged](/javascript/api/excel/excel.worksheet#onrowhiddenchanged)|当特定工作表上的行隐藏状态更改时发生。|
||[onRowSorted](/javascript/api/excel/excel.worksheet#onrowsorted)|在行上排序时发生。|
||[onSingleClicked](/javascript/api/excel/excel.worksheet#onsingleclicked)|在工作表中进行左键单击/点击操作时发生。 在以下情况下单击时, 将不会触发此事件: [...]|
||[slicers](/javascript/api/excel/excel.worksheet#slicers)|返回作为工作表一部分的切片器集合。 只读。|
||[showOutlineLevels (rowLevels: 数字, columnLevels: 数字)](/javascript/api/excel/excel.worksheet#showoutlinelevels-rowlevels--columnlevels-)|按行或列的大纲级别显示组。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|将工作簿的指定工作表插入当前工作簿。|
||[onColumnSorted](/javascript/api/excel/excel.worksheetcollection#oncolumnsorted)|在列上排序时发生。|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|在工作簿中应用任何工作表的筛选器时发生。|
||[onRowHiddenChanged](/javascript/api/excel/excel.worksheetcollection#onrowhiddenchanged)|当工作簿中的任何工作表的行隐藏状态更改时发生。|
||[onRowSorted](/javascript/api/excel/excel.worksheetcollection#onrowsorted)|在行上排序时发生。|
||[onSingleClicked](/javascript/api/excel/excel.worksheetcollection#onsingleclicked)|在工作表集合中发生左击或螺纹操作时发生。 在以下情况下单击时, 将不会触发此事件: [...]|
|[WorksheetColumnSortedEventArgs](/javascript/api/excel/excel.worksheetcolumnsortedeventargs)|[address](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#address)|获取区域地址，该地址表示特定工作表上的选定区域。|
||[source](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#source)|获取事件源。 有关详细信息，请参阅 Excel.EventSource。|
||[type](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#type)|获取事件的类型。 有关详细信息，请参阅 Excel.EventType。|
||[worksheetId](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#worksheetid)|获取发生排序的工作表的 id。|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|表示事件的类型。 有关详细信息，请参阅 Excel.EventType。|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|表示在其中应用筛选器的工作表的 ID。|
|[WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs)|[address](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#address)|获取区域地址，该地址表示特定工作表上的更改区域。|
||[changeType](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#changetype)|获取表示如何触发 RowHiddenChanged 事件的更改类型。 有关详细信息, 请参阅 RowHiddenChangeType。|
||[source](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#source)|获取事件源。 有关详细信息，请参阅 Excel.EventSource。|
||[type](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#type)|获取事件的类型。 有关详细信息，请参阅 Excel.EventType。|
||[worksheetId](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#worksheetid)|获取其中的数据发生更改的工作表的 ID。|
|[WorksheetRowSortedEventArgs](/javascript/api/excel/excel.worksheetrowsortedeventargs)|[address](/javascript/api/excel/excel.worksheetrowsortedeventargs#address)|获取区域地址，该地址表示特定工作表上的选定区域。|
||[source](/javascript/api/excel/excel.worksheetrowsortedeventargs#source)|获取事件源。 有关详细信息，请参阅 Excel.EventSource。|
||[type](/javascript/api/excel/excel.worksheetrowsortedeventargs#type)|获取事件的类型。 有关详细信息，请参阅 Excel.EventType。|
||[worksheetId](/javascript/api/excel/excel.worksheetrowsortedeventargs#worksheetid)|获取发生排序的工作表的 id。|
|[WorksheetSingleClickedEventArgs](/javascript/api/excel/excel.worksheetsingleclickedeventargs)|[address](/javascript/api/excel/excel.worksheetsingleclickedeventargs#address)|获取特定工作表中表示被左键单击/点击的单元格的地址。|
||[OffsetX](/javascript/api/excel/excel.worksheetsingleclickedeventargs#offsetx)|从左击/点击的点到左侧 (或从右到左语言的右侧) 的网格线边缘的距离, 以磅为单位的左击的单元格的网格线边缘。|
||[OffsetY](/javascript/api/excel/excel.worksheetsingleclickedeventargs#offsety)|从左键单击/点击的点到左键单击/点击的单元格的顶部网格线边缘的距离（以磅为单位）。|
||[type](/javascript/api/excel/excel.worksheetsingleclickedeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.worksheetsingleclickedeventargs#worksheetid)|获取已在其中左键单击/点击单元格的工作表的 ID。|

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 参考文档](/javascript/api/excel?view=excel-js-preview)
- [Excel JavaScript API 要求集](./excel-api-requirement-sets.md)
