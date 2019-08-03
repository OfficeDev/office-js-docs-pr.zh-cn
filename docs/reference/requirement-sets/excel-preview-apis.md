---
title: Excel JavaScript 预览 API
description: 有关即将推出的 Excel JavaScript Api 的详细信息
ms.date: 07/25/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 200b187b059c1b03ae3713b5afa11b2152aba0da
ms.sourcegitcommit: 3f5d7f4794e3d3c8bc3a79fa05c54157613b9376
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/02/2019
ms.locfileid: "36064849"
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
||[set (properties: Excel. 注释)](/javascript/api/excel/excel.comment#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: CommentUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.comment#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[add(content: string, cellAddress: Range \| string, contentType?: "Plain")](/javascript/api/excel/excel.commentcollection#add-content--celladdress--contenttype-)|在给定单元格上创建具有给定内容的新注释 (注释线程)。 如果`InvalidArgument`提供的范围大于一个单元格, 则会引发错误。|
||[add(content: string, cellAddress: Range \| string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentcollection#add-content--celladdress--contenttype-)|在给定单元格上创建具有给定内容的新注释 (注释线程)。 如果`InvalidArgument`提供的范围大于一个单元格, 则会引发错误。|
||[getCount()](/javascript/api/excel/excel.commentcollection#getcount--)|获取集合中的批注数量。|
||[getItem(commentId: string)](/javascript/api/excel/excel.commentcollection#getitem-commentid-)|根据其 ID 从集合中获取批注。 只读。|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentcollection#getitemat-index-)|根据其位置从集合中获取批注。|
||[getItemByCell(cellAddress: Range \| string)](/javascript/api/excel/excel.commentcollection#getitembycell-celladdress-)|从指定单元格获取的批注。|
||[getItemByReplyId(replyId: string)](/javascript/api/excel/excel.commentcollection#getitembyreplyid-replyid-)|获取与集合中的回复 ID 相关的批注。|
||[items](/javascript/api/excel/excel.commentcollection#items)|获取此集合中已加载的子项。|
|[CommentCollectionData](/javascript/api/excel/excel.commentcollectiondata)|[items](/javascript/api/excel/excel.commentcollectiondata#items)||
|[CommentCollectionLoadOptions](/javascript/api/excel/excel.commentcollectionloadoptions)|[$all](/javascript/api/excel/excel.commentcollectionloadoptions#$all)||
||[authorEmail](/javascript/api/excel/excel.commentcollectionloadoptions#authoremail)|对于集合中的每一项: 获取批注作者的电子邮件。|
||[authorName](/javascript/api/excel/excel.commentcollectionloadoptions#authorname)|对于集合中的每一项: 获取批注作者的姓名。|
||[content](/javascript/api/excel/excel.commentcollectionloadoptions#content)|对于集合中的每一项: 获取或设置批注的内容。 字符串为纯文本。|
||[creationDate](/javascript/api/excel/excel.commentcollectionloadoptions#creationdate)|对于集合中的每一项: 获取注释的创建时间。 如果批注是从备注转换而来的，则返回 null，因为批注没有创建日期。|
||[id](/javascript/api/excel/excel.commentcollectionloadoptions#id)|对于集合中的每一项: 代表注释标识符。 只读。|
|[CommentCollectionUpdateData](/javascript/api/excel/excel.commentcollectionupdatedata)|[items](/javascript/api/excel/excel.commentcollectionupdatedata#items)||
|[CommentData](/javascript/api/excel/excel.commentdata)|[authorEmail](/javascript/api/excel/excel.commentdata#authoremail)|获取批注作者的电子邮件。|
||[authorName](/javascript/api/excel/excel.commentdata#authorname)|获取批注作者的姓名。|
||[content](/javascript/api/excel/excel.commentdata#content)|获取或设置批注的内容。 字符串为纯文本。|
||[creationDate](/javascript/api/excel/excel.commentdata#creationdate)|获取批注的创建时间。 如果批注是从备注转换而来的，则返回 null，因为批注没有创建日期。|
||[id](/javascript/api/excel/excel.commentdata#id)|表示批注标识符。 只读。|
||[replies](/javascript/api/excel/excel.commentdata#replies)|表示与批注关联的回复对象的集合。 只读。|
|[CommentLoadOptions](/javascript/api/excel/excel.commentloadoptions)|[$all](/javascript/api/excel/excel.commentloadoptions#$all)||
||[authorEmail](/javascript/api/excel/excel.commentloadoptions#authoremail)|获取批注作者的电子邮件。|
||[authorName](/javascript/api/excel/excel.commentloadoptions#authorname)|获取批注作者的姓名。|
||[content](/javascript/api/excel/excel.commentloadoptions#content)|获取或设置批注的内容。 字符串为纯文本。|
||[creationDate](/javascript/api/excel/excel.commentloadoptions#creationdate)|获取批注的创建时间。 如果批注是从备注转换而来的，则返回 null，因为批注没有创建日期。|
||[id](/javascript/api/excel/excel.commentloadoptions#id)|表示批注标识符。 只读。|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[content](/javascript/api/excel/excel.commentreply#content)|获取或设置批注回复的内容。 字符串为纯文本。|
||[delete()](/javascript/api/excel/excel.commentreply#delete--)|删除批注回复。|
||[getLocation()](/javascript/api/excel/excel.commentreply#getlocation--)|获取此批注答复所在的单元格。|
||[getParentComment()](/javascript/api/excel/excel.commentreply#getparentcomment--)|获取此回复的父注释。|
||[authorEmail](/javascript/api/excel/excel.commentreply#authoremail)|获取批注回复作者的电子邮件。|
||[authorName](/javascript/api/excel/excel.commentreply#authorname)|获取批注回复作者的姓名。|
||[creationDate](/javascript/api/excel/excel.commentreply#creationdate)|获取批注回复的创建时间。|
||[id](/javascript/api/excel/excel.commentreply#id)|表示批注回复标识符。 只读。|
||[set (properties: CommentReply)](/javascript/api/excel/excel.commentreply#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: CommentReplyUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.commentreply#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[add(content: string, contentType?: "Plain")](/javascript/api/excel/excel.commentreplycollection#add-content--contenttype-)|为批注创建批注回复。|
||[add(content: string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentreplycollection#add-content--contenttype-)|为批注创建批注回复。|
||[getCount()](/javascript/api/excel/excel.commentreplycollection#getcount--)|获取集合中的批注回复数量。|
||[getItem(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#getitem-commentreplyid-)|返回由其 ID 标识的批注回复。 只读。|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentreplycollection#getitemat-index-)|根据其在集合中的位置获取批注回复。|
||[items](/javascript/api/excel/excel.commentreplycollection#items)|获取此集合中已加载的子项。|
|[CommentReplyCollectionData](/javascript/api/excel/excel.commentreplycollectiondata)|[items](/javascript/api/excel/excel.commentreplycollectiondata#items)||
|[CommentReplyCollectionLoadOptions](/javascript/api/excel/excel.commentreplycollectionloadoptions)|[$all](/javascript/api/excel/excel.commentreplycollectionloadoptions#$all)||
||[authorEmail](/javascript/api/excel/excel.commentreplycollectionloadoptions#authoremail)|对于集合中的每一项: 获取批注答复作者的电子邮件。|
||[authorName](/javascript/api/excel/excel.commentreplycollectionloadoptions#authorname)|对于集合中的每一项: 获取批注答复作者的姓名。|
||[content](/javascript/api/excel/excel.commentreplycollectionloadoptions#content)|对于集合中的每一项: 获取或设置批注答复的内容。 字符串为纯文本。|
||[creationDate](/javascript/api/excel/excel.commentreplycollectionloadoptions#creationdate)|对于集合中的每一项: 获取批注答复的创建时间。|
||[id](/javascript/api/excel/excel.commentreplycollectionloadoptions#id)|对于集合中的每一项: 表示批注答复标识符。 只读。|
|[CommentReplyCollectionUpdateData](/javascript/api/excel/excel.commentreplycollectionupdatedata)|[items](/javascript/api/excel/excel.commentreplycollectionupdatedata#items)||
|[CommentReplyData](/javascript/api/excel/excel.commentreplydata)|[authorEmail](/javascript/api/excel/excel.commentreplydata#authoremail)|获取批注回复作者的电子邮件。|
||[authorName](/javascript/api/excel/excel.commentreplydata#authorname)|获取批注回复作者的姓名。|
||[content](/javascript/api/excel/excel.commentreplydata#content)|获取或设置批注回复的内容。 字符串为纯文本。|
||[creationDate](/javascript/api/excel/excel.commentreplydata#creationdate)|获取批注回复的创建时间。|
||[id](/javascript/api/excel/excel.commentreplydata#id)|表示批注回复标识符。 只读。|
|[CommentReplyLoadOptions](/javascript/api/excel/excel.commentreplyloadoptions)|[$all](/javascript/api/excel/excel.commentreplyloadoptions#$all)||
||[authorEmail](/javascript/api/excel/excel.commentreplyloadoptions#authoremail)|获取批注回复作者的电子邮件。|
||[authorName](/javascript/api/excel/excel.commentreplyloadoptions#authorname)|获取批注回复作者的姓名。|
||[content](/javascript/api/excel/excel.commentreplyloadoptions#content)|获取或设置批注回复的内容。 字符串为纯文本。|
||[creationDate](/javascript/api/excel/excel.commentreplyloadoptions#creationdate)|获取批注回复的创建时间。|
||[id](/javascript/api/excel/excel.commentreplyloadoptions#id)|表示批注回复标识符。 只读。|
|[CommentReplyUpdateData](/javascript/api/excel/excel.commentreplyupdatedata)|[content](/javascript/api/excel/excel.commentreplyupdatedata#content)|获取或设置批注回复的内容。 字符串为纯文本。|
|[CommentUpdateData](/javascript/api/excel/excel.commentupdatedata)|[content](/javascript/api/excel/excel.commentupdatedata#content)|获取或设置批注的内容。 字符串为纯文本。|
|[GroupShapeCollectionLoadOptions](/javascript/api/excel/excel.groupshapecollectionloadoptions)|[placement](/javascript/api/excel/excel.groupshapecollectionloadoptions#placement)|对于集合中的每个项目: 代表如何将对象附加到其下的单元格。|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[enableFieldList](/javascript/api/excel/excel.pivotlayout#enablefieldlist)|指定是否可以 UI 中显示字段列表。|
||[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|根据数据层次结构以及各自层次结构的行和列项，获取数据透视表中的唯一单元格。 返回的单元格是给定行和列的交集，其中包含来自给定层次结构的数据。 此方法与在特定单元格上调用 getPivotItems 和 getDataHierarchy 相反。|
|[PivotLayoutData](/javascript/api/excel/excel.pivotlayoutdata)|[enableFieldList](/javascript/api/excel/excel.pivotlayoutdata#enablefieldlist)|指定是否可以 UI 中显示字段列表。|
|[PivotLayoutLoadOptions](/javascript/api/excel/excel.pivotlayoutloadoptions)|[enableFieldList](/javascript/api/excel/excel.pivotlayoutloadoptions#enablefieldlist)|指定是否可以 UI 中显示字段列表。|
|[PivotLayoutUpdateData](/javascript/api/excel/excel.pivotlayoutupdatedata)|[enableFieldList](/javascript/api/excel/excel.pivotlayoutupdatedata#enablefieldlist)|指定是否可以 UI 中显示字段列表。|
|[PivotTableStyle](/javascript/api/excel/excel.pivottablestyle)|[delete()](/javascript/api/excel/excel.pivottablestyle#delete--)|删除 PivotTableStyle。|
||[duplicate()](/javascript/api/excel/excel.pivottablestyle#duplicate--)|使用所有样式元素的副本创建此 PivotTableStyle 的副本。|
||[name](/javascript/api/excel/excel.pivottablestyle#name)|获取 PivotTableStyle 的名称。|
||[readOnly](/javascript/api/excel/excel.pivottablestyle#readonly)|指定此 PivotTableStyle 对象是否为只读。 只读。|
||[set (properties: PivotTableStyle)](/javascript/api/excel/excel.pivottablestyle#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: PivotTableStyleUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.pivottablestyle#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
|[PivotTableStyleCollection](/javascript/api/excel/excel.pivottablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.pivottablestylecollection#add-name--makeuniquename-)|使用指定名称创建空白 PivotTableStyle。|
||[getCount()](/javascript/api/excel/excel.pivottablestylecollection#getcount--)|获取集合中 PivotTable 的数量。|
||[getDefault()](/javascript/api/excel/excel.pivottablestylecollection#getdefault--)|获取父对象范围的默认 PivotTableStyle。|
||[getItem(name: string)](/javascript/api/excel/excel.pivottablestylecollection#getitem-name-)|按名称获取 PivotTableStyle。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablestylecollection#getitemornullobject-name-)|按名称获取 PivotTableStyle。 如果没有 PivotTableStyle，将返回 null 对象。|
||[items](/javascript/api/excel/excel.pivottablestylecollection#items)|获取此集合中已加载的子项。|
||[setDefault(newDefaultStyle: PivotTableStyle \| string)](/javascript/api/excel/excel.pivottablestylecollection#setdefault-newdefaultstyle-)|设置在父对象范围内使用的默认 PivotTableStyle。|
|[PivotTableStyleCollectionData](/javascript/api/excel/excel.pivottablestylecollectiondata)|[items](/javascript/api/excel/excel.pivottablestylecollectiondata#items)||
|[PivotTableStyleCollectionLoadOptions](/javascript/api/excel/excel.pivottablestylecollectionloadoptions)|[$all](/javascript/api/excel/excel.pivottablestylecollectionloadoptions#$all)||
||[name](/javascript/api/excel/excel.pivottablestylecollectionloadoptions#name)|对于集合中的每一项: 获取 PivotTableStyle 的名称。|
||[readOnly](/javascript/api/excel/excel.pivottablestylecollectionloadoptions#readonly)|对于集合中的每一项: 指定此 PivotTableStyle 对象是否为只读。 只读。|
|[PivotTableStyleCollectionUpdateData](/javascript/api/excel/excel.pivottablestylecollectionupdatedata)|[items](/javascript/api/excel/excel.pivottablestylecollectionupdatedata#items)||
|[PivotTableStyleData](/javascript/api/excel/excel.pivottablestyledata)|[name](/javascript/api/excel/excel.pivottablestyledata#name)|获取 PivotTableStyle 的名称。|
||[readOnly](/javascript/api/excel/excel.pivottablestyledata#readonly)|指定此 PivotTableStyle 对象是否为只读。 只读。|
|[PivotTableStyleLoadOptions](/javascript/api/excel/excel.pivottablestyleloadoptions)|[$all](/javascript/api/excel/excel.pivottablestyleloadoptions#$all)||
||[name](/javascript/api/excel/excel.pivottablestyleloadoptions#name)|获取 PivotTableStyle 的名称。|
||[readOnly](/javascript/api/excel/excel.pivottablestyleloadoptions#readonly)|指定此 PivotTableStyle 对象是否为只读。 只读。|
|[PivotTableStyleUpdateData](/javascript/api/excel/excel.pivottablestyleupdatedata)|[name](/javascript/api/excel/excel.pivottablestyleupdatedata#name)|获取 PivotTableStyle 的名称。|
|[Range](/javascript/api/excel/excel.range)|[getSpillParent()](/javascript/api/excel/excel.range#getspillparent--)|获取 Range 对象，它包含要将某个单元格溢出到的定位单元格。 如果应用于具有多个单元格的区域，则会失败。 只读。|
||[getSpillParentOrNullObject()](/javascript/api/excel/excel.range#getspillparentornullobject--)|获取 Range 对象，它包含要将某个单元格溢出到的定位单元格。 只读。|
||[getSpillingToRange()](/javascript/api/excel/excel.range#getspillingtorange--)|获取 Range 对象，它在调用定位单元格时包含溢出区域。 如果应用于具有多个单元格的区域，则会失败。 只读。|
||[getSpillingToRangeOrNullObject()](/javascript/api/excel/excel.range#getspillingtorangeornullobject--)|获取 Range 对象，它在调用定位单元格时包含溢出区域。 只读。|
||[hasSpill](/javascript/api/excel/excel.range#hasspill)|表示所有单元格是否都具有溢出边框。|
||[height](/javascript/api/excel/excel.range#height)|返回从区域的上边缘到区域的下边缘的 100％ 缩放的距离（以磅为单位）。 只读。|
||[left](/javascript/api/excel/excel.range#left)|返回从工作表的左边缘到区域的左边缘的 100％ 缩放的距离（以磅为单位）。 只读。|
||[savedAsArray](/javascript/api/excel/excel.range#savedasarray)|表示是否将所有单元格都保存为数组公式。|
||[top](/javascript/api/excel/excel.range#top)|返回从工作表的上边缘到区域的上边缘的 100％ 缩放的距离（以磅为单位）。 只读。|
||[width](/javascript/api/excel/excel.range#width)|返回从区域的左边缘到区域的右边缘的 100％ 缩放的距离（以磅为单位）。 只读。|
|[RangeCollectionLoadOptions](/javascript/api/excel/excel.rangecollectionloadoptions)|[hasSpill](/javascript/api/excel/excel.rangecollectionloadoptions#hasspill)|对于集合中的每一项: 表示是否所有单元格都有溢出边框。|
||[height](/javascript/api/excel/excel.rangecollectionloadoptions#height)|对于集合中的每一项: 返回以磅为单位的距离, 100 以磅为单位, 从区域的上边缘到区域的下边缘之间的距离 (以磅为单位)。 只读。|
||[left](/javascript/api/excel/excel.rangecollectionloadoptions#left)|对于集合中的每个项目: 返回以磅为单位的距离, 以磅为单位, 从工作表左边缘到区域左边缘的比例为 100%。 只读。|
||[savedAsArray](/javascript/api/excel/excel.rangecollectionloadoptions#savedasarray)|对于集合中的每一项: 表示是否将所有单元格都保存为数组公式。|
||[top](/javascript/api/excel/excel.rangecollectionloadoptions#top)|对于集合中的每个项目: 返回以磅为单位的距离, 以磅100为单位, 从工作表的上边缘到区域的上边缘之间的距离 (以磅为单位)。 只读。|
||[width](/javascript/api/excel/excel.rangecollectionloadoptions#width)|对于集合中的每个项目: 返回以磅为单位的距离, 以磅为单位, 从区域的左边缘到该范围的右边缘之间的 100% 缩放。 只读。|
|[RangeData](/javascript/api/excel/excel.rangedata)|[hasSpill](/javascript/api/excel/excel.rangedata#hasspill)|表示所有单元格是否都具有溢出边框。|
||[height](/javascript/api/excel/excel.rangedata#height)|返回从区域的上边缘到区域的下边缘的 100％ 缩放的距离（以磅为单位）。 只读。|
||[left](/javascript/api/excel/excel.rangedata#left)|返回从工作表的左边缘到区域的左边缘的 100％ 缩放的距离（以磅为单位）。 只读。|
||[savedAsArray](/javascript/api/excel/excel.rangedata#savedasarray)|表示是否将所有单元格都保存为数组公式。|
||[top](/javascript/api/excel/excel.rangedata#top)|返回从工作表的上边缘到区域的上边缘的 100％ 缩放的距离（以磅为单位）。 只读。|
||[width](/javascript/api/excel/excel.rangedata#width)|返回从区域的左边缘到区域的右边缘的 100％ 缩放的距离（以磅为单位）。 只读。|
|[RangeLoadOptions](/javascript/api/excel/excel.rangeloadoptions)|[hasSpill](/javascript/api/excel/excel.rangeloadoptions#hasspill)|表示所有单元格是否都具有溢出边框。|
||[height](/javascript/api/excel/excel.rangeloadoptions#height)|返回从区域的上边缘到区域的下边缘的 100％ 缩放的距离（以磅为单位）。 只读。|
||[left](/javascript/api/excel/excel.rangeloadoptions#left)|返回从工作表的左边缘到区域的左边缘的 100％ 缩放的距离（以磅为单位）。 只读。|
||[savedAsArray](/javascript/api/excel/excel.rangeloadoptions#savedasarray)|表示是否将所有单元格都保存为数组公式。|
||[top](/javascript/api/excel/excel.rangeloadoptions#top)|返回从工作表的上边缘到区域的上边缘的 100％ 缩放的距离（以磅为单位）。 只读。|
||[width](/javascript/api/excel/excel.rangeloadoptions#width)|返回从区域的左边缘到区域的右边缘的 100％ 缩放的距离（以磅为单位）。 只读。|
|[Shape](/javascript/api/excel/excel.shape)|[copyTo(destinationSheet?: Worksheet \| string)](/javascript/api/excel/excel.shape#copyto-destinationsheet-)|复制并粘贴 Shape 对象。|
||[placement](/javascript/api/excel/excel.shape#placement)|表示对象如何附加到其下方的单元格。|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|从 XML 字符串创建可缩放的矢量图形 (SVG) 并将其添加到工作表。 返回表示新图片的 Shape 对象。|
|[ShapeCollectionLoadOptions](/javascript/api/excel/excel.shapecollectionloadoptions)|[placement](/javascript/api/excel/excel.shapecollectionloadoptions#placement)|对于集合中的每个项目: 代表如何将对象附加到其下的单元格。|
|[ShapeData](/javascript/api/excel/excel.shapedata)|[placement](/javascript/api/excel/excel.shapedata#placement)|表示对象如何附加到其下方的单元格。|
|[ShapeLoadOptions](/javascript/api/excel/excel.shapeloadoptions)|[placement](/javascript/api/excel/excel.shapeloadoptions#placement)|表示对象如何附加到其下方的单元格。|
|[ShapeUpdateData](/javascript/api/excel/excel.shapeupdatedata)|[placement](/javascript/api/excel/excel.shapeupdatedata#placement)|表示对象如何附加到其下方的单元格。|
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
||[set (properties: Excel. 切片器)](/javascript/api/excel/excel.slicer#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: SlicerUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.slicer#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
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
|[SlicerCollectionData](/javascript/api/excel/excel.slicercollectiondata)|[items](/javascript/api/excel/excel.slicercollectiondata#items)||
|[SlicerCollectionLoadOptions](/javascript/api/excel/excel.slicercollectionloadoptions)|[$all](/javascript/api/excel/excel.slicercollectionloadoptions#$all)||
||[caption](/javascript/api/excel/excel.slicercollectionloadoptions#caption)|对于集合中的每一项: 表示切片器的标题。|
||[height](/javascript/api/excel/excel.slicercollectionloadoptions#height)|对于集合中的每一项: 表示切片器的高度 (以磅为单位)。|
||[id](/javascript/api/excel/excel.slicercollectionloadoptions#id)|对于集合中的每一项: 表示切片器的唯一 id。 只读。|
||[isFilterCleared](/javascript/api/excel/excel.slicercollectionloadoptions#isfiltercleared)|对于集合中的每一项: 如果清除当前应用于切片器上的所有筛选器, 则为 True。|
||[left](/javascript/api/excel/excel.slicercollectionloadoptions#left)|对于集合中的每一项: 代表从切片器左侧到工作表左侧的距离 (以磅为单位)。|
||[name](/javascript/api/excel/excel.slicercollectionloadoptions#name)|对于集合中的每一项: 表示切片器的名称。|
||[nameInFormula](/javascript/api/excel/excel.slicercollectionloadoptions#nameinformula)|对于集合中的每一项: 代表在公式中使用的切片器名称。|
||[sortBy](/javascript/api/excel/excel.slicercollectionloadoptions#sortby)|对于集合中的每一项: 表示切片器中项目的排序顺序。 可能的值为：DataSourceOrder、Ascending、Descending。|
||[style](/javascript/api/excel/excel.slicercollectionloadoptions#style)|对于集合中的每一项: 表示切片器样式的常量值。 可能的值为: "SlicerStyleLight1" 到 "SlicerStyleLight6"、"TableStyleOther1" 通过 "TableStyleOther2"、"SlicerStyleDark1" 到 "SlicerStyleDark6"。 还可以指定工作簿中显示的用户定义的自定义样式。|
||[top](/javascript/api/excel/excel.slicercollectionloadoptions#top)|对于集合中的每一项: 表示从切片器的上边缘到工作表顶部的距离 (以磅为单位)。|
||[width](/javascript/api/excel/excel.slicercollectionloadoptions#width)|对于集合中的每一项: 表示切片器的宽度 (以磅为单位)。|
||[worksheet](/javascript/api/excel/excel.slicercollectionloadoptions#worksheet)|对于集合中的每一项: 代表包含切片器的工作表。|
|[SlicerCollectionUpdateData](/javascript/api/excel/excel.slicercollectionupdatedata)|[items](/javascript/api/excel/excel.slicercollectionupdatedata#items)||
|[SlicerData](/javascript/api/excel/excel.slicerdata)|[caption](/javascript/api/excel/excel.slicerdata#caption)|表示切片器的标题。|
||[height](/javascript/api/excel/excel.slicerdata#height)|表示切片器的高度（以磅为单位）。|
||[id](/javascript/api/excel/excel.slicerdata#id)|表示切片器的唯一 ID。 只读。|
||[isFilterCleared](/javascript/api/excel/excel.slicerdata#isfiltercleared)|如果已清除当前切片器上应用的所有筛选器，则为 True。|
||[left](/javascript/api/excel/excel.slicerdata#left)|表示从切片器左侧到工作表左侧的距离（以磅为单位）。|
||[name](/javascript/api/excel/excel.slicerdata#name)|表示切片器的名称。|
||[nameInFormula](/javascript/api/excel/excel.slicerdata#nameinformula)|表示公式中使用切片器名称。|
||[slicerItems](/javascript/api/excel/excel.slicerdata#sliceritems)|表示作为切片器一部分的 SlicerItems 的集合。 只读。|
||[sortBy](/javascript/api/excel/excel.slicerdata#sortby)|表示切片器中的项目的排序顺序。 可能的值为：DataSourceOrder、Ascending、Descending。|
||[style](/javascript/api/excel/excel.slicerdata#style)|表示切片器样式的常量值。 可能的值为: "SlicerStyleLight1" 到 "SlicerStyleLight6"、"TableStyleOther1" 通过 "TableStyleOther2"、"SlicerStyleDark1" 到 "SlicerStyleDark6"。 还可以指定工作簿中显示的用户定义的自定义样式。|
||[top](/javascript/api/excel/excel.slicerdata#top)|表示从切片器上边缘到工作表顶部的距离（以磅为单位）。|
||[width](/javascript/api/excel/excel.slicerdata#width)|表示切片器的宽度（以磅为单位）。|
||[worksheet](/javascript/api/excel/excel.slicerdata#worksheet)|表示包含切片器的工作表。 只读。|
|[SlicerItem](/javascript/api/excel/excel.sliceritem)|[isSelected](/javascript/api/excel/excel.sliceritem#isselected)|如果选择了切片器项，则为 True。|
||[hasData](/javascript/api/excel/excel.sliceritem#hasdata)|如果切片器项包含数据，则为 True。|
||[key](/javascript/api/excel/excel.sliceritem#key)|表示代表切片器项的唯一值。|
||[name](/javascript/api/excel/excel.sliceritem#name)|代表 UI 中显示的标题。|
||[set (properties: SlicerItem)](/javascript/api/excel/excel.sliceritem#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: SlicerItemUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.sliceritem#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
|[SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection)|[getCount()](/javascript/api/excel/excel.sliceritemcollection#getcount--)|返回切片器中的切片器项的数量。|
||[getItem(key: string)](/javascript/api/excel/excel.sliceritemcollection#getitem-key-)|使用其键或名称获取切片器项对象。|
||[getItemAt(index: number)](/javascript/api/excel/excel.sliceritemcollection#getitemat-index-)|根据其在集合中的位置获取切片器项。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.sliceritemcollection#getitemornullobject-key-)|使用其键或名称获取切片器项。 如果没有切片器项，将返回 null 对象。|
||[items](/javascript/api/excel/excel.sliceritemcollection#items)|获取此集合中已加载的子项。|
|[SlicerItemCollectionData](/javascript/api/excel/excel.sliceritemcollectiondata)|[items](/javascript/api/excel/excel.sliceritemcollectiondata#items)||
|[SlicerItemCollectionLoadOptions](/javascript/api/excel/excel.sliceritemcollectionloadoptions)|[$all](/javascript/api/excel/excel.sliceritemcollectionloadoptions#$all)||
||[hasData](/javascript/api/excel/excel.sliceritemcollectionloadoptions#hasdata)|对于集合中的每一项: 如果切片器项包含数据, 则为 True。|
||[isSelected](/javascript/api/excel/excel.sliceritemcollectionloadoptions#isselected)|对于集合中的每一项: 如果选择了切片器项, 则为 True。|
||[key](/javascript/api/excel/excel.sliceritemcollectionloadoptions#key)|对于集合中的每一项: 表示切片器项的唯一值。|
||[name](/javascript/api/excel/excel.sliceritemcollectionloadoptions#name)|对于集合中的每一项: 代表 UI 中显示的标题。|
|[SlicerItemCollectionUpdateData](/javascript/api/excel/excel.sliceritemcollectionupdatedata)|[items](/javascript/api/excel/excel.sliceritemcollectionupdatedata#items)||
|[SlicerItemData](/javascript/api/excel/excel.sliceritemdata)|[hasData](/javascript/api/excel/excel.sliceritemdata#hasdata)|如果切片器项包含数据，则为 True。|
||[isSelected](/javascript/api/excel/excel.sliceritemdata#isselected)|如果选择了切片器项，则为 True。|
||[key](/javascript/api/excel/excel.sliceritemdata#key)|表示代表切片器项的唯一值。|
||[name](/javascript/api/excel/excel.sliceritemdata#name)|代表 UI 中显示的标题。|
|[SlicerItemLoadOptions](/javascript/api/excel/excel.sliceritemloadoptions)|[$all](/javascript/api/excel/excel.sliceritemloadoptions#$all)||
||[hasData](/javascript/api/excel/excel.sliceritemloadoptions#hasdata)|如果切片器项包含数据，则为 True。|
||[isSelected](/javascript/api/excel/excel.sliceritemloadoptions#isselected)|如果选择了切片器项，则为 True。|
||[key](/javascript/api/excel/excel.sliceritemloadoptions#key)|表示代表切片器项的唯一值。|
||[name](/javascript/api/excel/excel.sliceritemloadoptions#name)|代表 UI 中显示的标题。|
|[SlicerItemUpdateData](/javascript/api/excel/excel.sliceritemupdatedata)|[isSelected](/javascript/api/excel/excel.sliceritemupdatedata#isselected)|如果选择了切片器项，则为 True。|
|[SlicerLoadOptions](/javascript/api/excel/excel.slicerloadoptions)|[$all](/javascript/api/excel/excel.slicerloadoptions#$all)||
||[caption](/javascript/api/excel/excel.slicerloadoptions#caption)|表示切片器的标题。|
||[height](/javascript/api/excel/excel.slicerloadoptions#height)|表示切片器的高度（以磅为单位）。|
||[id](/javascript/api/excel/excel.slicerloadoptions#id)|表示切片器的唯一 ID。 只读。|
||[isFilterCleared](/javascript/api/excel/excel.slicerloadoptions#isfiltercleared)|如果已清除当前切片器上应用的所有筛选器，则为 True。|
||[left](/javascript/api/excel/excel.slicerloadoptions#left)|表示从切片器左侧到工作表左侧的距离（以磅为单位）。|
||[name](/javascript/api/excel/excel.slicerloadoptions#name)|表示切片器的名称。|
||[nameInFormula](/javascript/api/excel/excel.slicerloadoptions#nameinformula)|表示公式中使用切片器名称。|
||[sortBy](/javascript/api/excel/excel.slicerloadoptions#sortby)|表示切片器中的项目的排序顺序。 可能的值为：DataSourceOrder、Ascending、Descending。|
||[style](/javascript/api/excel/excel.slicerloadoptions#style)|表示切片器样式的常量值。 可能的值为: "SlicerStyleLight1" 到 "SlicerStyleLight6"、"TableStyleOther1" 通过 "TableStyleOther2"、"SlicerStyleDark1" 到 "SlicerStyleDark6"。 还可以指定工作簿中显示的用户定义的自定义样式。|
||[top](/javascript/api/excel/excel.slicerloadoptions#top)|表示从切片器上边缘到工作表顶部的距离（以磅为单位）。|
||[width](/javascript/api/excel/excel.slicerloadoptions#width)|表示切片器的宽度（以磅为单位）。|
||[worksheet](/javascript/api/excel/excel.slicerloadoptions#worksheet)|表示包含切片器的工作表。|
|[SlicerStyle](/javascript/api/excel/excel.slicerstyle)|[delete()](/javascript/api/excel/excel.slicerstyle#delete--)|删除 SlicerStyle。|
||[duplicate()](/javascript/api/excel/excel.slicerstyle#duplicate--)|使用所有样式元素的副本创建此 SlicerStyle 的副本。|
||[name](/javascript/api/excel/excel.slicerstyle#name)|获取 SlicerStyle 的名称。|
||[readOnly](/javascript/api/excel/excel.slicerstyle#readonly)|指定此 SlicerStyle 对象是否为只读。 只读。|
||[set (properties: SlicerStyle)](/javascript/api/excel/excel.slicerstyle#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: SlicerStyleUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.slicerstyle#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
|[SlicerStyleCollection](/javascript/api/excel/excel.slicerstylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.slicerstylecollection#add-name--makeuniquename-)|使用指定名称创建空白 SlicerStyle。|
||[getCount()](/javascript/api/excel/excel.slicerstylecollection#getcount--)|获取集合中的切片器样式数量。|
||[getDefault()](/javascript/api/excel/excel.slicerstylecollection#getdefault--)|获取父对象范围的默认 SlicerStyle。|
||[getItem(name: string)](/javascript/api/excel/excel.slicerstylecollection#getitem-name-)|按名称获取 SlicerStyle。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.slicerstylecollection#getitemornullobject-name-)|按名称获取 SlicerStyle。 如果没有 SlicerStyle，将返回 null 对象。|
||[items](/javascript/api/excel/excel.slicerstylecollection#items)|获取此集合中已加载的子项。|
||[setDefault(newDefaultStyle: SlicerStyle \| string)](/javascript/api/excel/excel.slicerstylecollection#setdefault-newdefaultstyle-)|设置在父对象范围内使用的默认 SlicerStyle。|
|[SlicerStyleCollectionData](/javascript/api/excel/excel.slicerstylecollectiondata)|[items](/javascript/api/excel/excel.slicerstylecollectiondata#items)||
|[SlicerStyleCollectionLoadOptions](/javascript/api/excel/excel.slicerstylecollectionloadoptions)|[$all](/javascript/api/excel/excel.slicerstylecollectionloadoptions#$all)||
||[name](/javascript/api/excel/excel.slicerstylecollectionloadoptions#name)|对于集合中的每一项: 获取 SlicerStyle 的名称。|
||[readOnly](/javascript/api/excel/excel.slicerstylecollectionloadoptions#readonly)|对于集合中的每一项: 指定此 SlicerStyle 对象是否为只读。 只读。|
|[SlicerStyleCollectionUpdateData](/javascript/api/excel/excel.slicerstylecollectionupdatedata)|[items](/javascript/api/excel/excel.slicerstylecollectionupdatedata#items)||
|[SlicerStyleData](/javascript/api/excel/excel.slicerstyledata)|[name](/javascript/api/excel/excel.slicerstyledata#name)|获取 SlicerStyle 的名称。|
||[readOnly](/javascript/api/excel/excel.slicerstyledata#readonly)|指定此 SlicerStyle 对象是否为只读。 只读。|
|[SlicerStyleLoadOptions](/javascript/api/excel/excel.slicerstyleloadoptions)|[$all](/javascript/api/excel/excel.slicerstyleloadoptions#$all)||
||[name](/javascript/api/excel/excel.slicerstyleloadoptions#name)|获取 SlicerStyle 的名称。|
||[readOnly](/javascript/api/excel/excel.slicerstyleloadoptions#readonly)|指定此 SlicerStyle 对象是否为只读。 只读。|
|[SlicerStyleUpdateData](/javascript/api/excel/excel.slicerstyleupdatedata)|[name](/javascript/api/excel/excel.slicerstyleupdatedata#name)|获取 SlicerStyle 的名称。|
|[SlicerUpdateData](/javascript/api/excel/excel.slicerupdatedata)|[caption](/javascript/api/excel/excel.slicerupdatedata#caption)|表示切片器的标题。|
||[height](/javascript/api/excel/excel.slicerupdatedata#height)|表示切片器的高度（以磅为单位）。|
||[left](/javascript/api/excel/excel.slicerupdatedata#left)|表示从切片器左侧到工作表左侧的距离（以磅为单位）。|
||[name](/javascript/api/excel/excel.slicerupdatedata#name)|表示切片器的名称。|
||[nameInFormula](/javascript/api/excel/excel.slicerupdatedata#nameinformula)|表示公式中使用切片器名称。|
||[sortBy](/javascript/api/excel/excel.slicerupdatedata#sortby)|表示切片器中的项目的排序顺序。 可能的值为：DataSourceOrder、Ascending、Descending。|
||[style](/javascript/api/excel/excel.slicerupdatedata#style)|表示切片器样式的常量值。 可能的值为: "SlicerStyleLight1" 到 "SlicerStyleLight6"、"TableStyleOther1" 通过 "TableStyleOther2"、"SlicerStyleDark1" 到 "SlicerStyleDark6"。 还可以指定工作簿中显示的用户定义的自定义样式。|
||[top](/javascript/api/excel/excel.slicerupdatedata#top)|表示从切片器上边缘到工作表顶部的距离（以磅为单位）。|
||[width](/javascript/api/excel/excel.slicerupdatedata#width)|表示切片器的宽度（以磅为单位）。|
||[worksheet](/javascript/api/excel/excel.slicerupdatedata#worksheet)|表示包含切片器的工作表。|
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
||[set (properties: TableStyle)](/javascript/api/excel/excel.tablestyle#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: TableStyleUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.tablestyle#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
|[TableStyleCollection](/javascript/api/excel/excel.tablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.tablestylecollection#add-name--makeuniquename-)|使用指定名称创建空白 TableStyle。|
||[getCount()](/javascript/api/excel/excel.tablestylecollection#getcount--)|获取集合中表格样式的数量。|
||[getDefault()](/javascript/api/excel/excel.tablestylecollection#getdefault--)|获取父对象范围的默认 TableStyle。|
||[getItem(name: string)](/javascript/api/excel/excel.tablestylecollection#getitem-name-)|按名称获取 TableStyle。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.tablestylecollection#getitemornullobject-name-)|按名称获取 TableStyle。 如果没有 TableStyle，将返回 null 对象。|
||[items](/javascript/api/excel/excel.tablestylecollection#items)|获取此集合中已加载的子项。|
||[setDefault(newDefaultStyle: TableStyle \| string)](/javascript/api/excel/excel.tablestylecollection#setdefault-newdefaultstyle-)|设置在父对象范围内使用的默认 TableStyle。|
|[TableStyleCollectionData](/javascript/api/excel/excel.tablestylecollectiondata)|[items](/javascript/api/excel/excel.tablestylecollectiondata#items)||
|[TableStyleCollectionLoadOptions](/javascript/api/excel/excel.tablestylecollectionloadoptions)|[$all](/javascript/api/excel/excel.tablestylecollectionloadoptions#$all)||
||[name](/javascript/api/excel/excel.tablestylecollectionloadoptions#name)|对于集合中的每一项: 获取 TableStyle 的名称。|
||[readOnly](/javascript/api/excel/excel.tablestylecollectionloadoptions#readonly)|对于集合中的每一项: 指定此 TableStyle 对象是否为只读。 只读。|
|[TableStyleCollectionUpdateData](/javascript/api/excel/excel.tablestylecollectionupdatedata)|[items](/javascript/api/excel/excel.tablestylecollectionupdatedata#items)||
|[TableStyleData](/javascript/api/excel/excel.tablestyledata)|[name](/javascript/api/excel/excel.tablestyledata#name)|获取 TableStyle 的名称。|
||[readOnly](/javascript/api/excel/excel.tablestyledata#readonly)|指定此 TableStyle 对象是否为只读。 只读。|
|[TableStyleLoadOptions](/javascript/api/excel/excel.tablestyleloadoptions)|[$all](/javascript/api/excel/excel.tablestyleloadoptions#$all)||
||[name](/javascript/api/excel/excel.tablestyleloadoptions#name)|获取 TableStyle 的名称。|
||[readOnly](/javascript/api/excel/excel.tablestyleloadoptions#readonly)|指定此 TableStyle 对象是否为只读。 只读。|
|[TableStyleUpdateData](/javascript/api/excel/excel.tablestyleupdatedata)|[name](/javascript/api/excel/excel.tablestyleupdatedata#name)|获取 TableStyle 的名称。|
|[TimelineStyle](/javascript/api/excel/excel.timelinestyle)|[delete()](/javascript/api/excel/excel.timelinestyle#delete--)|删除 TableStyle。|
||[duplicate()](/javascript/api/excel/excel.timelinestyle#duplicate--)|使用所有样式元素的副本创建此 TimelineStyle 的副本。|
||[name](/javascript/api/excel/excel.timelinestyle#name)|获取 TimelineStyle 的名称。|
||[readOnly](/javascript/api/excel/excel.timelinestyle#readonly)|指定此 TimelineStyle 对象是否为只读。 只读。|
||[set (properties: TimelineStyle)](/javascript/api/excel/excel.timelinestyle#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: TimelineStyleUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.timelinestyle#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
|[TimelineStyleCollection](/javascript/api/excel/excel.timelinestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.timelinestylecollection#add-name--makeuniquename-)|使用指定名称创建空白 TimelineStyle。|
||[getCount()](/javascript/api/excel/excel.timelinestylecollection#getcount--)|获取集合中日程表样式的数量。|
||[getDefault()](/javascript/api/excel/excel.timelinestylecollection#getdefault--)|获取父对象范围的默认 TimelineStyle。|
||[getItem(name: string)](/javascript/api/excel/excel.timelinestylecollection#getitem-name-)|按名称获取 TimelineStyle。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.timelinestylecollection#getitemornullobject-name-)|按名称获取 TimelineStyle。 如果没有 TimelineStyle，将返回 null 对象。|
||[items](/javascript/api/excel/excel.timelinestylecollection#items)|获取此集合中已加载的子项。|
||[setDefault(newDefaultStyle: TimelineStyle \| string)](/javascript/api/excel/excel.timelinestylecollection#setdefault-newdefaultstyle-)|设置在父对象范围内使用的默认 TimelineStyle。|
|[TimelineStyleCollectionData](/javascript/api/excel/excel.timelinestylecollectiondata)|[items](/javascript/api/excel/excel.timelinestylecollectiondata#items)||
|[TimelineStyleCollectionLoadOptions](/javascript/api/excel/excel.timelinestylecollectionloadoptions)|[$all](/javascript/api/excel/excel.timelinestylecollectionloadoptions#$all)||
||[name](/javascript/api/excel/excel.timelinestylecollectionloadoptions#name)|对于集合中的每一项: 获取 TimelineStyle 的名称。|
||[readOnly](/javascript/api/excel/excel.timelinestylecollectionloadoptions#readonly)|对于集合中的每一项: 指定此 TimelineStyle 对象是否为只读。 只读。|
|[TimelineStyleCollectionUpdateData](/javascript/api/excel/excel.timelinestylecollectionupdatedata)|[items](/javascript/api/excel/excel.timelinestylecollectionupdatedata#items)||
|[TimelineStyleData](/javascript/api/excel/excel.timelinestyledata)|[name](/javascript/api/excel/excel.timelinestyledata#name)|获取 TimelineStyle 的名称。|
||[readOnly](/javascript/api/excel/excel.timelinestyledata#readonly)|指定此 TimelineStyle 对象是否为只读。 只读。|
|[TimelineStyleLoadOptions](/javascript/api/excel/excel.timelinestyleloadoptions)|[$all](/javascript/api/excel/excel.timelinestyleloadoptions#$all)||
||[name](/javascript/api/excel/excel.timelinestyleloadoptions#name)|获取 TimelineStyle 的名称。|
||[readOnly](/javascript/api/excel/excel.timelinestyleloadoptions#readonly)|指定此 TimelineStyle 对象是否为只读。 只读。|
|[TimelineStyleUpdateData](/javascript/api/excel/excel.timelinestyleupdatedata)|[name](/javascript/api/excel/excel.timelinestyleupdatedata#name)|获取 TimelineStyle 的名称。|
|[Workbook](/javascript/api/excel/excel.workbook)|[close(closeBehavior?: "Save" \| "SkipSave")](/javascript/api/excel/excel.workbook#close-closebehavior-)|关闭当前工作簿。|
||[close(closeBehavior?: Excel.CloseBehavior)](/javascript/api/excel/excel.workbook#close-closebehavior-)|关闭当前工作簿。|
||[getActiveSlicer()](/javascript/api/excel/excel.workbook#getactiveslicer--)|获取工作簿中当前处于活动状态的切片器。 如果没有活动切片器, 则会`ItemNotFound`引发异常。|
||[getActiveSlicerOrNullObject()](/javascript/api/excel/excel.workbook#getactiveslicerornullobject--)|获取工作簿中当前处于活动状态的切片器。 如果没有处于活动状态的切片器，则返回 null 对象。|
||[comments](/javascript/api/excel/excel.workbook#comments)|表示与工作簿关联的批注集合。 只读。|
||[pivotTableStyles](/javascript/api/excel/excel.workbook#pivottablestyles)|表示一组与工作簿相关联的 PivotTableStyles。 只读。|
||[slicerStyles](/javascript/api/excel/excel.workbook#slicerstyles)|表示一组与工作簿相关联的 SlicerStyles。 只读。|
||[slicers](/javascript/api/excel/excel.workbook#slicers)|表示与工作簿关联的切片器集合。 只读。|
||[tableStyles](/javascript/api/excel/excel.workbook#tablestyles)|表示一组与工作簿相关联的 TableStyles。 只读。|
||[timelineStyles](/javascript/api/excel/excel.workbook#timelinestyles)|表示一组与工作簿相关联的 TimelineStyles。 只读。|
||[save(saveBehavior?: "Save" \| "Prompt")](/javascript/api/excel/excel.workbook#save-savebehavior-)|保存当前工作簿。|
||[save(saveBehavior?: Excel.SaveBehavior)](/javascript/api/excel/excel.workbook#save-savebehavior-)|保存当前工作簿。|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|如果工作簿使用 1904 日期系统，则为 True。|
|[WorkbookData](/javascript/api/excel/excel.workbookdata)|[comments](/javascript/api/excel/excel.workbookdata#comments)|表示与工作簿关联的批注集合。 只读。|
||[pivotTableStyles](/javascript/api/excel/excel.workbookdata#pivottablestyles)|表示一组与工作簿相关联的 PivotTableStyles。 只读。|
||[slicerStyles](/javascript/api/excel/excel.workbookdata#slicerstyles)|表示一组与工作簿相关联的 SlicerStyles。 只读。|
||[slicers](/javascript/api/excel/excel.workbookdata#slicers)|表示与工作簿关联的切片器集合。 只读。|
||[tableStyles](/javascript/api/excel/excel.workbookdata#tablestyles)|表示一组与工作簿相关联的 TableStyles。 只读。|
||[timelineStyles](/javascript/api/excel/excel.workbookdata#timelinestyles)|表示一组与工作簿相关联的 TimelineStyles。 只读。|
||[use1904DateSystem](/javascript/api/excel/excel.workbookdata#use1904datesystem)|如果工作簿使用 1904 日期系统，则为 True。|
|[WorkbookLoadOptions](/javascript/api/excel/excel.workbookloadoptions)|[use1904DateSystem](/javascript/api/excel/excel.workbookloadoptions#use1904datesystem)|如果工作簿使用 1904 日期系统，则为 True。|
|[WorkbookUpdateData](/javascript/api/excel/excel.workbookupdatedata)|[use1904DateSystem](/javascript/api/excel/excel.workbookupdatedata#use1904datesystem)|如果工作簿使用 1904 日期系统，则为 True。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[comments](/javascript/api/excel/excel.worksheet#comments)|返回工作表上的所有 Comments 对象的集合。 只读。|
||[onColumnSorted](/javascript/api/excel/excel.worksheet#oncolumnsorted)|在列上排序时发生。|
||[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|在特定工作表上应用筛选器时发生。|
||[onRowHiddenChanged](/javascript/api/excel/excel.worksheet#onrowhiddenchanged)|当特定工作表上的行隐藏状态更改时发生。|
||[onRowSorted](/javascript/api/excel/excel.worksheet#onrowsorted)|在行上排序时发生。|
||[onSingleClicked](/javascript/api/excel/excel.worksheet#onsingleclicked)|在工作表中进行左键单击/点击操作时发生。|
||[slicers](/javascript/api/excel/excel.worksheet#slicers)|返回作为工作表一部分的切片器集合。 只读。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: "None" \| "Before" \| "After" \| "Beginning" \| "End", relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|将工作簿的指定工作表插入当前工作簿。|
||[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|将工作簿的指定工作表插入当前工作簿。|
||[onColumnSorted](/javascript/api/excel/excel.worksheetcollection#oncolumnsorted)|在列上排序时发生。|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|在工作簿中应用任何工作表的筛选器时发生。|
||[onRowHiddenChanged](/javascript/api/excel/excel.worksheetcollection#onrowhiddenchanged)|当工作簿中的任何工作表的行隐藏状态更改时发生。|
||[onRowSorted](/javascript/api/excel/excel.worksheetcollection#onrowsorted)|在行上排序时发生。|
||[onSingleClicked](/javascript/api/excel/excel.worksheetcollection#onsingleclicked)|在工作表集合中发生左击或螺纹操作时发生。|
|[WorksheetColumnSortedEventArgs](/javascript/api/excel/excel.worksheetcolumnsortedeventargs)|[address](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#address)|获取区域地址，该地址表示特定工作表上的选定区域。|
||[source](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#source)|获取事件源。 有关详细信息，请参阅 Excel.EventSource。|
||[type](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#type)|获取事件的类型。 有关详细信息，请参阅 Excel.EventType。|
||[worksheetId](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#worksheetid)|获取发生排序的工作表的 id。|
|[WorksheetData](/javascript/api/excel/excel.worksheetdata)|[comments](/javascript/api/excel/excel.worksheetdata#comments)|返回工作表上的所有 Comments 对象的集合。 只读。|
||[slicers](/javascript/api/excel/excel.worksheetdata#slicers)|返回作为工作表一部分的切片器集合。 只读。|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|表示事件的类型。 有关详细信息，请参阅 Excel.EventType。|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|表示在其中应用筛选器的工作表的 ID。|
|[WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs)|[address](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#address)|获取区域地址，该地址表示特定工作表上的更改区域。|
||[changeType](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#changetype)|获取更改类型，该类型表示 Changed 事件的触发方式。 有关详细信息, 请参阅 RowHiddenChangeType。|
||[source](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#source)|获取事件源。 有关详细信息，请参阅 Excel.EventSource。|
||[type](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#type)|获取事件的类型。 有关详细信息，请参阅 Excel.EventType。|
||[worksheetId](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#worksheetid)|获取其中的数据发生更改的工作表的 ID。|
|[WorksheetRowSortedEventArgs](/javascript/api/excel/excel.worksheetrowsortedeventargs)|[address](/javascript/api/excel/excel.worksheetrowsortedeventargs#address)|获取区域地址，该地址表示特定工作表上的选定区域。|
||[source](/javascript/api/excel/excel.worksheetrowsortedeventargs#source)|获取事件源。 有关详细信息，请参阅 Excel.EventSource。|
||[type](/javascript/api/excel/excel.worksheetrowsortedeventargs#type)|获取事件的类型。 有关详细信息，请参阅 Excel.EventType。|
||[worksheetId](/javascript/api/excel/excel.worksheetrowsortedeventargs#worksheetid)|获取发生排序的工作表的 id。|
|[WorksheetSingleClickedEventArgs](/javascript/api/excel/excel.worksheetsingleclickedeventargs)|[address](/javascript/api/excel/excel.worksheetsingleclickedeventargs#address)|获取特定工作表中表示被左键单击/点击的单元格的地址。|
||[OffsetX](/javascript/api/excel/excel.worksheetsingleclickedeventargs#offsetx)|从左键单击/点击的点到左键单击/点击的单元格的左（RTL 则为右侧）网格线边缘的距离（以磅为单位）。|
||[OffsetY](/javascript/api/excel/excel.worksheetsingleclickedeventargs#offsety)|从左键单击/点击的点到左键单击/点击的单元格的顶部网格线边缘的距离（以磅为单位）。|
||[type](/javascript/api/excel/excel.worksheetsingleclickedeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.worksheetsingleclickedeventargs#worksheetid)|获取已在其中左键单击/点击单元格的工作表的 ID。|

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 参考文档](/javascript/api/excel?view=excel-js-preview)
- [Excel JavaScript API 要求集](./excel-api-requirement-sets.md)
