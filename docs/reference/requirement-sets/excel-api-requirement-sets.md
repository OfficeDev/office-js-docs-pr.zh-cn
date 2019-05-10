---
title: Excel JavaScript API 要求集
description: ''
ms.date: 05/06/2019
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 1735c01a8c17c31e632432d914770a800846508e
ms.sourcegitcommit: 5b9c2b39dfe76cabd98bf28d5287d9718788e520
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/07/2019
ms.locfileid: "33659640"
---
# <a name="excel-javascript-api-requirement-sets"></a>Excel JavaScript API 要求集

要求集是指各组已命名的 API 成员。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 主机是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](/office/dev/add-ins/develop/office-versions-and-requirement-sets)。

Excel 加载项在多个 Office 版本中运行，包括 Office 2016 for Windows 或更高版本、Office for iPad、Office for Mac 和 Office Online。下表列出了 Excel 要求集、支持各个要求集的 Office 主机应用程序，以及这些应用程序的内部版本或版本号。

> [!NOTE]
> 若要在任何编号的要求集中使用 API，你应该引用 CDN 上的**生产**库：https://appsforoffice.microsoft.com/lib/1/hosted/office.js。
>
> 有关使用预览 API 的信息，请参阅本文的 [Excel JavaScript 预览 API](#excel-javascript-preview-apis) 部分。

|  要求集  |  Office 365 for Windows  |  Office 365 for iPad  |  Office 365 for Mac  | Office Online  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|
| [预览](/javascript/api/excel)  | 请使用最新的 Office 版本来试用预览 API（你可能需要加入 [Office 预览体验成员计划](https://products.office.com/office-insider)） |
| ExcelApi 1.9  | 版本 1903 (内部版本 11425.20204) 或更高版本 | 2.24 或更高版本 | 16.24 或更高版本 | 2019 年 5 月       | 即将推出 |
| ExcelApi 1.8  | 版本 1808（内部版本 10730.20102）或更高版本 | 2.17 或更高版本 | 16.17 或更高版本 | 2018 年 9 月 | 即将推出 |
| ExcelApi 1.7  | 版本 1801（内部版本 9001.2171）或更高版本   | 2.9 或更高版本  | 16.9 或更高版本  | 2018 年 4 月     | 即将推出 |
| ExcelApi 1.6  | 版本 1704（生成号 8201.2001）或更高版本   | 2.2 或更高版本  | 15.36 或更高版本 | 2017 年 4 月     | 即将推出 |
| ExcelApi 1.5  | 版本 1703（内部版本 8067.2070）或更高版本   | 2.2 或更高版本  | 15.36 或更高版本 | 2017 年 3 月     | 即将推出 |
| ExcelApi 1.4  | 版本 1701（内部版本 7870.2024）或更高版本   | 2.2 或更高版本  | 15.36 或更高版本 | 2017 年 1 月   | 即将推出 |
| ExcelApi 1.3  | 版本 1608（内部版本 7369.2055）或更高版本   | 1.27 或更高版本 | 15.27 或更高版本 | 2016 年 9 月 | 版本 1608（内部版本 7601.6800）或更高版本|
| ExcelApi 1.2  | 版本 1601（内部版本 6741.2088）或更高版本   | 1.21 或更高版本 | 15.22 或更高版本 | 2016 年 1 月   ||
| ExcelApi 1.1  | 版本 1509（内部版本 4266.1001）或更高版本   | 1.19 或更高版本 | 15.20 或更高版本 | 2016 年 1 月   ||

> [!NOTE]
> 通过 MSI 安装的 Office 2016 的内部版本号为 16.0.4266.1001。 此版本只包含 ExcelApi 1.1 要求集。

## <a name="custom-functions"></a>自定义函数

[自定义函数](../../excel/custom-functions-overview.md)使用独立于核心 Excel JavaScript API 的要求集。 下表列出了自定义函数要求集、支持的 Office 主机应用程序，以及这些应用程序的内部版本或版本号。

|  要求集  |  Office 365 for Windows  |  Office 365 for iPad  |  Office 365 for Mac  | Office Online | Office Online Server |
|:-----|-----|:-----|:-----|:-----|:-----|
| CustomFunctionsRuntime 1.1 | 版本 1904（内部版本 11601.20144）或更高版本 | 不支持 | 16.24 或更高版本 | 2019 年 4 月 | 即将推出 |

有关版本、内部版本号和 Office Online Server 的详细信息，请参阅：

- [更新频道发布的 Office 365 客户端版本号和内部版本号](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [使用的是哪一版 Office？](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [在哪里可以找到 Office 365 客户端应用程序的版本号和内部版本号](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Office Online Server 概述](/officeonlineserver/office-online-server-overview)

## <a name="excel-javascript-preview-apis"></a>Excel JavaScript 预览 API

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

以下是当前预览版中的 API 的完整列表。

| Class | 域 | 说明 |
|:---|:---|:---|
|[Comment](/javascript/api/excel/excel.comment)|[content](/javascript/api/excel/excel.comment#content)|获取或设置批注的内容。|
||[delete()](/javascript/api/excel/excel.comment#delete--)|删除批注线程。|
||[getLocation()](/javascript/api/excel/excel.comment#getlocation--)|获取批注的位置。|
||[authorEmail](/javascript/api/excel/excel.comment#authoremail)|获取批注作者的电子邮件。|
||[authorName](/javascript/api/excel/excel.comment#authorname)|获取批注作者的姓名。|
||[creationDate](/javascript/api/excel/excel.comment#creationdate)|获取批注的创建时间。 如果批注是从备注转换而来的，则返回 null，因为批注没有创建日期。|
||[id](/javascript/api/excel/excel.comment#id)|表示批注标识符。 只读。|
||[replies](/javascript/api/excel/excel.comment#replies)|表示与批注关联的回复对象的集合。 只读。|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[add(content: string, cellAddress: Range \| string, contentType?: "Plain")](/javascript/api/excel/excel.commentcollection#add-content--celladdress--contenttype-)|使用给定单元格上的给定内容创建新批注。 如果所提供的区域大于一个单元格，则会引发无效的参数错误。|
||[add(content: string, cellAddress: Range \| string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentcollection#add-content--celladdress--contenttype-)|使用给定单元格上的给定内容创建新批注。 如果所提供的区域大于一个单元格，则会引发无效的参数错误。|
||[getCount()](/javascript/api/excel/excel.commentcollection#getcount--)|获取集合中的批注数量。|
||[getItem(commentId: string)](/javascript/api/excel/excel.commentcollection#getitem-commentid-)|根据其 ID 从集合中获取批注。 只读。|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentcollection#getitemat-index-)|根据其位置从集合中获取批注。|
||[getItemByCell(cellAddress: Range \| string)](/javascript/api/excel/excel.commentcollection#getitembycell-celladdress-)|从指定单元格获取的批注。|
||[getItemByReplyId(replyId: string)](/javascript/api/excel/excel.commentcollection#getitembyreplyid-replyid-)|从具有相应回复 ID 的集合中获取批注。|
||[items](/javascript/api/excel/excel.commentcollection#items)|获取此集合中已加载的子项。|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[content](/javascript/api/excel/excel.commentreply#content)|获取或设置批注回复的内容。|
||[delete()](/javascript/api/excel/excel.commentreply#delete--)|删除批注回复。|
||[getLocation()](/javascript/api/excel/excel.commentreply#getlocation--)|获取批注回复的位置。|
||[getParentComment()](/javascript/api/excel/excel.commentreply#getparentcomment--)|获取此回复的父批注。|
||[authorEmail](/javascript/api/excel/excel.commentreply#authoremail)|获取批注回复作者的电子邮件。|
||[authorName](/javascript/api/excel/excel.commentreply#authorname)|获取批注回复作者的姓名。|
||[creationDate](/javascript/api/excel/excel.commentreply#creationdate)|获取批注回复的创建时间。|
||[id](/javascript/api/excel/excel.commentreply#id)|表示批注回复标识符。 只读。|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[add(content: string, contentType?: "Plain")](/javascript/api/excel/excel.commentreplycollection#add-content--contenttype-)|为批注创建批注回复。|
||[add(content: string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentreplycollection#add-content--contenttype-)|为批注创建批注回复。|
||[getCount()](/javascript/api/excel/excel.commentreplycollection#getcount--)|获取集合中的批注回复数量。|
||[getItem(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#getitem-commentreplyid-)|返回由其 ID 标识的批注回复。 只读。|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentreplycollection#getitemat-index-)|根据其在集合中的位置获取批注回复。|
||[items](/javascript/api/excel/excel.commentreplycollection#items)|获取此集合中已加载的子项。|
|[CustomFunctionEventArgs](/javascript/api/excel/excel.customfunctioneventargs)|[higherTicks](/javascript/api/excel/excel.customfunctioneventargs#higherticks)||
||[lowerTicks](/javascript/api/excel/excel.customfunctioneventargs#lowerticks)||
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
||[hasSpill](/javascript/api/excel/excel.range#hasspill)|表示所有单元格是否都具有溢出边框。|
||[height](/javascript/api/excel/excel.range#height)|返回从区域的上边缘到区域的下边缘的 100％ 缩放的距离（以磅为单位）。 只读。|
||[left](/javascript/api/excel/excel.range#left)|返回从工作表的左边缘到区域的左边缘的 100％ 缩放的距离（以磅为单位）。 只读。|
||[top](/javascript/api/excel/excel.range#top)|返回从工作表的上边缘到区域的上边缘的 100％ 缩放的距离（以磅为单位）。 只读。|
||[width](/javascript/api/excel/excel.range#width)|返回从区域的左边缘到区域的右边缘的 100％ 缩放的距离（以磅为单位）。 只读。|
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
||[style](/javascript/api/excel/excel.slicer#style)|表示切片器样式的常量值。 可能的值为：SlicerStyleLight1 through SlicerStyleLight6、TableStyleOther1 through TableStyleOther2、 SlicerStyleDark1 through SlicerStyleDark6。 还可以指定工作簿中显示的用户定义的自定义样式。|
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
||[name](/javascript/api/excel/excel.sliceritem#name)|表示 UI 上显示的值。|
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
|[Workbook](/javascript/api/excel/excel.workbook)|[close(closeBehavior?: "Save" \| "SkipSave")](/javascript/api/excel/excel.workbook#close-closebehavior-)|关闭当前工作簿。|
||[close(closeBehavior?: Excel.CloseBehavior)](/javascript/api/excel/excel.workbook#close-closebehavior-)|关闭当前工作簿。|
||[getActiveSlicer()](/javascript/api/excel/excel.workbook#getactiveslicer--)|获取工作簿中当前处于活动状态的切片器。 如果没有处于活动状态的切片器，则会引发异常情况。|
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
|[Worksheet](/javascript/api/excel/excel.worksheet)|[comments](/javascript/api/excel/excel.worksheet#comments)|返回工作表上的所有 Comments 对象的集合。 只读。|
||[onColumnSorted](/javascript/api/excel/excel.worksheet#oncolumnsorted)|在列上排序时发生。|
||[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|在特定工作表上应用筛选器时发生。|
||[onRowSorted](/javascript/api/excel/excel.worksheet#onrowsorted)|在行上排序时发生。|
||[onSingleClicked](/javascript/api/excel/excel.worksheet#onsingleclicked)|在工作表中进行左键单击/点击操作时发生。|
||[slicers](/javascript/api/excel/excel.worksheet#slicers)|返回作为工作表一部分的切片器集合。 只读。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: "None" \| "Before" \| "After" \| "Beginning" \| "End", relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|将工作簿的指定工作表插入当前工作簿。|
||[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|将工作簿的指定工作表插入当前工作簿。|
||[onColumnSorted](/javascript/api/excel/excel.worksheetcollection#oncolumnsorted)|在列上排序时发生。|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|在工作簿中应用任何工作表的筛选器时发生。|
||[onRowSorted](/javascript/api/excel/excel.worksheetcollection#onrowsorted)|在行上排序时发生。|
|[WorksheetColumnSortedEventArgs](/javascript/api/excel/excel.worksheetcolumnsortedeventargs)|[address](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#address)|获取区域地址，该地址表示特定工作表上的选定区域。|
||[source](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#source)|获取事件源。 有关详细信息，请参阅 Excel.EventSource。|
||[type](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#type)|获取事件的类型。 有关详细信息，请参阅 Excel.EventType。|
||[worksheetId](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#worksheetid)|获取发生排序的工作表的 id。|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|表示事件的类型。 有关详细信息，请参阅 Excel.EventType。|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|表示在其中应用筛选器的工作表的 ID。|
|[WorksheetRowSortedEventArgs](/javascript/api/excel/excel.worksheetrowsortedeventargs)|[address](/javascript/api/excel/excel.worksheetrowsortedeventargs#address)|获取区域地址，该地址表示特定工作表上的选定区域。|
||[source](/javascript/api/excel/excel.worksheetrowsortedeventargs#source)|获取事件源。 有关详细信息，请参阅 Excel.EventSource。|
||[type](/javascript/api/excel/excel.worksheetrowsortedeventargs#type)|获取事件的类型。 有关详细信息，请参阅 Excel.EventType。|
||[worksheetId](/javascript/api/excel/excel.worksheetrowsortedeventargs#worksheetid)|获取发生排序的工作表的 id。|
|[WorksheetSingleClickedEventArgs](/javascript/api/excel/excel.worksheetsingleclickedeventargs)|[address](/javascript/api/excel/excel.worksheetsingleclickedeventargs#address)|获取特定工作表中表示被左键单击/点击的单元格的地址。|
||[OffsetX](/javascript/api/excel/excel.worksheetsingleclickedeventargs#offsetx)|从左键单击/点击的点到左键单击/点击的单元格的左（RTL 则为右侧）网格线边缘的距离（以磅为单位）。|
||[OffsetY](/javascript/api/excel/excel.worksheetsingleclickedeventargs#offsety)|从左键单击/点击的点到左键单击/点击的单元格的顶部网格线边缘的距离（以磅为单位）。|
||[type](/javascript/api/excel/excel.worksheetsingleclickedeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.worksheetsingleclickedeventargs#worksheetid)|获取已在其中左键单击/点击单元格的工作表的 ID。|

## <a name="whats-new-in-excel-javascript-api-19"></a>Excel JavaScript API 1.9 的最近更新

超过 500 个新  Excel API 随 1.9 要求集一起推出。 第一个表提供了 API 的简明摘要，而后续表提供了详细列表。

| 功能区域 | 说明 | 相关对象 |
|:--- |:--- |:--- |
| [Shape](../../excel/excel-add-ins-shapes.md) | 插入、定位和格式化图像、几何形状和文本框。 | [ShapeCollection](/javascript/api/excel/excel.shapecollection) [Shape](/javascript/api/excel/excel.shape) [GeometricShape](/javascript/api/excel/excel.geometricshape)  [Image](/javascript/api/excel/excel.image) |
| [自动筛选](../../excel/excel-add-ins-worksheets.md#filter-data) | 为区域添加筛选器。 | [AutoFilter](/javascript/api/excel/excel.autofilter) |
| [Areas](../../excel/excel-add-ins-multiple-ranges.md) | 支持非连续区域。 | [RangeAreas](/javascript/api/excel/excel.rangeareas) |
| [特殊单元格](../../excel/excel-add-ins-multiple-ranges.md#get-special-cells-from-multiple-ranges) | 获取在区域内包含日期、备注或公式的单元格。 | [Range](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-)|
| [查找](../../excel/excel-add-ins-ranges.md#find-a-cell-using-string-matching) | 查找区域或工作表中的值或公式。 | [Range](/javascript/api/excel/excel.range#find-text--criteria-)[Worksheet](/javascript/api/excel/excel.worksheet#findall-text--criteria-) |
| [复制和粘贴](../../excel/excel-add-ins-ranges-advanced.md#copy-and-paste) | 将值、格式和公式从一个区域复制到另一个区域。 | [Range](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-) |
| [Calculation](../../excel/performance.md#suspend-calculation-temporarily) | 更好地控制 Excel 计算引擎。 | [应用程序](/javascript/api/excel/excel.application) |
| 新图表 | 了解我们支持的新图表类型：地图、箱形图、瀑布图、旭日图、排列图 和漏斗图。 | [Chart](/javascript/api/excel/excel.charttype) |
| RangeFormat | 新功能及区域格式。 | [Range](/javascript/api/excel/excel.rangeformat) |

以下是 ExcelApi 1.9 要求集中 API 的完整列表。

| Class | 域 | 说明 |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[calculationEngineVersion](/javascript/api/excel/excel.application#calculationengineversion)|返回用于上次完整重新计算的 Excel 计算引擎版本。 只读。|
||[calculationState](/javascript/api/excel/excel.application#calculationstate)|返回应用程序的计算状态。 有关详细信息，请参阅 Excel.CalculationState。 只读。|
||[iterativeCalculation](/javascript/api/excel/excel.application#iterativecalculation)|返回“迭代计算”设置。|
||[suspendScreenUpdatingUntilNextSync()](/javascript/api/excel/excel.application#suspendscreenupdatinguntilnextsync--)|在下一次调用“context.sync()”前暂停屏幕更新。|
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[apply(range: Range \| string, columnIndex?: number, criteria?: Excel.FilterCriteria)](/javascript/api/excel/excel.autofilter#apply-range--columnindex--criteria-)|将自动筛选器应用于区域。 如果指定了列索引和筛选条件，则筛选列。|
||[clearCriteria()](/javascript/api/excel/excel.autofilter#clearcriteria--)|清除自动筛选器的筛选条件。|
||[getRange()](/javascript/api/excel/excel.autofilter#getrange--)|返回 Range 对象，该对象表示“自动筛选”应用的区域。|
||[getRangeOrNullObject()](/javascript/api/excel/excel.autofilter#getrangeornullobject--)|返回 Range 对象，该对象表示“自动筛选”应用的区域。|
||[load(propertyNames?: string \| string[])](/javascript/api/excel/excel.autofilter#load-propertynames-)|将命令加入队列以加载对象的指定属性。 在读取属性之前，你必须调用“context.sync()”。|
||[criteria](/javascript/api/excel/excel.autofilter#criteria)|在自动筛选区域中保留所有筛选条件的数组。 只读。|
||[enabled](/javascript/api/excel/excel.autofilter#enabled)|指示是否启用了自动筛选。 只读。|
||[isDataFiltered](/javascript/api/excel/excel.autofilter#isdatafiltered)|指示自动筛选是否具有筛选条件。 只读。|
||[reapply()](/javascript/api/excel/excel.autofilter#reapply--)|应用当前位于区域上的指定 Autofilter 对象。|
||[remove()](/javascript/api/excel/excel.autofilter#remove--)|删除区域的自动筛选。|
|[CellBorder](/javascript/api/excel/excel.cellborder)|[color](/javascript/api/excel/excel.cellborder#color)|表示`color`单个边框的属性。|
||[style](/javascript/api/excel/excel.cellborder#style)|表示`style`单个边框的属性。|
||[tintAndShade](/javascript/api/excel/excel.cellborder#tintandshade)|表示`tintAndShade`单个边框的属性。|
||[weight](/javascript/api/excel/excel.cellborder#weight)|表示`weight`单个边框的属性。|
|[CellBorderCollection](/javascript/api/excel/excel.cellbordercollection)|[bottom](/javascript/api/excel/excel.cellbordercollection#bottom)|表示`format.borders.bottom`属性。|
||[diagonalDown](/javascript/api/excel/excel.cellbordercollection#diagonaldown)|表示`format.borders.diagonalDown`属性。|
||[diagonalUp](/javascript/api/excel/excel.cellbordercollection#diagonalup)|表示`format.borders.diagonalUp`属性。|
||[horizontal](/javascript/api/excel/excel.cellbordercollection#horizontal)|表示`format.borders.horizontal`属性。|
||[left](/javascript/api/excel/excel.cellbordercollection#left)|表示`format.borders.left`属性。|
||[right](/javascript/api/excel/excel.cellbordercollection#right)|表示`format.borders.right`属性。|
||[top](/javascript/api/excel/excel.cellbordercollection#top)|表示`format.borders.top`属性。|
||[vertical](/javascript/api/excel/excel.cellbordercollection#vertical)|表示`format.borders.vertical`属性。|
|[CellProperties](/javascript/api/excel/excel.cellproperties)|[address](/javascript/api/excel/excel.cellproperties#address)|表示`address`属性。|
||[addressLocal](/javascript/api/excel/excel.cellproperties#addresslocal)|表示`addressLocal`属性。|
||[hidden](/javascript/api/excel/excel.cellproperties#hidden)|表示`hidden`属性。|
|[CellPropertiesFill](/javascript/api/excel/excel.cellpropertiesfill)|[color](/javascript/api/excel/excel.cellpropertiesfill#color)|表示`format.fill.color`属性。|
||[pattern](/javascript/api/excel/excel.cellpropertiesfill#pattern)|表示`format.fill.pattern`属性。|
||[patternColor](/javascript/api/excel/excel.cellpropertiesfill#patterncolor)|表示`format.fill.patternColor`属性。|
||[patternTintAndShade](/javascript/api/excel/excel.cellpropertiesfill#patterntintandshade)|表示`format.fill.patternTintAndShade`属性。|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfill#tintandshade)|表示`format.fill.tintAndShade`属性。|
|[CellPropertiesFont](/javascript/api/excel/excel.cellpropertiesfont)|[bold](/javascript/api/excel/excel.cellpropertiesfont#bold)|表示`format.font.bold`属性。|
||[color](/javascript/api/excel/excel.cellpropertiesfont#color)|表示`format.font.color`属性。|
||[italic](/javascript/api/excel/excel.cellpropertiesfont#italic)|表示`format.font.italic`属性。|
||[name](/javascript/api/excel/excel.cellpropertiesfont#name)|表示`format.font.name`属性。|
||[size](/javascript/api/excel/excel.cellpropertiesfont#size)|表示`format.font.size`属性。|
||[strikethrough](/javascript/api/excel/excel.cellpropertiesfont#strikethrough)|表示`format.font.strikethrough`属性。|
||[subscript](/javascript/api/excel/excel.cellpropertiesfont#subscript)|表示`format.font.subscript`属性。|
||[superscript](/javascript/api/excel/excel.cellpropertiesfont#superscript)|表示`format.font.superscript`属性。|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfont#tintandshade)|表示`format.font.tintAndShade`属性。|
||[underline](/javascript/api/excel/excel.cellpropertiesfont#underline)|表示`format.font.underline`属性。|
|[CellPropertiesFormat](/javascript/api/excel/excel.cellpropertiesformat)|[autoIndent](/javascript/api/excel/excel.cellpropertiesformat#autoindent)|表示`autoIndent`属性。|
||[Borders](/javascript/api/excel/excel.cellpropertiesformat#borders)|表示`borders`属性。|
||[fill](/javascript/api/excel/excel.cellpropertiesformat#fill)|表示`fill`属性。|
||[font](/javascript/api/excel/excel.cellpropertiesformat#font)|表示`font`属性。|
||[horizontalAlignment](/javascript/api/excel/excel.cellpropertiesformat#horizontalalignment)|表示`horizontalAlignment`属性。|
||[indentLevel](/javascript/api/excel/excel.cellpropertiesformat#indentlevel)|表示`indentLevel`属性。|
||[protection](/javascript/api/excel/excel.cellpropertiesformat#protection)|表示`protection`属性。|
||[readingOrder](/javascript/api/excel/excel.cellpropertiesformat#readingorder)|表示`readingOrder`属性。|
||[shrinkToFit](/javascript/api/excel/excel.cellpropertiesformat#shrinktofit)|表示`shrinkToFit`属性。|
||[textOrientation](/javascript/api/excel/excel.cellpropertiesformat#textorientation)|表示`textOrientation`属性。|
||[useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#usestandardheight)|表示`useStandardHeight`属性。|
||[useStandardWidth](/javascript/api/excel/excel.cellpropertiesformat#usestandardwidth)|表示`useStandardWidth`属性。|
||[verticalAlignment](/javascript/api/excel/excel.cellpropertiesformat#verticalalignment)|表示`verticalAlignment`属性。|
||[wrapText](/javascript/api/excel/excel.cellpropertiesformat#wraptext)|表示`wrapText`属性。|
|[CellPropertiesProtection](/javascript/api/excel/excel.cellpropertiesprotection)|[formulaHidden](/javascript/api/excel/excel.cellpropertiesprotection#formulahidden)|表示`format.protection.formulaHidden`属性。|
||[locked](/javascript/api/excel/excel.cellpropertiesprotection#locked)|表示`format.protection.locked`属性。|
|[ChangedEventDetail](/javascript/api/excel/excel.changedeventdetail)|[valueAfter](/javascript/api/excel/excel.changedeventdetail#valueafter)|表示更改之后的值。 返回的数据可能是字符串、数字，也可能是布尔值。 包含错误的单元格将返回错误字符串。|
||[valueBefore](/javascript/api/excel/excel.changedeventdetail#valuebefore)|表示更改之前的值。 返回的数据可能是字符串、数字，也可能是布尔值。 包含错误的单元格将返回错误字符串。|
||[valueTypeAfter](/javascript/api/excel/excel.changedeventdetail#valuetypeafter)|表示更改之后的值类型。|
||[valueTypeBefore](/javascript/api/excel/excel.changedeventdetail#valuetypebefore)|表示更改之前的值类型。|
|[Chart](/javascript/api/excel/excel.chart)|[activate()](/javascript/api/excel/excel.chart#activate--)|在 Excel UI 中激活图表。|
||[pivotOptions](/javascript/api/excel/excel.chart#pivotoptions)|封装数据透视图的选项。 只读。|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[colorScheme](/javascript/api/excel/excel.chartareaformat#colorscheme)|返回或设置图表的配色方案。 读/写。|
||[roundedCorners](/javascript/api/excel/excel.chartareaformat#roundedcorners)|指定图表的图表区域是否有圆角。 读/写。|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[linkNumberFormat](/javascript/api/excel/excel.chartaxis#linknumberformat)|表示数字格式是否链接到单元格。 如果为 true，则数字格式会在单元格中更改时在标签中更改。|
|[ChartBinOptions](/javascript/api/excel/excel.chartbinoptions)|[allowOverflow](/javascript/api/excel/excel.chartbinoptions#allowoverflow)|指定是否在直方图或排列图中启用容器溢出。 读/写。|
||[allowUnderflow](/javascript/api/excel/excel.chartbinoptions#allowunderflow)|指定是否在直方图或排列图中启用容器下溢。 读/写。|
||[count](/javascript/api/excel/excel.chartbinoptions#count)|返回或设置直方图或排列图的容器数量。 读/写。|
||[load(propertyNames?: string \| string[])](/javascript/api/excel/excel.chartbinoptions#load-propertynames-)|将命令加入队列以加载对象的指定属性。 在读取属性之前，你必须调用“context.sync()”。|
||[overflowValue](/javascript/api/excel/excel.chartbinoptions#overflowvalue)|返回或设置直方图或排列图的容器溢出值。 读/写。|
||[type](/javascript/api/excel/excel.chartbinoptions#type)|返回或设置直方图或排列图的容器类型。 读/写。|
||[underflowValue](/javascript/api/excel/excel.chartbinoptions#underflowvalue)|返回或设置直方图或排列图的容器下溢值。 读/写。|
||[width](/javascript/api/excel/excel.chartbinoptions#width)|返回或设置直方图或排列图的容器宽度值。 读/写。|
|[ChartBoxwhiskerOptions](/javascript/api/excel/excel.chartboxwhiskeroptions)|[load(propertyNames?: string \| string[])](/javascript/api/excel/excel.chartboxwhiskeroptions#load-propertynames-)|将命令加入队列以加载对象的指定属性。 在读取属性之前，你必须调用“context.sync()”。|
||[quartileCalculation](/javascript/api/excel/excel.chartboxwhiskeroptions#quartilecalculation)|返回或设置箱形图的四分位点计算类型。 读/写。|
||[showInnerPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#showinnerpoints)|指定在箱形图中是否显示内部点。 读/写。|
||[showMeanLine](/javascript/api/excel/excel.chartboxwhiskeroptions#showmeanline)|指定在箱形图中是否显示中线。 读/写。|
||[showMeanMarker](/javascript/api/excel/excel.chartboxwhiskeroptions#showmeanmarker)|指定在箱形图中是否显示平均值标记。 读/写。|
||[showOutlierPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#showoutlierpoints)|指定在箱形图中是否显示离群值点。 读/写。|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabel#linknumberformat)|布尔值，表示数字格式是否链接到单元格（以便在单元格中更改时标签中的数字格式会发生改变）。|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabels#linknumberformat)|表示数字格式是否链接到单元格。 如果为 true，则数字格式会在单元格中更改时在标签中更改。|
|[ChartErrorBars](/javascript/api/excel/excel.charterrorbars)|[endStyleCap](/javascript/api/excel/excel.charterrorbars#endstylecap)|指定误差线是否具有终止端样式。|
||[include](/javascript/api/excel/excel.charterrorbars#include)|指定包含误差线的哪些部分。|
||[load(propertyNames?: string \| string[])](/javascript/api/excel/excel.charterrorbars#load-propertynames-)|将命令加入队列以加载对象的指定属性。 在读取属性之前，你必须调用“context.sync()”。|
||[format](/javascript/api/excel/excel.charterrorbars#format)|指定误差线的格式类型。|
||[type](/javascript/api/excel/excel.charterrorbars#type)|误差线标记的区域类型。|
||[visible](/javascript/api/excel/excel.charterrorbars#visible)|指定是否显示误差线。|
|[ChartErrorBarsFormat](/javascript/api/excel/excel.charterrorbarsformat)|[load(propertyNames?: string \| string[])](/javascript/api/excel/excel.charterrorbarsformat#load-propertynames-)|将命令加入队列以加载对象的指定属性。 在读取属性之前，你必须调用“context.sync()”。|
||[line](/javascript/api/excel/excel.charterrorbarsformat#line)|表示图表线条格式。|
|[ChartMapOptions](/javascript/api/excel/excel.chartmapoptions)|[labelStrategy](/javascript/api/excel/excel.chartmapoptions#labelstrategy)|返回或设置区域地图图表的系列地图标签策略。 读/写。|
||[level](/javascript/api/excel/excel.chartmapoptions#level)|返回或设置区域地图图表的系列映射级别。 读/写。|
||[load(propertyNames?: string \| string[])](/javascript/api/excel/excel.chartmapoptions#load-propertynames-)|将命令加入队列以加载对象的指定属性。 在读取属性之前，你必须调用“context.sync()”。|
||[projectionType](/javascript/api/excel/excel.chartmapoptions#projectiontype)|返回或设置区域地图图表的系列投影类型。 读/写。|
|[ChartPivotOptions](/javascript/api/excel/excel.chartpivotoptions)|[load(propertyNames?: string \| string[])](/javascript/api/excel/excel.chartpivotoptions#load-propertynames-)|将命令加入队列以加载对象的指定属性。 在读取属性之前，你必须调用“context.sync()”。|
||[showAxisFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showaxisfieldbuttons)|指定是否在数据透视图上显示轴字段按钮。 ShowAxisFieldButtons 属性对应于“分析”选项卡（在选择数据透视图时可用）的“字段按钮”下拉列表中的“显示坐标轴字段按钮”命令。|
||[showLegendFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showlegendfieldbuttons)|指定是否在数据透视图上显示图例字段按钮。|
||[showReportFilterFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showreportfilterfieldbuttons)|指定是否在数据透视图上显示报表筛选字段按钮。|
||[showValueFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showvaluefieldbuttons)|指定是否在数据透视图上显示值字段按钮。|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[bubbleScale](/javascript/api/excel/excel.chartseries#bubblescale)|这可以是从 0（零）到 300 的整数值，表示默认大小的百分比。 该属性仅适用于气泡图。 读/写。|
||[gradientMaximumColor](/javascript/api/excel/excel.chartseries#gradientmaximumcolor)|返回或设置区域地图图表系列的最大值的颜色。 读/写。|
||[gradientMaximumType](/javascript/api/excel/excel.chartseries#gradientmaximumtype)|返回或设置区域地图图表系列的最大值的类型。 读/写。|
||[gradientMaximumValue](/javascript/api/excel/excel.chartseries#gradientmaximumvalue)|返回或设置区域地图图表系列的最大值。 读/写。|
||[gradientMidpointColor](/javascript/api/excel/excel.chartseries#gradientmidpointcolor)|返回或设置区域地图图表系列的中间值的颜色。 读/写。|
||[gradientMidpointType](/javascript/api/excel/excel.chartseries#gradientmidpointtype)|返回或设置区域地图图表系列的中间值的类型。 读/写。|
||[gradientMidpointValue](/javascript/api/excel/excel.chartseries#gradientmidpointvalue)|返回或设置区域地图图表系列的中间值。 读/写。|
||[gradientMinimumColor](/javascript/api/excel/excel.chartseries#gradientminimumcolor)|返回或设置区域地图图表系列的最小值的颜色。 读/写。|
||[gradientMinimumType](/javascript/api/excel/excel.chartseries#gradientminimumtype)|返回或设置区域地图图表系列的最小值的类型。 读/写。|
||[gradientMinimumValue](/javascript/api/excel/excel.chartseries#gradientminimumvalue)|返回或设置区域地图图表系列的最小值。 读/写。|
||[gradientStyle](/javascript/api/excel/excel.chartseries#gradientstyle)|返回或设置区域地图图表的系列渐变样式。 读/写。|
||[invertColor](/javascript/api/excel/excel.chartseries#invertcolor)|返回或设置系列中负数据点的填充颜色。 读/写。|
||[parentLabelStrategy](/javascript/api/excel/excel.chartseries#parentlabelstrategy)|返回或设置树状图的系列父标签策略区域。 读/写。|
||[binOptions](/javascript/api/excel/excel.chartseries#binoptions)|封装直方图和排列图的容器选项。 只读。|
||[boxwhiskerOptions](/javascript/api/excel/excel.chartseries#boxwhiskeroptions)|封装箱形图的选项。 只读。|
||[mapOptions](/javascript/api/excel/excel.chartseries#mapoptions)|封装区域地图图表的选项。 只读。|
||[xErrorBars](/javascript/api/excel/excel.chartseries#xerrorbars)|表示图表系列的误差线对象。|
||[yErrorBars](/javascript/api/excel/excel.chartseries#yerrorbars)|表示图表系列的误差线对象。|
||[showConnectorLines](/javascript/api/excel/excel.chartseries#showconnectorlines)|指定是否在瀑布图中显示连接线。 读/写。|
||[showLeaderLines](/javascript/api/excel/excel.chartseries#showleaderlines)|指定是否在系列中显示每个数据标签的引导线。 读/写。|
||[splitValue](/javascript/api/excel/excel.chartseries#splitvalue)|返回或设置复合饼图或复合条饼图中分隔两部分的阈值。 读/写。|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[linkNumberFormat](/javascript/api/excel/excel.charttrendlinelabel#linknumberformat)|布尔值，表示数字格式是否链接到单元格（以便在单元格中更改时标签中的数字格式会发生改变）。|
|[ColumnProperties](/javascript/api/excel/excel.columnproperties)|[address](/javascript/api/excel/excel.columnproperties#address)|表示`address`属性。|
||[addressLocal](/javascript/api/excel/excel.columnproperties#addresslocal)|表示`addressLocal`属性。|
||[columnIndex](/javascript/api/excel/excel.columnproperties#columnindex)|表示`columnIndex`属性。|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[getRanges()](/javascript/api/excel/excel.conditionalformat#getranges--)|返回将为其应用条件格式的 RangeAreas，它包含一个或多个矩形区域。 只读。|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[getInvalidCells()](/javascript/api/excel/excel.datavalidation#getinvalidcells--)|返回包含一个或多个矩形区域的 RangeAreas，它具有无效单元格值。 如果所有单元格值都有效，则此函数将引发 ItemNotFound 错误。|
||[getInvalidCellsOrNullObject()](/javascript/api/excel/excel.datavalidation#getinvalidcellsornullobject--)|返回包含一个或多个矩形区域的 RangeAreas，它具有无效单元格值。 如果所有单元格值都有效，则此函数将返回 null。|
|[FilterCriteria](/javascript/api/excel/excel.filtercriteria)|[subField](/javascript/api/excel/excel.filtercriteria#subfield)|筛选器使用该属性对 richvalue 执行丰富的筛选。|
|[GeometricShape](/javascript/api/excel/excel.geometricshape)|[load(propertyNames?: string \| string[])](/javascript/api/excel/excel.geometricshape#load-propertynames-)|将命令加入队列以加载对象的指定属性。 在读取属性之前，你必须调用“context.sync()”。|
||[id](/javascript/api/excel/excel.geometricshape#id)|返回形状标识符。 只读。|
||[shape](/javascript/api/excel/excel.geometricshape#shape)|返回几何形状的形状对象。 只读。|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getCount()](/javascript/api/excel/excel.groupshapecollection#getcount--)|返回形状组中的形状数量。 只读。|
||[getItem(key: string)](/javascript/api/excel/excel.groupshapecollection#getitem-key-)|按名称或 ID 获取形状。|
||[getItemAt(index: number)](/javascript/api/excel/excel.groupshapecollection#getitemat-index-)|根据其在集合中的位置获取形状。|
||[load(option?: OfficeExtension.LoadOption)](/javascript/api/excel/excel.groupshapecollection#load-option-)||
||[load(propertyNames?: string \| string[])](/javascript/api/excel/excel.groupshapecollection#load-propertynames-)|将命令加入队列以加载对象的指定属性。 在读取属性之前，你必须调用“context.sync()”。|
||[items](/javascript/api/excel/excel.groupshapecollection#items)|获取此集合中已加载的子项。|
|[HeaderFooter](/javascript/api/excel/excel.headerfooter)|[centerFooter](/javascript/api/excel/excel.headerfooter#centerfooter)|获取或设置工作表的中心页脚。|
||[centerHeader](/javascript/api/excel/excel.headerfooter#centerheader)|获取或设置工作表的中心页眉。|
||[LeftFooter](/javascript/api/excel/excel.headerfooter#leftfooter)|获取或设置工作表的左侧页脚。|
||[leftHeader](/javascript/api/excel/excel.headerfooter#leftheader)|获取或设置工作表的左侧页眉。|
||[load(propertyNames?: string \| string[])](/javascript/api/excel/excel.headerfooter#load-propertynames-)|将命令加入队列以加载对象的指定属性。 在读取属性之前，你必须调用“context.sync()”。|
||[rightFooter](/javascript/api/excel/excel.headerfooter#rightfooter)|获取或设置工作表的右侧页脚。|
||[rightHeader](/javascript/api/excel/excel.headerfooter#rightheader)|获取或设置工作表的右侧页眉。|
|[HeaderFooterGroup](/javascript/api/excel/excel.headerfootergroup)|[load(propertyNames?: string \| string[])](/javascript/api/excel/excel.headerfootergroup#load-propertynames-)|将命令加入队列以加载对象的指定属性。 在读取属性之前，你必须调用“context.sync()”。|
||[defaultForAllPages](/javascript/api/excel/excel.headerfootergroup#defaultforallpages)|常规页眉/页脚，除非指定偶数页/奇数页或首页，否则适用于所有页面，|
||[evenPages](/javascript/api/excel/excel.headerfootergroup#evenpages)|用于偶数页的页眉/页脚，需要为奇数页指定奇数页页眉/页脚。|
||[firstPage](/javascript/api/excel/excel.headerfootergroup#firstpage)|首页的页眉/页脚，为所有其他页使用常规或偶数页/奇数页页眉/页脚。|
||[oddPages](/javascript/api/excel/excel.headerfootergroup#oddpages)|用于奇数页的页眉/页脚，需要为偶数页指定偶数页页眉/页脚。|
||[state](/javascript/api/excel/excel.headerfootergroup#state)|获取或设置所设置的页眉/页脚的状态。 有关详细信息，请参阅 Excel.HeaderFooterState。|
||[useSheetMargins](/javascript/api/excel/excel.headerfootergroup#usesheetmargins)|获取或设置一个标记，指示页眉/页脚是否与工作表的页面布局选项中设置的页边距对齐。|
||[useSheetScale](/javascript/api/excel/excel.headerfootergroup#usesheetscale)|获取或设置一个标记，指示是否应按照工作表的页面布局选项中设置的页面缩放百分比来缩放页眉/页脚。|
|[Image](/javascript/api/excel/excel.image)|[load(propertyNames?: string \| string[])](/javascript/api/excel/excel.image#load-propertynames-)|将命令加入队列以加载对象的指定属性。 在读取属性之前，你必须调用“context.sync()”。|
||[format](/javascript/api/excel/excel.image#format)|返回图像的格式。 只读。|
||[id](/javascript/api/excel/excel.image#id)|表示图像对象的形状标识符。 只读。|
||[shape](/javascript/api/excel/excel.image#shape)|返回与图像关联的形状对象。 只读。|
|[IterativeCalculation](/javascript/api/excel/excel.iterativecalculation)|[enabled](/javascript/api/excel/excel.iterativecalculation#enabled)|如果 Excel 使用迭代来处理循环引用，则为 True。|
||[load(propertyNames?: string \| string[])](/javascript/api/excel/excel.iterativecalculation#load-propertynames-)|将命令加入队列以加载对象的指定属性。 在读取属性之前，你必须调用“context.sync()”。|
||[maxChange](/javascript/api/excel/excel.iterativecalculation#maxchange)|返回或设置 Excel 处理循环引用时迭代之间的最大变化值。|
||[maxIteration](/javascript/api/excel/excel.iterativecalculation#maxiteration)|返回或设置 Excel 处理循环引用的最大迭代次数。|
|[Line](/javascript/api/excel/excel.line)|[beginArrowheadLength](/javascript/api/excel/excel.line#beginarrowheadlength)|表示指定线条始端的箭头长度。|
||[beginArrowheadStyle](/javascript/api/excel/excel.line#beginarrowheadstyle)|表示指定线条始端的箭头样式。|
||[beginArrowheadWidth](/javascript/api/excel/excel.line#beginarrowheadwidth)|表示指定线条始端的箭头宽度。|
||[connectBeginShape(shape: Excel.Shape, connectionSite: number)](/javascript/api/excel/excel.line#connectbeginshape-shape--connectionsite-)|将指定连接线的始端附加到指定形状。|
||[connectEndShape(shape: Excel.Shape, connectionSite: number)](/javascript/api/excel/excel.line#connectendshape-shape--connectionsite-)|将指定连接线的末端附加到指定形状。|
||[connectorType](/javascript/api/excel/excel.line#connectortype)|表示线条的连接器类型。|
||[disconnectBeginShape()](/javascript/api/excel/excel.line#disconnectbeginshape--)|使指定连接线的始端与形状脱离。|
||[disconnectEndShape()](/javascript/api/excel/excel.line#disconnectendshape--)|使指定连接线的末端与形状脱离。|
||[endArrowheadLength](/javascript/api/excel/excel.line#endarrowheadlength)|表示指定线条末端的箭头长度。|
||[endArrowheadStyle](/javascript/api/excel/excel.line#endarrowheadstyle)|表示指定线条末端的箭头样式。|
||[endArrowheadWidth](/javascript/api/excel/excel.line#endarrowheadwidth)|表示指定线条末端的箭头宽度。|
||[load(propertyNames?: string \| string[])](/javascript/api/excel/excel.line#load-propertynames-)|将命令加入队列以加载对象的指定属性。 在读取属性之前，你必须调用“context.sync()”。|
||[beginConnectedShape](/javascript/api/excel/excel.line#beginconnectedshape)|表示指定线条始端所附加到的形状。 只读。|
||[beginConnectedSite](/javascript/api/excel/excel.line#beginconnectedsite)|表示连接线始端所连接的连接站点。 只读。 当线条的始端没有附加到任何形状时，返回 null。|
||[endConnectedShape](/javascript/api/excel/excel.line#endconnectedshape)|表示指定线条末端所附加到的形状。 只读。|
||[endConnectedSite](/javascript/api/excel/excel.line#endconnectedsite)|表示连接线末端所连接的连接站点。 只读。 当线条的末端没有附加到任何形状时，返回 null。|
||[id](/javascript/api/excel/excel.line#id)|表示形状标识符。 只读。|
||[isBeginConnected](/javascript/api/excel/excel.line#isbeginconnected)|指定指定线条的始端是否连接到形状。 只读。|
||[isEndConnected](/javascript/api/excel/excel.line#isendconnected)|指定指定线条的末端是否连接到形状。 只读。|
||[shape](/javascript/api/excel/excel.line#shape)|返回与线条关联的形状对象。 只读。|
|[PageBreak](/javascript/api/excel/excel.pagebreak)|[delete()](/javascript/api/excel/excel.pagebreak#delete--)|删除分页符对象。|
||[getCellAfterBreak()](/javascript/api/excel/excel.pagebreak#getcellafterbreak--)|获取分页符后的第一个单元格。|
||[load(propertyNames?: string \| string[])](/javascript/api/excel/excel.pagebreak#load-propertynames-)|将命令加入队列以加载对象的指定属性。 在读取属性之前，你必须调用“context.sync()”。|
||[columnIndex](/javascript/api/excel/excel.pagebreak#columnindex)|表示分页符的列索引|
||[rowIndex](/javascript/api/excel/excel.pagebreak#rowindex)|表示分页符的行索引|
|[PageBreakCollection](/javascript/api/excel/excel.pagebreakcollection)|[add(pageBreakRange: Range \| string)](/javascript/api/excel/excel.pagebreakcollection#add-pagebreakrange-)|在指定区域的左上角单元格之前添加分页符。|
||[getCount()](/javascript/api/excel/excel.pagebreakcollection#getcount--)|获取集合中的分页符数量。|
||[getItem(index: number)](/javascript/api/excel/excel.pagebreakcollection#getitem-index-)|通过索引获取分页符对象。|
||[load(option?: OfficeExtension.LoadOption)](/javascript/api/excel/excel.pagebreakcollection#load-option-)||
||[load(propertyNames?: string \| string[])](/javascript/api/excel/excel.pagebreakcollection#load-propertynames-)|将命令加入队列以加载对象的指定属性。 在读取属性之前，你必须调用“context.sync()”。|
||[items](/javascript/api/excel/excel.pagebreakcollection#items)|获取此集合中已加载的子项。|
||[removePageBreaks()](/javascript/api/excel/excel.pagebreakcollection#removepagebreaks--)|重置集合中的所有手动分页符。|
|[PageLayout](/javascript/api/excel/excel.pagelayout)|[blackAndWhite](/javascript/api/excel/excel.pagelayout#blackandwhite)|获取或设置工作表的黑白打印选项。|
||[bottomMargin](/javascript/api/excel/excel.pagelayout#bottommargin)|获取或设置要用于打印的工作表的底部页边距（以磅为单位）。|
||[centerHorizontally](/javascript/api/excel/excel.pagelayout#centerhorizontally)|获取或设置工作表的中心水平标记。 此标记确定在打印时是否水平居中工作表。|
||[centerVertically](/javascript/api/excel/excel.pagelayout#centervertically)|获取或设置工作表的中心垂直标记。 此标记确定在打印时是否垂直居中工作表。|
||[draftMode](/javascript/api/excel/excel.pagelayout#draftmode)|获取或设置工作表的草稿模式选项。 如果为 True，则将打印没有图形的工作表。|
||[firstPageNumber](/javascript/api/excel/excel.pagelayout#firstpagenumber)|获取或设置要打印的工作表的首页页码。 Null 值表示“自动”页码编号。|
||[footerMargin](/javascript/api/excel/excel.pagelayout#footermargin)|获取或设置在打印时使用的工作表的页脚边距（以磅为单位）。|
||[getPrintArea()](/javascript/api/excel/excel.pagelayout#getprintarea--)|获取包含一个或多个矩形区域的 RangeAreas 对象，它表示工作表的打印区域。 如果没有打印区域，则将引发 ItemNotFound 错误。|
||[getPrintAreaOrNullObject()](/javascript/api/excel/excel.pagelayout#getprintareaornullobject--)|获取包含一个或多个矩形区域的 RangeAreas 对象，它表示工作表的打印区域。 如果没有打印区域，则将返回 null 对象。|
||[getPrintTitleColumns()](/javascript/api/excel/excel.pagelayout#getprinttitlecolumns--)|获取表示标题列的 Range 对象。|
||[getPrintTitleColumnsOrNullObject()](/javascript/api/excel/excel.pagelayout#getprinttitlecolumnsornullobject--)|获取表示标题列的 Range 对象。 如果未设置，则将返回 null 对象。|
||[getPrintTitleRows()](/javascript/api/excel/excel.pagelayout#getprinttitlerows--)|获取表示标题行的 Range 对象。|
||[getPrintTitleRowsOrNullObject()](/javascript/api/excel/excel.pagelayout#getprinttitlerowsornullobject--)|获取表示标题行的 Range 对象。 如果未设置，则将返回 null 对象。|
||[headerMargin](/javascript/api/excel/excel.pagelayout#headermargin)|获取或设置在打印时使用的工作表的页眉边距（以磅为单位）。|
||[leftMargin](/javascript/api/excel/excel.pagelayout#leftmargin)|获取或设置在打印时使用的工作表的左边距（以磅为单位）。|
||[load(propertyNames?: string \| string[])](/javascript/api/excel/excel.pagelayout#load-propertynames-)|将命令加入队列以加载对象的指定属性。 在读取属性之前，你必须调用“context.sync()”。|
||[orientation](/javascript/api/excel/excel.pagelayout#orientation)|获取或设置工作表的页面方向。|
||[paperSize](/javascript/api/excel/excel.pagelayout#papersize)|获取或设置工作表的页面纸张大小。|
||[printComments](/javascript/api/excel/excel.pagelayout#printcomments)|获取或设置在打印时是否应该显示工作表的批注。|
||[printErrors](/javascript/api/excel/excel.pagelayout#printerrors)|获取或设置工作表的打印错误选项。|
||[printGridlines](/javascript/api/excel/excel.pagelayout#printgridlines)|获取或设置工作表的打印网格线标记。 此标记确定是否打印网格线。|
||[printHeadings](/javascript/api/excel/excel.pagelayout#printheadings)|获取或设置工作表的打印标题标记。 此标记确定是否打印标题。|
||[printOrder](/javascript/api/excel/excel.pagelayout#printorder)|获取或设置工作表的页面打印顺序选项。 它指定用于处理打印页码的顺序。|
||[headersFooters](/javascript/api/excel/excel.pagelayout#headersfooters)|工作表的页眉和页脚配置。|
||[rightMargin](/javascript/api/excel/excel.pagelayout#rightmargin)|获取或设置在打印时使用的工作表的右边距（以磅为单位）。|
||[setPrintArea(printArea: Range \| RangeAreas \| string)](/javascript/api/excel/excel.pagelayout#setprintarea-printarea-)|设置工作表的打印区域。|
||[setPrintMargins(unit: Excel.PrintMarginUnit, marginOptions: Excel.PageLayoutMarginOptions)](/javascript/api/excel/excel.pagelayout#setprintmargins-unit--marginoptions-)|设置带单位的工作表的页边距。|
||[setPrintMargins(unitString: "Points" \| "Inches" \| "Centimeters", marginOptions: Excel.PageLayoutMarginOptions)](/javascript/api/excel/excel.pagelayout#setprintmargins-unitstring--marginoptions-)|设置带单位的工作表的页边距。|
||[setPrintTitleColumns(printTitleColumns: Range \| string)](/javascript/api/excel/excel.pagelayout#setprinttitlecolumns-printtitlecolumns-)|设置列，这些列包含要在打印的工作表的每页左侧重复的单元格。|
||[setPrintTitleRows(printTitleRows: Range \| string)](/javascript/api/excel/excel.pagelayout#setprinttitlerows-printtitlerows-)|设置行，这些行包含要在打印的工作表的每页顶部重复的单元格。|
||[topMargin](/javascript/api/excel/excel.pagelayout#topmargin)|获取或设置在打印时使用的工作表的上边距（以磅为单位）。|
||[zoom](/javascript/api/excel/excel.pagelayout#zoom)|获取或设置工作表的打印缩放选项。|
|[PageLayoutMarginOptions](/javascript/api/excel/excel.pagelayoutmarginoptions)|[bottom](/javascript/api/excel/excel.pagelayoutmarginoptions#bottom)|表示要在打印时使用的页面布局下边距（使用指定的单位）。|
||[footer](/javascript/api/excel/excel.pagelayoutmarginoptions#footer)|表示要在打印时使用的页面布局页脚边距（使用指定的单位）。|
||[header](/javascript/api/excel/excel.pagelayoutmarginoptions#header)|表示要在打印时使用的页面布局页眉边距（使用指定的单位）。|
||[left](/javascript/api/excel/excel.pagelayoutmarginoptions#left)|表示要在打印时使用的页面布局左边距（使用指定的单位）。|
||[right](/javascript/api/excel/excel.pagelayoutmarginoptions#right)|表示要在打印时使用的页面布局右边距（使用指定的单位）。|
||[top](/javascript/api/excel/excel.pagelayoutmarginoptions#top)|表示要在打印时使用的页面布局上边距（使用指定的单位）。|
|[PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions)|[horizontalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#horizontalfittopages)|水平放置的页数。 如果使用百分比缩放，则此值可以为 null。|
||[scale](/javascript/api/excel/excel.pagelayoutzoomoptions#scale)|打印页面缩放值可以介于 10 至 400 之间。 如果已指定适应页面高度或宽度，则此值可以为 null。|
||[verticalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#verticalfittopages)|垂直放置的页数。 如果使用百分比缩放，则此值可以为 null。|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[sortByValues(sortBy: Excel.SortBy, valuesHierarchy: Excel.DataPivotHierarchy, pivotItemScope?: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotfield#sortbyvalues-sortby--valueshierarchy--pivotitemscope-)|按给定范围中的指定值对 PivotField 进行排序。 该范围定义将使用哪些特定值进行排序|
||[sortByValues(sortByString: "Ascending" \| "Descending", valuesHierarchy: Excel.DataPivotHierarchy, pivotItemScope?: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotfield#sortbyvalues-sortbystring--valueshierarchy--pivotitemscope-)|按给定范围中的指定值对 PivotField 进行排序。 该范围定义将使用哪些特定值进行排序|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[autoFormat](/javascript/api/excel/excel.pivotlayout#autoformat)|指定是否在刷新或移动字段时自动进行格式化|
||[getDataHierarchy(cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#getdatahierarchy-cell-)|获取 DataHierarchy，它用于计算数据透视表中指定区域内的值。|
||[getPivotItems(axis: Excel.PivotAxis, cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#getpivotitems-axis--cell-)|从构成数据透视表中指定区域内的值的轴获取 PivotItems。|
||[getPivotItems(axisString: "Unknown" \| "Row" \| "Column" \| "Data" \| "Filter", cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#getpivotitems-axisstring--cell-)|从构成数据透视表中指定区域内的值的轴获取 PivotItems。|
||[preserveFormatting](/javascript/api/excel/excel.pivotlayout#preserveformatting)|指定是否在通过透视、排序或更改页面字段项等操作来刷新或重新计算报表时保留格式。|
||[setAutoSortOnCell(cell: Range \| string, sortBy: Excel.SortBy)](/javascript/api/excel/excel.pivotlayout#setautosortoncell-cell--sortby-)|将数据透视表设置为使用指定的单元格设置自动排序，以自动选择排序的所有条件和上下文。 这与从 UI 应用自动排序的行为相同。|
||[setAutoSortOnCell(cell: Range \| string, sortByString: "Ascending" \| "Descending")](/javascript/api/excel/excel.pivotlayout#setautosortoncell-cell--sortbystring-)|将数据透视表设置为使用指定的单元格设置自动排序，以自动选择排序的所有条件和上下文。 这与从 UI 应用自动排序的行为相同。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[enableDataValueEditing](/javascript/api/excel/excel.pivottable#enabledatavalueediting)|指定数据透视表是否允许用户编辑数据体中的值。|
||[useCustomSortLists](/javascript/api/excel/excel.pivottable#usecustomsortlists)|指定排序时，数据透视表是否使用自定义列表。|
|[Range](/javascript/api/excel/excel.range)|[autoFill(destinationRange: Range \| string, autoFillType?: Excel.AutoFillType)](/javascript/api/excel/excel.range#autofill-destinationrange--autofilltype-)|填充区域从当前区域到目标区域。|
||[autoFill(destinationRange: Range \| string, autoFillType?: "FillDefault" \| "FillCopy" \| "FillSeries" \| "FillFormats" \| "FillValues" \| "FillDays" \| "FillWeekdays" \| "FillMonths" \| "FillYears" \| "LinearTrend" \| "GrowthTrend" \| "FlashFill")](/javascript/api/excel/excel.range#autofill-destinationrange--autofilltypestring-)|填充区域从当前区域到目标区域。|
||[convertDataTypeToText()](/javascript/api/excel/excel.range#convertdatatypetotext--)|将具有数据类型的区域单元格转换为文本。|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.range#converttolinkeddatatype-serviceid--languageculture-)|将区域单元格转换为工作表中的链接数据类型。|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-)|将单元格数据或格式从源区域或 RangeAreas 复制到当前区域。|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyTypeString?: "All" \| "Formulas" \| "Values" \| "Formats", skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytypestring--skipblanks--transpose-)|将单元格数据或格式从源区域或 RangeAreas 复制到当前区域。|
||[find(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#find-text--criteria-)|根据指定的条件查找给定的字符串。|
||[findOrNullObject(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#findornullobject-text--criteria-)|根据指定的条件查找给定的字符串。|
||[flashFill()](/javascript/api/excel/excel.range#flashfill--)|对当前区域进行快速填充。快速填充在感知到模式时可自动填充数据，因此该区域必须是单列区域且周围有数据以便查找模式。|
||[getCellProperties(cellPropertiesLoadOptions: CellPropertiesLoadOptions)](/javascript/api/excel/excel.range#getcellproperties-cellpropertiesloadoptions-)|返回一个 2D 数组，其中封装了每个单元格的字体、填充、边框、对齐方式和其他属性数据。|
||[getColumnProperties(columnPropertiesLoadOptions: ColumnPropertiesLoadOptions)](/javascript/api/excel/excel.range#getcolumnproperties-columnpropertiesloadoptions-)|返回一个一维数组，其中封装了每个列的字体、填充、边框、对齐方式和其他属性数据。  对于给定列中每个单元格不一致的属性，将返回 null。|
||[getRowProperties(rowPropertiesLoadOptions: RowPropertiesLoadOptions)](/javascript/api/excel/excel.range#getrowproperties-rowpropertiesloadoptions-)|返回一个一维数组，其中封装了每个行的字体、填充、边框、对齐方式和其他属性数据。  对于给定行中每个单元格不一致的属性，将返回 null。|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-)|获取包含一个或多个矩形区域的 RangeAreas 对象，它表示匹配指定类型和值的所有单元格。|
||[getSpecialCells(cellTypeString: "ConditionalFormats" \| "DataValidations" \| "Blanks" \| "Constants" \| "Formulas" \| "SameConditionalFormat" \| "SameDataValidation" \| "Visible", cellValueType?: "All" \| "Errors" \| "ErrorsLogical" \| "ErrorsNumbers" \| "ErrorsText" \| "ErrorsLogicalNumber" \| "ErrorsLogicalText" \| "ErrorsNumberText" \| "Logical" \| "LogicalNumbers" \| "LogicalText" \| "LogicalNumbersText" \| "Numbers" \| "NumbersText" \| "Text")](/javascript/api/excel/excel.range#getspecialcells-celltypestring--cellvaluetype-)|获取包含一个或多个矩形区域的 RangeAreas 对象，它表示匹配指定类型和值的所有单元格。|
||[getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#getspecialcellsornullobject-celltype--cellvaluetype-)|获取包含一个或多个区域的 RangeAreas 对象，它表示匹配指定类型和值的所有单元格。|
||[getSpecialCellsOrNullObject(cellTypeString: "ConditionalFormats" \| "DataValidations" \| "Blanks" \| "Constants" \| "Formulas" \| "SameConditionalFormat" \| "SameDataValidation" \| "Visible", cellValueType?: "All" \| "Errors" \| "ErrorsLogical" \| "ErrorsNumbers" \| "ErrorsText" \| "ErrorsLogicalNumber" \| "ErrorsLogicalText" \| "ErrorsNumberText" \| "Logical" \| "LogicalNumbers" \| "LogicalText" \| "LogicalNumbersText" \| "Numbers" \| "NumbersText" \| "Text")](/javascript/api/excel/excel.range#getspecialcellsornullobject-celltypestring--cellvaluetype-)|获取包含一个或多个区域的 RangeAreas 对象，它表示匹配指定类型和值的所有单元格。|
||[getTables(fullyContained?: boolean)](/javascript/api/excel/excel.range#gettables-fullycontained-)|获取与区域重叠的限定范围的表格集合。|
||[linkedDataTypeState](/javascript/api/excel/excel.range#linkeddatatypestate)|表示每个单元格的数据类型状态。 只读。|
||[removeDuplicates(columns: number[], includesHeader: boolean)](/javascript/api/excel/excel.range#removeduplicates-columns--includesheader-)|从列指定的区域中删除重复值。|
||[replaceAll(text: string, replacement: string, criteria: Excel.ReplaceCriteria)](/javascript/api/excel/excel.range#replaceall-text--replacement--criteria-)|根据当前区域内指定的条件查找并替换给定的字符串。|
||[setCellProperties(cellPropertiesData: SettableCellProperties[][])](/javascript/api/excel/excel.range#setcellproperties-cellpropertiesdata-)|根据单元格属性的 2D 数组更新区域，它封装了字体、填充、边框、对齐方式等内容。|
||[setColumnProperties(columnPropertiesData: SettableColumnProperties[])](/javascript/api/excel/excel.range#setcolumnproperties-columnpropertiesdata-)|根据列属性的一维数组更新区域，它封装了字体、填充、边框、对齐方式等内容。|
||[setDirty()](/javascript/api/excel/excel.range#setdirty--)|设置下一次重新计算发生时要重新计算的区域。|
||[setRowProperties(rowPropertiesData: SettableRowProperties[])](/javascript/api/excel/excel.range#setrowproperties-rowpropertiesdata-)|根据行属性的一维数组更新区域，它封装了字体、填充、边框、对齐方式等内容。|
|[RangeAreas](/javascript/api/excel/excel.rangeareas)|[calculate()](/javascript/api/excel/excel.rangeareas#calculate--)|计算 RangeAreas 中的所有单元格。|
||[clear(applyTo?: Excel.ClearApplyTo)](/javascript/api/excel/excel.rangeareas#clear-applyto-)|清除包含此 RangeAreas 对象的每个区域的值、格式、填充、边框等。|
||[clear(applyToString?: "All" \| "Formats" \| "Contents" \| "Hyperlinks" \| "RemoveHyperlinks")](/javascript/api/excel/excel.rangeareas#clear-applytostring-)|清除包含此 RangeAreas 对象的每个区域的值、格式、填充、边框等。|
||[convertDataTypeToText()](/javascript/api/excel/excel.rangeareas#convertdatatypetotext--)|将 RangeAreas 中具有数据类型的所有单元格转换为文本。|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.rangeareas#converttolinkeddatatype-serviceid--languageculture-)|将 RangeAreas 中的所有单元格转换为链接数据类型。|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.rangeareas#copyfrom-sourcerange--copytype--skipblanks--transpose-)|将单元格数据或格式从源区域或 RangeAreas 复制到当前 RangeAreas。|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyTypeString?: "All" \| "Formulas" \| "Values" \| "Formats", skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.rangeareas#copyfrom-sourcerange--copytypestring--skipblanks--transpose-)|将单元格数据或格式从源区域或 RangeAreas 复制到当前 RangeAreas。|
||[getEntireColumn()](/javascript/api/excel/excel.rangeareas#getentirecolumn--)|返回表示 RangeAreas 的整个列的 RangeAreas 对象（例如，如果当前 RangeAreas 表示单元格“B4:E11, H2”，它将返回表示列“B:E, H:H”的 RangeAreas）。|
||[getEntireRow()](/javascript/api/excel/excel.rangeareas#getentirerow--)|获取表示 RangeAreas 的整个行的 RangeAreas 对象（例如，如果当前 RangeAreas 表示单元格“B4:E11”，它将返回表示行“4:11”的 RangeAreas）。|
||[getIntersection(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#getintersection-anotherrange-)|返回 RangeAreas 对象，它表示给定区域或 RangeAreas 的交集。 如果未找到任何交集，则将引发 ItemNotFound 错误。|
||[getIntersectionOrNullObject(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#getintersectionornullobject-anotherrange-)|返回 RangeAreas 对象，它表示给定区域或 RangeAreas 的交集。 如果未找到任何交集，将返回 null 对象。|
||[getOffsetRangeAreas(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.rangeareas#getoffsetrangeareas-rowoffset--columnoffset-)|返回 RangeAreas 对象，它按特定的行和列偏移量进行移动。 返回的 RangeAreas 的维度将与原始对象匹配。 如果生成的 RangeAreas 强行超出工作表网格的边界，则将引发错误。|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#getspecialcells-celltype--cellvaluetype-)|返回一个 RangeAreas 对象，它表示匹配指定类型和值的所有单元格。 如果未找到符合条件的特殊单元格，则会引发错误。|
||[getSpecialCells(cellTypeString: "ConditionalFormats" \| "DataValidations" \| "Blanks" \| "Constants" \| "Formulas" \| "SameConditionalFormat" \| "SameDataValidation" \| "Visible", cellValueType?: "All" \| "Errors" \| "ErrorsLogical" \| "ErrorsNumbers" \| "ErrorsText" \| "ErrorsLogicalNumber" \| "ErrorsLogicalText" \| "ErrorsNumberText" \| "Logical" \| "LogicalNumbers" \| "LogicalText" \| "LogicalNumbersText" \| "Numbers" \| "NumbersText" \| "Text")](/javascript/api/excel/excel.rangeareas#getspecialcells-celltypestring--cellvaluetype-)|返回一个 RangeAreas 对象，它表示匹配指定类型和值的所有单元格。 如果未找到符合条件的特殊单元格，则会引发错误。|
||[getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#getspecialcellsornullobject-celltype--cellvaluetype-)|返回一个 RangeAreas 对象，它表示匹配指定类型和值的所有单元格。 如果未找到符合条件的特殊单元格，则返回 null 对象。|
||[getSpecialCellsOrNullObject(cellTypeString: "ConditionalFormats" \| "DataValidations" \| "Blanks" \| "Constants" \| "Formulas" \| "SameConditionalFormat" \| "SameDataValidation" \| "Visible", cellValueType?: "All" \| "Errors" \| "ErrorsLogical" \| "ErrorsNumbers" \| "ErrorsText" \| "ErrorsLogicalNumber" \| "ErrorsLogicalText" \| "ErrorsNumberText" \| "Logical" \| "LogicalNumbers" \| "LogicalText" \| "LogicalNumbersText" \| "Numbers" \| "NumbersText" \| "Text")](/javascript/api/excel/excel.rangeareas#getspecialcellsornullobject-celltypestring--cellvaluetype-)|返回一个 RangeAreas 对象，它表示匹配指定类型和值的所有单元格。 如果未找到符合条件的特殊单元格，则返回 null 对象。|
||[getTables(fullyContained?: boolean)](/javascript/api/excel/excel.rangeareas#gettables-fullycontained-)|返回与此 RangeAreas 对象中的任何区域重叠的限定范围的表格集合。|
||[getUsedRangeAreas(valuesOnly?: boolean)](/javascript/api/excel/excel.rangeareas#getusedrangeareas-valuesonly-)|返回使用的 RangeAreas，它包含 RangeAreas 对象中的各个矩形区域的所有已用区域。|
||[getUsedRangeAreasOrNullObject(valuesOnly?: boolean)](/javascript/api/excel/excel.rangeareas#getusedrangeareasornullobject-valuesonly-)|返回使用的 RangeAreas，它包含 RangeAreas 对象中的各个矩形区域的所有已用区域。|
||[load(propertyNames?: string \| string[])](/javascript/api/excel/excel.rangeareas#load-propertynames-)|将命令加入队列以加载对象的指定属性。 在读取属性之前，你必须调用“context.sync()”。|
||[address](/javascript/api/excel/excel.rangeareas#address)|返回 A1 样式中的 RageAreas 引用。 地址值将包含单元格的每个矩形块的工作表名称（例如“Sheet1!A1:B4, Sheet1!D1:D4”）。 只读。|
||[addressLocal](/javascript/api/excel/excel.rangeareas#addresslocal)|返回用户区域设置中的 RageAreas 引用。 只读。|
||[areaCount](/javascript/api/excel/excel.rangeareas#areacount)|返回包含此 RangeAreas 对象的矩形区域的数量。|
||[areas](/javascript/api/excel/excel.rangeareas#areas)|返回包含此 RangeAreas 对象的矩形区域的集合。|
||[cellCount](/javascript/api/excel/excel.rangeareas#cellcount)|返回 RangeAreas 对象中的单元格数量，即总计各个矩形区域的单元格计数。 如果单元格计数超过 2^31-1 (2,147,483,647)，则返回 -1。 只读。|
||[conditionalFormats](/javascript/api/excel/excel.rangeareas#conditionalformats)|返回与此 RangeAreas 对象中的任何单元格相交的 ConditionalFormats 集合。 只读。|
||[dataValidation](/javascript/api/excel/excel.rangeareas#datavalidation)|返回 RangeAreas 中的所有区域的 dataValidation 对象。|
||[format](/javascript/api/excel/excel.rangeareas#format)|返回一个 rangeFormat 对象，其中封装了 RangeAreas 对象中的所有区域的字体、填充、边框、对齐方式和其他属性。 只读。|
||[isEntireColumn](/javascript/api/excel/excel.rangeareas#isentirecolumn)|指示此 RangeAreas 对象上的所有区域是否表示整列（例如“A:C, Q:Z”）。 只读。|
||[isEntireRow](/javascript/api/excel/excel.rangeareas#isentirerow)|指示此 RangeAreas 对象上的所有区域是否表示整行（例如“1:3, 5:7”）。 只读。|
||[worksheet](/javascript/api/excel/excel.rangeareas#worksheet)|返回当前 RangeAreas 的工作表。 只读。|
||[setDirty()](/javascript/api/excel/excel.rangeareas#setdirty--)|设置要在下一次重新计算时重新进行计算的 RangeAreas。|
||[style](/javascript/api/excel/excel.rangeareas#style)|表示此 RangeAreas 对象中的所有区域的样式。|
||[track()](/javascript/api/excel/excel.rangeareas#track--)|根据文档中的相应更改来跟踪对象，以便进行自动调整。 此调用是 context.trackedObjects.add(thisObject) 的缩写。 如果你在“.sync”调用之间和按顺序执行“.run”批处理之外使用此对象，并且在对象上设置属性或调用方法时出现“InvalidObjectPath”错误，则需要在首次创建对象时为跟踪的对象集合添加对象。|
||[untrack()](/javascript/api/excel/excel.rangeareas#untrack--)|释放与此对象关联的内存（如果先前已跟踪过）。 此调用是 context.trackedObjects.add(thisObject) 的缩写。 拥有许多跟踪对象会降低主机应用程序的速度，因此请在使用完毕后释放所添加的任何对象。 在内存释放生效之前，你需要调用“context.sync()”。|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[tintAndShade](/javascript/api/excel/excel.rangeborder#tintandshade)|返回或设置一个使区域边框的颜色变亮或变暗的双精度数值，该值介于 -1（最暗）与 1（最亮）之间，初始颜色为 0。|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[tintAndShade](/javascript/api/excel/excel.rangebordercollection#tintandshade)|返回或设置一个使区域边框的颜色变亮或变暗的双精度数值，该值介于 -1（最暗）与 1（最亮）之间，初始颜色为 0。|
|[RangeCollection](/javascript/api/excel/excel.rangecollection)|[getCount()](/javascript/api/excel/excel.rangecollection#getcount--)|返回 RangeCollection 中的区域数量。|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangecollection#getitemat-index-)|根据其在 RangeCollection 中的位置返回 Range 对象。|
||[load(option?: OfficeExtension.LoadOption)](/javascript/api/excel/excel.rangecollection#load-option-)||
||[load(propertyNames?: string \| string[])](/javascript/api/excel/excel.rangecollection#load-propertynames-)|将命令加入队列以加载对象的指定属性。 在读取属性之前，你必须调用“context.sync()”。|
||[items](/javascript/api/excel/excel.rangecollection#items)|获取此集合中已加载的子项。|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[pattern](/javascript/api/excel/excel.rangefill#pattern)|获取或设置区域的图案。 有关详细信息，请参阅 Excel.FillPattern。 不支持 LinearGradient 和 RectangularGradient。|
||[patternColor](/javascript/api/excel/excel.rangefill#patterncolor)|设置 HTML 颜色代码，它表示窗体 #RRGGBB（例如“FFA500”）的区域图案颜色或作为已命名的 HTML 颜色（例如“orange”）。|
||[patternTintAndShade](/javascript/api/excel/excel.rangefill#patterntintandshade)|返回或设置一个使区域填充的图案颜色变亮或变暗的双精度数值，该值介于 -1（最暗）与 1（最亮）之间，初始颜色为 0。|
||[tintAndShade](/javascript/api/excel/excel.rangefill#tintandshade)|返回或设置一个使区域填充的颜色变亮或变暗的双精度数值，该值介于 -1（最暗）与 1（最亮）之间，初始颜色为 0。|
|[RangeFont](/javascript/api/excel/excel.rangefont)|[strikethrough](/javascript/api/excel/excel.rangefont#strikethrough)|表示字体的删除线状态。 Null 值表示整个区域没有统一的删除线设置。|
||[subscript](/javascript/api/excel/excel.rangefont#subscript)|表示字体的下标状态。|
||[superscript](/javascript/api/excel/excel.rangefont#superscript)|表示字体的上标状态。|
||[tintAndShade](/javascript/api/excel/excel.rangefont#tintandshade)|返回或设置一个使区域字体的颜色变亮或变暗的双精度数值，该值介于 -1（最暗）与 1（最亮）之间，初始颜色为 0。|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[autoIndent](/javascript/api/excel/excel.rangeformat#autoindent)|指示将文本对齐方式设为相等分布时文本是否会自动缩进。|
||[indentLevel](/javascript/api/excel/excel.rangeformat#indentlevel)|0 到 250 之间的一个整数，指示缩进水平。|
||[readingOrder](/javascript/api/excel/excel.rangeformat#readingorder)|区域的读取顺序。|
||[shrinkToFit](/javascript/api/excel/excel.rangeformat#shrinktofit)|指示文本是否自动缩小以适合可用列宽。|
|[RemoveDuplicatesResult](/javascript/api/excel/excel.removeduplicatesresult)|[load(propertyNames?: string \| string[])](/javascript/api/excel/excel.removeduplicatesresult#load-propertynames-)|将命令加入队列以加载对象的指定属性。 在读取属性之前，你必须调用“context.sync()”。|
||[removed](/javascript/api/excel/excel.removeduplicatesresult#removed)|由操作删除的重复行数。|
||[uniqueRemaining](/javascript/api/excel/excel.removeduplicatesresult#uniqueremaining)|所生成的区域中存在的剩余唯一行数。|
|[ReplaceCriteria](/javascript/api/excel/excel.replacecriteria)|[completeMatch](/javascript/api/excel/excel.replacecriteria#completematch)|指定匹配必须是完整匹配还是部分匹配。 默认值为 false（部分）。|
||[matchCase](/javascript/api/excel/excel.replacecriteria#matchcase)|指定匹配是否区分大小写。 默认值为 false（不区分大小写）。|
|[RowProperties](/javascript/api/excel/excel.rowproperties)|[address](/javascript/api/excel/excel.rowproperties#address)|表示`address`属性。|
||[addressLocal](/javascript/api/excel/excel.rowproperties#addresslocal)|表示`addressLocal`属性。|
||[rowIndex](/javascript/api/excel/excel.rowproperties#rowindex)|表示`rowIndex`属性。|
|[SearchCriteria](/javascript/api/excel/excel.searchcriteria)|[completeMatch](/javascript/api/excel/excel.searchcriteria#completematch)|指定匹配必须是完整匹配还是部分匹配。 默认值为 false（部分）。|
||[matchCase](/javascript/api/excel/excel.searchcriteria#matchcase)|指定匹配是否区分大小写。 默认值为 false（不区分大小写）。|
||[searchDirection](/javascript/api/excel/excel.searchcriteria#searchdirection)|指定搜索方向。 默认值为向前。 请参阅 Excel.SearchDirection。|
|[SettableCellProperties](/javascript/api/excel/excel.settablecellproperties)|[format](/javascript/api/excel/excel.settablecellproperties#format)|表示`format`属性。|
||[hyperlink](/javascript/api/excel/excel.settablecellproperties#hyperlink)|表示`hyperlink`属性。|
||[style](/javascript/api/excel/excel.settablecellproperties#style)|表示`style`属性。|
|[SettableColumnProperties](/javascript/api/excel/excel.settablecolumnproperties)|[columnHidden](/javascript/api/excel/excel.settablecolumnproperties#columnhidden)|表示`columnHidden`属性。|
||[format](/javascript/api/excel/excel.settablecolumnproperties#format)|表示`format`属性。|
|[SettableRowProperties](/javascript/api/excel/excel.settablerowproperties)|[格式：Excel.CellPropertiesFormat](/javascript/api/excel/excel.settablerowproperties#format)|表示`format`属性。|
||[rowHidden](/javascript/api/excel/excel.settablerowproperties#rowhidden)|表示`rowHidden`属性。|
|[Shape](/javascript/api/excel/excel.shape)|[altTextDescription](/javascript/api/excel/excel.shape#alttextdescription)|返回或设置形状对象的可选说明文本。|
||[altTextTitle](/javascript/api/excel/excel.shape#alttexttitle)|返回或设置形状对象的可选标题文本。|
||[delete()](/javascript/api/excel/excel.shape#delete--)|从工作表删除形状。|
||[geometricShapeType](/javascript/api/excel/excel.shape#geometricshapetype)|表示此几何形状的几何形状类型。 有关详细信息，请参阅 Excel.GeometricShapeType。 如果形状类型不是“GeometricShape”，返回 NULL。|
||[getAsImage(format: Excel.PictureFormat)](/javascript/api/excel/excel.shape#getasimage-format-)|将形状转换为图像并将图像返回为 base64 编码字符串。 DPI 为 96。 仅支持格式 `Excel.PictureFormat.BMP`、`Excel.PictureFormat.PNG`、`Excel.PictureFormat.JPEG` 和 `Excel.PictureFormat.GIF`。|
||[getAsImage(formatString: "UNKNOWN" \| "BMP" \| "JPEG" \| "GIF" \| "PNG" \| "SVG")](/javascript/api/excel/excel.shape#getasimage-formatstring-)|将形状转换为图像并将图像返回为 base64 编码字符串。 DPI 为 96。 仅支持格式 `Excel.PictureFormat.BMP`、`Excel.PictureFormat.PNG`、`Excel.PictureFormat.JPEG` 和 `Excel.PictureFormat.GIF`。|
||[height](/javascript/api/excel/excel.shape#height)|表示形状的高度（以磅为单位）。|
||[incrementLeft(increment: number)](/javascript/api/excel/excel.shape#incrementleft-increment-)|以指定磅数水平移动形状。|
||[incrementRotation(increment: number)](/javascript/api/excel/excel.shape#incrementrotation-increment-)|将形状围绕 z 轴旋转特定度数。|
||[incrementTop(increment: number)](/javascript/api/excel/excel.shape#incrementtop-increment-)|以指定磅数垂直移动形状。|
||[left](/javascript/api/excel/excel.shape#left)|从形状左侧到工作表左侧的距离（以磅为单位）。|
||[load(propertyNames?: string \| string[])](/javascript/api/excel/excel.shape#load-propertynames-)|将命令加入队列以加载对象的指定属性。 在读取属性之前，你必须调用“context.sync()”。|
||[lockAspectRatio](/javascript/api/excel/excel.shape#lockaspectratio)|指定此形状的纵横比是否锁定。|
||[名称](/javascript/api/excel/excel.shape#name)|表示形状的名称。|
||[connectionSiteCount](/javascript/api/excel/excel.shape#connectionsitecount)|返回此形状上的连接站点数。 只读。|
||[fill](/javascript/api/excel/excel.shape#fill)|返回此形状的填充格式。 只读。|
||[geometricShape](/javascript/api/excel/excel.shape#geometricshape)|返回与形状关联的几何形状。 如果形状类型不是“GeometricShape”，则会引发错误。|
||[组](/javascript/api/excel/excel.shape#group)|返回与形状关联的形状组。 如果形状类型不是“GroupShape”，则会引发错误。|
||[id](/javascript/api/excel/excel.shape#id)|表示形状标识符。 只读。|
||[image](/javascript/api/excel/excel.shape#image)|返回与形状关联的图像。 如果形状类型不是“Image”，则会引发错误。|
||[level](/javascript/api/excel/excel.shape#level)|表示指定形状的级别。 例如，级别 0 表示形状不是任何组的一部分，级别 1 表示形状是顶级组的一部分，级别 2 表示形状是顶级子组的一部分。|
||[line](/javascript/api/excel/excel.shape#line)|返回与形状关联的线条。 如果形状类型不是“Line”，则会引发错误。|
||[lineFormat](/javascript/api/excel/excel.shape#lineformat)|返回此形状的线条格式。 只读。|
||[onActivated](/javascript/api/excel/excel.shape#onactivated)|当激活形状时发生此事件。|
||[onDeactivated](/javascript/api/excel/excel.shape#ondeactivated)|当停用形状时发生此事件。|
||[parentGroup](/javascript/api/excel/excel.shape#parentgroup)|表示此形状的父组。|
||[textFrame](/javascript/api/excel/excel.shape#textframe)|返回此形状的文本框对象。 只读。|
||[type](/javascript/api/excel/excel.shape#type)|返回此形状的类型。 有关详细信息，请参阅 Excel.ShapeType。 只读。|
||[zOrderPosition](/javascript/api/excel/excel.shape#zorderposition)|返回指定形状在 z 顺序中的位置，0 表示顺序堆栈的底部。 只读。|
||[rotation](/javascript/api/excel/excel.shape#rotation)|表示形状的旋转度数。|
||[scaleHeight(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#scaleheight-scalefactor--scaletype--scalefrom-)|按指定因子缩放形状的高度。 对于图像，你可以说明是相对于原始尺寸还是当前尺寸缩放形状。 对于除图片以外的其他形状来说，缩放总是相对于其当前高度而言。|
||[scaleHeight(scaleFactor: number, scaleTypeString: "CurrentSize" \| "OriginalSize", scaleFrom?: "ScaleFromTopLeft" \| "ScaleFromMiddle" \| "ScaleFromBottomRight")](/javascript/api/excel/excel.shape#scaleheight-scalefactor--scaletypestring--scalefrom-)|按指定因子缩放形状的高度。 对于图像，你可以说明是相对于原始尺寸还是当前尺寸缩放形状。 对于除图片以外的其他形状来说，缩放总是相对于其当前高度而言。|
||[scaleWidth(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#scalewidth-scalefactor--scaletype--scalefrom-)|按指定因子缩放形状的宽度。 对于图像，你可以说明是相对于原始尺寸还是当前尺寸缩放形状。 对于除图片以外的其他形状来说，缩放总是相对于其当前宽度而言。|
||[scaleWidth(scaleFactor: number, scaleTypeString: "CurrentSize" \| "OriginalSize", scaleFrom?: "ScaleFromTopLeft" \| "ScaleFromMiddle" \| "ScaleFromBottomRight")](/javascript/api/excel/excel.shape#scalewidth-scalefactor--scaletypestring--scalefrom-)|按指定因子缩放形状的宽度。 对于图像，你可以说明是相对于原始尺寸还是当前尺寸缩放形状。 对于除图片以外的其他形状来说，缩放总是相对于其当前宽度而言。|
||[setZOrder(position: Excel.ShapeZOrder)](/javascript/api/excel/excel.shape#setzorder-position-)|将指定形状沿集合的 z 顺序向上或向下移动，将其移动到其他形状的前面或后面。|
||[setZOrder(positionString: "BringToFront" \| "BringForward" \| "SendToBack" \| "SendBackward")](/javascript/api/excel/excel.shape#setzorder-positionstring-)|将指定形状沿集合的 z 顺序向上或向下移动，将其移动到其他形状的前面或后面。|
||[top](/javascript/api/excel/excel.shape#top)|从形状上边缘到工作表上边缘之间的距离（以磅为单位）。|
||[visible](/javascript/api/excel/excel.shape#visible)|表示此形状的可视性。|
||[width](/javascript/api/excel/excel.shape#width)|表示形状的宽度（以磅为单位）。|
|[ShapeActivatedEventArgs](/javascript/api/excel/excel.shapeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapeactivatedeventargs#shapeid)|获取已激活的形状的 ID。|
||[type](/javascript/api/excel/excel.shapeactivatedeventargs#type)|获取事件的类型。 有关详细信息，请参阅 Excel.EventType。|
||[worksheetId](/javascript/api/excel/excel.shapeactivatedeventargs#worksheetid)|获取其中的形状已启用的工作表的 ID。|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addGeometricShape(geometricShapeType: Excel.GeometricShapeType)](/javascript/api/excel/excel.shapecollection#addgeometricshape-geometricshapetype-)|将几何形状添加到工作表。 返回一个 Shape 对象，该对象代表新图形。|
||[addGeometricShape(geometricShapeTypeString: "LineInverse" \| "Triangle" \| "RightTriangle" \| "Rectangle" \| "Diamond" \| "Parallelogram" \| "Trapezoid" \| "NonIsoscelesTrapezoid" \| "Pentagon" \| "Hexagon" \| "Heptagon" \| "Octagon" \| "Decagon" \| "Dodecagon" \| "Star4" \| "Star5" \| "Star6" \| "Star7" \| "Star8" \| "Star10" \| "Star12" \| "Star16" \| "Star24" \| "Star32" \| "RoundRectangle" \| "Round1Rectangle" \| "Round2SameRectangle" \| "Round2DiagonalRectangle" \| "SnipRoundRectangle" \| "Snip1Rectangle" \| "Snip2SameRectangle" \| "Snip2DiagonalRectangle" \| "Plaque" \| "Ellipse" \| "Teardrop" \| "HomePlate" \| "Chevron" \| "PieWedge" \| "Pie" \| "BlockArc" \| "Donut" \| "NoSmoking" \| "RightArrow" \| "LeftArrow" \| "UpArrow" \| "DownArrow" \| "StripedRightArrow" \| "NotchedRightArrow" \| "BentUpArrow" \| "LeftRightArrow" \| "UpDownArrow" \| "LeftUpArrow" \| "LeftRightUpArrow" \| "QuadArrow" \| "LeftArrowCallout" \| "RightArrowCallout" \| "UpArrowCallout" \| "DownArrowCallout" \| "LeftRightArrowCallout" \| "UpDownArrowCallout" \| "QuadArrowCallout" \| "BentArrow" \| "UturnArrow" \| "CircularArrow" \| "LeftCircularArrow" \| "LeftRightCircularArrow" \| "CurvedRightArrow" \| "CurvedLeftArrow" \| "CurvedUpArrow" \| "CurvedDownArrow" \| "SwooshArrow" \| "Cube" \| "Can" \| "LightningBolt" \| "Heart" \| "Sun" \| "Moon" \| "SmileyFace" \| "IrregularSeal1" \| "IrregularSeal2" \| "FoldedCorner" \| "Bevel" \| "Frame" \| "HalfFrame" \| "Corner" \| "DiagonalStripe" \| "Chord" \| "Arc" \| "LeftBracket" \| "RightBracket" \| "LeftBrace" \| "RightBrace" \| "BracketPair" \| "BracePair" \| "Callout1" \| "Callout2" \| "Callout3" \| "AccentCallout1" \| "AccentCallout2" \| "AccentCallout3" \| "BorderCallout1" \| "BorderCallout2" \| "BorderCallout3" \| "AccentBorderCallout1" \| "AccentBorderCallout2" \| "AccentBorderCallout3" \| "WedgeRectCallout" \| "WedgeRRectCallout" \| "WedgeEllipseCallout" \| "CloudCallout" \| "Cloud" \| "Ribbon" \| "Ribbon2" \| "EllipseRibbon" \| "EllipseRibbon2" \| "LeftRightRibbon" \| "VerticalScroll" \| "HorizontalScroll" \| "Wave" \| "DoubleWave" \| "Plus" \| "FlowChartProcess" \| "FlowChartDecision" \| "FlowChartInputOutput" \| "FlowChartPredefinedProcess" \| "FlowChartInternalStorage" \| "FlowChartDocument" \| "FlowChartMultidocument" \| "FlowChartTerminator" \| "FlowChartPreparation" \| "FlowChartManualInput" \| "FlowChartManualOperation" \| "FlowChartConnector" \| "FlowChartPunchedCard" \| "FlowChartPunchedTape" \| "FlowChartSummingJunction" \| "FlowChartOr" \| "FlowChartCollate" \| "FlowChartSort" \| "FlowChartExtract" \| "FlowChartMerge" \| "FlowChartOfflineStorage" \| "FlowChartOnlineStorage" \| "FlowChartMagneticTape" \| "FlowChartMagneticDisk" \| "FlowChartMagneticDrum" \| "FlowChartDisplay" \| "FlowChartDelay" \| "FlowChartAlternateProcess" \| "FlowChartOffpageConnector" \| "ActionButtonBlank" \| "ActionButtonHome" \| "ActionButtonHelp" \| "ActionButtonInformation" \| "ActionButtonForwardNext" \| "ActionButtonBackPrevious" \| "ActionButtonEnd" \| "ActionButtonBeginning" \| "ActionButtonReturn" \| "ActionButtonDocument" \| "ActionButtonSound" \| "ActionButtonMovie" \| "Gear6" \| "Gear9" \| "Funnel" \| "MathPlus" \| "MathMinus" \| "MathMultiply" \| "MathDivide" \| "MathEqual" \| "MathNotEqual" \| "CornerTabs" \| "SquareTabs" \| "PlaqueTabs" \| "ChartX" \| "ChartStar" \| "ChartPlus")](/javascript/api/excel/excel.shapecollection#addgeometricshape-geometricshapetypestring-)|将几何形状添加到工作表。 返回一个 Shape 对象，该对象代表新图形。|
||[addGroup(values: Array<string \| Shape>)](/javascript/api/excel/excel.shapecollection#addgroup-values-)|在此集合的工作表中对形状的子集进行分组。 返回表示新形状组的 Shape 对象。|
||[addImage(base64ImageString: string)](/javascript/api/excel/excel.shapecollection#addimage-base64imagestring-)|从 base64 编码的字符串创建图像并将其添加到工作表。 返回表示新图片的 Shape 对象。|
||[addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType)](/javascript/api/excel/excel.shapecollection#addline-startleft--starttop--endleft--endtop--connectortype-)|将线条添加到工作表。 返回表示新线条的 Shape 对象。|
||[addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: "Straight" \| "Elbow" \| "Curve")](/javascript/api/excel/excel.shapecollection#addline-startleft--starttop--endleft--endtop--connectortypestring-)|将线条添加到工作表。 返回表示新线条的 Shape 对象。|
||[addTextBox(text?: string)](/javascript/api/excel/excel.shapecollection#addtextbox-text-)|使用提供的文本作为内容，将文本框添加到工作表。 返回表示新文本框的 Shape 对象。|
||[getCount()](/javascript/api/excel/excel.shapecollection#getcount--)|返回工作表中的形状数。 只读。|
||[getItem(key: string)](/javascript/api/excel/excel.shapecollection#getitem-key-)|按名称或 ID 获取形状。|
||[getItemAt(index: number)](/javascript/api/excel/excel.shapecollection#getitemat-index-)|使用其在集合中的位置获取形状。|
||[load(option?: OfficeExtension.LoadOption)](/javascript/api/excel/excel.shapecollection#load-option-)||
||[load(propertyNames?: string \| string[])](/javascript/api/excel/excel.shapecollection#load-propertynames-)|将命令加入队列以加载对象的指定属性。 在读取属性之前，你必须调用“context.sync()”。|
||[items](/javascript/api/excel/excel.shapecollection#items)|获取此集合中已加载的子项。|
|[ShapeDeactivatedEventArgs](/javascript/api/excel/excel.shapedeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapedeactivatedeventargs#shapeid)|获取已停用的形状的 ID。|
||[type](/javascript/api/excel/excel.shapedeactivatedeventargs#type)|获取事件的类型。 有关详细信息，请参阅 Excel.EventType。|
||[worksheetId](/javascript/api/excel/excel.shapedeactivatedeventargs#worksheetid)|获取其中的形状已停用的工作表的 ID。|
|[ShapeFill](/javascript/api/excel/excel.shapefill)|[clear()](/javascript/api/excel/excel.shapefill#clear--)|清除此形状的填充格式。|
||[foregroundColor](/javascript/api/excel/excel.shapefill#foregroundcolor)|表示窗体 #RRGGBB（例如“FFA500”）的形状填充前景色（采用 HTML 颜色格式）或作为已命名的 HTML 颜色（例如“orange”）|
||[load(propertyNames?: string \| string[])](/javascript/api/excel/excel.shapefill#load-propertynames-)|将命令加入队列以加载对象的指定属性。 在读取属性之前，你必须调用“context.sync()”。|
||[type](/javascript/api/excel/excel.shapefill#type)|返回形状的填充类型。 只读。 有关详细信息，请参阅 Excel.ShapeFillType。|
||[setSolidColor(color: string)](/javascript/api/excel/excel.shapefill#setsolidcolor-color-)|将形状的填充格式设置为统一颜色。 这样可将填充类型更改为“Solid”。|
||[transparency](/javascript/api/excel/excel.shapefill#transparency)|返回或设置填充的透明度百分比，取值范围为 0.0（不透明）到 1.0（清晰）之间。 如果形状类型不支持透明度或形状填充透明度不一致（例如使用渐变填充类型），则返回 null。|
|[ShapeFont](/javascript/api/excel/excel.shapefont)|[bold](/javascript/api/excel/excel.shapefont#bold)|表示字体的加粗状态。 如果 TextRange 包含粗体和非粗体文本片段，则返回 null。|
||[color](/javascript/api/excel/excel.shapefont#color)|文本颜色的 HTML 颜色代码表示（例如，#FF0000 表示红色）。 如果 TextRange 包含具有不同颜色的文本片段，则返回 null。|
||[italic](/javascript/api/excel/excel.shapefont#italic)|表示字体的斜体状态。 如果 TextRange 包含斜体和非斜体文本片段，则返回 null。|
||[load(propertyNames?: string \| string[])](/javascript/api/excel/excel.shapefont#load-propertynames-)|将命令加入队列以加载对象的指定属性。 在读取属性之前，你必须调用“context.sync()”。|
||[name](/javascript/api/excel/excel.shapefont#name)|表示字体名称（例如“Calibri”）。 如果文本是复杂脚本或东亚语言，则这是相应的字体名称，否则是拉丁字体名称。|
||[size](/javascript/api/excel/excel.shapefont#size)|表示以磅为单位的字体大小（例如 11）。 如果 TextRange 包含具有不同字体大小的文本片段，则返回 null。|
||[underline](/javascript/api/excel/excel.shapefont#underline)|应用于字体的下划线类型。 如果 TextRange 包含具有不同下划线样式的文本片段，则返回 null。 有关详细信息，请参阅 Excel.ShapeFontUnderlineStyle。|
|[ShapeGroup](/javascript/api/excel/excel.shapegroup)|[load(propertyNames?: string \| string[])](/javascript/api/excel/excel.shapegroup#load-propertynames-)|将命令加入队列以加载对象的指定属性。 在读取属性之前，你必须调用“context.sync()”。|
||[id](/javascript/api/excel/excel.shapegroup#id)|表示形状标识符。 只读。|
||[shape](/javascript/api/excel/excel.shapegroup#shape)|返回与组关联的 Shape 对象。 只读。|
||[shapes](/javascript/api/excel/excel.shapegroup#shapes)|返回 Shape 对象的集合。 只读。|
||[ungroup()](/javascript/api/excel/excel.shapegroup#ungroup--)|取消分组指定形状组中的任何已分组形状。|
|[ShapeLineFormat](/javascript/api/excel/excel.shapelineformat)|[color](/javascript/api/excel/excel.shapelineformat#color)|表示窗体 #RRGGBB（例如“FFA500”）的线条颜色（采用 HTML 颜色格式）或作为已命名的 HTML 颜色（例如“orange”）|
||[dashStyle](/javascript/api/excel/excel.shapelineformat#dashstyle)|表示形状的线条样式。 当线条不可见或短划线样式不一致时，返回 null。 有关详细信息，请参阅 Excel.ShapeLineStyle。|
||[load(propertyNames?: string \| string[])](/javascript/api/excel/excel.shapelineformat#load-propertynames-)|将命令加入队列以加载对象的指定属性。 在读取属性之前，你必须调用“context.sync()”。|
||[style](/javascript/api/excel/excel.shapelineformat#style)|表示形状的线条样式。 当线条不可见或样式不一致时，返回 null。 有关详细信息，请参阅 Excel.ShapeLineStyle。|
||[transparency](/javascript/api/excel/excel.shapelineformat#transparency)|将指定线条的透明度表示为从 0.0（不透明）到 1.0（清晰）的值。 当形状的透明度不一致时，返回 null。|
||[visible](/javascript/api/excel/excel.shapelineformat#visible)|表示形状元素的线条格式是否可见。 当形状的可见性不一致时，返回 null。|
||[weight](/javascript/api/excel/excel.shapelineformat#weight)|表示线条的粗细（以磅为单位）。 当线条不可见或线条粗细不一致时，返回 null。|
|[SortField](/javascript/api/excel/excel.sortfield)|[subField](/javascript/api/excel/excel.sortfield#subfield)|表示子字段，它是要排序的复合值的目标属性名称。|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[getCount()](/javascript/api/excel/excel.stylecollection#getcount--)|获取集合中的样式数量。|
||[getItemAt(index: number)](/javascript/api/excel/excel.stylecollection#getitemat-index-)|根据其在集合中的位置获取样式。|
|[Table](/javascript/api/excel/excel.table)|[autoFilter](/javascript/api/excel/excel.table#autofilter)|表示表格的 AutoFilter 对象。 只读。|
|[TableAddedEventArgs](/javascript/api/excel/excel.tableaddedeventargs)|[source](/javascript/api/excel/excel.tableaddedeventargs#source)|获取事件源。 有关详细信息，请参阅 Excel.EventSource。|
||[tableId](/javascript/api/excel/excel.tableaddedeventargs#tableid)|获取已添加的表格的 ID。|
||[type](/javascript/api/excel/excel.tableaddedeventargs#type)|获取事件的类型。 有关详细信息，请参阅 Excel.EventType。|
||[worksheetId](/javascript/api/excel/excel.tableaddedeventargs#worksheetid)|获取已在其中添加表格的工作表的 ID。|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[details](/javascript/api/excel/excel.tablechangedeventargs#details)|表示更改详情的信息|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onAdded](/javascript/api/excel/excel.tablecollection#onadded)|在工作簿中添加新表格时发生。|
||[onDeleted](/javascript/api/excel/excel.tablecollection#ondeleted)|在工作簿中删除指定的表格时发生。|
|[TableDeletedEventArgs](/javascript/api/excel/excel.tabledeletedeventargs)|[source](/javascript/api/excel/excel.tabledeletedeventargs#source)|指定时间源。 有关详细信息，请参阅 Excel.EventSource。|
||[tableId](/javascript/api/excel/excel.tabledeletedeventargs#tableid)|指定已删除的表格的 ID。|
||[tableName](/javascript/api/excel/excel.tabledeletedeventargs#tablename)|指定已删除的表格的名称。|
||[type](/javascript/api/excel/excel.tabledeletedeventargs#type)|指定事件类型。 有关详细信息，请参阅 Excel.EventType。|
||[worksheetId](/javascript/api/excel/excel.tabledeletedeventargs#worksheetid)|指定已在其内删除表格的工作表的 ID。|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getCount()](/javascript/api/excel/excel.tablescopedcollection#getcount--)|获取集合中的表数量。|
||[getFirst()](/javascript/api/excel/excel.tablescopedcollection#getfirst--)|获取集合中的第一个表格。 集合中的表格按照从上到下、从左到右的顺序排列，因此左上表格是集合中的第一个表格。|
||[getItem(key: string)](/javascript/api/excel/excel.tablescopedcollection#getitem-key-)|按名称或 ID 获取表。|
||[load(option?: OfficeExtension.LoadOption)](/javascript/api/excel/excel.tablescopedcollection#load-option-)||
||[load(propertyNames?: string \| string[])](/javascript/api/excel/excel.tablescopedcollection#load-propertynames-)|将命令加入队列以加载对象的指定属性。 在读取属性之前，你必须调用“context.sync()”。|
||[items](/javascript/api/excel/excel.tablescopedcollection#items)|获取此集合中已加载的子项。|
|[TextFrame](/javascript/api/excel/excel.textframe)|[autoSizeSetting](/javascript/api/excel/excel.textframe#autosizesetting)|获取或设置文本框的自动调整大小设置。 可以将文本框设置为自动调整文本大小以适应文本框，或自动调整文本框大小以适应文本，或者不使用自动调整大小设置。|
||[bottomMargin](/javascript/api/excel/excel.textframe#bottommargin)|表示文本框的下边距（以磅为单位）。|
||[deleteText()](/javascript/api/excel/excel.textframe#deletetext--)|删除文本框中的所有文本。|
||[horizontalAlignment](/javascript/api/excel/excel.textframe#horizontalalignment)|表示文本框的水平对齐方式。 有关详细信息，请参阅 Excel.ShapeTextHorizontalAlignment。|
||[horizontalOverflow](/javascript/api/excel/excel.textframe#horizontaloverflow)|表示文本框的水平溢出行为。 有关详细信息，请参阅 Excel.ShapeTextHorizontalOverflow。|
||[leftMargin](/javascript/api/excel/excel.textframe#leftmargin)|表示文本框的左边距（以磅为单位）。|
||[load(propertyNames?: string \| string[])](/javascript/api/excel/excel.textframe#load-propertynames-)|将命令加入队列以加载对象的指定属性。 在读取属性之前，你必须调用“context.sync()”。|
||[orientation](/javascript/api/excel/excel.textframe#orientation)|表示文本框的文本方向。 有关详细信息，请参阅 Excel.ShapeTextOrientation。|
||[readingOrder](/javascript/api/excel/excel.textframe#readingorder)|表示文本框从左到右或从右到左的读取顺序。 有关详细信息，请参阅 Excel.ShapeTextReadingOrder。|
||[hasText](/javascript/api/excel/excel.textframe#hastext)|指定文本框是否包含文本。|
||[textRange](/javascript/api/excel/excel.textframe#textrange)|表示附加到文本框中形状上的文本，以及用于操作文本的属性和方法。 有关详细信息，请参阅 Excel.TextRange。|
||[rightMargin](/javascript/api/excel/excel.textframe#rightmargin)|表示文本框的右边距（以磅为单位）。|
||[topMargin](/javascript/api/excel/excel.textframe#topmargin)|表示文本框的上边距（以磅为单位）。|
||[verticalAlignment](/javascript/api/excel/excel.textframe#verticalalignment)|表示文本框的垂直对齐方式。 有关详细信息，请参阅 Excel.ShapeTextVerticalAlignment。|
||[verticalOverflow](/javascript/api/excel/excel.textframe#verticaloverflow)|表示文本框的垂直溢出行为。 有关详细信息，请参阅 Excel.ShapeTextVerticalOverflow。|
|[TextRange](/javascript/api/excel/excel.textrange)|[getSubstring(start: number, length?: number)](/javascript/api/excel/excel.textrange#getsubstring-start--length-)|返回给定区域内子字符串的 TextRange 对象。|
||[load(propertyNames?: string \| string[])](/javascript/api/excel/excel.textrange#load-propertynames-)|将命令加入队列以加载对象的指定属性。 在读取属性之前，你必须调用“context.sync()”。|
||[font](/javascript/api/excel/excel.textrange#font)|返回一个 ShapeFont 对象，该对象表示文本范围的字体属性。 只读。|
||[text](/javascript/api/excel/excel.textrange#text)|表示文本范围的纯文本内容。|
|[Workbook](/javascript/api/excel/excel.workbook)|[chartDataPointTrack](/javascript/api/excel/excel.workbook#chartdatapointtrack)|如果工作簿中的所有图表都跟踪它们所附加的实际数据点，则为 True。|
||[getActiveChart()](/javascript/api/excel/excel.workbook#getactivechart--)|获取工作簿中的当前活动图表。 如果没有活动图表，则在调用此语句时将引发异常|
||[getActiveChartOrNullObject()](/javascript/api/excel/excel.workbook#getactivechartornullobject--)|获取工作簿中的当前活动图表。 如果没有活动图表，则返回 null 对象|
||[getIsActiveCollabSession()](/javascript/api/excel/excel.workbook#getisactivecollabsession--)|如果多个用户正在编辑工作簿（共同创作），则为 True。|
||[getSelectedRanges()](/javascript/api/excel/excel.workbook#getselectedranges--)|从工作簿中获取当前选定的一个或多个区域。 与 getSelectedRange() 不同，此方法返回表示所有选定区域的 RangeAreas 对象。|
||[isDirty](/javascript/api/excel/excel.workbook#isdirty)|指定自上次保存以来是否对指定的工作簿进行任何更改。|
||[autoSave](/javascript/api/excel/excel.workbook#autosave)|指定工作簿是否处于自动保存模式。 只读。|
||[calculationEngineVersion](/javascript/api/excel/excel.workbook#calculationengineversion)|返回有关 Excel 计算引擎的版本号。 只读。|
||[onAutoSaveSettingChanged](/javascript/api/excel/excel.workbook#onautosavesettingchanged)|在工作簿上更改“自动保存”设置时发生。|
||[previouslySaved](/javascript/api/excel/excel.workbook#previouslysaved)|指定工作簿是否已在本地或在线保存。 只读。|
||[usePrecisionAsDisplayed](/javascript/api/excel/excel.workbook#useprecisionasdisplayed)|如果此工作簿中的计算仅使用显示的数字精度来完成，则为 True。|
|[WorkbookAutoSaveSetting ChangedEventArgs](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs)|[type](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs#type)|表示事件的类型。 有关详细信息，请参阅 Excel.EventType。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[enableCalculation](/javascript/api/excel/excel.worksheet#enablecalculation)|获取或设置工作表的 enableCalculation 属性。|
||[findAll(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#findall-text--criteria-)|根据指定的条件查找给定字符串的所有匹配项，并将它们作为包含一个或多个矩形区域的 RangeAreas 对象返回。|
||[findAllOrNullObject(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#findallornullobject-text--criteria-)|根据指定的条件查找给定字符串的所有匹配项，并将它们作为包含一个或多个矩形区域的 RangeAreas 对象返回。|
||[getRanges(address?: string)](/javascript/api/excel/excel.worksheet#getranges-address-)|获取按地址或名称指定的 RangeAreas 对象，它表示一个或多个矩形区域块。|
||[autoFilter](/javascript/api/excel/excel.worksheet#autofilter)|表示工作表的 AutoFilter 对象。 只读。|
||[horizontalPageBreaks](/javascript/api/excel/excel.worksheet#horizontalpagebreaks)|获取工作表的水平分页符集合。 此集合仅包含手动分页符。|
||[onFormatChanged](/javascript/api/excel/excel.worksheet#onformatchanged)|在特定工作表上更改格式时发生。|
||[pageLayout](/javascript/api/excel/excel.worksheet#pagelayout)|获取工作表的 PageLayout 对象。|
||[shapes](/javascript/api/excel/excel.worksheet#shapes)|返回工作表上的所有 Shape 对象的集合。 只读。|
||[verticalPageBreaks](/javascript/api/excel/excel.worksheet#verticalpagebreaks)|获取工作表的垂直分页符集合。 此集合仅包含手动分页符。|
||[replaceAll(text: string, replacement: string, criteria: Excel.ReplaceCriteria)](/javascript/api/excel/excel.worksheet#replaceall-text--replacement--criteria-)|根据当前工作表中指定的条件查找并替换给定的字符串。|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[details](/javascript/api/excel/excel.worksheetchangedeventargs#details)|表示更改详情的信息|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onChanged](/javascript/api/excel/excel.worksheetcollection#onchanged)|在更改工作簿中的任何工作表时发生。|
||[onFormatChanged](/javascript/api/excel/excel.worksheetcollection#onformatchanged)|在更改工作簿中的任何工作表的格式时发生。|
||[onSelectionChanged](/javascript/api/excel/excel.worksheetcollection#onselectionchanged)|在任何工作表上更改选择时发生。|
|[WorksheetFormatChangedEventArgs](/javascript/api/excel/excel.worksheetformatchangedeventargs)|[address](/javascript/api/excel/excel.worksheetformatchangedeventargs#address)|获取区域地址，该地址表示特定工作表上的更改区域。|
||[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#getrange-ctx-)|获取区域，该区域表示特定工作表上的更改区域。|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#getrangeornullobject-ctx-)|获取区域，该区域表示特定工作表上的更改区域。 它可能会返回 null 对象。|
||[source](/javascript/api/excel/excel.worksheetformatchangedeventargs#source)|获取事件源。 有关详细信息，请参阅 Excel.EventSource。|
||[type](/javascript/api/excel/excel.worksheetformatchangedeventargs#type)|获取事件的类型。 有关详细信息，请参阅 Excel.EventType。|
||[worksheetId](/javascript/api/excel/excel.worksheetformatchangedeventargs#worksheetid)|获取其中的数据发生更改的工作表的 ID。|
|[WorksheetSearchCriteria](/javascript/api/excel/excel.worksheetsearchcriteria)|[completeMatch](/javascript/api/excel/excel.worksheetsearchcriteria#completematch)|指定匹配必须是完整匹配还是部分匹配。 默认值为 false（部分）。|
||[matchCase](/javascript/api/excel/excel.worksheetsearchcriteria#matchcase)|指定匹配是否区分大小写。 默认值为 false（不区分大小写）。|

## <a name="whats-new-in-excel-javascript-api-18"></a>Excel JavaScript API 1.8 的最近更新

Excel JavaScript API 要求集 1.8 的功能包括适用于数据透视表、数据验证、图表、图表事件、性能选项和工作簿创建的 API。

### <a name="pivottable"></a>数据透视表

加载项通过数据透视表 API 的波形 2 设置数据透视表的层次结构。 现在可以控制数据及其聚合方式。 [数据透视表](/office/dev/add-ins/excel/excel-add-ins-pivottables)一文详细介绍了新的数据透视表功能。

### <a name="data-validation"></a>数据有效性

数据有效性可以控制用户在工作表中输入的内容。 可以将单元格限制为预定义的答案集，或者在用户输入无效数据时提供弹出警告。 立即详细了解[向区域添加数据有效性](/office/dev/add-ins/excel/excel-add-ins-data-validation)。

### <a name="charts"></a>图表

另一轮图表 API 可更好地对图表元素进行编程控制。 现在，你对图例、坐标轴、趋势线和绘图区拥有更高的访问权限。

### <a name="events"></a>事件

已为图表添加更多[事件](/office/dev/add-ins/excel/excel-add-ins-events)。 让加载项处理用于与图表的交互。 此外，你还可以在整个工作簿中[触发事件](/office/dev/add-ins/excel/performance#enable-and-disable-events)。

|对象| 最近更新| 说明|要求集|
|:----|:----|:----|:----|
|[application](/javascript/api/excel/excel.application)|_方法_ > [createWorkbook(base64File: string)](/javascript/api/excel/excel.application)|使用可选 base64 编码的 .xlsx 文件创建新的隐藏工作簿。|1.8|
|[basicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|_属性_ > formula1|获取或设置 Formula1，即最小值或取决于运算符的值。|1.8|
|[basicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|_属性_ > formula2|获取或设置 Formula2，即最大值或取决于运算符的值。|1.8|
|[basicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|_关系_ > operator|用于验证数据有效性的运算符。|1.8|
|[chart](/javascript/api/excel/excel.chart)|_属性_ > categoryLabelLevel|返回或设置 ChartCategoryLabelLevel 枚举常量，该常量代表分类标签源自位置的级别。 读/写。|1.8|
|[chart](/javascript/api/excel/excel.chart)|_属性_ > plotVisibleOnly|如果仅绘制可见单元格，则为 True。 如果绘制可见单元格和隐藏单元格，则为 False。 读写。|1.8|
|[chart](/javascript/api/excel/excel.chart)|_属性_ > seriesNameLevel|返回或设置 ChartSeriesNameLevel 枚举常量，该常量代表系列名称源自位置的级别。 读/写。|1.8|
|[chart](/javascript/api/excel/excel.chart)|_属性_ > showDataLabelsOverMaximum|表示当值大于数值轴上的最大值时是否显示数据标签。|1.8|
|[chart](/javascript/api/excel/excel.chart)|_属性_ > style|返回或设置图表的图表样式。 读写。|1.8|
|[chart](/javascript/api/excel/excel.chart)|_关系_ > displayBlanksAs|返回或设置图表上的空白单元格的绘制方式。 读写。|1.8|
|[chart](/javascript/api/excel/excel.chart)|_关系_ > plotArea|表示图表的绘制区域。 只读。|1.8|
|[chart](/javascript/api/excel/excel.chart)|_关系_ > plotBy|返回或设置图表上的列或行用作数据系列的方式。 读写。|1.8|
|[chartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|_属性_ > chartId|获取已启用图表的 ID。|1.8|
|[chartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|_属性_ > type|获取事件的类型。|1.8|
|[chartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|_属性_ > worksheetId|获取其中的图表已启用的工作表的 ID。|1.8|
|[chartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|_属性_ > chartId|获取已添加至工作表的图表的 ID。|1.8|
|[chartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|_属性_ > type|获取事件的类型。|1.8|
|[chartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|_属性_ > worksheetId|获取已在其中添加图表的工作表的 ID。|1.8|
|[chartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|_关系_ > source|获取事件源。|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_ > isBetweenCategories|表示数值轴是否与分类之间的分类轴交叉。|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_ > multiLevel|表示是否为多级轴。|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_ > numberFormat|表示轴刻度线标签的格式代码。|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_ > offset|表示不同标签级别之间的距离以及一级标签和轴线之间的距离。 此值应该是 0 到 1000 之间的整数。|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_ > positionAt|表示两轴交叉的特定轴位置。 应使用 SetPositionAt(double) 方法设置此属性。 只读。|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_ > textOrientation|表示轴刻度线标签的文本方向。 此值应是 -90 到 90 或 180（垂直文本）之间的整数。|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_关系_ > alignment|表示指定轴刻度线标签的对齐方式。|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_关系_ > position|表示两轴交叉的特定轴位置。|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_方法_ > [setPositionAt(value: double)](/javascript/api/excel/excel.chartaxis)|设置两轴交叉的特定轴位置。|1.8|
|[chartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|_关系_ > fill|表示图表填充格式。 只读。|1.8|
|[chartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|_方法_ > [setFormula(formula: string)](/javascript/api/excel/excel.chartaxistitle)|该字符串值表示采用 A1 表示法的图表轴标题的公式。|1.8|
|[chartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|_关系_ > border|表示边框格式，包括颜色、线条样式和粗细。 只读。|1.8|
|[chartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|_关系_ > fill|表示图表填充格式。 只读。|1.8|
|[chartBorder](/javascript/api/excel/excel.chartborder)|_方法_ > [clear()](/javascript/api/excel/excel.chartborder)|清除图表元素的边框格式。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_属性_ > autoText|该布尔值表示数据标签是否根据上下文自动生成相应的文本。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_属性_ > formula|该字符串值表示采用 A1 表示法的图表数据标签的公式。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_属性_ > height|返回图表数据标签的高度，以磅为单位。 只读。 如果图表数据标签不可见，则为 Null。 只读。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_属性_ > left|表示图表数据标签左边缘到图表区域左边缘的距离，以磅为单位。 如果图表数据标签不可见，则为 Null。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_属性_ > numberFormat|该字符串值表示数据标签的格式代码。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_属性_ > text|该字符串表示图表上的数据标签文本。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_属性_ > textOrientation|表示图表数据标签的文本方向。 此值应是 -90 到 90 或 180（垂直文本）之间的整数。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_属性_ > top|表示图表数据标签上边缘到图表区域顶部的距离，以磅为单位。 如果图表数据标签不可见，则为 Null。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_属性_ > width|返回图表数据标签的宽度，以磅为单位。 只读。 如果图表数据标签不可见，则为 Null。 只读。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_关系_ > format|表示图表数据标签的格式。 只读。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_关系_ > horizontalAlignment|表示图表数据标签水平对齐。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_关系_ > verticalAlignment|表示图表数据标签垂直对齐。|1.8|
|[chartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|_关系_ > border|表示边框格式，包括颜色、线条样式和粗细。 只读。|1.8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_属性_ > autoText|表示数据标签是否根据上下文自动生成相应的文本。|1.8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_属性_ > numberFormat|表示数据标签的格式代码。|1.8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_属性_ > textOrientation|表示数据标签的文本方向。 此值应是 -90 到 90 或 0 到 180（垂直文本）之间的整数。|1.8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_关系_ > horizontalAlignment|表示图表数据标签水平对齐。|1.8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_关系_ > verticalAlignment|表示图表数据标签垂直对齐。|1.8|
|[chartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|_属性_ > chartId|获取停用图表的 ID。|1.8|
|[chartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|_属性_ > type|获取事件的类型。|1.8|
|[chartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|_属性_ > worksheetId|获取其中的图表已停用的工作表的 ID。|1.8|
|[chartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|_属性_ > chartId|获取已从工作表删除的图表的 ID。|1.8|
|[chartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|_属性_ > type|获取事件的类型。|1.8|
|[chartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|_属性_ > worksheetId|获取已在其中删除图表的工作表的 ID。|1.8|
|[chartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|_关系_ > source|获取事件源。|1.8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_属性_ > height|表示图表图例上的 legendEntry 的高度。 只读。|1.8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_属性_ > index|表示图表图例中的 legendEntry 的索引。 只读。|1.8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_属性_ > left|表示图表 legendEntry 的左侧。 只读。|1.8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_属性_ > top|表示图表 legendEntry 的顶部。 只读。|1.8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_属性_ > width|表示图表图例上的 legendEntry 的宽度。 只读。|1.8|
|[chartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|_关系_ > border|表示边框格式，包括颜色、线条样式和粗细。 只读。|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_属性_ > height|表示 plotArea 的高度值。|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_属性_ > insideHeight|表示 plotArea 的 insideHeight 值。|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_属性_ > insideLeft|表示 plotArea 的 insideLeft 值。|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_属性_ > insideTop|表示 plotArea 的 insideTop 值。|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_属性_ > insideWidth|表示 plotArea 的 insideWidth 值。|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_属性_ > left|表示 plotArea 的 left 值。|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_属性_ > top|表示 plotArea 的 top 值。|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_属性_ > width|表示 plotArea 的宽度值。|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_关系_ > format|表示图表 plotArea 的格式。 只读。|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_关系_ > position|表示 plotArea 的位置。|1.8|
|[chartPlotAreaFormat](/javascript/api/excel/excel.chartplotareaformat)|_关系_ > border|表示图表 plotArea 的边框属性。 只读。|1.8|
|[chartPlotAreaFormat](/javascript/api/excel/excel.chartplotareaformat)|_关系_ > fill|表示对象的填充格式，包括背景格式信息。只读。|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_属性_ > explosion|返回或设置饼图或圆环图切片的分解程度值。 如果未分解（切片尖端位于饼图中心），则返回 0（零）。 读写。|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_属性_ > firstSliceAngle|返回或设置第一个饼图或圆环图切片的角度，以度为单位（从垂直方向起为顺时针）。 只适用于饼图、三维饼图和圆环图。 可以是 0 到 360 之间的值。 读写|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_属性_ > invertIfNegative|如果 Excel 在值为负数时反转项目中的图案，则为 True。 读写。|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_属性_ > overlap|指定条柱的摆放方式。 可以是 -100 到 100 之间的值。 只适用于二维条形图和二维柱形图。 读写。|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_属性_ > secondPlotSize|返回或设置复合饼图或复合条饼图的辅助分区的大小，以占主饼图大小的百分比表示。 可以是 5 到 200 之间的值。 读写。|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_属性_ > varyByCategories|如果 Excel 为每个数据标记分配不同的颜色或图案，则为 True。 图表必须只包含一个系列。 读写。|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_关系_ > axisGroup|返回或设置指定系列的组。读写|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_关系_ > dataLabels|表示系列中所有数据标签的集合。 只读。|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_关系_ > splitType|返回或设置复合饼图或复合条饼图中两个分区的拆分方式。 读写。|1.8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_属性_ > backwardPeriod|表示趋势线向后延伸的周期数。|1.8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_属性_ > forwardPeriod|表示趋势线向前延伸的周期数。|1.8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_属性_ > showEquation|如果图表上显示趋势线公式，则为 True。|1.8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_属性_ > showRSquared|如果图表上显示趋势线的 R 平方值，则为 True。|1.8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_关系_ > label|表示图表趋势线的标签。 只读。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_属性_ > autoText|该布尔值表示趋势线标签是否根据上下文自动生成相应的文本。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_属性_ > formula|该字符串值表示采用 A1 表示法的图表趋势线标签的公式。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_属性_ > height|返回图表趋势线标签的高度，以磅为单位。 只读。 如果图表趋势线标签不可见，则为 Null。 只读。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_属性_ > left|表示图表趋势线标签左边缘到图表区域左边缘的距离，以磅为单位。 如果图表趋势线标签不可见，则为 Null。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_属性_ > numberFormat|该字符串值表示趋势线标签的格式代码。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_属性_ > text|该字符串表示图表上的趋势线标签文本。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_属性_ > textOrientation|表示图表趋势线标签的文本方向。 此值应是 -90 到 90 或 180（垂直文本）之间的整数。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_属性_ > top|表示图表趋势线标签上边缘到图表区域顶部的距离，以磅为单位。 如果图表趋势线标签不可见，则为 Null。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_属性_ > width|返回图表趋势线标签的宽度，以磅为单位。 只读。 如果图表趋势线标签不可见，则为 Null。 只读。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_关系_ > format|表示图表趋势线标签的格式。 只读。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_关系_ > horizontalAlignment|表示图表趋势线标签水平对齐。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_关系_ > verticalAlignment|表示图表趋势线标签垂直对齐。|1.8|
|[chartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|_关系_ > border|表示边框格式，包括颜色、线条样式和粗细。 只读。|1.8|
|[chartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|_关系_ > fill|表示当前图表趋势线标签的填充格式。 只读。|1.8|
|[chartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|_关系_ > font|表示图表趋势线标签的字体属性（字体名称、字体大小、颜色等）。 只读。|1.8|
|[customDataValidation](/javascript/api/excel/excel.customdatavalidation)|_属性_ > formula| 自定义数据验证公式。 这将创建特殊输入规则，例如阻止重复或显示单元格区域的总值。|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_属性_ > id|DataPivotHierarchy ID。 只读。|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_属性_ > name|DataPivotHierarchy 的名称。|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_属性_ > numberFormat|DataPivotHierarchy 的数字格式。|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_属性_ > position|DataPivotHierarchy 的位置。|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_关系_ > field|返回与 DataPivotHierarchy 相关联的 PivotFields。 只读。|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_关系_ > showAs|确定数据是否应显示为特定计算汇总。|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_关系_ > summarizeBy|确定是否显示 DataPivotHierarchy 的所有项。|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_方法_ > [setToDefault()](/javascript/api/excel/excel.datapivothierarchy#settodefault--)|将 DataPivotHierarchy 重置回其默认值。|1.8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_属性_ > items|DataPivotHierarchy 对象的集合。 只读。|1.8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_方法_ > [add(pivotHierarchy: PivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection)|将 PivotHierarchy 添加到当前轴。|1.8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_方法_ > [getCount()](/javascript/api/excel/excel.datapivothierarchycollection)|获取集合中的透视层级结构的数量。|1.8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_方法_ > [getItem(name: string)](/javascript/api/excel/excel.datapivothierarchycollection)|按名称或 ID 获取 DataPivotHierarchy。|1.8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_方法_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.datapivothierarchycollection)|按名称获取 DataPivotHierarchy。 如果 DataPivotHierarchy 不存在，则返回 Null 对象。|1.8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_方法_ > [remove(DataPivotHierarchy: DataPivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection)|从当前轴删除 PivotHierarchy。|1.8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_属性_ > ignoreBlanks|忽略空白：不会对空白单元格执行数据严重，默认为 true。|1.8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_属性_ > valid|表示所有单元格值根据数据有效性规则是否全部有效。 只读。|1.8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_关系_ > errorAlert|用户输入无效数据时，出现错误警报。|1.8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_关系_ > prompt|用户选择某个单元格时进行提示。|1.8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_关系_ > rule|包含不同类型的数据有效性标准的数据有效性规则。|1.8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_关系_ > type|数据有效性类型，有关详细信息，请参阅 [Excel.DataValidationType](/javascript/api/excel/excel.datavalidationtype)。 只读。|1.8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_方法_ > [clear()](/javascript/api/excel/excel.datavalidation)|清除当前区域中的数据有效性。|1.8|
|[dataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|_属性_ > message|表示错误警报消息。|1.8|
|[dataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|_属性_ > showAlert|确定在用户输入无效数据时是否显示错误警报对话框。 默认值为 true。|1.8|
|[dataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|_属性_ > title|表示错误警报对话框标题。|1.8|
|[dataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|_关系_ > style|表示数据有效性警报类型，有关详细信息，请参阅 [Excel.DataValidationAlertStyle](/javascript/api/excel/excel.datavalidationalertstyle)。|1.8|
|[dataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)|_属性_ > message|表示提示消息。|1.8|
|[dataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)|_属性_ > showPrompt|确定在用户选择具有数据有效性的单元格时是否显示提示。|1.8|
|[dataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)|_属性_ > title|表示提示标题。|1.8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_关系_ > custom|自定义数据有效性条件。|1.8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_关系_ > date|日期数据有效性条件。|1.8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_关系_ > decimal|小数数据有效性条件。|1.8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_关系_ > list|列表数据有效性条件。|1.8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_关系_ > textLength|TextLength 数据有效性条件。|1.8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_关系_ > time|时间数据有效性条件。|1.8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_关系_ > wholeNumber|WholeNumber 数据有效性条件。|1.8|
|[dateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|_属性_ > formula1|获取或设置 Formula1，即最小值或取决于运算符的值。|1.8|
|[dateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|_属性_ > formula2|获取或设置 Formula2，即最大值或取决于运算符的值。|1.8|
|[dateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|_关系_ > operator|用于验证数据有效性的运算符。|1.8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_属性_ > enableMultipleFilterItems|确定是否允许多个筛选项。|1.8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_属性_ > id|FilterPivotHierarchy 的 ID。 只读。|1.8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_属性_ > name|FilterPivotHierarchy 的名称。|1.8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_属性_ > position|FilterPivotHierarchy 的位置。|1.8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_关系_ > fields|返回与 FilterPivotHierarchy 相关联的 PivotFields。 只读。|1.8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_方法_ > [setToDefault()](/javascript/api/excel/excel.filterpivothierarchy)|将 FilterPivotHierarchy 重置回其默认值。|1.8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_属性_ > items|FilterPivotHierarchy 对象的集合。 只读。|1.8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_方法_ > [add(pivotHierarchy: PivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection)|将 PivotHierarchy 添加到当前轴。 如果行、列或筛选轴上的其他位置存在层次结构，则会将该层次结构从相应的位置移除。|1.8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_方法_ > [getCount()](/javascript/api/excel/excel.filterpivothierarchycollection)|获取集合中的透视层级结构的数量。|1.8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_方法_ > [getItem(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection)|按名称或 ID 获取 FilterPivotHierarchy。|1.8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_方法_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection)|按名称获取 FilterPivotHierarchy。 如果 FilterPivotHierarchy 不存在，则返回 Null 对象。|1.8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_方法_ > [remove(filterPivotHierarchy: FilterPivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection)|从当前轴删除 PivotHierarchy。|1.8|
|[listDataValidation](/javascript/api/excel/excel.listdatavalidation)|_属性_ > inCellDropDown|是否显示单元格下拉菜单中的列表，默认为 true。|1.8|
|[listDataValidation](/javascript/api/excel/excel.listdatavalidation)|_属性_ > source|数据有效性列表源|1.8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_属性_ > id|PivotField 的 ID。 只读。|1.8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_属性_ > name|PivotField 的名称。|1.8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_属性_ > showAllItems|确定是否显示 PivotField 的所有项。|1.8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_关系_ > items|返回与 PivotField 相关联的 PivotFields。 只读。|1.8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_关系_ > subtotals|PivotField 小计。|1.8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_方法_ > [sortByLabels(sortby: SortBy)](/javascript/api/excel/excel.pivotfield)|PivotField 排序。 如果指定 DataPivotHierarchy，则会基于它进行排序，如果未指定，则会基于 PivotField 本身进行排序。|1.8|
|[pivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|_属性_ > items|pivotField 对象的集合。 只读。|1.8|
|[pivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|_方法_ > [getCount()](/javascript/api/excel/excel.pivotfieldcollection)|获取集合中的透视层级结构的数量。|1.8|
|[pivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|_方法_ > [getItem(name: string)](/javascript/api/excel/excel.pivotfieldcollection)|按名称或 ID 获取 PivotHierarchy。|1.8|
|[pivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|_方法_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotfieldcollection)|按名称获取 PivotHierarchy。 如果 PivotHierarchy 不存在，则返回 Null 对象。|1.8|
|[pivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|_属性_ > id|PivotHierarchy 的 ID。 只读。|1.8|
|[pivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|_属性_ > name|PivotHierarchy 的名称。|1.8|
|[pivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|_关系_ > fields|返回与 PivotHierarchy 相关联的 PivotFields。 只读。|1.8|
|[pivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|_属性_ > items|PivotHierarchy 对象的集合。 只读。|1.8|
|[pivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|_方法_ > [getCount()](/javascript/api/excel/excel.pivothierarchycollection)|获取集合中的透视层级结构的数量。|1.8|
|[pivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|_方法_ > [getItem(name: string)](/javascript/api/excel/excel.pivothierarchycollection)|按名称或 ID 获取 PivotHierarchy。|1.8|
|[pivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|_方法_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivothierarchycollection)|按名称获取 PivotHierarchy。 如果 PivotHierarchy 不存在，则返回 Null 对象。|1.8|
|[pivotItem](/javascript/api/excel/excel.pivotitem)|_属性_ > id|PivotItem 的 ID。 只读。|1.8|
|[pivotItem](/javascript/api/excel/excel.pivotitem)|_属性_ > isExpanded|确定是展开项以显示子项还是折叠项并隐藏子项。|1.8|
|[pivotItem](/javascript/api/excel/excel.pivotitem)|_属性_ > name|PivotItem 的名称。|1.8|
|[pivotItem](/javascript/api/excel/excel.pivotitem)|_属性_ > visible|确定 PivotItem 是否可见。|1.8|
|[pivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|_属性_ > items|pivotItem 对象的集合。 只读。|1.8|
|[pivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|_方法_ > [getCount()](/javascript/api/excel/excel.pivotitemcollection)|获取集合中的透视层级结构的数量。|1.8|
|[pivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|_方法_ > [getItem(name: string)](/javascript/api/excel/excel.pivotitemcollection)|按名称或 ID 获取 PivotHierarchy。|1.8|
|[pivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|_方法_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotitemcollection)|按名称获取 PivotHierarchy。 如果 PivotHierarchy 不存在，则返回 Null 对象。|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_属性_ > showColumnGrandTotals|如果数据透视表显示列总计，则为 True。|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_属性_ > showRowGrandTotals|如果数据透视表显示行总计，则为 True。|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_属性_ > subtotalLocation|此属性指示数据透视表上的所有字段的 SubtotalLocationType。 如果字段状态不同，则为 null。 可能的值包括：AtTop、AtBottom。|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_关系_ > layoutType|此属性指示数据透视表上的所有字段的 PivotLayoutType。 如果字段状态不同，则为 null。|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_方法_ > [getColumnLabelRange()](/javascript/api/excel/excel.pivotlayout)|返回数据透视表列标签所在位置的区域。|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_方法_ > [getDataBodyRange()](/javascript/api/excel/excel.pivotlayout)|返回数据透视表数据值所在位置的区域。|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_方法_ > [getFilterAxisRange()](/javascript/api/excel/excel.pivotlayout)|返回数据透视表筛选区的区域。|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_方法_ > [getRange()](/javascript/api/excel/excel.pivotlayout)|返回存在数据透视表的区域，不包括筛选区。|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_方法_ > [getRowLabelRange()](/javascript/api/excel/excel.pivotlayout)|返回数据透视表行标签所在位置的区域。|1.8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_关系_ > columnHierarchies|数据透视表的列透视层级结构。 只读。|1.8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_关系_ > dataHierarchies|数据透视表的数据透视层级结构。 只读。|1.8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_关系_ > filterHierarchies|数据透视表的筛选器透视层级结构。 只读。|1.8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_关系_ > hierarchies|数据透视表的透视层级结构。 只读。|1.8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_关系_ > layout|PivotLayout，用于说明数据透视表的布局和可视化结构。 只读。|1.8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_关系_ > rowHierarchies|数据透视表的行透视层级结构。 只读。|1.8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_方法_  >  [delete()](/javascript/api/excel/excel.pivottable)|删除 PivotTable 对象。|1.8|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_方法_ > [add(name: string, source: object, destination: object)](/javascript/api/excel/excel.pivottablecollection)|基于指定的数据源添加数据透视表，并将其插入到目标区域的左上单元格。|1.8|
|[range](/javascript/api/excel/excel.range)|_关系_ > dataValidation|返回数据有效性对象。 只读。|1.8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_属性_ > id|RowColumnPivotHierarchy 的 ID。 只读。|1.8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_属性_ > name|RowColumnPivotHierarchy 的名称。|1.8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_属性_ > position|RowColumnPivotHierarchy 的位置。|1.8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_关系_ > fields|返回与 RowColumnPivotHierarchy 相关联的 PivotFields。 只读。|1.8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_方法_ > [setToDefault()](/javascript/api/excel/excel.rowcolumnpivothierarchy)|将 RowColumnPivotHierarchy 重置回其默认值。|1.8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_属性_ > items|RowColumnPivotHierarchy 对象的集合。 只读。|1.8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_方法_ > [add(pivotHierarchy: PivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|将 PivotHierarchy 添加到当前轴。 行和列上的其他位置是否存在层次结构。|1.8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_方法_ > [getCount()](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|获取集合中的透视层级结构的数量。|1.8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_方法_ > [getItem(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|按名称或 ID 获取 RowColumnPivotHierarchy。|1.8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_方法_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|按名称获取 RowColumnPivotHierarchy。 如果 RowColumnPivotHierarchy 不存在，则返回 Null 对象。|1.8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_方法_ > [remove(rowColumnPivotHierarchy: RowColumnPivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|从当前轴删除 PivotHierarchy。|1.8|
|[runtime](/javascript/api/excel/excel.runtime)|_属性_ > enableEvents|切换当前任务窗格或内容加载项中的 JavaScript 事件。|1.8|
|[showAsRule](/javascript/api/excel/excel.showasrule)|_关系_ > baseField|基于 ShowAs 计算的基础 PivotField，如适用，基于 ShowAsCalculation 类型，否则为 null。|1.8|
|[showAsRule](/javascript/api/excel/excel.showasrule)|_关系_ > baseItem|基于 ShowAs 计算的基础 Item，如适用，基于 ShowAsCalculation 类型，否则为 null。|1.8|
|[showAsRule](/javascript/api/excel/excel.showasrule)|_关系_ > calculation|数据 PivotField 使用的 ShowAs 计算。|1.8|
|[style](/javascript/api/excel/excel.style)|_属性_ > autoIndent|指示将单元格中的文本对齐方式设为相等分布时文本是否会自动缩进。|1.8|
|[style](/javascript/api/excel/excel.style)|_属性_ > textOrientation|此样式中的文本方向。|1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_属性_ > automatic|如果将“Automatic”设为 true，则在设置 Subtotals 时，所有其他值均会被忽略。|1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_属性_ > average| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_属性_ > count| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_属性_> countNumbers| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_属性_ > max| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_属性_ > min| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_属性_ > product| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_属性_ > standardDeviation| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_属性_ > standardDeviationP| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_属性_ > sum| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_属性_ > variance| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_属性_ > varianceP| |1.8|
|[table](/javascript/api/excel/excel.table)|_属性_ > legacyId|返回数字 ID。只读。|1.8|
|[workbook](/javascript/api/excel/excel.workbook)|_属性_ > readOnly|如果在只读模式下打开工作簿，则为 True。 只读。|1.8|
|[workbookCreated](/javascript/api/excel/excel.workbookcreated)|_属性_ > id|返回用于唯一标识 WorkbookCreated 对象的值。 只读。|1.8|
|[workbookCreated](/javascript/api/excel/excel.workbookcreated)|_方法_ > [open()](/javascript/api/excel/excel.workbookcreated)|打开工作表。|1.8|
|[worksheet](/javascript/api/excel/excel.worksheet)|_属性_ > showGridlines|获取或设置工作表的网格线标志。|1.8|
|[worksheet](/javascript/api/excel/excel.worksheet)|_属性_ > showHeadings|获取或设置工作表的标题标志。|1.8|
|[worksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|_属性_ > type|获取事件的类型。|1.8|
|[worksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|_属性_ > worksheetId|获取计算的工作表的 ID。|1.8|

## <a name="whats-new-in-excel-javascript-api-17"></a>Excel JavaScript API 1.7 的最近更新

Excel JavaScript API 要求集 1.7 的功能包括用于图表、事件、工作表、区域、文档属性、已命名项目、保护选项和样式的 API。

### <a name="customize-charts"></a>自定义图表

通过新的图表 API，你可以创建其他图表类型、向图表中添加数据系列、设置图表标题、添加轴标题、添加显示单位、添加采用移动平均值的趋势线、将趋势线更改为线性趋势线等。 下面是一些示例：

* 图表轴 - 获取、设置、格式化和删除图表中的轴单位、标签和标题。
* 图表系列 - 添加、设置和删除图表中的某个系列。  更改系列标记、绘制顺序和大小。
* 图表趋势线 - 添加、获取和格式化图表中的趋势线。
* 图表图例 - 设置图表中的图例字体的格式。
* 图表点 - 设置图表点颜色。
* 图表标题子字符串 - 获取和设置图表的标题子字符串。
* 图表类型 - 用于创建更多图表类型的选项。

### <a name="events"></a>事件

Excel 事件 API 提供了多个事件处理程序，以便加载项能够在发生特定事件时自动运行指定的函数。 可以将函数设计为执行方案所需的任何操作。 有关当前可用的事件列表，请参阅[使用 Excel JavaScript API 处理事件](/office/dev/add-ins/excel/excel-add-ins-events)。

### <a name="customize-the-appearance-of-worksheets-and-ranges"></a>自定义工作表和区域的外观

使用新的 API 可以通过多种方式自定义工作表的外观：

* 冻结窗格，使特定行或列在你滚动工作表时保持可见。 例如，如果工作表中的第一行包含标题，则可以冻结此行，以便在你向下滚动工作表时列标题保持可见。
* 修改工作表标签颜色。
* 添加工作表标题。


可以通过多种方式自定义区域的外观：

* 设置某个区域的单元格样式，确保该区域内的所有单元格采用一致的格式。 单元格 样式是一组定义的格式特征，例如字体和字号、数字格式、单元格边框和单元格底纹。 使用 Excel 中的任意内置单元格样式，或者使用自己的自定义单元格样式。
* 设置区域的文本方向。
* 添加或修改区域上链接至工作表中的其他位置或外部位置的超链接。

### <a name="manage-document-properties"></a>管理文档属性

使用文档属性 API，你可以访问内置文档属性，并且还可以创建和管理自定义文档属性，以存储工作表的状态和驱动工作流和业务逻辑。

### <a name="copy-worksheets"></a>复制工作表

使用工作表复制 APIs，你可以将一个工作表中的数据和格式复制到相同工作簿中的另一个工作表，从而减少所需的数据传输量。

### <a name="handle-ranges-with-ease"></a>轻松地处理区域

使用各种区域 API，你可以完成诸如获取周围区域、获取大小经过重设的区域之类的任务。 这些 API 可以显著提高诸如区域操作和寻址之类任务的效率。

此外：

* 工作簿和工作表保护选项 - 使用这些 API 可保护工作表和工作簿结构中的数据。
* 更新已命名项目 - 使用此 API 可更新已命名项目。
* 获取活动单元格 - 使用此 API 可获取工作表中的活动单元格。

|对象| 最近更新| 说明|要求集|
|:----|:----|:----|:----|
|[chart](/javascript/api/excel/excel.chart)|_属性_ > chartType|表示图表的类型。 可能的值包括：ColumnClustered、ColumnStacked、ColumnStacked100、BarClustered、BarStacked、BarStacked100、LineStacked、LineStacked100、LineMarkers、LineMarkersStacked、LineMarkersStacked100、PieOfPie 等。|1.7|
|[chart](/javascript/api/excel/excel.chart)|_属性_ > id|图表的唯一 ID。 只读。|1.7|
|[chart](/javascript/api/excel/excel.chart)|_属性_ > showAllFieldButtons|表示是否在数据透视图上显示所有字段按钮。|1.7|
|[chartAreaFormat](/javascript/api/excel/excel.chartareaformat)|_关系_ > border|表示图表区域的边框格式，包括颜色、线条样式和粗细。 只读。|1.7|
|[chartAxes](/javascript/api/excel/excel.chartaxes)|_方法_ > getItem(type: string, group: string)|返回通过类型和组标识的特定轴。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_ > axisBetweenCategories|表示数值轴是否与分类之间的分类轴交叉。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_ > axisGroup|返回或设置指定轴的组。 只读。 可能的值包括：Primary、Secondary。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_ > categoryType|返回或设置分类轴类型。 可能的值包括：Automatic、TextAxis、DateAxis。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_ > crosses|表示两轴交叉处的特定轴。 可能的值包括：Automatic、Maximum、Minimum、Custom。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_ > crossesAt|表示两轴交叉处的特定轴。 只读。 设置此属性应使用 SetCrossesAt(double) 方法。 只读。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_ > customDisplayUnit|标识自定义轴显示单位值。 只读。 要设置此属性，请使用 SetCustomDisplayUnit(double) 方法。 只读。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_ > displayUnit|表示轴显示单位。 可能的值包括：None、Hundreds、Thousands、TenThousands、HundredThousands、Millions、TenMillions、HundredMillions、Billions、Trillions、Custom。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_ > height|表示图表轴的高度，以磅为单位。 如果轴不可见，则为 Null。 只读。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_ > left|表示轴的左边缘到图表区域左侧的距离，以磅为单位。 如果轴不可见，则为 Null。 只读。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_ > logBase|表示使用对数刻度时对数的底数。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_ > reversePlotOrder|表示 Microsoft Excel 是否按照最后一个到第一个的顺序绘制数据点。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_ > scaleType|表示数值轴刻度类型。 可能的值包括：Linear、Logarithmic。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_ > showDisplayUnitLabel|表示轴显示单位标签是否可见。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_ > tickLabelSpacing|表示刻度线标签之间的分类或系列数。 可以是 1 到 31999 的值或空字符串（自动设置）。 返回的值始终为数字。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_ > tickMarkSpacing|表示刻度线之间的分类或系列数。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_ > top|表示轴的上边缘到图表区域顶部的距离，以磅为单位。 如果轴不可见，则为 Null。 只读。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_ > type|表示轴类型。 只读。 可能的值包括：Invalid、Category、Value、Series。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_ > visible|该布尔值表示轴的可见性。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_ > width|表示图表轴的宽度，以磅为单位。 如果轴不可见，则为 Null。 只读。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_关系_ > baseTimeUnit|返回或设置指定分类轴的基本单位。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_关系_ > majorTickMark|表示特定轴的主要刻度线类型。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_关系_ > majorTimeUnitScale|返回或设置当 CategoryType 属性设为 TimeScale 时分类轴的主要单位刻度值。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_关系_ > minorTickMark|表示指定轴的次要刻度线类型。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_关系_ > minorTimeUnitScale|返回或设置当 CategoryType 属性设为 TimeScale 时分类轴的次要单位刻度值。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_关系_ > tickLabelPosition|表示特定轴上的刻度线标签位置。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_方法_ > setCategoryNames(sourceData: Range)|设置指定轴的所有分类名称。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_方法_ > setCrossesAt(value: double)|设置两轴交叉处的特定轴。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_方法_ > setCustomDisplayUnit(value: double)|将轴显示单位设为自定义值。|1.7|
|[chartBorder](/javascript/api/excel/excel.chartborder)|_属性_ > color|表示图表中的边框颜色的 HTML 颜色代码。|1.7|
|[chartBorder](/javascript/api/excel/excel.chartborder)|_属性_ > weight|表示边框的粗细，以磅为单位。|1.7|
|[chartBorder](/javascript/api/excel/excel.chartborder)|_关系_ > lineStyle|表示边框的线条样式。|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_属性_ > position|表示数据标签的位置的 DataLabelPosition 值。可能的值是：None、Center、InsideEnd、InsideBase、OutsideEnd、Left、Right、Top、Bottom、BestFit、Callout。|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_属性_ > separator|该字符串表示用于图表中数据标签的分隔符。|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_属性_ > showBubbleSize|该布尔值表示数据标签气泡大小是否可见。|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_属性_ > showCategoryName|表示数据标签分类名称是否可见的布尔值。|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_属性_ > showLegendKey|该布尔值表示数据标签图例标示是否可见。|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_属性_ > showPercentage|该布尔值表示数据标签百分比是否可见。|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_属性_ > showSeriesName|该布尔值表示数据标签系列名称是否可见。|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_属性_ > showValue|该布尔值表示数据标签值是否可见。|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_属性_ > height|表示图表上的图例高度。|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_属性_ > left|表示图表图例左侧。|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_属性_ > showShadow|表示图表上的图例是否有阴影。|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_属性_ > top|表示图表图例顶部。|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_属性_ > width|表示图表上的图例宽度。|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_关系_ > legendEntries|表示图例中 legendEntries 的集合。 只读。|1.7|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_属性_ > visible|表示图表图例条目可见。|1.7|
|[chartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|_属性_ > items|ChartLegendEntry 对象的集合。 只读。|1.7|
|[chartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|_方法_ > getCount()|返回集合中的 legendEntry 数量。|1.7|
|[chartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|_方法_ > getItemAt(index: number)|返回给定索引处的 legendEntry。|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_属性_ > hasDataLabel|表示数据点是否具有数据标签。 不适用于曲面图。|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_属性_ > markerBackgroundColor|表示数据点的标记背景色的 HTML 颜色代码。 例如， #FF0000 代表红色。|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_属性_ > markerForegroundColor|表示数据点的标记前景色的 HTML 颜色代码。 例如， #FF0000 代表红色。|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_属性_ > markerSize|表示数据点的标记大小。|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_属性_ > markerStyle|表示图表数据点的标记样式。 可能的值包括：Invalid、Automatic、None、Square、Diamond、Triangle、X、Star、Dot、Dash、Circle、Plus、Picture。|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_关系_ > dataLabel|返回图表点的数据标签。 只读。|1.7|
|[chartPointFormat](/javascript/api/excel/excel.chartpointformat)|_关系_ > border|表示图表数据点的边框格式，包括颜色、样式和粗细信息。 只读。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_属性_ > chartType|表示系列的图表类型。 可能的值包括：ColumnClustered、ColumnStacked、ColumnStacked100、BarClustered、BarStacked、BarStacked100、LineStacked、LineStacked100、LineMarkers、LineMarkersStacked、LineMarkersStacked100、PieOfPie 等。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_属性_ > doughnutHoleSize|表示图表系列的圆环孔大小。  仅对圆环图和分离型圆环图有效。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_属性_ > filtered|该布尔值表示是否筛选系列。 不适用于曲面图。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_属性_ > gapWidth|表示图表系列的间隙宽度。  有效对象：条形图和柱形图，以及|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_属性_ > hasDataLabels|该布尔值表示系列是否具有数据标签。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_属性_ > markerBackgroundColor|表示图表系列的标记背景色。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_属性_ > markerForegroundColor|表示图表系列的标记前景色。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_属性_ > markerSize|表示图表系列的标记大小。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_属性_ > markerStyle|表示图表系列的标记类型。 可能的值包括：Invalid、Automatic、None、Square、Diamond、Triangle、X、Star、Dot、Dash、Circle、Plus、Picture。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_属性_ > plotOrder|表示图表组中某个图表系列的绘制顺序。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_属性_ > showShadow|该布尔值表示系列是否具有阴影。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_属性_ > smooth|该布尔值表示系列是否平滑。 仅适用于折线图和散点图。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_关系_ > dataLabels|表示系列中所有数据标签的集合。 只读。|ApiSet.InProgressFeatures.ChartingAPI|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_关系_ > trendlines|表示系列中趋势线的集合。 只读。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_方法_ > delete()|删除 chart series 对象。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_方法_ > setBubbleSizes(sourceData: Range)|设置图表系列的气泡大小。 仅适用于气泡图。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_方法_ > setValues(sourceData: Range)|设置图表系列的值。 对于散点图，它表示 Y 轴的值。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_方法_ > setXAxisValues(sourceData: Range)|设置图表系列 X 轴的值。 仅适用于散点图。|1.7|
|[chartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|_方法_ > add(name: string, index: number)|向集合添加新系列。|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_属性_ > height|返回图表标题的高度，以磅为单位。 只读。 如果图表标题不可见，则为 Null。 只读。|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_属性_ > horizontalAlignment|表示图表标题水平对齐。 可能的值包括：Center、Left、Justify、Distributed、Right。|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_属性_ > left|表示图表标题左边缘到图表区域左边缘的距离，以磅为单位。 如果图表标题不可见，则为 Null。|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_属性_ > position|表示图表标题的位置。 可能的值包括：Top、Automatic、Bottom、Right、Left。|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_属性_ > showShadow|表示一个布尔值，用于确定图表标题是否具有阴影。|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_属性_ > textOrientation|表示图表标题的文本方向。 此值应是 -90 到 90 或 180（垂直文本）之间的整数。|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_属性_ > top|表示图表标题上边缘到图表区域顶部的距离，以磅为单位。 如果图表标题不可见，则为 Null。|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_属性_ > verticalAlignment|表示图表标题垂直对齐。 可能的值包括：Center、Bottom、Top、Justify、Distributed。|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_属性_ > width|返回图表标题的宽度，以磅为单位。 只读。 如果图表标题不可见，则为 Null。 只读。|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_方法_ > setFormula(formula: string)|设置一个字符串值，用于表示采用 A1 表示法的图表标题的公式。|1.7|
|[chartTitleFormat](/javascript/api/excel/excel.charttitleformat)|_关系_ > border|表示图表标题的边框格式，包括颜色、线条样式和粗细。 只读。|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_属性_ > backward|表示趋势线向后延伸的周期数。|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_属性_ > displayEquation|如果图表上显示趋势线公式，则为 True。|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_属性_ > displayRSquared|如果图表上显示趋势线的 R 平方值，则为 True。|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_属性_ > forward|表示趋势线向前延伸的周期数。|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_属性_ > intercept|表示趋势线的截距值。 可以设置为数字值或空字符串（对于自动值）。 返回的值始终为数字。|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_属性_ > movingAveragePeriod|表示图表趋势线的周期，仅适用于 MovingAverage 类型的趋势线。|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_属性_ > name|表示趋势线的名称。 可设为字符串值，或者设为 Null 值（表示自动值）。 返回的值始终为字符串|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_属性_ > polynomialOrder|表示图表趋势线的顺序，仅适用于 Polynomial 类型的趋势线。|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_属性_ > type|表示图表趋势线的类型。 可能的值包括：Linear、Exponential、Logarithmic、MovingAverage、Polynomial、Power。|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_关系_ > format|表示图表趋势线的格式。 只读。|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_方法_ > delete()|删除 Trendline 对象。|1.7|
|[chartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|_属性_ > items|chartTrendline 对象的集合。 只读。|1.7|
|[chartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|_方法_ > add(type: string)|向趋势线集合添加新的趋势线。|1.7|
|[chartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|_方法_ > getCount()|返回集合中的趋势线数量。|1.7|
|[chartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|_方法_ > getItem(index: number)|按索引（在项目数组中的插入顺序）获取 Trendline 对象。|1.7|
|[chartTrendlineFormat](/javascript/api/excel/excel.charttrendlineformat)|_关系_ > line|表示图表线条格式。只读。|1.7|
|[customProperty](/javascript/api/excel/excel.customproperty)|_属性_ > key|获取 customProperty 的键。只读。只读。|1.7|
|[customProperty](/javascript/api/excel/excel.customproperty)|_属性_ > type|获取自定义属性的值类型。 只读。 只读。 可能的值包括：Number、Boolean、Date、String、Float。|1.7|
|[customProperty](/javascript/api/excel/excel.customproperty)|_属性_ > value|获取或设置自定义属性的值。|1.7|
|[customProperty](/javascript/api/excel/excel.customproperty)|_方法_ > delete()|删除 custom property 对象。|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_属性_ > items|customProperty 对象的集合。只读。|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_方法_ > add(key: string, value: object)|新建自定义属性或设置现有自定义属性。|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_方法_ > deleteAll()|删除此集合中的所有自定义属性。|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_方法_ > getCount()|获取自定义属性的计数。|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_方法_ > getItem(key: string)|按键获取 custom property 对象（不区分大小写）。当不存在自定义属性时引发。|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_方法_ > getItemOrNullObject(key: string)|按键获取 custom property 对象（不区分大小写）。如果不存在自定义属性，则返回 null 对象。|1.7|
|[dataConnectionCollection](/javascript/api/excel/excel.dataconnectioncollection)|_属性_ > items|dataConnection 对象的集合。 只读。|1.7|
|[dataConnectionCollection](/javascript/api/excel/excel.dataconnectioncollection)|_方法_ > refreshAll()|刷新集合中的所有数据连接。|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_属性_ > author|获取或设置工作簿的作者。|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_属性_ > category|获取或设置工作簿的类别。|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_属性_ > comments|获取或设置工作簿的注释。|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_属性_ > company|获取或设置工作簿的公司。|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_属性_ > keywords|获取或设置工作簿的关键字。|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_属性_ > lastAuthor|获取工作簿的最终作者。 只读。 只读。|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_属性_ > manager|获取或设置工作簿的管理者。|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_属性_ > revisionNumber|获取工作簿的修订号。 只读。|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_属性_ > subject|获取或设置工作簿的主题。|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_属性_ > title|获取或设置工作簿的标题。|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_关系_ > creationDate|获取工作簿的创建日期。 只读。 只读。|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_关系_ > custom|获取工作簿的自定义属性的集合。 只读。 只读。|1.7|
|[namedItem](/javascript/api/excel/excel.nameditem)|_属性_ > formula|获取或设置的已命名项目的公式。  公式始终以等号 (=) 开头。|1.7|
|[namedItem](/javascript/api/excel/excel.nameditem)|_关系_ > arrayValues|返回包含已命名项目的值和类型的对象。 只读。|1.7|
|[namedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|_属性_ > types|表示已命名项目数组中每个项目的类型。只读。 可能的值包括：Unknown、Empty、String、Integer、Double、Boolean、Error。|1.7|
|[namedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|_属性_ > values|表示已命名项目数组中每个项目的值。 只读。|1.7|
|[range](/javascript/api/excel/excel.range)|_属性_ > isEntireColumn|表示当前区域是否为整列。 只读。|1.7|
|[range](/javascript/api/excel/excel.range)|_属性_ > isEntireRow|表示当前区域是否为整行。 只读。|1.7|
|[range](/javascript/api/excel/excel.range)|_属性_ > numberFormatLocal|表示 Excel 中的给定区域的数字格式代码，以用户语言的字符串表示。|1.7|
|[range](/javascript/api/excel/excel.range)|_属性_ > style|表示当前区域的样式。 它将返回 null 或字符串。|1.7|
|[range](/javascript/api/excel/excel.range)|_方法_ > getAbsoluteResizedRange(numRows: number, numColumns: number)|获取一个 Range 对象，该对象的左上单元格与当前 Range 对象相同，但具有指定的行数和列数。|1.7|
|[range](/javascript/api/excel/excel.range)|_方法_ > getImage()|将区域呈现为 base64 编码图像。|1.7|
|[range](/javascript/api/excel/excel.range)|_方法_ > getSurroundingRegion()|返回一个 Range 对象，该对象表示此区域左上单元格的周围区域。 周围区域是由相对于该区域的空白行和空白列的任何组合所限定的区域。|1.7|
|[range](/javascript/api/excel/excel.range)|_方法_ > showCard()|显示活动单元格的卡片（如果该单元格具有富值内容）。|1.7|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_属性_ > textOrientation|获取或设置区域内的所有单元格的文本方向。|1.7|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_属性_ > useStandardHeight|确定 Range 对象的行高是否等于工作表的标准行高。|1.7|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_属性_ > useStandardWidth|确定 Range 对象的列宽是否等于工作表的标准列宽。|1.7|
|[rangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|_属性_ > address|表示超链接的 URL 目标。|1.7|
|[rangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|_属性_ > document..|表示超链接的 文档目标。|1.7|
|[rangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|_属性_ > screenTip|表示鼠标悬停在超链接上时显示的字符串。|1.7|
|[rangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|_属性_ > textToDisplay|表示区域最左上方单元格中显示的字符串。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_ > addIndent|指示将单元格中的文本对齐方式设为相等分布时文本是否会自动缩进。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_ > autoIndent|指示将单元格中的文本对齐方式设为相等分布时文本是否会自动缩进。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_ > builtIn|指示样式是否为内置样式。 只读。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_ > formulaHidden|指示工作表受保护时是否隐藏公式。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_ > horizontalAlignment|表示样式水平对齐。 可能的值包括：General、Left、Center、Right、Fill、Justify、CenterAcrossSelection、Distributed。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_ > includeAlignment|指示样式是否包含 AutoIndent、HorizontalAlignment、VerticalAlignment、WrapText、IndentLevel 和 TextOrientation 属性。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_ > includeBorder|指示样式是否包含 Color、ColorIndex、LineStyle 和 Weight 边框属性。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_ > includeFont|指示样式是否包含 Background、Bold、Color、ColorIndex、FontStyle、Italic、Name、Size、Strikethrough、Subscrip、Superscript 和 Underline 字体属性。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_ > includeNumber|指示样式是否包含 NumberFormat 属性。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_ > includePatterns|指示样式是否包含 Color、ColorIndex、InvertIfNegative、Pattern、PatternColor 和 PatternColorIndex 内部属性。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_ > includeProtection|指示样式是否包含 FormulaHidden 和 Locked 保护属性。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_ > indentLevel|0 到 250 之间的一个整数，指示样式的缩进水平。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_ > locked|指示工作表受保护时是否锁定对象。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_ > name|样式的名称。 只读。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_ > numberFormat|样式中数字格式的格式代码。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_ > numberFormatLocal|样式中数字格式的本地化格式代码。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_ > orientation|此样式中的文本方向。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_ > readingOrder|样式中的阅读顺序。 可能的值包括：Context、LeftToRight、RightToLeft。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_ > shrinkToFit|指示文本是否自动缩小以适合可用列宽。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_ > textOrientation|此样式中的文本方向。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_ > verticalAlignment|表示样式的垂直对齐方式。 可能的值包括：Top、Center、Bottom、Justify、Distributed。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_ > wrapText|指示 Microsoft Excel 是否将对象中的文本换行。|1.7|
|[style](/javascript/api/excel/excel.style)|_关系_ > borders|四个 Border 对象的 Border 集合，表示四个边框的样式。 只读。|1.7|
|[style](/javascript/api/excel/excel.style)|_关系_ > fill|样式的填充。 只读。|1.7|
|[style](/javascript/api/excel/excel.style)|_关系_ > font|该 Font 对象表示样式的字体。 只读。|1.7|
|[style](/javascript/api/excel/excel.style)|_方法_ > delete()|删除此样式。|1.7|
|[styleCollection](/javascript/api/excel/excel.stylecollection)|_属性_ > items|style 对象的集合。 只读。|1.7|
|[styleCollection](/javascript/api/excel/excel.stylecollection)|_方法_ > add(name: string)]|向集合添加新样式。|1.7|
|[styleCollection](/javascript/api/excel/excel.stylecollection)|_方法_ > getItem(name: string)|按名称获取样式。|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_属性_ > address|获取地址，该地址表示特定工作表上的表格的更改区域。|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_属性_ > changeType|获取更改类型，该类型表示 Changed 事件的触发方式。 可能的值包括：Others、RangeEdited、RowInserted、RowDeleted、ColumnInserted、ColumnDeleted、CellInserted、CellDeleted。|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_属性_ > source|获取事件源。 可能的值包括：Local、Remote。|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_属性_ > tableId|获取其中的数据发生更改的表格的 ID。|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_属性_ > type|获取事件的类型。 可能的值包括：WorksheetDataChanged、WorksheetSelectionChanged、WorksheetAdded、WorksheetActivated、WorksheetDeactivated、TableDataChanged、TableSelectionChanged、WorksheetDeleted。|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_属性_ > worksheetId|获取其中的数据发生更改的工作表的 ID。|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_属性_ > address|获取区域地址，该地址表示特定工作表上的表格选定区域。|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_属性_ > isInsideTable|指示选定区域是否在表格内，如果 IsInsideTable 为 false，则地址无效。|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_属性_ > tableId|获取其中的选定区域发生更改的表格 ID。|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_属性_ > type|获取事件的类型。 可能的值包括：WorksheetDataChanged、WorksheetSelectionChanged、WorksheetAdded、WorksheetActivated、WorksheetDeactivated、TableDataChanged、TableSelectionChanged、WorksheetDeleted。|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_属性_ > worksheetId|获取其中的选定区域发生更改的工作表的 ID。|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_属性_ > name|获取工作簿名称。 只读。|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_关系_ > dataConnections|刷新工作簿中的所有数据连接。 只读。|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_关系_ > properties|获取工作簿属性。 只读。|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_关系_ > protection|返回工作簿的工作簿保护对象。 只读。|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_关系_ > styles|表示与工作簿关联的样式的集合。 只读。|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_方法_ > getActiveCell()|获取工作簿中当前处于活动状态的单元格。|1.7|
|[workbookProtection](/javascript/api/excel/excel.workbookprotection)|_属性_ > protected|指示工作簿是否受保护。 只读。 只读。|1.7|
|[workbookProtection](/javascript/api/excel/excel.workbookprotection)|_方法_ > protect(password: string)|保护工作簿。 如果工作簿处于受保护状态，则无法执行此方法。|1.7|
|[workbookProtection](/javascript/api/excel/excel.workbookprotection)|_方法_ > unprotect(password: string)|解除保护工作簿。|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_属性_ > gridlines|获取或设置工作表的网格线标志。|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_属性_ > headings|获取或设置工作表的标题标志。|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_属性_ > showHeadings|获取或设置工作表的标题标志。|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_属性_ > standardHeight|返回工作表中所有行的标准（默认）行高，以磅为单位。 只读。|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_属性_ > standardWidth|返回或设置工作表中所有列的标准（默认）列宽。|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_属性_ > tabColor|获取或设置工作表标签颜色。|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_关系_ > freezePanes|获取可用于控制工作表上的冻结窗格的对象。只读。|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_方法_ > copy(positionType: WorksheetPositionType, relativeTo: Worksheet)|复制工作表并将其置于指定位置。 返回复制的工作表。|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_方法_ > getRangeByIndexes(startRow: number, startColumn: number, rowCount: number, columnCount: number)|获取以特定行索引和列索引开始并跨越了一定数量的行和列的 range 对象。|1.7|
|[worksheetActivatedEventArgs](/javascript/api/excel/excel.worksheetactivatedeventargs)|_属性_ > type|获取事件的类型。 可能的值包括：WorksheetDataChanged、WorksheetSelectionChanged、WorksheetAdded、WorksheetActivated、WorksheetDeactivated、TableDataChanged、TableSelectionChanged、WorksheetDeleted。|1.7|
|[worksheetActivatedEventArgs](/javascript/api/excel/excel.worksheetactivatedeventargs)|_属性_ > worksheetId|获取已启用的工作表的 ID。|1.7|
|[worksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|_属性_ > source|获取事件源。 可能的值包括：Local、Remote。|1.7|
|[worksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|_属性_ > type|获取事件的类型。 可能的值包括：WorksheetDataChanged、WorksheetSelectionChanged、WorksheetAdded、WorksheetActivated、WorksheetDeactivated、TableDataChanged、TableSelectionChanged、WorksheetDeleted。|1.7|
|[worksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|_属性_ > worksheetId|获取已添加至工作簿的工作表的 ID。|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_属性_ > address|获取区域地址，该地址表示特定工作表上的更改区域。|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_属性_ > changeType|获取更改类型，该类型表示 Changed 事件的触发方式。 可能的值包括：Others、RangeEdited、RowInserted、RowDeleted、ColumnInserted、ColumnDeleted、CellInserted、CellDeleted。|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_属性_ > source|获取事件源。 可能的值包括：Local、Remote。|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_属性_ > type|获取事件的类型。 可能的值包括：WorksheetDataChanged、WorksheetSelectionChanged、WorksheetAdded、WorksheetActivated、WorksheetDeactivated、TableDataChanged、TableSelectionChanged、WorksheetDeleted。|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_属性_ > worksheetId|获取其中的数据发生更改的工作表的 ID。|1.7|
|[worksheetDeactivatedEventArgs](/javascript/api/excel/excel.worksheetdeactivatedeventargs)|_属性_ > type|获取事件的类型。 可能的值包括：WorksheetDataChanged、WorksheetSelectionChanged、WorksheetAdded、WorksheetActivated、WorksheetDeactivated、TableDataChanged、TableSelectionChanged、WorksheetDeleted。|1.7|
|[worksheetDeactivatedEventArgs](/javascript/api/excel/excel.worksheetdeactivatedeventargs)|_属性_ > worksheetId|获取已停用的工作表的 ID。|1.7|
|[worksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|_属性_ > source|获取事件源。 可能的值包括：Local、Remote。|1.7|
|[worksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|_属性_ > type|获取事件的类型。 可能的值包括：WorksheetDataChanged、WorksheetSelectionChanged、WorksheetAdded、WorksheetActivated、WorksheetDeactivated、TableDataChanged、TableSelectionChanged、WorksheetDeleted。|1.7|
|[worksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|_属性_ > worksheetId|获取已从工作簿删除的工作表的 ID。|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_方法_ > freezeAt(frozenRange: Range or string)|设置活动工作表视图中的冻结单元格。|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_方法_ > freezeColumns(count: number)|就地冻结工作表的第一列。|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_方法_ > freezeRows(count: number)|就地冻结工作表的顶行。|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_方法_ > getLocation()|获取用于描述活动工作表视图中的冻结单元格的区域。|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_方法_ > getLocationOrNullObject()|获取用于描述活动工作表视图中的冻结单元格的区域。|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_方法_ > unfreeze()|移除工作表中的所有冻结窗格。|1.7|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_属性_ > allowEditObjects|表示允许编辑对象的工作表保护选项。|1.7|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_属性_ > allowEditScenarios|表示允许编辑应用场景的工作表保护选项。|1.7|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_关系_ > selectionMode|表示选择模式的工作表保护选项。|1.7|
|[worksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|_属性_ > address|获取区域地址，该地址表示特定工作表上的选定区域。|1.7|
|[worksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|_属性_ > type|获取事件的类型。 可能的值包括：WorksheetDataChanged、WorksheetSelectionChanged、WorksheetAdded、WorksheetActivated、WorksheetDeactivated、TableDataChanged、TableSelectionChanged、WorksheetDeleted。|1.7|
|[worksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|_属性_ > worksheetId|获取其中的选定区域发生更改的工作表的 ID。|1.7|


## <a name="whats-new-in-excel-javascript-api-16"></a>Excel JavaScript API 1.6 的最近更新 

### <a name="conditional-formatting"></a>条件格式

引入了区域的条件格式。 允许以下条件格式类型：

* 色阶
* 数据栏
* 图标集
* 自定义

此外：

* 返回应用条件格式的区域。 
* 删除条件格式。 
* 提供优先级和 stopifTrue 功能。 
* 获取给定区域内所有条件格式的集合。 
* 清除当前指定区域中处于活动状态的所有条件格式。 

|对象| 最近更新| 说明|要求集|
|:----|:----|:----|:----|
|[application](/javascript/api/excel/excel.application)|_方法_ > suspendApiCalculationUntilNextSync()|在下一次调用“context.sync()”前暂停计算。设置后，开发者负责重新计算工作簿，以确保传播所有依赖项。|1.6|
|[cellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)|_关系_ > format|返回 format 对象，该对象用于封装条件格式字体、填充、边框和其他属性。只读。|1.6|
|[cellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)|_关系_ > rule|表示此条件格式中的 Rule 对象。|1.6|
|[colorScaleConditionalFormat](/javascript/api/excel/excel.colorscaleconditionalformat)|_属性_ > threeColorScale|如果为 true，则色阶有三个点（最小、中点、最大），否则将有两个点（最小、最大）。只读。|1.6|
|[colorScaleConditionalFormat](/javascript/api/excel/excel.colorscaleconditionalformat)|_关系_ > criteria|色阶的条件。使用两点色阶时，中点可选。|1.6|
|[conditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|_属性_ > formula1|如果需要，公式可对条件格式规则进行求值。|1.6|
|[conditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|_属性_ > formula2|如果需要，公式可对条件格式规则进行求值。|1.6|
|[conditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|_属性_ > operator|文本条件格式的运算符。可能的值包括：Invalid、Between、NotBetween、EqualTo、NotEqualTo、GreaterThan、LessThan、GreaterThanOrEqual、LessThanOrEqual。|1.6|
|[conditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|_关系_ > maximum|最大点色阶条件。|1.6|
|[conditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|_关系_ > midpoint|色阶为 3 色阶时的中点色阶条件。|1.6|
|[conditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|_关系_ > minimum|最小点色阶条件。|1.6|
|[conditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|_属性_ > color|色阶颜色的 HTML 颜色代码表示。例如，#FF0000 代表红色。|1.6|
|[conditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|_属性_ > formula|数字、公式或 null（如果类型为 LowestValue）。|1.6|
|[conditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|_属性_ > type|图标条件公式的依据。可能的值包括：Invalid、LowestValue、HighestValue、Number、Percent、Formula、Percentile。|1.6|
|[conditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|_属性_ > borderColor|表示窗体 #RRGGBB（例如“FFA500”）的边框线条颜色或作为已命名的 HTML 颜色（例如“orange”）的 HTML 颜色代码。|1.6|
|[conditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|_属性_ > fillColor|表示窗体 #RRGGBB（例如“FFA500”）的填充颜色或已命名 HTML 颜色（例如“orange”）的 HTML 颜色代码。|1.6|
|[conditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|_属性_ > matchPositiveBorderColor|该布尔值表示负 DataBar 是否与正 DataBar 具有相同边框颜色。|1.6|
|[conditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|_属性_ > matchPositiveFillColor|该布尔值表示负 DataBar 是否与正 DataBar 具有相同填充颜色。|1.6|
|[conditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|_属性_ > borderColor|表示窗体 #RRGGBB（例如“FFA500”）的边框线条颜色或作为已命名的 HTML 颜色（例如“orange”）的 HTML 颜色代码。|1.6|
|[conditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|_属性_ > fillColor|表示窗体 #RRGGBB（例如“FFA500”）的填充颜色或已命名 HTML 颜色（例如“orange”）的 HTML 颜色代码。|1.6|
|[conditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|_属性_ > gradientFill|该布尔值表示 DataBar 是否具有渐变。|1.6|
|[conditionalDataBarRule](/javascript/api/excel/excel.conditionaldatabarrule)|_属性_ > formula|如果需要，公式可对 databar 规则进行求值。|1.6|
|[conditionalDataBarRule](/javascript/api/excel/excel.conditionaldatabarrule)|_属性_ > type|数据栏的规则类型。可能的值包括：LowestValue、HighestValue、Number、Percent、Formula、Percentile、Automatic。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_属性_ > id|当前 ConditionalFormatCollection 内的条件格式的优先级。 只读。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_属性_ > priority|优先级（或索引）位于此条件格式当前存在的条件格式集合内。更改此属性也会|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_属性_ > stopIfTrue|如果满足此条件格式的条件，则不会有任何低优先级格式应在此单元格上生效。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_属性_ > type|条件格式的类型。一次只能设置一种类型。只读。只读。可能的值包括：Custom、DataBar、ColorScale、IconSet。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_关系_ > cellValue|如果当前的条件格式是 CellValue 类型，则返回单元值条件格式属性。只读。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_关系_ > cellValueOrNullObject|如果当前的条件格式是 CellValue 类型，则返回单元值条件格式属性。只读。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_关系_ > colorScale|如果当前的条件格式为 ColorScale 类型，返回 ColorScale 条件格式属性。只读。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_关系_ > colorScaleOrNullObject|如果当前的条件格式为 ColorScale 类型，返回 ColorScale 条件格式属性。只读。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_关系_ > custom|如果当前的条件格式是自定义类型，则返回自定义条件格式属性。只读。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_关系_ > customOrNullObject|如果当前的条件格式是自定义类型，则返回自定义条件格式属性。只读。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_关系_ > dataBar|如果当前的条件格式是数据栏，则返回数据栏属性。只读。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_关系_ > dataBarOrNullObject|如果当前的条件格式是数据栏，则返回数据栏属性。只读。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_关系_ > iconSet|如果当前的条件格式是 IconSet 类型，则返回 IconSet 条件格式属性。只读。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_关系_ > iconSetOrNullObject|如果当前的条件格式是 IconSet 类型，则返回 IconSet 条件格式属性。只读。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_关系_ > preset|返回预设条件的条件格式，如上述 averagebelow averageunique valuescontains blanknonblankerrornoerror 属性。只读。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_关系_ > presetOrNullObject|返回预设条件的条件格式，如上述 averagebelow averageunique valuescontains blanknonblankerrornoerror 属性。只读。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_关系_ > textComparison|如果当前的条件格式是文本类型，则返回特定文本条件格式属性。只读。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_关系_ > textComparisonOrNullObject|如果当前的条件格式是文本类型，则返回特定文本条件格式属性。只读。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_关系_ > topBottom|如果当前的条件格式是 TopBottom 类型，则返回 TopBottom 条件格式属性。只读。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_关系_ > topBottomOrNullObject|如果当前的条件格式是 TopBottom 类型，则返回 TopBottom 条件格式属性。只读。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_方法_ > delete()|删除此条件格式。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_方法_ > getRange()|返回条件格式应用的区域，如果区域不连续，则返回 NULL 对象。只读。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_方法_ > getRangeOrNullObject()|返回条件格式应用的区域，如果区域不连续，则返回 NULL 对象。只读。|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_属性_ > items|ConditionalFormat 对象的集合。只读。|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_方法_ > add(type: string)|向 firsttop 优先级的集合添加新的条件格式。|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_方法_ > clearAll()|清除当前指定区域中处于活动状态的所有条件格式。|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_方法_ > getCount()|返回工作簿中的条件格式数。只读。|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_方法_ > getItem(id: string)|返回给定 ID 的条件格式。|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_方法_ > getItemAt(index: number)|返回给定索引处的条件格式。|1.6|
|[conditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|_属性_ > formula|如果需要，公式可对条件格式规则进行求值。|1.6|
|[conditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|_属性_ > formulaLocal|如果需要，公式可采用用户的语言对条件格式规则进行求值。|1.6|
|[conditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|_属性_ > formulaR1C1|如果需要，公式可采用 R1C1 表示法对条件格式规则进行求值。|1.6|
|[conditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|_属性_ > formula|取决于类型的数字或公式。|1.6|
|[conditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|_属性_ > operator|Icon 条件格式的每个规则类型的 GreaterThan 或 GreaterThanOrEqual。可能的值包括：Invalid、GreaterThan、GreaterThanOrEqual。|1.6|
|[conditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|_关系_ > customIcon|如果与默认 IconSet 不同，返回当前条件的自定义图标，否则将返回 null。|1.6|
|[conditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|_关系_ > type|应基于的图标条件公式。|1.6|
|[conditionalPresetCriteriaRule](/javascript/api/excel/excel.conditionalpresetcriteriarule)|_属性_ > criterion|条件格式的条件。可能的值是：Invalid、Blanks、NonBlanks、Errors、NonErrors、Yesterday、Today、Tomorrow、LastSevenDays、LastWeek、ThisWeek、NextWeek、LastMonth、ThisMonth、NextMonth、AboveAverage、BelowAverage、EqualOrAboveAverage、EqualOrBelowAverage、OneStdDevAboveAverage、OneStdDevBelowAverage、TwoStdDevAboveAverage、TwoStdDevBelowAverage、ThreeStdDevAboveAverage、ThreeStdDevBelowAverage、UniqueValues、DuplicateValues。|1.6|
|[conditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|_属性_ > color|表示窗体 #RRGGBB（例如“FFA500”）的边框线条颜色或作为已命名的 HTML 颜色（例如“orange”）的 HTML 颜色代码。|1.6|
|[conditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|_属性_ > id|表示边框标识符。只读。可能的值是：EdgeTop、EdgeBottom、EdgeLeft、EdgeRight。|1.6|
|[conditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|_属性_ > sideIndex|指示边框的特定边的常量值。只读。可能的值是：EdgeTop、EdgeBottom、EdgeLeft、EdgeRight。|1.6|
|[conditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|_属性_ > style|线条样式的常量之一，指定边框的线条样式。可能的值是：None、Continuous、Dash、DashDot、DashDotDot、Dot、Double、SlantDashDot。|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_属性_ > count|集合中的 border 对象数量。只读。|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_属性_ > items|conditionalRangeBorder 对象的集合。只读。|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_关系_ > bottom|获取只读的顶部边框。|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Relationship_ > left|获取只读的顶部边框。|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_关系_ > right|获取只读的顶部边框。|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_关系_ > top|获取只读的顶部边框。|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_方法_ > getItem(index: string)|使用其名称获取 border 对象|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_方法_ > getItemAt(index: number)|使用其索引获取 border 对象|1.6|
|[conditionalRangeFill](/javascript/api/excel/excel.conditionalrangefill)|_属性_ > color|表示窗体 #RRGGBB（例如“FFA500”）的填充颜色或已命名 HTML 颜色（例如“orange”）的 HTML 颜色代码。|1.6|
|[conditionalRangeFill](/javascript/api/excel/excel.conditionalrangefill)|_方法_ > clear()|重置填充。|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_属性_ > bold|表示字体的加粗状态。|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_属性_ > color|文本颜色的 HTML 颜色代码表示。例如，#FF0000 表示红色。|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_属性_ > italic|表示字体的斜体状态。|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_属性_ > strikethrough|表示字体的删除线状态。|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_属性_ > underline|应用于字体的下划线类型。可能的值是：None、Single、Double。|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_方法_ > clear()|重置字体格式。|1.6|
|[conditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|_属性_ > numberFormat|表示 Excel 中指定范围的数字格式代码。当传递 null 时清除。|1.6|
|[conditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|_关系_ > borders|应用于整个条件格式范围的 border 对象的集合。只读。|1.6|
|[conditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|_关系_ > fill|返回在整个条件格式范围内定义的 fill 对象。只读。|1.6|
|[conditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|_关系_ > font|返回在整个条件格式范围内定义的 font 对象。只读。|1.6|
|[conditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule)|_属性_ > operator|本文条件格式的运算符。可能的值包括：Invalid、Contains、NotContains、BeginsWith、EndsWith。|1.6|
|[conditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule)|_属性_ > text|条件格式的文本值。|1.6|
|[conditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|_属性_ > rank|1 和 1000 之间的数字排名或 1 和 100 之间的百分比排名。|1.6|
|[conditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|_属性_ > type|基于排名第一或排名最后的格式值。可能的值包括：Invalid、TopItems、TopPercent、BottomItems、BottomPercent。|1.6|
|[customConditionalFormat](/javascript/api/excel/excel.customconditionalformat)|_关系_ > format|返回 format 对象，该对象用于封装条件格式字体、填充、边框和其他属性。只读。|1.6|
|[customConditionalFormat](/javascript/api/excel/excel.customconditionalformat)|_关系_ > rule|表示此条件格式中的 Rule 对象。只读。|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_属性_ > axisColor|表示窗体 #RRGGBB（例如 "FFA500"）的轴行颜色或作为已命名的 HTML 颜色（例如 "orange"）的 HTML 颜色代码。|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_属性_ > axisFormat|如何确定 Excel 数据栏的轴的表示形式。可能的值包括：Automatic、None、CellMidPoint。|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_属性_ > barDirection|表示数据栏图形应遵循的方向。可能的值包括：Context、LeftToRight、RightToLeft。|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_属性_ > showDataBarOnly|如果为 true，则对应用数据栏的单元格隐藏值。|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_关系_ > lowerBoundRule|构成数据栏的下限（以及如何计算，如果适用）的规则。|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_关系_ > negativeFormat|Excel 数据栏中轴左侧的所有值的表示形式。只读。。|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_关系_ > positiveFormat|Excel 数据栏中轴右侧的所有值的表示形式。只读。。|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_关系_ > upperBoundRule|构成数据栏的上限（以及如何计算，如果适用）的规则。|1.6|
|[iconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|_属性_ > reverseIconOrder|如果为 true，则反转 IconSet 的图标顺序。注意，如果使用自定义图标，则不能进行设置。|1.6|
|[iconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|_属性_ > showIconOnly|如果为 true，则隐藏值并仅显示图标。|1.6|
|[iconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|_属性_ > style|如果设置，则显示条件格式的 IconSet 选项。可能的值是：Invalid、ThreeArrows、ThreeArrowsGray、ThreeFlags、ThreeTrafficLights1、ThreeTrafficLights2、ThreeSigns、ThreeSymbols、ThreeSymbols2、FourArrows、FourArrowsGray、FourRedToBlack、FourRating、FourTrafficLights、FiveArrows、FiveArrowsGray、FiveRating、FiveQuarters、ThreeStars、ThreeTriangles、FiveBoxes。|1.6|
|[iconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|_关系_ > criteria|规则的 Criteria 和 IconSet 数组，以及条件图标的潜在自定义图标。注意，对于第一个条件，只能修改自定义图标，类型、公式和运算符在设置时将忽略。|1.6|
|[presetCriteriaConditionalFormat](/javascript/api/excel/excel.presetcriteriaconditionalformat)|_关系_ > format|返回 format 对象，该对象用于封装条件格式字体、填充、边框和其他属性。只读。|1.6|
|[presetCriteriaConditionalFormat](/javascript/api/excel/excel.presetcriteriaconditionalformat)|_关系_ > rule|条件格式的规则。|1.6|
|[range](/javascript/api/excel/excel.range)|_关系_ > conditionalFormats|区域交叉的 ConditionalFormats 的集合。只读。|1.6|
|[range](/javascript/api/excel/excel.range)|_方法_ > calculate()|计算工作表上的单元格区域。|1.6|
|[textConditionalFormat](/javascript/api/excel/excel.textconditionalformat)|_关系_ > format|返回 format 对象，该对象用于封装条件格式字体、填充、边框和其他属性。只读。|1.6|
|[textConditionalFormat](/javascript/api/excel/excel.textconditionalformat)|_关系_ > rule|条件格式的规则。|1.6|
|[topBottomConditionalFormat](/javascript/api/excel/excel.topbottomconditionalformat)|_关系_ > format|返回 format 对象，该对象用于封装条件格式字体、填充、边框和其他属性。只读。|1.6|
|[topBottomConditionalFormat](/javascript/api/excel/excel.topbottomconditionalformat)|_关系_ > rule|表示 TopBottom 条件格式的条件。|1.6|
|[workbook](/javascript/api/excel/excel.workbook)|_关系_ > internalTest|仅供内部使用。只读。|1.6|
|[worksheet](/javascript/api/excel/excel.worksheet)|_方法_ > calculate(markAllDirty: bool)|计算工作表上的所有单元格。|1.6|

##  <a name="whats-new-in-excel-javascript-api-15"></a>Excel JavaScript API 1.5 的最近更新

### <a name="custom-xml-part"></a>自定义 XML 部件

* 将自定义 XML 部件集合添加到工作簿对象中。
* 使用 ID 获取自定义 XML 部件
* 获取其命名空间匹配给定命名空间的自定义 XML 部件的新作用域内集合。
* 获取与某个部件关联的 XML 字符串。
* 提供部件的 ID 和命名空间。
* 向工作簿添加新的自定义 XML 部件。
* 设置整个 XML 部件。
* 删除自定义 XML 部件。
* 删除其给定名称来自由 xpath 标识的元素的属性。
* 按 xpath 查询 XML 内容。
* 插入、更新和删除属性。

**参考实现：** 请参阅[此处](https://github.com/mandren/Excel-CustomXMLPart-Demo)，了解说明如何在外接程序中使用自定义 XML 部件的参考实现。

### <a name="others"></a>其他
* `range.getSurroundingRegion()` 返回一个 Range 对象，该对象表示此范围的周围区域。周围区域是由相对于该范围的空白行和空白列的任何组合所限定的范围。
* 对表列执行 `getNextColumn()`、`getPreviousColumn()` 以及 `getLast() 操作。
* 对工作簿执行 `getActiveWorksheet()` 操作。
* 工作簿的 `getRange(address: string)` 关闭。
* `getBoundingRange(ranges: )` 获取包含提供的范围的最小 range 对象。例如，介于 “B2:C5” 和 “D10:E15” 之间的边界范围为 “B2:E15”。
* 对各种集合（例如已命名项目、工作表、表等）执行 `getCount()` 操作以获取集合中的项目数。 `workbook.worksheets.getCount()`
* 对各种集合（如工作表、表列、图标点、范围视图集合）执行 `getFirst()` 和 `getLast()` 以及 get last 操作。
* 对工作表、表列集合执行 `getNext()` 和 `getPrevious()` 操作。
* `getRangeR1C1()` 获取以特定行索引和列索引开始并跨越了一定数量的行和列的 range 对象。

|对象| 最近更新| 说明|要求集|
|:----|:----|:----|:----|
|[customXmlPart](/javascript/api/excel/excel.customxmlpart)|_属性_ > id|自定义 XML 部件的 ID。只读。|1.5|
|[customXmlPart](/javascript/api/excel/excel.customxmlpart)|_属性_ > namespaceUri|自定义 XML 部件的命名空间 URI。只读。|1.5|
|[customXmlPart](/javascript/api/excel/excel.customxmlpart)|_方法_ > delete()|删除自定义 XML 部件。|1.5|
|[customXmlPart](/javascript/api/excel/excel.customxmlpart)|_方法_ > getXml()|获取自定义 XML 部件的完整 XML 内容。|1.5|
|[customXmlPart](/javascript/api/excel/excel.customxmlpart)|_方法_ > setXml(xml: string)|设置自定义 XML 部件的完整 XML 内容。|1.5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_属性_ > items|customXmlPart 对象的集合。只读。|1.5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_方法_ > add(xml: string)|向工作簿添加新的自定义 XML 部件。|1.5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_方法_ > getByNamespace(namespaceUri: string)|获取其命名空间匹配给定命名空间的自定义 XML 部件的新作用域内集合。|1.5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_方法_ > getCount()|获取此集合中 CustomXml 部件的数量。|1.5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_方法_ > getItem(id: string)|获取基于其 ID 的自定义 XML 部件。|1.5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_方法_ > getItemOrNullObject(id: string)|获取基于其 ID 的自定义 XML 部件。|1.5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_属性_ > items|CustomXmlPartScoped 对象的集合。只读。|1.5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_方法_ > getCount()|获取此集合中 CustomXML 部件的数量。|1.5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_方法_ > getItem(id: string)|获取基于其 ID 的自定义 XML 部件。|1.5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_方法_ > getItemOrNullObject(id: string)|获取基于其 ID 的自定义 XML 部件。|1.5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_方法_ > getOnlyItem()|如果集合仅包含一个项，则此方法返回该项。|1.5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_方法_ > getOnlyItemOrNullObject()|如果集合仅包含一个项，则此方法返回该项。|1.5|
|[workbook](/javascript/api/excel/excel.workbook)|_关系_ > customXmlParts|代表此工作簿包含的自定义 XML 部件的集合。只读。|1.5|
|[worksheet](/javascript/api/excel/excel.worksheet)|_方法_ > getNext(visibleOnly: bool)|获取该工作表之后的工作表。如果该工作表后没有工作表，此方法将引发错误。|1.5|
|[worksheet](/javascript/api/excel/excel.worksheet)|_方法_ > getNextOrNullObject(visibleOnly: bool)|获取该工作表之后的工作表。如果该工作表后没有工作表，此方法将返回 null 值。|1.5|
|[worksheet](/javascript/api/excel/excel.worksheet)|_方法_ > getPrevious(visibleOnly: bool)|获取该工作表之前的工作表。如果该工作表前没有工作表，此方法将引发错误。|1.5|
|[worksheet](/javascript/api/excel/excel.worksheet)|_方法_ > getPreviousOrNullObject(visibleOnly: bool)|获取该工作表之前的工作表。如果该工作表前没有工作表，此方法将返回 null 值。|1.5|
|[worksheetCollection](/javascript/api/excel/excel.worksheetcollection)|_方法_ > getFirst(visibleOnly: bool)|获取集合中的第一个工作表。|1.5|
|[worksheetCollection](/javascript/api/excel/excel.worksheetcollection)|_方法_ > getLast(visibleOnly: bool)|获取集合中的最后一个工作表。|1.5|

## <a name="whats-new-in-excel-javascript-api-14"></a>Excel JavaScript API 1.4 的最近更新
下面介绍了要求集 1.4 中 Excel JavaScript API 的新增内容。

### <a name="named-item-add-and-new-properties"></a>添加了已命名项和新属性

新属性：

* `comment`
* `scope`：限定到工作表或工作簿的项。
* `worksheet`：返回已命名项限定到的工作表。

新方法：

* `add(name: string, reference: Range or string, comment: string)`：将新名称添加到给定范围的集合。
* `addFormulaLocal(name: string, formula: string, comment: string)`：使用用户的公式区域设置，将新名称添加到给定范围的集合。

### <a name="settings-api-in-the-excel-namespace"></a>Excel 命名空间中的设置 API

[Setting](/javascript/api/excel/excel.setting) 对象表示文档保留设置的键值对。 `Excel.Setting` 的功能等同于 `Office.Settings`，但使用批处理 API 语法，而不是通用 API 的回调模型。

API 包括通过键获取设置条目的 `getItem()`，以及将指定键值设置对添加到工作簿的 `add()`。

### <a name="others"></a>其他

* 设置表列名称（旧版只允许读取）。
* 将表列添加到表的末尾（旧版只允许添加到除末尾之外的其他任何位置）。
* 一次性向表中添加多行（旧版只允许一次添加 1 行）。
* `range.getColumnsAfter(count: number)` 和 `range.getColumnsBefore(count: number)` 分别用于获取当前 Range 对象的右/左侧的一定数量的列。
* 获取项或 NULL 对象函数：此功能允许使用键获取对象。如果没有对象，返回的对象的 isNullObject 属性为 true。这样一来，开发者可以检查对象是否存在，而无需通过异常处理来进行处理。适用于工作表、已命名项、绑定、图表系列等

    ```javascript
    worksheet.GetItemOrNullObject()
    ```

|对象| 最近更新| 说明|要求集|
|:----|:----|:----|:----|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_方法_ > getCount()|获取集合中的绑定数量。|1.4|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_方法_ > getItemOrNullObject(id: string)|按 ID 获取 Binding 对象。如果没有 Binding 对象，将返回 NULL 对象。|1.4|
|[chartCollection](/javascript/api/excel/excel.chartcollection)|_方法_ > getCount()|返回工作表中的图表数。|1.4|
|[chartCollection](/javascript/api/excel/excel.chartcollection)|_方法_ > getItemOrNullObject(name: string)|使用图表名称获取图表。如果存在多个名称相同的图表，将返回第一个图表。|1.4|
|[chartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|_方法_ > getCount()|返回系列中的图表点数。|1.4|
|[chartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|_方法_ > getCount()|返回集合中的系列数量。|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_属性_ > comment|表示与此名称相关联的注释。|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_属性_ > scope|指明是否将 name 限定到工作簿或特定工作表。只读。可取值为：Equal、Greater、GreaterEqual、Less、LessEqual、NotEqual。|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_关系_ > worksheet|返回已命名项限定到的工作表。如果项改为限定到工作簿，将引发错误。只读。|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_关系_ > worksheetOrNullObject|返回已命名项限定到的工作表。如果项改为限定到工作簿，将返回 NULL 对象。只读。|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_方法_ > delete()|删除给定的名称。|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_方法_ > getRangeOrNullObject()|返回与名称相关联的 Range 对象。如果已命名项的类型不是 Range，将返回 NULL 对象。|1.4|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_方法_ > add(name: string, reference: Range or string, comment: string)|将新名称添加到给定范围的集合。|1.4|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_方法_ > addFormulaLocal(name: string, formula: string, comment: string)|使用用户的公式区域设置，将新名称添加到给定范围的集合。|1.4|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_方法_ > getCount()|获取集合中已命名项的数量。|1.4|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_方法_ > getItemOrNullObject(name: string)|按 NamedItem 对象的名称获取此对象。如果没有 NamedItem 对象，将返回 NULL 对象。|1.4|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_方法_ > getCount()|获取集合中的数据透视表的数量。|1.4|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_方法_ > getItemOrNullObject(name: string)|按 PivotTable 对象的名称获取此对象。如果没有 PivotTable 对象，将返回 NULL 对象。|1.4|
|[range](/javascript/api/excel/excel.range)|_方法_ > getIntersectionOrNullObject(anotherRange: Range or string)|获取表示指定区域的矩形交集的 range 对象。如果找不到任何交集，则此方法返回空对象。|1.4|
|[range](/javascript/api/excel/excel.range)|_方法_ > getUsedRangeOrNullObject(valuesOnly: bool)|返回指定 Range 对象的所用区域。如果区域内没有使用单元格，此函数将返回 NULL 对象。|1.4|
|[rangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|_方法_ > getCount()|获取集合中 RangeView 对象的数量。|1.4|
|[setting](/javascript/api/excel/excel.setting)|_属性_ > key|返回表示 setting 对象的 ID 的键。只读。|1.4|
|[setting](/javascript/api/excel/excel.setting)|_属性_ > value|表示为此设置存储的值。|1.4|
|[setting](/javascript/api/excel/excel.setting)|_方法_ > delete()|删除 Setting 对象。|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_属性_ > items|一组 setting 对象。只读。|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_方法_ > add(key: string, value: (any))|设置指定的 Setting 对象，或将其添加到工作簿中。|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_方法_ > getCount()|获取集合中的 Setting 对象的数量。|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_方法_ > getItem(key: string)|按键获取 Setting 项。|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_方法_ > getItemOrNullObject(key: string)|按键获取 Setting 项。如果没有 Setting 项，将返回 NULL 对象。|1.4|
|[settingsChangedEventArgs](/javascript/api/excel/excel.settingschangedeventargs)|_关系_ > settings|获取表示引发了 SettingsChanged 事件的绑定的 Setting 对象。|1.4|
|[tableCollection](/javascript/api/excel/excel.tablecollection)|_方法_ > getCount()]|获取集合中的表数量。|1.4|
|[tableCollection](/javascript/api/excel/excel.tablecollection)|_方法_ > getItemOrNullObject(key: number or string)|按名称或 ID 获取表。如果没有表，将返回 NULL 对象。|1.4|
|[tableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|_方法_ > getCount()|获取表中的列数。|1.4|
|[tableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|_方法_ > getItemOrNullObject(key: number or string)|按名称或 ID 获取 column 对象。如果没有 column 对象，将返回 NULL 对象。|1.4|
|[tableRowCollection](/javascript/api/excel/excel.tablerowcollection)|_方法_ > getCount()|获取表格中的行数。|1.4|
|[workbook](/javascript/api/excel/excel.workbook)|_关系_ > settings|表示一组与 workbook 相关联的 setting 对象。只读。|1.4|
|[worksheet](/javascript/api/excel/excel.worksheet)|_关系_ > names|一组范围限定到当前工作表的名称。只读。|1.4|
|[worksheet](/javascript/api/excel/excel.worksheet)|_方法_ > getUsedRangeOrNullObject(valuesOnly: bool)|使用的区域是包含分配了值或格式的任意单元格的最小区域。如果整个工作表为空，此函数将返回 NULL 对象。|1.4|
|[worksheetCollection](/javascript/api/excel/excel.worksheetcollection)|_方法_ > getCount(visibleOnly: bool)|获取集合中的工作表数量。|1.4|
|[worksheetCollection](/javascript/api/excel/excel.worksheetcollection)|_方法_ > getItemOrNullObject(key: string)|按 Worksheet 对象的名称或 ID 获取此对象。如果没有 Worksheet 对象，将返回 NULL 对象。|1.4|

## <a name="whats-new-in-excel-javascript-api-13"></a>Excel JavaScript API 1.3 的最近更新

下面介绍了要求集 1.3 中 Excel JavaScript API 的新增内容。

|对象| 最近更新| 说明|要求集|
|:----|:----|:----|:----|
|[binding](/javascript/api/excel/excel.binding)|_方法_ > delete()|删除 binding 对象。|1.3|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_方法_ > add(range: Range or string, bindingType: string, id: string)|将新的 binding 对象添加到特定区域。|1.3|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_方法_ > addFromNamedItem(name: string, bindingType: string, id: string)|根据工作簿中的命名项添加新的 binding 对象。|1.3|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_方法_ > addFromSelection(bindingType: string, id: string)|根据当前选择的内容添加新的 binding 对象。|1.3|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_方法_ > getItemOrNull(id: string)|按 ID 获取 binding 对象。如果 binding 对象不存在，则返回对象的 isNull 属性为 true。|1.3|
|[chartCollection](/javascript/api/excel/excel.chartcollection)|_方法_ > getItemOrNull(name: string)|使用图表名称获取图表。如果存在多个名称相同的图表，将返回第一个图表。|1.3|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_方法_ > getItemOrNull(name: string)|按 nameditem 对象的名称获取此对象。如果 nameditem 对象不存在，则返回对象的 isNull 属性为 true。|1.3|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_属性_ > name|PivotTable 对象的名称。|1.3|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_关系_ > worksheet|包含当前 PivotTable 对象的工作表。只读。|1.3|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_方法_ > refresh()|刷新 PivotTable 对象。|1.3|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_属性_ > items|一组 PivotTable 对象。只读。|1.3|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_方法_ > getItem(name: string)|按名称获取 PivotTable 对象。|1.3|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_方法_ > getItemOrNull(name: string)|按名称获取 PivotTable 对象。如果 PivotTable 对象不存在，则返回对象的 isNull 属性为 true。|1.3|
|[range](/javascript/api/excel/excel.range)|_方法_ > getIntersectionOrNull(anotherRange: Range or string)|获取表示指定区域的矩形交集的 range 对象。如果找不到任何交集，则此方法返回空对象。|1.3|
|[range](/javascript/api/excel/excel.range)|_方法_ > getVisibleView()|表示当前 range 对象的可见行。|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_属性_ > cellAddresses|表示 RangeView 的单元格地址。只读。|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_属性_ > columnCount|返回可见列数。只读。|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_属性_ > formulas|表示采用 A1 表示法的公式。|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_属性_ > formulasLocal|使用用户语言和数字格式区域设置表示采用 A1 表示法的公式。例如，用英语表示的公式 "=SUM(A1, introduced in 1.5)" 在德语中将变为 "=SUMME(A1; 1,5)"。|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_属性_ > formulasR1C1|表示采用 R1C1 表示法的公式。|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_属性_ > index|返回表示 RangeView 的索引的值。只读。|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_属性_ > numberFormat|表示 Excel 中指定单元格的数字格式代码。|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_属性_ > rowCount|返回可见行数。只读。|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_属性_ > text|指定区域的文本值。文本值与单元格宽度无关。在 Excel UI 中替代 # 符号不会影响 API 返回的文本值。只读。|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_属性_ > valueTypes|表示每个单元格的数据类型。只读。可能的值是：Unknown、Empty、String、Integer、Double、Boolean、Error。|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_属性_ > values|表示指定的 RangeView 的原始值。返回的数据可能是字符串、数字，也可能是布尔值。包含错误的单元格将返回错误字符串。|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_关系_ > rows|表示一组与 range 相关联的 RangeView。只读。|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_方法_ > getRange()|获取与当前 RangeView 相关联的父 range。|1.3|
|[rangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|_属性_ > items|一组 rangeView 对象。只读。|1.3|
|[rangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|_方法_ > getItemAt(index: number)|按索引获取 RangeView 行。从零开始编制索引。|1.3|
|[setting](/javascript/api/excel/excel.setting)|_属性_ > key|返回表示 setting 对象的 ID 的键。只读。|1.3|
|[setting](/javascript/api/excel/excel.setting)|_方法_ > delete()|删除 Setting 对象。|1.3|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_属性_ > items|一组 setting 对象。只读。|1.3|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_方法_ > getItem(key: string)|按键获取 Setting 项。|1.3|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_方法_ > getItemOrNull(key: string)|按键获取 setting 项。如果 setting 对象不存在，则返回对象的 isNull 属性为 true。|1.3|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_方法_ > set(key: string, value: string)|设置指定的 Setting 对象，或将其添加到工作簿中。|1.3|
|[settingsChangedEventArgs](/javascript/api/excel/excel.settingschangedeventargs)|_关系_ > settingCollection|获取表示引发了 SettingsChanged 事件的 binding 的 setting 对象。|1.3|
|[table](/javascript/api/excel/excel.table)|_属性_ > highlightFirstColumn|指明第一列是否包含特殊格式。|1.3|
|[table](/javascript/api/excel/excel.table)|_属性_ > highlightLastColumn|指明最后一列是否包含特殊格式。|1.3|
|[table](/javascript/api/excel/excel.table)|_属性_ > showBandedColumns|指明列是否采用镶边格式来以不同的方式突出显示奇数列与偶数列，让表更易于阅读。|1.3|
|[table](/javascript/api/excel/excel.table)|_属性_ > showBandedRows|指明行是否采用镶边格式来以不同的方式突出显示奇数行与偶数行，让表更易于阅读。|1.3|
|[table](/javascript/api/excel/excel.table)|_属性_ > showFilterButton|指明是否在每个列标题的顶部显示筛选器按钮。仅当 table 中包含标题行时，才允许设定此设置。|1.3|
|[tableCollection](/javascript/api/excel/excel.tablecollection)|_方法_ > getItemOrNull(key: number or string)|按名称或 ID 获取 table 对象。如果 table 对象不存在，则返回对象的 isNull 属性为 true。|1.3|
|[tableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|_方法_ > getItemOrNull(key: number or string)|按名称或 ID 获取 column 对象。如果 column 对象不存在，则返回对象的 isNull 属性为 true。|1.3|
|[workbook](/javascript/api/excel/excel.workbook)|_关系_ > pivotTables|表示一组与 workbook 相关联的 PivotTable 对象。只读。|1.3|
|[workbook](/javascript/api/excel/excel.workbook)|_关系_ > settings|表示一组与 workbook 相关联的 setting 对象。只读。|1.3|
|[worksheet](/javascript/api/excel/excel.worksheet)|_关系_ > pivotTables|一组属于 worksheet 的 PivotTable 对象。只读。|1.3|

## <a name="whats-new-in-excel-javascript-api-12"></a>Excel JavaScript API 1.2 的最近更新

下面介绍了要求集 1.2 中 Excel JavaScript API 的新增内容。

|对象| 最近更新| 说明|要求集|
|:----|:----|:----|:----|
|[chart](/javascript/api/excel/excel.chart)|_属性_ > id|根据其在集合中的位置获取图表。只读。|1.2|
|[chart](/javascript/api/excel/excel.chart)|_关系_ > worksheet|包含当前 chart 的 worksheet 对象。只读。|1.2|
|[chart](/javascript/api/excel/excel.chart)|_方法_ > getImage(height: number, width: number, fittingMode: string)|通过缩放图表以适应指定的尺寸，将图表呈现为 base64 编码的图像。|1.2|
|[filter](/javascript/api/excel/excel.filter)|_关系_ > criteria|给定列上当前应用的筛选器。只读。|1.2|
|[filter](/javascript/api/excel/excel.filter)|_方法_ > apply(criteria: FilterCriteria)|在给定列中应用给定的筛选条件。|1.2|
|[filter](/javascript/api/excel/excel.filter)|_方法_ > applyBottomItemsFilter(count: number)|将“Bottom Item”筛选器应用于列，以获取给定数量的元素。|1.2|
|[filter](/javascript/api/excel/excel.filter)|_方法_ > applyBottomPercentFilter(percent: number)]|将“Bottom Percent”筛选器应用于列，以获取给定比例的元素。|1.2|
|[filter](/javascript/api/excel/excel.filter)|_方法_ > applyCellColorFilter(color: string)|将“Cell Color”筛选器应用于列，以获取给定颜色。|1.2|
|[filter](/javascript/api/excel/excel.filter)|_方法_ > applyCustomFilter(criteria1: string, criteria2: string, oper: string)|将“Icon”筛选器应用于列，以获取给定的条件字符串。|1.2|
|[filter](/javascript/api/excel/excel.filter)|_方法_ > applyDynamicFilter(criteria: string)|将“Dynamic”筛选器应用于列。|1.2|
|[filter](/javascript/api/excel/excel.filter)|_方法_ > applyFontColorFilter(color: string)|将“Font Color”筛选器应用于列，以获取给定颜色。|1.2|
|[filter](/javascript/api/excel/excel.filter)|_方法_ > applyIconFilter(icon: Icon)|将“Icon”筛选器应用于列，以获取给定图标。|1.2|
|[filter](/javascript/api/excel/excel.filter)|_方法_ > applyTopItemsFilter(count: number)|将“Top Item”筛选器应用于列，以获取给定数量的元素。|1.2|
|[filter](/javascript/api/excel/excel.filter)|_方法_ > applyTopPercentFilter(percent: number)|将“Top Percent”筛选器应用于列，以获取给定比例的元素。|1.2|
|[filter](/javascript/api/excel/excel.filter)|_方法_ > applyValuesFilter(values: ())|将“Values”筛选器应用于列，获取给定值。|1.2|
|[filter](/javascript/api/excel/excel.filter)|_方法_ > clear()|清除给定列上的 filter。|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_属性_ > color|用于筛选单元格的 HTML 颜色字符串。与“cellColor”和“fontColor”筛选一起使用。|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_属性_ > criterion1|第一个条件用于筛选数据。在“自定义”筛选中用作运算符。|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_属性_ > criterion2|第二个条件用于筛选数据。在“自定义”筛选中仅用作运算符。|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_属性_ > dynamicCriteria|Excel.DynamicFilterCriteria 集中的动态条件将应用于此列。与“动态”筛选一起使用。可能的值是：Unknown、AboveAverage、AllDatesInPeriodApril、AllDatesInPeriodAugust、AllDatesInPeriodDecember、AllDatesInPeriodFebruray、AllDatesInPeriodJanuary、AllDatesInPeriodJuly、AllDatesInPeriodJune、AllDatesInPeriodMarch、AllDatesInPeriodMay、AllDatesInPeriodNovember、AllDatesInPeriodOctober、AllDatesInPeriodQuarter1、AllDatesInPeriodQuarter2、AllDatesInPeriodQuarter3、AllDatesInPeriodQuarter4、AllDatesInPeriodSeptember、BelowAverage、LastMonth、LastQuarter、LastWeek、LastYear、NextMonth、NextQuarter、NextWeek、NextYear、ThisMonth、ThisQuarter、ThisWeek、ThisYear、Today、Tomorrow、YearToDate、Yesterday。|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_属性_ > filterOn|filter 使用的属性，用于确定值是否应一直可见。可取值为：BottomItems、BottomPercent、CellColor、Dynamic、FontColor、Values、TopItems、TopPercent、Icon、Custom。|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_属性_ > operator|使用“自定义”筛选器时，用于组合条件 1 和 2 的运算符。可取值为：And、Or。|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_属性_ > values|一组用于“values”筛选器的值。|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_关系_ > icon|用于筛选单元格的图标。与“图标”筛选一起使用。|1.2|
|[filterDatetime](/javascript/api/excel/excel.filterdatetime)|_属性_ > date|用于筛选数据的采用 ISO8601 格式的日期。|1.2|
|[filterDatetime](/javascript/api/excel/excel.filterdatetime)|_属性_ > specificity|用于保留数据的日期的具体程度。例如，如果日期是 2005-04-02 并且将特殊性设置为“月”，则筛选操作将保留包含 2009 年 4 月日期的所有行。可能的值是：Year、Monday、Day、Hour、Minute、Second。|1.2|
|[formatProtection](/javascript/api/excel/excel.formatprotection)|_属性_ > formulaHidden|表示 Excel 是否隐藏区域中的单元格公式。指示整个区域不具有统一公式隐藏设置的空值。|1.2|
|[formatProtection](/javascript/api/excel/excel.formatprotection)|_属性_ > locked|指示 Excel 是否锁定对象中的单元格。指示整个区域不具有统一锁定设置的空值。|1.2|
|[icon](/javascript/api/excel/excel.icon)|_属性_ > index|表示 icon 在给定集中的索引。|1.2|
|[icon](/javascript/api/excel/excel.icon)|_属性_ > set|表示图标所属的集合。可能的值是：Invalid、ThreeArrows、ThreeArrowsGray、ThreeFlags、ThreeTrafficLights1、ThreeTrafficLights2、ThreeSigns、ThreeSymbols、ThreeSymbols2、FourArrows、FourArrowsGray、FourRedToBlack、FourRating、FourTrafficLights、FiveArrows、FiveArrowsGray、FiveRating、FiveQuarters、ThreeStars、ThreeTriangles、FiveBoxes。|1.2|
|[range](/javascript/api/excel/excel.range)|_属性_ > columnHidden|表示当前 range 的所有列均已隐藏。|1.2|
|[range](/javascript/api/excel/excel.range)|_属性_ > formulasR1C1|表示采用 R1C1 表示法的公式。|1.2|
|[range](/javascript/api/excel/excel.range)|_属性_ > hidden|表示当前区域中的所有单元格是否隐藏。只读。|1.2|
|[range](/javascript/api/excel/excel.range)|_属性_ > rowHidden|表示当前 range 的所有行均已隐藏。|1.2|
|[range](/javascript/api/excel/excel.range)|_关系_ > sort|表示当前 range 的区域排序。只读。|1.2|
|[range](/javascript/api/excel/excel.range)|_方法_ > merge(across: bool)|在工作表中，将 range 单元格合并到一个区域中。|1.2|
|[range](/javascript/api/excel/excel.range)|_方法_ > unmerge()|将范围单元格取消合并为各个单元格。|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_属性_ > columnWidth|获取或设置区域内的所有列的宽度。如果列宽不统一，则返回 NULL。|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_属性_ > rowHeight|获取或设置区域中所有行的高度。如果行高不统一，则返回 NULL。|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_关系_ > protection|返回某一区域的格式保护对象。只读。|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_方法_ > autofitColumns()|根据列中的当前数据更改当前范围的列宽，以达到最佳宽度。|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_方法_ > autofitRows()|根据列中的当前数据，更改当前范围的行高以达到最佳高度。|1.2|
|[rangeReference](/javascript/api/excel/excel.rangereference)|_属性_ > address|表示当前 range 对象的可见行。|1.2|
|[rangeSort](/javascript/api/excel/excel.rangesort)|_方法_ > apply(fields: SortField, matchCase: bool, hasHeaders: bool, orientation: string, method: string)|执行排序操作。|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_属性_ > ascending|表示是否执行升序排序。|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_属性_ > color|表示按字体或单元格颜色进行排序时，条件的目标颜色。|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_属性_ > dataOption|表示此字段的其他排序选项。可能的值是：Normal、TextAsNumber。|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_属性_ > key|表示条件所在的列（或行，具体取决于排序方向）。表示与第一列（或行）的偏移量。|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_属性_ > sortOn|表示此条件的排序类型。可能的值是：Value、CellColor、FontColor、Icon。|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_关系_ > icon|表示对单元格图标进行排序时，条件的目标图标。|1.2|
|[table](/javascript/api/excel/excel.table)|_关系_ > sort|表示表的排序。只读。|1.2|
|[table](/javascript/api/excel/excel.table)|_关系_ > worksheet|包含当前表的工作表。只读。|1.2|
|[table](/javascript/api/excel/excel.table)|_方法_ > clearFilters()|清除当前表上应用的所有筛选器。|1.2|
|[table](/javascript/api/excel/excel.table)|_方法_ > convertToRange()|将表转换为普通单元格区域。保留所有数据。|1.2|
|[table](/javascript/api/excel/excel.table)|_方法_ > reapplyFilters()|重新应用当前在 table 上应用的所有 filter。|1.2|
|[tableColumn](/javascript/api/excel/excel.tablecolumn)|_关系_ > filter|检索应用于列的筛选器。只读。|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_属性_ > matchCase|表示最后一次对表进行排序时大小写是否有影响。只读。|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_属性_ > method|表示最后一次对表排序所使用的中文字符排序方法。只读。可能的值是：PinYin、StrokeCount。|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_关系_ > fields|表示最后一次对表排序所使用的当前条件。只读。|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_方法_ > apply(fields: SortField, matchCase: bool, method: string)|执行排序操作。|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_方法_ > clear()|清除表上的当前排序。尽管这不能修改表的排序，但它会清除标题按钮的状态。|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_方法_ > reapply()|对 table 重新应用当前的排序参数。|1.2|
|[workbook](/javascript/api/excel/excel.workbook)|_关系_ > functions|表示包含此工作簿的 Excel 应用程序实例。只读。|1.2|
|[worksheet](/javascript/api/excel/excel.worksheet)|_关系_ > protection|返回表工作表的工作表保护对象。只读。|1.2|
|[worksheetProtection](/javascript/api/excel/excel.worksheetprotection)|_属性_ > protected|指明 worksheet 是否受保护。只读。只读。|1.2|
|[worksheetProtection](/javascript/api/excel/excel.worksheetprotection)|_关系_ > options|工作表保护选项。只读。|1.2|
|[worksheetProtection](/javascript/api/excel/excel.worksheetprotection)|_方法_ > protect(options: WorksheetProtectionOptions)|保护 worksheet。如果 worksheet 处于受保护状态，则无法执行此方法。|1.2|
|[worksheetProtection](/javascript/api/excel/excel.worksheetprotection)|_方法_ > unprotect()|解除对 worksheet 的保护。|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_属性_ > allowAutoFilter|表示允许使用自动筛选功能的工作表保护选项。|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_属性_ > allowDeleteColumns|表示允许删除列的工作表保护选项。|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_属性_ > allowDeleteRows|表示允许删除行的工作表保护选项。|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_属性_ > allowFormatCells|表示允许格式化单元格的工作表保护选项。|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_属性_ > allowFormatColumns|表示允许格式化列的工作表保护选项。|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_属性_ > allowFormatRows|表示允许格式化行的工作表保护选项。|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_属性_ > allowInsertColumns|表示允许插入列的工作表保护选项。|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_属性_ > allowInsertHyperlinks|表示允许插入超链接的工作表保护选项。|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_属性_ > allowInsertRows|表示允许插入行的工作表保护选项。|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_属性_ > allowPivotTables|表示允许使用数据透视表功能的工作表保护选项。|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_属性_ > allowSort|表示允许使用排序功能的工作表保护选项。|1.2|

## <a name="excel-javascript-api-11"></a>Excel JavaScript API 1.1

Excel JavaScript API 1.1 是首版 API。有关 API 的详细信息，请参阅 [Excel JavaScript API](/javascript/api/excel) 参考主题。

## <a name="see-also"></a>另请参阅

- [Office 版本和要求集](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [指定 Office 主机和 API 要求](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office 外接程序 XML 清单](/office/dev/add-ins/develop/add-in-manifests)
