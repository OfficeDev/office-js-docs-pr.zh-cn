---
title: Excel JavaScript 预览 API
description: 有关即将推出的 Excel JavaScript Api 的详细信息
ms.date: 10/22/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: dc0a2a3b23fbf4ccffb5de3b0689b0de0ed08b75
ms.sourcegitcommit: 5ba325cc88183a3f230cd89d615fd49c695addcf
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/24/2019
ms.locfileid: "37682541"
---
# <a name="excel-javascript-preview-apis"></a>Excel JavaScript 预览 API

新的 Excel JavaScript API 首先在“预览版”中引入，在进行充分测试并获得用户反馈后，它将成为编号的特定要求集的一部分。

第一个表提供了 API 的简明摘要，而后续表提供了详细列表。

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

| 功能区域 | 说明 | 相关对象 |
|:--- |:--- |:--- |
| [注释提到](../../excel/excel-add-ins-comments.md#mentions-preview) | 在评论中提及其他人以发送通知。 | [Comment](/javascript/api/excel/excel.comment)、 [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) |
| [插入工作簿](../../excel/excel-add-ins-workbooks.md#insert-a-copy-of-an-existing-workbook-into-the-current-one-preview) | 将一个工作簿插入另一个工作簿。  | [Workbook](/javascript/api/excel/excel.worksheetcollection) |
| 工作簿[保存](../../excel/excel-add-ins-workbooks.md#save-the-workbook-preview)和[关闭](../../excel/excel-add-ins-workbooks.md#close-the-workbook-preview) | 保存和关闭工作簿。  | [Workbook](/javascript/api/excel/excel.workbook) |

## <a name="api-list"></a>API 列表

下表列出了当前预览中的 Excel JavaScript Api。 若要查看所有 Excel JavaScript Api （包括预览 Api 和之前发布的 Api）的完整列表，请参阅[所有 Excel Javascript api](/javascript/api/excel?view=excel-js-preview)。

| Class | 域 | 说明 |
|:---|:---|:---|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[getDimensionValues （维： ChartSeriesDimension）](/javascript/api/excel/excel.chartseries#getdimensionvalues-dimension-)|从图表系列的单个维中获取值。 这些值可以是类别值，也可以是数据值，具体取决于指定的维度和为图表系列映射数据的方式。|
|[Comment](/javascript/api/excel/excel.comment)|[提及](/javascript/api/excel/excel.comment#mentions)|获取注释中提到的实体（如人员）。|
||[richContent](/javascript/api/excel/excel.comment#richcontent)|获取丰富的注释内容（例如，注释中的提及）。 此字符串不应显示给最终用户。 您的外接程序应仅使用此信息分析丰富的注释内容。|
||[经过](/javascript/api/excel/excel.comment#resolved)|获取或设置批注线程的状态。 值为 "true" 表示注释线程处于 "已解决" 状态。|
||[updateMentions （contentWithMentions： CommentRichContent）](/javascript/api/excel/excel.comment#updatementions-contentwithmentions-)|使用特殊格式的字符串和提及列表更新注释内容。|
|[CommentMention](/javascript/api/excel/excel.commentmention)|[email](/javascript/api/excel/excel.commentmention#email)|获取或设置注释中提到的实体的电子邮件地址。|
||[id](/javascript/api/excel/excel.commentmention#id)|获取或设置实体的 id。 这与中`CommentRichContent.richContent`的 id 信息对齐。|
||[name](/javascript/api/excel/excel.commentmention#name)|获取或设置注释中提到的实体的名称。|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[提及](/javascript/api/excel/excel.commentreply#mentions)|获取注释中提到的实体（如人员）。|
||[经过](/javascript/api/excel/excel.commentreply#resolved)|获取或设置批注答复状态。 值为 "true" 表示批注答复处于 "已解决" 状态。|
||[richContent](/javascript/api/excel/excel.commentreply#richcontent)|获取丰富的注释内容（例如，注释中的提及）。 此字符串不应显示给最终用户。 您的外接程序应仅使用此信息分析丰富的注释内容。|
||[updateMentions （contentWithMentions： CommentRichContent）](/javascript/api/excel/excel.commentreply#updatementions-contentwithmentions-)|使用特殊格式的字符串和提及列表更新注释内容。|
|[CommentRichContent](/javascript/api/excel/excel.commentrichcontent)|[提及](/javascript/api/excel/excel.commentrichcontent#mentions)|包含注释中提到的所有实体（例如，人员）的数组。|
||[richContent](/javascript/api/excel/excel.commentrichcontent#richcontent)||
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|根据数据层次结构以及各自层次结构的行和列项，获取数据透视表中的唯一单元格。 返回的单元格是给定行和列的交集，其中包含来自给定层次结构的数据。 此方法与在特定单元格上调用 getPivotItems 和 getDataHierarchy 相反。|
|[Range](/javascript/api/excel/excel.range)|[getSpillParent()](/javascript/api/excel/excel.range#getspillparent--)|获取 Range 对象，它包含要将某个单元格溢出到的定位单元格。 如果应用于具有多个单元格的区域，则会失败。 只读。|
||[getSpillParentOrNullObject()](/javascript/api/excel/excel.range#getspillparentornullobject--)|获取 Range 对象，它包含要将某个单元格溢出到的定位单元格。 只读。|
||[getSpillingToRange()](/javascript/api/excel/excel.range#getspillingtorange--)|获取 Range 对象，它在调用定位单元格时包含溢出区域。 如果应用于具有多个单元格的区域，则会失败。 只读。|
||[getSpillingToRangeOrNullObject()](/javascript/api/excel/excel.range#getspillingtorangeornullobject--)|获取 Range 对象，它在调用定位单元格时包含溢出区域。 只读。|
||[hasSpill](/javascript/api/excel/excel.range#hasspill)|表示所有单元格是否都具有溢出边框。|
||[savedAsArray](/javascript/api/excel/excel.range#savedasarray)|表示是否将所有单元格都保存为数组公式。|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[adjustIndent （金额：数字）](/javascript/api/excel/excel.rangeformat#adjustindent-amount-)|调整范围格式的缩进量。 缩进值的范围为0到250。|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|从 XML 字符串创建可缩放的矢量图形 (SVG) 并将其添加到工作表。 返回表示新图片的 Shape 对象。|
|[Slicer](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameinformula)|表示公式中使用切片器名称。|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|将表格更改为使用默认表格样式。|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|在特定表格上应用筛选器时发生。|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|在工作簿或工作表中的任何表格上应用筛选器时发生。|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|表示应用了筛选器的表的 id。|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#type)|表示事件的类型。 有关详细信息，请参阅 Excel.EventType。|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|表示包含表格的工作表的 ID。|
|[Workbook](/javascript/api/excel/excel.workbook)|[close(closeBehavior?: Excel.CloseBehavior)](/javascript/api/excel/excel.workbook#close-closebehavior-)|关闭当前工作簿。|
||[save(saveBehavior?: Excel.SaveBehavior)](/javascript/api/excel/excel.workbook#save-savebehavior-)|保存当前工作簿。|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|如果工作簿使用 1904 日期系统，则为 True。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|在特定工作表上应用筛选器时发生。|
||[onRowHiddenChanged](/javascript/api/excel/excel.worksheet#onrowhiddenchanged)|在特定工作表上的一个或多个行的隐藏状态更改时发生。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|将工作簿的指定工作表插入当前工作簿。|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|在工作簿中应用任何工作表的筛选器时发生。|
||[onRowHiddenChanged](/javascript/api/excel/excel.worksheetcollection#onrowhiddenchanged)|在特定工作表上的一个或多个行的隐藏状态更改时发生。|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|表示事件的类型。 有关详细信息，请参阅 Excel.EventType。|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|表示在其中应用筛选器的工作表的 ID。|
|[WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs)|[address](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#address)|获取区域地址，该地址表示特定工作表上的更改区域。|
||[changeType](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#changetype)|获取表示事件触发方式的更改类型。 有关详细信息，请参阅 `Excel.RowHiddenChangeType`。|
||[source](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#source)|获取事件源。 有关详细信息，请参阅 Excel.EventSource。|
||[type](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#type)|获取事件的类型。 有关详细信息，请参阅 Excel.EventType。|
||[worksheetId](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#worksheetid)|获取其中的数据发生更改的工作表的 ID。|

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 参考文档](/javascript/api/excel?view=excel-js-preview)
- [Excel JavaScript API 要求集](./excel-api-requirement-sets.md)
