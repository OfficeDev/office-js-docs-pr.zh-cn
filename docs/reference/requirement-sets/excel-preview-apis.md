---
title: Excel JavaScript 预览 API
description: 有关即将推出的 Excel JavaScript API 的详细信息。
ms.date: 02/24/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 0663b6330c402f64e7ed7e8f598a52848bbe1319
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505533"
---
# <a name="excel-javascript-preview-apis"></a>Excel JavaScript 预览 API

新的 Excel JavaScript API 首先在“预览版”中引入，在进行充分测试并获得用户反馈后，它将成为编号的特定要求集的一部分。

第一个表提供了 API 的简明摘要，而后续表提供了详细列表。

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

| 功能区域 | 说明 | 相关对象 |
|:--- |:--- |:--- |
| 公式更改事件 | 跟踪对公式的更改，包括导致更改的事件的源和类型。 | [Worksheet.onFormulaChanged](/javascript/api/excel/excel.worksheet#onFormulaChanged)|
| 链接的数据类型 | 添加对从外部源连接到 Excel 的数据类型的支持。 | [LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|
| 命名工作表视图 | 以编程方式控制每用户工作表视图。 | [NamedSheetView](/javascript/api/excel/excel.namedsheetview) |
| 任务 | 将注释转换为分配给用户的任务。 | [任务](/javascript/api/excel/excel.task) |

## <a name="api-list"></a>API 列表

下表列出了当前预览版中的 Excel JavaScript API。 有关所有 Excel JavaScript API 的完整列表， (预览 API 和以前发布的 API) ，请参阅所有[Excel JavaScript API。](/javascript/api/excel?view=excel-js-preview&preserve-view=true)

| 类 | 域 | 说明 |
|:---|:---|:---|
|[Comment](/javascript/api/excel/excel.comment)|[assignTask (电子邮件：字符串) ](/javascript/api/excel/excel.comment#assigntask-email-)|将附加到注释的任务作为唯一的代理人分配给给定用户。|
||[getTask () ](/javascript/api/excel/excel.comment#gettask--)|获取与此注释关联的任务。|
||[getTaskOrNullObject () ](/javascript/api/excel/excel.comment#gettaskornullobject--)|获取与此注释关联的任务。|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[assignTask (电子邮件：字符串) ](/javascript/api/excel/excel.commentreply#assigntask-email-)|将附加到注释的任务作为唯一的代理人分配给给定用户。|
||[getTask () ](/javascript/api/excel/excel.commentreply#gettask--)|获取与此注释关联的任务。|
||[getTaskOrNullObject () ](/javascript/api/excel/excel.commentreply#gettaskornullobject--)|获取与此注释关联的任务。|
|[FormulaChangedEventDetail](/javascript/api/excel/excel.formulachangedeventdetail)|[cellAddress](/javascript/api/excel/excel.formulachangedeventdetail#celladdress)|包含已更改公式的单元格的地址。|
||[previousFormula](/javascript/api/excel/excel.formulachangedeventdetail#previousformula)|表示更改前一个公式。|
|[LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|[dataProvider](/javascript/api/excel/excel.linkeddatatype#dataprovider)|链接数据提供程序的数据提供程序数据类型。|
||[lastRefreshed](/javascript/api/excel/excel.linkeddatatype#lastrefreshed)|自上次刷新链接的工作簿时打开数据类型时区日期和时间。|
||[名称](/javascript/api/excel/excel.linkeddatatype#name)|链接对象的名称数据类型。|
||[periodicRefreshInterval](/javascript/api/excel/excel.linkeddatatype#periodicrefreshinterval)|链接对象刷新的频率（以秒数据类型设置为 `refreshMode` "Periodic"时刷新。|
||[refreshMode](/javascript/api/excel/excel.linkeddatatype#refreshmode)|检索链接数据数据类型的机制。|
||[serviceId](/javascript/api/excel/excel.linkeddatatype#serviceid)|链接对象的唯一数据类型。|
||[supportedRefreshModes](/javascript/api/excel/excel.linkeddatatype#supportedrefreshmodes)|返回一个数组，该数组支持链接对象支持的所有数据类型。|
||[requestRefresh () ](/javascript/api/excel/excel.linkeddatatype#requestrefresh--)|请求刷新链接数据类型。|
||[requestSetRefreshMode (refreshMode： Excel.LinkedDataTypeRefreshMode) ](/javascript/api/excel/excel.linkeddatatype#requestsetrefreshmode-refreshmode-)|请求更改此链接的刷新数据类型。|
|[LinkedDataTypeAddedEventArgs](/javascript/api/excel/excel.linkeddatatypeaddedeventargs)|[serviceId](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#serviceid)|新链接对象的唯一数据类型。|
||[source](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#source)|获取事件源。|
||[type](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#type)|获取事件的类型。|
|[LinkedDataTypeCollection](/javascript/api/excel/excel.linkeddatatypecollection)|[getCount()](/javascript/api/excel/excel.linkeddatatypecollection#getcount--)|获取集合中链接的数据类型的数量。|
||[getItem (键：数字) ](/javascript/api/excel/excel.linkeddatatypecollection#getitem-key-)|按服务 id 数据类型链接的标识符。|
||[getItemAt(index: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitemat-index-)|按集合数据类型索引获取链接的索引。|
||[getItemOrNullObject (键：数字) ](/javascript/api/excel/excel.linkeddatatypecollection#getitemornullobject-key-)|按 ID 获取数据类型链接对象。|
||[items](/javascript/api/excel/excel.linkeddatatypecollection#items)|获取此集合中已加载的子项。|
||[requestRefreshAll () ](/javascript/api/excel/excel.linkeddatatypecollection#requestrefreshall--)|请求刷新集合中所有链接的数据类型。|
|[NamedSheetView](/javascript/api/excel/excel.namedsheetview)|[activate()](/javascript/api/excel/excel.namedsheetview#activate--)|激活此工作表视图。|
||[delete()](/javascript/api/excel/excel.namedsheetview#delete--)|从工作表中删除工作表视图。|
||[重复 (名称？：字符串) ](/javascript/api/excel/excel.namedsheetview#duplicate-name-)|创建此工作表视图的副本。|
||[名称](/javascript/api/excel/excel.namedsheetview#name)|获取或设置工作表视图的名称。|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[add(name: string)](/javascript/api/excel/excel.namedsheetviewcollection#add-name-)|创建具有给定名称的新工作表视图。|
||[enterTemporary () ](/javascript/api/excel/excel.namedsheetviewcollection#entertemporary--)|创建并激活新的临时工作表视图。|
||[exit () ](/javascript/api/excel/excel.namedsheetviewcollection#exit--)|退出当前活动的工作表视图。|
||[getActive () ](/javascript/api/excel/excel.namedsheetviewcollection#getactive--)|获取工作表当前的活动工作表视图。|
||[getCount()](/javascript/api/excel/excel.namedsheetviewcollection#getcount--)|获取此工作表中的工作表视图数。|
||[getItem(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getitem-key-)|使用工作表视图的名称获取工作表视图。|
||[getItemAt(index: number)](/javascript/api/excel/excel.namedsheetviewcollection#getitemat-index-)|按工作表视图在集合中的索引获取工作表视图。|
||[items](/javascript/api/excel/excel.namedsheetviewcollection#items)|获取此集合中已加载的子项。|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[altTextDescription](/javascript/api/excel/excel.pivotlayout#alttextdescription)|数据透视表的替换文字说明。|
||[altTextTitle](/javascript/api/excel/excel.pivotlayout#alttexttitle)|数据透视表的替换文字标题。|
||[displayBlankLineAfterEachItem (显示： boolean) ](/javascript/api/excel/excel.pivotlayout#displayblanklineaftereachitem-display-)|设置是否在每一项后显示一个空行。|
||[emptyCellText](/javascript/api/excel/excel.pivotlayout#emptycelltext)|自动填充到数据透视表中任何空单元格的文本（如果 `fillEmptyCells == true` ）。|
||[fillEmptyCells](/javascript/api/excel/excel.pivotlayout#fillemptycells)|指定数据透视表中是否应当填充空单元格 `emptyCellText` 。|
||[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|根据数据层次结构以及各自层次结构的行和列项，获取数据透视表中的唯一单元格。|
||[pivotStyle](/javascript/api/excel/excel.pivotlayout#pivotstyle)|应用于数据透视表的样式。|
||[repeatAllItemLabels (repeatLabels： boolean) ](/javascript/api/excel/excel.pivotlayout#repeatallitemlabels-repeatlabels-)|设置数据透视表中所有字段的"重复所有项目标签"设置。|
||[setStyle (：string \| PivotTableStyle \| BuiltInPivotTableStyle) ](/javascript/api/excel/excel.pivotlayout#setstyle-style-)|设置应用于数据透视表的样式。|
||[showFieldHeaders](/javascript/api/excel/excel.pivotlayout#showfieldheaders)|指定数据透视表是否显示字段标题 (字段标题和筛选下拉) 。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[refreshOnOpen](/javascript/api/excel/excel.pivottable#refreshonopen)|指定工作簿打开时数据透视表是否刷新。|
|[区域](/javascript/api/excel/excel.range)|[getPrecedents () ](/javascript/api/excel/excel.range#getprecedents--)|返回一个对象，该对象代表包含同一工作表或多个工作表中单元格的所有引用 `WorkbookRangeAreas` 单元格的范围。|
|[RefreshModeChangedEventArgs](/javascript/api/excel/excel.refreshmodechangedeventargs)|[refreshMode](/javascript/api/excel/excel.refreshmodechangedeventargs#refreshmode)|链接的数据类型刷新模式。|
||[serviceId](/javascript/api/excel/excel.refreshmodechangedeventargs#serviceid)|其刷新模式已更改的对象的唯一 ID。|
||[source](/javascript/api/excel/excel.refreshmodechangedeventargs#source)|获取事件源。|
||[type](/javascript/api/excel/excel.refreshmodechangedeventargs#type)|获取事件的类型。|
|[RefreshRequestCompletedEventArgs](/javascript/api/excel/excel.refreshrequestcompletedeventargs)|[已刷新](/javascript/api/excel/excel.refreshrequestcompletedeventargs#refreshed)|指示刷新请求是否成功。|
||[serviceId](/javascript/api/excel/excel.refreshrequestcompletedeventargs#serviceid)|已完成刷新请求的对象的唯一 ID。|
||[source](/javascript/api/excel/excel.refreshrequestcompletedeventargs#source)|获取事件源。|
||[type](/javascript/api/excel/excel.refreshrequestcompletedeventargs#type)|获取事件的类型。|
||[警告](/javascript/api/excel/excel.refreshrequestcompletedeventargs#warnings)|包含从刷新请求生成的任何警告的数组。|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|从 XML 字符串创建可缩放的矢量图形 (SVG) 并将其添加到工作表。|
|[Slicer](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameinformula)|表示公式中使用切片器名称。|
||[slicerStyle](/javascript/api/excel/excel.slicer#slicerstyle)|应用于切片器的样式。|
||[setStyle (：string \| SlicerStyle \| BuiltInSlicerStyle) ](/javascript/api/excel/excel.slicer#setstyle-style-)|设置应用于切片器的样式。|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|将表格更改为使用默认表格样式。|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|在特定表格上应用筛选器时发生。|
||[tableStyle](/javascript/api/excel/excel.table#tablestyle)|应用于 Table 的样式。|
||[setStyle (样式：string \| TableStyle \| BuiltInTableStyle) ](/javascript/api/excel/excel.table#setstyle-style-)|设置应用于表格的样式。|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|在工作簿或工作表中的任何表格上应用筛选器时发生。|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|获取应用筛选器的表的 ID。|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|获取包含表格的工作表的 ID。|
|[任务](/javascript/api/excel/excel.task)|[addAssignee (电子邮件：字符串) ](/javascript/api/excel/excel.task#addassignee-email-)|向任务添加一个收件人。|
||[applyChanges (taskChanges： Excel.TaskChanges) ](/javascript/api/excel/excel.task#applychanges-taskchanges-)|对任务应用给定的更改。|
||[被分配者](/javascript/api/excel/excel.task#assignees)|获取任务分配给的用户。|
||[comment](/javascript/api/excel/excel.task#comment)|获取与任务关联的注释。|
||[dueDate](/javascript/api/excel/excel.task#duedate)|获取任务的截止日期。|
||[historyRecords](/javascript/api/excel/excel.task#historyrecords)|获取任务的历史记录。|
||[id](/javascript/api/excel/excel.task#id)|获取任务的 ID。|
||[percentComplete](/javascript/api/excel/excel.task#percentcomplete)|获取任务的完成百分比。|
||[priority](/javascript/api/excel/excel.task#priority)|获取任务的优先级。|
||[startDate](/javascript/api/excel/excel.task#startdate)|获取任务应开始的日期和时间。|
||[title](/javascript/api/excel/excel.task#title)|获取任务的标题。|
||[removeAllAssignees () ](/javascript/api/excel/excel.task#removeallassignees--)|从任务中删除所有被分配者。|
||[removeAssignee (电子邮件：字符串) ](/javascript/api/excel/excel.task#removeassignee-email-)|从任务中删除一个被委派者。|
||[setPercentComplete (percentComplete： number) ](/javascript/api/excel/excel.task#setpercentcomplete-percentcomplete-)|更改任务的完成情况。|
||[setPriority (优先级：数字) ](/javascript/api/excel/excel.task#setpriority-priority-)|更改任务的优先级。|
||[setStartDateAndDueDate (startDate： Date， dueDate： Date) ](/javascript/api/excel/excel.task#setstartdateandduedate-startdate--duedate-)|更改任务的开始日期和截止日期。|
||[setTitle (标题：字符串) ](/javascript/api/excel/excel.task#settitle-title-)|更改任务的标题。|
|[TaskChanges](/javascript/api/excel/excel.taskchanges)|[dueDate](/javascript/api/excel/excel.taskchanges#duedate)|设置任务的新截止日期（UTC 时区）。|
||[emailsToAssign](/javascript/api/excel/excel.taskchanges#emailstoassign)|设置要分配给任务的用户的电子邮件地址。|
||[emailsToUnassign](/javascript/api/excel/excel.taskchanges#emailstounassign)|将用户的电子邮件地址设置为取消分配任务。|
||[percentComplete](/javascript/api/excel/excel.taskchanges#percentcomplete)|设置任务的新完成百分比。|
||[priority](/javascript/api/excel/excel.taskchanges#priority)|设置任务的新优先级。|
||[removeAllPreviousAssignees](/javascript/api/excel/excel.taskchanges#removeallpreviousassignees)|设置更改是否应当从任务中删除所有以前的分配者。|
||[startDate](/javascript/api/excel/excel.taskchanges#startdate)|设置任务的新开始日期（UTC 时区）。|
||[title](/javascript/api/excel/excel.taskchanges#title)|设置任务的新标题。|
|[TaskCollection](/javascript/api/excel/excel.taskcollection)|[getCount()](/javascript/api/excel/excel.taskcollection#getcount--)|获取集合中的任务数。|
||[getItem(key: string)](/javascript/api/excel/excel.taskcollection#getitem-key-)|使用其 ID 获取任务。|
||[getItemAt(index: number)](/javascript/api/excel/excel.taskcollection#getitemat-index-)|按任务在集合中的索引获取任务。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.taskcollection#getitemornullobject-key-)|使用其 ID 获取任务。|
||[items](/javascript/api/excel/excel.taskcollection#items)|获取此集合中已加载的子项。|
|[TaskHistoryRecord](/javascript/api/excel/excel.taskhistoryrecord)|[anchorId](/javascript/api/excel/excel.taskhistoryrecord#anchorid)|表示任务定位到的对象的 ID (例如，附加到批注或批注) 。|
||[被分配者](/javascript/api/excel/excel.taskhistoryrecord#assignee)|表示分配给"分配"历史记录记录类型任务的用户，或者从"未分配"历史记录记录类型的任务取消分配的用户。|
||[attributionUser](/javascript/api/excel/excel.taskhistoryrecord#attributionuser)|表示创建或更改任务的用户。|
||[dueDate](/javascript/api/excel/excel.taskhistoryrecord#duedate)|表示任务的截止日期。|
||[historyRecordCreatedDate](/javascript/api/excel/excel.taskhistoryrecord#historyrecordcreateddate)|表示任务历史记录的创建日期。|
||[id](/javascript/api/excel/excel.taskhistoryrecord#id)|历史记录的 ID。|
||[percentComplete](/javascript/api/excel/excel.taskhistoryrecord#percentcomplete)|表示任务的完成百分比。|
||[priority](/javascript/api/excel/excel.taskhistoryrecord#priority)|表示任务的优先级。|
||[startDate](/javascript/api/excel/excel.taskhistoryrecord#startdate)|表示任务的开始日期。|
||[title](/javascript/api/excel/excel.taskhistoryrecord#title)|表示任务的标题。|
||[type](/javascript/api/excel/excel.taskhistoryrecord#type)|表示任务历史记录记录的类型。|
||[undoHistoryId](/javascript/api/excel/excel.taskhistoryrecord#undohistoryid)|代表TaskHistoryRecord.id"Undo"历史记录记录类型撤消的一个属性。|
|[TaskHistoryRecordCollection](/javascript/api/excel/excel.taskhistoryrecordcollection)|[getCount()](/javascript/api/excel/excel.taskhistoryrecordcollection#getcount--)|获取任务集合中的历史记录记录数。|
||[getItemAt(index: number)](/javascript/api/excel/excel.taskhistoryrecordcollection#getitemat-index-)|使用任务历史记录记录在集合中的索引获取该记录。|
||[items](/javascript/api/excel/excel.taskhistoryrecordcollection#items)|获取此集合中已加载的子项。|
|[用户](/javascript/api/excel/excel.user)|[displayName](/javascript/api/excel/excel.user#displayname)|表示用户的显示名称。|
||[email](/javascript/api/excel/excel.user#email)|表示用户的电子邮件地址。|
||[uid](/javascript/api/excel/excel.user#uid)|表示用户的唯一 ID。|
|[Workbook](/javascript/api/excel/excel.workbook)|[linkedDataTypes](/javascript/api/excel/excel.workbook#linkeddatatypes)|返回属于工作簿的链接数据类型的集合。|
||[tasks](/javascript/api/excel/excel.workbook#tasks)|返回工作簿中的任务的集合。|
||[showPivotFieldList](/javascript/api/excel/excel.workbook#showpivotfieldlist)|指定是否在工作簿级别显示数据透视表的字段列表窗格。|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|如果工作簿使用 1904 日期系统，则为 True。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[namedSheetViews](/javascript/api/excel/excel.worksheet#namedsheetviews)|返回工作表中呈现的工作表视图的集合。|
||[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|在特定工作表上应用筛选器时发生。|
||[onFormulaChanged](/javascript/api/excel/excel.worksheet#onformulachanged)|在此工作表中更改一个或多个公式时发生。|
||[tasks](/javascript/api/excel/excel.worksheet#tasks)|返回工作表中的任务集合。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|将工作簿的指定工作表插入当前工作簿。|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|在工作簿中应用任何工作表的筛选器时发生。|
||[onFormulaChanged](/javascript/api/excel/excel.worksheetcollection#onformulachanged)|在此集合的任何工作表中更改一个或多个公式时发生。|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|获取应用筛选器的工作表的 ID。|
|[WorksheetFormulaChangedEventArgs](/javascript/api/excel/excel.worksheetformulachangedeventargs)|[formulaDetails](/javascript/api/excel/excel.worksheetformulachangedeventargs#formuladetails)|获取 FormulaChangedEventDetail 对象的数组，其中包含有关所有已更改公式的详细信息。|
||[source](/javascript/api/excel/excel.worksheetformulachangedeventargs#source)|事件的源。|
||[type](/javascript/api/excel/excel.worksheetformulachangedeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.worksheetformulachangedeventargs#worksheetid)|获取公式发生更改的工作表的 ID。|

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 参考文档](/javascript/api/excel?view=excel-js-preview&preserve-view=true)
- [Excel JavaScript API 要求集](excel-api-requirement-sets.md)
