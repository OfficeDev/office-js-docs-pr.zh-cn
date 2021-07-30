---
title: Excel JavaScript 预览 API
description: 有关即将推出的 JavaScript Excel的详细信息。
ms.date: 07/23/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 4bceda6229270332ed7624b693913e47a065a066
ms.sourcegitcommit: 3cc8f6adee0c7c68c61a42da0d97ed5ea61be0ac
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/30/2021
ms.locfileid: "53661276"
---
# <a name="excel-javascript-preview-apis"></a>Excel JavaScript 预览 API

新的 Excel JavaScript API 首先在“预览版”中引入，在进行充分测试并获得用户反馈后，它将成为编号的特定要求集的一部分。

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

下表提供了 API 的简要摘要，而后续 [API](#api-list) 列表表提供了一个详细列表。

| 功能区域 | 说明 | 相关对象 |
|:--- |:--- |:--- |
| 图表数据表 | 控制图表上数据表的外观、格式和可见性。 | [](/javascript/api/excel/excel.chart)Chart、ChartDataTable、ChartDataTableFormat [](/javascript/api/excel/excel.chartdatatable) [](/javascript/api/excel/excel.chartdatatableformat) |
| 记录任务 | 将注释转换为分配给用户的任务。 | [DocumentTask](/javascript/api/excel/excel.documenttask) |
| 身份 | 管理用户标识，包括显示名称和电子邮件地址。 | [](/javascript/api/excel/excel.identity) [Identity、IdentityCollection、IdentityEntity](/javascript/api/excel/excel.identitycollection) [](/javascript/api/excel/excel.identityentity) |
| 链接的数据类型 | 添加对从外部源连接到Excel类型的支持。 | [LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|
| 链接的工作簿 | 管理工作簿之间的链接，包括对刷新和断开工作簿链接的支持。 | [LinkedWorkbook](/javascript/api/excel/excel.linkedworkbook) [、LinkedWorkbookCollection](/javascript/api/excel/excel.linkedworkbookcollection) |
| 表样式 | 提供对字体、边框、填充颜色以及表格样式的其他方面的控制。 | [表](/javascript/api/excel/excel.table)、[数据透视表](/javascript/api/excel/excel.pivottable)[、切片器](/javascript/api/excel/excel.slicer) |
| 查询 | 检索查询属性，如名称、刷新日期和查询计数。 | [Query](/javascript/api/excel/excel.query) [、QueryCollection](/javascript/api/excel/excel.querycollection)|

## <a name="api-list"></a>API 列表

下表列出了当前预览Excel JavaScript API 的列表。 有关所有 JavaScript EXCEL的完整列表 (包括预览 API 和以前发布的 API) ，请参阅所有 Excel [JavaScript API。](/javascript/api/excel?view=excel-js-preview&preserve-view=true)

| 类 | 域 | 说明 |
|:---|:---|:---|
|[ChangeDirectionState](/javascript/api/excel/excel.changedirectionstate)|[deleteShiftDirection](/javascript/api/excel/excel.changedirectionstate#deleteshiftdirection)|代表在 (单元格时剩余单元格) 移动的方向，例如向上或向左移动。|
||[insertShiftDirection](/javascript/api/excel/excel.changedirectionstate#insertshiftdirection)|代表插入 (单元格时现有单元格) 向右或向下移动的方向。|
|[Chart](/javascript/api/excel/excel.chart)|[getDataTable () ](/javascript/api/excel/excel.chart#getdatatable--)|获取图表上的数据表。|
||[getDataTableOrNullObject () ](/javascript/api/excel/excel.chart#getdatatableornullobject--)|获取图表上的数据表。|
|[ChartDataTable](/javascript/api/excel/excel.chartdatatable)|[format](/javascript/api/excel/excel.chartdatatable#format)|表示图表数据表的格式，包括填充、字体和边框格式。|
||[showHorizontalBorder](/javascript/api/excel/excel.chartdatatable#showhorizontalborder)|指定是否显示数据表的水平边框。|
||[showLegendKey](/javascript/api/excel/excel.chartdatatable#showlegendkey)|指定是否显示数据表的 legendkey。|
||[showOutlineBorder](/javascript/api/excel/excel.chartdatatable#showoutlineborder)|指定是否显示数据表的外边框。|
||[showVerticalBorder](/javascript/api/excel/excel.chartdatatable#showverticalborder)|指定是否显示数据表的垂直边框。|
||[visible](/javascript/api/excel/excel.chartdatatable#visible)|指定是否显示图表的数据表。|
|[ChartDataTableFormat](/javascript/api/excel/excel.chartdatatableformat)|[border](/javascript/api/excel/excel.chartdatatableformat#border)|表示图表数据表的边框格式，其中包括颜色、线条样式和粗细。|
||[fill](/javascript/api/excel/excel.chartdatatableformat#fill)|表示对象的填充格式，包括背景格式信息。|
||[font](/javascript/api/excel/excel.chartdatatableformat#font)|表示当前 (字体名称、字号和颜色) 字体属性。|
|[Comment](/javascript/api/excel/excel.comment)|[assignTask (：Identity) ](/javascript/api/excel/excel.comment#assigntask-assignee-)|将附加到注释的任务作为委派者分配给给定用户。|
||[getTask () ](/javascript/api/excel/excel.comment#gettask--)|获取与此注释关联的任务。|
||[getTaskOrNullObject () ](/javascript/api/excel/excel.comment#gettaskornullobject--)|获取与此注释关联的任务。|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[getItemOrNullObject (commentId： string) ](/javascript/api/excel/excel.commentcollection#getitemornullobject-commentid-)|根据其 ID 从集合中获取批注。|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[assignTask (：Identity) ](/javascript/api/excel/excel.commentreply#assigntask-assignee-)|将附加到注释的任务分配给指定用户作为唯一的代理人。|
||[getTask () ](/javascript/api/excel/excel.commentreply#gettask--)|获取与此批注回复线程相关联的任务。|
||[getTaskOrNullObject () ](/javascript/api/excel/excel.commentreply#gettaskornullobject--)|获取与此批注回复线程相关联的任务。|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[getItemOrNullObject (commentReplyId： string) ](/javascript/api/excel/excel.commentreplycollection#getitemornullobject-commentreplyid-)|返回由其 ID 标识的批注回复。|
|[ConditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|[getItemOrNullObject(id: string)](/javascript/api/excel/excel.conditionalformatcollection#getitemornullobject-id-)|返回由 ID 标识的条件格式。|
|[DocumentTask](/javascript/api/excel/excel.documenttask)|[percentComplete](/javascript/api/excel/excel.documenttask#percentcomplete)|指定任务的完成百分比。|
||[priority](/javascript/api/excel/excel.documenttask#priority)|指定任务的优先级。|
||[assignees](/javascript/api/excel/excel.documenttask#assignees)|返回任务的被分配者的集合。|
||[更改](/javascript/api/excel/excel.documenttask#changes)|获取任务的更改记录。|
||[comment](/javascript/api/excel/excel.documenttask#comment)|获取与任务关联的注释。|
||[completedBy](/javascript/api/excel/excel.documenttask#completedby)|获取已完成该任务的最新用户。|
||[completedDateTime](/javascript/api/excel/excel.documenttask#completeddatetime)|获取任务的完成日期和时间。|
||[createdBy](/javascript/api/excel/excel.documenttask#createdby)|获取创建任务的用户。|
||[createdDateTime](/javascript/api/excel/excel.documenttask#createddatetime)|获取任务的创建日期和时间。|
||[id](/javascript/api/excel/excel.documenttask#id)|获取任务的 ID。|
||[setStartAndDueDateTime (startDateTime： Date， dueDateTime： Date) ](/javascript/api/excel/excel.documenttask#setstartandduedatetime-startdatetime--duedatetime-)|更改任务的开始日期和截止日期。|
||[startAndDueDateTime](/javascript/api/excel/excel.documenttask#startandduedatetime)|获取或设置任务应开始和到期的日期和时间。|
||[title](/javascript/api/excel/excel.documenttask#title)|指定任务的标题。|
|[DocumentTaskChange](/javascript/api/excel/excel.documenttaskchange)|[被分派人](/javascript/api/excel/excel.documenttaskchange#assignee)|表示分配给更改记录类型的任务的用户，或者从更改记录类型的任务 `assign` 中取消 `unassign` 分配的用户。|
||[changedBy](/javascript/api/excel/excel.documenttaskchange#changedby)|表示创建或更改任务的用户。|
||[commentId](/javascript/api/excel/excel.documenttaskchange#commentid)|表示 任务更改锁定的 或 `Comment` `CommentReply` 的 ID。|
||[createdDateTime](/javascript/api/excel/excel.documenttaskchange#createddatetime)|表示任务更改记录的创建日期和时间。|
||[dueDateTime](/javascript/api/excel/excel.documenttaskchange#duedatetime)|表示任务的截止日期和时间，以 UTC 时区表示。|
||[id](/javascript/api/excel/excel.documenttaskchange#id)|任务更改记录的 ID。|
||[percentComplete](/javascript/api/excel/excel.documenttaskchange#percentcomplete)|表示任务的完成百分比。|
||[priority](/javascript/api/excel/excel.documenttaskchange#priority)|表示任务的优先级。|
||[startDateTime](/javascript/api/excel/excel.documenttaskchange#startdatetime)|表示任务的开始日期和时间，以 UTC 时区表示。|
||[title](/javascript/api/excel/excel.documenttaskchange#title)|表示任务的标题。|
||[type](/javascript/api/excel/excel.documenttaskchange#type)|表示任务更改记录的操作类型。|
||[undoHistoryId](/javascript/api/excel/excel.documenttaskchange#undohistoryid)|表示 `DocumentTaskChange.id` 对更改记录类型撤消 `undo` 的属性。|
|[DocumentTaskChangeCollection](/javascript/api/excel/excel.documenttaskchangecollection)|[getCount()](/javascript/api/excel/excel.documenttaskchangecollection#getcount--)|获取任务集合中的更改记录数。|
||[getItemAt(index: number)](/javascript/api/excel/excel.documenttaskchangecollection#getitemat-index-)|使用任务更改记录在集合中的索引获取该记录。|
||[items](/javascript/api/excel/excel.documenttaskchangecollection#items)|获取此集合中已加载的子项。|
|[DocumentTaskCollection](/javascript/api/excel/excel.documenttaskcollection)|[getCount()](/javascript/api/excel/excel.documenttaskcollection#getcount--)|获取集合中的任务数。|
||[getItem(key: string)](/javascript/api/excel/excel.documenttaskcollection#getitem-key-)|使用其 ID 获取任务。|
||[getItemAt(index: number)](/javascript/api/excel/excel.documenttaskcollection#getitemat-index-)|按任务在集合中的索引获取任务。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.documenttaskcollection#getitemornullobject-key-)|使用其 ID 获取任务。|
||[items](/javascript/api/excel/excel.documenttaskcollection#items)|获取此集合中已加载的子项。|
|[DocumentTaskSchedule](/javascript/api/excel/excel.documenttaskschedule)|[dueDateTime](/javascript/api/excel/excel.documenttaskschedule#duedatetime)|获取任务到期的日期和时间。|
||[startDateTime](/javascript/api/excel/excel.documenttaskschedule#startdatetime)|获取任务应开始的日期和时间。|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.groupshapecollection#getitemornullobject-key-)|使用形状的名称或 ID 获取形状。|
|[身份](/javascript/api/excel/excel.identity)|[displayName](/javascript/api/excel/excel.identity#displayname)|表示用户的显示名称。|
||[email](/javascript/api/excel/excel.identity#email)|表示用户的电子邮件地址。|
||[id](/javascript/api/excel/excel.identity#id)|表示用户的唯一 ID。|
|[IdentityCollection](/javascript/api/excel/excel.identitycollection)|[add (assignee： Identity) ](/javascript/api/excel/excel.identitycollection#add-assignee-)|向集合添加用户标识。|
||[clear()](/javascript/api/excel/excel.identitycollection#clear--)|从集合中删除所有的用户标识。|
||[getCount()](/javascript/api/excel/excel.identitycollection#getcount--)|获取集合中项的数目。|
||[getItemAt(index: number)](/javascript/api/excel/excel.identitycollection#getitemat-index-)|使用文档在集合中的索引获取文档用户标识。|
||[items](/javascript/api/excel/excel.identitycollection#items)|获取此集合中已加载的子项。|
||[删除 (：标识) ](/javascript/api/excel/excel.identitycollection#remove-assignee-)|从集合中删除用户标识。|
|[IdentityEntity](/javascript/api/excel/excel.identityentity)|[displayName](/javascript/api/excel/excel.identityentity#displayname)|表示用户的显示名称。|
||[email](/javascript/api/excel/excel.identityentity#email)|表示用户的电子邮件地址。|
||[id](/javascript/api/excel/excel.identityentity#id)|表示用户的唯一 ID。|
|[LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|[dataProvider](/javascript/api/excel/excel.linkeddatatype#dataprovider)|链接数据提供程序的数据提供程序数据类型。|
||[lastRefreshed](/javascript/api/excel/excel.linkeddatatype#lastrefreshed)|自上次刷新链接工作簿时打开工作簿以来的本地数据类型日期和时间。|
||[名称](/javascript/api/excel/excel.linkeddatatype#name)|链接对象数据类型。|
||[periodicRefreshInterval](/javascript/api/excel/excel.linkeddatatype#periodicrefreshinterval)|链接对象刷新频率（以秒数据类型设置为 `refreshMode` "Periodic"时刷新。|
||[refreshMode](/javascript/api/excel/excel.linkeddatatype#refreshmode)|用于检索链接数据数据类型的机制。|
||[服务 Id](/javascript/api/excel/excel.linkeddatatype#serviceid)|链接对象的唯一数据类型。|
||[supportedRefreshModes](/javascript/api/excel/excel.linkeddatatype#supportedrefreshmodes)|返回一个数组，该数组包含链接对象支持的所有数据类型。|
||[requestRefresh () ](/javascript/api/excel/excel.linkeddatatype#requestrefresh--)|请求刷新链接数据类型。|
||[requestSetRefreshMode (refreshMode： Excel。LinkedDataTypeRefreshMode) ](/javascript/api/excel/excel.linkeddatatype#requestsetrefreshmode-refreshmode-)|请求更改此链接的刷新数据类型。|
|[LinkedDataTypeAddedEventArgs](/javascript/api/excel/excel.linkeddatatypeaddedeventargs)|[服务 Id](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#serviceid)|新链接对象的唯一 ID 数据类型。|
||[source](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#source)|获取事件源。|
||[type](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#type)|获取事件的类型。|
|[LinkedDataTypeCollection](/javascript/api/excel/excel.linkeddatatypecollection)|[getCount()](/javascript/api/excel/excel.linkeddatatypecollection#getcount--)|获取集合中链接的数据类型的数量。|
||[getItem (键：number) ](/javascript/api/excel/excel.linkeddatatypecollection#getitem-key-)|按服务 ID 数据类型链接的标识符。|
||[getItemAt(index: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitemat-index-)|按集合数据类型索引获取链接对象。|
||[getItemOrNullObject (键：number) ](/javascript/api/excel/excel.linkeddatatypecollection#getitemornullobject-key-)|按 ID 数据类型链接的标识符。|
||[items](/javascript/api/excel/excel.linkeddatatypecollection#items)|获取此集合中已加载的子项。|
||[requestRefreshAll () ](/javascript/api/excel/excel.linkeddatatypecollection#requestrefreshall--)|请求刷新集合中所有链接的数据类型。|
|[LinkedWorkbook](/javascript/api/excel/excel.linkedworkbook)|[breakLinks () ](/javascript/api/excel/excel.linkedworkbook#breaklinks--)|请求断开指向链接工作簿的链接。|
||[id](/javascript/api/excel/excel.linkedworkbook#id)|指向链接工作簿的原始 URL。|
||[refresh()](/javascript/api/excel/excel.linkedworkbook#refresh--)|请求刷新从链接工作簿检索到的数据。|
|[LinkedWorkbookCollection](/javascript/api/excel/excel.linkedworkbookcollection)|[breakAllLinks () ](/javascript/api/excel/excel.linkedworkbookcollection#breakalllinks--)|断开指向链接工作簿的所有链接。|
||[getItem(key: string)](/javascript/api/excel/excel.linkedworkbookcollection#getitem-key-)|按 URL 获取链接工作簿的信息。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.linkedworkbookcollection#getitemornullobject-key-)|按 URL 获取链接工作簿的信息。|
||[items](/javascript/api/excel/excel.linkedworkbookcollection#items)|获取此集合中已加载的子项。|
||[refreshAll () ](/javascript/api/excel/excel.linkedworkbookcollection#refreshall--)|请求刷新所有工作簿链接。|
||[workbookLinksRefreshMode](/javascript/api/excel/excel.linkedworkbookcollection#workbooklinksrefreshmode)|表示工作簿链接的更新模式。|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getitemornullobject-key-)|使用工作表视图的名称获取工作表视图。|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|根据数据层次结构以及各自层次结构的行和列项，获取数据透视表中的唯一单元格。|
||[pivotStyle](/javascript/api/excel/excel.pivotlayout#pivotstyle)|应用于数据透视表的样式。|
||[setStyle (样式：string \| PivotTableStyle \| BuiltInPivotTableStyle) ](/javascript/api/excel/excel.pivotlayout#setstyle-style-)|设置应用于数据透视表的样式。|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getFirstOrNullObject () ](/javascript/api/excel/excel.pivottablescopedcollection#getfirstornullobject--)|获取集合中的第一个数据透视表。|
|[Query](/javascript/api/excel/excel.query)|[error](/javascript/api/excel/excel.query#error)|从上次刷新查询时获取查询错误消息。|
||[loadedTo](/javascript/api/excel/excel.query#loadedto)|获取查询"加载到"对象类型。|
||[loadedToDataModel](/javascript/api/excel/excel.query#loadedtodatamodel)|指定是否将查询加载到数据模型。|
||[名称](/javascript/api/excel/excel.query#name)|获取查询的名称。|
||[refreshDate](/javascript/api/excel/excel.query#refreshdate)|获取上次刷新查询的日期和时间。|
||[rowsLoadedCount](/javascript/api/excel/excel.query#rowsloadedcount)|获取上次刷新查询时加载的行数。|
|[QueryCollection](/javascript/api/excel/excel.querycollection)|[getCount()](/javascript/api/excel/excel.querycollection#getcount--)|获取工作簿中的查询数。|
||[getItem(key: string)](/javascript/api/excel/excel.querycollection#getitem-key-)|根据名称从集合获取查询。|
||[items](/javascript/api/excel/excel.querycollection#items)|获取此集合中已加载的子项。|
|[区域](/javascript/api/excel/excel.range)|[getDependents () ](/javascript/api/excel/excel.range#getdependents--)|返回一个对象，该对象表示包含同一工作表或多个工作表中单元格的所有 `WorkbookRangeAreas` 从属单元格的范围。|
||[getPrecedents () ](/javascript/api/excel/excel.range#getprecedents--)|返回一个对象，该对象代表包含同一工作表或多个工作表中单元格的所有引用 `WorkbookRangeAreas` 单元格的范围。|
|[RefreshModeChangedEventArgs](/javascript/api/excel/excel.refreshmodechangedeventargs)|[refreshMode](/javascript/api/excel/excel.refreshmodechangedeventargs#refreshmode)|链接的数据类型刷新模式。|
||[服务 Id](/javascript/api/excel/excel.refreshmodechangedeventargs#serviceid)|刷新模式已更改的对象的唯一 ID。|
||[source](/javascript/api/excel/excel.refreshmodechangedeventargs#source)|获取事件源。|
||[type](/javascript/api/excel/excel.refreshmodechangedeventargs#type)|获取事件的类型。|
|[RefreshRequestCompletedEventArgs](/javascript/api/excel/excel.refreshrequestcompletedeventargs)|[已刷新](/javascript/api/excel/excel.refreshrequestcompletedeventargs#refreshed)|指示刷新请求是否成功。|
||[服务 Id](/javascript/api/excel/excel.refreshrequestcompletedeventargs#serviceid)|已完成刷新请求的对象的唯一 ID。|
||[source](/javascript/api/excel/excel.refreshrequestcompletedeventargs#source)|获取事件源。|
||[type](/javascript/api/excel/excel.refreshrequestcompletedeventargs#type)|获取事件的类型。|
||[warnings](/javascript/api/excel/excel.refreshrequestcompletedeventargs#warnings)|包含从刷新请求生成的任何警告的数组。|
|[Shape](/javascript/api/excel/excel.shape)|[displayName](/javascript/api/excel/excel.shape#displayname)|获取显示名称的大小。|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|从 XML 字符串创建可缩放的矢量图形 (SVG) 并将其添加到工作表。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.shapecollection#getitemornullobject-key-)|使用形状的名称或 ID 获取形状。|
|[Slicer](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameinformula)|表示公式中使用切片器名称。|
||[slicerStyle](/javascript/api/excel/excel.slicer#slicerstyle)|应用于切片器的样式。|
||[setStyle (样式：字符串 \| SlicerStyle \| BuiltInSlicerStyle) ](/javascript/api/excel/excel.slicer#setstyle-style-)|设置应用于切片器的样式。|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[getItemOrNullObject(name: string)](/javascript/api/excel/excel.stylecollection#getitemornullobject-name-)|按名称获取样式。|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|将表格更改为使用默认表格样式。|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|在将筛选器应用于特定表时发生。|
||[tableStyle](/javascript/api/excel/excel.table#tablestyle)|应用于表格的样式。|
||[setStyle (样式：string \| TableStyle \| BuiltInTableStyle) ](/javascript/api/excel/excel.table#setstyle-style-)|设置应用于表格的样式。|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|在工作簿或工作表的任何表上应用筛选器时发生。|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|获取应用筛选器的表的 ID。|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|获取包含表格的工作表的 ID。|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[deleteRows (行： number[] \| TableRow[]) ](/javascript/api/excel/excel.tablerowcollection#deleterows-rows-)|从表中删除多行。|
||[deleteRowsAt (索引： number， count？： number) ](/javascript/api/excel/excel.tablerowcollection#deleterowsat-index--count-)|从给定索引开始，从表中删除指定数量的行。|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.tablescopedcollection#getitemornullobject-key-)|按名称或 ID 获取表。|
|[Workbook](/javascript/api/excel/excel.workbook)|[linkedDataTypes](/javascript/api/excel/excel.workbook#linkeddatatypes)|返回属于工作簿的链接数据类型的集合。|
||[linkedWorkbooks](/javascript/api/excel/excel.workbook#linkedworkbooks)|返回链接工作簿的集合。|
||[查询](/javascript/api/excel/excel.workbook#queries)|返回属于工作簿的 Power Query 查询的集合。|
||[任务](/javascript/api/excel/excel.workbook#tasks)|返回工作簿中的任务集合。|
||[showPivotFieldList](/javascript/api/excel/excel.workbook#showpivotfieldlist)|指定是否在工作簿级别显示数据透视表的字段列表窗格。|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|如果工作簿使用 1904 日期系统，则为 True。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|在将筛选器应用于特定工作表时发生。|
||[onProtectionChanged](/javascript/api/excel/excel.worksheet#onprotectionchanged)|工作表保护状态更改时发生。|
||[tabId](/javascript/api/excel/excel.worksheet#tabid)|返回一个值，该值代表此工作表，该工作表可通过 Open Office XML 读取。|
||[任务](/javascript/api/excel/excel.worksheet#tasks)|返回工作表中的任务集合。|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[changeDirectionState](/javascript/api/excel/excel.worksheetchangedeventargs#changedirectionstate)|表示工作表中单元格在删除或插入时移动的方向的变化。|
||[triggerSource](/javascript/api/excel/excel.worksheetchangedeventargs#triggersource)|表示事件的触发源。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|将工作簿的指定工作表插入当前工作簿。|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|在工作簿中应用任何工作表的筛选器时发生。|
||[onProtectionChanged](/javascript/api/excel/excel.worksheetcollection#onprotectionchanged)|工作表保护状态更改时发生。|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|获取应用筛选器的工作表的 ID。|
|[WorksheetProtectionChangedEventArgs](/javascript/api/excel/excel.worksheetprotectionchangedeventargs)|[isProtected](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#isprotected)|获取工作表的当前保护状态。|
||[source](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#source)|事件的源。|
||[type](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#worksheetid)|获取其中保护状态发生更改的工作表的 ID。|

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 参考文档](/javascript/api/excel?view=excel-js-preview&preserve-view=true)
- [Excel JavaScript API 要求集](excel-api-requirement-sets.md)
