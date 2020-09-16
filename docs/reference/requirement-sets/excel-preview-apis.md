---
title: Excel JavaScript 预览 API
description: 有关即将推出的 Excel JavaScript Api 的详细信息。
ms.date: 09/15/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 9ddc1405d4bc13087780e8950b36d9b3b4b04069
ms.sourcegitcommit: ed2a98b6fb5b432fa99c6cefa5ce52965dc25759
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/16/2020
ms.locfileid: "47819789"
---
# <a name="excel-javascript-preview-apis"></a>Excel JavaScript 预览 API

新的 Excel JavaScript API 首先在“预览版”中引入，在进行充分测试并获得用户反馈后，它将成为编号的特定要求集的一部分。

第一个表提供了 API 的简明摘要，而后续表提供了详细列表。

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

| 功能区域 | 说明 | 相关对象 |
|:--- |:--- |:--- |
| 链接的数据类型 | 为从外部源连接到 Excel 的数据类型添加支持。 | [LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|
| 命名工作表视图 | 提供对每个用户的工作表视图的编程控制。 | [NamedSheetView](/javascript/api/excel/excel.namedsheetview) |

## <a name="api-list"></a>API 列表

下表列出了当前预览中的 Excel JavaScript Api。 若要查看所有 Excel JavaScript Api 的完整列表 (包括预览 Api 和之前发布的 Api) ，请参阅 [所有 Excel Javascript api](/javascript/api/excel?view=excel-js-preview&preserve-view=true)。

| Class | 域 | 说明 |
|:---|:---|:---|
|[LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|[dataProvider](/javascript/api/excel/excel.linkeddatatype#dataprovider)|链接数据类型的数据提供程序的名称。 当从服务中检索信息时，这可能会发生变化。|
||[lastRefreshed](/javascript/api/excel/excel.linkeddatatype#lastrefreshed)|上次刷新链接数据类型时，自工作簿打开时的本地时区日期和时间。|
||[name](/javascript/api/excel/excel.linkeddatatype#name)|链接的数据类型的名称。 当从服务中检索信息时，这可能会发生变化。|
||[periodicRefreshInterval](/javascript/api/excel/excel.linkeddatatype#periodicrefreshinterval)|如果 `refreshMode` 设置为 "定期"，则刷新链接的数据类型的频率（以秒为单位）。|
||[Microsoft.sharepoint.linq.refreshmode](/javascript/api/excel/excel.linkeddatatype#refreshmode)|检索链接数据类型的数据所依据的机制。|
||[serviceId](/javascript/api/excel/excel.linkeddatatype#serviceid)|链接的数据类型的唯一 id。|
||[supportedRefreshModes](/javascript/api/excel/excel.linkeddatatype#supportedrefreshmodes)|返回一个数组，其中包含已链接的数据类型支持的所有刷新模式。 当从服务中检索信息时，数组的内容可能会发生变化。|
||[requestRefresh ( # B1 ](/javascript/api/excel/excel.linkeddatatype#requestrefresh--)|发出刷新链接数据类型的请求。 如果服务正忙或暂时无法访问，则不会满足该请求。|
||[requestSetRefreshMode (Microsoft.sharepoint.linq.refreshmode： LinkedDataTypeRefreshMode) ](/javascript/api/excel/excel.linkeddatatype#requestsetrefreshmode-refreshmode-)|发出请求，以更改此链接数据类型的刷新模式。|
|[LinkedDataTypeAddedEventArgs](/javascript/api/excel/excel.linkeddatatypeaddedeventargs)|[serviceId](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#serviceid)|新的链接数据类型的唯一 id。|
||[source](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#source)|获取事件源。 有关详细信息，请参阅 Excel.EventSource。|
||[type](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#type)|获取事件的类型。 有关详细信息，请参阅 Excel.EventType。|
|[LinkedDataTypeCollection](/javascript/api/excel/excel.linkeddatatypecollection)|[getCount()](/javascript/api/excel/excel.linkeddatatypecollection#getcount--)|获取集合中链接的数据类型的数目。|
||[getItem (项： number) ](/javascript/api/excel/excel.linkeddatatypecollection#getitem-key-)|按服务 id 获取链接的数据类型。|
||[getItemAt(index: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitemat-index-)|按其在集合中的索引获取链接的数据类型。|
||[getItemOrNullObject (项： number) ](/javascript/api/excel/excel.linkeddatatypecollection#getitemornullobject-key-)|按 ID 获取链接的数据类型。 如果链接的数据类型不存在，则为其 `isNullObject` 属性设置为的对象 `true` 。 有关详细信息，请参阅 {@link https://docs.microsoft.com/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties | * OrNullObject 方法和属性}。|
||[items](/javascript/api/excel/excel.linkeddatatypecollection#items)|获取此集合中已加载的子项。|
||[requestRefreshAll ( # B1 ](/javascript/api/excel/excel.linkeddatatypecollection#requestrefreshall--)|发出请求以刷新集合中的所有链接数据类型。|
|[NamedSheetView](/javascript/api/excel/excel.namedsheetview)|[activate()](/javascript/api/excel/excel.namedsheetview#activate--)|激活此工作表视图。 这等效于在 Excel UI 中使用 "切换到"。|
||[delete()](/javascript/api/excel/excel.namedsheetview#delete--)|将工作表视图从工作表中删除。|
||[重复 (名称？： string) ](/javascript/api/excel/excel.namedsheetview#duplicate-name-)|创建此工作表视图的副本。|
||[name](/javascript/api/excel/excel.namedsheetview#name)|获取或设置工作表视图的名称。|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[add(name: string)](/javascript/api/excel/excel.namedsheetviewcollection#add-name-)|创建具有给定名称的新工作表视图。|
||[enterTemporary ( # B1 ](/javascript/api/excel/excel.namedsheetviewcollection#entertemporary--)|创建并激活一个新的临时表视图。|
||[退出 ( # B1 ](/javascript/api/excel/excel.namedsheetviewcollection#exit--)|退出当前的活动工作表视图。|
||[getActive ( # B1 ](/javascript/api/excel/excel.namedsheetviewcollection#getactive--)|获取工作表的当前活动工作表视图。|
||[getCount()](/javascript/api/excel/excel.namedsheetviewcollection#getcount--)|获取此工作表中的工作表视图数。|
||[getItem(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getitem-key-)|使用其名称获取工作表视图。|
||[getItemAt(index: number)](/javascript/api/excel/excel.namedsheetviewcollection#getitemat-index-)|按其在集合中的索引获取工作表视图。|
||[items](/javascript/api/excel/excel.namedsheetviewcollection#items)|获取此集合中已加载的子项。|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[altTextDescription](/javascript/api/excel/excel.pivotlayout#alttextdescription)|数据透视表的替换文字说明。|
||[altTextTitle](/javascript/api/excel/excel.pivotlayout#alttexttitle)|数据透视表的 alt 文本标题。|
||[displayBlankLineAfterEachItem (显示： boolean) ](/javascript/api/excel/excel.pivotlayout#displayblanklineaftereachitem-display-)|设置是否在每个项目后显示一个空行。 这是在数据透视表的全局级别上设置的，并应用于各个透视字段。|
||[emptyCellText](/javascript/api/excel/excel.pivotlayout#emptycelltext)|自动填充到数据透视表中的任何空单元格的文本（如果有） `fillEmptyCells == true` 。|
||[fillEmptyCells](/javascript/api/excel/excel.pivotlayout#fillemptycells)|指定是否应使用数据透视表中的空单元格填充 `emptyCellText` 。 默认值为 False。|
||[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|根据数据层次结构以及各自层次结构的行和列项，获取数据透视表中的唯一单元格。 返回的单元格是给定行和列的交集，其中包含来自给定层次结构的数据。 此方法与在特定单元格上调用 getPivotItems 和 getDataHierarchy 相反。|
||[repeatAllItemLabels (repeatLabels： boolean) ](/javascript/api/excel/excel.pivotlayout#repeatallitemlabels-repeatlabels-)|在数据透视表中的所有字段之间设置 "重复所有项目标签" 设置。|
||[setStyle (样式： string \| PivotTableStyle \| BuiltInPivotTableStyle) ](/javascript/api/excel/excel.pivotlayout#setstyle-style-)|设置应用于数据透视表的样式。|
||[showFieldHeaders](/javascript/api/excel/excel.pivotlayout#showfieldheaders)|指定数据透视表是否显示字段标题 (字段标题和筛选下拉) 。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[refreshOnOpen](/javascript/api/excel/excel.pivottable#refreshonopen)|指定在工作簿打开时是否刷新数据透视表。 与 UI 中的 "加载时刷新" 设置相对应。|
|[区域](/javascript/api/excel/excel.range)|[getMergedAreas ( # B1 ](/javascript/api/excel/excel.range#getmergedareas--)|返回一个 `RangeAreas` 对象，表示此范围中的合并区域。 请注意，如果此范围中的合并区域计数超过512，API 将无法返回结果。|
||[getPrecedents ( # B1 ](/javascript/api/excel/excel.range#getprecedents--)|返回一个 `WorkbookRangeAreas` 对象，表示包含同一工作表或多个工作表中的单元格的所有引用单元格的区域。|
|[RefreshModeChangedEventArgs](/javascript/api/excel/excel.refreshmodechangedeventargs)|[Microsoft.sharepoint.linq.refreshmode](/javascript/api/excel/excel.refreshmodechangedeventargs#refreshmode)|链接的数据类型刷新模式。|
||[serviceId](/javascript/api/excel/excel.refreshmodechangedeventargs#serviceid)|已更改其刷新模式的对象的唯一 id。|
||[source](/javascript/api/excel/excel.refreshmodechangedeventargs#source)|获取事件源。 有关详细信息，请参阅 Excel.EventSource。|
||[type](/javascript/api/excel/excel.refreshmodechangedeventargs#type)|获取事件的类型。 有关详细信息，请参阅 Excel.EventType。|
|[RefreshRequestCompletedEventArgs](/javascript/api/excel/excel.refreshrequestcompletedeventargs)|[焕然一新](/javascript/api/excel/excel.refreshrequestcompletedeventargs#refreshed)|指示刷新请求是否成功。|
||[serviceId](/javascript/api/excel/excel.refreshrequestcompletedeventargs#serviceid)|已完成其刷新请求的对象的唯一 id。|
||[source](/javascript/api/excel/excel.refreshrequestcompletedeventargs#source)|获取事件源。 有关详细信息，请参阅 Excel.EventSource。|
||[type](/javascript/api/excel/excel.refreshrequestcompletedeventargs#type)|获取事件的类型。 有关详细信息，请参阅 Excel.EventType。|
||[发出](/javascript/api/excel/excel.refreshrequestcompletedeventargs#warnings)|包含从刷新请求生成的任何警告的数组。|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|从 XML 字符串创建可缩放的矢量图形 (SVG) 并将其添加到工作表。 返回表示新图片的 Shape 对象。|
|[Slicer](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameinformula)|表示公式中使用切片器名称。|
||[setStyle (样式： string \| SlicerStyle \| BuiltInSlicerStyle) ](/javascript/api/excel/excel.slicer#setstyle-style-)|设置应用于切片器的样式。|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|将表格更改为使用默认表格样式。|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|在特定表格上应用筛选器时发生。|
||[tableStyle](/javascript/api/excel/excel.table#tablestyle)|应用于表的样式。|
||[setStyle (样式： string \| TableStyle \| BuiltInTableStyle) ](/javascript/api/excel/excel.table#setstyle-style-)|设置应用于表的样式。|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|在工作簿或工作表中的任何表格上应用筛选器时发生。|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|获取应用了筛选器的表的 id。|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#type)|获取事件的类型。 有关详细信息，请参阅 Excel.EventType。|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|获取包含表的工作表的 id。|
|[Workbook](/javascript/api/excel/excel.workbook)|[linkedDataTypes](/javascript/api/excel/excel.workbook#linkeddatatypes)|返回属于工作簿的链接数据类型的集合。|
||[showPivotFieldList](/javascript/api/excel/excel.workbook#showpivotfieldlist)|指定是否在工作簿级别显示数据透视表的 "字段列表" 窗格。|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|如果工作簿使用 1904 日期系统，则为 True。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[namedSheetViews](/javascript/api/excel/excel.worksheet#namedsheetviews)|返回工作表视图的集合，这些视图显示在工作表中。|
||[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|在特定工作表上应用筛选器时发生。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|将工作簿的指定工作表插入当前工作簿。|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|在工作簿中应用任何工作表的筛选器时发生。|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|获取事件的类型。 有关详细信息，请参阅 Excel.EventType。|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|获取应用了筛选器的工作表的 id。|

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 参考文档](/javascript/api/excel?view=excel-js-preview&preserve-view=true)
- [Excel JavaScript API 要求集](excel-api-requirement-sets.md)
