---
title: Excel JavaScript 预览 API
description: 有关即将推出的 Excel JavaScript Api 的详细信息
ms.date: 06/29/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: d1701ad393b96e33f0007bfcb5609c93c13608a2
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/11/2020
ms.locfileid: "47430763"
---
# <a name="excel-javascript-preview-apis"></a>Excel JavaScript 预览 API

新的 Excel JavaScript API 首先在“预览版”中引入，在进行充分测试并获得用户反馈后，它将成为编号的特定要求集的一部分。

第一个表提供了 API 的简明摘要，而后续表提供了详细列表。

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

| 功能区域 | 说明 | 相关对象 |
|:--- |:--- |:--- |
| 日期和时间 [区域性设置](../../excel/excel-add-ins-workbooks.md#access-application-culture-settings) | 提供有关日期和时间格式设置的其他区域性设置的访问权限。 | [CultureInfo](/javascript/api/excel/excel.cultureinfo)、 [NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo) [应用程序](/javascript/api/excel/excel.application) |
| [插入工作簿](../../excel/excel-add-ins-workbooks.md#insert-a-copy-of-an-existing-workbook-into-the-current-one-preview) | 将一个工作簿插入另一个工作簿。  | [Workbook](/javascript/api/excel/excel.worksheetcollection) |
| 透视筛选器 | 对数据透视表的字段应用数值驱动的筛选器。 | [透视字段](/javascript/api/excel/excel.pivotfield#applyfilter-filter-)、 [PivotFilters](/javascript/api/excel/excel.pivotFilters) |
|区域 spilling | 允许外接程序查找与 [动态数组](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531) 结果关联的范围。 | [区域](/javascript/api/excel/excel.range) |

## <a name="api-list"></a>API 列表

下表列出了当前预览中的 Excel JavaScript Api。 若要查看所有 Excel JavaScript Api 的完整列表 (包括预览 Api 和之前发布的 Api) ，请参阅 [所有 Excel Javascript api](/javascript/api/excel?view=excel-js-preview&preserve-view=true)。

| Class | 域 | 说明 |
|:---|:---|:---|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[getDimensionValues (维： ChartSeriesDimension) ](/javascript/api/excel/excel.chartseries#getdimensionvalues-dimension-)|从图表系列的单个维中获取值。 这些值可以是类别值，也可以是数据值，具体取决于指定的维度和为图表系列映射数据的方式。|
|[Comment](/javascript/api/excel/excel.comment)|[contentType](/javascript/api/excel/excel.comment#contenttype)|获取注释的内容类型。|
|[CommentAddedEventArgs](/javascript/api/excel/excel.commentaddedeventargs)|[commentDetails](/javascript/api/excel/excel.commentaddedeventargs#commentdetails)|获取 `CommentDetail` 包含其相关答复的注释 ID 和 id 的数组。|
||[source](/javascript/api/excel/excel.commentaddedeventargs#source)|指定时间源。 有关详细信息，请参阅 `Excel.EventSource`。|
||[type](/javascript/api/excel/excel.commentaddedeventargs#type)|获取事件的类型。 有关详细信息，请参阅 `Excel.EventType`。|
||[worksheetId](/javascript/api/excel/excel.commentaddedeventargs#worksheetid)|获取发生事件的工作表的 ID。|
|[CommentChangedEventArgs](/javascript/api/excel/excel.commentchangedeventargs)|[changeType](/javascript/api/excel/excel.commentchangedeventargs#changetype)|获取表示已更改事件的触发方式的更改类型。|
||[commentDetails](/javascript/api/excel/excel.commentchangedeventargs#commentdetails)|获取 `CommentDetail` 包含其相关答复的注释 ID 和 id 的数组。|
||[source](/javascript/api/excel/excel.commentchangedeventargs#source)|指定时间源。 有关详细信息，请参阅 `Excel.EventSource`。|
||[type](/javascript/api/excel/excel.commentchangedeventargs#type)|获取事件的类型。 有关详细信息，请参阅 `Excel.EventType`。|
||[worksheetId](/javascript/api/excel/excel.commentchangedeventargs#worksheetid)|获取发生事件的工作表的 ID。|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[onAdded](/javascript/api/excel/excel.commentcollection#onadded)|添加注释时发生。|
||[onChanged](/javascript/api/excel/excel.commentcollection#onchanged)|当注释集合中的批注或答复发生更改时发生，包括答复被删除的时间。|
||[onDeleted](/javascript/api/excel/excel.commentcollection#ondeleted)|在注释集合中删除批注时发生。|
|[CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs)|[commentDetails](/javascript/api/excel/excel.commentdeletedeventargs#commentdetails)|获取 `CommentDetail` 包含其相关答复的注释 ID 和 id 的数组。|
||[source](/javascript/api/excel/excel.commentdeletedeventargs#source)|指定时间源。 有关详细信息，请参阅 `Excel.EventSource`。|
||[type](/javascript/api/excel/excel.commentdeletedeventargs#type)|获取事件的类型。 有关详细信息，请参阅 `Excel.EventType`。|
||[worksheetId](/javascript/api/excel/excel.commentdeletedeventargs#worksheetid)|获取发生事件的工作表的 ID。|
|[CommentDetail](/javascript/api/excel/excel.commentdetail)|[commentId](/javascript/api/excel/excel.commentdetail#commentid)|表示注释的 ID。|
||[replyIds](/javascript/api/excel/excel.commentdetail#replyids)|表示相关答复的 Id 属于注释。|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[contentType](/javascript/api/excel/excel.commentreply#contenttype)|答复的内容类型。|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[datetimeFormat](/javascript/api/excel/excel.cultureinfo#datetimeformat)|定义适当的区域性格式，以显示日期和时间。 这取决于当前的系统区域性设置。|
|[格式](/javascript/api/excel/excel.datetimeformatinfo)|[dateSeparator](/javascript/api/excel/excel.datetimeformatinfo#dateseparator)|获取用作日期分隔符的字符串。 这取决于当前的系统设置。|
||[longDatePattern](/javascript/api/excel/excel.datetimeformatinfo#longdatepattern)|获取长日期值的格式字符串。 这取决于当前的系统设置。|
||[longTimePattern](/javascript/api/excel/excel.datetimeformatinfo#longtimepattern)|获取长时间值的格式字符串。 这取决于当前的系统设置。|
||[shortDatePattern](/javascript/api/excel/excel.datetimeformatinfo#shortdatepattern)|获取短日期值的格式字符串。 这取决于当前的系统设置。|
||[timeSeparator](/javascript/api/excel/excel.datetimeformatinfo#timeseparator)|获取用作时间分隔符的字符串。 这取决于当前的系统设置。|
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
|[PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter)|[运算符](/javascript/api/excel/excel.pivotdatefilter#comparator)|比较运算符是其他值要与其进行比较的静态值。 比较的类型由条件定义。|
||[表达式](/javascript/api/excel/excel.pivotdatefilter#condition)|指定筛选器的条件，该条件定义了必要的筛选条件。|
||[异](/javascript/api/excel/excel.pivotdatefilter#exclusive)|如果为 true，则筛选 *排除* 满足条件的项目。 默认值为 false (筛选器以包含满足条件) 的项目。|
||[lowerBound](/javascript/api/excel/excel.pivotdatefilter#lowerbound)|筛选条件范围的下限 `Between` 。|
||[upperBound](/javascript/api/excel/excel.pivotdatefilter#upperbound)|筛选条件范围的上限 `Between` 。|
||[wholeDays](/javascript/api/excel/excel.pivotdatefilter#wholedays)|对于 `Equals` 、 `Before` 、 `After` 和 `Between` 筛选条件，指示是否应将比较作为全天进行。|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[applyFilter (筛选器： PivotFilters) ](/javascript/api/excel/excel.pivotfield#applyfilter-filter-)|设置一个或多个字段的当前 PivotFilters，并将其应用于字段。|
||[clearAllFilters ( # B1 ](/javascript/api/excel/excel.pivotfield#clearallfilters--)|从字段的所有筛选器中清除所有条件。 这将删除对字段的任何活动筛选。|
||[clearFilter (filterType： PivotFilterType) ](/javascript/api/excel/excel.pivotfield#clearfilter-filtertype-)|清除给定类型的字段筛选器中的所有现有条件 (如果当前已将其应用) 。|
||[getFilters ( # B1 ](/javascript/api/excel/excel.pivotfield#getfilters--)|获取当前应用于字段的所有筛选器。|
||[isFiltered (filterType？： PivotFilterType) ](/javascript/api/excel/excel.pivotfield#isfiltered-filtertype-)|检查字段上是否有任何已应用的筛选器。|
|[PivotFilters](/javascript/api/excel/excel.pivotfilters)|[dateFilter](/javascript/api/excel/excel.pivotfilters#datefilter)|透视字段当前应用的日期筛选器。 如果未应用任何值，则为 Null。|
||[labelFilter](/javascript/api/excel/excel.pivotfilters#labelfilter)|透视字段当前应用的标签筛选器。 如果未应用任何值，则为 Null。|
||[manualFilter](/javascript/api/excel/excel.pivotfilters#manualfilter)|透视字段当前应用的手动筛选。 如果未应用任何值，则为 Null。|
||[valueFilter](/javascript/api/excel/excel.pivotfilters#valuefilter)|透视字段当前应用的值筛选器。 如果未应用任何值，则为 Null。|
|[PivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter)|[运算符](/javascript/api/excel/excel.pivotlabelfilter#comparator)|比较运算符是其他值要与其进行比较的静态值。 比较的类型由条件定义。|
||[表达式](/javascript/api/excel/excel.pivotlabelfilter#condition)|指定筛选器的条件，该条件定义了必要的筛选条件。|
||[异](/javascript/api/excel/excel.pivotlabelfilter#exclusive)|如果为 true，则筛选 *排除* 满足条件的项目。 默认值为 false (筛选器以包含满足条件) 的项目。|
||[lowerBound](/javascript/api/excel/excel.pivotlabelfilter#lowerbound)|筛选条件之间的范围的下限。|
||[substring](/javascript/api/excel/excel.pivotlabelfilter#substring)|用于 `BeginsWith` 、 `EndsWith` 和筛选条件的子字符串 `Contains` 。|
||[upperBound](/javascript/api/excel/excel.pivotlabelfilter#upperbound)|筛选条件之间的范围的上限。|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|根据数据层次结构以及各自层次结构的行和列项，获取数据透视表中的唯一单元格。 返回的单元格是给定行和列的交集，其中包含来自给定层次结构的数据。 此方法与在特定单元格上调用 getPivotItems 和 getDataHierarchy 相反。|
||[pivotStyle](/javascript/api/excel/excel.pivotlayout#pivotstyle)|应用于数据透视表的样式。|
||[setStyle (样式： string \| PivotTableStyle \| BuiltInPivotTableStyle) ](/javascript/api/excel/excel.pivotlayout#setstyle-style-)|设置应用于数据透视表的样式。|
|[PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter)|[selectedItems](/javascript/api/excel/excel.pivotmanualfilter#selecteditems)|要手动筛选的选定项的列表。 这些项必须是所选字段中现有和有效的项。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[allowMultipleFiltersPerField](/javascript/api/excel/excel.pivottable#allowmultiplefiltersperfield)|指定数据透视表是否允许对表中给定的透视字段上的多个 PivotFilters 进行应用。|
|[PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter)|[运算符](/javascript/api/excel/excel.pivotvaluefilter#comparator)|比较运算符是其他值要与其进行比较的静态值。 比较的类型由条件定义。|
||[表达式](/javascript/api/excel/excel.pivotvaluefilter#condition)|指定筛选器的条件，该条件定义了必要的筛选条件。|
||[异](/javascript/api/excel/excel.pivotvaluefilter#exclusive)|如果为 true，则筛选 *排除* 满足条件的项目。 默认值为 false (筛选器以包含满足条件) 的项目。|
||[lowerBound](/javascript/api/excel/excel.pivotvaluefilter#lowerbound)|筛选条件范围的下限 `Between` 。|
||[selectionType](/javascript/api/excel/excel.pivotvaluefilter#selectiontype)|指定筛选器是用于顶部/底部 N 项、顶部/底部 N 百分比还是顶部/底部 N 求和。|
||[极限](/javascript/api/excel/excel.pivotvaluefilter#threshold)|要针对顶部/底部筛选条件筛选的项、百分比或 sum 的 "N" 阈值数。|
||[upperBound](/javascript/api/excel/excel.pivotvaluefilter#upperbound)|筛选条件范围的上限 `Between` 。|
||[value](/javascript/api/excel/excel.pivotvaluefilter#value)|筛选所依据的字段中所选的 "值" 的名称。|
|[区域](/javascript/api/excel/excel.range)|[getDirectPrecedents ( # B1 ](/javascript/api/excel/excel.range#getdirectprecedents--)|返回一个 `WorkbookRangeAreas` object 类型的对象，该对象代表包含同一工作表或多个工作表中的单元格的所有直接引用单元格的区域。|
||[getMergedAreas ( # B1 ](/javascript/api/excel/excel.range#getmergedareas--)|返回一个 RangeAreas 对象，该对象代表此区域中的合并区域。 请注意，如果此范围中的合并区域计数超过512，API 将无法返回结果。|
||[getPrecedents ( # B1 ](/javascript/api/excel/excel.range#getprecedents--)|返回一个 `WorkbookRangeAreas` 对象，表示包含同一工作表或多个工作表中的单元格的所有引用单元格的区域。|
||[getSpillParent()](/javascript/api/excel/excel.range#getspillparent--)|获取 Range 对象，它包含要将某个单元格溢出到的定位单元格。 如果应用于具有多个单元格的区域，则会失败。|
||[getSpillParentOrNullObject()](/javascript/api/excel/excel.range#getspillparentornullobject--)|获取 Range 对象，它包含要将某个单元格溢出到的定位单元格。|
||[getSpillingToRange()](/javascript/api/excel/excel.range#getspillingtorange--)|获取 Range 对象，它在调用定位单元格时包含溢出区域。 如果应用于具有多个单元格的区域，则会失败。|
||[getSpillingToRangeOrNullObject()](/javascript/api/excel/excel.range#getspillingtorangeornullobject--)|获取 Range 对象，它在调用定位单元格时包含溢出区域。|
||[hasSpill](/javascript/api/excel/excel.range#hasspill)|表示所有单元格是否都具有溢出边框。|
||[numberFormatCategories](/javascript/api/excel/excel.range#numberformatcategories)|表示每个单元格的数字格式的类别。|
||[savedAsArray](/javascript/api/excel/excel.range#savedasarray)|表示是否将所有单元格都保存为数组公式。|
|[RangeAreasCollection](/javascript/api/excel/excel.rangeareascollection)|[getCount()](/javascript/api/excel/excel.rangeareascollection#getcount--)|获取此集合中的 RangeAreas 对象的数目。|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangeareascollection#getitemat-index-)|根据集合中的位置返回 RangeAreas 对象。|
||[items](/javascript/api/excel/excel.rangeareascollection#items)|获取此集合中已加载的子项。|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|从 XML 字符串创建可缩放的矢量图形 (SVG) 并将其添加到工作表。 返回表示新图片的 Shape 对象。|
|[Slicer](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameinformula)|表示公式中使用切片器名称。|
||[slicerStyle](/javascript/api/excel/excel.slicer#slicerstyle)|应用于切片器的样式。|
||[setStyle (样式： string \| PivotTableStyle \| BuiltInSlicerStyle) ](/javascript/api/excel/excel.slicer#setstyle-style-)|设置应用于切片器的样式。|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|将表格更改为使用默认表格样式。|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|在特定表格上应用筛选器时发生。|
||[tableStyle](/javascript/api/excel/excel.table#tablestyle)|应用于表的样式。|
||[setStyle (样式： string \| PivotTableStyle \| BuiltInTableStyle) ](/javascript/api/excel/excel.table#setstyle-style-)|设置应用于切片器的样式。|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|在工作簿或工作表中的任何表格上应用筛选器时发生。|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|获取应用了筛选器的表的 id。|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#type)|获取事件的类型。 有关详细信息，请参阅 Excel.EventType。|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|获取包含表的工作表的 id。|
|[Workbook](/javascript/api/excel/excel.workbook)|[showPivotFieldList](/javascript/api/excel/excel.workbook#showpivotfieldlist)|指定是否在工作簿级别显示数据透视表的 "字段列表" 窗格。|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|如果工作簿使用 1904 日期系统，则为 True。|
|[WorkbookRangeAreas](/javascript/api/excel/excel.workbookrangeareas)|[getRangeAreasBySheet (项： string) ](/javascript/api/excel/excel.workbookrangeareas#getrangeareasbysheet-key-)|`RangeAreas`基于集合中的工作表 id 或名称返回对象。|
||[getRangeAreasOrNullObjectBySheet (项： string) ](/javascript/api/excel/excel.workbookrangeareas#getrangeareasornullobjectbysheet-key-)|`RangeAreas`基于集合中的工作表名称或 id 返回对象。 如果没有 Worksheet 对象，将返回 NULL 对象。|
||[地址](/javascript/api/excel/excel.workbookrangeareas#addresses)|返回 A1 样式的地址数组。 Address 值将包含单元格每个矩形块的工作表名称 (例如，"Sheet1！A1： B4、Sheet1！D1： D4 ") 。 只读。|
||[areas](/javascript/api/excel/excel.workbookrangeareas#areas)|返回 RangeAreasCollection 对象，集合中的每个 RangeAreas 代表一个工作表中的一个或多个矩形区域。|
||[区域](/javascript/api/excel/excel.workbookrangeareas#ranges)|返回构成此对象的范围的集合。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[customProperties](/javascript/api/excel/excel.worksheet#customproperties)|获取工作表级自定义属性的集合。|
||[namedSheetViews](/javascript/api/excel/excel.worksheet#namedsheetviews)|返回工作表视图的集合，这些视图显示在工作表中。|
||[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|在特定工作表上应用筛选器时发生。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|将工作簿的指定工作表插入当前工作簿。|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|在工作簿中应用任何工作表的筛选器时发生。|
|[WorksheetCustomProperty](/javascript/api/excel/excel.worksheetcustomproperty)|[delete()](/javascript/api/excel/excel.worksheetcustomproperty#delete--)|删除 custom property 对象。|
||[key](/javascript/api/excel/excel.worksheetcustomproperty#key)|获取 customProperty 的键。 自定义属性键不区分大小写。 密钥限制为255个字符 (较大的值将导致引发 "InvalidArgument" 错误。 ) |
||[value](/javascript/api/excel/excel.worksheetcustomproperty#value)|获取或设置自定义属性的值。|
|[WorksheetCustomPropertyCollection](/javascript/api/excel/excel.worksheetcustompropertycollection)|[add (key： string，value： string) ](/javascript/api/excel/excel.worksheetcustompropertycollection#add-key--value-)|添加映射到所提供的键的新自定义属性。 这将使用该密钥覆盖现有的自定义属性。|
||[getCount()](/javascript/api/excel/excel.worksheetcustompropertycollection#getcount--)|获取此工作表上的自定义属性的数目。|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitem-key-)|按键获取自定义属性对象（不区分大小写）。 如果自定义属性不存在，则引发此异常。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitemornullobject-key-)|按键获取自定义属性对象（不区分大小写）。 如果自定义属性不存在，则返回 null 对象。|
||[items](/javascript/api/excel/excel.worksheetcustompropertycollection#items)|获取此集合中已加载的子项。|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|获取事件的类型。 有关详细信息，请参阅 Excel.EventType。|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|获取应用了筛选器的工作表的 id。|

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 参考文档](/javascript/api/excel?view=excel-js-preview&preserve-view=true)
- [Excel JavaScript API 要求集](./excel-api-requirement-sets.md)
