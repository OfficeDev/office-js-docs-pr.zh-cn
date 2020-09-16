---
title: Excel JavaScript API 要求集1.12
description: 有关 ExcelApi 1.12 要求集的详细信息。
ms.date: 09/15/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: a88c511e90fe48e1a9997d19cb4a2851cb718f6b
ms.sourcegitcommit: ed2a98b6fb5b432fa99c6cefa5ce52965dc25759
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/16/2020
ms.locfileid: "47819839"
---
# <a name="whats-new-in-excel-javascript-api-112"></a>Excel JavaScript API 1.12 中的新增功能

通过添加用于跟踪动态数组和查找公式的直接引用单元格的 Api，ExcelApi 1.12 增加了对区域中的公式的支持。 它还添加了对数据透视表筛选器的 API 控制。 此外，在注释、区域性设置和自定义属性功能区域中也进行了改进。

| 功能区域 | 说明 | 相关对象 |
|:--- |:--- |:--- |
| [注释事件](../../excel/excel-add-ins-events.md) | 将添加、更改和删除的事件添加到注释集合中。| [CommentCollection](/javascript/api/excel/excel.commentcollection) |
| 日期和时间 [区域性设置](../../excel/excel-add-ins-workbooks.md#access-application-culture-settings) | 提供有关日期和时间格式设置的其他区域性设置的访问权限。 | [CultureInfo](/javascript/api/excel/excel.cultureinfo)、 [NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo) [应用程序](/javascript/api/excel/excel.application) |
| 直接引用单元格 | 返回用于计算单元格的公式的范围。| [区域](/javascript/api/excel/excel.range#getdirectprecedents--) |
| 透视筛选器 | 对数据透视表的字段应用数值驱动的筛选器。 | [透视字段](/javascript/api/excel/excel.pivotfield#applyfilter-filter-)、 [PivotFilters](/javascript/api/excel/excel.pivotFilters) |
| [区域 spilling](../../excel/excel-add-ins-ranges-advanced.md#handle-dynamic-arrays-and-spilling) | 允许外接程序查找与 [动态数组](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531) 结果关联的范围。 | [区域](/javascript/api/excel/excel.range) |
| [工作表级自定义属性](../../excel/excel-add-ins-workbooks.md#worksheet-level-custom-properties) | 允许自定义属性的范围限定为工作表级别，除了作用于工作簿级别之外。 | [WorksheetCustomProperty](/javascript/api/excel/excel.worksheetcustomproperty)、 [WorksheetCustomPropertyCollection](/javascript/api/excel/excel.worksheetcustompropertycollection)|

## <a name="api-list"></a>API 列表

下表列出了 Excel JavaScript API 要求集1.12 中的 Api。 若要查看 Excel JavaScript API 要求集1.12 或更早版本支持的所有 Api 的 API 参考文档，请参阅 [要求集1.12 或更早版本中的 Excel api](/javascript/api/excel?view=excel-js-1.12&preserve-view=true)。

| Class | 域 | 说明 |
|:---|:---|:---|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[textOrientation](/javascript/api/excel/excel.chartaxistitle#textorientation)|指定文本面向图表轴标题的角度。 该值应为-90 到90的整数或垂直方向的文本的整数180。|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[getDimensionValues (维： ChartSeriesDimension) ](/javascript/api/excel/excel.chartseries#getdimensionvalues-dimension-)|从图表系列的单个维中获取值。 这些值可以是类别值，也可以是数据值，具体取决于指定的维度和为图表系列映射数据的方式。|
|[Comment](/javascript/api/excel/excel.comment)|[contentType](/javascript/api/excel/excel.comment#contenttype)|获取注释的内容类型。|
|[CommentAddedEventArgs](/javascript/api/excel/excel.commentaddedeventargs)|[commentDetails](/javascript/api/excel/excel.commentaddedeventargs#commentdetails)|获取包含其相关答复的注释 Id 和 Id 的 CommentDetail 数组。|
||[source](/javascript/api/excel/excel.commentaddedeventargs#source)|指定时间源。 有关详细信息，请参阅 Excel.EventSource。|
||[type](/javascript/api/excel/excel.commentaddedeventargs#type)|获取事件的类型。 有关详细信息，请参阅 Excel.EventType。|
||[worksheetId](/javascript/api/excel/excel.commentaddedeventargs#worksheetid)|获取发生事件的工作表的 Id。|
|[CommentChangedEventArgs](/javascript/api/excel/excel.commentchangedeventargs)|[changeType](/javascript/api/excel/excel.commentchangedeventargs#changetype)|获取表示已更改事件的触发方式的更改类型。|
||[commentDetails](/javascript/api/excel/excel.commentchangedeventargs#commentdetails)|获取包含其相关答复的注释 Id 和 Id 的 CommentDetail 数组。|
||[source](/javascript/api/excel/excel.commentchangedeventargs#source)|指定时间源。 有关详细信息，请参阅 Excel.EventSource。|
||[type](/javascript/api/excel/excel.commentchangedeventargs#type)|获取事件的类型。 有关详细信息，请参阅 Excel.EventType。|
||[worksheetId](/javascript/api/excel/excel.commentchangedeventargs#worksheetid)|获取发生事件的工作表的 Id。|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[onAdded](/javascript/api/excel/excel.commentcollection#onadded)|添加注释时发生。|
||[onChanged](/javascript/api/excel/excel.commentcollection#onchanged)|当注释集合中的批注或答复发生更改时发生，包括答复被删除的时间。|
||[onDeleted](/javascript/api/excel/excel.commentcollection#ondeleted)|在注释集合中删除批注时发生。|
|[CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs)|[commentDetails](/javascript/api/excel/excel.commentdeletedeventargs#commentdetails)|获取包含其相关答复的注释 Id 和 Id 的 CommentDetail 数组。|
||[source](/javascript/api/excel/excel.commentdeletedeventargs#source)|指定时间源。 有关详细信息，请参阅 Excel.EventSource。|
||[type](/javascript/api/excel/excel.commentdeletedeventargs#type)|获取事件的类型。 有关详细信息，请参阅 Excel.EventType。|
||[worksheetId](/javascript/api/excel/excel.commentdeletedeventargs#worksheetid)|获取发生事件的工作表的 Id。|
|[CommentDetail](/javascript/api/excel/excel.commentdetail)|[commentId](/javascript/api/excel/excel.commentdetail#commentid)|表示注释的 id。|
||[replyIds](/javascript/api/excel/excel.commentdetail#replyids)|表示相关答复的 id 属于注释。|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[contentType](/javascript/api/excel/excel.commentreply#contenttype)|答复的内容类型。|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[datetimeFormat](/javascript/api/excel/excel.cultureinfo#datetimeformat)|定义适当的区域性格式，以显示日期和时间。 这取决于当前的系统区域性设置。|
|[格式](/javascript/api/excel/excel.datetimeformatinfo)|[dateSeparator](/javascript/api/excel/excel.datetimeformatinfo#dateseparator)|获取用作日期分隔符的字符串。 这取决于当前的系统设置。|
||[longDatePattern](/javascript/api/excel/excel.datetimeformatinfo#longdatepattern)|获取长日期值的格式字符串。 这取决于当前的系统设置。|
||[longTimePattern](/javascript/api/excel/excel.datetimeformatinfo#longtimepattern)|获取长时间值的格式字符串。 这取决于当前的系统设置。|
||[shortDatePattern](/javascript/api/excel/excel.datetimeformatinfo#shortdatepattern)|获取短日期值的格式字符串。 这取决于当前的系统设置。|
||[timeSeparator](/javascript/api/excel/excel.datetimeformatinfo#timeseparator)|获取用作时间分隔符的字符串。 这取决于当前的系统设置。|
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
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[pivotStyle](/javascript/api/excel/excel.pivotlayout#pivotstyle)|应用于数据透视表的样式。|
|[PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter)|[selectedItems](/javascript/api/excel/excel.pivotmanualfilter#selecteditems)|要手动筛选的选定项的列表。 这些项必须是所选字段中现有和有效的项。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[allowMultipleFiltersPerField](/javascript/api/excel/excel.pivottable#allowmultiplefiltersperfield)|指定数据透视表是否允许对表中给定的透视字段上的多个 PivotFilters 进行应用。|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getCount()](/javascript/api/excel/excel.pivottablescopedcollection#getcount--)|获取集合中的数据透视表的数目。|
||[getFirst()](/javascript/api/excel/excel.pivottablescopedcollection#getfirst--)|获取集合中的第一个数据透视表。 集合中的数据透视表按从上到下、从左到右的顺序排序，因此左上角的表格是集合中的第一个数据透视表。|
||[getItem(key: string)](/javascript/api/excel/excel.pivottablescopedcollection#getitem-key-)|按名称获取 PivotTable 对象。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablescopedcollection#getitemornullobject-name-)|按 PivotTable 对象的名称获取此对象。 如果没有 PivotTable 对象，将返回 NULL 对象。|
||[items](/javascript/api/excel/excel.pivottablescopedcollection#items)|获取此集合中已加载的子项。|
|[PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter)|[运算符](/javascript/api/excel/excel.pivotvaluefilter#comparator)|比较运算符是其他值要与其进行比较的静态值。 比较的类型由条件定义。|
||[表达式](/javascript/api/excel/excel.pivotvaluefilter#condition)|指定筛选器的条件，该条件定义了必要的筛选条件。|
||[异](/javascript/api/excel/excel.pivotvaluefilter#exclusive)|如果为 true，则筛选 *排除* 满足条件的项目。 默认值为 false (筛选器以包含满足条件) 的项目。|
||[lowerBound](/javascript/api/excel/excel.pivotvaluefilter#lowerbound)|筛选条件范围的下限 `Between` 。|
||[selectionType](/javascript/api/excel/excel.pivotvaluefilter#selectiontype)|指定筛选器是用于顶部/底部 N 项、顶部/底部 N 百分比还是顶部/底部 N 求和。|
||[极限](/javascript/api/excel/excel.pivotvaluefilter#threshold)|要针对顶部/底部筛选条件筛选的项、百分比或 sum 的 "N" 阈值数。|
||[upperBound](/javascript/api/excel/excel.pivotvaluefilter#upperbound)|筛选条件范围的上限 `Between` 。|
||[value](/javascript/api/excel/excel.pivotvaluefilter#value)|筛选所依据的字段中所选的 "值" 的名称。|
|[区域](/javascript/api/excel/excel.range)|[getDirectPrecedents ( # B1 ](/javascript/api/excel/excel.range#getdirectprecedents--)|返回一个 WorkbookRangeAreas 对象，该对象代表包含同一工作表或多个工作表中的单元格的所有直接引用单元格的区域。|
||[getPivotTables (fullyContained？： boolean) ](/javascript/api/excel/excel.range#getpivottables-fullycontained-)|获取与区域重叠的数据透视表的限定集合。|
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
|[Slicer](/javascript/api/excel/excel.slicer)|[slicerStyle](/javascript/api/excel/excel.slicer#slicerstyle)|应用于切片器的样式。|
|[WorkbookRangeAreas](/javascript/api/excel/excel.workbookrangeareas)|[getRangeAreasBySheet (项： string) ](/javascript/api/excel/excel.workbookrangeareas#getrangeareasbysheet-key-)|`RangeAreas`基于集合中的工作表 id 或名称返回对象。|
||[getRangeAreasOrNullObjectBySheet (项： string) ](/javascript/api/excel/excel.workbookrangeareas#getrangeareasornullobjectbysheet-key-)|`RangeAreas`基于集合中的工作表名称或 id 返回对象。 如果没有 Worksheet 对象，将返回 NULL 对象。|
||[地址](/javascript/api/excel/excel.workbookrangeareas#addresses)|返回 A1 样式的地址数组。 Address 值将包含单元格每个矩形块的工作表名称 (例如，"Sheet1！A1： B4、Sheet1！D1： D4 ") 。 只读。|
||[areas](/javascript/api/excel/excel.workbookrangeareas#areas)|返回 `RangeAreasCollection` 对象。 `RangeAreas`集合中的每个对象代表一张工作表中的一个或多个矩形区域。|
||[区域](/javascript/api/excel/excel.workbookrangeareas#ranges)|返回在对象中组成此对象的范围 `RangeCollection` 。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[customProperties](/javascript/api/excel/excel.worksheet#customproperties)|获取工作表级自定义属性的集合。|
|[WorksheetCustomProperty](/javascript/api/excel/excel.worksheetcustomproperty)|[delete()](/javascript/api/excel/excel.worksheetcustomproperty#delete--)|删除 custom property 对象。|
||[key](/javascript/api/excel/excel.worksheetcustomproperty#key)|获取 customProperty 的键。 自定义属性键不区分大小写。 密钥限制为255个字符 (较大的值将导致引发 "InvalidArgument" 错误。 ) |
||[value](/javascript/api/excel/excel.worksheetcustomproperty#value)|获取或设置自定义属性的值。|
|[WorksheetCustomPropertyCollection](/javascript/api/excel/excel.worksheetcustompropertycollection)|[add (key： string，value： string) ](/javascript/api/excel/excel.worksheetcustompropertycollection#add-key--value-)|添加映射到所提供的键的新自定义属性。 这将使用该密钥覆盖现有的自定义属性。|
||[getCount()](/javascript/api/excel/excel.worksheetcustompropertycollection#getcount--)|获取此工作表上的自定义属性的数目。|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitem-key-)|按键获取自定义属性对象（不区分大小写）。 如果自定义属性不存在，则引发此异常。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitemornullobject-key-)|按键获取自定义属性对象（不区分大小写）。 如果自定义属性不存在，则返回 null 对象。|
||[items](/javascript/api/excel/excel.worksheetcustompropertycollection#items)|获取此集合中已加载的子项。|

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 参考文档](/javascript/api/excel?view=excel-js-1.12&preserve-view=true)
- [Excel JavaScript API 要求集](excel-api-requirement-sets.md)
