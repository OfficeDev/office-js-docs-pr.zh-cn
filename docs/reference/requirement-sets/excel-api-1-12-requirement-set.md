---
title: ExcelJavaScript API 要求集 1.12
description: 有关 ExcelApi 1.12 要求集的详细信息。
ms.date: 04/01/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: f2780e188eeebc70d754d7c2ea26cde1b1974702
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2021
ms.locfileid: "59152297"
---
# <a name="whats-new-in-excel-javascript-api-112"></a>JavaScript API 1.12 Excel的新增功能

ExcelApi 1.12 通过添加用于跟踪动态数组和查找公式的直接引用单元格的 API 来增加对范围中公式的支持。 它还添加了数据透视表筛选器的 API 控件。 注释、区域性设置和自定义属性功能区域也进行了改进。

| 功能区域 | 说明 | 相关对象 |
|:--- |:--- |:--- |
| [注释事件](../../excel/excel-add-ins-comments.md#comment-events) | 将添加、更改和删除事件添加到注释集合。| [CommentCollection](/javascript/api/excel/excel.commentcollection) |
| 日期和时间 [区域性设置](../../excel/excel-add-ins-workbooks.md#access-application-culture-settings) | 提供对日期和时间格式的其他文化设置的访问权限。 | [CultureInfo](/javascript/api/excel/excel.cultureinfo) [、NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo) [应用程序](/javascript/api/excel/excel.application) |
| [直接引用单元格](../../excel/excel-add-ins-ranges-precedents.md) | 返回用于计算单元格公式的范围。| [Range](/javascript/api/excel/excel.range#getdirectprecedents--) |
| 透视筛选器 | 将值驱动的筛选器应用于数据透视表的字段。 | [PivotField](/javascript/api/excel/excel.pivotfield#applyfilter-filter-) [、PivotFilters](/javascript/api/excel/excel.pivotFilters) |
| [区域溢出](../../excel/excel-add-ins-ranges-dynamic-arrays.md) | 允许外接程序查找与动态数组结果 [关联的](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531) 区域。 | [Range](/javascript/api/excel/excel.range) |
| [工作表级别的自定义属性](../../excel/excel-add-ins-workbooks.md#worksheet-level-custom-properties) | 除了将自定义属性的范围限制到工作簿级别外，还可以将自定义属性的范围缩小到工作表级别。 | [WorksheetCustomProperty](/javascript/api/excel/excel.worksheetcustomproperty) [、WorksheetCustomPropertyCollection](/javascript/api/excel/excel.worksheetcustompropertycollection)|

## <a name="api-list"></a>API 列表

下表列出了 JavaScript API 要求集 1.12 Excel中的 API。 若要查看受 Excel JavaScript API 要求集 1.12 或更早版本支持的所有 API 的 API 参考文档，请参阅要求集[1.12](/javascript/api/excel?view=excel-js-1.12&preserve-view=true)或更早中的 Excel API。

| 类 | 域 | 说明 |
|:---|:---|:---|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[textOrientation](/javascript/api/excel/excel.chartaxistitle#textOrientation)|指定文本面向图表坐标轴标题的角度。|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[getDimensionValues (维度：Excel。ChartSeriesDimension) ](/javascript/api/excel/excel.chartseries#getDimensionValues_dimension_)|获取图表系列的单个维度中的值。|
|[Comment](/javascript/api/excel/excel.comment)|[contentType](/javascript/api/excel/excel.comment#contentType)|获取注释的内容类型。|
|[CommentAddedEventArgs](/javascript/api/excel/excel.commentaddedeventargs)|[commentDetails](/javascript/api/excel/excel.commentaddedeventargs#commentDetails)|获取 `CommentDetail` 包含注释 ID 和相关回复的 ID 的数组。|
||[source](/javascript/api/excel/excel.commentaddedeventargs#source)|指定时间源。|
||[type](/javascript/api/excel/excel.commentaddedeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.commentaddedeventargs#worksheetId)|获取发生事件的工作表的 ID。|
|[CommentChangedEventArgs](/javascript/api/excel/excel.commentchangedeventargs)|[changeType](/javascript/api/excel/excel.commentchangedeventargs#changeType)|获取更改类型，该类型表示如何触发更改事件。|
||[commentDetails](/javascript/api/excel/excel.commentchangedeventargs#commentDetails)|获取包含注释 ID 和相关回复的 ID `CommentDetail` 的数组。|
||[source](/javascript/api/excel/excel.commentchangedeventargs#source)|指定时间源。|
||[type](/javascript/api/excel/excel.commentchangedeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.commentchangedeventargs#worksheetId)|获取发生事件的工作表的 ID。|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[onAdded](/javascript/api/excel/excel.commentcollection#onAdded)|添加注释时发生。|
||[onChanged](/javascript/api/excel/excel.commentcollection#onChanged)|在批注集合中的批注或答复发生更改时发生，包括删除答复时。|
||[onDeleted](/javascript/api/excel/excel.commentcollection#onDeleted)|在批注集合中删除批注时发生。|
|[CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs)|[commentDetails](/javascript/api/excel/excel.commentdeletedeventargs#commentDetails)|获取 `CommentDetail` 包含注释 ID 和相关回复的 ID 的数组。|
||[source](/javascript/api/excel/excel.commentdeletedeventargs#source)|指定时间源。|
||[type](/javascript/api/excel/excel.commentdeletedeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.commentdeletedeventargs#worksheetId)|获取发生事件的工作表的 ID。|
|[CommentDetail](/javascript/api/excel/excel.commentdetail)|[commentId](/javascript/api/excel/excel.commentdetail#commentId)|表示批注的 ID。|
||[replyIds](/javascript/api/excel/excel.commentdetail#replyIds)|表示属于注释的相关回复的 ID。|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[contentType](/javascript/api/excel/excel.commentreply#contentType)|回复的内容类型。|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[datetimeFormat](/javascript/api/excel/excel.cultureinfo#datetimeFormat)|定义显示日期和时间的适合文化的格式。|
|[DatetimeFormatInfo](/javascript/api/excel/excel.datetimeformatinfo)|[dateSeparator](/javascript/api/excel/excel.datetimeformatinfo#dateSeparator)|获取用作日期分隔符的字符串。|
||[longDatePattern](/javascript/api/excel/excel.datetimeformatinfo#longDatePattern)|获取长日期值的格式字符串。|
||[longTimePattern](/javascript/api/excel/excel.datetimeformatinfo#longTimePattern)|获取长时间值的格式字符串。|
||[shortDatePattern](/javascript/api/excel/excel.datetimeformatinfo#shortDatePattern)|获取短日期值的格式字符串。|
||[timeSeparator](/javascript/api/excel/excel.datetimeformatinfo#timeSeparator)|获取用作时间分隔符的字符串。|
|[PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter)|[比较器](/javascript/api/excel/excel.pivotdatefilter#comparator)|比较器是比较其他值的静态值。|
||[condition](/javascript/api/excel/excel.pivotdatefilter#condition)|指定筛选器的条件，该条件定义必要的筛选条件。|
||[exclusive](/javascript/api/excel/excel.pivotdatefilter#exclusive)|如果 `true` 为 ，筛选器 *将* 排除满足条件的项目。|
||[lowerBound](/javascript/api/excel/excel.pivotdatefilter#lowerBound)|筛选条件的范围的 `between` 下限。|
||[upperBound](/javascript/api/excel/excel.pivotdatefilter#upperBound)|筛选条件的范围 `between` 上限。|
||[wholeDays](/javascript/api/excel/excel.pivotdatefilter#wholeDays)|对于 、 、 和 筛选条件， `equals` `before` `after` `between` 指示是否按整日进行比较。|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[applyFilter (筛选器：Excel。PivotFilters) ](/javascript/api/excel/excel.pivotfield#applyFilter_filter_)|设置字段的一个或多个当前 PivotFilter，并应用于该字段。|
||[clearAllFilters () ](/javascript/api/excel/excel.pivotfield#clearAllFilters__)|清除字段的所有筛选器的所有条件。|
||[clearFilter (filterType： Excel。PivotFilterType) ](/javascript/api/excel/excel.pivotfield#clearFilter_filterType_)|从给定类型的字段筛选器中清除所有现有条件 (如果当前已对给定) 。|
||[getFilters () ](/javascript/api/excel/excel.pivotfield#getFilters__)|获取当前应用于字段的所有筛选器。|
||[isFiltered (filterType？： Excel。PivotFilterType) ](/javascript/api/excel/excel.pivotfield#isFiltered_filterType_)|检查字段上是否有已应用的筛选器。|
|[PivotFilters](/javascript/api/excel/excel.pivotfilters)|[dateFilter](/javascript/api/excel/excel.pivotfilters#dateFilter)|透视字段当前应用的日期筛选器。|
||[labelFilter](/javascript/api/excel/excel.pivotfilters#labelFilter)|透视字段当前应用的标签筛选器。|
||[manualFilter](/javascript/api/excel/excel.pivotfilters#manualFilter)|透视字段当前应用的手动筛选器。|
||[valueFilter](/javascript/api/excel/excel.pivotfilters#valueFilter)|透视字段当前应用的值筛选器。|
|[PivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter)|[比较器](/javascript/api/excel/excel.pivotlabelfilter#comparator)|比较器是比较其他值的静态值。|
||[condition](/javascript/api/excel/excel.pivotlabelfilter#condition)|指定筛选器的条件，该条件定义必要的筛选条件。|
||[exclusive](/javascript/api/excel/excel.pivotlabelfilter#exclusive)|如果 `true` 为 ，筛选器 *将* 排除满足条件的项目。|
||[lowerBound](/javascript/api/excel/excel.pivotlabelfilter#lowerBound)|筛选条件的范围的 `between` 下限。|
||[substring](/javascript/api/excel/excel.pivotlabelfilter#substring)|用于 、 和 `beginsWith` `endsWith` 筛选条件的 `contains` 子字符串。|
||[upperBound](/javascript/api/excel/excel.pivotlabelfilter#upperBound)|筛选条件的范围 `between` 上限。|
|[PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter)|[selectedItems](/javascript/api/excel/excel.pivotmanualfilter#selectedItems)|要手动筛选的选定项的列表。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[allowMultipleFiltersPerField](/javascript/api/excel/excel.pivottable#allowMultipleFiltersPerField)|指定数据透视表是否允许在表中的给定透视字段上应用多个 PivotFilter。|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getCount()](/javascript/api/excel/excel.pivottablescopedcollection#getCount__)|获取集合中数据透视表的数量。|
||[getFirst()](/javascript/api/excel/excel.pivottablescopedcollection#getFirst__)|获取集合中的第一个数据透视表。|
||[getItem(key: string)](/javascript/api/excel/excel.pivottablescopedcollection#getItem_key_)|按名称获取 PivotTable 对象。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablescopedcollection#getItemOrNullObject_name_)|按名称获取 PivotTable 对象。|
||[items](/javascript/api/excel/excel.pivottablescopedcollection#items)|获取此集合中已加载的子项。|
|[PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter)|[比较器](/javascript/api/excel/excel.pivotvaluefilter#comparator)|比较器是比较其他值的静态值。|
||[condition](/javascript/api/excel/excel.pivotvaluefilter#condition)|指定筛选器的条件，该条件定义必要的筛选条件。|
||[exclusive](/javascript/api/excel/excel.pivotvaluefilter#exclusive)|如果 `true` 为 ，筛选器 *将* 排除满足条件的项目。|
||[lowerBound](/javascript/api/excel/excel.pivotvaluefilter#lowerBound)|筛选条件的范围的 `between` 下限。|
||[selectionType](/javascript/api/excel/excel.pivotvaluefilter#selectionType)|指定筛选器是针对上/下 N 项、上/下 N% 还是上/下 N 个和。|
||[阈值](/javascript/api/excel/excel.pivotvaluefilter#threshold)|要针对顶部/底部筛选条件进行筛选的项目数、百分比或总和的"N"阈值。|
||[upperBound](/javascript/api/excel/excel.pivotvaluefilter#upperBound)|筛选条件的范围 `between` 上限。|
||[value](/javascript/api/excel/excel.pivotvaluefilter#value)|要筛选的字段中所选"值"的名称。|
|[Range](/javascript/api/excel/excel.range)|[getDirectPrecedents () ](/javascript/api/excel/excel.range#getDirectPrecedents__)|返回一个对象，该对象表示包含同一工作表或多个工作表中单元格的所有直接引用 `WorkbookRangeAreas` 单元格的范围。|
||[getPivotTables (fullyContained？： boolean) ](/javascript/api/excel/excel.range#getPivotTables_fullyContained_)|获取与区域重叠的数据透视表的范围集合。|
||[getSpillParent()](/javascript/api/excel/excel.range#getSpillParent__)|获取 Range 对象，它包含要将某个单元格溢出到的定位单元格。|
||[getSpillParentOrNullObject()](/javascript/api/excel/excel.range#getSpillParentOrNullObject__)|获取包含被溢出到的单元格的定位单元格的 range 对象。|
||[getSpillingToRange()](/javascript/api/excel/excel.range#getSpillingToRange__)|获取 Range 对象，它在调用定位单元格时包含溢出区域。|
||[getSpillingToRangeOrNullObject()](/javascript/api/excel/excel.range#getSpillingToRangeOrNullObject__)|获取 Range 对象，它在调用定位单元格时包含溢出区域。|
||[hasSpill](/javascript/api/excel/excel.range#hasSpill)|表示所有单元格是否都具有溢出边框。|
||[numberFormatCategories](/javascript/api/excel/excel.range#numberFormatCategories)|表示每个单元格的编号格式类别。|
||[savedAsArray](/javascript/api/excel/excel.range#savedAsArray)|表示所有单元格是否将另存为数组公式。|
|[RangeAreasCollection](/javascript/api/excel/excel.rangeareascollection)|[getCount()](/javascript/api/excel/excel.rangeareascollection#getCount__)|获取此集合 `RangeAreas` 中的对象数。|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangeareascollection#getItemAt_index_)|基于 `RangeAreas` 集合中的位置返回对象。|
||[items](/javascript/api/excel/excel.rangeareascollection#items)|获取此集合中已加载的子项。|
|[WorkbookRangeAreas](/javascript/api/excel/excel.workbookrangeareas)|[getRangeAreasBySheet (键：string) ](/javascript/api/excel/excel.workbookrangeareas#getRangeAreasBySheet_key_)|基于 `RangeAreas` 集合中的工作表 ID 或名称返回对象。|
||[getRangeAreasOrNullObjectBySheet (键：string) ](/javascript/api/excel/excel.workbookrangeareas#getRangeAreasOrNullObjectBySheet_key_)|基于 `RangeAreas` 集合中的工作表名称或 ID 返回对象。|
||[地址](/javascript/api/excel/excel.workbookrangeareas#addresses)|返回 A1 样式的地址数组。|
||[areas](/javascript/api/excel/excel.workbookrangeareas#areas)|返回 `RangeAreasCollection` 对象。|
||[ranges](/javascript/api/excel/excel.workbookrangeareas#ranges)|返回在对象中组成此对象 `RangeCollection` 的范围。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[customProperties](/javascript/api/excel/excel.worksheet#customProperties)|获取工作表级自定义属性的集合。|
|[WorksheetCustomProperty](/javascript/api/excel/excel.worksheetcustomproperty)|[delete()](/javascript/api/excel/excel.worksheetcustomproperty#delete__)|删除 custom property 对象。|
||[key](/javascript/api/excel/excel.worksheetcustomproperty#key)|获取 customProperty 的键。|
||[value](/javascript/api/excel/excel.worksheetcustomproperty#value)|获取或设置自定义属性的值。|
|[WorksheetCustomPropertyCollection](/javascript/api/excel/excel.worksheetcustompropertycollection)|[add (key： string， value： string) ](/javascript/api/excel/excel.worksheetcustompropertycollection#add_key__value_)|添加映射到提供的键的新自定义属性。|
||[getCount()](/javascript/api/excel/excel.worksheetcustompropertycollection#getCount__)|获取此工作表上的自定义属性数。|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getItem_key_)|按键获取自定义属性对象（不区分大小写）。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getItemOrNullObject_key_)|按键获取自定义属性对象（不区分大小写）。|
||[items](/javascript/api/excel/excel.worksheetcustompropertycollection#items)|获取此集合中已加载的子项。|

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 参考文档](/javascript/api/excel?view=excel-js-1.12&preserve-view=true)
- [Excel JavaScript API 要求集](excel-api-requirement-sets.md)
