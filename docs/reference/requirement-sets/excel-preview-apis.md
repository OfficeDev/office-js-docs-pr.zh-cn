---
title: Excel JavaScript 预览 API
description: 有关即将推出的 Excel JavaScript Api 的详细信息
ms.date: 05/11/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: e3652e60f6ecea9c8269d0ed6b348848c026fe5c
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611370"
---
# <a name="excel-javascript-preview-apis"></a>Excel JavaScript 预览 API

新的 Excel JavaScript API 首先在“预览版”中引入，在进行充分测试并获得用户反馈后，它将成为编号的特定要求集的一部分。

第一个表提供了 API 的简明摘要，而后续表提供了详细列表。

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

| 功能区域 | 说明 | 相关对象 |
|:--- |:--- |:--- |
| 日期和时间[区域性设置](../../excel/excel-add-ins-workbooks.md#access-application-culture-settings) | 提供有关日期和时间格式设置的其他区域性设置的访问权限。 | [CultureInfo](/javascript/api/excel/excel.cultureinfo)、 [NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo) [应用程序](/javascript/api/excel/excel.application) |
| [插入工作簿](../../excel/excel-add-ins-workbooks.md#insert-a-copy-of-an-existing-workbook-into-the-current-one-preview) | 将一个工作簿插入另一个工作簿。  | [Workbook](/javascript/api/excel/excel.worksheetcollection) |
| 透视筛选器 | 对数据透视表的字段应用数值驱动的筛选器。 | [透视字段](/javascript/api/excel/excel.pivotfield#applyfilter-filter-)、 [PivotFilters](/javascript/api/excel/excel.pivotFilters) |
|区域 spilling | 允许外接程序查找与[动态数组](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531)结果关联的范围。 | [Range](/javascript/api/excel/excel.range) |

## <a name="api-list"></a>API 列表

下表列出了当前预览中的 Excel JavaScript Api。 若要查看所有 Excel JavaScript Api （包括预览 Api 和之前发布的 Api）的完整列表，请参阅[所有 Excel Javascript api](/javascript/api/excel?view=excel-js-preview)。

| Class | 域 | 说明 |
|:---|:---|:---|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[getDimensionValues （维： ChartSeriesDimension）](/javascript/api/excel/excel.chartseries#getdimensionvalues-dimension-)|从图表系列的单个维中获取值。 这些值可以是类别值，也可以是数据值，具体取决于指定的维度和为图表系列映射数据的方式。|
|[Comment](/javascript/api/excel/excel.comment)|[contentType](/javascript/api/excel/excel.comment#contenttype)|获取注释的内容类型。|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[contentType](/javascript/api/excel/excel.commentreply#contenttype)|答复的内容类型。|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[datetimeFormat](/javascript/api/excel/excel.cultureinfo#datetimeformat)|定义适当的区域性格式，以显示日期和时间。 这取决于当前的系统区域性设置。|
|[格式](/javascript/api/excel/excel.datetimeformatinfo)|[dateSeparator](/javascript/api/excel/excel.datetimeformatinfo#dateseparator)|获取用作日期分隔符的字符串。 这取决于当前的系统设置。|
||[longDatePattern](/javascript/api/excel/excel.datetimeformatinfo#longdatepattern)|获取长日期值的格式字符串。 这取决于当前的系统设置。|
||[longTimePattern](/javascript/api/excel/excel.datetimeformatinfo#longtimepattern)|获取长时间值的格式字符串。 这取决于当前的系统设置。|
||[shortDatePattern](/javascript/api/excel/excel.datetimeformatinfo#shortdatepattern)|获取短日期值的格式字符串。 这取决于当前的系统设置。|
||[timeSeparator](/javascript/api/excel/excel.datetimeformatinfo#timeseparator)|获取用作时间分隔符的字符串。 这取决于当前的系统设置。|
|[PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter)|[运算符](/javascript/api/excel/excel.pivotdatefilter#comparator)|比较运算符是其他值要与其进行比较的静态值。 比较的类型由条件定义。|
||[表达式](/javascript/api/excel/excel.pivotdatefilter#condition)|指定筛选器的条件，该条件定义了必要的筛选条件。|
||[异](/javascript/api/excel/excel.pivotdatefilter#exclusive)|如果为 true，则筛选*排除*满足条件的项目。 默认值为 false （筛选以包含满足条件的项目）。|
||[lowerBound](/javascript/api/excel/excel.pivotdatefilter#lowerbound)|筛选条件范围的下限 `Between` 。|
||[upperBound](/javascript/api/excel/excel.pivotdatefilter#upperbound)|筛选条件范围的上限 `Between` 。|
||[wholeDays](/javascript/api/excel/excel.pivotdatefilter#wholedays)|对于 `Equals` 、 `Before` 、 `After` 和 `Between` 筛选条件，指示是否应将比较作为全天进行。|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[applyFilter （filter： PivotFilters）](/javascript/api/excel/excel.pivotfield#applyfilter-filter-)|设置字段的当前 PivotFilters 的一个或多个，并将它们应用于字段。|
||[clearAllFilters （）](/javascript/api/excel/excel.pivotfield#clearallfilters--)|从字段的所有筛选器中清除所有条件。 这将删除对字段的任何活动筛选。|
||[clearFilter （filterType： PivotFilterType）](/javascript/api/excel/excel.pivotfield#clearfilter-filtertype-)|清除给定类型的字段筛选器中的所有现有条件（如果当前应用了一个条件）。|
||[getFilters()](/javascript/api/excel/excel.pivotfield#getfilters--)|获取当前应用于字段的所有筛选器。|
||[isFiltered （filterType？： PivotFilterType）](/javascript/api/excel/excel.pivotfield#isfiltered-filtertype-)|检查字段上是否有任何已应用的筛选器。|
|[PivotFilters](/javascript/api/excel/excel.pivotfilters)|[dateFilter](/javascript/api/excel/excel.pivotfilters#datefilter)|透视字段当前应用的日期筛选器。 如果未应用任何值，则为 Null。|
||[labelFilter](/javascript/api/excel/excel.pivotfilters#labelfilter)|透视字段当前应用的标签筛选器。 如果未应用任何值，则为 Null。|
||[manualFilter](/javascript/api/excel/excel.pivotfilters#manualfilter)|透视字段当前应用的手动筛选。 如果未应用任何值，则为 Null。|
||[valueFilter](/javascript/api/excel/excel.pivotfilters#valuefilter)|透视字段当前应用的值筛选器。 如果未应用任何值，则为 Null。|
|[PivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter)|[运算符](/javascript/api/excel/excel.pivotlabelfilter#comparator)|比较运算符是其他值要与其进行比较的静态值。 比较的类型由条件定义。|
||[表达式](/javascript/api/excel/excel.pivotlabelfilter#condition)|指定筛选器的条件，该条件定义了必要的筛选条件。|
||[异](/javascript/api/excel/excel.pivotlabelfilter#exclusive)|如果为 true，则筛选*排除*满足条件的项目。 默认值为 false （筛选以包含满足条件的项目）。|
||[lowerBound](/javascript/api/excel/excel.pivotlabelfilter#lowerbound)|筛选条件之间的范围的下限。|
||[substring](/javascript/api/excel/excel.pivotlabelfilter#substring)|用于 `BeginsWith` 、 `EndsWith` 和筛选条件的子字符串 `Contains` 。|
||[upperBound](/javascript/api/excel/excel.pivotlabelfilter#upperbound)|筛选条件之间的范围的上限。|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|根据数据层次结构以及各自层次结构的行和列项，获取数据透视表中的唯一单元格。 返回的单元格是给定行和列的交集，其中包含来自给定层次结构的数据。 此方法与在特定单元格上调用 getPivotItems 和 getDataHierarchy 相反。|
||[pivotStyle](/javascript/api/excel/excel.pivotlayout#pivotstyle)|应用于数据透视表的样式。|
||[setStyle （style： string \| PivotTableStyle \| BuiltInPivotTableStyle）](/javascript/api/excel/excel.pivotlayout#setstyle-style-)|设置应用于数据透视表的样式。|
|[PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter)|[selectedItems](/javascript/api/excel/excel.pivotmanualfilter#selecteditems)|要手动筛选的选定项的列表。 这些项必须是所选字段中现有和有效的项。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[allowMultipleFiltersPerField](/javascript/api/excel/excel.pivottable#allowmultiplefiltersperfield)|指定数据透视表是否允许对表中给定的透视字段上的多个 PivotFilters 进行应用。|
|[PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter)|[运算符](/javascript/api/excel/excel.pivotvaluefilter#comparator)|比较运算符是其他值要与其进行比较的静态值。 比较的类型由条件定义。|
||[表达式](/javascript/api/excel/excel.pivotvaluefilter#condition)|指定筛选器的条件，该条件定义了必要的筛选条件。|
||[异](/javascript/api/excel/excel.pivotvaluefilter#exclusive)|如果为 true，则筛选*排除*满足条件的项目。 默认值为 false （筛选以包含满足条件的项目）。|
||[lowerBound](/javascript/api/excel/excel.pivotvaluefilter#lowerbound)|筛选条件范围的下限 `Between` 。|
||[selectionType](/javascript/api/excel/excel.pivotvaluefilter#selectiontype)|指定筛选器是用于顶部/底部 N 项、顶部/底部 N 百分比还是顶部/底部 N 求和。|
||[极限](/javascript/api/excel/excel.pivotvaluefilter#threshold)|要针对顶部/底部筛选条件筛选的项、百分比或 sum 的 "N" 阈值数。|
||[upperBound](/javascript/api/excel/excel.pivotvaluefilter#upperbound)|筛选条件范围的上限 `Between` 。|
||[value](/javascript/api/excel/excel.pivotvaluefilter#value)|筛选所依据的字段中所选的 "值" 的名称。|
|[Range](/javascript/api/excel/excel.range)|[getSpillParent()](/javascript/api/excel/excel.range#getspillparent--)|获取 Range 对象，它包含要将某个单元格溢出到的定位单元格。 如果应用于具有多个单元格的区域，则会失败。|
||[getSpillParentOrNullObject()](/javascript/api/excel/excel.range#getspillparentornullobject--)|获取 Range 对象，它包含要将某个单元格溢出到的定位单元格。|
||[getSpillingToRange()](/javascript/api/excel/excel.range#getspillingtorange--)|获取 Range 对象，它在调用定位单元格时包含溢出区域。 如果应用于具有多个单元格的区域，则会失败。|
||[getSpillingToRangeOrNullObject()](/javascript/api/excel/excel.range#getspillingtorangeornullobject--)|获取 Range 对象，它在调用定位单元格时包含溢出区域。|
||[hasSpill](/javascript/api/excel/excel.range#hasspill)|表示所有单元格是否都具有溢出边框。|
||[numberFormatCategories](/javascript/api/excel/excel.range#numberformatcategories)|表示每个单元格的数字格式的类别。|
||[savedAsArray](/javascript/api/excel/excel.range#savedasarray)|表示是否将所有单元格都保存为数组公式。|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|从 XML 字符串创建可缩放的矢量图形 (SVG) 并将其添加到工作表。 返回表示新图片的 Shape 对象。|
|[Slicer](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameinformula)|表示公式中使用切片器名称。|
||[slicerStyle](/javascript/api/excel/excel.slicer#slicerstyle)|应用于切片器的样式。|
||[setStyle （style： string \| PivotTableStyle \| BuiltInSlicerStyle）](/javascript/api/excel/excel.slicer#setstyle-style-)|设置应用于切片器的样式。|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|将表格更改为使用默认表格样式。|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|在特定表格上应用筛选器时发生。|
||[tableStyle](/javascript/api/excel/excel.table#tablestyle)|应用于表的样式。|
||[setStyle （style： string \| PivotTableStyle \| BuiltInTableStyle）](/javascript/api/excel/excel.table#setstyle-style-)|设置应用于切片器的样式。|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|在工作簿或工作表中的任何表格上应用筛选器时发生。|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|获取应用了筛选器的表的 id。|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#type)|获取事件的类型。 有关详细信息，请参阅 Excel.EventType。|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|获取包含表的工作表的 id。|
|[Workbook](/javascript/api/excel/excel.workbook)|[showPivotFieldList](/javascript/api/excel/excel.workbook#showpivotfieldlist)|指定是否在工作簿级别显示数据透视表的 "字段列表" 窗格。|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|如果工作簿使用 1904 日期系统，则为 True。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[customProperties](/javascript/api/excel/excel.worksheet#customproperties)|获取工作表级自定义属性的集合。|
||[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|在特定工作表上应用筛选器时发生。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|将工作簿的指定工作表插入当前工作簿。|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|在工作簿中应用任何工作表的筛选器时发生。|
|[WorksheetCustomProperty](/javascript/api/excel/excel.worksheetcustomproperty)|[delete()](/javascript/api/excel/excel.worksheetcustomproperty#delete--)|删除 custom property 对象。|
||[key](/javascript/api/excel/excel.worksheetcustomproperty#key)|获取 customProperty 的键。 自定义属性键不区分大小写。|
||[value](/javascript/api/excel/excel.worksheetcustomproperty#value)|获取或设置自定义属性的值。|
|[WorksheetCustomPropertyCollection](/javascript/api/excel/excel.worksheetcustompropertycollection)|[add （key： string，value： string）](/javascript/api/excel/excel.worksheetcustompropertycollection#add-key--value-)|添加映射到所提供的键的新自定义属性。 这将使用该密钥覆盖现有的自定义属性。|
||[getCount()](/javascript/api/excel/excel.worksheetcustompropertycollection#getcount--)|获取此工作表上的自定义属性的数目。|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitem-key-)|按键获取自定义属性对象（不区分大小写）。 如果自定义属性不存在，则引发此异常。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitemornullobject-key-)|按键获取自定义属性对象（不区分大小写）。 如果自定义属性不存在，则返回 null 对象。|
||[items](/javascript/api/excel/excel.worksheetcustompropertycollection#items)|获取此集合中已加载的子项。|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|获取事件的类型。 有关详细信息，请参阅 Excel.EventType。|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|获取应用了筛选器的工作表的 id。|

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 参考文档](/javascript/api/excel?view=excel-js-preview)
- [Excel JavaScript API 要求集](./excel-api-requirement-sets.md)
