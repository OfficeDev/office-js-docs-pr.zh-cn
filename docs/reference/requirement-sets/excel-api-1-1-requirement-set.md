---
title: ExcelJavaScript API 要求集 1.1
description: 有关 ExcelApi 1.1 要求集的详细信息。
ms.date: 11/09/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 7bc378c200d8aa7c200158d7fe50fdbd71b8251a
ms.sourcegitcommit: 3fa8c754a47bab909e559ae3e5d4237ba27fdbe4
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/30/2021
ms.locfileid: "53671441"
---
# <a name="excel-javascript-api-requirement-set-11"></a>ExcelJavaScript API 要求集 1.1

Excel JavaScript API 1.1 是首版 API。 这是唯一Excel支持的特定要求Excel 2016。

## <a name="api-list"></a>API 列表

下表列出了 JavaScript API 要求Excel集 1.1 中的 API。 若要查看受 Excel JavaScript API 要求集 1.1 支持的所有 API 的 API 参考文档，请参阅要求集[1.1](/javascript/api/excel?view=excel-js-1.1&preserve-view=true)中的 Excel API。

| 类 | 域 | 说明 |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[calculate (calculationType：Excel。CalculationType) ](/javascript/api/excel/excel.application#calculate_calculationType_)|重新计算 Excel 中当前打开的所有工作簿。|
||[calculationMode](/javascript/api/excel/excel.application#calculationMode)|返回工作簿中使用的计算模式，如 中的常量所定义 `Excel.CalculationMode` 。|
|[Binding](/javascript/api/excel/excel.binding)|[getRange()](/javascript/api/excel/excel.binding#getRange__)|返回绑定表示的区域。|
||[getTable()](/javascript/api/excel/excel.binding#getTable__)|返回绑定表示的表。|
||[getText()](/javascript/api/excel/excel.binding#getText__)|返回绑定表示的文本。|
||[id](/javascript/api/excel/excel.binding#id)|表示绑定标识符。|
||[type](/javascript/api/excel/excel.binding#type)|返回绑定的类型。|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[getItem(id: string)](/javascript/api/excel/excel.bindingcollection#getItem_id_)|按 ID 获取绑定对象。|
||[getItemAt(index: number)](/javascript/api/excel/excel.bindingcollection#getItemAt_index_)|根据其在项目数组中的位置获取绑定对象。|
||[count](/javascript/api/excel/excel.bindingcollection#count)|返回集合中绑定的数量。|
||[items](/javascript/api/excel/excel.bindingcollection#items)|获取此集合中已加载的子项。|
|[Chart](/javascript/api/excel/excel.chart)|[delete()](/javascript/api/excel/excel.chart#delete__)|删除 chart 对象。|
||[height](/javascript/api/excel/excel.chart#height)|指定图表对象的高度（以点表示）。|
||[left](/javascript/api/excel/excel.chart#left)|从图表左侧到工作表原点的距离，以磅为单位。|
||[名称](/javascript/api/excel/excel.chart#name)|指定图表对象的名称。|
||[axes](/javascript/api/excel/excel.chart#axes)|表示图表坐标轴。|
||[dataLabels](/javascript/api/excel/excel.chart#dataLabels)|表示图表上的数据标签。|
||[format](/javascript/api/excel/excel.chart#format)|封装图表区域的格式属性。|
||[legend](/javascript/api/excel/excel.chart#legend)|表示图表的图例。|
||[series](/javascript/api/excel/excel.chart#series)|表示单个系列或图表中的系列集合。|
||[title](/javascript/api/excel/excel.chart#title)|表示指定图表的标题，包括标题的文本、可见性、位置和格式。|
||[setData (sourceData： Range， seriesBy？： Excel。ChartSeriesBy) ](/javascript/api/excel/excel.chart#setData_sourceData__seriesBy_)|重置图表的源数据。|
||[setPosition (startCell： Range \| string， endCell？： Range \| string) ](/javascript/api/excel/excel.chart#setPosition_startCell__endCell_)|相对于工作表上的单元格放置图表。|
||[top](/javascript/api/excel/excel.chart#top)|指定从对象上边缘到工作表上第 1 (行顶端的距离（以) 或图表 (图表区域的顶部) ）。|
||[width](/javascript/api/excel/excel.chart#width)|指定图表对象的宽度（以点表示）。|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[fill](/javascript/api/excel/excel.chartareaformat#fill)|表示对象的填充格式，包括背景格式信息。|
||[font](/javascript/api/excel/excel.chartareaformat#font)|表示当前对象的字体属性（字体名称、字体大小、颜色等）。|
|[ChartAxes](/javascript/api/excel/excel.chartaxes)|[categoryAxis](/javascript/api/excel/excel.chartaxes#categoryAxis)|表示图表中的类别轴。|
||[seriesAxis](/javascript/api/excel/excel.chartaxes#seriesAxis)|表示三维图表的系列轴。|
||[valueAxis](/javascript/api/excel/excel.chartaxes#valueAxis)|表示坐标轴中的数值轴。|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[majorUnit](/javascript/api/excel/excel.chartaxis#majorUnit)|表示两个主要刻度标记之间的间隔。|
||[maximum](/javascript/api/excel/excel.chartaxis#maximum)|表示数值轴上的最大值。|
||[minimum](/javascript/api/excel/excel.chartaxis#minimum)|表示数值轴上的最小值。|
||[minorUnit](/javascript/api/excel/excel.chartaxis#minorUnit)|表示两个次要刻度标记之间的间隔。|
||[format](/javascript/api/excel/excel.chartaxis#format)|表示 chart 对象的格式，包括线条和字体格式。|
||[majorGridlines](/javascript/api/excel/excel.chartaxis#majorGridlines)|返回一个对象，该对象代表指定坐标轴的主要网格线。|
||[minorGridlines](/javascript/api/excel/excel.chartaxis#minorGridlines)|返回一个对象，该对象代表指定坐标轴的次要网格线。|
||[title](/javascript/api/excel/excel.chartaxis#title)|表示坐标轴标题。|
|[ChartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|[font](/javascript/api/excel/excel.chartaxisformat#font)|指定图表坐标轴 (的字体属性) 字体名称、字体大小、颜色等。|
||[line](/javascript/api/excel/excel.chartaxisformat#line)|指定图表线条格式。|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[format](/javascript/api/excel/excel.chartaxistitle#format)|指定图表坐标轴标题的格式。|
||[text](/javascript/api/excel/excel.chartaxistitle#text)|指定坐标轴标题。|
||[visible](/javascript/api/excel/excel.chartaxistitle#visible)|指定坐标轴标题是否可见。|
|[ChartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|[font](/javascript/api/excel/excel.chartaxistitleformat#font)|指定图表坐标轴标题的字体属性，如图表坐标轴标题对象的字体名称、字号或颜色。|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[添加 (类型：Excel。ChartType， sourceData： Range， seriesBy？： Excel。ChartSeriesBy) ](/javascript/api/excel/excel.chartcollection#add_type__sourceData__seriesBy_)|创建新图表。|
||[getItem(name: string)](/javascript/api/excel/excel.chartcollection#getItem_name_)|使用图表名称获取图表。|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartcollection#getItemAt_index_)|根据其在集合中的位置获取图表。|
||[count](/javascript/api/excel/excel.chartcollection#count)|返回工作表中的图表数。|
||[items](/javascript/api/excel/excel.chartcollection#items)|获取此集合中已加载的子项。|
|[ChartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|[fill](/javascript/api/excel/excel.chartdatalabelformat#fill)|表示当前图表数据标签的填充格式。|
||[font](/javascript/api/excel/excel.chartdatalabelformat#font)|表示图表数据标签 (字体名称、字号和颜色) 字体属性。|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[position](/javascript/api/excel/excel.chartdatalabels#position)|表示数据标签的位置的值。|
||[format](/javascript/api/excel/excel.chartdatalabels#format)|指定图表数据标签的格式，包括填充和字体格式。|
||[分隔符](/javascript/api/excel/excel.chartdatalabels#separator)|表示用于图表中数据标签的分隔符的字符串。|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabels#showBubbleSize)|指定数据标签气泡大小是否可见。|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabels#showCategoryName)|指定数据标签类别名称是否可见。|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabels#showLegendKey)|指定数据标签图例项标示是否可见。|
||[showPercentage](/javascript/api/excel/excel.chartdatalabels#showPercentage)|指定数据标签百分比是否可见。|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabels#showSeriesName)|指定数据标签系列名称是否可见。|
||[showValue](/javascript/api/excel/excel.chartdatalabels#showValue)|指定数据标签值是否可见。|
|[ChartFill](/javascript/api/excel/excel.chartfill)|[clear()](/javascript/api/excel/excel.chartfill#clear__)|清除图表元素的填充颜色。|
||[setSolidColor(color: string)](/javascript/api/excel/excel.chartfill#setSolidColor_color_)|将图表元素的填充格式设置为统一颜色。|
|[ChartFont](/javascript/api/excel/excel.chartfont)|[bold](/javascript/api/excel/excel.chartfont#bold)|表示字体的加粗状态。|
||[color](/javascript/api/excel/excel.chartfont#color)|文本颜色格式的 HTML 颜色代码表示 (例如，#FF0000红色) 。|
||[italic](/javascript/api/excel/excel.chartfont#italic)|表示字体的斜体状态。|
||[名称](/javascript/api/excel/excel.chartfont#name)|字体名称 (例如"Calibri") |
||[size](/javascript/api/excel/excel.chartfont#size)|字体大小 (例如 11) |
||[underline](/javascript/api/excel/excel.chartfont#underline)|应用于字体的下划线类型。|
|[ChartGridlines](/javascript/api/excel/excel.chartgridlines)|[format](/javascript/api/excel/excel.chartgridlines#format)|表示图表网格线的格式。|
||[visible](/javascript/api/excel/excel.chartgridlines#visible)|指定坐标轴网格线是否可见。|
|[ChartGridlinesFormat](/javascript/api/excel/excel.chartgridlinesformat)|[line](/javascript/api/excel/excel.chartgridlinesformat#line)|表示图表线条格式。|
|[ChartLegend](/javascript/api/excel/excel.chartlegend)|[重叠](/javascript/api/excel/excel.chartlegend#overlay)|指定图表图例是否与图表的主主体重叠。|
||[position](/javascript/api/excel/excel.chartlegend#position)|指定图例在图表上的位置。|
||[format](/javascript/api/excel/excel.chartlegend#format)|表示图表图例的格式，包括填充和字体格式。|
||[visible](/javascript/api/excel/excel.chartlegend#visible)|指定图表图例是否可见。|
|[ChartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|[fill](/javascript/api/excel/excel.chartlegendformat#fill)|表示对象的填充格式，包括背景格式信息。|
||[font](/javascript/api/excel/excel.chartlegendformat#font)|表示字体属性，如图表图例的字体名称、字号和颜色。|
|[ChartLineFormat](/javascript/api/excel/excel.chartlineformat)|[clear()](/javascript/api/excel/excel.chartlineformat#clear__)|清除图表元素的线条格式。|
||[color](/javascript/api/excel/excel.chartlineformat#color)|表示图表中的线条颜色的 HTML 颜色代码。|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|[format](/javascript/api/excel/excel.chartpoint#format)|封装图表点的格式属性。|
||[value](/javascript/api/excel/excel.chartpoint#value)|返回图表点的值。|
|[ChartPointFormat](/javascript/api/excel/excel.chartpointformat)|[fill](/javascript/api/excel/excel.chartpointformat#fill)|表示图表的填充格式，其中包括背景格式信息。|
|[ChartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|[getItemAt(index: number)](/javascript/api/excel/excel.chartpointscollection#getItemAt_index_)|根据其在系列中的位置检索点。|
||[count](/javascript/api/excel/excel.chartpointscollection#count)|返回系列中的图表点数。|
||[items](/javascript/api/excel/excel.chartpointscollection#items)|获取此集合中已加载的子项。|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[名称](/javascript/api/excel/excel.chartseries#name)|指定图表中系列的名称。|
||[format](/javascript/api/excel/excel.chartseries#format)|表示图表系列的格式，包括填充和线条格式。|
||[points](/javascript/api/excel/excel.chartseries#points)|返回系列中所有数据点的集合。|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[getItemAt(index: number)](/javascript/api/excel/excel.chartseriescollection#getItemAt_index_)|根据其在集合中的位置检索系列|
||[count](/javascript/api/excel/excel.chartseriescollection#count)|返回集合中的系列数量。|
||[items](/javascript/api/excel/excel.chartseriescollection#items)|获取此集合中已加载的子项。|
|[ChartSeriesFormat](/javascript/api/excel/excel.chartseriesformat)|[fill](/javascript/api/excel/excel.chartseriesformat#fill)|表示图表系列的填充格式，包括背景格式信息。|
||[line](/javascript/api/excel/excel.chartseriesformat#line)|表示线条格式。|
|[ChartTitle](/javascript/api/excel/excel.charttitle)|[重叠](/javascript/api/excel/excel.charttitle#overlay)|指定图表标题是否覆盖图表。|
||[format](/javascript/api/excel/excel.charttitle#format)|表示图表标题的格式，包括填充和字体格式。|
||[text](/javascript/api/excel/excel.charttitle#text)|指定图表的标题文本。|
||[visible](/javascript/api/excel/excel.charttitle#visible)|指定图表标题是否可见。|
|[ChartTitleFormat](/javascript/api/excel/excel.charttitleformat)|[fill](/javascript/api/excel/excel.charttitleformat#fill)|表示对象的填充格式，包括背景格式信息。|
||[font](/javascript/api/excel/excel.charttitleformat#font)|代表对象的字体 (字体名称、字号和颜色) 等。|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[getRange()](/javascript/api/excel/excel.nameditem#getRange__)|返回与名称相关的 range 对象。|
||[名称](/javascript/api/excel/excel.nameditem#name)|对象的名称。|
||[type](/javascript/api/excel/excel.nameditem#type)|指定名称的公式返回的值的类型。|
||[value](/javascript/api/excel/excel.nameditem#value)|表示 name 公式计算出的值。|
||[visible](/javascript/api/excel/excel.nameditem#visible)|指定对象是否可见。|
|[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection)|[getItem(name: string)](/javascript/api/excel/excel.nameditemcollection#getItem_name_)|使用 `NamedItem` 对象的名称获取对象。|
||[items](/javascript/api/excel/excel.nameditemcollection#items)|获取此集合中已加载的子项。|
|[区域](/javascript/api/excel/excel.range)|[clear(applyTo?: Excel.ClearApplyTo)](/javascript/api/excel/excel.range#clear_applyTo_)|清除区域值、格式、填充、边框等。|
||[删除 (班次：Excel。DeleteShiftDirection) ](/javascript/api/excel/excel.range#delete_shift_)|删除与区域相关的单元格。|
||[formulas](/javascript/api/excel/excel.range#formulas)|表示采用 A1 表示法的公式。|
||[formulasLocal](/javascript/api/excel/excel.range#formulasLocal)|表示采用 A1 样式表示法的公式，使用用户的语言和数字格式区域设置。|
||[getBoundingRect (anotherRange： Range \| string) ](/javascript/api/excel/excel.range#getBoundingRect_anotherRange_)|获取包含指定区域的最小 range 对象。|
||[getCell(row: number, column: number)](/javascript/api/excel/excel.range#getCell_row__column_)|根据行和列编号获取包含单个单元格的 range 对象。|
||[getColumn(column: number)](/javascript/api/excel/excel.range#getColumn_column_)|获取区域中包含的列。|
||[getEntireColumn()](/javascript/api/excel/excel.range#getEntireColumn__)|获取一个对象，该对象代表区域"B4：E11" (例如，如果当前区域代表单元格"B4：E11"，则它是表示列 `getEntireColumn` "B：E") 。|
||[getEntireRow()](/javascript/api/excel/excel.range#getEntireRow__)|获取一个对象，该对象代表区域整行 (例如，如果当前区域代表单元格"B4：E11"，则其为代表行 `GetEntireRow` "4：11") 。|
||[getIntersection (anotherRange： Range \| string) ](/javascript/api/excel/excel.range#getIntersection_anotherRange_)|获取表示指定区域的矩形交集的 range 对象。|
||[getLastCell () ](/javascript/api/excel/excel.range#getLastCell__)|获取区域内的最后一个单元格。|
||[getLastColumn () ](/javascript/api/excel/excel.range#getLastColumn__)|获取区域内的最后一列。|
||[getLastRow () ](/javascript/api/excel/excel.range#getLastRow__)|获取区域内的最后一行。|
||[getOffsetRange(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.range#getOffsetRange_rowOffset__columnOffset_)|获取表示与指定区域偏移的区域的对象。|
||[getRow(row: number)](/javascript/api/excel/excel.range#getRow_row_)|获取范围中包含的行。|
||[插入 (班次：Excel。InsertShiftDirection) ](/javascript/api/excel/excel.range#insert_shift_)|将单个单元格或一系列单元格插入到工作表中取代此区域，并移动其他单元格以留出空间。|
||[numberFormat](/javascript/api/excel/excel.range#numberFormat)|表示Excel区域的电话号码格式代码。|
||[address](/javascript/api/excel/excel.range#address)|指定 A1 样式的范围引用。|
||[addressLocal](/javascript/api/excel/excel.range#addressLocal)|表示以用户语言表示指定区域的范围引用。|
||[cellCount](/javascript/api/excel/excel.range#cellCount)|指定区域中的单元格数。|
||[columnCount](/javascript/api/excel/excel.range#columnCount)|指定范围中的列总数。|
||[columnIndex](/javascript/api/excel/excel.range#columnIndex)|指定范围中第一个单元格的列号。|
||[format](/javascript/api/excel/excel.range#format)|返回一个格式对象，其中封装了区域的字体、填充、边框、对齐方式和其他属性。|
||[rowCount](/javascript/api/excel/excel.range#rowCount)|返回区域中的总行数。|
||[rowIndex](/javascript/api/excel/excel.range#rowIndex)|返回区域中第一个单元格的行编号。|
||[text](/javascript/api/excel/excel.range#text)|指定区域的文本值。|
||[valueTypes](/javascript/api/excel/excel.range#valueTypes)|指定每个单元格中数据类型。|
||[worksheet](/javascript/api/excel/excel.range#worksheet)|包含当前区域的工作表。|
||[select()](/javascript/api/excel/excel.range#select__)|在 Excel UI 中选择指定的区域。|
||[values](/javascript/api/excel/excel.range#values)|表示指定区域的原始值。|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[color](/javascript/api/excel/excel.rangeborder#color)|HTML 颜色代码，表示边框线条的颜色，格式为 #RRGGBB (例如"FFA500") ，或作为已命名的 HTML 颜色 (例如"orange") 。|
||[sideIndex](/javascript/api/excel/excel.rangeborder#sideIndex)|指示边框的特定边的常量值。|
||[style](/javascript/api/excel/excel.rangeborder#style)|线条样式的常量之一，指定边框的线条样式。|
||[weight](/javascript/api/excel/excel.rangeborder#weight)|指定区域周围的边框的粗细。|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[getItem (索引：Excel。BorderIndex) ](/javascript/api/excel/excel.rangebordercollection#getItem_index_)|使用其名称获取 border 对象|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangebordercollection#getItemAt_index_)|使用其索引获取 border 对象|
||[count](/javascript/api/excel/excel.rangebordercollection#count)|集合中的 border 对象数量。|
||[items](/javascript/api/excel/excel.rangebordercollection#items)|获取此集合中已加载的子项。|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[clear()](/javascript/api/excel/excel.rangefill#clear__)|重置区域背景。|
||[color](/javascript/api/excel/excel.rangefill#color)|HTML 颜色代码，表示背景的颜色，格式为 #RRGGBB (例如"FFA500") ，或作为已命名的 HTML 颜色 (例如"orange") |
|[RangeFont](/javascript/api/excel/excel.rangefont)|[bold](/javascript/api/excel/excel.rangefont#bold)|表示字体的加粗状态。|
||[color](/javascript/api/excel/excel.rangefont#color)|文本颜色格式的 HTML 颜色代码表示 (例如，#FF0000红色) 。|
||[italic](/javascript/api/excel/excel.rangefont#italic)|指定字体的 italic 状态。|
||[名称](/javascript/api/excel/excel.rangefont#name)|字体名称 (例如"Calibri") 。|
||[size](/javascript/api/excel/excel.rangefont#size)|字号|
||[underline](/javascript/api/excel/excel.rangefont#underline)|应用于字体的下划线类型。|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[horizontalAlignment](/javascript/api/excel/excel.rangeformat#horizontalAlignment)|表示指定对象的水平对齐方式。|
||[Borders](/javascript/api/excel/excel.rangeformat#borders)|应用于整个区域的 Border 对象的集合。|
||[fill](/javascript/api/excel/excel.rangeformat#fill)|返回在整个区域内定义的 fill 对象。|
||[font](/javascript/api/excel/excel.rangeformat#font)|返回在整个区域内定义的 Font 对象。|
||[verticalAlignment](/javascript/api/excel/excel.rangeformat#verticalAlignment)|表示指定对象的垂直对齐方式。|
||[wrapText](/javascript/api/excel/excel.rangeformat#wrapText)|指定是否Excel对象中的文本换行。|
|[Table](/javascript/api/excel/excel.table)|[delete()](/javascript/api/excel/excel.table#delete__)|删除表。|
||[getDataBodyRange () ](/javascript/api/excel/excel.table#getDataBodyRange__)|获取与表的数据体相关的 range 对象。|
||[getHeaderRowRange () ](/javascript/api/excel/excel.table#getHeaderRowRange__)|获取与表的标题行相关的 range 对象。|
||[getRange()](/javascript/api/excel/excel.table#getRange__)|获取与整个表相关的 range 对象。|
||[getTotalRowRange () ](/javascript/api/excel/excel.table#getTotalRowRange__)|获取与表的总计行相关的 range 对象。|
||[名称](/javascript/api/excel/excel.table#name)|表的名称。|
||[columns](/javascript/api/excel/excel.table#columns)|表示表中所有列的集合。|
||[id](/javascript/api/excel/excel.table#id)|返回用于唯一标识指定工作簿中表的值。|
||[rows](/javascript/api/excel/excel.table#rows)|表示表中所有行的集合。|
||[showHeaders](/javascript/api/excel/excel.table#showHeaders)|指定标题行是否可见。|
||[showTotals](/javascript/api/excel/excel.table#showTotals)|指定总计行是否可见。|
||[style](/javascript/api/excel/excel.table#style)|代表表格样式的常量值。|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[add (address： Range \| string， hasHeaders： boolean) ](/javascript/api/excel/excel.tablecollection#add_address__hasHeaders_)|创建一个新表。|
||[getItem(key: string)](/javascript/api/excel/excel.tablecollection#getItem_key_)|按名称或 ID 获取表。|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablecollection#getItemAt_index_)|根据其在集合中的位置获取表。|
||[count](/javascript/api/excel/excel.tablecollection#count)|返回工作簿中的表数目。|
||[items](/javascript/api/excel/excel.tablecollection#items)|获取此集合中已加载的子项。|
|[TableColumn](/javascript/api/excel/excel.tablecolumn)|[delete()](/javascript/api/excel/excel.tablecolumn#delete__)|从表中删除列。|
||[getDataBodyRange () ](/javascript/api/excel/excel.tablecolumn#getDataBodyRange__)|获取与列的数据体相关的 range 对象。|
||[getHeaderRowRange () ](/javascript/api/excel/excel.tablecolumn#getHeaderRowRange__)|获取与列的标头行相关的 range 对象。|
||[getRange()](/javascript/api/excel/excel.tablecolumn#getRange__)|获取与整个列相关的 range 对象。|
||[getTotalRowRange () ](/javascript/api/excel/excel.tablecolumn#getTotalRowRange__)|获取与列的总计行相关的 range 对象。|
||[名称](/javascript/api/excel/excel.tablecolumn#name)|指定表列的名称。|
||[id](/javascript/api/excel/excel.tablecolumn#id)|返回标识表内的列的唯一键。|
||[index](/javascript/api/excel/excel.tablecolumn#index)|返回表的列集合内列的索引编号。|
||[values](/javascript/api/excel/excel.tablecolumn#values)|表示指定区域的原始值。|
|[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|[add (index？： number， values？： Array<Array<boolean \| string \| number>> \| boolean \| string \| number， name？： string) ](/javascript/api/excel/excel.tablecolumncollection#add_index__values__name_)|向表中添加新列。|
||[getItem (键：数字 \| 字符串) ](/javascript/api/excel/excel.tablecolumncollection#getItem_key_)|按名称或 ID 获取 column 对象。|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablecolumncollection#getItemAt_index_)|根据其在集合中的位置获取列。|
||[count](/javascript/api/excel/excel.tablecolumncollection#count)|返回表中的列数。|
||[items](/javascript/api/excel/excel.tablecolumncollection#items)|获取此集合中已加载的子项。|
|[TableRow](/javascript/api/excel/excel.tablerow)|[delete()](/javascript/api/excel/excel.tablerow#delete__)|从表中删除行。|
||[getRange()](/javascript/api/excel/excel.tablerow#getRange__)|返回与整个行相关的 range 对象。|
||[index](/javascript/api/excel/excel.tablerow#index)|返回表的行集合内行的索引编号。|
||[values](/javascript/api/excel/excel.tablerow#values)|表示指定区域的原始值。|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[add (index？： number， values？： Array<Array<boolean \| string \| number>> \| boolean \| string number \|) ](/javascript/api/excel/excel.tablerowcollection#add_index__values_)|向表中添加一行或多行。|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablerowcollection#getItemAt_index_)|根据其在集合中的位置获取行。|
||[count](/javascript/api/excel/excel.tablerowcollection#count)|返回表中的行数。|
||[items](/javascript/api/excel/excel.tablerowcollection#items)|获取此集合中已加载的子项。|
|[Workbook](/javascript/api/excel/excel.workbook)|[getSelectedRange () ](/javascript/api/excel/excel.workbook#getSelectedRange__)|从工作簿获取当前选定的单个区域。|
||[application](/javascript/api/excel/excel.workbook#application)|表示Excel工作簿的应用程序实例。|
||[bindings](/javascript/api/excel/excel.workbook#bindings)|表示属于工作簿的绑定的集合。|
||[names](/javascript/api/excel/excel.workbook#names)|代表工作簿范围的命名项集合， (范围和常量) 。|
||[表](/javascript/api/excel/excel.workbook#tables)|表示与工作簿关联的表的集合。|
||[worksheets](/javascript/api/excel/excel.workbook#worksheets)|表示与工作簿关联的工作表的集合。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[activate()](/javascript/api/excel/excel.worksheet#activate__)|在 Excel UI 中激活工作表。|
||[delete()](/javascript/api/excel/excel.worksheet#delete__)|从工作簿中删除工作表。|
||[getCell(row: number, column: number)](/javascript/api/excel/excel.worksheet#getCell_row__column_)|基于 `Range` 行号和列号获取包含单个单元格的对象。|
||[getRange (address？： string) ](/javascript/api/excel/excel.worksheet#getRange_address_)|获取 `Range` 对象，该对象代表由地址或名称指定的单个单元格矩形块。|
||[名称](/javascript/api/excel/excel.worksheet#name)|工作表的显示名称。|
||[position](/javascript/api/excel/excel.worksheet#position)|工作表在工作簿中的位置，从零开始。|
||[charts](/javascript/api/excel/excel.worksheet#charts)|返回属于工作表的图表集合。|
||[id](/javascript/api/excel/excel.worksheet#id)|返回用于唯一标识指定工作簿中工作表的值。|
||[表](/javascript/api/excel/excel.worksheet#tables)|属于工作表的表的集合。|
||[visibility](/javascript/api/excel/excel.worksheet#visibility)|工作表的可见性。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[add (name？： string) ](/javascript/api/excel/excel.worksheetcollection#add_name_)|向工作簿添加新工作表。|
||[getActiveWorksheet () ](/javascript/api/excel/excel.worksheetcollection#getActiveWorksheet__)|获取工作簿中当前处于活动状态的工作表。|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcollection#getItem_key_)|使用其名称或 ID 获取 worksheet 对象。|
||[items](/javascript/api/excel/excel.worksheetcollection#items)|获取此集合中已加载的子项。|

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 参考文档](/javascript/api/excel?view=excel-js-1.1&preserve-view=true)
- [Excel JavaScript API 要求集](excel-api-requirement-sets.md)
