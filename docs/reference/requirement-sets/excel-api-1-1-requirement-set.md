---
title: Excel JavaScript API 要求集1。1
description: 有关 ExcelApi 1.1 要求集的详细信息
ms.date: 07/26/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 815b90b18135be22632c39a9824f862149852a84
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/11/2020
ms.locfileid: "47430917"
---
# <a name="excel-javascript-api-requirement-set-11"></a>Excel JavaScript API 要求集1。1

Excel JavaScript API 1.1 是首版 API。 它是 Excel 2016 支持的唯一特定于 Excel 的要求集。

## <a name="api-list"></a>API 列表

下表列出了 Excel JavaScript API 要求集1.1 中的 Api。 若要查看 Excel JavaScript API 要求集1.1 支持的所有 Api 的 API 参考文档，请参阅 [要求集1.1 中的 Excel api](/javascript/api/excel?view=excel-js-1.1&preserve-view=true)。

| Class | 域 | 说明 |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[计算 (calculationType： CalculationType) ](/javascript/api/excel/excel.application#calculate-calculationtype-)|重新计算 Excel 中当前打开的所有工作簿。|
||[calculationMode](/javascript/api/excel/excel.application#calculationmode)|返回工作簿中使用的计算模式，如 CalculationMode 中的常量所定义。 可能的值为： `Automatic` ，excel 控制重新计算的位置; `AutomaticExceptTables` excel 控制重新计算但忽略表中的更改; 在 `Manual` 用户请求计算时完成计算。|
|[Binding](/javascript/api/excel/excel.binding)|[getRange()](/javascript/api/excel/excel.binding#getrange--)|返回绑定表示的区域。如果绑定类型不正确，将引发错误。|
||[getTable()](/javascript/api/excel/excel.binding#gettable--)|返回绑定表示的表。如果绑定类型不正确，将引发错误。|
||[getText()](/javascript/api/excel/excel.binding#gettext--)|返回绑定表示的文本。 如果绑定类型不正确，将引发错误。|
||[id](/javascript/api/excel/excel.binding#id)|表示绑定标识符。 只读。|
||[type](/javascript/api/excel/excel.binding#type)|返回绑定的类型。 有关详细信息，请参阅 BindingType。 只读。|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[getItem(id: string)](/javascript/api/excel/excel.bindingcollection#getitem-id-)|按 ID 获取绑定对象。|
||[getItemAt(index: number)](/javascript/api/excel/excel.bindingcollection#getitemat-index-)|根据其在项目数组中的位置获取绑定对象。|
||[count](/javascript/api/excel/excel.bindingcollection#count)|返回集合中绑定的数量。 只读。|
||[items](/javascript/api/excel/excel.bindingcollection#items)|获取此集合中已加载的子项。|
|[Chart](/javascript/api/excel/excel.chart)|[delete()](/javascript/api/excel/excel.chart#delete--)|删除 chart 对象。|
||[height](/javascript/api/excel/excel.chart#height)|表示 chart 对象的高度，以磅值表示。|
||[left](/javascript/api/excel/excel.chart#left)|从图表左侧到工作表原点的距离，以磅值表示。|
||[name](/javascript/api/excel/excel.chart#name)|表示 chart 对象的名称。|
||[根](/javascript/api/excel/excel.chart#axes)|表示图表坐标轴。 只读。|
||[dataLabels](/javascript/api/excel/excel.chart#datalabels)|表示图表上的数据标签。 只读。|
||[format](/javascript/api/excel/excel.chart#format)|封装图表区域的格式属性。 只读。|
||[图例](/javascript/api/excel/excel.chart#legend)|表示图表的图例。 只读。|
||[series](/javascript/api/excel/excel.chart#series)|表示单个系列或图表中的系列集合。 只读。|
||[title](/javascript/api/excel/excel.chart#title)|表示指定图表的标题，包括标题的文本、可见性、位置和格式。 只读。|
||[setData (sourceData： Range，seriesBy？： ChartSeriesBy) ](/javascript/api/excel/excel.chart#setdata-sourcedata--seriesby-)|重置图表的源数据。|
||[setPosition (startCell： Range \| string，endCell？： range \| string) ](/javascript/api/excel/excel.chart#setposition-startcell--endcell-)|相对于工作表上的单元格放置图表。|
||[top](/javascript/api/excel/excel.chart#top)|表示从对象左边界至第 1 行顶部（在工作表上）或图表区域顶部（在图表上）的距离，以磅值表示。|
||[width](/javascript/api/excel/excel.chart#width)|表示 chart 对象的宽度，以磅值表示。|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[fill](/javascript/api/excel/excel.chartareaformat#fill)|表示对象的填充格式，包括背景格式信息。 只读。|
||[font](/javascript/api/excel/excel.chartareaformat#font)|表示当前对象的字体属性（字体名称、字体大小、颜色等）。 只读。|
|[ChartAxes](/javascript/api/excel/excel.chartaxes)|[categoryAxis](/javascript/api/excel/excel.chartaxes#categoryaxis)|表示图表中的类别轴。 只读。|
||[seriesAxis](/javascript/api/excel/excel.chartaxes#seriesaxis)|表示三维图表的系列轴。 只读。|
||[值坐标轴](/javascript/api/excel/excel.chartaxes#valueaxis)|表示坐标轴中的数值轴。 只读。|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[majorUnit](/javascript/api/excel/excel.chartaxis#majorunit)|表示两个主要刻度标记之间的间隔。可以设置为数字值或空字符串。返回的值始终为数字。|
||[maximum](/javascript/api/excel/excel.chartaxis#maximum)|表示数值轴上的最大值。可以设置为数字值或空字符串（对于自动坐标轴值）。返回的值始终为数字。|
||[minimum](/javascript/api/excel/excel.chartaxis#minimum)|表示数值轴上的最小值。可以设置为数字值或空字符串（对于自动坐标轴值）。返回的值始终为数字。|
||[minorUnit](/javascript/api/excel/excel.chartaxis#minorunit)|表示两个次要刻度标记之间的间隔。 可以设置为数字值或空字符串（对于自动坐标轴值）。 返回的值始终为数字。|
||[format](/javascript/api/excel/excel.chartaxis#format)|表示 chart 对象的格式，包括线条和字体格式。 只读。|
||[majorGridlines](/javascript/api/excel/excel.chartaxis#majorgridlines)|返回一个表示指定坐标轴的主要网格线的网格线对象。 只读。|
||[minorGridlines](/javascript/api/excel/excel.chartaxis#minorgridlines)|返回一个表示指定坐标轴的次要网格线的网格线对象。 只读。|
||[title](/javascript/api/excel/excel.chartaxis#title)|表示坐标轴标题。 只读。|
|[ChartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|[font](/javascript/api/excel/excel.chartaxisformat#font)|表示图表坐标轴元素的字体属性（字体名称、字体大小、颜色等）。 只读。|
||[line](/javascript/api/excel/excel.chartaxisformat#line)|表示图表线条格式。 只读。|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[format](/javascript/api/excel/excel.chartaxistitle#format)|表示图表坐标轴标题的格式。 只读。|
||[text](/javascript/api/excel/excel.chartaxistitle#text)|表示坐标轴标题。|
||[visible](/javascript/api/excel/excel.chartaxistitle#visible)|指定坐标轴标题是否可见的布尔值。|
|[ChartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|[font](/javascript/api/excel/excel.chartaxistitleformat#font)|表示图表坐标轴标题对象的字体属性，例如字体名称、字体大小、颜色等。 只读。|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[添加 (类型： ChartType、sourceData： Range、seriesBy？： ChartSeriesBy) ](/javascript/api/excel/excel.chartcollection#add-type--sourcedata--seriesby-)|创建新图表。|
||[getItem(name: string)](/javascript/api/excel/excel.chartcollection#getitem-name-)|使用图表名称获取图表。 如果存在多个名称相同的图表，将返回第一个图表。|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartcollection#getitemat-index-)|根据其在集合中的位置获取图表。|
||[count](/javascript/api/excel/excel.chartcollection#count)|返回工作表中的图表数。 只读。|
||[items](/javascript/api/excel/excel.chartcollection#items)|获取此集合中已加载的子项。|
|[ChartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|[fill](/javascript/api/excel/excel.chartdatalabelformat#fill)|表示当前图表数据标签的填充格式。 只读。|
||[font](/javascript/api/excel/excel.chartdatalabelformat#font)|表示图表数据标签的字体属性（字体名称、字体大小、颜色等）。 只读。|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[position](/javascript/api/excel/excel.chartdatalabels#position)|表示数据标签的位置的 DataLabelPosition 值。 有关详细信息，请参阅 ChartDataLabelPosition。|
||[format](/javascript/api/excel/excel.chartdatalabels#format)|表示图表数据标签的格式，包括填充和字体格式。 只读。|
||[分隔符](/javascript/api/excel/excel.chartdatalabels#separator)|表示用于图表中数据标签的分隔符的字符串。|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabels#showbubblesize)|该布尔值表示数据标签气泡大小是否可见。|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabels#showcategoryname)|表示数据标签分类名称是否可见的布尔值。|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabels#showlegendkey)|该布尔值表示数据标签图例标示是否可见。|
||[showPercentage](/javascript/api/excel/excel.chartdatalabels#showpercentage)|该布尔值表示数据标签百分比是否可见。|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabels#showseriesname)|该布尔值表示数据标签系列名称是否可见。|
||[showValue](/javascript/api/excel/excel.chartdatalabels#showvalue)|该布尔值表示数据标签值是否可见。|
|[ChartFill](/javascript/api/excel/excel.chartfill)|[clear()](/javascript/api/excel/excel.chartfill#clear--)|清除图表元素的填充颜色。|
||[setSolidColor(color: string)](/javascript/api/excel/excel.chartfill#setsolidcolor-color-)|将图表元素的填充格式设置为统一颜色。|
|[ChartFont](/javascript/api/excel/excel.chartfont)|[bold](/javascript/api/excel/excel.chartfont#bold)|表示字体的加粗状态。|
||[color](/javascript/api/excel/excel.chartfont#color)|文本颜色的 HTML 颜色代码表示。 例如， #FF0000 代表红色。|
||[italic](/javascript/api/excel/excel.chartfont#italic)|表示字体的斜体状态。|
||[name](/javascript/api/excel/excel.chartfont#name)|字体名称（例如"Calibri"）|
||[size](/javascript/api/excel/excel.chartfont#size)|字号（例如，11）|
||[underline](/javascript/api/excel/excel.chartfont#underline)|应用于字体的下划线类型。 有关详细信息，请参阅 ChartUnderlineStyle。|
|[ChartGridlines](/javascript/api/excel/excel.chartgridlines)|[format](/javascript/api/excel/excel.chartgridlines#format)|表示图表网格线的格式。 只读。|
||[visible](/javascript/api/excel/excel.chartgridlines#visible)|表示坐标轴网格线是否可见的布尔值。|
|[ChartGridlinesFormat](/javascript/api/excel/excel.chartgridlinesformat)|[line](/javascript/api/excel/excel.chartgridlinesformat#line)|表示图表线条格式。 只读。|
|[ChartLegend](/javascript/api/excel/excel.chartlegend)|[重叠](/javascript/api/excel/excel.chartlegend#overlay)|表示图表图例是否应该与图表的主体重叠的布尔值。|
||[position](/javascript/api/excel/excel.chartlegend#position)|表示图例在图表上的位置。 有关详细信息，请参阅 ChartLegendPosition。|
||[format](/javascript/api/excel/excel.chartlegend#format)|表示图表图例的格式，包括填充和字体格式。 只读。|
||[visible](/javascript/api/excel/excel.chartlegend#visible)|表示 ChartLegend 对象是否可见的布尔值。|
|[ChartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|[fill](/javascript/api/excel/excel.chartlegendformat#fill)|表示对象的填充格式，包括背景格式信息。 只读。|
||[font](/javascript/api/excel/excel.chartlegendformat#font)|表示图表图例的字体属性，例如字体名称、字体大小、颜色等。 只读。|
|[ChartLineFormat](/javascript/api/excel/excel.chartlineformat)|[clear()](/javascript/api/excel/excel.chartlineformat#clear--)|清除图表元素的线条格式。|
||[color](/javascript/api/excel/excel.chartlineformat#color)|表示图表中的线条颜色的 HTML 颜色代码。|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|[format](/javascript/api/excel/excel.chartpoint#format)|封装图表点的格式属性。 只读。|
||[value](/javascript/api/excel/excel.chartpoint#value)|返回图表点的值。 只读。|
|[ChartPointFormat](/javascript/api/excel/excel.chartpointformat)|[fill](/javascript/api/excel/excel.chartpointformat#fill)|代表图表的填充格式，其中包括背景格式信息。 只读。|
|[ChartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|[getItemAt(index: number)](/javascript/api/excel/excel.chartpointscollection#getitemat-index-)|根据其在系列中的位置检索点。|
||[count](/javascript/api/excel/excel.chartpointscollection#count)|返回系列中的图表点数。 只读。|
||[items](/javascript/api/excel/excel.chartpointscollection#items)|获取此集合中已加载的子项。|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[name](/javascript/api/excel/excel.chartseries#name)|表示图表中某个系列的名称。|
||[format](/javascript/api/excel/excel.chartseries#format)|表示图表系列的格式，包括填充和线条格式。 只读。|
||[点](/javascript/api/excel/excel.chartseries#points)|表示系列中所有数据点的集合。 只读。|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[getItemAt(index: number)](/javascript/api/excel/excel.chartseriescollection#getitemat-index-)|根据其在集合中的位置检索系列|
||[count](/javascript/api/excel/excel.chartseriescollection#count)|返回集合中的系列数量。 只读。|
||[items](/javascript/api/excel/excel.chartseriescollection#items)|获取此集合中已加载的子项。|
|[ChartSeriesFormat](/javascript/api/excel/excel.chartseriesformat)|[fill](/javascript/api/excel/excel.chartseriesformat#fill)|表示图表系列的填充格式，包括背景格式信息。 只读。|
||[line](/javascript/api/excel/excel.chartseriesformat#line)|表示线条格式。 只读。|
|[ChartTitle](/javascript/api/excel/excel.charttitle)|[重叠](/javascript/api/excel/excel.charttitle#overlay)|表示图表标题是否将叠加在图表上的布尔值。|
||[format](/javascript/api/excel/excel.charttitle#format)|表示图表标题的格式，包括填充和字体格式。 只读。|
||[text](/javascript/api/excel/excel.charttitle#text)|表示图表的标题文本。|
||[visible](/javascript/api/excel/excel.charttitle#visible)|表示图表标题对象是否可见的布尔值。|
|[ChartTitleFormat](/javascript/api/excel/excel.charttitleformat)|[fill](/javascript/api/excel/excel.charttitleformat#fill)|表示对象的填充格式，包括背景格式信息。 只读。|
||[font](/javascript/api/excel/excel.charttitleformat#font)|表示对象的字体属性（字体名称、字体大小、颜色等）。 只读。|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[getRange()](/javascript/api/excel/excel.nameditem#getrange--)|返回与名称相关联的 Range 对象。 如果已命名项的类型不是 Range，将引发错误。|
||[name](/javascript/api/excel/excel.nameditem#name)|对象的名称。 只读。|
||[type](/javascript/api/excel/excel.nameditem#type)|指明 name 公式返回的值的类型。 有关详细信息，请参阅 NamedItemType。 只读。|
||[value](/javascript/api/excel/excel.nameditem#value)|表示 name 公式计算出的值。 对于已命名区域，将返回区域地址。 只读。|
||[visible](/javascript/api/excel/excel.nameditem#visible)|指定对象是否可见。|
|[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection)|[getItem(name: string)](/javascript/api/excel/excel.nameditemcollection#getitem-name-)|使用其名称获取 NamedItem 对象。|
||[items](/javascript/api/excel/excel.nameditemcollection#items)|获取此集合中已加载的子项。|
|[区域](/javascript/api/excel/excel.range)|[clear(applyTo?: Excel.ClearApplyTo)](/javascript/api/excel/excel.range#clear-applyto-)|清除区域值、格式、填充、边框等。|
||[删除 (shift： DeleteShiftDirection) ](/javascript/api/excel/excel.range#delete-shift-)|删除与区域相关的单元格。|
||[formulas](/javascript/api/excel/excel.range#formulas)|表示采用 A1 表示法的公式。|
||[formulasLocal](/javascript/api/excel/excel.range#formulaslocal)|表示采用 A1 样式表示法的公式，使用用户的语言和数字格式区域设置。例如，英语中的公式 "=SUM(A1, 1.5)" 在德语中将变为 "=SUMME(A1; 1,5)"。|
||[getBoundingRect (anotherRange： Range \| string) ](/javascript/api/excel/excel.range#getboundingrect-anotherrange-)|获取包含指定区域的最小 range 对象。 例如，“B2:C5”和“D10:E15”的 GetBoundingRect 为“B2:E15”。|
||[getCell(row: number, column: number)](/javascript/api/excel/excel.range#getcell-row--column-)|根据行和列编号获取包含单个单元格的 range 对象。 单元格可以位于其父区域的边界之外，但前提是它停留在工作表网格中。 返回的单元格位于相对于区域左上角的单元格的位置。|
||[getColumn(column: number)](/javascript/api/excel/excel.range#getcolumn-column-)|获取区域中包含的列。|
||[getEntireColumn()](/javascript/api/excel/excel.range#getentirecolumn--)|获取一个对象，该对象表示区域中 (的整列。例如，如果当前区域表示单元格 "B4： E11"， `getEntireColumn` 则它是表示列 "B:E" 的区域 ) 。|
||[getEntireRow()](/javascript/api/excel/excel.range#getentirerow--)|获取表示区域中整行的对象 (例如，如果当前区域表示单元格 "B4： E11"， `GetEntireRow` 则它是表示行 "4:11" ) 的区域。|
||[getIntersection (anotherRange： Range \| string) ](/javascript/api/excel/excel.range#getintersection-anotherrange-)|获取表示指定区域的矩形交集的 range 对象。|
||[getLastCell ( # B1 ](/javascript/api/excel/excel.range#getlastcell--)|获取区域内的最后一个单元格。例如，“B2:D5”的最后一个单元格是“D5”。|
||[getLastColumn ( # B1 ](/javascript/api/excel/excel.range#getlastcolumn--)|获取区域内的最后一列。例如，“B2:D5”的最后一列是“D2:D5”。|
||[getLastRow ( # B1 ](/javascript/api/excel/excel.range#getlastrow--)|获取区域内的最后一行。例如，“B2:D5”的最后一行是“B5:D5”。|
||[getOffsetRange(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-)|获取表示与指定区域偏移的区域的对象。返回的区域的尺寸将与此区域一致。如果强制在工作表网格的边界之外生成区域，将引发错误。|
||[getRow(row: number)](/javascript/api/excel/excel.range#getrow-row-)|获取范围中包含的行。|
||[插入 (shift： InsertShiftDirection) ](/javascript/api/excel/excel.range#insert-shift-)|将单个单元格或一系列单元格插入到工作表中取代此区域，并移动其他单元格以留出空间。在现在空白的空间返回新的 Range 对象。|
||[numberFormat](/javascript/api/excel/excel.range#numberformat)|表示给定范围的 Excel 数字格式代码。|
||[address](/javascript/api/excel/excel.range#address)|表示 A1 样式的区域引用。 Address 值将包含工作表引用 (例如 "Sheet1！A1： B4 ") 。 只读。|
||[addressLocal](/javascript/api/excel/excel.range#addresslocal)|以用户语言表示对指定区域的区域引用。 只读。|
||[cellCount](/javascript/api/excel/excel.range#cellcount)|范围中的单元格数。 如果单元格数超过 2^31-1 (2,147,483,647)，此 API 返回 -1。 只读。|
||[columnCount](/javascript/api/excel/excel.range#columncount)|表示区域中的列总数。 只读。|
||[columnIndex](/javascript/api/excel/excel.range#columnindex)|表示区域中第一个单元格的列编号。 从零开始编制索引。 只读。|
||[format](/javascript/api/excel/excel.range#format)|返回一个格式对象，其中封装了区域的字体、填充、边框、对齐方式和其他属性。 只读。|
||[rowCount](/javascript/api/excel/excel.range#rowcount)|返回区域中的总行数。 只读。|
||[rowIndex](/javascript/api/excel/excel.range#rowindex)|返回区域中第一个单元格的行编号。 从零开始编制索引。 只读。|
||[text](/javascript/api/excel/excel.range#text)|指定区域的文本值。文本值与单元格宽度无关。在 Excel UI 中替代 # 符号不会影响 API 返回的文本值。只读。|
||[valueTypes](/javascript/api/excel/excel.range#valuetypes)|表示每个单元格的数据类型。 只读。|
||[worksheet](/javascript/api/excel/excel.range#worksheet)|包含当前区域的工作表。 只读。|
||[select()](/javascript/api/excel/excel.range#select--)|在 Excel UI 中选择指定的区域。|
||[values](/javascript/api/excel/excel.range#values)|表示指定区域的原始值。 返回的数据可能是字符串、数字，也可能是布尔值。 包含错误的单元格将返回错误字符串。|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[color](/javascript/api/excel/excel.rangeborder#color)|表示窗体 #RRGGBB（例如“FFA500”）的边框线条颜色或作为已命名的 HTML 颜色（例如“orange”）的 HTML 颜色代码。|
||[sideIndex](/javascript/api/excel/excel.rangeborder#sideindex)|指示边框的特定边的常量值。 有关详细信息，请参阅 BorderIndex。 只读。|
||[style](/javascript/api/excel/excel.rangeborder#style)|线条样式的常量之一，指定边框的线条样式。 有关详细信息，请参阅 BorderLineStyle。|
||[weight](/javascript/api/excel/excel.rangeborder#weight)|指定区域周围的边框的粗细。 有关详细信息，请参阅 BorderWeight。|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[getItem (索引： BorderIndex) ](/javascript/api/excel/excel.rangebordercollection#getitem-index-)|使用其名称获取 border 对象|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangebordercollection#getitemat-index-)|使用其索引获取 border 对象|
||[count](/javascript/api/excel/excel.rangebordercollection#count)|集合中的 border 对象数量。 只读。|
||[items](/javascript/api/excel/excel.rangebordercollection#items)|获取此集合中已加载的子项。|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[clear()](/javascript/api/excel/excel.rangefill#clear--)|重置区域背景。|
||[color](/javascript/api/excel/excel.rangefill#color)|表示窗体 #RRGGBB（例如“FFA500”）的边框线条颜色或作为已命名的 HTML 颜色（例如“orange”）的 HTML 颜色代码。|
|[RangeFont](/javascript/api/excel/excel.rangefont)|[bold](/javascript/api/excel/excel.rangefont#bold)|表示字体的加粗状态。|
||[color](/javascript/api/excel/excel.rangefont#color)|文本颜色的 HTML 颜色代码表示。 例如， #FF0000 代表红色。|
||[italic](/javascript/api/excel/excel.rangefont#italic)|表示字体的斜体状态。|
||[name](/javascript/api/excel/excel.rangefont#name)|字体名称（例如"Calibri"）|
||[size](/javascript/api/excel/excel.rangefont#size)|字号|
||[underline](/javascript/api/excel/excel.rangefont#underline)|应用于字体的下划线类型。 有关详细信息，请参阅 RangeUnderlineStyle。|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[horizontalAlignment](/javascript/api/excel/excel.rangeformat#horizontalalignment)|表示指定对象的水平对齐方式。 有关详细信息，请参阅 HorizontalAlignment。|
||[Borders](/javascript/api/excel/excel.rangeformat#borders)|应用于整个区域的 Border 对象的集合。 只读。|
||[fill](/javascript/api/excel/excel.rangeformat#fill)|返回在整个区域内定义的 fill 对象。 只读。|
||[font](/javascript/api/excel/excel.rangeformat#font)|返回在整个区域内定义的 Font 对象。 只读。|
||[verticalAlignment](/javascript/api/excel/excel.rangeformat#verticalalignment)|表示指定对象的垂直对齐方式。 有关详细信息，请参阅 VerticalAlignment。|
||[wrapText](/javascript/api/excel/excel.rangeformat#wraptext)|指示 Excel 是否将对象中的文本换行。 指示整个区域不具有统一换行设置的空值|
|[Table](/javascript/api/excel/excel.table)|[delete()](/javascript/api/excel/excel.table#delete--)|删除表。|
||[getDataBodyRange ( # B1 ](/javascript/api/excel/excel.table#getdatabodyrange--)|获取与表的数据体相关的 range 对象。|
||[getHeaderRowRange ( # B1 ](/javascript/api/excel/excel.table#getheaderrowrange--)|获取与表的标题行相关的 range 对象。|
||[getRange()](/javascript/api/excel/excel.table#getrange--)|获取与整个表相关的 range 对象。|
||[getTotalRowRange ( # B1 ](/javascript/api/excel/excel.table#gettotalrowrange--)|获取与表的总计行相关的 range 对象。|
||[name](/javascript/api/excel/excel.table#name)|表的名称。|
||[columns](/javascript/api/excel/excel.table#columns)|表示表中所有列的集合。 只读。|
||[id](/javascript/api/excel/excel.table#id)|返回用于唯一标识指定工作簿中表的值。即使表被重命名，标识符的值仍然相同。只读。|
||[rows](/javascript/api/excel/excel.table#rows)|表示表中所有行的集合。 只读。|
||[showHeaders](/javascript/api/excel/excel.table#showheaders)|指示标头行是否可见。 该值可以设置为显示或删除标头行。|
||[showTotals](/javascript/api/excel/excel.table#showtotals)|指示总计行是否可见。 该值可以设置为显示或删除总计行。|
||[style](/javascript/api/excel/excel.table#style)|表示表格样式的常量值。可能的值是：TableStyleLight1 thru TableStyleLight21、TableStyleMedium1 thru TableStyleMedium28、TableStyleStyleDark1 thru TableStyleStyleDark11。还可以指定工作簿中显示的用户定义的自定义样式。|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[add (address： Range \| string，hasHeaders： boolean) ](/javascript/api/excel/excel.tablecollection#add-address--hasheaders-)|新建表。范围对象或源地址决定了在哪个工作表下添加表。如果无法添加表（例如，由于地址无效，或者表与另一个表重叠），则会引发错误。|
||[getItem(key: string)](/javascript/api/excel/excel.tablecollection#getitem-key-)|按名称或 ID 获取表。|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablecollection#getitemat-index-)|根据其在集合中的位置获取表。|
||[count](/javascript/api/excel/excel.tablecollection#count)|返回工作簿中的表数目。 只读。|
||[items](/javascript/api/excel/excel.tablecollection#items)|获取此集合中已加载的子项。|
|[TableColumn](/javascript/api/excel/excel.tablecolumn)|[delete()](/javascript/api/excel/excel.tablecolumn#delete--)|从表中删除列。|
||[getDataBodyRange ( # B1 ](/javascript/api/excel/excel.tablecolumn#getdatabodyrange--)|获取与列的数据体相关的 range 对象。|
||[getHeaderRowRange ( # B1 ](/javascript/api/excel/excel.tablecolumn#getheaderrowrange--)|获取与列的标头行相关的 range 对象。|
||[getRange()](/javascript/api/excel/excel.tablecolumn#getrange--)|获取与整个列相关的 range 对象。|
||[getTotalRowRange ( # B1 ](/javascript/api/excel/excel.tablecolumn#gettotalrowrange--)|获取与列的总计行相关的 range 对象。|
||[name](/javascript/api/excel/excel.tablecolumn#name)|表示表列的名称。|
||[id](/javascript/api/excel/excel.tablecolumn#id)|返回标识表内的列的唯一键。 只读。|
||[index](/javascript/api/excel/excel.tablecolumn#index)|返回表的列集合内列的索引编号。 从零开始编制索引。 只读。|
||[values](/javascript/api/excel/excel.tablecolumn#values)|表示指定区域的原始值。 返回的数据可能是字符串、数字，也可能是布尔值。 包含错误的单元格将返回错误字符串。|
|[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|[是否添加 (索引？： number，values？： Array<数组<布尔字符串编号 \| \|>> \| 布尔 \| 字符串 \| 编号，name？： string) ](/javascript/api/excel/excel.tablecolumncollection#add-index--values--name-)|向表中添加新列。|
||[getItem (项：数字 \| 字符串) ](/javascript/api/excel/excel.tablecolumncollection#getitem-key-)|按名称或 ID 获取 column 对象。|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablecolumncollection#getitemat-index-)|根据其在集合中的位置获取列。|
||[count](/javascript/api/excel/excel.tablecolumncollection#count)|返回表中的列数。 只读。|
||[items](/javascript/api/excel/excel.tablecolumncollection#items)|获取此集合中已加载的子项。|
|[TableRow](/javascript/api/excel/excel.tablerow)|[delete()](/javascript/api/excel/excel.tablerow#delete--)|从表中删除行。|
||[getRange()](/javascript/api/excel/excel.tablerow#getrange--)|返回与整个行相关的 range 对象。|
||[index](/javascript/api/excel/excel.tablerow#index)|返回表的行集合内行的索引编号。 从零开始编制索引。 只读。|
||[values](/javascript/api/excel/excel.tablerow#values)|表示指定区域的原始值。 返回的数据可能是字符串、数字，也可能是布尔值。 包含错误的单元格将返回错误字符串。|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[是否添加 (索引？： number，values？： Array<数组<布尔字符串数字 \| \|>> \| 布尔 \| 字符串数字 \|) ](/javascript/api/excel/excel.tablerowcollection#add-index--values-)|向表中添加一行或多行。 返回对象是新添加的首行。|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablerowcollection#getitemat-index-)|根据其在集合中的位置获取行。|
||[count](/javascript/api/excel/excel.tablerowcollection#count)|返回表中的行数。 只读。|
||[items](/javascript/api/excel/excel.tablerowcollection#items)|获取此集合中已加载的子项。|
|[Workbook](/javascript/api/excel/excel.workbook)|[getSelectedRange ( # B1 ](/javascript/api/excel/excel.workbook#getselectedrange--)|从工作簿中获取当前选定的单个区域。 如果选择了多个区域，则此方法将引发错误。|
||[application](/javascript/api/excel/excel.workbook#application)|表示包含此工作簿的 Excel 应用程序实例。 只读。|
||[bindings](/javascript/api/excel/excel.workbook#bindings)|表示属于工作簿的绑定的集合。 只读。|
||[names](/javascript/api/excel/excel.workbook#names)|表示工作簿范围内的已命名项目（称为区域和常量）的集合。 只读。|
||[表](/javascript/api/excel/excel.workbook#tables)|表示与工作簿关联的表的集合。 只读。|
||[单](/javascript/api/excel/excel.workbook#worksheets)|表示与工作簿关联的工作表的集合。 只读。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[activate()](/javascript/api/excel/excel.worksheet#activate--)|在 Excel UI 中激活工作表。|
||[delete()](/javascript/api/excel/excel.worksheet#delete--)|从工作簿中删除工作表。 请注意，如果工作表的可见性设置为 "VeryHidden"，则删除操作将失败，并出现 GeneralException。|
||[getCell(row: number, column: number)](/javascript/api/excel/excel.worksheet#getcell-row--column-)|根据行和列编号获取包含单个单元格的 range 对象。 单元格可以位于其父区域的边界之外，但前提是它停留在工作表网格中。|
||[getRange (address？： string) ](/javascript/api/excel/excel.worksheet#getrange-address-)|获取一个 range 对象，该对象代表由地址或名称指定的单个矩形单元格块。|
||[name](/javascript/api/excel/excel.worksheet#name)|工作表的显示名称。|
||[position](/javascript/api/excel/excel.worksheet#position)|工作表在工作簿中的位置，从零开始。|
||[直方图](/javascript/api/excel/excel.worksheet#charts)|返回属于工作表的图表的集合。 只读。|
||[id](/javascript/api/excel/excel.worksheet#id)|返回用于唯一标识指定工作簿中工作表的值。即使工作表被重命名或移动，标识符的值仍然相同。只读。|
||[表](/javascript/api/excel/excel.worksheet#tables)|属于工作表的表的集合。 只读。|
||[visibility](/javascript/api/excel/excel.worksheet#visibility)|工作表的可见性。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[添加 (名称？： string) ](/javascript/api/excel/excel.worksheetcollection#add-name-)|向工作簿添加新工作表。工作表将添加到现有工作表的末尾。如果您想要激活新添加的工作表，请对其调用 ".activate()。|
||[getActiveWorksheet ( # B1 ](/javascript/api/excel/excel.worksheetcollection#getactiveworksheet--)|获取工作簿中当前处于活动状态的工作表。|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcollection#getitem-key-)|按 Worksheet 对象的名称或 ID 获取此对象。|
||[items](/javascript/api/excel/excel.worksheetcollection#items)|获取此集合中已加载的子项。|

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 参考文档](/javascript/api/excel?view=excel-js-1.1&preserve-view=true)
- [Excel JavaScript API 要求集](./excel-api-requirement-sets.md)
