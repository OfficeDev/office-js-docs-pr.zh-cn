---
title: Excel JavaScript API 要求集 1.1
description: 有关 ExcelApi 1.1 要求集的详细信息。
ms.date: 11/09/2020
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 45061afc7e401e18a67377bf88fa1670bb7a8ece
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/23/2022
ms.locfileid: "63745953"
---
# <a name="excel-javascript-api-requirement-set-11"></a>Excel JavaScript API 要求集 1.1

Excel JavaScript API 1.1 是首版 API。 这是唯一Excel支持的特定要求Excel 2016。

## <a name="api-list"></a>API 列表

下表列出了 JavaScript API 要求集 1.1 Excel中的 API。 若要查看 JavaScript API 要求集 1.1 Excel所有 API 的 API 参考文档，请参阅要求集 [1.1](/javascript/api/excel?view=excel-js-1.1&preserve-view=true) Excel API。

| 类 | 域 | 说明 |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[calculate (calculationType： Excel.CalculationType) ](/javascript/api/excel/excel.application#excel-excel-application-calculate-member(1))|重新计算 Excel 中当前打开的所有工作簿。|
||[calculationMode](/javascript/api/excel/excel.application#excel-excel-application-calculationmode-member)|返回工作簿中使用的计算模式，如 中的常量所定义 `Excel.CalculationMode`。|
|[Binding](/javascript/api/excel/excel.binding)|[getRange()](/javascript/api/excel/excel.binding#excel-excel-binding-getrange-member(1))|返回绑定表示的区域。|
||[getTable()](/javascript/api/excel/excel.binding#excel-excel-binding-gettable-member(1))|返回绑定表示的表。|
||[getText()](/javascript/api/excel/excel.binding#excel-excel-binding-gettext-member(1))|返回绑定表示的文本。|
||[id](/javascript/api/excel/excel.binding#excel-excel-binding-id-member)|表示绑定标识符。|
||[type](/javascript/api/excel/excel.binding#excel-excel-binding-type-member)|返回绑定的类型。|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[count](/javascript/api/excel/excel.bindingcollection#excel-excel-bindingcollection-count-member)|返回集合中绑定的数量。|
||[getItem(id: string)](/javascript/api/excel/excel.bindingcollection#excel-excel-bindingcollection-getitem-member(1))|按 ID 获取绑定对象。|
||[getItemAt(index: number)](/javascript/api/excel/excel.bindingcollection#excel-excel-bindingcollection-getitemat-member(1))|根据其在项目数组中的位置获取绑定对象。|
||[items](/javascript/api/excel/excel.bindingcollection#excel-excel-bindingcollection-items-member)|获取此集合中已加载的子项。|
|[Chart](/javascript/api/excel/excel.chart)|[axes](/javascript/api/excel/excel.chart#excel-excel-chart-axes-member)|表示图表坐标轴。|
||[dataLabels](/javascript/api/excel/excel.chart#excel-excel-chart-datalabels-member)|表示图表上的数据标签。|
||[delete()](/javascript/api/excel/excel.chart#excel-excel-chart-delete-member(1))|删除 chart 对象。|
||[format](/javascript/api/excel/excel.chart#excel-excel-chart-format-member)|封装图表区域的格式属性。|
||[height](/javascript/api/excel/excel.chart#excel-excel-chart-height-member)|指定图表对象的高度（以点表示）。|
||[left](/javascript/api/excel/excel.chart#excel-excel-chart-left-member)|从图表左侧到工作表原点的距离，以磅为单位。|
||[legend](/javascript/api/excel/excel.chart#excel-excel-chart-legend-member)|表示图表的图例。|
||[name](/javascript/api/excel/excel.chart#excel-excel-chart-name-member)|指定图表对象的名称。|
||[series](/javascript/api/excel/excel.chart#excel-excel-chart-series-member)|表示单个系列或图表中的系列集合。|
||[setData (sourceData： Range， seriesBy？： Excel。ChartSeriesBy) ](/javascript/api/excel/excel.chart#excel-excel-chart-setdata-member(1))|重置图表的源数据。|
||[setPosition (startCell： Range \| string， endCell？： Range \| string) ](/javascript/api/excel/excel.chart#excel-excel-chart-setposition-member(1))|相对于工作表上的单元格放置图表。|
||[title](/javascript/api/excel/excel.chart#excel-excel-chart-title-member)|表示指定图表的标题，包括标题的文本、可见性、位置和格式。|
||[top](/javascript/api/excel/excel.chart#excel-excel-chart-top-member)|指定从对象上边缘到工作表工作表上第 1 (行顶端的距离（以) 或图表图表 (顶部的) ）。|
||[width](/javascript/api/excel/excel.chart#excel-excel-chart-width-member)|指定图表对象的宽度（以点表示）。|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[fill](/javascript/api/excel/excel.chartareaformat#excel-excel-chartareaformat-fill-member)|表示对象的填充格式，包括背景格式信息。|
||[font](/javascript/api/excel/excel.chartareaformat#excel-excel-chartareaformat-font-member)|表示当前对象的字体属性（字体名称、字体大小、颜色等）。|
|[ChartAxes](/javascript/api/excel/excel.chartaxes)|[categoryAxis](/javascript/api/excel/excel.chartaxes#excel-excel-chartaxes-categoryaxis-member)|表示图表中的类别轴。|
||[seriesAxis](/javascript/api/excel/excel.chartaxes#excel-excel-chartaxes-seriesaxis-member)|表示三维图表的系列轴。|
||[valueAxis](/javascript/api/excel/excel.chartaxes#excel-excel-chartaxes-valueaxis-member)|表示坐标轴中的数值轴。|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[format](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-format-member)|表示 chart 对象的格式，包括线条和字体格式。|
||[majorGridlines](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-majorgridlines-member)|返回一个对象，该对象代表指定坐标轴的主要网格线。|
||[majorUnit](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-majorunit-member)|表示两个主要刻度标记之间的间隔。|
||[maximum](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-maximum-member)|表示数值轴上的最大值。|
||[minimum](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-minimum-member)|表示数值轴上的最小值。|
||[minorGridlines](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-minorgridlines-member)|返回一个对象，该对象代表指定坐标轴的次要网格线。|
||[minorUnit](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-minorunit-member)|表示两个次要刻度标记之间的间隔。|
||[title](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-title-member)|表示坐标轴标题。|
|[ChartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|[font](/javascript/api/excel/excel.chartaxisformat#excel-excel-chartaxisformat-font-member)|指定图表坐标轴元素 (字体名称、字体大小、颜色等) 字体属性。|
||[line](/javascript/api/excel/excel.chartaxisformat#excel-excel-chartaxisformat-line-member)|指定图表线条格式。|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[format](/javascript/api/excel/excel.chartaxistitle#excel-excel-chartaxistitle-format-member)|指定图表坐标轴标题的格式。|
||[text](/javascript/api/excel/excel.chartaxistitle#excel-excel-chartaxistitle-text-member)|指定坐标轴标题。|
||[visible](/javascript/api/excel/excel.chartaxistitle#excel-excel-chartaxistitle-visible-member)|指定坐标轴标题是否可见。|
|[ChartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|[font](/javascript/api/excel/excel.chartaxistitleformat#excel-excel-chartaxistitleformat-font-member)|指定图表坐标轴标题的字体属性，如图表坐标轴标题对象的字体名称、字号或颜色。|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[添加 (类型：Excel。ChartType，sourceData：Range，seriesBy？： Excel。ChartSeriesBy) ](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-add-member(1))|创建新图表。|
||[count](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-count-member)|返回工作表中的图表数。|
||[getItem(name: string)](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-getitem-member(1))|使用图表名称获取图表。|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-getitemat-member(1))|根据其在集合中的位置获取图表。|
||[items](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-items-member)|获取此集合中已加载的子项。|
|[ChartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|[fill](/javascript/api/excel/excel.chartdatalabelformat#excel-excel-chartdatalabelformat-fill-member)|表示当前图表数据标签的填充格式。|
||[font](/javascript/api/excel/excel.chartdatalabelformat#excel-excel-chartdatalabelformat-font-member)|表示图表数据 (字体名称、字号和颜色) 字体属性。|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[format](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-format-member)|指定图表数据标签的格式，包括填充和字体格式。|
||[position](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-position-member)|表示数据标签的位置的值。|
||[分隔符](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-separator-member)|表示用于图表中数据标签的分隔符的字符串。|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-showbubblesize-member)|指定数据标签气泡大小是否可见。|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-showcategoryname-member)|指定数据标签类别名称是否可见。|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-showlegendkey-member)|指定数据标签图例项标示是否可见。|
||[showPercentage](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-showpercentage-member)|指定数据标签百分比是否可见。|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-showseriesname-member)|指定数据标签系列名称是否可见。|
||[showValue](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-showvalue-member)|指定数据标签值是否可见。|
|[ChartFill](/javascript/api/excel/excel.chartfill)|[clear()](/javascript/api/excel/excel.chartfill#excel-excel-chartfill-clear-member(1))|清除图表元素的填充颜色。|
||[setSolidColor(color: string)](/javascript/api/excel/excel.chartfill#excel-excel-chartfill-setsolidcolor-member(1))|将图表元素的填充格式设置为统一颜色。|
|[ChartFont](/javascript/api/excel/excel.chartfont)|[bold](/javascript/api/excel/excel.chartfont#excel-excel-chartfont-bold-member)|表示字体的加粗状态。|
||[color](/javascript/api/excel/excel.chartfont#excel-excel-chartfont-color-member)|文本颜色格式的 HTML 颜色代码表示 (例如，#FF0000红色) 。|
||[italic](/javascript/api/excel/excel.chartfont#excel-excel-chartfont-italic-member)|表示字体的斜体状态。|
||[name](/javascript/api/excel/excel.chartfont#excel-excel-chartfont-name-member)|字体名称 (例如"Calibri") |
||[size](/javascript/api/excel/excel.chartfont#excel-excel-chartfont-size-member)|字体大小 (例如 11) |
||[underline](/javascript/api/excel/excel.chartfont#excel-excel-chartfont-underline-member)|应用于字体的下划线类型。|
|[ChartGridlines](/javascript/api/excel/excel.chartgridlines)|[format](/javascript/api/excel/excel.chartgridlines#excel-excel-chartgridlines-format-member)|表示图表网格线的格式。|
||[visible](/javascript/api/excel/excel.chartgridlines#excel-excel-chartgridlines-visible-member)|指定坐标轴网格线是否可见。|
|[ChartGridlinesFormat](/javascript/api/excel/excel.chartgridlinesformat)|[line](/javascript/api/excel/excel.chartgridlinesformat#excel-excel-chartgridlinesformat-line-member)|表示图表线条格式。|
|[ChartLegend](/javascript/api/excel/excel.chartlegend)|[format](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-format-member)|表示图表图例的格式，包括填充和字体格式。|
||[重叠](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-overlay-member)|指定图表图例是否与图表的主主体重叠。|
||[position](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-position-member)|指定图例在图表上的位置。|
||[visible](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-visible-member)|指定图表图例是否可见。|
|[ChartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|[fill](/javascript/api/excel/excel.chartlegendformat#excel-excel-chartlegendformat-fill-member)|表示对象的填充格式，包括背景格式信息。|
||[font](/javascript/api/excel/excel.chartlegendformat#excel-excel-chartlegendformat-font-member)|表示字体属性，如图表图例的字体名称、字号和颜色。|
|[ChartLineFormat](/javascript/api/excel/excel.chartlineformat)|[clear()](/javascript/api/excel/excel.chartlineformat#excel-excel-chartlineformat-clear-member(1))|清除图表元素的线条格式。|
||[color](/javascript/api/excel/excel.chartlineformat#excel-excel-chartlineformat-color-member)|表示图表中的线条颜色的 HTML 颜色代码。|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|[format](/javascript/api/excel/excel.chartpoint#excel-excel-chartpoint-format-member)|封装图表点的格式属性。|
||[value](/javascript/api/excel/excel.chartpoint#excel-excel-chartpoint-value-member)|返回图表点的值。|
|[ChartPointFormat](/javascript/api/excel/excel.chartpointformat)|[fill](/javascript/api/excel/excel.chartpointformat#excel-excel-chartpointformat-fill-member)|表示图表的填充格式，其中包括背景格式信息。|
|[ChartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|[count](/javascript/api/excel/excel.chartpointscollection#excel-excel-chartpointscollection-count-member)|返回系列中的图表点数。|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartpointscollection#excel-excel-chartpointscollection-getitemat-member(1))|根据其在系列中的位置检索点。|
||[items](/javascript/api/excel/excel.chartpointscollection#excel-excel-chartpointscollection-items-member)|获取此集合中已加载的子项。|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[format](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-format-member)|表示图表系列的格式，包括填充和线条格式。|
||[name](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-name-member)|指定图表中系列的名称。|
||[points](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-points-member)|返回系列中所有数据点的集合。|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[count](/javascript/api/excel/excel.chartseriescollection#excel-excel-chartseriescollection-count-member)|返回集合中的系列数量。|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartseriescollection#excel-excel-chartseriescollection-getitemat-member(1))|根据其在集合中的位置检索系列|
||[items](/javascript/api/excel/excel.chartseriescollection#excel-excel-chartseriescollection-items-member)|获取此集合中已加载的子项。|
|[ChartSeriesFormat](/javascript/api/excel/excel.chartseriesformat)|[fill](/javascript/api/excel/excel.chartseriesformat#excel-excel-chartseriesformat-fill-member)|表示图表系列的填充格式，包括背景格式信息。|
||[line](/javascript/api/excel/excel.chartseriesformat#excel-excel-chartseriesformat-line-member)|表示线条格式。|
|[ChartTitle](/javascript/api/excel/excel.charttitle)|[format](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-format-member)|表示图表标题的格式，包括填充和字体格式。|
||[重叠](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-overlay-member)|指定图表标题是否覆盖图表。|
||[text](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-text-member)|指定图表的标题文本。|
||[visible](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-visible-member)|指定图表标题是否可见。|
|[ChartTitleFormat](/javascript/api/excel/excel.charttitleformat)|[fill](/javascript/api/excel/excel.charttitleformat#excel-excel-charttitleformat-fill-member)|表示对象的填充格式，包括背景格式信息。|
||[font](/javascript/api/excel/excel.charttitleformat#excel-excel-charttitleformat-font-member)|代表对象的字体 (字体名称、字号和颜色) 属性。|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[getRange()](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-getrange-member(1))|返回与名称相关的 range 对象。|
||[name](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-name-member)|对象的名称。|
||[type](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-type-member)|指定名称的公式返回的值的类型。|
||[value](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-value-member)|表示 name 公式计算出的值。|
||[visible](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-visible-member)|指定对象是否可见。|
|[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection)|[getItem(name: string)](/javascript/api/excel/excel.nameditemcollection#excel-excel-nameditemcollection-getitem-member(1))|使用对象 `NamedItem` 的名称获取对象。|
||[items](/javascript/api/excel/excel.nameditemcollection#excel-excel-nameditemcollection-items-member)|获取此集合中已加载的子项。|
|[范围](/javascript/api/excel/excel.range)|[address](/javascript/api/excel/excel.range#excel-excel-range-address-member)|指定 A1 样式的范围引用。|
||[addressLocal](/javascript/api/excel/excel.range#excel-excel-range-addresslocal-member)|表示以用户语言表示指定区域的范围引用。|
||[cellCount](/javascript/api/excel/excel.range#excel-excel-range-cellcount-member)|指定区域中的单元格数。|
||[clear(applyTo?: Excel.ClearApplyTo)](/javascript/api/excel/excel.range#excel-excel-range-clear-member(1))|清除区域值、格式、填充、边框等。|
||[columnCount](/javascript/api/excel/excel.range#excel-excel-range-columncount-member)|指定范围中的列总数。|
||[columnIndex](/javascript/api/excel/excel.range#excel-excel-range-columnindex-member)|指定范围中第一个单元格的列号。|
||[删除 (班次：Excel。DeleteShiftDirection) ](/javascript/api/excel/excel.range#excel-excel-range-delete-member(1))|删除与区域相关的单元格。|
||[format](/javascript/api/excel/excel.range#excel-excel-range-format-member)|返回一个格式对象，其中封装了区域的字体、填充、边框、对齐方式和其他属性。|
||[formulas](/javascript/api/excel/excel.range#excel-excel-range-formulas-member)|表示采用 A1 表示法的公式。|
||[formulasLocal](/javascript/api/excel/excel.range#excel-excel-range-formulaslocal-member)|表示采用 A1 样式表示法的公式，使用用户的语言和数字格式区域设置。|
||[getBoundingRect (anotherRange： Range \| string) ](/javascript/api/excel/excel.range#excel-excel-range-getboundingrect-member(1))|获取包含指定区域的最小 range 对象。|
||[getCell(row: number, column: number)](/javascript/api/excel/excel.range#excel-excel-range-getcell-member(1))|根据行和列编号获取包含单个单元格的 range 对象。|
||[getColumn(column: number)](/javascript/api/excel/excel.range#excel-excel-range-getcolumn-member(1))|获取区域中包含的列。|
||[getEntireColumn()](/javascript/api/excel/excel.range#excel-excel-range-getentirecolumn-member(1))|获取一个对象，该对象代表区域区域整列 (例如，如果当前区域代表单元格"B4：E11" `getEntireColumn` ，则它是表示列"B：E") 。|
||[getEntireRow()](/javascript/api/excel/excel.range#excel-excel-range-getentirerow-member(1))|获取一个对象，该对象代表区域整行 (例如，如果当前区域代表单元格"B4：E11" `GetEntireRow` ，则其为表示行"4：11") 。|
||[getIntersection (anotherRange： Range \| string) ](/javascript/api/excel/excel.range#excel-excel-range-getintersection-member(1))|获取表示指定区域的矩形交集的 range 对象。|
||[getLastCell () ](/javascript/api/excel/excel.range#excel-excel-range-getlastcell-member(1))|获取区域内的最后一个单元格。|
||[getLastColumn () ](/javascript/api/excel/excel.range#excel-excel-range-getlastcolumn-member(1))|获取区域内的最后一列。|
||[getLastRow () ](/javascript/api/excel/excel.range#excel-excel-range-getlastrow-member(1))|获取区域内的最后一行。|
||[getOffsetRange(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.range#excel-excel-range-getoffsetrange-member(1))|获取表示与指定区域偏移的区域的对象。|
||[getRow(row: number)](/javascript/api/excel/excel.range#excel-excel-range-getrow-member(1))|获取范围中包含的行。|
||[插入 (班次：Excel。InsertShiftDirection) ](/javascript/api/excel/excel.range#excel-excel-range-insert-member(1))|将单个单元格或一系列单元格插入到工作表中取代此区域，并移动其他单元格以留出空间。|
||[numberFormat](/javascript/api/excel/excel.range#excel-excel-range-numberformat-member)|表示Excel区域的电话号码格式代码。|
||[rowCount](/javascript/api/excel/excel.range#excel-excel-range-rowcount-member)|返回区域中的总行数。|
||[rowIndex](/javascript/api/excel/excel.range#excel-excel-range-rowindex-member)|返回区域中第一个单元格的行编号。|
||[select()](/javascript/api/excel/excel.range#excel-excel-range-select-member(1))|在 Excel UI 中选择指定的区域。|
||[text](/javascript/api/excel/excel.range#excel-excel-range-text-member)|指定区域的文本值。|
||[valueTypes](/javascript/api/excel/excel.range#excel-excel-range-valuetypes-member)|指定每个单元格中数据类型。|
||[values](/javascript/api/excel/excel.range#excel-excel-range-values-member)|表示指定区域的原始值。|
||[worksheet](/javascript/api/excel/excel.range#excel-excel-range-worksheet-member)|包含当前区域的工作表。|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[color](/javascript/api/excel/excel.rangeborder#excel-excel-rangeborder-color-member)|HTML 颜色代码，表示边框线的颜色，格式为 #RRGGBB (例如"FFA500") ，或作为已命名的 HTML 颜色 (例如"orange") 。|
||[sideIndex](/javascript/api/excel/excel.rangeborder#excel-excel-rangeborder-sideindex-member)|指示边框的特定边的常量值。|
||[style](/javascript/api/excel/excel.rangeborder#excel-excel-rangeborder-style-member)|线条样式的常量之一，指定边框的线条样式。|
||[weight](/javascript/api/excel/excel.rangeborder#excel-excel-rangeborder-weight-member)|指定区域周围的边框的粗细。|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[count](/javascript/api/excel/excel.rangebordercollection#excel-excel-rangebordercollection-count-member)|集合中的 border 对象数量。|
||[getItem (索引：Excel。BorderIndex) ](/javascript/api/excel/excel.rangebordercollection#excel-excel-rangebordercollection-getitem-member(1))|使用其名称获取 border 对象|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangebordercollection#excel-excel-rangebordercollection-getitemat-member(1))|使用其索引获取 border 对象|
||[items](/javascript/api/excel/excel.rangebordercollection#excel-excel-rangebordercollection-items-member)|获取此集合中已加载的子项。|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[clear()](/javascript/api/excel/excel.rangefill#excel-excel-rangefill-clear-member(1))|重置区域背景。|
||[color](/javascript/api/excel/excel.rangefill#excel-excel-rangefill-color-member)|HTML 颜色代码，表示背景的颜色，格式为 #RRGGBB (例如"FFA500") ，或作为已命名的 HTML 颜色 (例如"orange") |
|[RangeFont](/javascript/api/excel/excel.rangefont)|[bold](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-bold-member)|表示字体的加粗状态。|
||[color](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-color-member)|文本颜色格式的 HTML 颜色代码表示 (例如，#FF0000红色) 。|
||[italic](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-italic-member)|指定字体的 italic 状态。|
||[name](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-name-member)|字体名称 (，例如"Calibri") 。|
||[size](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-size-member)|字号|
||[underline](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-underline-member)|应用于字体的下划线类型。|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[Borders](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-borders-member)|应用于整个区域的 Border 对象的集合。|
||[fill](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-fill-member)|返回在整个区域内定义的 fill 对象。|
||[font](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-font-member)|返回在整个区域内定义的 Font 对象。|
||[horizontalAlignment](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-horizontalalignment-member)|表示指定对象的水平对齐方式。|
||[verticalAlignment](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-verticalalignment-member)|表示指定对象的垂直对齐方式。|
||[wrapText](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-wraptext-member)|指定是否Excel对象中的文本换行。|
|[Table](/javascript/api/excel/excel.table)|[columns](/javascript/api/excel/excel.table#excel-excel-table-columns-member)|表示表中所有列的集合。|
||[delete()](/javascript/api/excel/excel.table#excel-excel-table-delete-member(1))|删除表。|
||[getDataBodyRange () ](/javascript/api/excel/excel.table#excel-excel-table-getdatabodyrange-member(1))|获取与表的数据体相关的 range 对象。|
||[getHeaderRowRange () ](/javascript/api/excel/excel.table#excel-excel-table-getheaderrowrange-member(1))|获取与表的标题行相关的 range 对象。|
||[getRange()](/javascript/api/excel/excel.table#excel-excel-table-getrange-member(1))|获取与整个表相关的 range 对象。|
||[getTotalRowRange () ](/javascript/api/excel/excel.table#excel-excel-table-gettotalrowrange-member(1))|获取与表的总计行相关的 range 对象。|
||[id](/javascript/api/excel/excel.table#excel-excel-table-id-member)|返回用于唯一标识指定工作簿中表的值。|
||[name](/javascript/api/excel/excel.table#excel-excel-table-name-member)|表的名称。|
||[rows](/javascript/api/excel/excel.table#excel-excel-table-rows-member)|表示表中所有行的集合。|
||[showHeaders](/javascript/api/excel/excel.table#excel-excel-table-showheaders-member)|指定标题行是否可见。|
||[showTotals](/javascript/api/excel/excel.table#excel-excel-table-showtotals-member)|指定总计行是否可见。|
||[style](/javascript/api/excel/excel.table#excel-excel-table-style-member)|代表表格样式的常量值。|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[add (address： Range \| string， hasHeaders： boolean) ](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-add-member(1))|创建一个新表。|
||[count](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-count-member)|返回工作簿中的表数目。|
||[getItem(key: string)](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-getitem-member(1))|按名称或 ID 获取表。|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-getitemat-member(1))|根据其在集合中的位置获取表。|
||[items](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-items-member)|获取此集合中已加载的子项。|
|[TableColumn](/javascript/api/excel/excel.tablecolumn)|[delete()](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-delete-member(1))|从表中删除列。|
||[getDataBodyRange () ](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-getdatabodyrange-member(1))|获取与列的数据体相关的 range 对象。|
||[getHeaderRowRange () ](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-getheaderrowrange-member(1))|获取与列的标头行相关的 range 对象。|
||[getRange()](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-getrange-member(1))|获取与整个列相关的 range 对象。|
||[getTotalRowRange () ](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-gettotalrowrange-member(1))|获取与列的总计行相关的 range 对象。|
||[id](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-id-member)|返回标识表内的列的唯一键。|
||[index](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-index-member)|返回表的列集合内列的索引编号。|
||[name](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-name-member)|指定表列的名称。|
||[values](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-values-member)|表示指定区域的原始值。|
|[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|[add (index？： number， values？： Array<Array<boolean \| string \| number>> \| boolean \| string \| number， name？： string) ](/javascript/api/excel/excel.tablecolumncollection#excel-excel-tablecolumncollection-add-member(1))|向表中添加新列。|
||[count](/javascript/api/excel/excel.tablecolumncollection#excel-excel-tablecolumncollection-count-member)|返回表中的列数。|
||[getItem (键：数字 \| 字符串) ](/javascript/api/excel/excel.tablecolumncollection#excel-excel-tablecolumncollection-getitem-member(1))|按名称或 ID 获取 column 对象。|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablecolumncollection#excel-excel-tablecolumncollection-getitemat-member(1))|根据其在集合中的位置获取列。|
||[items](/javascript/api/excel/excel.tablecolumncollection#excel-excel-tablecolumncollection-items-member)|获取此集合中已加载的子项。|
|[TableRow](/javascript/api/excel/excel.tablerow)|[delete()](/javascript/api/excel/excel.tablerow#excel-excel-tablerow-delete-member(1))|从表中删除行。|
||[getRange()](/javascript/api/excel/excel.tablerow#excel-excel-tablerow-getrange-member(1))|返回与整个行相关的 range 对象。|
||[index](/javascript/api/excel/excel.tablerow#excel-excel-tablerow-index-member)|返回表的行集合内行的索引编号。|
||[values](/javascript/api/excel/excel.tablerow#excel-excel-tablerow-values-member)|表示指定区域的原始值。|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[add (index？： number， values？： Array<Array<boolean \| string \| number>> \| boolean \| string \| number， alwaysInsert？： boolean) ](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-add-member(1))|向表中添加一行或多行。|
||[count](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-count-member)|返回表中的行数。|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-getitemat-member(1))|根据其在集合中的位置获取行。|
||[items](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-items-member)|获取此集合中已加载的子项。|
|[Workbook](/javascript/api/excel/excel.workbook)|[application](/javascript/api/excel/excel.workbook#excel-excel-workbook-application-member)|表示Excel工作簿的应用程序实例。|
||[bindings](/javascript/api/excel/excel.workbook#excel-excel-workbook-bindings-member)|表示属于工作簿的绑定的集合。|
||[getSelectedRange () ](/javascript/api/excel/excel.workbook#excel-excel-workbook-getselectedrange-member(1))|从工作簿获取当前选定的单个区域。|
||[names](/javascript/api/excel/excel.workbook#excel-excel-workbook-names-member)|表示工作簿范围的命名项集合， (范围和常量) 。|
||[表](/javascript/api/excel/excel.workbook#excel-excel-workbook-tables-member)|表示与工作簿关联的表的集合。|
||[worksheets](/javascript/api/excel/excel.workbook#excel-excel-workbook-worksheets-member)|表示与工作簿关联的工作表的集合。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[activate()](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-activate-member(1))|在 Excel UI 中激活工作表。|
||[charts](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-charts-member)|返回属于工作表的图表集合。|
||[delete()](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-delete-member(1))|从工作簿中删除工作表。|
||[getCell(row: number, column: number)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getcell-member(1))|`Range`基于行号和列号获取包含单个单元格的对象。|
||[getRange (address？： string) ](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getrange-member(1))|获取对象 `Range` ，该对象代表由地址或名称指定的单个单元格矩形块。|
||[id](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-id-member)|返回用于唯一标识指定工作簿中工作表的值。|
||[name](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-name-member)|工作表的显示名称。|
||[position](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-position-member)|工作表在工作簿中的位置，从零开始。|
||[表](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-tables-member)|属于工作表的表的集合。|
||[visibility](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-visibility-member)|工作表的可见性。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[添加 (名称？：string) ](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-add-member(1))|向工作簿添加新工作表。|
||[getActiveWorksheet () ](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-getactiveworksheet-member(1))|获取工作簿中当前处于活动状态的工作表。|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-getitem-member(1))|使用其名称或 ID 获取 worksheet 对象。|
||[items](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-items-member)|获取此集合中已加载的子项。|

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 参考文档](/javascript/api/excel?view=excel-js-1.1&preserve-view=true)
- [Excel JavaScript API 要求集](excel-api-requirement-sets.md)
