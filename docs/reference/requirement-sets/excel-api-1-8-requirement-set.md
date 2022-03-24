---
title: Excel JavaScript API 要求集 1.8
description: 有关 ExcelApi 1.8 要求集的详细信息。
ms.date: 03/19/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 39f3a5daf89849d3f8517794ab8cd4214309a667
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/23/2022
ms.locfileid: "63746859"
---
# <a name="whats-new-in-excel-javascript-api-18"></a>JavaScript API 1.8 Excel的新增功能

Excel JavaScript API 要求集 1.8 的功能包括适用于数据透视表、数据验证、图表、图表事件、性能选项和工作簿创建的 API。

## <a name="pivottable"></a>数据透视表

加载项通过数据透视表 API 的波形 2 设置数据透视表的层次结构。 现在可以控制数据及其聚合方式。 [数据透视表](../../excel/excel-add-ins-pivottables.md)一文详细介绍了新的数据透视表功能。

## <a name="data-validation"></a>数据有效性

数据有效性可以控制用户在工作表中输入的内容。 可以将单元格限制为预定义的答案集，或者在用户输入无效数据时提供弹出警告。 立即详细了解[向区域添加数据有效性](../../excel/excel-add-ins-data-validation.md)。

## <a name="charts"></a>图表

另一轮图表 API 可更好地对图表元素进行编程控制。 现在，你对图例、坐标轴、趋势线和绘图区拥有更高的访问权限。

## <a name="events"></a>事件

已为图表添加更多[事件](../../excel/excel-add-ins-events.md)。 让加载项处理用于与图表的交互。 此外，你还可以在整个工作簿中[触发事件](../../excel/performance.md#enable-and-disable-events)。

## <a name="api-list"></a>API 列表

下表列出了 JavaScript API 要求集 1.8 Excel中的 API。 若要查看受 Excel JavaScript API 要求集 1.8 或更早版本支持的所有 API 的 API 参考文档，请参阅[要求集 1.8](/javascript/api/excel?view=excel-js-1.8&preserve-view=true) 或更早中的 Excel API。

| 类 | 域 | 说明 |
|:---|:---|:---|
|[BasicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|[formula1](/javascript/api/excel/excel.basicdatavalidation#excel-excel-basicdatavalidation-formula1-member)|当 operator 属性设置为二进制运算符（如 GreaterThan (）时，指定右侧操作数 (左侧操作数是用户尝试在单元格单元格中输入的值) 。|
||[formula2](/javascript/api/excel/excel.basicdatavalidation#excel-excel-basicdatavalidation-formula2-member)|使用三元运算符 Between 和 NotBetween 指定上限操作数。|
||[operator](/javascript/api/excel/excel.basicdatavalidation#excel-excel-basicdatavalidation-operator-member)|用于验证数据有效性的运算符。|
|[Chart](/javascript/api/excel/excel.chart)|[categoryLabelLevel](/javascript/api/excel/excel.chart#excel-excel-chart-categorylabellevel-member)|指定图表类别标签级别枚举常量，该常量引用源分类标签的级别。|
||[displayBlanksAs](/javascript/api/excel/excel.chart#excel-excel-chart-displayblanksas-member)|指定在图表上绘制空白单元格的方式。|
||[onActivated](/javascript/api/excel/excel.chart#excel-excel-chart-onactivated-member)|激活图表时发生。|
||[onDeactivated](/javascript/api/excel/excel.chart#excel-excel-chart-ondeactivated-member)|在图表被停用时发生。|
||[plotArea](/javascript/api/excel/excel.chart#excel-excel-chart-plotarea-member)|表示图表的绘图区。|
||[plotBy](/javascript/api/excel/excel.chart#excel-excel-chart-plotby-member)|指定列或行在图表上用作数据系列的方式。|
||[plotVisibleOnly](/javascript/api/excel/excel.chart#excel-excel-chart-plotvisibleonly-member)|如果仅绘制可见单元格，则为 True。|
||[seriesNameLevel](/javascript/api/excel/excel.chart#excel-excel-chart-seriesnamelevel-member)|指定图表系列名称级别枚举常量，该常量引用源系列名称的级别。|
||[showDataLabelsOverMaximum](/javascript/api/excel/excel.chart#excel-excel-chart-showdatalabelsovermaximum-member)|指定当值大于数值轴上的最大值时是否显示数据标签。|
||[style](/javascript/api/excel/excel.chart#excel-excel-chart-style-member)|指定图表的图表样式。|
|[ChartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|[chartId](/javascript/api/excel/excel.chartactivatedeventargs#excel-excel-chartactivatedeventargs-chartid-member)|获取已激活图表的 ID。|
||[type](/javascript/api/excel/excel.chartactivatedeventargs#excel-excel-chartactivatedeventargs-type-member)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.chartactivatedeventargs#excel-excel-chartactivatedeventargs-worksheetid-member)|获取激活图表的工作表的 ID。|
|[ChartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|[chartId](/javascript/api/excel/excel.chartaddedeventargs#excel-excel-chartaddedeventargs-chartid-member)|获取添加到工作表的图表的 ID。|
||[source](/javascript/api/excel/excel.chartaddedeventargs#excel-excel-chartaddedeventargs-source-member)|获取事件源。|
||[type](/javascript/api/excel/excel.chartaddedeventargs#excel-excel-chartaddedeventargs-type-member)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.chartaddedeventargs#excel-excel-chartaddedeventargs-worksheetid-member)|获取添加图表的工作表的 ID。|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[alignment](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-alignment-member)|指定指定坐标轴刻度线标签的对齐方式。|
||[isBetweenCategories](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-isbetweencategories-member)|指定数值轴是否与分类之间的分类轴相交。|
||[multiLevel](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-multilevel-member)|指定坐标轴是否多级。|
||[numberFormat](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-numberformat-member)|指定坐标轴刻度线标签的格式代码。|
||[offset](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-offset-member)|指定标签级别之间的距离，以及第一级标签与轴线之间的距离。|
||[position](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-position-member)|指定两轴交叉的指定坐标轴位置。|
||[positionAt](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-positionat-member)|指定两轴交叉的轴位置。|
||[setPositionAt (值：number) ](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-setpositionat-member(1))|设置指定坐标轴与其他坐标轴交叉的位置。|
||[textOrientation](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-textorientation-member)|指定文本面向图表坐标轴刻度线标签的角度。|
|[ChartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|[fill](/javascript/api/excel/excel.chartaxisformat#excel-excel-chartaxisformat-fill-member)|指定图表填充格式。|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[setFormula (公式：string) ](/javascript/api/excel/excel.chartaxistitle#excel-excel-chartaxistitle-setformula-member(1))|该字符串值表示采用 A1 表示法的图表轴标题的公式。|
|[ChartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|[border](/javascript/api/excel/excel.chartaxistitleformat#excel-excel-chartaxistitleformat-border-member)|指定图表坐标轴标题的边框格式，包括颜色、线条和粗细。|
||[fill](/javascript/api/excel/excel.chartaxistitleformat#excel-excel-chartaxistitleformat-fill-member)|指定图表坐标轴标题的填充格式。|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[clear()](/javascript/api/excel/excel.chartborder#excel-excel-chartborder-clear-member(1))|清除图表元素的边框格式。|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[onActivated](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-onactivated-member)|激活图表时发生。|
||[onAdded](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-onadded-member)|在将新图表添加到工作表时发生。|
||[onDeactivated](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-ondeactivated-member)|停用图表时发生。|
||[onDeleted](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-ondeleted-member)|删除图表时发生。|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[autoText](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-autotext-member)|指定数据标签是否根据上下文自动生成相应的文本。|
||[format](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-format-member)|表示图表数据标签的格式。|
||[formula](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-formula-member)|该字符串值表示采用 A1 表示法的图表数据标签的公式。|
||[height](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-height-member)|返回图表数据标签的高度，以磅为单位。|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-horizontalalignment-member)|表示图表数据标签水平对齐。|
||[left](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-left-member)|表示图表数据标签左边缘到图表区域左边缘的距离，以磅为单位。|
||[numberFormat](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-numberformat-member)|该字符串值表示数据标签的格式代码。|
||[text](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-text-member)|该字符串表示图表上的数据标签文本。|
||[textOrientation](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-textorientation-member)|表示文本面向图表数据标签的角度。|
||[top](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-top-member)|表示图表数据标签上边缘到图表区域顶部的距离，以磅为单位。|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-verticalalignment-member)|表示图表数据标签垂直对齐。|
||[width](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-width-member)|返回图表数据标签的宽度，以磅为单位。|
|[ChartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|[border](/javascript/api/excel/excel.chartdatalabelformat#excel-excel-chartdatalabelformat-border-member)|表示边框格式，包括颜色、线条样式和粗细。|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[autoText](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-autotext-member)|指定数据标签是否根据上下文自动生成相应的文本。|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-horizontalalignment-member)|指定图表数据标签的水平对齐方式。|
||[numberFormat](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-numberformat-member)|指定数据标签的格式代码。|
||[textOrientation](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-textorientation-member)|表示文本面向数据标签的角度。|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-verticalalignment-member)|表示图表数据标签垂直对齐。|
|[ChartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|[chartId](/javascript/api/excel/excel.chartdeactivatedeventargs#excel-excel-chartdeactivatedeventargs-chartid-member)|获取已停用图表的 ID。|
||[type](/javascript/api/excel/excel.chartdeactivatedeventargs#excel-excel-chartdeactivatedeventargs-type-member)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.chartdeactivatedeventargs#excel-excel-chartdeactivatedeventargs-worksheetid-member)|获取在其中停用图表的工作表的 ID。|
|[ChartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|[chartId](/javascript/api/excel/excel.chartdeletedeventargs#excel-excel-chartdeletedeventargs-chartid-member)|获取从工作表中删除的图表的 ID。|
||[source](/javascript/api/excel/excel.chartdeletedeventargs#excel-excel-chartdeletedeventargs-source-member)|获取事件源。|
||[type](/javascript/api/excel/excel.chartdeletedeventargs#excel-excel-chartdeletedeventargs-type-member)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.chartdeletedeventargs#excel-excel-chartdeletedeventargs-worksheetid-member)|获取删除图表的工作表的 ID。|
|[ChartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|[height](/javascript/api/excel/excel.chartlegendentry#excel-excel-chartlegendentry-height-member)|指定图表图例上图例项的高度。|
||[index](/javascript/api/excel/excel.chartlegendentry#excel-excel-chartlegendentry-index-member)|指定图表图例中图例项的索引。|
||[left](/javascript/api/excel/excel.chartlegendentry#excel-excel-chartlegendentry-left-member)|指定图表图例项的左侧值。|
||[top](/javascript/api/excel/excel.chartlegendentry#excel-excel-chartlegendentry-top-member)|指定图表图例项的顶部。|
||[width](/javascript/api/excel/excel.chartlegendentry#excel-excel-chartlegendentry-width-member)|表示图表 Legend 上的图例项的宽度。|
|[ChartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|[border](/javascript/api/excel/excel.chartlegendformat#excel-excel-chartlegendformat-border-member)|表示边框格式，包括颜色、线条样式和粗细。|
|[ChartPlotArea](/javascript/api/excel/excel.chartplotarea)|[format](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-format-member)|指定图表绘图区的格式。|
||[height](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-height-member)|指定绘图区的高度值。|
||[insideHeight](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-insideheight-member)|指定绘图区的内部高度值。|
||[insideLeft](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-insideleft-member)|指定绘图区的左内值。|
||[insideTop](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-insidetop-member)|指定绘图区内部的顶部值。|
||[insideWidth](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-insidewidth-member)|指定绘图区的内部宽度值。|
||[left](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-left-member)|指定绘图区的左侧值。|
||[position](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-position-member)|指定绘图区的位置。|
||[top](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-top-member)|指定绘图区的顶部值。|
||[width](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-width-member)|指定绘图区的宽度值。|
|[ChartPlotAreaFormat](/javascript/api/excel/excel.chartplotareaformat)|[border](/javascript/api/excel/excel.chartplotareaformat#excel-excel-chartplotareaformat-border-member)|指定图表绘图区的边框属性。|
||[fill](/javascript/api/excel/excel.chartplotareaformat#excel-excel-chartplotareaformat-fill-member)|指定对象的填充格式，其中包括背景格式信息。|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[axisGroup](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-axisgroup-member)|指定指定系列的组。|
||[dataLabels](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-datalabels-member)|表示系列中所有数据标签的集合。|
||[explosion](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-explosion-member)|指定饼图或圆环图扇区的分解值。|
||[firstSliceAngle](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-firstsliceangle-member)|指定第一个饼图或圆环图扇区的角度，以度 (垂直方向顺时针) 。|
||[invertIfNegative](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-invertifnegative-member)|如此 Excel当项对应于一个负数时反转图案。|
||[overlap](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-overlap-member)|指定条柱的摆放方式。|
||[secondPlotSize](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-secondplotsize-member)|指定复合饼图或复合条饼图的第二部分的大小，以主饼图大小的百分比表示。|
||[splitType](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-splittype-member)|指定拆分复合饼图或复合条饼图的两部分的方式。|
||[varyByCategories](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-varybycategories-member)|如此 如果Excel为每个数据标记分配不同的颜色或图案。|
|[ChartTrendline](/javascript/api/excel/excel.charttrendline)|[backwardPeriod](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-backwardperiod-member)|表示趋势线向后延伸的周期数。|
||[forwardPeriod](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-forwardperiod-member)|表示趋势线向前延伸的周期数。|
||[label](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-label-member)|表示图表趋势线的标签。|
||[showEquation](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-showequation-member)|如果图表上显示趋势线公式，则为 True。|
||[showRSquared](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-showrsquared-member)|如此 如果趋势线的 r 平方值显示在图表上。|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[autoText](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-autotext-member)|指定趋势线标签是否根据上下文自动生成相应的文本。|
||[format](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-format-member)|图表趋势线标签的格式。|
||[formula](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-formula-member)|字符串值，表示使用 A1 样式表示法的图表趋势线标签的公式。|
||[height](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-height-member)|返回图表趋势线标签的高度，以磅为单位。|
||[horizontalAlignment](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-horizontalalignment-member)|表示图表趋势线标签的水平对齐方式。|
||[left](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-left-member)|表示从图表趋势线标签左边缘到图表区左边缘的距离（以点表示）。|
||[numberFormat](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-numberformat-member)|字符串值，表示趋势线标签的格式代码。|
||[text](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-text-member)|该字符串表示图表上的趋势线标签文本。|
||[textOrientation](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-textorientation-member)|表示文本面向图表趋势线标签的角度。|
||[top](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-top-member)|表示从图表趋势线标签的上边缘到图表区域顶部的距离（以点表示）。|
||[verticalAlignment](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-verticalalignment-member)|表示图表趋势线标签的垂直对齐方式。|
||[width](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-width-member)|返回图表趋势线标签的宽度，以磅为单位。|
|[ChartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|[border](/javascript/api/excel/excel.charttrendlinelabelformat#excel-excel-charttrendlinelabelformat-border-member)|指定边框格式，包括颜色、线条和粗细。|
||[fill](/javascript/api/excel/excel.charttrendlinelabelformat#excel-excel-charttrendlinelabelformat-fill-member)|指定当前图表趋势线标签的填充格式。|
||[font](/javascript/api/excel/excel.charttrendlinelabelformat#excel-excel-charttrendlinelabelformat-font-member)|指定图表趋势 (的字体属性，如字体名称) 字体大小和颜色类型。|
|[CustomDataValidation](/javascript/api/excel/excel.customdatavalidation)|[formula](/javascript/api/excel/excel.customdatavalidation#excel-excel-customdatavalidation-formula-member)|自定义数据验证公式。|
|[DataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|[field](/javascript/api/excel/excel.datapivothierarchy#excel-excel-datapivothierarchy-field-member)|返回与 DataPivotHierarchy 相关联的 PivotFields。|
||[id](/javascript/api/excel/excel.datapivothierarchy#excel-excel-datapivothierarchy-id-member)|DataPivotHierarchy 的 ID。|
||[name](/javascript/api/excel/excel.datapivothierarchy#excel-excel-datapivothierarchy-name-member)|DataPivotHierarchy 的名称。|
||[numberFormat](/javascript/api/excel/excel.datapivothierarchy#excel-excel-datapivothierarchy-numberformat-member)|DataPivotHierarchy 的数字格式。|
||[position](/javascript/api/excel/excel.datapivothierarchy#excel-excel-datapivothierarchy-position-member)|DataPivotHierarchy 的位置。|
||[setToDefault () ](/javascript/api/excel/excel.datapivothierarchy#excel-excel-datapivothierarchy-settodefault-member(1))|将 DataPivotHierarchy 重置回其默认值。|
||[showAs](/javascript/api/excel/excel.datapivothierarchy#excel-excel-datapivothierarchy-showas-member)|指定数据是否应显示为特定的摘要计算。|
||[summarizeBy](/javascript/api/excel/excel.datapivothierarchy#excel-excel-datapivothierarchy-summarizeby-member)|指定是否显示 DataPivotHierarchy 的所有项。|
|[DataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|[添加 (pivotHierarchy： Excel。PivotHierarchy) ](/javascript/api/excel/excel.datapivothierarchycollection#excel-excel-datapivothierarchycollection-add-member(1))|将 PivotHierarchy 添加到当前轴。|
||[getCount()](/javascript/api/excel/excel.datapivothierarchycollection#excel-excel-datapivothierarchycollection-getcount-member(1))|获取集合中的透视层级结构的数量。|
||[getItem(name: string)](/javascript/api/excel/excel.datapivothierarchycollection#excel-excel-datapivothierarchycollection-getitem-member(1))|按名称或 ID 获取 DataPivotHierarchy。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.datapivothierarchycollection#excel-excel-datapivothierarchycollection-getitemornullobject-member(1))|按名称获取 DataPivotHierarchy。|
||[items](/javascript/api/excel/excel.datapivothierarchycollection#excel-excel-datapivothierarchycollection-items-member)|获取此集合中已加载的子项。|
||[remove (DataPivotHierarchy： Excel。DataPivotHierarchy) ](/javascript/api/excel/excel.datapivothierarchycollection#excel-excel-datapivothierarchycollection-remove-member(1))|从当前轴删除 PivotHierarchy。|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[clear()](/javascript/api/excel/excel.datavalidation#excel-excel-datavalidation-clear-member(1))|清除当前区域中的数据有效性。|
||[errorAlert](/javascript/api/excel/excel.datavalidation#excel-excel-datavalidation-erroralert-member)|用户输入无效数据时，出现错误警报。|
||[ignoreBlanks](/javascript/api/excel/excel.datavalidation#excel-excel-datavalidation-ignoreblanks-member)|指定是否对空白单元格执行数据验证。|
||[prompt](/javascript/api/excel/excel.datavalidation#excel-excel-datavalidation-prompt-member)|当用户选择单元格时提示。|
||[rule](/javascript/api/excel/excel.datavalidation#excel-excel-datavalidation-rule-member)|包含不同类型的数据有效性条件的数据有效性规则。|
||[type](/javascript/api/excel/excel.datavalidation#excel-excel-datavalidation-type-member)|数据验证的类型，有关详细信息 `Excel.DataValidationType` ，请参阅。|
||[valid](/javascript/api/excel/excel.datavalidation#excel-excel-datavalidation-valid-member)|表示所有单元格值根据数据有效性规则是否全部有效。|
|[DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|[邮件](/javascript/api/excel/excel.datavalidationerroralert#excel-excel-datavalidationerroralert-message-member)|表示错误警报消息。|
||[showAlert](/javascript/api/excel/excel.datavalidationerroralert#excel-excel-datavalidationerroralert-showalert-member)|指定在用户输入无效数据时是否显示错误警报对话框。|
||[style](/javascript/api/excel/excel.datavalidationerroralert#excel-excel-datavalidationerroralert-style-member)|数据有效性警报类型，请参阅了解 `Excel.DataValidationAlertStyle` 详细信息。|
||[title](/javascript/api/excel/excel.datavalidationerroralert#excel-excel-datavalidationerroralert-title-member)|代表错误警报对话框标题。|
|[DataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)|[邮件](/javascript/api/excel/excel.datavalidationprompt#excel-excel-datavalidationprompt-message-member)|指定提示消息。|
||[showPrompt](/javascript/api/excel/excel.datavalidationprompt#excel-excel-datavalidationprompt-showprompt-member)|指定当用户选择具有数据有效性的单元格时是否显示提示。|
||[title](/javascript/api/excel/excel.datavalidationprompt#excel-excel-datavalidationprompt-title-member)|指定提示的标题。|
|[DataValidationRule](/javascript/api/excel/excel.datavalidationrule)|[custom](/javascript/api/excel/excel.datavalidationrule#excel-excel-datavalidationrule-custom-member)|自定义数据有效性条件。|
||[date](/javascript/api/excel/excel.datavalidationrule#excel-excel-datavalidationrule-date-member)|日期数据有效性条件。|
||[decimal](/javascript/api/excel/excel.datavalidationrule#excel-excel-datavalidationrule-decimal-member)|小数数据有效性条件。|
||[列表](/javascript/api/excel/excel.datavalidationrule#excel-excel-datavalidationrule-list-member)|列表数据有效性条件。|
||[textLength](/javascript/api/excel/excel.datavalidationrule#excel-excel-datavalidationrule-textlength-member)|文本长度数据有效性条件。|
||[time](/javascript/api/excel/excel.datavalidationrule#excel-excel-datavalidationrule-time-member)|时间数据有效性条件。|
||[wholeNumber](/javascript/api/excel/excel.datavalidationrule#excel-excel-datavalidationrule-wholenumber-member)|整个数字数据有效性条件。|
|[DateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|[formula1](/javascript/api/excel/excel.datetimedatavalidation#excel-excel-datetimedatavalidation-formula1-member)|当 operator 属性设置为二进制运算符（如 GreaterThan (）时，指定右侧操作数 (左侧操作数是用户尝试在单元格单元格中输入的值) 。|
||[formula2](/javascript/api/excel/excel.datetimedatavalidation#excel-excel-datetimedatavalidation-formula2-member)|使用三元运算符 Between 和 NotBetween 指定上限操作数。|
||[operator](/javascript/api/excel/excel.datetimedatavalidation#excel-excel-datetimedatavalidation-operator-member)|用于验证数据有效性的运算符。|
|[FilterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|[enableMultipleFilterItems](/javascript/api/excel/excel.filterpivothierarchy#excel-excel-filterpivothierarchy-enablemultiplefilteritems-member)|确定是否允许多个筛选项。|
||[fields](/javascript/api/excel/excel.filterpivothierarchy#excel-excel-filterpivothierarchy-fields-member)|返回与 FilterPivotHierarchy 相关联的 PivotFields。|
||[id](/javascript/api/excel/excel.filterpivothierarchy#excel-excel-filterpivothierarchy-id-member)|FilterPivotHierarchy 的 ID。|
||[name](/javascript/api/excel/excel.filterpivothierarchy#excel-excel-filterpivothierarchy-name-member)|FilterPivotHierarchy 的名称。|
||[position](/javascript/api/excel/excel.filterpivothierarchy#excel-excel-filterpivothierarchy-position-member)|FilterPivotHierarchy 的位置。|
||[setToDefault () ](/javascript/api/excel/excel.filterpivothierarchy#excel-excel-filterpivothierarchy-settodefault-member(1))|将 FilterPivotHierarchy 重置回其默认值。|
|[FilterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|[添加 (pivotHierarchy： Excel。PivotHierarchy) ](/javascript/api/excel/excel.filterpivothierarchycollection#excel-excel-filterpivothierarchycollection-add-member(1))|将 PivotHierarchy 添加到当前轴。|
||[getCount()](/javascript/api/excel/excel.filterpivothierarchycollection#excel-excel-filterpivothierarchycollection-getcount-member(1))|获取集合中的透视层级结构的数量。|
||[getItem(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection#excel-excel-filterpivothierarchycollection-getitem-member(1))|按名称或 ID 获取 FilterPivotHierarchy。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection#excel-excel-filterpivothierarchycollection-getitemornullobject-member(1))|按名称获取 FilterPivotHierarchy。|
||[items](/javascript/api/excel/excel.filterpivothierarchycollection#excel-excel-filterpivothierarchycollection-items-member)|获取此集合中已加载的子项。|
||[remove (filterPivotHierarchy： Excel。FilterPivotHierarchy) ](/javascript/api/excel/excel.filterpivothierarchycollection#excel-excel-filterpivothierarchycollection-remove-member(1))|从当前轴删除 PivotHierarchy。|
|[ListDataValidation](/javascript/api/excel/excel.listdatavalidation)|[inCellDropDown](/javascript/api/excel/excel.listdatavalidation#excel-excel-listdatavalidation-incelldropdown-member)|指定是否在单元格下拉列表中显示列表。|
||[source](/javascript/api/excel/excel.listdatavalidation#excel-excel-listdatavalidation-source-member)|数据有效性列表源|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[id](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-id-member)|透视字段的 ID。|
||[项目](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-items-member)|返回与 PivotField 关联的 PivotItems。|
||[name](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-name-member)|PivotField 的名称。|
||[showAllItems](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-showallitems-member)|确定是否显示 PivotField 的所有项。|
||[sortByLabels (sortBy： SortBy) ](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-sortbylabels-member(1))|PivotField 排序。|
||[subtotals](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-subtotals-member)|PivotField 小计。|
|[PivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|[getCount()](/javascript/api/excel/excel.pivotfieldcollection#excel-excel-pivotfieldcollection-getcount-member(1))|获取集合中透视字段的数量。|
||[getItem(name: string)](/javascript/api/excel/excel.pivotfieldcollection#excel-excel-pivotfieldcollection-getitem-member(1))|按名称或 ID 获取 PivotField。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotfieldcollection#excel-excel-pivotfieldcollection-getitemornullobject-member(1))|按名称获取 PivotField。|
||[items](/javascript/api/excel/excel.pivotfieldcollection#excel-excel-pivotfieldcollection-items-member)|获取此集合中已加载的子项。|
|[PivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|[fields](/javascript/api/excel/excel.pivothierarchy#excel-excel-pivothierarchy-fields-member)|返回与 PivotHierarchy 相关联的 PivotFields。|
||[id](/javascript/api/excel/excel.pivothierarchy#excel-excel-pivothierarchy-id-member)|PivotHierarchy 的 ID。|
||[name](/javascript/api/excel/excel.pivothierarchy#excel-excel-pivothierarchy-name-member)|PivotHierarchy 的名称。|
|[PivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|[getCount()](/javascript/api/excel/excel.pivothierarchycollection#excel-excel-pivothierarchycollection-getcount-member(1))|获取集合中的透视层级结构的数量。|
||[getItem(name: string)](/javascript/api/excel/excel.pivothierarchycollection#excel-excel-pivothierarchycollection-getitem-member(1))|按名称或 ID 获取 PivotHierarchy。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivothierarchycollection#excel-excel-pivothierarchycollection-getitemornullobject-member(1))|按名称获取 PivotHierarchy。|
||[items](/javascript/api/excel/excel.pivothierarchycollection#excel-excel-pivothierarchycollection-items-member)|获取此集合中已加载的子项。|
|[PivotItem](/javascript/api/excel/excel.pivotitem)|[id](/javascript/api/excel/excel.pivotitem#excel-excel-pivotitem-id-member)|PivotItem 的 ID。|
||[isExpanded](/javascript/api/excel/excel.pivotitem#excel-excel-pivotitem-isexpanded-member)|确定是展开项以显示子项还是折叠项并隐藏子项。|
||[name](/javascript/api/excel/excel.pivotitem#excel-excel-pivotitem-name-member)|PivotItem 的名称。|
||[visible](/javascript/api/excel/excel.pivotitem#excel-excel-pivotitem-visible-member)|指定 PivotItem 是否可见。|
|[PivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|[getCount()](/javascript/api/excel/excel.pivotitemcollection#excel-excel-pivotitemcollection-getcount-member(1))|获取集合中 PivotItems 的数量。|
||[getItem(name: string)](/javascript/api/excel/excel.pivotitemcollection#excel-excel-pivotitemcollection-getitem-member(1))|按名称或 ID 获取 PivotItem。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotitemcollection#excel-excel-pivotitemcollection-getitemornullobject-member(1))|按名称获取 PivotItem。|
||[items](/javascript/api/excel/excel.pivotitemcollection#excel-excel-pivotitemcollection-items-member)|获取此集合中已加载的子项。|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getColumnLabelRange () ](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-getcolumnlabelrange-member(1))|返回数据透视表列标签所在位置的区域。|
||[getDataBodyRange () ](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-getdatabodyrange-member(1))|返回数据透视表数据值所在位置的区域。|
||[getFilterAxisRange () ](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-getfilteraxisrange-member(1))|返回数据透视表筛选区的区域。|
||[getRange()](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-getrange-member(1))|返回存在数据透视表的区域，不包括筛选区。|
||[getRowLabelRange () ](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-getrowlabelrange-member(1))|返回数据透视表行标签所在位置的区域。|
||[layoutType](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-layouttype-member)|此属性指示数据透视表上的所有字段的 PivotLayoutType。|
||[showColumnGrandTotals](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-showcolumngrandtotals-member)|指定数据透视表是否显示列的总计。|
||[showRowGrandTotals](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-showrowgrandtotals-member)|指定数据透视表是否显示行的总计。|
||[subtotalLocation](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-subtotallocation-member)|此属性指示数据 `SubtotalLocationType` 透视表上所有字段的 。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[columnHierarchies](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-columnhierarchies-member)|数据透视表的列透视层级结构。|
||[dataHierarchies](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-datahierarchies-member)|数据透视表的数据透视层级结构。|
||[delete()](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-delete-member(1))|删除 PivotTable 对象。|
||[filterHierarchies](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-filterhierarchies-member)|数据透视表的筛选器透视层级结构。|
||[hierarchies](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-hierarchies-member)|数据透视表的透视层级结构。|
||[layout](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-layout-member)|PivotLayout，用于说明数据透视表的布局和可视化结构。|
||[rowHierarchies](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-rowhierarchies-member)|数据透视表的行透视层级结构。|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[add (name： string， source： Range \| string \| Table， destination： Range \| string) ](/javascript/api/excel/excel.pivottablecollection#excel-excel-pivottablecollection-add-member(1))|添加基于指定源数据的数据透视表，并将其插入到目标区域左上方的单元格中。|
|[范围](/javascript/api/excel/excel.range)|[dataValidation](/javascript/api/excel/excel.range#excel-excel-range-datavalidation-member)|返回数据有效性对象。|
|[RowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|[fields](/javascript/api/excel/excel.rowcolumnpivothierarchy#excel-excel-rowcolumnpivothierarchy-fields-member)|返回与 RowColumnPivotHierarchy 相关联的 PivotFields。|
||[id](/javascript/api/excel/excel.rowcolumnpivothierarchy#excel-excel-rowcolumnpivothierarchy-id-member)|RowColumnPivotHierarchy 的 ID。|
||[name](/javascript/api/excel/excel.rowcolumnpivothierarchy#excel-excel-rowcolumnpivothierarchy-name-member)|RowColumnPivotHierarchy 的名称。|
||[position](/javascript/api/excel/excel.rowcolumnpivothierarchy#excel-excel-rowcolumnpivothierarchy-position-member)|RowColumnPivotHierarchy 的位置。|
||[setToDefault () ](/javascript/api/excel/excel.rowcolumnpivothierarchy#excel-excel-rowcolumnpivothierarchy-settodefault-member(1))|将 RowColumnPivotHierarchy 重置回其默认值。|
|[RowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|[添加 (pivotHierarchy： Excel。PivotHierarchy) ](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#excel-excel-rowcolumnpivothierarchycollection-add-member(1))|将 PivotHierarchy 添加到当前轴。|
||[getCount()](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#excel-excel-rowcolumnpivothierarchycollection-getcount-member(1))|获取集合中的透视层级结构的数量。|
||[getItem(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#excel-excel-rowcolumnpivothierarchycollection-getitem-member(1))|按名称或 ID 获取 RowColumnPivotHierarchy。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#excel-excel-rowcolumnpivothierarchycollection-getitemornullobject-member(1))|按名称获取 RowColumnPivotHierarchy。|
||[items](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#excel-excel-rowcolumnpivothierarchycollection-items-member)|获取此集合中已加载的子项。|
||[remove (rowColumnPivotHierarchy： Excel。RowColumnPivotHierarchy) ](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#excel-excel-rowcolumnpivothierarchycollection-remove-member(1))|从当前轴删除 PivotHierarchy。|
|[运行时](/javascript/api/excel/excel.runtime)|[enableEvents](/javascript/api/excel/excel.runtime#excel-excel-runtime-enableevents-member)|在当前任务窗格或内容加载项中切换 JavaScript 事件。|
|[ShowAsRule](/javascript/api/excel/excel.showasrule)|[baseField](/javascript/api/excel/excel.showasrule#excel-excel-showasrule-basefield-member)|要基于的透视 `ShowAs` 字段（如果适用，则根据 `ShowAsCalculation` 类型）进行计算，否则 `null`为 。|
||[baseItem](/javascript/api/excel/excel.showasrule#excel-excel-showasrule-baseitem-member)|要基于计算依据 `ShowAs` 的项（如果适用，则根据 `ShowAsCalculation` 类型）进行计算，否则 `null`为 。|
||[calculation](/javascript/api/excel/excel.showasrule#excel-excel-showasrule-calculation-member)|用于 `ShowAs` 透视字段的计算。|
|[Style](/javascript/api/excel/excel.style)|[autoIndent](/javascript/api/excel/excel.style#excel-excel-style-autoindent-member)|指定当单元格中的文本对齐方式设置为相等分布时文本是否自动缩进。|
||[textOrientation](/javascript/api/excel/excel.style#excel-excel-style-textorientation-member)|此样式中的文本方向。|
|[Subtotals](/javascript/api/excel/excel.subtotals)|[automatic](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-automatic-member)|如果 `Automatic` 设置为 `true`，则设置 时将忽略所有其他值 `Subtotals`。|
||[average](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-average-member)||
||[count](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-count-member)||
||[countNumbers](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-countnumbers-member)||
||[max](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-max-member)||
||[min](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-min-member)||
||[product](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-product-member)||
||[standardDeviation](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-standarddeviation-member)||
||[standardDeviationP](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-standarddeviationp-member)||
||[sum](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-sum-member)||
||[variance](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-variance-member)||
||[varianceP](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-variancep-member)||
|[Table](/javascript/api/excel/excel.table)|[legacyId](/javascript/api/excel/excel.table#excel-excel-table-legacyid-member)|返回数字 ID。|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.tablechangedeventargs#excel-excel-tablechangedeventargs-getrange-member(1))|获取表示特定工作表上表格的已更改区域的区域。|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.tablechangedeventargs#excel-excel-tablechangedeventargs-getrangeornullobject-member(1))|获取表示特定工作表上表格的已更改区域的区域。|
|[Workbook](/javascript/api/excel/excel.workbook)|[readOnly](/javascript/api/excel/excel.workbook#excel-excel-workbook-readonly-member)|如果 `true` 工作簿以只读模式打开，则返回 。|
|[WorkbookCreated](/javascript/api/excel/excel.workbookcreated)|||
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onCalculated](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-oncalculated-member)|计算工作表时发生。|
||[showGridlines](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-showgridlines-member)|指定网格线是否对用户可见。|
||[showHeadings](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-showheadings-member)|指定标题是否对用户可见。|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[type](/javascript/api/excel/excel.worksheetcalculatedeventargs#excel-excel-worksheetcalculatedeventargs-type-member)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.worksheetcalculatedeventargs#excel-excel-worksheetcalculatedeventargs-worksheetid-member)|获取进行计算的工作表的 ID。|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-getrange-member(1))|获取区域，该区域表示特定工作表上的更改区域。|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-getrangeornullobject-member(1))|获取区域，该区域表示特定工作表上的更改区域。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onCalculated](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-oncalculated-member)|计算工作簿中任何工作表时发生。|

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 参考文档](/javascript/api/excel?view=excel-js-1.8&preserve-view=true)
- [Excel JavaScript API 要求集](excel-api-requirement-sets.md)
