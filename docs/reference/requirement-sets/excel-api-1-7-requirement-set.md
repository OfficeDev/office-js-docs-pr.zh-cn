---
title: Excel JavaScript API 要求集 1.7
description: 有关 ExcelApi 1.7 要求集的详细信息。
ms.date: 11/09/2020
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: cd8f0f333b76306a6feecff95b9ba8831428606a
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744531"
---
# <a name="whats-new-in-excel-javascript-api-17"></a>Excel JavaScript API 1.7 的最近更新

Excel JavaScript API 要求集 1.7 的功能包括用于图表、事件、工作表、区域、文档属性、已命名项目、保护选项和样式的 API。

## <a name="customize-charts"></a>自定义图表

通过新的图表 API，你可以创建其他图表类型、向图表中添加数据系列、设置图表标题、添加轴标题、添加显示单位、添加采用移动平均值的趋势线、将趋势线更改为线性趋势线等。 以下是一些示例。

- 图表轴 - 获取、设置、格式化和删除图表中的轴单位、标签和标题。
- 图表系列 - 添加、设置和删除图表中的某个系列。  更改系列标记、绘制顺序和大小。
- 图表趋势线 - 添加、获取和格式化图表中的趋势线。
- 图表图例 - 设置图表中的图例字体的格式。
- 图表点 - 设置图表点颜色。
- 图表标题子字符串 - 获取和设置图表的标题子字符串。
- 图表类型 - 用于创建更多图表类型的选项。

## <a name="events"></a>事件

Excel 事件 API 提供了多个事件处理程序，以便加载项能够在发生特定事件时自动运行指定的函数。 可以将函数设计为执行方案所需的任何操作。 有关当前可用的事件列表，请参阅[使用 Excel JavaScript API 处理事件](../../excel/excel-add-ins-events.md)。

## <a name="customize-the-appearance-of-worksheets-and-ranges"></a>自定义工作表和区域的外观

使用新的 API 可以通过多种方式自定义工作表的外观：

- 冻结窗格，使特定行或列在你滚动工作表时保持可见。 例如，如果工作表中的第一行包含标题，则可以冻结此行，以便在你向下滚动工作表时列标题保持可见。
- 修改工作表标签颜色。
- 添加工作表标题。

可以通过多种方式自定义区域的外观：

- 设置某个区域的单元格样式，确保该区域内的所有单元格采用一致的格式。 单元格 样式是一组定义的格式特征，例如字体和字号、数字格式、单元格边框和单元格底纹。 使用 Excel 中的任意内置单元格样式，或者使用自己的自定义单元格样式。
- 设置区域的文本方向。
- 添加或修改区域上链接至工作表中的其他位置或外部位置的超链接。

## <a name="manage-document-properties"></a>管理文档属性

使用文档属性 API，你可以访问内置文档属性，并且还可以创建和管理自定义文档属性，以存储工作表的状态和驱动工作流和业务逻辑。

## <a name="copy-worksheets"></a>复制工作表

使用工作表复制 APIs，你可以将一个工作表中的数据和格式复制到相同工作簿中的另一个工作表，从而减少所需的数据传输量。

## <a name="handle-ranges-with-ease"></a>轻松地处理区域

使用各种区域 API，你可以完成诸如获取周围区域、获取大小经过重设的区域之类的任务。 这些 API 可以显著提高诸如区域操作和寻址之类任务的效率。

此外：

- 工作簿和工作表保护选项 - 使用这些 API 可保护工作表和工作簿结构中的数据。
- 更新已命名项目 - 使用此 API 可更新已命名项目。
- 获取活动单元格 - 使用此 API 可获取工作表中的活动单元格。

## <a name="api-list"></a>API 列表

下表列出了 JavaScript API 要求集 1.7 Excel中的 API。 若要查看受 Excel JavaScript API 要求集 1.7 或更早版本支持的所有 API 的 API 参考文档，请参阅[要求集 1.7](/javascript/api/excel?view=excel-js-1.7&preserve-view=true) 或更早中的 Excel API。

| 类 | 域 | 说明 |
|:---|:---|:---|
|[Chart](/javascript/api/excel/excel.chart)|[chartType](/javascript/api/excel/excel.chart#excel-excel-chart-charttype-member)|指定图表的类型。|
||[id](/javascript/api/excel/excel.chart#excel-excel-chart-id-member)|图表的唯一 ID。|
||[showAllFieldButtons](/javascript/api/excel/excel.chart#excel-excel-chart-showallfieldbuttons-member)|指定是否显示项目上的所有字段数据透视图。|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[border](/javascript/api/excel/excel.chartareaformat#excel-excel-chartareaformat-border-member)|表示图表区的边框格式，包括颜色、线条和粗细。|
|[ChartAxes](/javascript/api/excel/excel.chartaxes)|[getItem (类型：Excel。ChartAxisType， group？： Excel。ChartAxisGroup) ](/javascript/api/excel/excel.chartaxes#excel-excel-chartaxes-getitem-member(1))|返回通过类型和组标识的特定轴。|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[axisGroup](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-axisgroup-member)|指定指定坐标轴的组。|
||[baseTimeUnit](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-basetimeunit-member)|指定指定分类轴的基本单位。|
||[categoryType](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-categorytype-member)|指定分类轴类型。|
||[customDisplayUnit](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-customdisplayunit-member)|指定自定义轴显示单位值。|
||[displayUnit](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-displayunit-member)|表示轴显示单位。|
||[height](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-height-member)|指定图表轴的高度（以点表示）。|
||[left](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-left-member)|指定从坐标轴左边缘到图表区左侧的距离（以点表示）。|
||[logBase](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-logbase-member)|指定使用对数刻度时对数的底数。|
||[majorTickMark](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-majortickmark-member)|指定指定坐标轴的主要刻度线类型。|
||[majorTimeUnitScale](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-majortimeunitscale-member)|当属性设置为 时，指定分类 `categoryType` 轴的主要单位刻度值 `dateAxis`。|
||[minorTickMark](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-minortickmark-member)|指定指定坐标轴的次要刻度线类型。|
||[minorTimeUnitScale](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-minortimeunitscale-member)|当属性设置为 时 `categoryType` ，指定分类轴的次要单位刻度值 `dateAxis`。|
||[reversePlotOrder](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-reverseplotorder-member)|指定是否Excel从最后一个到第一个数据点。|
||[scaleType](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-scaletype-member)|指定数值轴刻度类型。|
||[setCategoryNames (sourceData：Range) ](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-setcategorynames-member(1))|设置指定轴的所有分类名称。|
||[setCustomDisplayUnit (值：number) ](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-setcustomdisplayunit-member(1))|将轴显示单位设为自定义值。|
||[showDisplayUnitLabel](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-showdisplayunitlabel-member)|指定坐标轴显示单位标签是否可见。|
||[tickLabelPosition](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-ticklabelposition-member)|在指定坐标轴上指定刻度线标签的位置。|
||[tickLabelSpacing](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-ticklabelspacing-member)|指定刻度线标签之间的分类数或系列数。|
||[tickMarkSpacing](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-tickmarkspacing-member)|指定刻度线之间的分类数或系列数。|
||[top](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-top-member)|指定从坐标轴上边缘到图表区顶部的距离（以点表示）。|
||[type](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-type-member)|指定坐标轴类型。|
||[visible](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-visible-member)|指定坐标轴是否可见。|
||[width](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-width-member)|指定图表轴的宽度（以点表示）。|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[color](/javascript/api/excel/excel.chartborder#excel-excel-chartborder-color-member)|表示图表中的边框颜色的 HTML 颜色代码。|
||[lineStyle](/javascript/api/excel/excel.chartborder#excel-excel-chartborder-linestyle-member)|表示边框的线条样式。|
||[weight](/javascript/api/excel/excel.chartborder#excel-excel-chartborder-weight-member)|表示边框的粗细，以磅为单位。|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[position](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-position-member)|表示数据标签的位置的值。|
||[分隔符](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-separator-member)|该字符串表示用于图表中数据标签的分隔符。|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-showbubblesize-member)|指定数据标签气泡大小是否可见。|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-showcategoryname-member)|指定数据标签类别名称是否可见。|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-showlegendkey-member)|指定数据标签图例项标示是否可见。|
||[showPercentage](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-showpercentage-member)|指定数据标签百分比是否可见。|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-showseriesname-member)|指定数据标签系列名称是否可见。|
||[showValue](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-showvalue-member)|指定数据标签值是否可见。|
|[ChartFormatString](/javascript/api/excel/excel.chartformatstring)|[font](/javascript/api/excel/excel.chartformatstring#excel-excel-chartformatstring-font-member)|表示字体属性，如字体名称、字体大小和图表字符对象的颜色。|
|[ChartLegend](/javascript/api/excel/excel.chartlegend)|[height](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-height-member)|指定图表上图例的高度（以点表示）。|
||[left](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-left-member)|指定图表上图例的左值（以点表示）。|
||[legendEntries](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-legendentries-member)|表示图例中 legendEntries 的集合。|
||[showShadow](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-showshadow-member)|指定图例在图表上是否具有阴影。|
||[top](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-top-member)|指定图表图例的顶部。|
||[width](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-width-member)|指定图表上的图例的宽度（以点表示）。|
|[ChartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|[visible](/javascript/api/excel/excel.chartlegendentry#excel-excel-chartlegendentry-visible-member)|表示图表图例项的可见性。|
|[ChartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|[getCount()](/javascript/api/excel/excel.chartlegendentrycollection#excel-excel-chartlegendentrycollection-getcount-member(1))|返回集合中的图例项数。|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartlegendentrycollection#excel-excel-chartlegendentrycollection-getitemat-member(1))|返回给定索引位置的图例项。|
||[items](/javascript/api/excel/excel.chartlegendentrycollection#excel-excel-chartlegendentrycollection-items-member)|获取此集合中已加载的子项。|
|[ChartLineFormat](/javascript/api/excel/excel.chartlineformat)|[lineStyle](/javascript/api/excel/excel.chartlineformat#excel-excel-chartlineformat-linestyle-member)|代表线条样式。|
||[weight](/javascript/api/excel/excel.chartlineformat#excel-excel-chartlineformat-weight-member)|表示线条的粗细（以磅为单位）。|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|[dataLabel](/javascript/api/excel/excel.chartpoint#excel-excel-chartpoint-datalabel-member)|返回图表点的数据标签。|
||[hasDataLabel](/javascript/api/excel/excel.chartpoint#excel-excel-chartpoint-hasdatalabel-member)|表示一个数据点是否具有数据标签。|
||[markerBackgroundColor](/javascript/api/excel/excel.chartpoint#excel-excel-chartpoint-markerbackgroundcolor-member)|数据点的数据标记背景色的 HTML 颜色代码表示 (例如，#FF0000红色) 。|
||[markerForegroundColor](/javascript/api/excel/excel.chartpoint#excel-excel-chartpoint-markerforegroundcolor-member)|数据点标记前景色的 HTML 颜色代码表示 (例如，#FF0000红色) 。|
||[markerSize](/javascript/api/excel/excel.chartpoint#excel-excel-chartpoint-markersize-member)|表示数据点的标记大小。|
||[markerStyle](/javascript/api/excel/excel.chartpoint#excel-excel-chartpoint-markerstyle-member)|表示图表数据点的标记样式。|
|[ChartPointFormat](/javascript/api/excel/excel.chartpointformat)|[border](/javascript/api/excel/excel.chartpointformat#excel-excel-chartpointformat-border-member)|表示图表数据点的边框格式，包括颜色、样式和粗细信息。|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[chartType](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-charttype-member)|表示系列的图表类型。|
||[delete()](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-delete-member(1))|删除 chart series 对象。|
||[doughnutHoleSize](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-doughnutholesize-member)|表示图表系列的圆环孔大小。|
||[filtered](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-filtered-member)|指定是否筛选系列。|
||[gapWidth](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gapwidth-member)|表示图表系列的间隙宽度。|
||[hasDataLabels](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-hasdatalabels-member)|指定系列是否具有数据标签。|
||[markerBackgroundColor](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-markerbackgroundcolor-member)|指定图表系列的标记背景色。|
||[markerForegroundColor](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-markerforegroundcolor-member)|指定图表系列的标记前景色。|
||[markerSize](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-markersize-member)|指定图表系列的标记大小。|
||[markerStyle](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-markerstyle-member)|指定图表系列的标记样式。|
||[plotOrder](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-plotorder-member)|指定图表组中图表系列的绘制顺序。|
||[setBubbleSizes (sourceData：Range) ](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-setbubblesizes-member(1))|设置图表系列的气泡大小。|
||[setValues (sourceData：Range) ](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-setvalues-member(1))|设置图表系列的值。|
||[setXAxisValues (sourceData：Range) ](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-setxaxisvalues-member(1))|设置图表系列的 x 轴的值。|
||[showShadow](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-showshadow-member)|指定系列是否具有阴影。|
||[平滑](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-smooth-member)|指定系列是否平滑。|
||[trendlines](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-trendlines-member)|系列中趋势线的集合。|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[add (name？： string， index？： number) ](/javascript/api/excel/excel.chartseriescollection#excel-excel-chartseriescollection-add-member(1))|向集合添加新系列。|
|[ChartTitle](/javascript/api/excel/excel.charttitle)|[getSubstring (start： number， length： number) ](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-getsubstring-member(1))|获取图表标题的子字符串。|
||[height](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-height-member)|返回图表标题的高度，以磅为单位。|
||[horizontalAlignment](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-horizontalalignment-member)|指定图表标题的水平对齐方式。|
||[left](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-left-member)|指定图表标题左边缘到图表区域左边缘的距离（以点表示）。|
||[position](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-position-member)|表示图表标题的位置。|
||[setFormula (公式：string) ](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-setformula-member(1))|设置一个字符串值，用于表示采用 A1 表示法的图表标题的公式。|
||[showShadow](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-showshadow-member)|表示一个布尔值，用于确定图表标题是否具有阴影。|
||[textOrientation](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-textorientation-member)|指定文本面向图表标题的角度。|
||[top](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-top-member)|指定图表标题的上边缘到图表区域顶部的距离（以点表示）。|
||[verticalAlignment](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-verticalalignment-member)|指定图表标题的垂直对齐方式。|
||[width](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-width-member)|指定图表标题的宽度（以点表示）。|
|[ChartTitleFormat](/javascript/api/excel/excel.charttitleformat)|[border](/javascript/api/excel/excel.charttitleformat#excel-excel-charttitleformat-border-member)|表示图表标题的边框格式，包括颜色、线条和粗细。|
|[ChartTrendline](/javascript/api/excel/excel.charttrendline)|[delete()](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-delete-member(1))|删除 Trendline 对象。|
||[format](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-format-member)|表示图表趋势线的格式。|
||[intercept](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-intercept-member)|表示趋势线的截距值。|
||[movingAveragePeriod](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-movingaverageperiod-member)|表示图表趋势线的周期。|
||[name](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-name-member)|表示趋势线的名称。|
||[polynomialOrder](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-polynomialorder-member)|表示图表趋势线的顺序。|
||[type](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-type-member)|表示图表趋势线的类型。|
|[ChartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|[添加 (类型？：Excel。ChartTrendlineType) ](/javascript/api/excel/excel.charttrendlinecollection#excel-excel-charttrendlinecollection-add-member(1))|向趋势线集合添加新的趋势线。|
||[getCount()](/javascript/api/excel/excel.charttrendlinecollection#excel-excel-charttrendlinecollection-getcount-member(1))|返回集合中的趋势线数量。|
||[getItem(index: number)](/javascript/api/excel/excel.charttrendlinecollection#excel-excel-charttrendlinecollection-getitem-member(1))|按索引获取趋势线对象，即项数组中的插入顺序。|
||[items](/javascript/api/excel/excel.charttrendlinecollection#excel-excel-charttrendlinecollection-items-member)|获取此集合中已加载的子项。|
|[ChartTrendlineFormat](/javascript/api/excel/excel.charttrendlineformat)|[line](/javascript/api/excel/excel.charttrendlineformat#excel-excel-charttrendlineformat-line-member)|表示图表线条格式。|
|[CustomProperty](/javascript/api/excel/excel.customproperty)|[delete()](/javascript/api/excel/excel.customproperty#excel-excel-customproperty-delete-member(1))|删除 custom property 对象。|
||[key](/javascript/api/excel/excel.customproperty#excel-excel-customproperty-key-member)|自定义属性的键。|
||[type](/javascript/api/excel/excel.customproperty#excel-excel-customproperty-type-member)|用于自定义属性的值的类型。|
||[value](/javascript/api/excel/excel.customproperty#excel-excel-customproperty-value-member)|自定义属性的值。|
|[CustomPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|[add (key： string， value： any) ](/javascript/api/excel/excel.custompropertycollection#excel-excel-custompropertycollection-add-member(1))|新建自定义属性或设置现有自定义属性。|
||[deleteAll () ](/javascript/api/excel/excel.custompropertycollection#excel-excel-custompropertycollection-deleteall-member(1))|删除此集合中的所有自定义属性。|
||[getCount()](/javascript/api/excel/excel.custompropertycollection#excel-excel-custompropertycollection-getcount-member(1))|获取自定义属性的计数。|
||[getItem(key: string)](/javascript/api/excel/excel.custompropertycollection#excel-excel-custompropertycollection-getitem-member(1))|按键获取自定义属性对象（不区分大小写）。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.custompropertycollection#excel-excel-custompropertycollection-getitemornullobject-member(1))|按键获取自定义属性对象（不区分大小写）。|
||[items](/javascript/api/excel/excel.custompropertycollection#excel-excel-custompropertycollection-items-member)|获取此集合中已加载的子项。|
|[DataConnectionCollection](/javascript/api/excel/excel.dataconnectioncollection)|[refreshAll () ](/javascript/api/excel/excel.dataconnectioncollection#excel-excel-dataconnectioncollection-refreshall-member(1))|刷新集合中所有数据连接。|
|[DocumentProperties](/javascript/api/excel/excel.documentproperties)|[author](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-author-member)|工作簿的作者。|
||[category](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-category-member)|工作簿的类别。|
||[comments](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-comments-member)|工作簿的注释。|
||[company](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-company-member)|工作簿的公司。|
||[creationDate](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-creationdate-member)|获取工作簿的创建日期。|
||[custom](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-custom-member)|获取工作簿的自定义属性的集合。|
||[keywords](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-keywords-member)|工作簿的关键字。|
||[lastAuthor](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-lastauthor-member)|获取工作簿的最终作者。|
||[manager](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-manager-member)|工作簿的管理者。|
||[revisionNumber](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-revisionnumber-member)|获取工作簿的修订号。|
||[subject](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-subject-member)|工作簿的主题。|
||[title](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-title-member)|工作簿的标题。|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[arrayValues](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-arrayvalues-member)|返回包含已命名项目的值和类型的对象。|
||[formula](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-formula-member)|已命名项目的公式。|
|[NamedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|[types](/javascript/api/excel/excel.nameditemarrayvalues#excel-excel-nameditemarrayvalues-types-member)|表示已命名项目数组中每个项目的类型|
||[values](/javascript/api/excel/excel.nameditemarrayvalues#excel-excel-nameditemarrayvalues-values-member)|表示已命名项目数组中每个项目的值。|
|[范围](/javascript/api/excel/excel.range)|[getAbsoluteResizedRange (numRows： number， numColumns： number) ](/javascript/api/excel/excel.range#excel-excel-range-getabsoluteresizedrange-member(1))|获取一 `Range` 个对象，该对象的左上 `Range` 单元格与当前对象相同，但具有指定的行数和列数。|
||[getImage () ](/javascript/api/excel/excel.range#excel-excel-range-getimage-member(1))|将区域呈现为 base64 编码 png 图像。|
||[getSurroundingRegion () ](/javascript/api/excel/excel.range#excel-excel-range-getsurroundingregion-member(1))|返回一 `Range` 个对象，该对象代表此区域左上单元格的周围区域。|
||[hyperlink](/javascript/api/excel/excel.range#excel-excel-range-hyperlink-member)|表示当前范围的超链接。|
||[isEntireColumn](/javascript/api/excel/excel.range#excel-excel-range-isentirecolumn-member)|表示当前区域是否为整列。|
||[isEntireRow](/javascript/api/excel/excel.range#excel-excel-range-isentirerow-member)|表示当前区域是否为整行。|
||[numberFormatLocal](/javascript/api/excel/excel.range#excel-excel-range-numberformatlocal-member)|表示Excel用户的语言设置，指定区域的电话号码格式代码。|
||[showCard () ](/javascript/api/excel/excel.range#excel-excel-range-showcard-member(1))|显示活动单元格的卡片（如果该单元格具有富值内容）。|
||[style](/javascript/api/excel/excel.range#excel-excel-range-style-member)|表示当前区域的样式。|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[textOrientation](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-textorientation-member)|区域内所有单元格的文本方向。|
||[useStandardHeight](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-usestandardheight-member)|确定对象的行高 `Range` 是否等于工作表的标准高度。|
||[useStandardWidth](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-usestandardwidth-member)|指定对象的列宽 `Range` 是否等于工作表的标准宽度。|
|[RangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|[address](/javascript/api/excel/excel.rangehyperlink#excel-excel-rangehyperlink-address-member)|表示超链接的 URL 目标。|
||[documentReference](/javascript/api/excel/excel.rangehyperlink#excel-excel-rangehyperlink-documentreference-member)|表示超链接的文档引用目标。|
||[screenTip](/javascript/api/excel/excel.rangehyperlink#excel-excel-rangehyperlink-screentip-member)|表示鼠标悬停在超链接上时显示的字符串。|
||[textToDisplay](/javascript/api/excel/excel.rangehyperlink#excel-excel-rangehyperlink-texttodisplay-member)|表示区域最左上方单元格中显示的字符串。|
|[Style](/javascript/api/excel/excel.style)|[Borders](/javascript/api/excel/excel.style#excel-excel-style-borders-member)|四个 border 对象的集合，这些对象代表四个边框的样式。|
||[builtIn](/javascript/api/excel/excel.style#excel-excel-style-builtin-member)|指定样式是否内置样式。|
||[delete()](/javascript/api/excel/excel.style#excel-excel-style-delete-member(1))|删除此样式。|
||[fill](/javascript/api/excel/excel.style#excel-excel-style-fill-member)|样式的填充。|
||[font](/javascript/api/excel/excel.style#excel-excel-style-font-member)|一 `Font` 个代表样式字体的对象。|
||[formulaHidden](/javascript/api/excel/excel.style#excel-excel-style-formulahidden-member)|指定在工作表受保护时是否隐藏公式。|
||[horizontalAlignment](/javascript/api/excel/excel.style#excel-excel-style-horizontalalignment-member)|表示样式水平对齐。|
||[includeAlignment](/javascript/api/excel/excel.style#excel-excel-style-includealignment-member)|指定样式是否包括自动缩进、水平对齐、垂直对齐、自动换行、缩进级别和文本方向属性。|
||[includeBorder](/javascript/api/excel/excel.style#excel-excel-style-includeborder-member)|指定样式是否包括颜色、颜色索引、线条样式和粗边框属性。|
||[includeFont](/javascript/api/excel/excel.style#excel-excel-style-includefont-member)|指定样式是否包括背景、粗体、颜色、颜色索引、字体样式、italic、名称、大小、删除线、下标、上标和下划线字体属性。|
||[includeNumber](/javascript/api/excel/excel.style#excel-excel-style-includenumber-member)|指定样式是否包含数字格式属性。|
||[includePatterns](/javascript/api/excel/excel.style#excel-excel-style-includepatterns-member)|指定样式是否包括颜色、颜色索引、反转（如果为负值）、图案、图案颜色和图案颜色索引内部属性。|
||[includeProtection](/javascript/api/excel/excel.style#excel-excel-style-includeprotection-member)|指定样式是否包含隐藏和锁定保护属性的公式。|
||[indentLevel](/javascript/api/excel/excel.style#excel-excel-style-indentlevel-member)|0 到 250 之间的一个整数，指示样式的缩进水平。|
||[locked](/javascript/api/excel/excel.style#excel-excel-style-locked-member)|指定在工作表受保护时对象是否被锁定。|
||[name](/javascript/api/excel/excel.style#excel-excel-style-name-member)|样式的名称。|
||[numberFormat](/javascript/api/excel/excel.style#excel-excel-style-numberformat-member)|样式中数字格式的格式代码。|
||[numberFormatLocal](/javascript/api/excel/excel.style#excel-excel-style-numberformatlocal-member)|样式中数字格式的本地化格式代码。|
||[readingOrder](/javascript/api/excel/excel.style#excel-excel-style-readingorder-member)|样式中的阅读顺序。|
||[shrinkToFit](/javascript/api/excel/excel.style#excel-excel-style-shrinktofit-member)|指定文本是否自动缩小以适应可用列宽。|
||[verticalAlignment](/javascript/api/excel/excel.style#excel-excel-style-verticalalignment-member)|指定样式的垂直对齐方式。|
||[wrapText](/javascript/api/excel/excel.style#excel-excel-style-wraptext-member)|指定是否Excel对象中的文本换行。|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[add(name: string)](/javascript/api/excel/excel.stylecollection#excel-excel-stylecollection-add-member(1))|向集合添加新样式。|
||[getItem(name: string)](/javascript/api/excel/excel.stylecollection#excel-excel-stylecollection-getitem-member(1))|按名称 `Style` 获取 。|
||[items](/javascript/api/excel/excel.stylecollection#excel-excel-stylecollection-items-member)|获取此集合中已加载的子项。|
|[Table](/javascript/api/excel/excel.table)|[onChanged](/javascript/api/excel/excel.table#excel-excel-table-onchanged-member)|特定表格上的单元格数据发生更改时发生。|
||[onSelectionChanged](/javascript/api/excel/excel.table#excel-excel-table-onselectionchanged-member)|当特定表格上的所选内容发生更改时发生。|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[address](/javascript/api/excel/excel.tablechangedeventargs#excel-excel-tablechangedeventargs-address-member)|获取地址，该地址表示特定工作表上的表格的更改区域。|
||[changeType](/javascript/api/excel/excel.tablechangedeventargs#excel-excel-tablechangedeventargs-changetype-member)|获取更改类型，该类型表示如何触发更改事件。|
||[source](/javascript/api/excel/excel.tablechangedeventargs#excel-excel-tablechangedeventargs-source-member)|获取事件源。|
||[tableId](/javascript/api/excel/excel.tablechangedeventargs#excel-excel-tablechangedeventargs-tableid-member)|获取其中的数据发生更改的表的 ID。|
||[type](/javascript/api/excel/excel.tablechangedeventargs#excel-excel-tablechangedeventargs-type-member)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.tablechangedeventargs#excel-excel-tablechangedeventargs-worksheetid-member)|获取其中的数据发生更改的工作表的 ID。|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onChanged](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-onchanged-member)|在工作簿或工作表中任何表上的数据发生更改时发生。|
|[TableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|[address](/javascript/api/excel/excel.tableselectionchangedeventargs#excel-excel-tableselectionchangedeventargs-address-member)|获取区域地址，该地址表示特定工作表上的表格选定区域。|
||[isInsideTable](/javascript/api/excel/excel.tableselectionchangedeventargs#excel-excel-tableselectionchangedeventargs-isinsidetable-member)|指定所选内容是否位于表格内部。|
||[tableId](/javascript/api/excel/excel.tableselectionchangedeventargs#excel-excel-tableselectionchangedeventargs-tableid-member)|获取所选内容发生更改的表的 ID。|
||[type](/javascript/api/excel/excel.tableselectionchangedeventargs#excel-excel-tableselectionchangedeventargs-type-member)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.tableselectionchangedeventargs#excel-excel-tableselectionchangedeventargs-worksheetid-member)|获取选定内容发生更改的工作表的 ID。|
|[Workbook](/javascript/api/excel/excel.workbook)|[dataConnections](/javascript/api/excel/excel.workbook#excel-excel-workbook-dataconnections-member)|表示工作簿中所有数据连接。|
||[getActiveCell () ](/javascript/api/excel/excel.workbook#excel-excel-workbook-getactivecell-member(1))|获取工作簿中当前处于活动状态的单元格。|
||[name](/javascript/api/excel/excel.workbook#excel-excel-workbook-name-member)|获取工作簿名称。|
||[properties](/javascript/api/excel/excel.workbook#excel-excel-workbook-properties-member)|获取工作簿属性。|
||[protection](/javascript/api/excel/excel.workbook#excel-excel-workbook-protection-member)|返回工作簿的保护对象。|
||[styles](/javascript/api/excel/excel.workbook#excel-excel-workbook-styles-member)|表示与工作簿关联的样式的集合。|
|[WorkbookProtection](/javascript/api/excel/excel.workbookprotection)|[保护 (密码？：字符串) ](/javascript/api/excel/excel.workbookprotection#excel-excel-workbookprotection-protect-member(1))|保护工作簿。|
||[protected](/javascript/api/excel/excel.workbookprotection#excel-excel-workbookprotection-protected-member)|指定工作簿是否受保护。|
||[不 (密码？：string) ](/javascript/api/excel/excel.workbookprotection#excel-excel-workbookprotection-unprotect-member(1))|解除保护工作簿。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[copy (positionType？： Excel。WorksheetPositionType， relativeTo？： Excel。工作表) ](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-copy-member(1))|复制工作表，并放置于指定位置。|
||[freezePanes](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-freezepanes-member)|获取一个对象，该对象可用于处理工作表上的冻结窗格。|
||[getRangeByIndexes (startRow： number， startColumn： number， rowCount： number， columnCount： number) ](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getrangebyindexes-member(1))|获取从 `Range` 特定行索引和列索引开始并跨越一定数量的行和列的对象。|
||[onActivated](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onactivated-member)|在激活工作表时发生。|
||[onChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onchanged-member)|特定工作表中的数据发生更改时发生。|
||[onDeactivated](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-ondeactivated-member)|在工作表被停用时发生。|
||[onSelectionChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onselectionchanged-member)|当特定工作表上的选择更改时发生。|
||[standardHeight](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-standardheight-member)|返回工作表中所有行的标准（默认）行高，以磅为单位。|
||[standardWidth](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-standardwidth-member)|指定工作表中 (列) 列的默认列宽。|
||[tabColor](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-tabcolor-member)|工作表的选项卡颜色。|
|[WorksheetActivatedEventArgs](/javascript/api/excel/excel.worksheetactivatedeventargs)|[type](/javascript/api/excel/excel.worksheetactivatedeventargs#excel-excel-worksheetactivatedeventargs-type-member)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.worksheetactivatedeventargs#excel-excel-worksheetactivatedeventargs-worksheetid-member)|获取已激活工作表的 ID。|
|[WorksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|[source](/javascript/api/excel/excel.worksheetaddedeventargs#excel-excel-worksheetaddedeventargs-source-member)|获取事件源。|
||[type](/javascript/api/excel/excel.worksheetaddedeventargs#excel-excel-worksheetaddedeventargs-type-member)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.worksheetaddedeventargs#excel-excel-worksheetaddedeventargs-worksheetid-member)|获取添加到工作簿的工作表的 ID。|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[address](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-address-member)|获取区域地址，该地址表示特定工作表上的更改区域。|
||[changeType](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-changetype-member)|获取更改类型，该类型表示如何触发更改事件。|
||[source](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-source-member)|获取事件源。|
||[type](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-type-member)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-worksheetid-member)|获取其中的数据发生更改的工作表的 ID。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onActivated](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onactivated-member)|在工作簿中任何工作表被激活时发生。|
||[onAdded](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onadded-member)|在将新工作表添加到工作簿时发生。|
||[onDeactivated](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-ondeactivated-member)|在工作簿中任何工作表被停用时发生。|
||[onDeleted](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-ondeleted-member)|从工作簿中删除工作表时发生。|
|[WorksheetDeactivatedEventArgs](/javascript/api/excel/excel.worksheetdeactivatedeventargs)|[type](/javascript/api/excel/excel.worksheetdeactivatedeventargs#excel-excel-worksheetdeactivatedeventargs-type-member)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.worksheetdeactivatedeventargs#excel-excel-worksheetdeactivatedeventargs-worksheetid-member)|获取已停用的工作表的 ID。|
|[WorksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|[source](/javascript/api/excel/excel.worksheetdeletedeventargs#excel-excel-worksheetdeletedeventargs-source-member)|获取事件源。|
||[type](/javascript/api/excel/excel.worksheetdeletedeventargs#excel-excel-worksheetdeletedeventargs-type-member)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.worksheetdeletedeventargs#excel-excel-worksheetdeletedeventargs-worksheetid-member)|获取从工作簿中删除的工作表的 ID。|
|[WorksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|[freezeAt (frozenRange：Range \| string) ](/javascript/api/excel/excel.worksheetfreezepanes#excel-excel-worksheetfreezepanes-freezeat-member(1))|设置活动工作表视图中的冻结单元格。|
||[freezeColumns (count？： number) ](/javascript/api/excel/excel.worksheetfreezepanes#excel-excel-worksheetfreezepanes-freezecolumns-member(1))|就地冻结工作表的第一列。|
||[freezeRows (count？： number) ](/javascript/api/excel/excel.worksheetfreezepanes#excel-excel-worksheetfreezepanes-freezerows-member(1))|就地冻结工作表的首行。|
||[getLocation()](/javascript/api/excel/excel.worksheetfreezepanes#excel-excel-worksheetfreezepanes-getlocation-member(1))|获取用于描述活动工作表视图中的冻结单元格的区域。|
||[getLocationOrNullObject () ](/javascript/api/excel/excel.worksheetfreezepanes#excel-excel-worksheetfreezepanes-getlocationornullobject-member(1))|获取用于描述活动工作表视图中的冻结单元格的区域。|
||[unfreeze () ](/javascript/api/excel/excel.worksheetfreezepanes#excel-excel-worksheetfreezepanes-unfreeze-member(1))|移除工作表中的所有冻结窗格。|
|[WorksheetProtection](/javascript/api/excel/excel.worksheetprotection)|[不 (密码？：string) ](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-unprotect-member(1))|解除对 worksheet 的保护。|
|[WorksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|[allowEditObjects](/javascript/api/excel/excel.worksheetprotectionoptions#excel-excel-worksheetprotectionoptions-alloweditobjects-member)|表示允许编辑对象的工作表保护选项。|
||[allowEditScenarios](/javascript/api/excel/excel.worksheetprotectionoptions#excel-excel-worksheetprotectionoptions-alloweditscenarios-member)|表示允许编辑方案的工作表保护选项。|
||[selectionMode](/javascript/api/excel/excel.worksheetprotectionoptions#excel-excel-worksheetprotectionoptions-selectionmode-member)|表示选择模式的工作表保护选项。|
|[WorksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|[address](/javascript/api/excel/excel.worksheetselectionchangedeventargs#excel-excel-worksheetselectionchangedeventargs-address-member)|获取区域地址，该地址表示特定工作表上的选定区域。|
||[type](/javascript/api/excel/excel.worksheetselectionchangedeventargs#excel-excel-worksheetselectionchangedeventargs-type-member)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.worksheetselectionchangedeventargs#excel-excel-worksheetselectionchangedeventargs-worksheetid-member)|获取选定内容发生更改的工作表的 ID。|

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 参考文档](/javascript/api/excel?view=excel-js-1.7&preserve-view=true)
- [Excel JavaScript API 要求集](excel-api-requirement-sets.md)
