---
title: Excel JavaScript API 要求集 1.9
description: 有关 ExcelApi 1.9 要求集的详细信息。
ms.date: 04/01/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: f34b109f95f013cf27f0abfca9c2a8c6b1e4e7c9
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/23/2022
ms.locfileid: "63746698"
---
# <a name="whats-new-in-excel-javascript-api-19"></a>JavaScript API 1.9 Excel的新增功能

超过 500 个新  Excel API 随 1.9 要求集一起推出。 第一个表提供了 API 的简明摘要，而后续表提供了详细列表。

| 功能区域 | 说明 | 相关对象 |
|:--- |:--- |:--- |
| [Shape](../../excel/excel-add-ins-shapes.md) | 插入、定位和格式化图像、几何形状和文本框。 | [ShapeCollection](/javascript/api/excel/excel.shapecollection) [Shape](/javascript/api/excel/excel.shape) [GeometricShape](/javascript/api/excel/excel.geometricshape)  [Image](/javascript/api/excel/excel.image) |
| [自动筛选](../../excel/excel-add-ins-worksheets.md#filter-data) | 为区域添加筛选器。 | [AutoFilter](/javascript/api/excel/excel.autofilter) |
| [Areas](../../excel/excel-add-ins-multiple-ranges.md) | 支持非连续区域。 | [RangeAreas](/javascript/api/excel/excel.rangeareas) |
| [特殊单元格](../../excel/excel-add-ins-multiple-ranges.md#get-special-cells-from-multiple-ranges) | 获取在区域内包含日期、备注或公式的单元格。 | [Range](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-)|
| [查找](../../excel/excel-add-ins-ranges-string-match.md) | 查找区域或工作表中的值或公式。 | [Range](/javascript/api/excel/excel.range#find-text--criteria-)[Worksheet](/javascript/api/excel/excel.worksheet#findall-text--criteria-) |
| [复制和粘贴](../../excel/excel-add-ins-ranges-cut-copy-paste.md) | 将值、格式和公式从一个区域复制到另一个区域。 | [Range](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-) |
| [Calculation](../../excel/performance.md#suspend-calculation-temporarily) | 更好地控制 Excel 计算引擎。 | [应用程序](/javascript/api/excel/excel.application) |
| 新图表 | 了解我们支持的新图表类型：地图、箱形图、瀑布图、旭日图、排列图 和漏斗图。 | [Chart](/javascript/api/excel/excel.charttype) |
| RangeFormat | 新功能及区域格式。 | [Range](/javascript/api/excel/excel.rangeformat) |

## <a name="api-list"></a>API 列表

下表列出了 JavaScript API 要求Excel集 1.9 中的 API。 若要查看受 Excel JavaScript API 要求集 1.9 或更早版本支持的所有 API 的 API 参考文档，请参阅[要求集 1.9](/javascript/api/excel?view=excel-js-1.9&preserve-view=true) 或更早中的 Excel API。

| 类 | 域 | 说明 |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[calculationEngineVersion](/javascript/api/excel/excel.application#excel-excel-application-calculationengineversion-member)|返回用于上次完整重新计算的 Excel 计算引擎版本。|
||[calculationState](/javascript/api/excel/excel.application#excel-excel-application-calculationstate-member)|返回应用程序的计算状态。|
||[iterativeCalculation](/javascript/api/excel/excel.application#excel-excel-application-iterativecalculation-member)|返回迭代计算设置。|
||[suspendScreenUpdatingUntilNextSync()](/javascript/api/excel/excel.application#excel-excel-application-suspendscreenupdatinguntilnextsync-member(1))|暂停屏幕更新，直到调用下一 `context.sync()` 个。|
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[apply(range: Range \| string, columnIndex?: number, criteria?: Excel.FilterCriteria)](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-apply-member(1))|将自动筛选器应用于区域。|
||[clearCriteria()](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-clearcriteria-member(1))|清除自动筛选的筛选条件及排序状态。|
||[criteria](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-criteria-member)|在自动筛选区域中保留所有筛选条件的数组。|
||[enabled](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-enabled-member)|指定是否启用自动筛选。|
||[getRange()](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-getrange-member(1))|`Range`返回一个对象，该对象代表应用自动筛选的范围。|
||[getRangeOrNullObject()](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-getrangeornullobject-member(1))|`Range`返回一个对象，该对象代表应用自动筛选的范围。|
||[isDataFiltered](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-isdatafiltered-member)|指定自动筛选是否具有筛选条件。|
||[reapply()](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-reapply-member(1))|应用当前位于区域上的指定 Autofilter 对象。|
||[remove()](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-remove-member(1))|删除区域的自动筛选。|
|[CellBorder](/javascript/api/excel/excel.cellborder)|[color](/javascript/api/excel/excel.cellborder#excel-excel-cellborder-color-member)|表示`color`单个边框的属性。|
||[style](/javascript/api/excel/excel.cellborder#excel-excel-cellborder-style-member)|表示`style`单个边框的属性。|
||[tintAndShade](/javascript/api/excel/excel.cellborder#excel-excel-cellborder-tintandshade-member)|表示`tintAndShade`单个边框的属性。|
||[weight](/javascript/api/excel/excel.cellborder#excel-excel-cellborder-weight-member)|表示`weight`单个边框的属性。|
|[CellBorderCollection](/javascript/api/excel/excel.cellbordercollection)|[bottom](/javascript/api/excel/excel.cellbordercollection#excel-excel-cellbordercollection-bottom-member)|表示`format.borders.bottom`属性。|
||[diagonalDown](/javascript/api/excel/excel.cellbordercollection#excel-excel-cellbordercollection-diagonaldown-member)|表示`format.borders.diagonalDown`属性。|
||[diagonalUp](/javascript/api/excel/excel.cellbordercollection#excel-excel-cellbordercollection-diagonalup-member)|表示`format.borders.diagonalUp`属性。|
||[horizontal](/javascript/api/excel/excel.cellbordercollection#excel-excel-cellbordercollection-horizontal-member)|表示`format.borders.horizontal`属性。|
||[left](/javascript/api/excel/excel.cellbordercollection#excel-excel-cellbordercollection-left-member)|表示`format.borders.left`属性。|
||[right](/javascript/api/excel/excel.cellbordercollection#excel-excel-cellbordercollection-right-member)|表示`format.borders.right`属性。|
||[top](/javascript/api/excel/excel.cellbordercollection#excel-excel-cellbordercollection-top-member)|表示`format.borders.top`属性。|
||[vertical](/javascript/api/excel/excel.cellbordercollection#excel-excel-cellbordercollection-vertical-member)|表示`format.borders.vertical`属性。|
|[CellProperties](/javascript/api/excel/excel.cellproperties)|[address](/javascript/api/excel/excel.cellproperties#excel-excel-cellproperties-address-member)|表示`address`属性。|
||[addressLocal](/javascript/api/excel/excel.cellproperties#excel-excel-cellproperties-addresslocal-member)|表示`addressLocal`属性。|
||[hidden](/javascript/api/excel/excel.cellproperties#excel-excel-cellproperties-hidden-member)|表示`hidden`属性。|
|[CellPropertiesFill](/javascript/api/excel/excel.cellpropertiesfill)|[color](/javascript/api/excel/excel.cellpropertiesfill#excel-excel-cellpropertiesfill-color-member)|表示`format.fill.color`属性。|
||[pattern](/javascript/api/excel/excel.cellpropertiesfill#excel-excel-cellpropertiesfill-pattern-member)|表示`format.fill.pattern`属性。|
||[patternColor](/javascript/api/excel/excel.cellpropertiesfill#excel-excel-cellpropertiesfill-patterncolor-member)|表示`format.fill.patternColor`属性。|
||[patternTintAndShade](/javascript/api/excel/excel.cellpropertiesfill#excel-excel-cellpropertiesfill-patterntintandshade-member)|表示`format.fill.patternTintAndShade`属性。|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfill#excel-excel-cellpropertiesfill-tintandshade-member)|表示`format.fill.tintAndShade`属性。|
|[CellPropertiesFont](/javascript/api/excel/excel.cellpropertiesfont)|[bold](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-bold-member)|表示`format.font.bold`属性。|
||[color](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-color-member)|表示`format.font.color`属性。|
||[italic](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-italic-member)|表示`format.font.italic`属性。|
||[name](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-name-member)|表示`format.font.name`属性。|
||[size](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-size-member)|表示`format.font.size`属性。|
||[strikethrough](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-strikethrough-member)|表示`format.font.strikethrough`属性。|
||[subscript](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-subscript-member)|表示`format.font.subscript`属性。|
||[superscript](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-superscript-member)|表示`format.font.superscript`属性。|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-tintandshade-member)|表示`format.font.tintAndShade`属性。|
||[underline](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-underline-member)|表示`format.font.underline`属性。|
|[CellPropertiesFormat](/javascript/api/excel/excel.cellpropertiesformat)|[autoIndent](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-autoindent-member)|表示`autoIndent`属性。|
||[Borders](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-borders-member)|表示`borders`属性。|
||[fill](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-fill-member)|表示`fill`属性。|
||[font](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-font-member)|表示`font`属性。|
||[horizontalAlignment](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-horizontalalignment-member)|表示`horizontalAlignment`属性。|
||[indentLevel](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-indentlevel-member)|表示`indentLevel`属性。|
||[protection](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-protection-member)|表示`protection`属性。|
||[readingOrder](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-readingorder-member)|表示`readingOrder`属性。|
||[shrinkToFit](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-shrinktofit-member)|表示`shrinkToFit`属性。|
||[textOrientation](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-textorientation-member)|表示`textOrientation`属性。|
||[useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-usestandardheight-member)|表示`useStandardHeight`属性。|
||[useStandardWidth](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-usestandardwidth-member)|表示`useStandardWidth`属性。|
||[verticalAlignment](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-verticalalignment-member)|表示`verticalAlignment`属性。|
||[wrapText](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-wraptext-member)|表示`wrapText`属性。|
|[CellPropertiesProtection](/javascript/api/excel/excel.cellpropertiesprotection)|[formulaHidden](/javascript/api/excel/excel.cellpropertiesprotection#excel-excel-cellpropertiesprotection-formulahidden-member)|表示`format.protection.formulaHidden`属性。|
||[locked](/javascript/api/excel/excel.cellpropertiesprotection#excel-excel-cellpropertiesprotection-locked-member)|表示`format.protection.locked`属性。|
|[ChangedEventDetail](/javascript/api/excel/excel.changedeventdetail)|[valueAfter](/javascript/api/excel/excel.changedeventdetail#excel-excel-changedeventdetail-valueafter-member)|表示更改后的值。|
||[valueBefore](/javascript/api/excel/excel.changedeventdetail#excel-excel-changedeventdetail-valuebefore-member)|表示更改前的值。|
||[valueTypeAfter](/javascript/api/excel/excel.changedeventdetail#excel-excel-changedeventdetail-valuetypeafter-member)|表示更改后的值的类型。|
||[valueTypeBefore](/javascript/api/excel/excel.changedeventdetail#excel-excel-changedeventdetail-valuetypebefore-member)|表示更改之前的值的类型。|
|[Chart](/javascript/api/excel/excel.chart)|[activate()](/javascript/api/excel/excel.chart#excel-excel-chart-activate-member(1))|在 Excel UI 中激活图表。|
||[pivotOptions](/javascript/api/excel/excel.chart#excel-excel-chart-pivotoptions-member)|封装数据透视图的选项。|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[colorScheme](/javascript/api/excel/excel.chartareaformat#excel-excel-chartareaformat-colorscheme-member)|指定图表的配色方案。|
||[roundedCorners](/javascript/api/excel/excel.chartareaformat#excel-excel-chartareaformat-roundedcorners-member)|指定图表的图表区域是否具有圆角。|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[linkNumberFormat](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-linknumberformat-member)|指定数字格式是否链接到单元格。|
|[ChartBinOptions](/javascript/api/excel/excel.chartbinoptions)|[allowOverflow](/javascript/api/excel/excel.chartbinoptions#excel-excel-chartbinoptions-allowoverflow-member)|指定在直方图或流程图中是否启用箱溢出。|
||[allowUnderflow](/javascript/api/excel/excel.chartbinoptions#excel-excel-chartbinoptions-allowunderflow-member)|指定在直方图或流程图中是否启用箱下溢。|
||[count](/javascript/api/excel/excel.chartbinoptions#excel-excel-chartbinoptions-count-member)|指定直方图或 pareto 图表的箱计数。|
||[overflowValue](/javascript/api/excel/excel.chartbinoptions#excel-excel-chartbinoptions-overflowvalue-member)|指定直方图或流程图的箱溢出值。|
||[type](/javascript/api/excel/excel.chartbinoptions#excel-excel-chartbinoptions-type-member)|指定直方图或流程图的箱类型。|
||[underflowValue](/javascript/api/excel/excel.chartbinoptions#excel-excel-chartbinoptions-underflowvalue-member)|指定直方图或流程图的箱下溢值。|
||[width](/javascript/api/excel/excel.chartbinoptions#excel-excel-chartbinoptions-width-member)|指定直方图或流程图的箱宽度值。|
|[ChartBoxwhiskerOptions](/javascript/api/excel/excel.chartboxwhiskeroptions)|[quartileCalculation](/javascript/api/excel/excel.chartboxwhiskeroptions#excel-excel-chartboxwhiskeroptions-quartilecalculation-member)|指定箱形图的四分位数计算类型。|
||[showInnerPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#excel-excel-chartboxwhiskeroptions-showinnerpoints-member)|指定内部点是否显示在箱形图中。|
||[showMeanLine](/javascript/api/excel/excel.chartboxwhiskeroptions#excel-excel-chartboxwhiskeroptions-showmeanline-member)|指定在箱形图中是否显示平均值。|
||[showMeanMarker](/javascript/api/excel/excel.chartboxwhiskeroptions#excel-excel-chartboxwhiskeroptions-showmeanmarker-member)|指定是否将平均值标记显示在箱形图中。|
||[showOutlierPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#excel-excel-chartboxwhiskeroptions-showoutlierpoints-member)|指定在箱形图中是否显示离群值点。|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-linknumberformat-member)|指定数字格式是否链接到单元格 (以便当数字格式在单元格区域更改时，标签中的数字) 。|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-linknumberformat-member)|指定数字格式是否链接到单元格。|
|[ChartErrorBars](/javascript/api/excel/excel.charterrorbars)|[endStyleCap](/javascript/api/excel/excel.charterrorbars#excel-excel-charterrorbars-endstylecap-member)|指定误差线是否具有结束样式上限。|
||[format](/javascript/api/excel/excel.charterrorbars#excel-excel-charterrorbars-format-member)|指定误差线的格式类型。|
||[include](/javascript/api/excel/excel.charterrorbars#excel-excel-charterrorbars-include-member)|指定包含误差线的哪些部分。|
||[type](/javascript/api/excel/excel.charterrorbars#excel-excel-charterrorbars-type-member)|误差线标记的区域类型。|
||[visible](/javascript/api/excel/excel.charterrorbars#excel-excel-charterrorbars-visible-member)|指定是否显示误差线。|
|[ChartErrorBarsFormat](/javascript/api/excel/excel.charterrorbarsformat)|[line](/javascript/api/excel/excel.charterrorbarsformat#excel-excel-charterrorbarsformat-line-member)|表示图表线条格式。|
|[ChartMapOptions](/javascript/api/excel/excel.chartmapoptions)|[labelStrategy](/javascript/api/excel/excel.chartmapoptions#excel-excel-chartmapoptions-labelstrategy-member)|指定区域地图图表的系列地图标签策略。|
||[level](/javascript/api/excel/excel.chartmapoptions#excel-excel-chartmapoptions-level-member)|指定区域地图图表的系列映射级别。|
||[projectionType](/javascript/api/excel/excel.chartmapoptions#excel-excel-chartmapoptions-projectiontype-member)|指定区域地图图表的系列投影类型。|
|[ChartPivotOptions](/javascript/api/excel/excel.chartpivotoptions)|[showAxisFieldButtons](/javascript/api/excel/excel.chartpivotoptions#excel-excel-chartpivotoptions-showaxisfieldbuttons-member)|指定是否在坐标轴上显示坐标轴字段数据透视图。|
||[showLegendFieldButtons](/javascript/api/excel/excel.chartpivotoptions#excel-excel-chartpivotoptions-showlegendfieldbuttons-member)|指定是否在图例上显示图例数据透视图。|
||[showReportFilterFieldButtons](/javascript/api/excel/excel.chartpivotoptions#excel-excel-chartpivotoptions-showreportfilterfieldbuttons-member)|指定是否在报表上显示报表筛选字段数据透视图。|
||[showValueFieldButtons](/javascript/api/excel/excel.chartpivotoptions#excel-excel-chartpivotoptions-showvaluefieldbuttons-member)|指定是否在项目上显示"显示值"字段数据透视图。|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[binOptions](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-binoptions-member)|封装直方图和排列图的容器选项。|
||[boxwhiskerOptions](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-boxwhiskeroptions-member)|封装箱形图的选项。|
||[bubbleScale](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-bubblescale-member)|这可以是从 0（零）到 300 的整数值，表示默认大小的百分比。|
||[gradientMaximumColor](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientmaximumcolor-member)|指定区域地图图表系列的最大值的颜色。|
||[gradientMaximumType](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientmaximumtype-member)|指定区域地图图表系列的最大值的类型。|
||[gradientMaximumValue](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientmaximumvalue-member)|指定区域地图图表系列的最大值。|
||[gradientMidpointColor](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientmidpointcolor-member)|指定区域地图图表系列的中点值的颜色。|
||[gradientMidpointType](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientmidpointtype-member)|指定区域地图图表系列的中点值的类型。|
||[gradientMidpointValue](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientmidpointvalue-member)|指定区域地图图表系列的中点值。|
||[gradientMinimumColor](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientminimumcolor-member)|指定区域地图图表系列的最小值的颜色。|
||[gradientMinimumType](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientminimumtype-member)|指定区域地图图表系列的最小值的类型。|
||[gradientMinimumValue](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientminimumvalue-member)|指定区域地图图表系列的最小值。|
||[gradientStyle](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientstyle-member)|指定区域地图图表的系列渐变样式。|
||[invertColor](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-invertcolor-member)|指定系列中负数据点的填充颜色。|
||[mapOptions](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-mapoptions-member)|封装区域地图图表的选项。|
||[parentLabelStrategy](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-parentlabelstrategy-member)|指定树状图的系列父标签策略区域。|
||[showConnectorLines](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-showconnectorlines-member)|指定是否在瀑布图中显示连接线。|
||[showLeaderLines](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-showleaderlines-member)|指定是否为系列中每个数据标签显示前导线。|
||[splitValue](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-splitvalue-member)|指定分隔复合饼图或复合条饼图的两个部分的阈值。|
||[xErrorBars](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-xerrorbars-member)|表示图表系列的误差线对象。|
||[yErrorBars](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-yerrorbars-member)|表示图表系列的误差线对象。|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[linkNumberFormat](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-linknumberformat-member)|指定数字格式是否链接到单元格 (以便当数字格式在单元格区域更改时，标签中的数字) 。|
|[ColumnProperties](/javascript/api/excel/excel.columnproperties)|[address](/javascript/api/excel/excel.columnproperties#excel-excel-columnproperties-address-member)|表示`address`属性。|
||[addressLocal](/javascript/api/excel/excel.columnproperties#excel-excel-columnproperties-addresslocal-member)|表示`addressLocal`属性。|
||[columnIndex](/javascript/api/excel/excel.columnproperties#excel-excel-columnproperties-columnindex-member)|表示`columnIndex`属性。|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[getRanges()](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-getranges-member(1))|`RangeAreas`返回 ，包含一个或多个矩形区域，其中应用了条件格式。|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[getInvalidCells()](/javascript/api/excel/excel.datavalidation#excel-excel-datavalidation-getinvalidcells-member(1))|返回包含 `RangeAreas` 一个或多个矩形区域的对象，其中单元格值无效。|
||[getInvalidCellsOrNullObject()](/javascript/api/excel/excel.datavalidation#excel-excel-datavalidation-getinvalidcellsornullobject-member(1))|返回包含 `RangeAreas` 一个或多个矩形区域的对象，其中单元格值无效。|
|[FilterCriteria](/javascript/api/excel/excel.filtercriteria)|[subField](/javascript/api/excel/excel.filtercriteria#excel-excel-filtercriteria-subfield-member)|筛选器用于对丰富值执行丰富筛选器的属性。|
|[GeometricShape](/javascript/api/excel/excel.geometricshape)|[id](/javascript/api/excel/excel.geometricshape#excel-excel-geometricshape-id-member)|返回形状标识符。|
||[shape](/javascript/api/excel/excel.geometricshape#excel-excel-geometricshape-shape-member)|`Shape`返回几何形状的对象。|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getCount()](/javascript/api/excel/excel.groupshapecollection#excel-excel-groupshapecollection-getcount-member(1))|返回形状组中的形状数量。|
||[getItem(key: string)](/javascript/api/excel/excel.groupshapecollection#excel-excel-groupshapecollection-getitem-member(1))|使用形状的名称或 ID 获取形状。|
||[getItemAt(index: number)](/javascript/api/excel/excel.groupshapecollection#excel-excel-groupshapecollection-getitemat-member(1))|根据其在集合中的位置获取形状。|
||[items](/javascript/api/excel/excel.groupshapecollection#excel-excel-groupshapecollection-items-member)|获取此集合中已加载的子项。|
|[HeaderFooter](/javascript/api/excel/excel.headerfooter)|[centerFooter](/javascript/api/excel/excel.headerfooter#excel-excel-headerfooter-centerfooter-member)|工作表的中心页脚。|
||[centerHeader](/javascript/api/excel/excel.headerfooter#excel-excel-headerfooter-centerheader-member)|工作表的中心页眉。|
||[LeftFooter](/javascript/api/excel/excel.headerfooter#excel-excel-headerfooter-leftfooter-member)|工作表的左页脚。|
||[leftHeader](/javascript/api/excel/excel.headerfooter#excel-excel-headerfooter-leftheader-member)|工作表的左页眉。|
||[rightFooter](/javascript/api/excel/excel.headerfooter#excel-excel-headerfooter-rightfooter-member)|工作表的右页脚。|
||[rightHeader](/javascript/api/excel/excel.headerfooter#excel-excel-headerfooter-rightheader-member)|工作表的右页眉。|
|[HeaderFooterGroup](/javascript/api/excel/excel.headerfootergroup)|[defaultForAllPages](/javascript/api/excel/excel.headerfootergroup#excel-excel-headerfootergroup-defaultforallpages-member)|常规页眉/页脚，除非指定偶数页/奇数页或首页，否则适用于所有页面，|
||[evenPages](/javascript/api/excel/excel.headerfootergroup#excel-excel-headerfootergroup-evenpages-member)|用于偶数页的页眉/页脚，需要为奇数页指定奇数页页眉/页脚。|
||[firstPage](/javascript/api/excel/excel.headerfootergroup#excel-excel-headerfootergroup-firstpage-member)|首页的页眉/页脚，为所有其他页使用常规或偶数页/奇数页页眉/页脚。|
||[oddPages](/javascript/api/excel/excel.headerfootergroup#excel-excel-headerfootergroup-oddpages-member)|用于奇数页的页眉/页脚，需要为偶数页指定偶数页页眉/页脚。|
||[state](/javascript/api/excel/excel.headerfootergroup#excel-excel-headerfootergroup-state-member)|设置页眉/页脚时所按的状态。|
||[useSheetMargins](/javascript/api/excel/excel.headerfootergroup#excel-excel-headerfootergroup-usesheetmargins-member)|获取或设置一个标记，指示页眉/页脚是否与工作表的页面布局选项中设置的页边距对齐。|
||[useSheetScale](/javascript/api/excel/excel.headerfootergroup#excel-excel-headerfootergroup-usesheetscale-member)|获取或设置一个标记，指示是否应按照工作表的页面布局选项中设置的页面缩放百分比来缩放页眉/页脚。|
|[Image](/javascript/api/excel/excel.image)|[format](/javascript/api/excel/excel.image#excel-excel-image-format-member)|返回图像的格式。|
||[id](/javascript/api/excel/excel.image#excel-excel-image-id-member)|指定图像对象的形状标识符。|
||[shape](/javascript/api/excel/excel.image#excel-excel-image-shape-member)|`Shape`返回与图像关联的对象。|
|[IterativeCalculation](/javascript/api/excel/excel.iterativecalculation)|[enabled](/javascript/api/excel/excel.iterativecalculation#excel-excel-iterativecalculation-enabled-member)|如果 Excel 使用迭代来处理循环引用，则为 True。|
||[maxChange](/javascript/api/excel/excel.iterativecalculation#excel-excel-iterativecalculation-maxchange-member)|指定每个迭代之间的最大更改量，因为Excel循环引用。|
||[maxIteration](/javascript/api/excel/excel.iterativecalculation#excel-excel-iterativecalculation-maxiteration-member)|指定可用于解析循环引用Excel迭代的最大次数。|
|[Line](/javascript/api/excel/excel.line)|[beginArrowheadLength](/javascript/api/excel/excel.line#excel-excel-line-beginarrowheadlength-member)|表示指定线条始端的箭头长度。|
||[beginArrowheadStyle](/javascript/api/excel/excel.line#excel-excel-line-beginarrowheadstyle-member)|表示指定线条始端的箭头样式。|
||[beginArrowheadWidth](/javascript/api/excel/excel.line#excel-excel-line-beginarrowheadwidth-member)|表示指定线条始端的箭头宽度。|
||[beginConnectedShape](/javascript/api/excel/excel.line#excel-excel-line-beginconnectedshape-member)|表示指定线条始端所附加到的形状。|
||[beginConnectedSite](/javascript/api/excel/excel.line#excel-excel-line-beginconnectedsite-member)|表示连接线始端所连接的连接站点。|
||[connectBeginShape(shape: Excel.Shape, connectionSite: number)](/javascript/api/excel/excel.line#excel-excel-line-connectbeginshape-member(1))|将指定连接线的始端附加到指定形状。|
||[connectEndShape(shape: Excel.Shape, connectionSite: number)](/javascript/api/excel/excel.line#excel-excel-line-connectendshape-member(1))|将指定连接线的末端附加到指定形状。|
||[connectorType](/javascript/api/excel/excel.line#excel-excel-line-connectortype-member)|表示线条的连接器类型。|
||[disconnectBeginShape()](/javascript/api/excel/excel.line#excel-excel-line-disconnectbeginshape-member(1))|使指定连接线的始端与形状脱离。|
||[disconnectEndShape()](/javascript/api/excel/excel.line#excel-excel-line-disconnectendshape-member(1))|使指定连接线的末端与形状脱离。|
||[endArrowheadLength](/javascript/api/excel/excel.line#excel-excel-line-endarrowheadlength-member)|表示指定线条末端的箭头长度。|
||[endArrowheadStyle](/javascript/api/excel/excel.line#excel-excel-line-endarrowheadstyle-member)|表示指定线条末端的箭头样式。|
||[endArrowheadWidth](/javascript/api/excel/excel.line#excel-excel-line-endarrowheadwidth-member)|表示指定线条末端的箭头宽度。|
||[endConnectedShape](/javascript/api/excel/excel.line#excel-excel-line-endconnectedshape-member)|表示指定线条末端所附加到的形状。|
||[endConnectedSite](/javascript/api/excel/excel.line#excel-excel-line-endconnectedsite-member)|表示连接线末端所连接的连接站点。|
||[id](/javascript/api/excel/excel.line#excel-excel-line-id-member)|指定形状标识符。|
||[isBeginConnected](/javascript/api/excel/excel.line#excel-excel-line-isbeginconnected-member)|指定指定线条的始值是否连接到形状。|
||[isEndConnected](/javascript/api/excel/excel.line#excel-excel-line-isendconnected-member)|指定指定线条的终点是否连接到形状。|
||[shape](/javascript/api/excel/excel.line#excel-excel-line-shape-member)|`Shape`返回与线条关联的对象。|
|[PageBreak](/javascript/api/excel/excel.pagebreak)|[columnIndex](/javascript/api/excel/excel.pagebreak#excel-excel-pagebreak-columnindex-member)|指定分页符的列索引。|
||[delete()](/javascript/api/excel/excel.pagebreak#excel-excel-pagebreak-delete-member(1))|删除分页符对象。|
||[getCellAfterBreak()](/javascript/api/excel/excel.pagebreak#excel-excel-pagebreak-getcellafterbreak-member(1))|获取分页符后的第一个单元格。|
||[rowIndex](/javascript/api/excel/excel.pagebreak#excel-excel-pagebreak-rowindex-member)|指定分页符的行索引。|
|[PageBreakCollection](/javascript/api/excel/excel.pagebreakcollection)|[add(pageBreakRange: Range \| string)](/javascript/api/excel/excel.pagebreakcollection#excel-excel-pagebreakcollection-add-member(1))|在指定区域的左上角单元格之前添加分页符。|
||[getCount()](/javascript/api/excel/excel.pagebreakcollection#excel-excel-pagebreakcollection-getcount-member(1))|获取集合中的分页符数量。|
||[getItem(index: number)](/javascript/api/excel/excel.pagebreakcollection#excel-excel-pagebreakcollection-getitem-member(1))|通过索引获取分页符对象。|
||[items](/javascript/api/excel/excel.pagebreakcollection#excel-excel-pagebreakcollection-items-member)|获取此集合中已加载的子项。|
||[removePageBreaks()](/javascript/api/excel/excel.pagebreakcollection#excel-excel-pagebreakcollection-removepagebreaks-member(1))|重置集合中的所有手动分页符。|
|[PageLayout](/javascript/api/excel/excel.pagelayout)|[blackAndWhite](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-blackandwhite-member)|工作表的黑白打印选项。|
||[bottomMargin](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-bottommargin-member)|要用于打印的工作表的底部页边距（以点表示）。|
||[centerHorizontally](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-centerhorizontally-member)|工作表的中心水平标记。|
||[centerVertically](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-centervertically-member)|工作表的中心垂直标记。|
||[draftMode](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-draftmode-member)|工作表的草稿模式选项。|
||[firstPageNumber](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-firstpagenumber-member)|要打印的工作表的第一页码。|
||[footerMargin](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-footermargin-member)|打印时工作表的页脚边距（以点表示）。|
||[getPrintArea()](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-getprintarea-member(1))|获取包含 `RangeAreas` 一个或多个矩形区域的对象，该对象表示工作表的打印区域。|
||[getPrintAreaOrNullObject()](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-getprintareaornullobject-member(1))|获取包含 `RangeAreas` 一个或多个矩形区域的对象，该对象表示工作表的打印区域。|
||[getPrintTitleColumns()](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-getprinttitlecolumns-member(1))|获取表示标题列的 Range 对象。|
||[getPrintTitleColumnsOrNullObject()](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-getprinttitlecolumnsornullobject-member(1))|获取表示标题列的 Range 对象。|
||[getPrintTitleRows()](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-getprinttitlerows-member(1))|获取表示标题行的 Range 对象。|
||[getPrintTitleRowsOrNullObject()](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-getprinttitlerowsornullobject-member(1))|获取表示标题行的 Range 对象。|
||[headerMargin](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-headermargin-member)|打印时工作表的页眉边距（以点表示）。|
||[headersFooters](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-headersfooters-member)|工作表的页眉和页脚配置。|
||[leftMargin](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-leftmargin-member)|打印时工作表的左边距（以点表示）。|
||[orientation](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-orientation-member)|工作表的页面方向。|
||[paperSize](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-papersize-member)|工作表的页面纸张大小。|
||[printComments](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-printcomments-member)|指定打印时是否显示工作表的注释。|
||[printErrors](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-printerrors-member)|工作表的打印错误选项。|
||[printGridlines](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-printgridlines-member)|指定是否打印工作表的网格线。|
||[printHeadings](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-printheadings-member)|指定是否打印工作表的标题。|
||[printOrder](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-printorder-member)|工作表的页面打印顺序选项。|
||[rightMargin](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-rightmargin-member)|打印时使用的工作表的右边距（以点表示）。|
||[setPrintArea(printArea: Range \| RangeAreas \| string)](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-setprintarea-member(1))|设置工作表的打印区域。|
||[setPrintMargins(unit: Excel.PrintMarginUnit, marginOptions: Excel.PageLayoutMarginOptions)](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-setprintmargins-member(1))|设置带单位的工作表的页边距。|
||[setPrintTitleColumns(printTitleColumns: Range \| string)](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-setprinttitlecolumns-member(1))|设置列，这些列包含要在打印的工作表的每页左侧重复的单元格。|
||[setPrintTitleRows(printTitleRows: Range \| string)](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-setprinttitlerows-member(1))|设置行，这些行包含要在打印的工作表的每页顶部重复的单元格。|
||[topMargin](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-topmargin-member)|打印时工作表的上边距（以点表示）。|
||[zoom](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-zoom-member)|工作表的打印缩放选项。|
|[PageLayoutMarginOptions](/javascript/api/excel/excel.pagelayoutmarginoptions)|[bottom](/javascript/api/excel/excel.pagelayoutmarginoptions#excel-excel-pagelayoutmarginoptions-bottom-member)|指定要用于打印的页面布局下边距（以指定单位表示）。|
||[footer](/javascript/api/excel/excel.pagelayoutmarginoptions#excel-excel-pagelayoutmarginoptions-footer-member)|指定要用于打印的页面布局页脚边距（以指定单位表示）。|
||[header](/javascript/api/excel/excel.pagelayoutmarginoptions#excel-excel-pagelayoutmarginoptions-header-member)|指定要用于打印的页面布局页眉边距（以指定单位表示）。|
||[left](/javascript/api/excel/excel.pagelayoutmarginoptions#excel-excel-pagelayoutmarginoptions-left-member)|指定要用于打印的页面布局左边距（以指定单位表示）。|
||[right](/javascript/api/excel/excel.pagelayoutmarginoptions#excel-excel-pagelayoutmarginoptions-right-member)|指定要用于打印的页面布局右边距（以指定单位表示）。|
||[top](/javascript/api/excel/excel.pagelayoutmarginoptions#excel-excel-pagelayoutmarginoptions-top-member)|指定要用于打印的页面布局上边距（以指定单位表示）。|
|[PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions)|[horizontalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#excel-excel-pagelayoutzoomoptions-horizontalfittopages-member)|水平放置的页数。|
||[scale](/javascript/api/excel/excel.pagelayoutzoomoptions#excel-excel-pagelayoutzoomoptions-scale-member)|打印页面缩放值可以介于 10 至 400 之间。|
||[verticalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#excel-excel-pagelayoutzoomoptions-verticalfittopages-member)|垂直放置的页数。|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[sortByValues(sortBy: Excel.SortBy, valuesHierarchy: Excel.DataPivotHierarchy, pivotItemScope?: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-sortbyvalues-member(1))|按给定范围中的指定值对 PivotField 进行排序。|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[autoFormat](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-autoformat-member)|指定在刷新格式或移动字段时是否自动设置格式。|
||[getDataHierarchy(cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-getdatahierarchy-member(1))|获取 DataHierarchy，它用于计算数据透视表中指定区域内的值。|
||[getPivotItems(axis: Excel.PivotAxis, cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-getpivotitems-member(1))|从构成数据透视表中指定区域内的值的轴获取 PivotItems。|
||[preserveFormatting](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-preserveformatting-member)|指定在通过透视、排序或更改页字段项等操作刷新或重新计算报表时是否保留格式。|
||[setAutoSortOnCell(cell: Range \| string, sortBy: Excel.SortBy)](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-setautosortoncell-member(1))|将数据透视表设置为使用指定的单元格设置自动排序，以自动选择排序的所有条件和上下文。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[enableDataValueEditing](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-enabledatavalueediting-member)|指定数据透视表是否允许用户编辑数据正文中的值。|
||[useCustomSortLists](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-usecustomsortlists-member)|指定数据透视表在排序时是否使用自定义列表。|
|[范围](/javascript/api/excel/excel.range)|[autoFill (destinationRange？： Range \| string， autoFillType？： Excel。AutoFillType) ](/javascript/api/excel/excel.range#excel-excel-range-autofill-member(1))|使用指定的自动填充逻辑填充从当前区域到目标区域的范围。|
||[convertDataTypeToText()](/javascript/api/excel/excel.range#excel-excel-range-convertdatatypetotext-member(1))|将数据类型为区域单元格转换为文本。|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.range#excel-excel-range-converttolinkeddatatype-member(1))|将区域单元格转换为工作表中的链接数据类型。|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.range#excel-excel-range-copyfrom-member(1))|将单元格数据或格式从源区域或当前 `RangeAreas` 区域复制。|
||[find(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#excel-excel-range-find-member(1))|根据指定的条件查找给定的字符串。|
||[findOrNullObject(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#excel-excel-range-findornullobject-member(1))|根据指定的条件查找给定的字符串。|
||[flashFill()](/javascript/api/excel/excel.range#excel-excel-range-flashfill-member(1))|对当前范围进行快速填充。|
||[getCellProperties(cellPropertiesLoadOptions: CellPropertiesLoadOptions)](/javascript/api/excel/excel.range#excel-excel-range-getcellproperties-member(1))|返回一个 2D 数组，其中封装了每个单元格的字体、填充、边框、对齐方式和其他属性数据。|
||[getColumnProperties(columnPropertiesLoadOptions: ColumnPropertiesLoadOptions)](/javascript/api/excel/excel.range#excel-excel-range-getcolumnproperties-member(1))|返回一个一维数组，其中封装了每个列的字体、填充、边框、对齐方式和其他属性数据。|
||[getRowProperties(rowPropertiesLoadOptions: RowPropertiesLoadOptions)](/javascript/api/excel/excel.range#excel-excel-range-getrowproperties-member(1))|返回一个一维数组，其中封装了每个行的字体、填充、边框、对齐方式和其他属性数据。|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#excel-excel-range-getspecialcells-member(1))|获取包含 `RangeAreas` 一个或多个矩形区域的对象，该对象表示与指定类型和值匹配的所有单元格。|
||[getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#excel-excel-range-getspecialcellsornullobject-member(1))|获取包含 `RangeAreas` 一个或多个区域的对象，该对象代表与指定类型和值匹配的所有单元格。|
||[getTables(fullyContained?: boolean)](/javascript/api/excel/excel.range#excel-excel-range-gettables-member(1))|获取与区域重叠的限定范围的表格集合。|
||[linkedDataTypeState](/javascript/api/excel/excel.range#excel-excel-range-linkeddatatypestate-member)|表示每个单元格的数据类型状态。|
||[removeDuplicates(columns: number[], includesHeader: boolean)](/javascript/api/excel/excel.range#excel-excel-range-removeduplicates-member(1))|从列指定的区域中删除重复值。|
||[replaceAll(text: string, replacement: string, criteria: Excel.ReplaceCriteria)](/javascript/api/excel/excel.range#excel-excel-range-replaceall-member(1))|根据当前区域内指定的条件查找并替换给定的字符串。|
||[setCellProperties(cellPropertiesData: SettableCellProperties[][])](/javascript/api/excel/excel.range#excel-excel-range-setcellproperties-member(1))|根据单元格属性的 2D 数组更新区域，并封装字体、填充、边框和对齐方式等内容。|
||[setColumnProperties(columnPropertiesData: SettableColumnProperties[])](/javascript/api/excel/excel.range#excel-excel-range-setcolumnproperties-member(1))|根据列属性的一维数组更新区域，并封装字体、填充、边框和对齐方式等内容。|
||[setDirty()](/javascript/api/excel/excel.range#excel-excel-range-setdirty-member(1))|设置下一次重新计算发生时要重新计算的区域。|
||[setRowProperties(rowPropertiesData: SettableRowProperties[])](/javascript/api/excel/excel.range#excel-excel-range-setrowproperties-member(1))|根据行属性的一维数组更新区域，并封装字体、填充、边框和对齐方式等内容。|
|[RangeAreas](/javascript/api/excel/excel.rangeareas)|[address](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-address-member)|`RangeAreas`以 A1 样式返回引用。|
||[addressLocal](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-addresslocal-member)|`RangeAreas`返回用户区域设置中的引用。|
||[areaCount](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-areacount-member)|返回包含此对象的矩形区域 `RangeAreas` 的数量。|
||[areas](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-areas-member)|返回包含此对象的矩形区域 `RangeAreas` 的集合。|
||[calculate()](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-calculate-member(1))|计算 中所有单元格 `RangeAreas`。|
||[cellCount](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-cellcount-member)|返回对象中的单元格数 `RangeAreas` ，并汇总所有单个矩形范围的单元格计数。|
||[clear(applyTo?: Excel.ClearApplyTo)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-clear-member(1))|清除包含此对象的每个区域的值、格式、填充、边框和其他 `RangeAreas` 属性。|
||[conditionalFormats](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-conditionalformats-member)|返回一组条件格式，这些格式与该对象中任何单元格相交 `RangeAreas` 。|
||[convertDataTypeToText()](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-convertdatatypetotext-member(1))|将 数据类型为 `RangeAreas` 的 中所有单元格转换为文本。|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-converttolinkeddatatype-member(1))|将 中所有单元格 `RangeAreas` 转换为链接数据类型。|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-copyfrom-member(1))|将单元格数据或格式从源区域或复制到 `RangeAreas` 当前 `RangeAreas`。|
||[dataValidation](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-datavalidation-member)|返回 中所有区域的数据验证对象 `RangeAreas`。|
||[format](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-format-member)|返回一 `RangeFormat` 个对象，该对象封装对象中所有范围的字体、填充、边框、对齐方式和其他 `RangeAreas` 属性。|
||[getEntireColumn()](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-getentirecolumn-member(1))|`RangeAreas` `RangeAreas` `RangeAreas` 返回一个对象，该对象代表 (例如，如果当前表示单元格"B4：E11， H2"`RangeAreas`，则返回表示列"B：E， H：H") 。|
||[getEntireRow()](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-getentirerow-member(1))|`RangeAreas`返回一个对象`RangeAreas``RangeAreas`，该对象代表 (例如，如果当前表示单元格"B4：E11"`RangeAreas`，则返回表示行"4：11") 。|
||[getIntersection(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-getintersection-member(1))|`RangeAreas`返回表示给定区域或 的交集的对象`RangeAreas`。|
||[getIntersectionOrNullObject(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-getintersectionornullobject-member(1))|`RangeAreas`返回表示给定区域或 的交集的对象`RangeAreas`。|
||[getOffsetRangeAreas(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-getoffsetrangeareas-member(1))|返回由 `RangeAreas` 特定行和列偏移移动的对象。|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-getspecialcells-member(1))|返回一 `RangeAreas` 个对象，该对象代表与指定类型和值匹配的所有单元格。|
||[getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-getspecialcellsornullobject-member(1))|返回一 `RangeAreas` 个对象，该对象代表与指定类型和值匹配的所有单元格。|
||[getTables(fullyContained?: boolean)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-gettables-member(1))|返回与该对象中任何区域重叠的表的范围集合 `RangeAreas` 。|
||[getUsedRangeAreas(valuesOnly?: boolean)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-getusedrangeareas-member(1))|返回由 `RangeAreas` 对象中各个矩形区域的所有已用区域组成的 `RangeAreas` used。|
||[getUsedRangeAreasOrNullObject(valuesOnly?: boolean)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-getusedrangeareasornullobject-member(1))|返回由 `RangeAreas` 对象中各个矩形区域的所有已用区域组成的 `RangeAreas` used。|
||[isEntireColumn](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-isentirecolumn-member)|指定此对象上的所有区域 `RangeAreas` 是否代表整个列 (例如，"A：C，Q：Z") 。|
||[isEntireRow](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-isentirerow-member)|指定此对象上的所有 `RangeAreas` 区域是否代表整个行 (例如，"1：3， 5：7") 。|
||[setDirty()](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-setdirty-member(1))|`RangeAreas`设置在下次重新计算时要重新计算的 。|
||[style](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-style-member)|表示此对象中所有范围的 `RangeAreas` 样式。|
||[worksheet](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-worksheet-member)|返回当前 的工作表 `RangeAreas`。|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[tintAndShade](/javascript/api/excel/excel.rangeborder#excel-excel-rangeborder-tintandshade-member)|指定使区域边框的颜色变亮或变暗的双精度值，该值介于 -1 (最暗) 和 1 (最亮) 之间，原始颜色为 0。|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[tintAndShade](/javascript/api/excel/excel.rangebordercollection#excel-excel-rangebordercollection-tintandshade-member)|指定使区域边框的颜色变亮或变暗的双精度型值。|
|[RangeCollection](/javascript/api/excel/excel.rangecollection)|[getCount()](/javascript/api/excel/excel.rangecollection#excel-excel-rangecollection-getcount-member(1))|返回 中区域的数量 `RangeCollection`。|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangecollection#excel-excel-rangecollection-getitemat-member(1))|基于 range 对象在 中的位置返回该对象 `RangeCollection`。|
||[items](/javascript/api/excel/excel.rangecollection#excel-excel-rangecollection-items-member)|获取此集合中已加载的子项。|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[pattern](/javascript/api/excel/excel.rangefill#excel-excel-rangefill-pattern-member)|范围的图案。|
||[patternColor](/javascript/api/excel/excel.rangefill#excel-excel-rangefill-patterncolor-member)|HTML 颜色代码，表示区域图案的颜色，格式为 #RRGGBB (例如"FFA500") ，或作为已命名的 HTML 颜色 (例如"orange") 。|
||[patternTintAndShade](/javascript/api/excel/excel.rangefill#excel-excel-rangefill-patterntintandshade-member)|指定使区域填充的图案颜色变亮或变暗的双精度型。|
||[tintAndShade](/javascript/api/excel/excel.rangefill#excel-excel-rangefill-tintandshade-member)|指定使区域填充的颜色变亮或变暗的双精度值。|
|[RangeFont](/javascript/api/excel/excel.rangefont)|[strikethrough](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-strikethrough-member)|指定字体的删除线状态。|
||[subscript](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-subscript-member)|指定字体的下标状态。|
||[superscript](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-superscript-member)|指定字体的上标状态。|
||[tintAndShade](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-tintandshade-member)|指定使区域字体的颜色变亮或变暗的双精度值。|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[autoIndent](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-autoindent-member)|指定文本对齐方式设置为相等分布时文本是否自动缩进。|
||[indentLevel](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-indentlevel-member)|0 到 250 之间的一个整数，指示缩进水平。|
||[readingOrder](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-readingorder-member)|区域的读取顺序。|
||[shrinkToFit](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-shrinktofit-member)|指定文本是否自动缩小以适应可用列宽。|
|[RemoveDuplicatesResult](/javascript/api/excel/excel.removeduplicatesresult)|[removed](/javascript/api/excel/excel.removeduplicatesresult#excel-excel-removeduplicatesresult-removed-member)|由操作删除的重复行数。|
||[uniqueRemaining](/javascript/api/excel/excel.removeduplicatesresult#excel-excel-removeduplicatesresult-uniqueremaining-member)|所生成的区域中存在的剩余唯一行数。|
|[ReplaceCriteria](/javascript/api/excel/excel.replacecriteria)|[completeMatch](/javascript/api/excel/excel.replacecriteria#excel-excel-replacecriteria-completematch-member)|指定匹配是需要完成还是部分匹配。|
||[matchCase](/javascript/api/excel/excel.replacecriteria#excel-excel-replacecriteria-matchcase-member)|指定匹配是否区分大小写。|
|[RowProperties](/javascript/api/excel/excel.rowproperties)|[address](/javascript/api/excel/excel.rowproperties#excel-excel-rowproperties-address-member)|表示`address`属性。|
||[addressLocal](/javascript/api/excel/excel.rowproperties#excel-excel-rowproperties-addresslocal-member)|表示`addressLocal`属性。|
||[rowIndex](/javascript/api/excel/excel.rowproperties#excel-excel-rowproperties-rowindex-member)|表示`rowIndex`属性。|
|[SearchCriteria](/javascript/api/excel/excel.searchcriteria)|[completeMatch](/javascript/api/excel/excel.searchcriteria#excel-excel-searchcriteria-completematch-member)|指定匹配是需要完成还是部分匹配。|
||[matchCase](/javascript/api/excel/excel.searchcriteria#excel-excel-searchcriteria-matchcase-member)|指定匹配是否区分大小写。|
||[searchDirection](/javascript/api/excel/excel.searchcriteria#excel-excel-searchcriteria-searchdirection-member)|指定搜索方向。|
|[SettableCellProperties](/javascript/api/excel/excel.settablecellproperties)|[format](/javascript/api/excel/excel.settablecellproperties#excel-excel-settablecellproperties-format-member)|表示`format`属性。|
||[hyperlink](/javascript/api/excel/excel.settablecellproperties#excel-excel-settablecellproperties-hyperlink-member)|表示`hyperlink`属性。|
||[style](/javascript/api/excel/excel.settablecellproperties#excel-excel-settablecellproperties-style-member)|表示`style`属性。|
|[SettableColumnProperties](/javascript/api/excel/excel.settablecolumnproperties)|[columnHidden](/javascript/api/excel/excel.settablecolumnproperties#excel-excel-settablecolumnproperties-columnhidden-member)|表示`columnHidden`属性。|
||[columnWidth](/javascript/api/excel/excel.settablecolumnproperties#excel-excel-settablecolumnproperties-columnwidth-member)||
||[format：Excel。CellPropertiesFormat & { columnWidth](/javascript/api/excel/excel.settablecolumnproperties#excel-excel-settablecolumnproperties-format-member)|表示`format`属性。|
|[SettableRowProperties](/javascript/api/excel/excel.settablerowproperties)|[format：Excel。CellPropertiesFormat & { rowHeight](/javascript/api/excel/excel.settablerowproperties#excel-excel-settablerowproperties-format-member)|表示`format`属性。|
||[rowHeight](/javascript/api/excel/excel.settablerowproperties#excel-excel-settablerowproperties-rowheight-member)||
||[rowHidden](/javascript/api/excel/excel.settablerowproperties#excel-excel-settablerowproperties-rowhidden-member)|表示`rowHidden`属性。|
|[Shape](/javascript/api/excel/excel.shape)|[altTextDescription](/javascript/api/excel/excel.shape#excel-excel-shape-alttextdescription-member)|指定对象的可选说明 `Shape` 文本。|
||[altTextTitle](/javascript/api/excel/excel.shape#excel-excel-shape-alttexttitle-member)|指定对象的可选标题 `Shape` 文本。|
||[connectionSiteCount](/javascript/api/excel/excel.shape#excel-excel-shape-connectionsitecount-member)|返回此形状上的连接站点数。|
||[delete()](/javascript/api/excel/excel.shape#excel-excel-shape-delete-member(1))|从工作表删除形状。|
||[fill](/javascript/api/excel/excel.shape#excel-excel-shape-fill-member)|返回此形状的填充格式。|
||[geometricShape](/javascript/api/excel/excel.shape#excel-excel-shape-geometricshape-member)|返回与形状关联的几何形状。|
||[geometricShapeType](/javascript/api/excel/excel.shape#excel-excel-shape-geometricshapetype-member)|指定此几何形状的几何形状类型。|
||[getAsImage(format: Excel.PictureFormat)](/javascript/api/excel/excel.shape#excel-excel-shape-getasimage-member(1))|将形状转换为图像并将图像返回为 base64 编码字符串。|
||[组](/javascript/api/excel/excel.shape#excel-excel-shape-group-member)|返回与形状关联的形状组。|
||[height](/javascript/api/excel/excel.shape#excel-excel-shape-height-member)|指定形状的高度（以点表示）。|
||[id](/javascript/api/excel/excel.shape#excel-excel-shape-id-member)|指定形状标识符。|
||[image](/javascript/api/excel/excel.shape#excel-excel-shape-image-member)|返回与形状关联的图像。|
||[incrementLeft(increment: number)](/javascript/api/excel/excel.shape#excel-excel-shape-incrementleft-member(1))|以指定磅数水平移动形状。|
||[incrementRotation(increment: number)](/javascript/api/excel/excel.shape#excel-excel-shape-incrementrotation-member(1))|将形状围绕 z 轴旋转特定度数。|
||[incrementTop(increment: number)](/javascript/api/excel/excel.shape#excel-excel-shape-incrementtop-member(1))|以指定磅数垂直移动形状。|
||[left](/javascript/api/excel/excel.shape#excel-excel-shape-left-member)|从形状左侧到工作表左侧的距离（以磅为单位）。|
||[level](/javascript/api/excel/excel.shape#excel-excel-shape-level-member)|指定指定形状的级别。|
||[line](/javascript/api/excel/excel.shape#excel-excel-shape-line-member)|返回与形状关联的线条。|
||[lineFormat](/javascript/api/excel/excel.shape#excel-excel-shape-lineformat-member)|返回此形状的线条格式。|
||[lockAspectRatio](/javascript/api/excel/excel.shape#excel-excel-shape-lockaspectratio-member)|指定是否锁定此形状的纵横比。|
||[name](/javascript/api/excel/excel.shape#excel-excel-shape-name-member)|指定形状的名称。|
||[onActivated](/javascript/api/excel/excel.shape#excel-excel-shape-onactivated-member)|当激活形状时发生此事件。|
||[onDeactivated](/javascript/api/excel/excel.shape#excel-excel-shape-ondeactivated-member)|当停用形状时发生此事件。|
||[parentGroup](/javascript/api/excel/excel.shape#excel-excel-shape-parentgroup-member)|指定此形状的父组。|
||[rotation](/javascript/api/excel/excel.shape#excel-excel-shape-rotation-member)|指定形状的旋转角度（以度数表示）。|
||[scaleHeight(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#excel-excel-shape-scaleheight-member(1))|按指定因子缩放形状的高度。|
||[scaleWidth(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#excel-excel-shape-scalewidth-member(1))|按指定因子缩放形状的宽度。|
||[setZOrder(position: Excel.ShapeZOrder)](/javascript/api/excel/excel.shape#excel-excel-shape-setzorder-member(1))|将指定形状沿集合的 z 顺序向上或向下移动，将其移动到其他形状的前面或后面。|
||[textFrame](/javascript/api/excel/excel.shape#excel-excel-shape-textframe-member)|返回此形状的文本框对象。|
||[top](/javascript/api/excel/excel.shape#excel-excel-shape-top-member)|从形状上边缘到工作表上边缘之间的距离（以磅为单位）。|
||[type](/javascript/api/excel/excel.shape#excel-excel-shape-type-member)|返回此形状的类型。|
||[visible](/javascript/api/excel/excel.shape#excel-excel-shape-visible-member)|指定形状是否可见。|
||[width](/javascript/api/excel/excel.shape#excel-excel-shape-width-member)|指定形状的宽度（以点表示）。|
||[zOrderPosition](/javascript/api/excel/excel.shape#excel-excel-shape-zorderposition-member)|返回指定形状在 z 顺序中的位置，0 表示顺序堆栈的底部。|
|[ShapeActivatedEventArgs](/javascript/api/excel/excel.shapeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapeactivatedeventargs#excel-excel-shapeactivatedeventargs-shapeid-member)|获取已激活形状的 ID。|
||[type](/javascript/api/excel/excel.shapeactivatedeventargs#excel-excel-shapeactivatedeventargs-type-member)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.shapeactivatedeventargs#excel-excel-shapeactivatedeventargs-worksheetid-member)|获取激活形状的工作表的 ID。|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addGeometricShape(geometricShapeType: Excel.GeometricShapeType)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addgeometricshape-member(1))|将几何形状添加到工作表。|
||[addGroup(values: Array<string \| Shape>)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addgroup-member(1))|在此集合的工作表中对形状的子集进行分组。|
||[addImage(base64ImageString: string)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addimage-member(1))|从 base64 编码的字符串创建图像并将其添加到工作表。|
||[addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addline-member(1))|将线条添加到工作表。|
||[addTextBox(text?: string)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addtextbox-member(1))|使用提供的文本作为内容，将文本框添加到工作表。|
||[getCount()](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-getcount-member(1))|返回工作表中的形状数。|
||[getItem(key: string)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-getitem-member(1))|使用形状的名称或 ID 获取形状。|
||[getItemAt(index: number)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-getitemat-member(1))|使用其在集合中的位置获取形状。|
||[items](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-items-member)|获取此集合中已加载的子项。|
|[ShapeDeactivatedEventArgs](/javascript/api/excel/excel.shapedeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapedeactivatedeventargs#excel-excel-shapedeactivatedeventargs-shapeid-member)|获取已停用形状的形状的 ID。|
||[type](/javascript/api/excel/excel.shapedeactivatedeventargs#excel-excel-shapedeactivatedeventargs-type-member)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.shapedeactivatedeventargs#excel-excel-shapedeactivatedeventargs-worksheetid-member)|获取在其中停用形状的工作表的 ID。|
|[ShapeFill](/javascript/api/excel/excel.shapefill)|[clear()](/javascript/api/excel/excel.shapefill#excel-excel-shapefill-clear-member(1))|清除此形状的填充格式。|
||[foregroundColor](/javascript/api/excel/excel.shapefill#excel-excel-shapefill-foregroundcolor-member)|以 HTML 颜色格式表示形状填充前景色，格式为 #RRGGBB (例如"FFA500") 或作为已命名的 HTML 颜色 (例如"orange") |
||[setSolidColor(color: string)](/javascript/api/excel/excel.shapefill#excel-excel-shapefill-setsolidcolor-member(1))|将形状的填充格式设置为统一颜色。|
||[transparency](/javascript/api/excel/excel.shapefill#excel-excel-shapefill-transparency-member)|将填充的透明度百分比指定为从 0.0 到 1.0 (不透明) 1.0 (透明) 。|
||[type](/javascript/api/excel/excel.shapefill#excel-excel-shapefill-type-member)|返回形状的填充类型。|
|[ShapeFont](/javascript/api/excel/excel.shapefont)|[bold](/javascript/api/excel/excel.shapefont#excel-excel-shapefont-bold-member)|表示字体的加粗状态。|
||[color](/javascript/api/excel/excel.shapefont#excel-excel-shapefont-color-member)|文本颜色格式的 HTML 颜色代码表示 (例如，"#FF0000"表示红色) 。|
||[italic](/javascript/api/excel/excel.shapefont#excel-excel-shapefont-italic-member)|表示字体的斜体状态。|
||[name](/javascript/api/excel/excel.shapefont#excel-excel-shapefont-name-member)|表示字体名称 (例如"Calibri") 。|
||[size](/javascript/api/excel/excel.shapefont#excel-excel-shapefont-size-member)|表示字号（以 (，例如 11) ）。|
||[underline](/javascript/api/excel/excel.shapefont#excel-excel-shapefont-underline-member)|应用于字体的下划线类型。|
|[ShapeGroup](/javascript/api/excel/excel.shapegroup)|[id](/javascript/api/excel/excel.shapegroup#excel-excel-shapegroup-id-member)|指定形状标识符。|
||[shape](/javascript/api/excel/excel.shapegroup#excel-excel-shapegroup-shape-member)|`Shape`返回与组关联的对象。|
||[shapes](/javascript/api/excel/excel.shapegroup#excel-excel-shapegroup-shapes-member)|返回对象 `Shape` 的集合。|
||[ungroup()](/javascript/api/excel/excel.shapegroup#excel-excel-shapegroup-ungroup-member(1))|取消分组指定形状组中的任何已分组形状。|
|[ShapeLineFormat](/javascript/api/excel/excel.shapelineformat)|[color](/javascript/api/excel/excel.shapelineformat#excel-excel-shapelineformat-color-member)|表示 HTML 颜色格式的线条颜色，格式为 #RRGGBB (例如"FFA500") 或作为已命名的 HTML 颜色 (例如"orange") 。|
||[dashStyle](/javascript/api/excel/excel.shapelineformat#excel-excel-shapelineformat-dashstyle-member)|表示形状的线条样式。|
||[style](/javascript/api/excel/excel.shapelineformat#excel-excel-shapelineformat-style-member)|表示形状的线条样式。|
||[transparency](/javascript/api/excel/excel.shapelineformat#excel-excel-shapelineformat-transparency-member)|将指定线条的透明度表示为从 0.0（不透明）到 1.0（清晰）的值。|
||[visible](/javascript/api/excel/excel.shapelineformat#excel-excel-shapelineformat-visible-member)|指定形状元素的线条格式是否可见。|
||[weight](/javascript/api/excel/excel.shapelineformat#excel-excel-shapelineformat-weight-member)|表示线条的粗细（以磅为单位）。|
|[SortField](/javascript/api/excel/excel.sortfield)|[subField](/javascript/api/excel/excel.sortfield#excel-excel-sortfield-subfield-member)|指定作为要排序的丰富值的目标属性名称的子字段。|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[getCount()](/javascript/api/excel/excel.stylecollection#excel-excel-stylecollection-getcount-member(1))|获取集合中的样式数量。|
||[getItemAt(index: number)](/javascript/api/excel/excel.stylecollection#excel-excel-stylecollection-getitemat-member(1))|根据其在集合中的位置获取样式。|
|[Table](/javascript/api/excel/excel.table)|[autoFilter](/javascript/api/excel/excel.table#excel-excel-table-autofilter-member)|`AutoFilter`表示 table 的对象。|
|[TableAddedEventArgs](/javascript/api/excel/excel.tableaddedeventargs)|[source](/javascript/api/excel/excel.tableaddedeventargs#excel-excel-tableaddedeventargs-source-member)|获取事件源。|
||[tableId](/javascript/api/excel/excel.tableaddedeventargs#excel-excel-tableaddedeventargs-tableid-member)|获取已添加的表的 ID。|
||[type](/javascript/api/excel/excel.tableaddedeventargs#excel-excel-tableaddedeventargs-type-member)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.tableaddedeventargs#excel-excel-tableaddedeventargs-worksheetid-member)|获取添加表格的工作表的 ID。|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[details](/javascript/api/excel/excel.tablechangedeventargs#excel-excel-tablechangedeventargs-details-member)|获取有关更改详细信息的信息。|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onAdded](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-onadded-member)|在工作簿中添加新表时发生。|
||[onDeleted](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-ondeleted-member)|在工作簿中删除指定的表格时发生。|
|[TableDeletedEventArgs](/javascript/api/excel/excel.tabledeletedeventargs)|[source](/javascript/api/excel/excel.tabledeletedeventargs#excel-excel-tabledeletedeventargs-source-member)|获取事件源。|
||[tableId](/javascript/api/excel/excel.tabledeletedeventargs#excel-excel-tabledeletedeventargs-tableid-member)|获取已删除的表的 ID。|
||[tableName](/javascript/api/excel/excel.tabledeletedeventargs#excel-excel-tabledeletedeventargs-tablename-member)|获取已删除的表的名称。|
||[type](/javascript/api/excel/excel.tabledeletedeventargs#excel-excel-tabledeletedeventargs-type-member)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.tabledeletedeventargs#excel-excel-tabledeletedeventargs-worksheetid-member)|获取删除表格的工作表的 ID。|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getCount()](/javascript/api/excel/excel.tablescopedcollection#excel-excel-tablescopedcollection-getcount-member(1))|获取集合中的表数量。|
||[getFirst()](/javascript/api/excel/excel.tablescopedcollection#excel-excel-tablescopedcollection-getfirst-member(1))|获取集合中的第一个表格。|
||[getItem(key: string)](/javascript/api/excel/excel.tablescopedcollection#excel-excel-tablescopedcollection-getitem-member(1))|按名称或 ID 获取表。|
||[items](/javascript/api/excel/excel.tablescopedcollection#excel-excel-tablescopedcollection-items-member)|获取此集合中已加载的子项。|
|[TextFrame](/javascript/api/excel/excel.textframe)|[autoSizeSetting](/javascript/api/excel/excel.textframe#excel-excel-textframe-autosizesetting-member)|文本框的自动大小调整设置。|
||[bottomMargin](/javascript/api/excel/excel.textframe#excel-excel-textframe-bottommargin-member)|表示文本框的下边距（以磅为单位）。|
||[deleteText()](/javascript/api/excel/excel.textframe#excel-excel-textframe-deletetext-member(1))|删除文本框中的所有文本。|
||[hasText](/javascript/api/excel/excel.textframe#excel-excel-textframe-hastext-member)|指定文本框是否包含文本。|
||[horizontalAlignment](/javascript/api/excel/excel.textframe#excel-excel-textframe-horizontalalignment-member)|表示文本框的水平对齐方式。|
||[horizontalOverflow](/javascript/api/excel/excel.textframe#excel-excel-textframe-horizontaloverflow-member)|表示文本框的水平溢出行为。|
||[leftMargin](/javascript/api/excel/excel.textframe#excel-excel-textframe-leftmargin-member)|表示文本框的左边距（以磅为单位）。|
||[orientation](/javascript/api/excel/excel.textframe#excel-excel-textframe-orientation-member)|表示文本面向文本框的角度。|
||[readingOrder](/javascript/api/excel/excel.textframe#excel-excel-textframe-readingorder-member)|表示文本框从左到右或从右到左的读取顺序。|
||[rightMargin](/javascript/api/excel/excel.textframe#excel-excel-textframe-rightmargin-member)|表示文本框的右边距（以磅为单位）。|
||[textRange](/javascript/api/excel/excel.textframe#excel-excel-textframe-textrange-member)|表示附加到文本框中形状上的文本，以及用于操作文本的属性和方法。|
||[topMargin](/javascript/api/excel/excel.textframe#excel-excel-textframe-topmargin-member)|表示文本框的上边距（以磅为单位）。|
||[verticalAlignment](/javascript/api/excel/excel.textframe#excel-excel-textframe-verticalalignment-member)|表示文本框的垂直对齐方式。|
||[verticalOverflow](/javascript/api/excel/excel.textframe#excel-excel-textframe-verticaloverflow-member)|表示文本框的垂直溢出行为。|
|[TextRange](/javascript/api/excel/excel.textrange)|[font](/javascript/api/excel/excel.textrange#excel-excel-textrange-font-member)|`ShapeFont`返回一个对象，该对象代表文本范围的字体属性。|
||[getSubstring(start: number, length?: number)](/javascript/api/excel/excel.textrange#excel-excel-textrange-getsubstring-member(1))|返回给定区域内子字符串的 TextRange 对象。|
||[text](/javascript/api/excel/excel.textrange#excel-excel-textrange-text-member)|表示文本范围的纯文本内容。|
|[Workbook](/javascript/api/excel/excel.workbook)|[autoSave](/javascript/api/excel/excel.workbook#excel-excel-workbook-autosave-member)|指定工作簿是否位于"自动保存"模式。|
||[calculationEngineVersion](/javascript/api/excel/excel.workbook#excel-excel-workbook-calculationengineversion-member)|返回有关 Excel 计算引擎的版本号。|
||[chartDataPointTrack](/javascript/api/excel/excel.workbook#excel-excel-workbook-chartdatapointtrack-member)|如果工作簿中的所有图表都跟踪它们所附加的实际数据点，则为 True。|
||[getActiveChart()](/javascript/api/excel/excel.workbook#excel-excel-workbook-getactivechart-member(1))|获取工作簿中的当前活动图表。|
||[getActiveChartOrNullObject()](/javascript/api/excel/excel.workbook#excel-excel-workbook-getactivechartornullobject-member(1))|获取工作簿中的当前活动图表。|
||[getIsActiveCollabSession()](/javascript/api/excel/excel.workbook#excel-excel-workbook-getisactivecollabsession-member(1))|返回 `true` 工作簿是否正由多个用户编辑， (共同创作) 。|
||[getSelectedRanges()](/javascript/api/excel/excel.workbook#excel-excel-workbook-getselectedranges-member(1))|从工作簿中获取当前选定的一个或多个区域。|
||[isDirty](/javascript/api/excel/excel.workbook#excel-excel-workbook-isdirty-member)|指定自上次保存工作簿以来是否进行了更改。|
||[onAutoSaveSettingChanged](/javascript/api/excel/excel.workbook#excel-excel-workbook-onautosavesettingchanged-member)|在工作簿上更改"自动保存"设置时发生。|
||[previouslySaved](/javascript/api/excel/excel.workbook#excel-excel-workbook-previouslysaved-member)|指定工作簿是在本地保存还是联机保存。|
||[usePrecisionAsDisplayed](/javascript/api/excel/excel.workbook#excel-excel-workbook-useprecisionasdisplayed-member)|如果此工作簿中的计算仅使用显示的数字精度来完成，则为 True。|
|[WorkbookAutoSaveSettingChangedEventArgs](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs)|[type](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs#excel-excel-workbookautosavesettingchangedeventargs-type-member)|获取事件的类型。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[autoFilter](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-autofilter-member)|`AutoFilter`表示工作表的对象。|
||[enableCalculation](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-enablecalculation-member)|确定是否Excel应重新计算工作表（ 如有必要）。|
||[findAll(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-findall-member(1))|根据指定的条件查找给定字符串 `RangeAreas` 的所有匹配项，并作为包含一个或多个矩形区域的对象返回。|
||[findAllOrNullObject(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-findallornullobject-member(1))|根据指定的条件查找给定字符串 `RangeAreas` 的所有匹配项，并作为包含一个或多个矩形区域的对象返回。|
||[getRanges(address?: string)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getranges-member(1))|`RangeAreas`获取表示由地址或名称指定的一个或多个矩形区域块的对象。|
||[horizontalPageBreaks](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-horizontalpagebreaks-member)|获取工作表的水平分页符集合。|
||[onFormatChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onformatchanged-member)|在特定工作表上更改格式时发生。|
||[pageLayout](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-pagelayout-member)|`PageLayout`获取工作表的对象。|
||[replaceAll(text: string, replacement: string, criteria: Excel.ReplaceCriteria)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-replaceall-member(1))|根据当前工作表中指定的条件查找并替换给定的字符串。|
||[shapes](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-shapes-member)|返回工作表上的所有 Shape 对象的集合。|
||[verticalPageBreaks](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-verticalpagebreaks-member)|获取工作表的垂直分页符集合。|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[details](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-details-member)|表示有关更改详细信息的信息。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onChanged](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onchanged-member)|在更改工作簿中的任何工作表时发生。|
||[onFormatChanged](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onformatchanged-member)|当工作簿中任何工作表的格式发生更改时发生。|
||[onSelectionChanged](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onselectionchanged-member)|在任何工作表上更改选择时发生。|
|[WorksheetFormatChangedEventArgs](/javascript/api/excel/excel.worksheetformatchangedeventargs)|[address](/javascript/api/excel/excel.worksheetformatchangedeventargs#excel-excel-worksheetformatchangedeventargs-address-member)|获取区域地址，该地址表示特定工作表上的更改区域。|
||[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#excel-excel-worksheetformatchangedeventargs-getrange-member(1))|获取区域，该区域表示特定工作表上的更改区域。|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#excel-excel-worksheetformatchangedeventargs-getrangeornullobject-member(1))|获取区域，该区域表示特定工作表上的更改区域。|
||[source](/javascript/api/excel/excel.worksheetformatchangedeventargs#excel-excel-worksheetformatchangedeventargs-source-member)|获取事件源。|
||[type](/javascript/api/excel/excel.worksheetformatchangedeventargs#excel-excel-worksheetformatchangedeventargs-type-member)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.worksheetformatchangedeventargs#excel-excel-worksheetformatchangedeventargs-worksheetid-member)|获取其中的数据发生更改的工作表的 ID。|
|[WorksheetSearchCriteria](/javascript/api/excel/excel.worksheetsearchcriteria)|[completeMatch](/javascript/api/excel/excel.worksheetsearchcriteria#excel-excel-worksheetsearchcriteria-completematch-member)|指定匹配是需要完成还是部分匹配。|
||[matchCase](/javascript/api/excel/excel.worksheetsearchcriteria#excel-excel-worksheetsearchcriteria-matchcase-member)|指定匹配是否区分大小写。|

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 参考文档](/javascript/api/excel?view=excel-js-1.9&preserve-view=true)
- [Excel JavaScript API 要求集](excel-api-requirement-sets.md)
