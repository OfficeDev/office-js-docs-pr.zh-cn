---
title: ExcelJavaScript API 要求集 1.9
description: 有关 ExcelApi 1.9 要求集的详细信息。
ms.date: 04/01/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: eb917ed75049f965178075f57e8d0e9e7630bc9081019763e7812b221a00f67c
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/07/2021
ms.locfileid: "57098285"
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

下表列出了 JavaScript API 要求集 1.9 Excel中的 API。 若要查看受 Excel JavaScript API 要求集 1.9 或更早版本支持的所有 API 的 API 参考文档，请参阅要求集[1.9](/javascript/api/excel?view=excel-js-1.9&preserve-view=true)或更早中的 Excel API。

| 类 | 域 | 说明 |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[calculationEngineVersion](/javascript/api/excel/excel.application#calculationEngineVersion)|返回用于上次完整重新计算的 Excel 计算引擎版本。|
||[calculationState](/javascript/api/excel/excel.application#calculationState)|返回应用程序的计算状态。|
||[iterativeCalculation](/javascript/api/excel/excel.application#iterativeCalculation)|返回迭代计算设置。|
||[suspendScreenUpdatingUntilNextSync()](/javascript/api/excel/excel.application#suspendScreenUpdatingUntilNextSync__)|暂停屏幕更新，直到调用下一 `context.sync()` 个。|
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[apply(range: Range \| string, columnIndex?: number, criteria?: Excel.FilterCriteria)](/javascript/api/excel/excel.autofilter#apply_range__columnIndex__criteria_)|将自动筛选器应用于区域。|
||[clearCriteria()](/javascript/api/excel/excel.autofilter#clearCriteria__)|清除自动筛选器的筛选条件。|
||[getRange()](/javascript/api/excel/excel.autofilter#getRange__)|返回 `Range` 一个对象，该对象代表应用自动筛选的范围。|
||[getRangeOrNullObject()](/javascript/api/excel/excel.autofilter#getRangeOrNullObject__)|返回 `Range` 一个对象，该对象代表应用自动筛选的范围。|
||[criteria](/javascript/api/excel/excel.autofilter#criteria)|在自动筛选区域中保留所有筛选条件的数组。|
||[enabled](/javascript/api/excel/excel.autofilter#enabled)|指定是否启用自动筛选。|
||[isDataFiltered](/javascript/api/excel/excel.autofilter#isDataFiltered)|指定自动筛选是否具有筛选条件。|
||[reapply()](/javascript/api/excel/excel.autofilter#reapply__)|应用当前位于区域上的指定 Autofilter 对象。|
||[remove()](/javascript/api/excel/excel.autofilter#remove__)|删除区域的自动筛选。|
|[CellBorder](/javascript/api/excel/excel.cellborder)|[color](/javascript/api/excel/excel.cellborder#color)|表示`color`单个边框的属性。|
||[style](/javascript/api/excel/excel.cellborder#style)|表示`style`单个边框的属性。|
||[tintAndShade](/javascript/api/excel/excel.cellborder#tintAndShade)|表示`tintAndShade`单个边框的属性。|
||[weight](/javascript/api/excel/excel.cellborder#weight)|表示`weight`单个边框的属性。|
|[CellBorderCollection](/javascript/api/excel/excel.cellbordercollection)|[bottom](/javascript/api/excel/excel.cellbordercollection#bottom)|表示`format.borders.bottom`属性。|
||[diagonalDown](/javascript/api/excel/excel.cellbordercollection#diagonalDown)|表示`format.borders.diagonalDown`属性。|
||[diagonalUp](/javascript/api/excel/excel.cellbordercollection#diagonalUp)|表示`format.borders.diagonalUp`属性。|
||[horizontal](/javascript/api/excel/excel.cellbordercollection#horizontal)|表示`format.borders.horizontal`属性。|
||[left](/javascript/api/excel/excel.cellbordercollection#left)|表示`format.borders.left`属性。|
||[right](/javascript/api/excel/excel.cellbordercollection#right)|表示`format.borders.right`属性。|
||[top](/javascript/api/excel/excel.cellbordercollection#top)|表示`format.borders.top`属性。|
||[vertical](/javascript/api/excel/excel.cellbordercollection#vertical)|表示`format.borders.vertical`属性。|
|[CellProperties](/javascript/api/excel/excel.cellproperties)|[address](/javascript/api/excel/excel.cellproperties#address)|表示`address`属性。|
||[addressLocal](/javascript/api/excel/excel.cellproperties#addressLocal)|表示`addressLocal`属性。|
||[hidden](/javascript/api/excel/excel.cellproperties#hidden)|表示`hidden`属性。|
|[CellPropertiesFill](/javascript/api/excel/excel.cellpropertiesfill)|[color](/javascript/api/excel/excel.cellpropertiesfill#color)|表示`format.fill.color`属性。|
||[pattern](/javascript/api/excel/excel.cellpropertiesfill#pattern)|表示`format.fill.pattern`属性。|
||[patternColor](/javascript/api/excel/excel.cellpropertiesfill#patternColor)|表示`format.fill.patternColor`属性。|
||[patternTintAndShade](/javascript/api/excel/excel.cellpropertiesfill#patternTintAndShade)|表示`format.fill.patternTintAndShade`属性。|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfill#tintAndShade)|表示`format.fill.tintAndShade`属性。|
|[CellPropertiesFont](/javascript/api/excel/excel.cellpropertiesfont)|[bold](/javascript/api/excel/excel.cellpropertiesfont#bold)|表示`format.font.bold`属性。|
||[color](/javascript/api/excel/excel.cellpropertiesfont#color)|表示`format.font.color`属性。|
||[italic](/javascript/api/excel/excel.cellpropertiesfont#italic)|表示`format.font.italic`属性。|
||[name](/javascript/api/excel/excel.cellpropertiesfont#name)|表示`format.font.name`属性。|
||[size](/javascript/api/excel/excel.cellpropertiesfont#size)|表示`format.font.size`属性。|
||[strikethrough](/javascript/api/excel/excel.cellpropertiesfont#strikethrough)|表示`format.font.strikethrough`属性。|
||[subscript](/javascript/api/excel/excel.cellpropertiesfont#subscript)|表示`format.font.subscript`属性。|
||[superscript](/javascript/api/excel/excel.cellpropertiesfont#superscript)|表示`format.font.superscript`属性。|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfont#tintAndShade)|表示`format.font.tintAndShade`属性。|
||[underline](/javascript/api/excel/excel.cellpropertiesfont#underline)|表示`format.font.underline`属性。|
|[CellPropertiesFormat](/javascript/api/excel/excel.cellpropertiesformat)|[autoIndent](/javascript/api/excel/excel.cellpropertiesformat#autoIndent)|表示`autoIndent`属性。|
||[Borders](/javascript/api/excel/excel.cellpropertiesformat#borders)|表示`borders`属性。|
||[fill](/javascript/api/excel/excel.cellpropertiesformat#fill)|表示`fill`属性。|
||[font](/javascript/api/excel/excel.cellpropertiesformat#font)|表示`font`属性。|
||[horizontalAlignment](/javascript/api/excel/excel.cellpropertiesformat#horizontalAlignment)|表示`horizontalAlignment`属性。|
||[indentLevel](/javascript/api/excel/excel.cellpropertiesformat#indentLevel)|表示`indentLevel`属性。|
||[protection](/javascript/api/excel/excel.cellpropertiesformat#protection)|表示`protection`属性。|
||[readingOrder](/javascript/api/excel/excel.cellpropertiesformat#readingOrder)|表示`readingOrder`属性。|
||[shrinkToFit](/javascript/api/excel/excel.cellpropertiesformat#shrinkToFit)|表示`shrinkToFit`属性。|
||[textOrientation](/javascript/api/excel/excel.cellpropertiesformat#textOrientation)|表示`textOrientation`属性。|
||[useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#useStandardHeight)|表示`useStandardHeight`属性。|
||[useStandardWidth](/javascript/api/excel/excel.cellpropertiesformat#useStandardWidth)|表示`useStandardWidth`属性。|
||[verticalAlignment](/javascript/api/excel/excel.cellpropertiesformat#verticalAlignment)|表示`verticalAlignment`属性。|
||[wrapText](/javascript/api/excel/excel.cellpropertiesformat#wrapText)|表示`wrapText`属性。|
|[CellPropertiesProtection](/javascript/api/excel/excel.cellpropertiesprotection)|[formulaHidden](/javascript/api/excel/excel.cellpropertiesprotection#formulaHidden)|表示`format.protection.formulaHidden`属性。|
||[locked](/javascript/api/excel/excel.cellpropertiesprotection#locked)|表示`format.protection.locked`属性。|
|[ChangedEventDetail](/javascript/api/excel/excel.changedeventdetail)|[valueAfter](/javascript/api/excel/excel.changedeventdetail#valueAfter)|表示更改后的值。|
||[valueBefore](/javascript/api/excel/excel.changedeventdetail#valueBefore)|表示更改前的值。|
||[valueTypeAfter](/javascript/api/excel/excel.changedeventdetail#valueTypeAfter)|表示更改后的值的类型。|
||[valueTypeBefore](/javascript/api/excel/excel.changedeventdetail#valueTypeBefore)|表示更改之前的值的类型。|
|[Chart](/javascript/api/excel/excel.chart)|[activate()](/javascript/api/excel/excel.chart#activate__)|在 Excel UI 中激活图表。|
||[pivotOptions](/javascript/api/excel/excel.chart#pivotOptions)|封装数据透视图的选项。|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[colorScheme](/javascript/api/excel/excel.chartareaformat#colorScheme)|指定图表的配色方案。|
||[roundedCorners](/javascript/api/excel/excel.chartareaformat#roundedCorners)|指定图表的图表区域是否具有圆角。|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[linkNumberFormat](/javascript/api/excel/excel.chartaxis#linkNumberFormat)|指定数字格式是否链接到单元格。|
|[ChartBinOptions](/javascript/api/excel/excel.chartbinoptions)|[allowOverflow](/javascript/api/excel/excel.chartbinoptions#allowOverflow)|指定在直方图或直方图中是否启用箱溢出。|
||[allowUnderflow](/javascript/api/excel/excel.chartbinoptions#allowUnderflow)|指定在直方图或流程图中是否启用箱下溢。|
||[count](/javascript/api/excel/excel.chartbinoptions#count)|指定直方图或流程图的箱计数。|
||[overflowValue](/javascript/api/excel/excel.chartbinoptions#overflowValue)|指定直方图或流程图的箱溢出值。|
||[type](/javascript/api/excel/excel.chartbinoptions#type)|指定直方图或流程图的箱类型。|
||[underflowValue](/javascript/api/excel/excel.chartbinoptions#underflowValue)|指定直方图或流程图的箱下溢值。|
||[width](/javascript/api/excel/excel.chartbinoptions#width)|指定直方图或 pareto 图表的箱宽度值。|
|[ChartBoxwhiskerOptions](/javascript/api/excel/excel.chartboxwhiskeroptions)|[quartileCalculation](/javascript/api/excel/excel.chartboxwhiskeroptions#quartileCalculation)|指定箱形图的四分位数计算类型。|
||[showInnerPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#showInnerPoints)|指定内部点是否显示在箱形图中。|
||[showMeanLine](/javascript/api/excel/excel.chartboxwhiskeroptions#showMeanLine)|指定在箱形图中是否显示平均值。|
||[showMeanMarker](/javascript/api/excel/excel.chartboxwhiskeroptions#showMeanMarker)|指定是否将平均值标记显示在箱形图中。|
||[showOutlierPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#showOutlierPoints)|指定在箱形图中是否显示离群值点。|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabel#linkNumberFormat)|指定数字格式是否链接到单元格 (以便当数字格式在单元格区域更改时标签) 。|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabels#linkNumberFormat)|指定数字格式是否链接到单元格。|
|[ChartErrorBars](/javascript/api/excel/excel.charterrorbars)|[endStyleCap](/javascript/api/excel/excel.charterrorbars#endStyleCap)|指定误差线是否具有结束样式上限。|
||[include](/javascript/api/excel/excel.charterrorbars#include)|指定包含误差线的哪些部分。|
||[format](/javascript/api/excel/excel.charterrorbars#format)|指定误差线的格式类型。|
||[type](/javascript/api/excel/excel.charterrorbars#type)|误差线标记的区域类型。|
||[visible](/javascript/api/excel/excel.charterrorbars#visible)|指定是否显示误差线。|
|[ChartErrorBarsFormat](/javascript/api/excel/excel.charterrorbarsformat)|[line](/javascript/api/excel/excel.charterrorbarsformat#line)|表示图表线条格式。|
|[ChartMapOptions](/javascript/api/excel/excel.chartmapoptions)|[labelStrategy](/javascript/api/excel/excel.chartmapoptions#labelStrategy)|指定区域地图图表的系列地图标签策略。|
||[level](/javascript/api/excel/excel.chartmapoptions#level)|指定区域地图图表的系列映射级别。|
||[projectionType](/javascript/api/excel/excel.chartmapoptions#projectionType)|指定区域地图图表的系列投影类型。|
|[ChartPivotOptions](/javascript/api/excel/excel.chartpivotoptions)|[showAxisFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showAxisFieldButtons)|指定是否在坐标轴上显示坐标轴字段数据透视图。|
||[showLegendFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showLegendFieldButtons)|指定是否在图例上显示图例数据透视图。|
||[showReportFilterFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showReportFilterFieldButtons)|指定是否在报表上显示报表筛选字段数据透视图。|
||[showValueFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showValueFieldButtons)|指定是否在项目上显示"显示值"字段数据透视图。|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[bubbleScale](/javascript/api/excel/excel.chartseries#bubbleScale)|这可以是从 0（零）到 300 的整数值，表示默认大小的百分比。|
||[gradientMaximumColor](/javascript/api/excel/excel.chartseries#gradientMaximumColor)|指定区域地图图表系列的最大值的颜色。|
||[gradientMaximumType](/javascript/api/excel/excel.chartseries#gradientMaximumType)|指定区域地图图表系列的最大值的类型。|
||[gradientMaximumValue](/javascript/api/excel/excel.chartseries#gradientMaximumValue)|指定区域地图图表系列的最大值。|
||[gradientMidpointColor](/javascript/api/excel/excel.chartseries#gradientMidpointColor)|指定区域地图图表系列的中点值的颜色。|
||[gradientMidpointType](/javascript/api/excel/excel.chartseries#gradientMidpointType)|指定区域地图图表系列的中点值的类型。|
||[gradientMidpointValue](/javascript/api/excel/excel.chartseries#gradientMidpointValue)|指定区域地图图表系列的中点值。|
||[gradientMinimumColor](/javascript/api/excel/excel.chartseries#gradientMinimumColor)|指定区域地图图表系列的最小值的颜色。|
||[gradientMinimumType](/javascript/api/excel/excel.chartseries#gradientMinimumType)|指定区域地图图表系列的最小值的类型。|
||[gradientMinimumValue](/javascript/api/excel/excel.chartseries#gradientMinimumValue)|指定区域地图图表系列的最小值。|
||[gradientStyle](/javascript/api/excel/excel.chartseries#gradientStyle)|指定区域地图图表的系列渐变样式。|
||[invertColor](/javascript/api/excel/excel.chartseries#invertColor)|指定系列中负数据点的填充颜色。|
||[parentLabelStrategy](/javascript/api/excel/excel.chartseries#parentLabelStrategy)|指定树图的系列父标签策略区域。|
||[binOptions](/javascript/api/excel/excel.chartseries#binOptions)|封装直方图和排列图的容器选项。|
||[boxwhiskerOptions](/javascript/api/excel/excel.chartseries#boxwhiskerOptions)|封装箱形图的选项。|
||[mapOptions](/javascript/api/excel/excel.chartseries#mapOptions)|封装区域地图图表的选项。|
||[xErrorBars](/javascript/api/excel/excel.chartseries#xErrorBars)|表示图表系列的误差线对象。|
||[yErrorBars](/javascript/api/excel/excel.chartseries#yErrorBars)|表示图表系列的误差线对象。|
||[showConnectorLines](/javascript/api/excel/excel.chartseries#showConnectorLines)|指定是否在瀑布图中显示连接线。|
||[showLeaderLines](/javascript/api/excel/excel.chartseries#showLeaderLines)|指定是否为系列中每个数据标签显示前导线。|
||[splitValue](/javascript/api/excel/excel.chartseries#splitValue)|指定分隔复合饼图或复合条饼图的两个部分的阈值。|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[linkNumberFormat](/javascript/api/excel/excel.charttrendlinelabel#linkNumberFormat)|指定数字格式是否链接到单元格 (以便当数字格式在单元格区域更改时标签) 。|
|[ColumnProperties](/javascript/api/excel/excel.columnproperties)|[address](/javascript/api/excel/excel.columnproperties#address)|表示`address`属性。|
||[addressLocal](/javascript/api/excel/excel.columnproperties#addressLocal)|表示`addressLocal`属性。|
||[columnIndex](/javascript/api/excel/excel.columnproperties#columnIndex)|表示`columnIndex`属性。|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[getRanges()](/javascript/api/excel/excel.conditionalformat#getRanges__)|返回 `RangeAreas` ，包含一个或多个矩形区域，其中应用了条件格式。|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[getInvalidCells()](/javascript/api/excel/excel.datavalidation#getInvalidCells__)|返回包含 `RangeAreas` 一个或多个矩形区域的对象，其中单元格值无效。|
||[getInvalidCellsOrNullObject()](/javascript/api/excel/excel.datavalidation#getInvalidCellsOrNullObject__)|返回包含 `RangeAreas` 一个或多个矩形区域的对象，其中单元格值无效。|
|[FilterCriteria](/javascript/api/excel/excel.filtercriteria)|[subField](/javascript/api/excel/excel.filtercriteria#subField)|筛选器用于对丰富值执行丰富筛选器的属性。|
|[GeometricShape](/javascript/api/excel/excel.geometricshape)|[id](/javascript/api/excel/excel.geometricshape#id)|返回形状标识符。|
||[shape](/javascript/api/excel/excel.geometricshape#shape)|返回 `Shape` 几何形状的对象。|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getCount()](/javascript/api/excel/excel.groupshapecollection#getCount__)|返回形状组中的形状数量。|
||[getItem(key: string)](/javascript/api/excel/excel.groupshapecollection#getItem_key_)|使用形状的名称或 ID 获取形状。|
||[getItemAt(index: number)](/javascript/api/excel/excel.groupshapecollection#getItemAt_index_)|根据其在集合中的位置获取形状。|
||[items](/javascript/api/excel/excel.groupshapecollection#items)|获取此集合中已加载的子项。|
|[HeaderFooter](/javascript/api/excel/excel.headerfooter)|[centerFooter](/javascript/api/excel/excel.headerfooter#centerFooter)|工作表的中心页脚。|
||[centerHeader](/javascript/api/excel/excel.headerfooter#centerHeader)|工作表的中心页眉。|
||[LeftFooter](/javascript/api/excel/excel.headerfooter#leftFooter)|工作表的左页脚。|
||[leftHeader](/javascript/api/excel/excel.headerfooter#leftHeader)|工作表的左页眉。|
||[rightFooter](/javascript/api/excel/excel.headerfooter#rightFooter)|工作表的右页脚。|
||[rightHeader](/javascript/api/excel/excel.headerfooter#rightHeader)|工作表的右页眉。|
|[HeaderFooterGroup](/javascript/api/excel/excel.headerfootergroup)|[defaultForAllPages](/javascript/api/excel/excel.headerfootergroup#defaultForAllPages)|常规页眉/页脚，除非指定偶数页/奇数页或首页，否则适用于所有页面，|
||[evenPages](/javascript/api/excel/excel.headerfootergroup#evenPages)|用于偶数页的页眉/页脚，需要为奇数页指定奇数页页眉/页脚。|
||[firstPage](/javascript/api/excel/excel.headerfootergroup#firstPage)|首页的页眉/页脚，为所有其他页使用常规或偶数页/奇数页页眉/页脚。|
||[oddPages](/javascript/api/excel/excel.headerfootergroup#oddPages)|用于奇数页的页眉/页脚，需要为偶数页指定偶数页页眉/页脚。|
||[state](/javascript/api/excel/excel.headerfootergroup#state)|设置页眉/页脚时所按的状态。|
||[useSheetMargins](/javascript/api/excel/excel.headerfootergroup#useSheetMargins)|获取或设置一个标记，指示页眉/页脚是否与工作表的页面布局选项中设置的页边距对齐。|
||[useSheetScale](/javascript/api/excel/excel.headerfootergroup#useSheetScale)|获取或设置一个标记，指示是否应按照工作表的页面布局选项中设置的页面缩放百分比来缩放页眉/页脚。|
|[Image](/javascript/api/excel/excel.image)|[format](/javascript/api/excel/excel.image#format)|返回图像的格式。|
||[id](/javascript/api/excel/excel.image#id)|指定图像对象的形状标识符。|
||[shape](/javascript/api/excel/excel.image#shape)|返回 `Shape` 与图像关联的对象。|
|[IterativeCalculation](/javascript/api/excel/excel.iterativecalculation)|[enabled](/javascript/api/excel/excel.iterativecalculation#enabled)|如果 Excel 使用迭代来处理循环引用，则为 True。|
||[maxChange](/javascript/api/excel/excel.iterativecalculation#maxChange)|指定每个迭代之间的最大更改量，因为Excel循环引用。|
||[maxIteration](/javascript/api/excel/excel.iterativecalculation#maxIteration)|指定可用于解析循环引用Excel迭代的最大次数。|
|[Line](/javascript/api/excel/excel.line)|[beginArrowheadLength](/javascript/api/excel/excel.line#beginArrowheadLength)|表示指定线条始端的箭头长度。|
||[beginArrowheadStyle](/javascript/api/excel/excel.line#beginArrowheadStyle)|表示指定线条始端的箭头样式。|
||[beginArrowheadWidth](/javascript/api/excel/excel.line#beginArrowheadWidth)|表示指定线条始端的箭头宽度。|
||[connectBeginShape(shape: Excel.Shape, connectionSite: number)](/javascript/api/excel/excel.line#connectBeginShape_shape__connectionSite_)|将指定连接线的始端附加到指定形状。|
||[connectEndShape(shape: Excel.Shape, connectionSite: number)](/javascript/api/excel/excel.line#connectEndShape_shape__connectionSite_)|将指定连接线的末端附加到指定形状。|
||[connectorType](/javascript/api/excel/excel.line#connectorType)|表示线条的连接器类型。|
||[disconnectBeginShape()](/javascript/api/excel/excel.line#disconnectBeginShape__)|使指定连接线的始端与形状脱离。|
||[disconnectEndShape()](/javascript/api/excel/excel.line#disconnectEndShape__)|使指定连接线的末端与形状脱离。|
||[endArrowheadLength](/javascript/api/excel/excel.line#endArrowheadLength)|表示指定线条末端的箭头长度。|
||[endArrowheadStyle](/javascript/api/excel/excel.line#endArrowheadStyle)|表示指定线条末端的箭头样式。|
||[endArrowheadWidth](/javascript/api/excel/excel.line#endArrowheadWidth)|表示指定线条末端的箭头宽度。|
||[beginConnectedShape](/javascript/api/excel/excel.line#beginConnectedShape)|表示指定线条始端所附加到的形状。|
||[beginConnectedSite](/javascript/api/excel/excel.line#beginConnectedSite)|表示连接线始端所连接的连接站点。|
||[endConnectedShape](/javascript/api/excel/excel.line#endConnectedShape)|表示指定线条末端所附加到的形状。|
||[endConnectedSite](/javascript/api/excel/excel.line#endConnectedSite)|表示连接线末端所连接的连接站点。|
||[id](/javascript/api/excel/excel.line#id)|指定形状标识符。|
||[isBeginConnected](/javascript/api/excel/excel.line#isBeginConnected)|指定指定线条的始值是否连接到形状。|
||[isEndConnected](/javascript/api/excel/excel.line#isEndConnected)|指定指定线条的终点是否连接到形状。|
||[shape](/javascript/api/excel/excel.line#shape)|返回 `Shape` 与线条关联的对象。|
|[PageBreak](/javascript/api/excel/excel.pagebreak)|[delete()](/javascript/api/excel/excel.pagebreak#delete__)|删除分页符对象。|
||[getCellAfterBreak()](/javascript/api/excel/excel.pagebreak#getCellAfterBreak__)|获取分页符后的第一个单元格。|
||[columnIndex](/javascript/api/excel/excel.pagebreak#columnIndex)|指定分页符的列索引。|
||[rowIndex](/javascript/api/excel/excel.pagebreak#rowIndex)|指定分页符的行索引。|
|[PageBreakCollection](/javascript/api/excel/excel.pagebreakcollection)|[add(pageBreakRange: Range \| string)](/javascript/api/excel/excel.pagebreakcollection#add_pageBreakRange_)|在指定区域的左上角单元格之前添加分页符。|
||[getCount()](/javascript/api/excel/excel.pagebreakcollection#getCount__)|获取集合中的分页符数量。|
||[getItem(index: number)](/javascript/api/excel/excel.pagebreakcollection#getItem_index_)|通过索引获取分页符对象。|
||[items](/javascript/api/excel/excel.pagebreakcollection#items)|获取此集合中已加载的子项。|
||[removePageBreaks()](/javascript/api/excel/excel.pagebreakcollection#removePageBreaks__)|重置集合中的所有手动分页符。|
|[PageLayout](/javascript/api/excel/excel.pagelayout)|[blackAndWhite](/javascript/api/excel/excel.pagelayout#blackAndWhite)|工作表的黑白打印选项。|
||[bottomMargin](/javascript/api/excel/excel.pagelayout#bottomMargin)|要用于打印的工作表的底部页边距（以点表示）。|
||[centerHorizontally](/javascript/api/excel/excel.pagelayout#centerHorizontally)|工作表的中心水平标记。|
||[centerVertically](/javascript/api/excel/excel.pagelayout#centerVertically)|工作表的中心垂直标记。|
||[draftMode](/javascript/api/excel/excel.pagelayout#draftMode)|工作表的草稿模式选项。|
||[firstPageNumber](/javascript/api/excel/excel.pagelayout#firstPageNumber)|要打印的工作表的第一页码。|
||[footerMargin](/javascript/api/excel/excel.pagelayout#footerMargin)|打印时工作表的页脚边距（以点表示）。|
||[getPrintArea()](/javascript/api/excel/excel.pagelayout#getPrintArea__)|获取包含一个或多个矩形区域的对象， `RangeAreas` 该对象表示工作表的打印区域。|
||[getPrintAreaOrNullObject()](/javascript/api/excel/excel.pagelayout#getPrintAreaOrNullObject__)|获取包含一个或多个矩形区域的对象， `RangeAreas` 该对象表示工作表的打印区域。|
||[getPrintTitleColumns()](/javascript/api/excel/excel.pagelayout#getPrintTitleColumns__)|获取表示标题列的 Range 对象。|
||[getPrintTitleColumnsOrNullObject()](/javascript/api/excel/excel.pagelayout#getPrintTitleColumnsOrNullObject__)|获取表示标题列的 Range 对象。|
||[getPrintTitleRows()](/javascript/api/excel/excel.pagelayout#getPrintTitleRows__)|获取表示标题行的 Range 对象。|
||[getPrintTitleRowsOrNullObject()](/javascript/api/excel/excel.pagelayout#getPrintTitleRowsOrNullObject__)|获取表示标题行的 Range 对象。|
||[headerMargin](/javascript/api/excel/excel.pagelayout#headerMargin)|打印时工作表的页眉边距（以点表示）。|
||[leftMargin](/javascript/api/excel/excel.pagelayout#leftMargin)|打印时工作表的左边距（以点表示）。|
||[orientation](/javascript/api/excel/excel.pagelayout#orientation)|工作表的页面方向。|
||[paperSize](/javascript/api/excel/excel.pagelayout#paperSize)|工作表的页面纸张大小。|
||[printComments](/javascript/api/excel/excel.pagelayout#printComments)|指定打印时是否显示工作表的注释。|
||[printErrors](/javascript/api/excel/excel.pagelayout#printErrors)|工作表的打印错误选项。|
||[printGridlines](/javascript/api/excel/excel.pagelayout#printGridlines)|指定是否打印工作表的网格线。|
||[printHeadings](/javascript/api/excel/excel.pagelayout#printHeadings)|指定是否打印工作表的标题。|
||[printOrder](/javascript/api/excel/excel.pagelayout#printOrder)|工作表的页面打印顺序选项。|
||[headersFooters](/javascript/api/excel/excel.pagelayout#headersFooters)|工作表的页眉和页脚配置。|
||[rightMargin](/javascript/api/excel/excel.pagelayout#rightMargin)|打印时工作表的右边距（以点表示）。|
||[setPrintArea(printArea: Range \| RangeAreas \| string)](/javascript/api/excel/excel.pagelayout#setPrintArea_printArea_)|设置工作表的打印区域。|
||[setPrintMargins(unit: Excel.PrintMarginUnit, marginOptions: Excel.PageLayoutMarginOptions)](/javascript/api/excel/excel.pagelayout#setPrintMargins_unit__marginOptions_)|设置带单位的工作表的页边距。|
||[setPrintTitleColumns(printTitleColumns: Range \| string)](/javascript/api/excel/excel.pagelayout#setPrintTitleColumns_printTitleColumns_)|设置列，这些列包含要在打印的工作表的每页左侧重复的单元格。|
||[setPrintTitleRows(printTitleRows: Range \| string)](/javascript/api/excel/excel.pagelayout#setPrintTitleRows_printTitleRows_)|设置行，这些行包含要在打印的工作表的每页顶部重复的单元格。|
||[topMargin](/javascript/api/excel/excel.pagelayout#topMargin)|打印时工作表的上边距（以点表示）。|
||[zoom](/javascript/api/excel/excel.pagelayout#zoom)|工作表的打印缩放选项。|
|[PageLayoutMarginOptions](/javascript/api/excel/excel.pagelayoutmarginoptions)|[bottom](/javascript/api/excel/excel.pagelayoutmarginoptions#bottom)|指定要用于打印的页面布局下边距（以指定单位表示）。|
||[footer](/javascript/api/excel/excel.pagelayoutmarginoptions#footer)|指定要用于打印的页面布局页脚边距（以指定单位表示）。|
||[header](/javascript/api/excel/excel.pagelayoutmarginoptions#header)|指定要用于打印的页面布局页眉边距（以指定单位表示）。|
||[left](/javascript/api/excel/excel.pagelayoutmarginoptions#left)|指定要用于打印的页面布局左边距（以指定单位表示）。|
||[right](/javascript/api/excel/excel.pagelayoutmarginoptions#right)|指定要用于打印的页面布局右边距（以指定单位表示）。|
||[top](/javascript/api/excel/excel.pagelayoutmarginoptions#top)|指定要用于打印的页面布局上边距（以指定单位表示）。|
|[PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions)|[horizontalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#horizontalFitToPages)|水平放置的页数。|
||[scale](/javascript/api/excel/excel.pagelayoutzoomoptions#scale)|打印页面缩放值可以介于 10 至 400 之间。|
||[verticalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#verticalFitToPages)|垂直放置的页数。|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[sortByValues(sortBy: Excel.SortBy, valuesHierarchy: Excel.DataPivotHierarchy, pivotItemScope?: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotfield#sortByValues_sortBy__valuesHierarchy__pivotItemScope_)|按给定范围中的指定值对 PivotField 进行排序。|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[autoFormat](/javascript/api/excel/excel.pivotlayout#autoFormat)|指定在刷新格式或移动字段时是否自动设置格式。|
||[getDataHierarchy(cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#getDataHierarchy_cell_)|获取 DataHierarchy，它用于计算数据透视表中指定区域内的值。|
||[getPivotItems(axis: Excel.PivotAxis, cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#getPivotItems_axis__cell_)|从构成数据透视表中指定区域内的值的轴获取 PivotItems。|
||[preserveFormatting](/javascript/api/excel/excel.pivotlayout#preserveFormatting)|指定在通过透视、排序或更改页字段项等操作刷新或重新计算报表时是否保留格式。|
||[setAutoSortOnCell(cell: Range \| string, sortBy: Excel.SortBy)](/javascript/api/excel/excel.pivotlayout#setAutoSortOnCell_cell__sortBy_)|将数据透视表设置为使用指定的单元格设置自动排序，以自动选择排序的所有条件和上下文。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[enableDataValueEditing](/javascript/api/excel/excel.pivottable#enableDataValueEditing)|指定数据透视表是否允许用户编辑数据正文中的值。|
||[useCustomSortLists](/javascript/api/excel/excel.pivottable#useCustomSortLists)|指定数据透视表在排序时是否使用自定义列表。|
|[区域](/javascript/api/excel/excel.range)|[autoFill (destinationRange？： Range \| string， autoFillType？： Excel。AutoFillType) ](/javascript/api/excel/excel.range#autoFill_destinationRange__autoFillType_)|使用指定的自动填充逻辑填充从当前区域到目标区域的范围。|
||[convertDataTypeToText()](/javascript/api/excel/excel.range#convertDataTypeToText__)|将数据类型为区域单元格转换为文本。|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.range#convertToLinkedDataType_serviceID__languageCulture_)|将区域单元格转换为工作表中的链接数据类型。|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.range#copyFrom_sourceRange__copyType__skipBlanks__transpose_)|将单元格数据或格式从源区域或 `RangeAreas` 当前区域复制。|
||[find(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#find_text__criteria_)|根据指定的条件查找给定的字符串。|
||[findOrNullObject(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#findOrNullObject_text__criteria_)|根据指定的条件查找给定的字符串。|
||[flashFill()](/javascript/api/excel/excel.range#flashFill__)|对当前范围进行快速填充。|
||[getCellProperties(cellPropertiesLoadOptions: CellPropertiesLoadOptions)](/javascript/api/excel/excel.range#getCellProperties_cellPropertiesLoadOptions_)|返回一个 2D 数组，其中封装了每个单元格的字体、填充、边框、对齐方式和其他属性数据。|
||[getColumnProperties(columnPropertiesLoadOptions: ColumnPropertiesLoadOptions)](/javascript/api/excel/excel.range#getColumnProperties_columnPropertiesLoadOptions_)|返回一个一维数组，其中封装了每个列的字体、填充、边框、对齐方式和其他属性数据。|
||[getRowProperties(rowPropertiesLoadOptions: RowPropertiesLoadOptions)](/javascript/api/excel/excel.range#getRowProperties_rowPropertiesLoadOptions_)|返回一个一维数组，其中封装了每个行的字体、填充、边框、对齐方式和其他属性数据。|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#getSpecialCells_cellType__cellValueType_)|获取包含一个或多个矩形区域的对象，该对象表示与指定类型和值 `RangeAreas` 匹配的所有单元格。|
||[getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#getSpecialCellsOrNullObject_cellType__cellValueType_)|获取包含一个或多个区域的对象，该对象代表与指定类型和值 `RangeAreas` 匹配的所有单元格。|
||[getTables(fullyContained?: boolean)](/javascript/api/excel/excel.range#getTables_fullyContained_)|获取与区域重叠的限定范围的表格集合。|
||[linkedDataTypeState](/javascript/api/excel/excel.range#linkedDataTypeState)|表示每个单元格的数据类型状态。|
||[removeDuplicates(columns: number[], includesHeader: boolean)](/javascript/api/excel/excel.range#removeDuplicates_columns__includesHeader_)|从列指定的区域中删除重复值。|
||[replaceAll(text: string, replacement: string, criteria: Excel.ReplaceCriteria)](/javascript/api/excel/excel.range#replaceAll_text__replacement__criteria_)|根据当前区域内指定的条件查找并替换给定的字符串。|
||[setCellProperties(cellPropertiesData: SettableCellProperties[][])](/javascript/api/excel/excel.range#setCellProperties_cellPropertiesData_)|根据单元格属性的 2D 数组更新区域，并封装字体、填充、边框和对齐方式等内容。|
||[setColumnProperties(columnPropertiesData: SettableColumnProperties[])](/javascript/api/excel/excel.range#setColumnProperties_columnPropertiesData_)|根据列属性的一维数组更新区域，并封装字体、填充、边框和对齐方式等内容。|
||[setDirty()](/javascript/api/excel/excel.range#setDirty__)|设置下一次重新计算发生时要重新计算的区域。|
||[setRowProperties(rowPropertiesData: SettableRowProperties[])](/javascript/api/excel/excel.range#setRowProperties_rowPropertiesData_)|根据行属性的一维数组更新区域，并封装字体、填充、边框和对齐方式等内容。|
|[RangeAreas](/javascript/api/excel/excel.rangeareas)|[calculate()](/javascript/api/excel/excel.rangeareas#calculate__)|计算 中所有单元格 `RangeAreas` 。|
||[clear(applyTo?: Excel.ClearApplyTo)](/javascript/api/excel/excel.rangeareas#clear_applyTo_)|清除包含此对象的每个区域的值、格式、填充、边框和其他 `RangeAreas` 属性。|
||[convertDataTypeToText()](/javascript/api/excel/excel.rangeareas#convertDataTypeToText__)|将 数据类型为 `RangeAreas` 的 中所有单元格转换为文本。|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.rangeareas#convertToLinkedDataType_serviceID__languageCulture_)|将 中所有单元格 `RangeAreas` 转换为链接数据类型。|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.rangeareas#copyFrom_sourceRange__copyType__skipBlanks__transpose_)|将单元格数据或格式从源区域或 `RangeAreas` 复制到当前 `RangeAreas` 。|
||[getEntireColumn()](/javascript/api/excel/excel.rangeareas#getEntireColumn__)|返回一个对象，该对象代表 (例如，如果当前表示单元格 `RangeAreas` `RangeAreas` `RangeAreas` "B4：E11， H2"，则返回表示列 `RangeAreas` "B：E， H：H") 。|
||[getEntireRow()](/javascript/api/excel/excel.rangeareas#getEntireRow__)|返回一个对象，该对象代表 (例如，如果当前表示单元格 `RangeAreas` `RangeAreas` `RangeAreas` "B4：E11"，则返回表示行 `RangeAreas` "4：11") 。|
||[getIntersection(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#getIntersection_anotherRange_)|返回 `RangeAreas` 表示给定区域或 的交集的对象 `RangeAreas` 。|
||[getIntersectionOrNullObject(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#getIntersectionOrNullObject_anotherRange_)|返回 `RangeAreas` 表示给定区域或 的交集的对象 `RangeAreas` 。|
||[getOffsetRangeAreas(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.rangeareas#getOffsetRangeAreas_rowOffset__columnOffset_)|返回 `RangeAreas` 由特定行和列偏移移动的对象。|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#getSpecialCells_cellType__cellValueType_)|返回 `RangeAreas` 一个对象，该对象代表与指定类型和值匹配的所有单元格。|
||[getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#getSpecialCellsOrNullObject_cellType__cellValueType_)|返回 `RangeAreas` 一个对象，该对象代表与指定类型和值匹配的所有单元格。|
||[getTables(fullyContained?: boolean)](/javascript/api/excel/excel.rangeareas#getTables_fullyContained_)|返回与该对象中任何区域重叠的表的范围集合 `RangeAreas` 。|
||[getUsedRangeAreas(valuesOnly?: boolean)](/javascript/api/excel/excel.rangeareas#getUsedRangeAreas_valuesOnly_)|返回由 `RangeAreas` 对象中各个矩形区域的所有已用区域组成的 `RangeAreas` used。|
||[getUsedRangeAreasOrNullObject(valuesOnly?: boolean)](/javascript/api/excel/excel.rangeareas#getUsedRangeAreasOrNullObject_valuesOnly_)|返回由 `RangeAreas` 对象中各个矩形区域的所有已用区域组成的 `RangeAreas` used。|
||[address](/javascript/api/excel/excel.rangeareas#address)|以 `RangeAreas` A1 样式返回引用。|
||[addressLocal](/javascript/api/excel/excel.rangeareas#addressLocal)|返回 `RangeAreas` 用户区域设置中的引用。|
||[areaCount](/javascript/api/excel/excel.rangeareas#areaCount)|返回包含此对象的矩形区域 `RangeAreas` 的数量。|
||[areas](/javascript/api/excel/excel.rangeareas#areas)|返回包含此对象的矩形区域 `RangeAreas` 的集合。|
||[cellCount](/javascript/api/excel/excel.rangeareas#cellCount)|返回对象中的单元格数，并汇总所有单个矩形范围的 `RangeAreas` 单元格计数。|
||[conditionalFormats](/javascript/api/excel/excel.rangeareas#conditionalFormats)|返回一组条件格式，这些格式与该对象中任何单元格 `RangeAreas` 相交。|
||[dataValidation](/javascript/api/excel/excel.rangeareas#dataValidation)|返回 中所有区域的数据验证对象 `RangeAreas` 。|
||[format](/javascript/api/excel/excel.rangeareas#format)|返回一个对象，该对象封装对象中所有范围的字体、填充、边框、 `RangeFormat` 对齐方式和其他 `RangeAreas` 属性。|
||[isEntireColumn](/javascript/api/excel/excel.rangeareas#isEntireColumn)|指定此对象上的所有区域是否代表整个列 (例如 `RangeAreas` ，"A：C，Q：Z") 。|
||[isEntireRow](/javascript/api/excel/excel.rangeareas#isEntireRow)|指定此对象上的所有区域是否代表整个行 (例如 `RangeAreas` ，"1：3， 5：7") 。|
||[worksheet](/javascript/api/excel/excel.rangeareas#worksheet)|返回当前 的工作表 `RangeAreas` 。|
||[setDirty()](/javascript/api/excel/excel.rangeareas#setDirty__)|设置 `RangeAreas` 在下次重新计算时要重新计算的 。|
||[style](/javascript/api/excel/excel.rangeareas#style)|表示此对象中所有范围的 `RangeAreas` 样式。|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[tintAndShade](/javascript/api/excel/excel.rangeborder#tintAndShade)|指定一个使区域边框的颜色变亮或变暗的双精度值，该值介于 -1 (最暗) 和 1 (最亮) 之间，原始颜色为 0。|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[tintAndShade](/javascript/api/excel/excel.rangebordercollection#tintAndShade)|指定使区域边框的颜色变亮或变暗的双精度型值。|
|[RangeCollection](/javascript/api/excel/excel.rangecollection)|[getCount()](/javascript/api/excel/excel.rangecollection#getCount__)|返回 中区域的数量 `RangeCollection` 。|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangecollection#getItemAt_index_)|基于 range 对象在 中的位置返回该对象 `RangeCollection` 。|
||[items](/javascript/api/excel/excel.rangecollection#items)|获取此集合中已加载的子项。|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[pattern](/javascript/api/excel/excel.rangefill#pattern)|范围的图案。|
||[patternColor](/javascript/api/excel/excel.rangefill#patternColor)|HTML 颜色代码，表示区域图案的颜色，格式为 #RRGGBB (例如"FFA500") ，或作为已命名的 HTML 颜色 (例如"orange") 。|
||[patternTintAndShade](/javascript/api/excel/excel.rangefill#patternTintAndShade)|指定使区域填充的图案颜色变亮或变暗的双精度型。|
||[tintAndShade](/javascript/api/excel/excel.rangefill#tintAndShade)|指定使区域填充的颜色变亮或变暗的双精度值。|
|[RangeFont](/javascript/api/excel/excel.rangefont)|[strikethrough](/javascript/api/excel/excel.rangefont#strikethrough)|指定字体的删除线状态。|
||[subscript](/javascript/api/excel/excel.rangefont#subscript)|指定字体的下标状态。|
||[superscript](/javascript/api/excel/excel.rangefont#superscript)|指定字体的上标状态。|
||[tintAndShade](/javascript/api/excel/excel.rangefont#tintAndShade)|指定使区域字体的颜色变亮或变暗的双精度值。|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[autoIndent](/javascript/api/excel/excel.rangeformat#autoIndent)|指定文本对齐方式设置为相等分布时文本是否自动缩进。|
||[indentLevel](/javascript/api/excel/excel.rangeformat#indentLevel)|0 到 250 之间的一个整数，指示缩进水平。|
||[readingOrder](/javascript/api/excel/excel.rangeformat#readingOrder)|区域的读取顺序。|
||[shrinkToFit](/javascript/api/excel/excel.rangeformat#shrinkToFit)|指定文本是否自动缩小以适应可用列宽。|
|[RemoveDuplicatesResult](/javascript/api/excel/excel.removeduplicatesresult)|[removed](/javascript/api/excel/excel.removeduplicatesresult#removed)|由操作删除的重复行数。|
||[uniqueRemaining](/javascript/api/excel/excel.removeduplicatesresult#uniqueRemaining)|所生成的区域中存在的剩余唯一行数。|
|[ReplaceCriteria](/javascript/api/excel/excel.replacecriteria)|[completeMatch](/javascript/api/excel/excel.replacecriteria#completeMatch)|指定匹配是需要完成还是部分匹配。|
||[matchCase](/javascript/api/excel/excel.replacecriteria#matchCase)|指定匹配是否区分大小写。|
|[RowProperties](/javascript/api/excel/excel.rowproperties)|[address](/javascript/api/excel/excel.rowproperties#address)|表示`address`属性。|
||[addressLocal](/javascript/api/excel/excel.rowproperties#addressLocal)|表示`addressLocal`属性。|
||[rowIndex](/javascript/api/excel/excel.rowproperties#rowIndex)|表示`rowIndex`属性。|
|[SearchCriteria](/javascript/api/excel/excel.searchcriteria)|[completeMatch](/javascript/api/excel/excel.searchcriteria#completeMatch)|指定匹配是需要完成还是部分匹配。|
||[matchCase](/javascript/api/excel/excel.searchcriteria#matchCase)|指定匹配是否区分大小写。|
||[searchDirection](/javascript/api/excel/excel.searchcriteria#searchDirection)|指定搜索方向。|
|[SettableCellProperties](/javascript/api/excel/excel.settablecellproperties)|[format](/javascript/api/excel/excel.settablecellproperties#format)|表示`format`属性。|
||[hyperlink](/javascript/api/excel/excel.settablecellproperties#hyperlink)|表示`hyperlink`属性。|
||[style](/javascript/api/excel/excel.settablecellproperties#style)|表示`style`属性。|
|[SettableColumnProperties](/javascript/api/excel/excel.settablecolumnproperties)|[columnHidden](/javascript/api/excel/excel.settablecolumnproperties#columnHidden)|表示`columnHidden`属性。|
||[columnWidth](/javascript/api/excel/excel.settablecolumnproperties#columnWidth)||
||[format：Excel。CellPropertiesFormat & {
            columnWidth？] (/javascript/api/excel/excel.settablecolumnproperties#format) |表示`format`属性。|
|[SettableRowProperties](/javascript/api/excel/excel.settablerowproperties)|[format：Excel。CellPropertiesFormat & {
            rowHeight？] (/javascript/api/excel/excel.settablerowproperties#format) |表示`format`属性。|
||[rowHeight](/javascript/api/excel/excel.settablerowproperties#rowHeight)||
||[rowHidden](/javascript/api/excel/excel.settablerowproperties#rowHidden)|表示`rowHidden`属性。|
|[Shape](/javascript/api/excel/excel.shape)|[altTextDescription](/javascript/api/excel/excel.shape#altTextDescription)|指定对象的可选说明 `Shape` 文本。|
||[altTextTitle](/javascript/api/excel/excel.shape#altTextTitle)|指定对象的可选标题 `Shape` 文本。|
||[delete()](/javascript/api/excel/excel.shape#delete__)|从工作表删除形状。|
||[geometricShapeType](/javascript/api/excel/excel.shape#geometricShapeType)|指定此几何形状的几何形状类型。|
||[getAsImage(format: Excel.PictureFormat)](/javascript/api/excel/excel.shape#getAsImage_format_)|将形状转换为图像并将图像返回为 base64 编码字符串。|
||[height](/javascript/api/excel/excel.shape#height)|指定形状的高度（以点表示）。|
||[incrementLeft(increment: number)](/javascript/api/excel/excel.shape#incrementLeft_increment_)|以指定磅数水平移动形状。|
||[incrementRotation(increment: number)](/javascript/api/excel/excel.shape#incrementRotation_increment_)|将形状围绕 z 轴旋转特定度数。|
||[incrementTop(increment: number)](/javascript/api/excel/excel.shape#incrementTop_increment_)|以指定磅数垂直移动形状。|
||[left](/javascript/api/excel/excel.shape#left)|从形状左侧到工作表左侧的距离（以磅为单位）。|
||[lockAspectRatio](/javascript/api/excel/excel.shape#lockAspectRatio)|指定是否锁定此形状的纵横比。|
||[名称](/javascript/api/excel/excel.shape#name)|指定形状的名称。|
||[connectionSiteCount](/javascript/api/excel/excel.shape#connectionSiteCount)|返回此形状上的连接站点数。|
||[fill](/javascript/api/excel/excel.shape#fill)|返回此形状的填充格式。|
||[geometricShape](/javascript/api/excel/excel.shape#geometricShape)|返回与形状关联的几何形状。|
||[组](/javascript/api/excel/excel.shape#group)|返回与形状关联的形状组。|
||[id](/javascript/api/excel/excel.shape#id)|指定形状标识符。|
||[image](/javascript/api/excel/excel.shape#image)|返回与形状关联的图像。|
||[level](/javascript/api/excel/excel.shape#level)|指定指定形状的级别。|
||[line](/javascript/api/excel/excel.shape#line)|返回与形状关联的线条。|
||[lineFormat](/javascript/api/excel/excel.shape#lineFormat)|返回此形状的线条格式。|
||[onActivated](/javascript/api/excel/excel.shape#onActivated)|当激活形状时发生此事件。|
||[onDeactivated](/javascript/api/excel/excel.shape#onDeactivated)|当停用形状时发生此事件。|
||[parentGroup](/javascript/api/excel/excel.shape#parentGroup)|指定此形状的父组。|
||[textFrame](/javascript/api/excel/excel.shape#textFrame)|返回此形状的文本框对象。|
||[type](/javascript/api/excel/excel.shape#type)|返回此形状的类型。|
||[zOrderPosition](/javascript/api/excel/excel.shape#zOrderPosition)|返回指定形状在 z 顺序中的位置，0 表示顺序堆栈的底部。|
||[rotation](/javascript/api/excel/excel.shape#rotation)|指定形状的旋转度数。|
||[scaleHeight(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#scaleHeight_scaleFactor__scaleType__scaleFrom_)|按指定因子缩放形状的高度。|
||[scaleWidth(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#scaleWidth_scaleFactor__scaleType__scaleFrom_)|按指定因子缩放形状的宽度。|
||[setZOrder(position: Excel.ShapeZOrder)](/javascript/api/excel/excel.shape#setZOrder_position_)|将指定形状沿集合的 z 顺序向上或向下移动，将其移动到其他形状的前面或后面。|
||[top](/javascript/api/excel/excel.shape#top)|从形状上边缘到工作表上边缘之间的距离（以磅为单位）。|
||[visible](/javascript/api/excel/excel.shape#visible)|指定形状是否可见。|
||[width](/javascript/api/excel/excel.shape#width)|指定形状的宽度（以点表示）。|
|[ShapeActivatedEventArgs](/javascript/api/excel/excel.shapeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapeactivatedeventargs#shapeId)|获取已激活形状的 ID。|
||[type](/javascript/api/excel/excel.shapeactivatedeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.shapeactivatedeventargs#worksheetId)|获取激活形状的工作表的 ID。|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addGeometricShape(geometricShapeType: Excel.GeometricShapeType)](/javascript/api/excel/excel.shapecollection#addGeometricShape_geometricShapeType_)|将几何形状添加到工作表。|
||[addGroup(values: Array<string \| Shape>)](/javascript/api/excel/excel.shapecollection#addGroup_values_)|在此集合的工作表中对形状的子集进行分组。|
||[addImage(base64ImageString: string)](/javascript/api/excel/excel.shapecollection#addImage_base64ImageString_)|从 base64 编码的字符串创建图像并将其添加到工作表。|
||[addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType)](/javascript/api/excel/excel.shapecollection#addLine_startLeft__startTop__endLeft__endTop__connectorType_)|将线条添加到工作表。|
||[addTextBox(text?: string)](/javascript/api/excel/excel.shapecollection#addTextBox_text_)|使用提供的文本作为内容，将文本框添加到工作表。|
||[getCount()](/javascript/api/excel/excel.shapecollection#getCount__)|返回工作表中的形状数。|
||[getItem(key: string)](/javascript/api/excel/excel.shapecollection#getItem_key_)|使用形状的名称或 ID 获取形状。|
||[getItemAt(index: number)](/javascript/api/excel/excel.shapecollection#getItemAt_index_)|使用其在集合中的位置获取形状。|
||[items](/javascript/api/excel/excel.shapecollection#items)|获取此集合中已加载的子项。|
|[ShapeDeactivatedEventArgs](/javascript/api/excel/excel.shapedeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapedeactivatedeventargs#shapeId)|获取已停用形状的形状的 ID。|
||[type](/javascript/api/excel/excel.shapedeactivatedeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.shapedeactivatedeventargs#worksheetId)|获取在其中停用形状的工作表的 ID。|
|[ShapeFill](/javascript/api/excel/excel.shapefill)|[clear()](/javascript/api/excel/excel.shapefill#clear__)|清除此形状的填充格式。|
||[foregroundColor](/javascript/api/excel/excel.shapefill#foregroundColor)|以 HTML 颜色格式表示形状填充前景色，格式为 #RRGGBB (例如"FFA500") 或作为已命名的 HTML 颜色 (例如"orange") |
||[type](/javascript/api/excel/excel.shapefill#type)|返回形状的填充类型。|
||[setSolidColor(color: string)](/javascript/api/excel/excel.shapefill#setSolidColor_color_)|将形状的填充格式设置为统一颜色。|
||[transparency](/javascript/api/excel/excel.shapefill#transparency)|将填充的透明度百分比指定为从 0.0 到 1.0 (不透明) 1.0 (透明) 。|
|[ShapeFont](/javascript/api/excel/excel.shapefont)|[bold](/javascript/api/excel/excel.shapefont#bold)|表示字体的加粗状态。|
||[color](/javascript/api/excel/excel.shapefont#color)|文本颜色格式的 HTML 颜色代码表示 (例如，"#FF0000"表示红色) 。|
||[italic](/javascript/api/excel/excel.shapefont#italic)|表示字体的斜体状态。|
||[名称](/javascript/api/excel/excel.shapefont#name)|表示字体名称 (例如"Calibri") 。|
||[size](/javascript/api/excel/excel.shapefont#size)|表示字号（以 (，例如 11) ）。|
||[underline](/javascript/api/excel/excel.shapefont#underline)|应用于字体的下划线类型。|
|[ShapeGroup](/javascript/api/excel/excel.shapegroup)|[id](/javascript/api/excel/excel.shapegroup#id)|指定形状标识符。|
||[shape](/javascript/api/excel/excel.shapegroup#shape)|返回 `Shape` 与组关联的对象。|
||[shapes](/javascript/api/excel/excel.shapegroup#shapes)|返回对象 `Shape` 的集合。|
||[ungroup()](/javascript/api/excel/excel.shapegroup#ungroup__)|取消分组指定形状组中的任何已分组形状。|
|[ShapeLineFormat](/javascript/api/excel/excel.shapelineformat)|[color](/javascript/api/excel/excel.shapelineformat#color)|表示 HTML 颜色格式的线条颜色，格式为 #RRGGBB (例如"FFA500") 或作为已命名的 HTML 颜色 (例如"orange") 。|
||[dashStyle](/javascript/api/excel/excel.shapelineformat#dashStyle)|表示形状的线条样式。|
||[style](/javascript/api/excel/excel.shapelineformat#style)|表示形状的线条样式。|
||[transparency](/javascript/api/excel/excel.shapelineformat#transparency)|将指定线条的透明度表示为从 0.0（不透明）到 1.0（清晰）的值。|
||[visible](/javascript/api/excel/excel.shapelineformat#visible)|指定形状元素的线条格式是否可见。|
||[weight](/javascript/api/excel/excel.shapelineformat#weight)|表示线条的粗细（以磅为单位）。|
|[SortField](/javascript/api/excel/excel.sortfield)|[subField](/javascript/api/excel/excel.sortfield#subField)|指定作为要排序的丰富值的目标属性名称的子字段。|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[getCount()](/javascript/api/excel/excel.stylecollection#getCount__)|获取集合中的样式数量。|
||[getItemAt(index: number)](/javascript/api/excel/excel.stylecollection#getItemAt_index_)|根据其在集合中的位置获取样式。|
|[Table](/javascript/api/excel/excel.table)|[autoFilter](/javascript/api/excel/excel.table#autoFilter)|表示 `AutoFilter` table 的对象。|
|[TableAddedEventArgs](/javascript/api/excel/excel.tableaddedeventargs)|[source](/javascript/api/excel/excel.tableaddedeventargs#source)|获取事件源。|
||[tableId](/javascript/api/excel/excel.tableaddedeventargs#tableId)|获取已添加的表的 ID。|
||[type](/javascript/api/excel/excel.tableaddedeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.tableaddedeventargs#worksheetId)|获取添加表格的工作表的 ID。|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[details](/javascript/api/excel/excel.tablechangedeventargs#details)|获取有关更改详细信息的信息。|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onAdded](/javascript/api/excel/excel.tablecollection#onAdded)|在工作簿中添加新表时发生。|
||[onDeleted](/javascript/api/excel/excel.tablecollection#onDeleted)|在工作簿中删除指定的表格时发生。|
|[TableDeletedEventArgs](/javascript/api/excel/excel.tabledeletedeventargs)|[source](/javascript/api/excel/excel.tabledeletedeventargs#source)|获取事件源。|
||[tableId](/javascript/api/excel/excel.tabledeletedeventargs#tableId)|获取已删除的表的 ID。|
||[tableName](/javascript/api/excel/excel.tabledeletedeventargs#tableName)|获取已删除的表的名称。|
||[type](/javascript/api/excel/excel.tabledeletedeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.tabledeletedeventargs#worksheetId)|获取删除表格的工作表的 ID。|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getCount()](/javascript/api/excel/excel.tablescopedcollection#getCount__)|获取集合中的表数量。|
||[getFirst()](/javascript/api/excel/excel.tablescopedcollection#getFirst__)|获取集合中的第一个表格。|
||[getItem(key: string)](/javascript/api/excel/excel.tablescopedcollection#getItem_key_)|按名称或 ID 获取表。|
||[items](/javascript/api/excel/excel.tablescopedcollection#items)|获取此集合中已加载的子项。|
|[TextFrame](/javascript/api/excel/excel.textframe)|[autoSizeSetting](/javascript/api/excel/excel.textframe#autoSizeSetting)|文本框的自动大小调整设置。|
||[bottomMargin](/javascript/api/excel/excel.textframe#bottomMargin)|表示文本框的下边距（以磅为单位）。|
||[deleteText()](/javascript/api/excel/excel.textframe#deleteText__)|删除文本框中的所有文本。|
||[horizontalAlignment](/javascript/api/excel/excel.textframe#horizontalAlignment)|表示文本框的水平对齐方式。|
||[horizontalOverflow](/javascript/api/excel/excel.textframe#horizontalOverflow)|表示文本框的水平溢出行为。|
||[leftMargin](/javascript/api/excel/excel.textframe#leftMargin)|表示文本框的左边距（以磅为单位）。|
||[orientation](/javascript/api/excel/excel.textframe#orientation)|表示文本面向文本框的角度。|
||[readingOrder](/javascript/api/excel/excel.textframe#readingOrder)|表示文本框从左到右或从右到左的读取顺序。|
||[hasText](/javascript/api/excel/excel.textframe#hasText)|指定文本框是否包含文本。|
||[textRange](/javascript/api/excel/excel.textframe#textRange)|表示附加到文本框中形状上的文本，以及用于操作文本的属性和方法。|
||[rightMargin](/javascript/api/excel/excel.textframe#rightMargin)|表示文本框的右边距（以磅为单位）。|
||[topMargin](/javascript/api/excel/excel.textframe#topMargin)|表示文本框的上边距（以磅为单位）。|
||[verticalAlignment](/javascript/api/excel/excel.textframe#verticalAlignment)|表示文本框的垂直对齐方式。|
||[verticalOverflow](/javascript/api/excel/excel.textframe#verticalOverflow)|表示文本框的垂直溢出行为。|
|[TextRange](/javascript/api/excel/excel.textrange)|[getSubstring(start: number, length?: number)](/javascript/api/excel/excel.textrange#getSubstring_start__length_)|返回给定区域内子字符串的 TextRange 对象。|
||[font](/javascript/api/excel/excel.textrange#font)|返回 `ShapeFont` 一个对象，该对象代表文本范围的字体属性。|
||[text](/javascript/api/excel/excel.textrange#text)|表示文本范围的纯文本内容。|
|[Workbook](/javascript/api/excel/excel.workbook)|[chartDataPointTrack](/javascript/api/excel/excel.workbook#chartDataPointTrack)|如果工作簿中的所有图表都跟踪它们所附加的实际数据点，则为 True。|
||[getActiveChart()](/javascript/api/excel/excel.workbook#getActiveChart__)|获取工作簿中的当前活动图表。|
||[getActiveChartOrNullObject()](/javascript/api/excel/excel.workbook#getActiveChartOrNullObject__)|获取工作簿中的当前活动图表。|
||[getIsActiveCollabSession()](/javascript/api/excel/excel.workbook#getIsActiveCollabSession__)|返回 `true` 工作簿是否正由多个用户编辑， (共同创作) 。|
||[getSelectedRanges()](/javascript/api/excel/excel.workbook#getSelectedRanges__)|从工作簿中获取当前选定的一个或多个区域。|
||[isDirty](/javascript/api/excel/excel.workbook#isDirty)|指定自上次保存工作簿以来是否进行了更改。|
||[autoSave](/javascript/api/excel/excel.workbook#autoSave)|指定工作簿是否位于"自动保存"模式。|
||[calculationEngineVersion](/javascript/api/excel/excel.workbook#calculationEngineVersion)|返回有关 Excel 计算引擎的版本号。|
||[onAutoSaveSettingChanged](/javascript/api/excel/excel.workbook#onAutoSaveSettingChanged)|在工作簿上更改"自动保存"设置时发生。|
||[previouslySaved](/javascript/api/excel/excel.workbook#previouslySaved)|指定工作簿是在本地保存还是联机保存。|
||[usePrecisionAsDisplayed](/javascript/api/excel/excel.workbook#usePrecisionAsDisplayed)|如果此工作簿中的计算仅使用显示的数字精度来完成，则为 True。|
|[WorkbookAutoSaveSettingChangedEventArgs](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs)|[type](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs#type)|获取事件的类型。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[enableCalculation](/javascript/api/excel/excel.worksheet#enableCalculation)|确定是否Excel重新计算工作表（ 如有必要）。|
||[findAll(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#findAll_text__criteria_)|根据指定的条件查找给定字符串的所有匹配项，并作为包含一个或多个矩形区域的对象 `RangeAreas` 返回。|
||[findAllOrNullObject(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#findAllOrNullObject_text__criteria_)|根据指定的条件查找给定字符串的所有匹配项，并作为包含一个或多个矩形区域的对象 `RangeAreas` 返回。|
||[getRanges(address?: string)](/javascript/api/excel/excel.worksheet#getRanges_address_)|获取 `RangeAreas` 表示由地址或名称指定的一个或多个矩形区域块的对象。|
||[autoFilter](/javascript/api/excel/excel.worksheet#autoFilter)|表示 `AutoFilter` 工作表的对象。|
||[horizontalPageBreaks](/javascript/api/excel/excel.worksheet#horizontalPageBreaks)|获取工作表的水平分页符集合。|
||[onFormatChanged](/javascript/api/excel/excel.worksheet#onFormatChanged)|在特定工作表上更改格式时发生。|
||[pageLayout](/javascript/api/excel/excel.worksheet#pageLayout)|获取 `PageLayout` 工作表的对象。|
||[shapes](/javascript/api/excel/excel.worksheet#shapes)|返回工作表上的所有 Shape 对象的集合。|
||[verticalPageBreaks](/javascript/api/excel/excel.worksheet#verticalPageBreaks)|获取工作表的垂直分页符集合。|
||[replaceAll(text: string, replacement: string, criteria: Excel.ReplaceCriteria)](/javascript/api/excel/excel.worksheet#replaceAll_text__replacement__criteria_)|根据当前工作表中指定的条件查找并替换给定的字符串。|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[details](/javascript/api/excel/excel.worksheetchangedeventargs#details)|表示有关更改详细信息的信息。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onChanged](/javascript/api/excel/excel.worksheetcollection#onChanged)|在更改工作簿中的任何工作表时发生。|
||[onFormatChanged](/javascript/api/excel/excel.worksheetcollection#onFormatChanged)|当工作簿中任何工作表的格式发生更改时发生。|
||[onSelectionChanged](/javascript/api/excel/excel.worksheetcollection#onSelectionChanged)|在任何工作表上更改选择时发生。|
|[WorksheetFormatChangedEventArgs](/javascript/api/excel/excel.worksheetformatchangedeventargs)|[address](/javascript/api/excel/excel.worksheetformatchangedeventargs#address)|获取区域地址，该地址表示特定工作表上的更改区域。|
||[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#getRange_ctx_)|获取区域，该区域表示特定工作表上的更改区域。|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#getRangeOrNullObject_ctx_)|获取区域，该区域表示特定工作表上的更改区域。|
||[source](/javascript/api/excel/excel.worksheetformatchangedeventargs#source)|获取事件源。|
||[type](/javascript/api/excel/excel.worksheetformatchangedeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.worksheetformatchangedeventargs#worksheetId)|获取其中的数据发生更改的工作表的 ID。|
|[WorksheetSearchCriteria](/javascript/api/excel/excel.worksheetsearchcriteria)|[completeMatch](/javascript/api/excel/excel.worksheetsearchcriteria#completeMatch)|指定匹配是需要完成还是部分匹配。|
||[matchCase](/javascript/api/excel/excel.worksheetsearchcriteria#matchCase)|指定匹配是否区分大小写。|

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 参考文档](/javascript/api/excel?view=excel-js-1.9&preserve-view=true)
- [Excel JavaScript API 要求集](excel-api-requirement-sets.md)
