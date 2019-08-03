---
title: Excel JavaScript API 要求集1。9
description: 有关 ExcelApi 1.9 要求集的详细信息
ms.date: 07/26/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: b28406f9792278e554ff055a59ef4833be915aba
ms.sourcegitcommit: 3f5d7f4794e3d3c8bc3a79fa05c54157613b9376
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/02/2019
ms.locfileid: "36064863"
---
# <a name="whats-new-in-excel-javascript-api-19"></a>Excel JavaScript API 1.9 中的新增功能

超过 500 个新  Excel API 随 1.9 要求集一起推出。 第一个表提供了 API 的简明摘要，而后续表提供了详细列表。

| 功能区域 | 说明 | 相关对象 |
|:--- |:--- |:--- |
| [Shape](../../excel/excel-add-ins-shapes.md) | 插入、定位和格式化图像、几何形状和文本框。 | [ShapeCollection](/javascript/api/excel/excel.shapecollection) [Shape](/javascript/api/excel/excel.shape) [GeometricShape](/javascript/api/excel/excel.geometricshape)  [Image](/javascript/api/excel/excel.image) |
| [自动筛选](../../excel/excel-add-ins-worksheets.md#filter-data) | 为区域添加筛选器。 | [AutoFilter](/javascript/api/excel/excel.autofilter) |
| [Areas](../../excel/excel-add-ins-multiple-ranges.md) | 支持非连续区域。 | [RangeAreas](/javascript/api/excel/excel.rangeareas) |
| [特殊单元格](../../excel/excel-add-ins-multiple-ranges.md#get-special-cells-from-multiple-ranges) | 获取在区域内包含日期、备注或公式的单元格。 | [Range](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-)|
| [查找](../../excel/excel-add-ins-ranges.md#find-a-cell-using-string-matching) | 查找区域或工作表中的值或公式。 | [Range](/javascript/api/excel/excel.range#find-text--criteria-)[Worksheet](/javascript/api/excel/excel.worksheet#findall-text--criteria-) |
| [复制和粘贴](../../excel/excel-add-ins-ranges-advanced.md#copy-and-paste) | 将值、格式和公式从一个区域复制到另一个区域。 | [Range](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-) |
| [Calculation](../../excel/performance.md#suspend-calculation-temporarily) | 更好地控制 Excel 计算引擎。 | [应用程序](/javascript/api/excel/excel.application) |
| 新图表 | 了解我们支持的新图表类型：地图、箱形图、瀑布图、旭日图、排列图 和漏斗图。 | [Chart](/javascript/api/excel/excel.charttype) |
| RangeFormat | 新功能及区域格式。 | [Range](/javascript/api/excel/excel.rangeformat) |

## <a name="api-list"></a>API 列表

下表列出了 Excel JavaScript API 要求集1.9 中的 Api。 若要查看 Excel JavaScript API 要求集1.9 或更早版本支持的所有 Api 的 API 参考文档, 请参阅[要求集1.9 或更早版本中的 Excel api](/javascript/api/excel?view=excel-js-1.9)。

| Class | 域 | 说明 |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[calculationEngineVersion](/javascript/api/excel/excel.application#calculationengineversion)|返回用于上次完整重新计算的 Excel 计算引擎版本。 只读。|
||[calculationState](/javascript/api/excel/excel.application#calculationstate)|返回应用程序的计算状态。 有关详细信息，请参阅 Excel.CalculationState。 只读。|
||[iterativeCalculation](/javascript/api/excel/excel.application#iterativecalculation)|返回“迭代计算”设置。|
||[suspendScreenUpdatingUntilNextSync()](/javascript/api/excel/excel.application#suspendscreenupdatinguntilnextsync--)|在下一次调用“context.sync()”前暂停屏幕更新。|
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[apply(range: Range \| string, columnIndex?: number, criteria?: Excel.FilterCriteria)](/javascript/api/excel/excel.autofilter#apply-range--columnindex--criteria-)|将自动筛选器应用于区域。 如果指定了列索引和筛选条件，则筛选列。|
||[clearCriteria()](/javascript/api/excel/excel.autofilter#clearcriteria--)|清除自动筛选器的筛选条件。|
||[getRange()](/javascript/api/excel/excel.autofilter#getrange--)|返回 Range 对象，该对象表示“自动筛选”应用的区域。|
||[getRangeOrNullObject()](/javascript/api/excel/excel.autofilter#getrangeornullobject--)|返回 Range 对象，该对象表示“自动筛选”应用的区域。|
||[criteria](/javascript/api/excel/excel.autofilter#criteria)|在自动筛选区域中保留所有筛选条件的数组。 只读。|
||[enabled](/javascript/api/excel/excel.autofilter#enabled)|指示是否启用了自动筛选。 只读。|
||[isDataFiltered](/javascript/api/excel/excel.autofilter#isdatafiltered)|指示自动筛选是否具有筛选条件。 只读。|
||[reapply()](/javascript/api/excel/excel.autofilter#reapply--)|应用当前位于区域上的指定 Autofilter 对象。|
||[remove()](/javascript/api/excel/excel.autofilter#remove--)|删除区域的自动筛选。|
|[CellBorder](/javascript/api/excel/excel.cellborder)|[color](/javascript/api/excel/excel.cellborder#color)|表示`color`单个边框的属性。|
||[style](/javascript/api/excel/excel.cellborder#style)|表示`style`单个边框的属性。|
||[tintAndShade](/javascript/api/excel/excel.cellborder#tintandshade)|表示`tintAndShade`单个边框的属性。|
||[weight](/javascript/api/excel/excel.cellborder#weight)|表示`weight`单个边框的属性。|
|[CellBorderCollection](/javascript/api/excel/excel.cellbordercollection)|[bottom](/javascript/api/excel/excel.cellbordercollection#bottom)|表示`format.borders.bottom`属性。|
||[diagonalDown](/javascript/api/excel/excel.cellbordercollection#diagonaldown)|表示`format.borders.diagonalDown`属性。|
||[diagonalUp](/javascript/api/excel/excel.cellbordercollection#diagonalup)|表示`format.borders.diagonalUp`属性。|
||[horizontal](/javascript/api/excel/excel.cellbordercollection#horizontal)|表示`format.borders.horizontal`属性。|
||[left](/javascript/api/excel/excel.cellbordercollection#left)|表示`format.borders.left`属性。|
||[right](/javascript/api/excel/excel.cellbordercollection#right)|表示`format.borders.right`属性。|
||[top](/javascript/api/excel/excel.cellbordercollection#top)|表示`format.borders.top`属性。|
||[vertical](/javascript/api/excel/excel.cellbordercollection#vertical)|表示`format.borders.vertical`属性。|
|[CellProperties](/javascript/api/excel/excel.cellproperties)|[address](/javascript/api/excel/excel.cellproperties#address)|表示`address`属性。|
||[addressLocal](/javascript/api/excel/excel.cellproperties#addresslocal)|表示`addressLocal`属性。|
||[hidden](/javascript/api/excel/excel.cellproperties#hidden)|表示`hidden`属性。|
|[CellPropertiesFill](/javascript/api/excel/excel.cellpropertiesfill)|[color](/javascript/api/excel/excel.cellpropertiesfill#color)|表示`format.fill.color`属性。|
||[pattern](/javascript/api/excel/excel.cellpropertiesfill#pattern)|表示`format.fill.pattern`属性。|
||[patternColor](/javascript/api/excel/excel.cellpropertiesfill#patterncolor)|表示`format.fill.patternColor`属性。|
||[patternTintAndShade](/javascript/api/excel/excel.cellpropertiesfill#patterntintandshade)|表示`format.fill.patternTintAndShade`属性。|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfill#tintandshade)|表示`format.fill.tintAndShade`属性。|
|[CellPropertiesFont](/javascript/api/excel/excel.cellpropertiesfont)|[bold](/javascript/api/excel/excel.cellpropertiesfont#bold)|表示`format.font.bold`属性。|
||[color](/javascript/api/excel/excel.cellpropertiesfont#color)|表示`format.font.color`属性。|
||[italic](/javascript/api/excel/excel.cellpropertiesfont#italic)|表示`format.font.italic`属性。|
||[name](/javascript/api/excel/excel.cellpropertiesfont#name)|表示`format.font.name`属性。|
||[size](/javascript/api/excel/excel.cellpropertiesfont#size)|表示`format.font.size`属性。|
||[strikethrough](/javascript/api/excel/excel.cellpropertiesfont#strikethrough)|表示`format.font.strikethrough`属性。|
||[subscript](/javascript/api/excel/excel.cellpropertiesfont#subscript)|表示`format.font.subscript`属性。|
||[superscript](/javascript/api/excel/excel.cellpropertiesfont#superscript)|表示`format.font.superscript`属性。|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfont#tintandshade)|表示`format.font.tintAndShade`属性。|
||[underline](/javascript/api/excel/excel.cellpropertiesfont#underline)|表示`format.font.underline`属性。|
|[CellPropertiesFormat](/javascript/api/excel/excel.cellpropertiesformat)|[autoIndent](/javascript/api/excel/excel.cellpropertiesformat#autoindent)|表示`autoIndent`属性。|
||[Borders](/javascript/api/excel/excel.cellpropertiesformat#borders)|表示`borders`属性。|
||[fill](/javascript/api/excel/excel.cellpropertiesformat#fill)|表示`fill`属性。|
||[font](/javascript/api/excel/excel.cellpropertiesformat#font)|表示`font`属性。|
||[horizontalAlignment](/javascript/api/excel/excel.cellpropertiesformat#horizontalalignment)|表示`horizontalAlignment`属性。|
||[indentLevel](/javascript/api/excel/excel.cellpropertiesformat#indentlevel)|表示`indentLevel`属性。|
||[protection](/javascript/api/excel/excel.cellpropertiesformat#protection)|表示`protection`属性。|
||[readingOrder](/javascript/api/excel/excel.cellpropertiesformat#readingorder)|表示`readingOrder`属性。|
||[shrinkToFit](/javascript/api/excel/excel.cellpropertiesformat#shrinktofit)|表示`shrinkToFit`属性。|
||[textOrientation](/javascript/api/excel/excel.cellpropertiesformat#textorientation)|表示`textOrientation`属性。|
||[useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#usestandardheight)|表示`useStandardHeight`属性。|
||[useStandardWidth](/javascript/api/excel/excel.cellpropertiesformat#usestandardwidth)|表示`useStandardWidth`属性。|
||[verticalAlignment](/javascript/api/excel/excel.cellpropertiesformat#verticalalignment)|表示`verticalAlignment`属性。|
||[wrapText](/javascript/api/excel/excel.cellpropertiesformat#wraptext)|表示`wrapText`属性。|
|[CellPropertiesProtection](/javascript/api/excel/excel.cellpropertiesprotection)|[formulaHidden](/javascript/api/excel/excel.cellpropertiesprotection#formulahidden)|表示`format.protection.formulaHidden`属性。|
||[locked](/javascript/api/excel/excel.cellpropertiesprotection#locked)|表示`format.protection.locked`属性。|
|[ChangedEventDetail](/javascript/api/excel/excel.changedeventdetail)|[valueAfter](/javascript/api/excel/excel.changedeventdetail#valueafter)|表示更改之后的值。 返回的数据可能是字符串、数字，也可能是布尔值。 包含错误的单元格将返回错误字符串。|
||[valueBefore](/javascript/api/excel/excel.changedeventdetail#valuebefore)|表示更改之前的值。 返回的数据可能是字符串、数字，也可能是布尔值。 包含错误的单元格将返回错误字符串。|
||[valueTypeAfter](/javascript/api/excel/excel.changedeventdetail#valuetypeafter)|表示更改之后的值类型。|
||[valueTypeBefore](/javascript/api/excel/excel.changedeventdetail#valuetypebefore)|表示更改之前的值类型。|
|[Chart](/javascript/api/excel/excel.chart)|[activate()](/javascript/api/excel/excel.chart#activate--)|在 Excel UI 中激活图表。|
||[pivotOptions](/javascript/api/excel/excel.chart#pivotoptions)|封装数据透视图的选项。 只读。|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[colorScheme](/javascript/api/excel/excel.chartareaformat#colorscheme)|返回或设置图表的配色方案。 读/写。|
||[roundedCorners](/javascript/api/excel/excel.chartareaformat#roundedcorners)|指定图表的图表区域是否有圆角。 读/写。|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[linkNumberFormat](/javascript/api/excel/excel.chartaxis#linknumberformat)|表示数字格式是否链接到单元格。 如果为 true，则数字格式会在单元格中更改时在标签中更改。|
|[ChartBinOptions](/javascript/api/excel/excel.chartbinoptions)|[allowOverflow](/javascript/api/excel/excel.chartbinoptions#allowoverflow)|指定是否在直方图或排列图中启用容器溢出。 读/写。|
||[allowUnderflow](/javascript/api/excel/excel.chartbinoptions#allowunderflow)|指定是否在直方图或排列图中启用容器下溢。 读/写。|
||[count](/javascript/api/excel/excel.chartbinoptions#count)|返回或设置直方图或排列图的容器数量。 读/写。|
||[overflowValue](/javascript/api/excel/excel.chartbinoptions#overflowvalue)|返回或设置直方图或排列图的容器溢出值。 读/写。|
||[type](/javascript/api/excel/excel.chartbinoptions#type)|返回或设置直方图或排列图的容器类型。 读/写。|
||[underflowValue](/javascript/api/excel/excel.chartbinoptions#underflowvalue)|返回或设置直方图或排列图的容器下溢值。 读/写。|
||[width](/javascript/api/excel/excel.chartbinoptions#width)|返回或设置直方图或排列图的容器宽度值。 读/写。|
|[ChartBoxwhiskerOptions](/javascript/api/excel/excel.chartboxwhiskeroptions)|[quartileCalculation](/javascript/api/excel/excel.chartboxwhiskeroptions#quartilecalculation)|返回或设置箱形图的四分位点计算类型。 读/写。|
||[showInnerPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#showinnerpoints)|指定在箱形图中是否显示内部点。 读/写。|
||[showMeanLine](/javascript/api/excel/excel.chartboxwhiskeroptions#showmeanline)|指定在箱形图中是否显示中线。 读/写。|
||[showMeanMarker](/javascript/api/excel/excel.chartboxwhiskeroptions#showmeanmarker)|指定在箱形图中是否显示平均值标记。 读/写。|
||[showOutlierPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#showoutlierpoints)|指定在箱形图中是否显示离群值点。 读/写。|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabel#linknumberformat)|布尔值，表示数字格式是否链接到单元格（以便在单元格中更改时标签中的数字格式会发生改变）。|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabels#linknumberformat)|表示数字格式是否链接到单元格。 如果为 true，则数字格式会在单元格中更改时在标签中更改。|
|[ChartErrorBars](/javascript/api/excel/excel.charterrorbars)|[endStyleCap](/javascript/api/excel/excel.charterrorbars#endstylecap)|指定误差线是否具有终止端样式。|
||[include](/javascript/api/excel/excel.charterrorbars#include)|指定包含误差线的哪些部分。|
||[format](/javascript/api/excel/excel.charterrorbars#format)|指定误差线的格式类型。|
||[type](/javascript/api/excel/excel.charterrorbars#type)|误差线标记的区域类型。|
||[visible](/javascript/api/excel/excel.charterrorbars#visible)|指定是否显示误差线。|
|[ChartErrorBarsFormat](/javascript/api/excel/excel.charterrorbarsformat)|[line](/javascript/api/excel/excel.charterrorbarsformat#line)|表示图表线条格式。|
|[ChartMapOptions](/javascript/api/excel/excel.chartmapoptions)|[labelStrategy](/javascript/api/excel/excel.chartmapoptions#labelstrategy)|返回或设置区域地图图表的系列地图标签策略。 读/写。|
||[level](/javascript/api/excel/excel.chartmapoptions#level)|返回或设置区域地图图表的系列映射级别。 读/写。|
||[projectionType](/javascript/api/excel/excel.chartmapoptions#projectiontype)|返回或设置区域地图图表的系列投影类型。 读/写。|
|[ChartPivotOptions](/javascript/api/excel/excel.chartpivotoptions)|[showAxisFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showaxisfieldbuttons)|指定是否在数据透视图上显示轴字段按钮。 ShowAxisFieldButtons 属性对应于“分析”选项卡（在选择数据透视图时可用）的“字段按钮”下拉列表中的“显示坐标轴字段按钮”命令。|
||[showLegendFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showlegendfieldbuttons)|指定是否在数据透视图上显示图例字段按钮。|
||[showReportFilterFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showreportfilterfieldbuttons)|指定是否在数据透视图上显示报表筛选字段按钮。|
||[showValueFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showvaluefieldbuttons)|指定是否在数据透视图上显示值字段按钮。|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[bubbleScale](/javascript/api/excel/excel.chartseries#bubblescale)|这可以是从 0（零）到 300 的整数值，表示默认大小的百分比。 该属性仅适用于气泡图。 读/写。|
||[gradientMaximumColor](/javascript/api/excel/excel.chartseries#gradientmaximumcolor)|返回或设置区域地图图表系列的最大值的颜色。 读/写。|
||[gradientMaximumType](/javascript/api/excel/excel.chartseries#gradientmaximumtype)|返回或设置区域地图图表系列的最大值的类型。 读/写。|
||[gradientMaximumValue](/javascript/api/excel/excel.chartseries#gradientmaximumvalue)|返回或设置区域地图图表系列的最大值。 读/写。|
||[gradientMidpointColor](/javascript/api/excel/excel.chartseries#gradientmidpointcolor)|返回或设置区域地图图表系列的中间值的颜色。 读/写。|
||[gradientMidpointType](/javascript/api/excel/excel.chartseries#gradientmidpointtype)|返回或设置区域地图图表系列的中间值的类型。 读/写。|
||[gradientMidpointValue](/javascript/api/excel/excel.chartseries#gradientmidpointvalue)|返回或设置区域地图图表系列的中间值。 读/写。|
||[gradientMinimumColor](/javascript/api/excel/excel.chartseries#gradientminimumcolor)|返回或设置区域地图图表系列的最小值的颜色。 读/写。|
||[gradientMinimumType](/javascript/api/excel/excel.chartseries#gradientminimumtype)|返回或设置区域地图图表系列的最小值的类型。 读/写。|
||[gradientMinimumValue](/javascript/api/excel/excel.chartseries#gradientminimumvalue)|返回或设置区域地图图表系列的最小值。 读/写。|
||[gradientStyle](/javascript/api/excel/excel.chartseries#gradientstyle)|返回或设置区域地图图表的系列渐变样式。 读/写。|
||[invertColor](/javascript/api/excel/excel.chartseries#invertcolor)|返回或设置系列中负数据点的填充颜色。 读/写。|
||[parentLabelStrategy](/javascript/api/excel/excel.chartseries#parentlabelstrategy)|返回或设置树状图的系列父标签策略区域。 读/写。|
||[binOptions](/javascript/api/excel/excel.chartseries#binoptions)|封装直方图和排列图的容器选项。 只读。|
||[boxwhiskerOptions](/javascript/api/excel/excel.chartseries#boxwhiskeroptions)|封装箱形图的选项。 只读。|
||[mapOptions](/javascript/api/excel/excel.chartseries#mapoptions)|封装区域地图图表的选项。 只读。|
||[xErrorBars](/javascript/api/excel/excel.chartseries#xerrorbars)|表示图表系列的误差线对象。|
||[yErrorBars](/javascript/api/excel/excel.chartseries#yerrorbars)|表示图表系列的误差线对象。|
||[showConnectorLines](/javascript/api/excel/excel.chartseries#showconnectorlines)|指定是否在瀑布图中显示连接线。 读/写。|
||[showLeaderLines](/javascript/api/excel/excel.chartseries#showleaderlines)|指定是否在系列中显示每个数据标签的引导线。 读/写。|
||[splitValue](/javascript/api/excel/excel.chartseries#splitvalue)|返回或设置复合饼图或复合条饼图中分隔两部分的阈值。 读/写。|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[linkNumberFormat](/javascript/api/excel/excel.charttrendlinelabel#linknumberformat)|布尔值，表示数字格式是否链接到单元格（以便在单元格中更改时标签中的数字格式会发生改变）。|
|[ColumnProperties](/javascript/api/excel/excel.columnproperties)|[address](/javascript/api/excel/excel.columnproperties#address)|表示`address`属性。|
||[addressLocal](/javascript/api/excel/excel.columnproperties#addresslocal)|表示`addressLocal`属性。|
||[columnIndex](/javascript/api/excel/excel.columnproperties#columnindex)|表示`columnIndex`属性。|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[getRanges()](/javascript/api/excel/excel.conditionalformat#getranges--)|返回将为其应用条件格式的 RangeAreas，它包含一个或多个矩形区域。 只读。|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[getInvalidCells()](/javascript/api/excel/excel.datavalidation#getinvalidcells--)|返回包含一个或多个矩形区域的 RangeAreas，它具有无效单元格值。 如果所有单元格值都有效，则此函数将引发 ItemNotFound 错误。|
||[getInvalidCellsOrNullObject()](/javascript/api/excel/excel.datavalidation#getinvalidcellsornullobject--)|返回包含一个或多个矩形区域的 RangeAreas，它具有无效单元格值。 如果所有单元格值都有效，则此函数将返回 null。|
|[FilterCriteria](/javascript/api/excel/excel.filtercriteria)|[subField](/javascript/api/excel/excel.filtercriteria#subfield)|筛选器使用该属性对 richvalue 执行丰富的筛选。|
|[GeometricShape](/javascript/api/excel/excel.geometricshape)|[id](/javascript/api/excel/excel.geometricshape#id)|返回形状标识符。 只读。|
||[shape](/javascript/api/excel/excel.geometricshape#shape)|返回几何形状的形状对象。 只读。|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getCount()](/javascript/api/excel/excel.groupshapecollection#getcount--)|返回形状组中的形状数量。 只读。|
||[getItem(key: string)](/javascript/api/excel/excel.groupshapecollection#getitem-key-)|按名称或 ID 获取形状。|
||[getItemAt(index: number)](/javascript/api/excel/excel.groupshapecollection#getitemat-index-)|根据其在集合中的位置获取形状。|
||[items](/javascript/api/excel/excel.groupshapecollection#items)|获取此集合中已加载的子项。|
|[HeaderFooter](/javascript/api/excel/excel.headerfooter)|[centerFooter](/javascript/api/excel/excel.headerfooter#centerfooter)|获取或设置工作表的中心页脚。|
||[centerHeader](/javascript/api/excel/excel.headerfooter#centerheader)|获取或设置工作表的中心页眉。|
||[LeftFooter](/javascript/api/excel/excel.headerfooter#leftfooter)|获取或设置工作表的左侧页脚。|
||[leftHeader](/javascript/api/excel/excel.headerfooter#leftheader)|获取或设置工作表的左侧页眉。|
||[rightFooter](/javascript/api/excel/excel.headerfooter#rightfooter)|获取或设置工作表的右侧页脚。|
||[rightHeader](/javascript/api/excel/excel.headerfooter#rightheader)|获取或设置工作表的右侧页眉。|
|[HeaderFooterGroup](/javascript/api/excel/excel.headerfootergroup)|[defaultForAllPages](/javascript/api/excel/excel.headerfootergroup#defaultforallpages)|常规页眉/页脚，除非指定偶数页/奇数页或首页，否则适用于所有页面，|
||[evenPages](/javascript/api/excel/excel.headerfootergroup#evenpages)|用于偶数页的页眉/页脚，需要为奇数页指定奇数页页眉/页脚。|
||[firstPage](/javascript/api/excel/excel.headerfootergroup#firstpage)|首页的页眉/页脚，为所有其他页使用常规或偶数页/奇数页页眉/页脚。|
||[oddPages](/javascript/api/excel/excel.headerfootergroup#oddpages)|用于奇数页的页眉/页脚，需要为偶数页指定偶数页页眉/页脚。|
||[state](/javascript/api/excel/excel.headerfootergroup#state)|获取或设置所设置的页眉/页脚的状态。 有关详细信息，请参阅 Excel.HeaderFooterState。|
||[useSheetMargins](/javascript/api/excel/excel.headerfootergroup#usesheetmargins)|获取或设置一个标记，指示页眉/页脚是否与工作表的页面布局选项中设置的页边距对齐。|
||[useSheetScale](/javascript/api/excel/excel.headerfootergroup#usesheetscale)|获取或设置一个标记，指示是否应按照工作表的页面布局选项中设置的页面缩放百分比来缩放页眉/页脚。|
|[Image](/javascript/api/excel/excel.image)|[format](/javascript/api/excel/excel.image#format)|返回图像的格式。 只读。|
||[id](/javascript/api/excel/excel.image#id)|表示图像对象的形状标识符。 只读。|
||[shape](/javascript/api/excel/excel.image#shape)|返回与图像关联的形状对象。 只读。|
|[IterativeCalculation](/javascript/api/excel/excel.iterativecalculation)|[enabled](/javascript/api/excel/excel.iterativecalculation#enabled)|如果 Excel 使用迭代来处理循环引用，则为 True。|
||[maxChange](/javascript/api/excel/excel.iterativecalculation#maxchange)|返回或设置 Excel 处理循环引用时迭代之间的最大变化值。|
||[maxIteration](/javascript/api/excel/excel.iterativecalculation#maxiteration)|返回或设置 Excel 处理循环引用的最大迭代次数。|
|[Line](/javascript/api/excel/excel.line)|[beginArrowheadLength](/javascript/api/excel/excel.line#beginarrowheadlength)|表示指定线条始端的箭头长度。|
||[beginArrowheadStyle](/javascript/api/excel/excel.line#beginarrowheadstyle)|表示指定线条始端的箭头样式。|
||[beginArrowheadWidth](/javascript/api/excel/excel.line#beginarrowheadwidth)|表示指定线条始端的箭头宽度。|
||[connectBeginShape(shape: Excel.Shape, connectionSite: number)](/javascript/api/excel/excel.line#connectbeginshape-shape--connectionsite-)|将指定连接线的始端附加到指定形状。|
||[connectEndShape(shape: Excel.Shape, connectionSite: number)](/javascript/api/excel/excel.line#connectendshape-shape--connectionsite-)|将指定连接线的末端附加到指定形状。|
||[connectorType](/javascript/api/excel/excel.line#connectortype)|表示线条的连接器类型。|
||[disconnectBeginShape()](/javascript/api/excel/excel.line#disconnectbeginshape--)|使指定连接线的始端与形状脱离。|
||[disconnectEndShape()](/javascript/api/excel/excel.line#disconnectendshape--)|使指定连接线的末端与形状脱离。|
||[endArrowheadLength](/javascript/api/excel/excel.line#endarrowheadlength)|表示指定线条末端的箭头长度。|
||[endArrowheadStyle](/javascript/api/excel/excel.line#endarrowheadstyle)|表示指定线条末端的箭头样式。|
||[endArrowheadWidth](/javascript/api/excel/excel.line#endarrowheadwidth)|表示指定线条末端的箭头宽度。|
||[beginConnectedShape](/javascript/api/excel/excel.line#beginconnectedshape)|表示指定线条始端所附加到的形状。 只读。|
||[beginConnectedSite](/javascript/api/excel/excel.line#beginconnectedsite)|表示连接线始端所连接的连接站点。 只读。 当线条的始端没有附加到任何形状时，返回 null。|
||[endConnectedShape](/javascript/api/excel/excel.line#endconnectedshape)|表示指定线条末端所附加到的形状。 只读。|
||[endConnectedSite](/javascript/api/excel/excel.line#endconnectedsite)|表示连接线末端所连接的连接站点。 只读。 当线条的末端没有附加到任何形状时，返回 null。|
||[id](/javascript/api/excel/excel.line#id)|表示形状标识符。 只读。|
||[isBeginConnected](/javascript/api/excel/excel.line#isbeginconnected)|指定指定线条的始端是否连接到形状。 只读。|
||[isEndConnected](/javascript/api/excel/excel.line#isendconnected)|指定指定线条的末端是否连接到形状。 只读。|
||[shape](/javascript/api/excel/excel.line#shape)|返回与线条关联的形状对象。 只读。|
|[PageBreak](/javascript/api/excel/excel.pagebreak)|[delete()](/javascript/api/excel/excel.pagebreak#delete--)|删除分页符对象。|
||[getCellAfterBreak()](/javascript/api/excel/excel.pagebreak#getcellafterbreak--)|获取分页符后的第一个单元格。|
||[columnIndex](/javascript/api/excel/excel.pagebreak#columnindex)|表示分页符的列索引|
||[rowIndex](/javascript/api/excel/excel.pagebreak#rowindex)|表示分页符的行索引|
|[PageBreakCollection](/javascript/api/excel/excel.pagebreakcollection)|[add(pageBreakRange: Range \| string)](/javascript/api/excel/excel.pagebreakcollection#add-pagebreakrange-)|在指定区域的左上角单元格之前添加分页符。|
||[getCount()](/javascript/api/excel/excel.pagebreakcollection#getcount--)|获取集合中的分页符数量。|
||[getItem(index: number)](/javascript/api/excel/excel.pagebreakcollection#getitem-index-)|通过索引获取分页符对象。|
||[items](/javascript/api/excel/excel.pagebreakcollection#items)|获取此集合中已加载的子项。|
||[removePageBreaks()](/javascript/api/excel/excel.pagebreakcollection#removepagebreaks--)|重置集合中的所有手动分页符。|
|[PageLayout](/javascript/api/excel/excel.pagelayout)|[blackAndWhite](/javascript/api/excel/excel.pagelayout#blackandwhite)|获取或设置工作表的黑白打印选项。|
||[bottomMargin](/javascript/api/excel/excel.pagelayout#bottommargin)|获取或设置要用于打印的工作表的底部页边距（以磅为单位）。|
||[centerHorizontally](/javascript/api/excel/excel.pagelayout#centerhorizontally)|获取或设置工作表的中心水平标记。 此标记确定在打印时是否水平居中工作表。|
||[centerVertically](/javascript/api/excel/excel.pagelayout#centervertically)|获取或设置工作表的中心垂直标记。 此标记确定在打印时是否垂直居中工作表。|
||[draftMode](/javascript/api/excel/excel.pagelayout#draftmode)|获取或设置工作表的草稿模式选项。 如果为 True，则将打印没有图形的工作表。|
||[firstPageNumber](/javascript/api/excel/excel.pagelayout#firstpagenumber)|获取或设置要打印的工作表的首页页码。 Null 值表示“自动”页码编号。|
||[footerMargin](/javascript/api/excel/excel.pagelayout#footermargin)|获取或设置在打印时使用的工作表的页脚边距（以磅为单位）。|
||[getPrintArea()](/javascript/api/excel/excel.pagelayout#getprintarea--)|获取包含一个或多个矩形区域的 RangeAreas 对象，它表示工作表的打印区域。 如果没有打印区域，则将引发 ItemNotFound 错误。|
||[getPrintAreaOrNullObject()](/javascript/api/excel/excel.pagelayout#getprintareaornullobject--)|获取包含一个或多个矩形区域的 RangeAreas 对象，它表示工作表的打印区域。 如果没有打印区域，则将返回 null 对象。|
||[getPrintTitleColumns()](/javascript/api/excel/excel.pagelayout#getprinttitlecolumns--)|获取表示标题列的 Range 对象。|
||[getPrintTitleColumnsOrNullObject()](/javascript/api/excel/excel.pagelayout#getprinttitlecolumnsornullobject--)|获取表示标题列的 Range 对象。 如果未设置，则将返回 null 对象。|
||[getPrintTitleRows()](/javascript/api/excel/excel.pagelayout#getprinttitlerows--)|获取表示标题行的 Range 对象。|
||[getPrintTitleRowsOrNullObject()](/javascript/api/excel/excel.pagelayout#getprinttitlerowsornullobject--)|获取表示标题行的 Range 对象。 如果未设置，则将返回 null 对象。|
||[headerMargin](/javascript/api/excel/excel.pagelayout#headermargin)|获取或设置在打印时使用的工作表的页眉边距（以磅为单位）。|
||[leftMargin](/javascript/api/excel/excel.pagelayout#leftmargin)|获取或设置在打印时使用的工作表的左边距（以磅为单位）。|
||[orientation](/javascript/api/excel/excel.pagelayout#orientation)|获取或设置工作表的页面方向。|
||[paperSize](/javascript/api/excel/excel.pagelayout#papersize)|获取或设置工作表的页面纸张大小。|
||[printComments](/javascript/api/excel/excel.pagelayout#printcomments)|获取或设置在打印时是否应该显示工作表的批注。|
||[printErrors](/javascript/api/excel/excel.pagelayout#printerrors)|获取或设置工作表的打印错误选项。|
||[printGridlines](/javascript/api/excel/excel.pagelayout#printgridlines)|获取或设置工作表的打印网格线标记。 此标记确定是否打印网格线。|
||[printHeadings](/javascript/api/excel/excel.pagelayout#printheadings)|获取或设置工作表的打印标题标记。 此标记确定是否打印标题。|
||[printOrder](/javascript/api/excel/excel.pagelayout#printorder)|获取或设置工作表的页面打印顺序选项。 它指定用于处理打印页码的顺序。|
||[headersFooters](/javascript/api/excel/excel.pagelayout#headersfooters)|工作表的页眉和页脚配置。|
||[rightMargin](/javascript/api/excel/excel.pagelayout#rightmargin)|获取或设置在打印时使用的工作表的右边距（以磅为单位）。|
||[setPrintArea(printArea: Range \| RangeAreas \| string)](/javascript/api/excel/excel.pagelayout#setprintarea-printarea-)|设置工作表的打印区域。|
||[setPrintMargins(unit: Excel.PrintMarginUnit, marginOptions: Excel.PageLayoutMarginOptions)](/javascript/api/excel/excel.pagelayout#setprintmargins-unit--marginoptions-)|设置带单位的工作表的页边距。|
||[setPrintTitleColumns(printTitleColumns: Range \| string)](/javascript/api/excel/excel.pagelayout#setprinttitlecolumns-printtitlecolumns-)|设置列，这些列包含要在打印的工作表的每页左侧重复的单元格。|
||[setPrintTitleRows(printTitleRows: Range \| string)](/javascript/api/excel/excel.pagelayout#setprinttitlerows-printtitlerows-)|设置行，这些行包含要在打印的工作表的每页顶部重复的单元格。|
||[topMargin](/javascript/api/excel/excel.pagelayout#topmargin)|获取或设置在打印时使用的工作表的上边距（以磅为单位）。|
||[zoom](/javascript/api/excel/excel.pagelayout#zoom)|获取或设置工作表的打印缩放选项。|
|[PageLayoutMarginOptions](/javascript/api/excel/excel.pagelayoutmarginoptions)|[bottom](/javascript/api/excel/excel.pagelayoutmarginoptions#bottom)|表示要在打印时使用的页面布局下边距（使用指定的单位）。|
||[footer](/javascript/api/excel/excel.pagelayoutmarginoptions#footer)|表示要在打印时使用的页面布局页脚边距（使用指定的单位）。|
||[header](/javascript/api/excel/excel.pagelayoutmarginoptions#header)|表示要在打印时使用的页面布局页眉边距（使用指定的单位）。|
||[left](/javascript/api/excel/excel.pagelayoutmarginoptions#left)|表示要在打印时使用的页面布局左边距（使用指定的单位）。|
||[right](/javascript/api/excel/excel.pagelayoutmarginoptions#right)|表示要在打印时使用的页面布局右边距（使用指定的单位）。|
||[top](/javascript/api/excel/excel.pagelayoutmarginoptions#top)|表示要在打印时使用的页面布局上边距（使用指定的单位）。|
|[PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions)|[horizontalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#horizontalfittopages)|水平放置的页数。 如果使用百分比缩放，则此值可以为 null。|
||[scale](/javascript/api/excel/excel.pagelayoutzoomoptions#scale)|打印页面缩放值可以介于 10 至 400 之间。 如果已指定适应页面高度或宽度，则此值可以为 null。|
||[verticalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#verticalfittopages)|垂直放置的页数。 如果使用百分比缩放，则此值可以为 null。|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[sortByValues(sortBy: Excel.SortBy, valuesHierarchy: Excel.DataPivotHierarchy, pivotItemScope?: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotfield#sortbyvalues-sortby--valueshierarchy--pivotitemscope-)|按给定范围中的指定值对 PivotField 进行排序。 该范围定义将使用哪些特定值进行排序|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[autoFormat](/javascript/api/excel/excel.pivotlayout#autoformat)|指定是否在刷新或移动字段时自动进行格式化|
||[getDataHierarchy(cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#getdatahierarchy-cell-)|获取 DataHierarchy，它用于计算数据透视表中指定区域内的值。|
||[getPivotItems(axis: Excel.PivotAxis, cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#getpivotitems-axis--cell-)|从构成数据透视表中指定区域内的值的轴获取 PivotItems。|
||[preserveFormatting](/javascript/api/excel/excel.pivotlayout#preserveformatting)|指定是否在通过透视、排序或更改页面字段项等操作来刷新或重新计算报表时保留格式。|
||[setAutoSortOnCell(cell: Range \| string, sortBy: Excel.SortBy)](/javascript/api/excel/excel.pivotlayout#setautosortoncell-cell--sortby-)|将数据透视表设置为使用指定的单元格设置自动排序，以自动选择排序的所有条件和上下文。 这与从 UI 应用自动排序的行为相同。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[enableDataValueEditing](/javascript/api/excel/excel.pivottable#enabledatavalueediting)|指定数据透视表是否允许用户编辑数据体中的值。|
||[useCustomSortLists](/javascript/api/excel/excel.pivottable#usecustomsortlists)|指定排序时，数据透视表是否使用自定义列表。|
|[Range](/javascript/api/excel/excel.range)|[autoFill(destinationRange: Range \| string, autoFillType?: Excel.AutoFillType)](/javascript/api/excel/excel.range#autofill-destinationrange--autofilltype-)|填充区域从当前区域到目标区域。|
||[convertDataTypeToText()](/javascript/api/excel/excel.range#convertdatatypetotext--)|将具有数据类型的区域单元格转换为文本。|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.range#converttolinkeddatatype-serviceid--languageculture-)|将区域单元格转换为工作表中的链接数据类型。|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-)|将单元格数据或格式从源区域或 RangeAreas 复制到当前区域。|
||[find(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#find-text--criteria-)|根据指定的条件查找给定的字符串。|
||[findOrNullObject(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#findornullobject-text--criteria-)|根据指定的条件查找给定的字符串。|
||[flashFill()](/javascript/api/excel/excel.range#flashfill--)|对当前区域进行快速填充。快速填充在感知到模式时可自动填充数据，因此该区域必须是单列区域且周围有数据以便查找模式。|
||[getCellProperties(cellPropertiesLoadOptions: CellPropertiesLoadOptions)](/javascript/api/excel/excel.range#getcellproperties-cellpropertiesloadoptions-)|返回一个 2D 数组，其中封装了每个单元格的字体、填充、边框、对齐方式和其他属性数据。|
||[getColumnProperties(columnPropertiesLoadOptions: ColumnPropertiesLoadOptions)](/javascript/api/excel/excel.range#getcolumnproperties-columnpropertiesloadoptions-)|返回一个一维数组，其中封装了每个列的字体、填充、边框、对齐方式和其他属性数据。  对于给定列中每个单元格不一致的属性，将返回 null。|
||[getRowProperties(rowPropertiesLoadOptions: RowPropertiesLoadOptions)](/javascript/api/excel/excel.range#getrowproperties-rowpropertiesloadoptions-)|返回一个一维数组，其中封装了每个行的字体、填充、边框、对齐方式和其他属性数据。  对于给定行中每个单元格不一致的属性，将返回 null。|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-)|获取包含一个或多个矩形区域的 RangeAreas 对象，它表示匹配指定类型和值的所有单元格。|
||[getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#getspecialcellsornullobject-celltype--cellvaluetype-)|获取包含一个或多个区域的 RangeAreas 对象，它表示匹配指定类型和值的所有单元格。|
||[getTables(fullyContained?: boolean)](/javascript/api/excel/excel.range#gettables-fullycontained-)|获取与区域重叠的限定范围的表格集合。|
||[linkedDataTypeState](/javascript/api/excel/excel.range#linkeddatatypestate)|表示每个单元格的数据类型状态。 只读。|
||[removeDuplicates(columns: number[], includesHeader: boolean)](/javascript/api/excel/excel.range#removeduplicates-columns--includesheader-)|从列指定的区域中删除重复值。|
||[replaceAll(text: string, replacement: string, criteria: Excel.ReplaceCriteria)](/javascript/api/excel/excel.range#replaceall-text--replacement--criteria-)|根据当前区域内指定的条件查找并替换给定的字符串。|
||[setCellProperties(cellPropertiesData: SettableCellProperties[][])](/javascript/api/excel/excel.range#setcellproperties-cellpropertiesdata-)|根据单元格属性的 2D 数组更新区域，它封装了字体、填充、边框、对齐方式等内容。|
||[setColumnProperties(columnPropertiesData: SettableColumnProperties[])](/javascript/api/excel/excel.range#setcolumnproperties-columnpropertiesdata-)|根据列属性的一维数组更新区域，它封装了字体、填充、边框、对齐方式等内容。|
||[setDirty()](/javascript/api/excel/excel.range#setdirty--)|设置下一次重新计算发生时要重新计算的区域。|
||[setRowProperties(rowPropertiesData: SettableRowProperties[])](/javascript/api/excel/excel.range#setrowproperties-rowpropertiesdata-)|根据行属性的一维数组更新区域，它封装了字体、填充、边框、对齐方式等内容。|
|[RangeAreas](/javascript/api/excel/excel.rangeareas)|[calculate()](/javascript/api/excel/excel.rangeareas#calculate--)|计算 RangeAreas 中的所有单元格。|
||[clear(applyTo?: Excel.ClearApplyTo)](/javascript/api/excel/excel.rangeareas#clear-applyto-)|清除包含此 RangeAreas 对象的每个区域的值、格式、填充、边框等。|
||[convertDataTypeToText()](/javascript/api/excel/excel.rangeareas#convertdatatypetotext--)|将 RangeAreas 中具有数据类型的所有单元格转换为文本。|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.rangeareas#converttolinkeddatatype-serviceid--languageculture-)|将 RangeAreas 中的所有单元格转换为链接数据类型。|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.rangeareas#copyfrom-sourcerange--copytype--skipblanks--transpose-)|将单元格数据或格式从源区域或 RangeAreas 复制到当前 RangeAreas。|
||[getEntireColumn()](/javascript/api/excel/excel.rangeareas#getentirecolumn--)|返回表示 RangeAreas 的整个列的 RangeAreas 对象（例如，如果当前 RangeAreas 表示单元格“B4:E11, H2”，它将返回表示列“B:E, H:H”的 RangeAreas）。|
||[getEntireRow()](/javascript/api/excel/excel.rangeareas#getentirerow--)|获取表示 RangeAreas 的整个行的 RangeAreas 对象（例如，如果当前 RangeAreas 表示单元格“B4:E11”，它将返回表示行“4:11”的 RangeAreas）。|
||[getIntersection(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#getintersection-anotherrange-)|返回 RangeAreas 对象，它表示给定区域或 RangeAreas 的交集。 如果未找到任何交集，则将引发 ItemNotFound 错误。|
||[getIntersectionOrNullObject(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#getintersectionornullobject-anotherrange-)|返回 RangeAreas 对象，它表示给定区域或 RangeAreas 的交集。 如果未找到任何交集，将返回 null 对象。|
||[getOffsetRangeAreas(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.rangeareas#getoffsetrangeareas-rowoffset--columnoffset-)|返回 RangeAreas 对象，它按特定的行和列偏移量进行移动。 返回的 RangeAreas 的维度将与原始对象匹配。 如果生成的 RangeAreas 强行超出工作表网格的边界，则将引发错误。|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#getspecialcells-celltype--cellvaluetype-)|返回一个 RangeAreas 对象，它表示匹配指定类型和值的所有单元格。 如果未找到符合条件的特殊单元格，则会引发错误。|
||[getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#getspecialcellsornullobject-celltype--cellvaluetype-)|返回一个 RangeAreas 对象，它表示匹配指定类型和值的所有单元格。 如果未找到符合条件的特殊单元格，则返回 null 对象。|
||[getTables(fullyContained?: boolean)](/javascript/api/excel/excel.rangeareas#gettables-fullycontained-)|返回与此 RangeAreas 对象中的任何区域重叠的限定范围的表格集合。|
||[getUsedRangeAreas(valuesOnly?: boolean)](/javascript/api/excel/excel.rangeareas#getusedrangeareas-valuesonly-)|返回使用的 RangeAreas，它包含 RangeAreas 对象中的各个矩形区域的所有已用区域。|
||[getUsedRangeAreasOrNullObject(valuesOnly?: boolean)](/javascript/api/excel/excel.rangeareas#getusedrangeareasornullobject-valuesonly-)|返回使用的 RangeAreas，它包含 RangeAreas 对象中的各个矩形区域的所有已用区域。|
||[address](/javascript/api/excel/excel.rangeareas#address)|返回 A1 样式中的 RageAreas 引用。 地址值将包含单元格的每个矩形块的工作表名称（例如“Sheet1!A1:B4, Sheet1!D1:D4”）。 只读。|
||[addressLocal](/javascript/api/excel/excel.rangeareas#addresslocal)|返回用户区域设置中的 RageAreas 引用。 只读。|
||[areaCount](/javascript/api/excel/excel.rangeareas#areacount)|返回包含此 RangeAreas 对象的矩形区域的数量。|
||[areas](/javascript/api/excel/excel.rangeareas#areas)|返回包含此 RangeAreas 对象的矩形区域的集合。|
||[cellCount](/javascript/api/excel/excel.rangeareas#cellcount)|返回 RangeAreas 对象中的单元格数量，即总计各个矩形区域的单元格计数。 如果单元格计数超过 2^31-1 (2,147,483,647)，则返回 -1。 只读。|
||[conditionalFormats](/javascript/api/excel/excel.rangeareas#conditionalformats)|返回与此 RangeAreas 对象中的任何单元格相交的 ConditionalFormats 集合。 只读。|
||[dataValidation](/javascript/api/excel/excel.rangeareas#datavalidation)|返回 RangeAreas 中的所有区域的 dataValidation 对象。|
||[format](/javascript/api/excel/excel.rangeareas#format)|返回一个 rangeFormat 对象，其中封装了 RangeAreas 对象中的所有区域的字体、填充、边框、对齐方式和其他属性。 只读。|
||[isEntireColumn](/javascript/api/excel/excel.rangeareas#isentirecolumn)|指示此 RangeAreas 对象上的所有区域是否表示整列（例如“A:C, Q:Z”）。 只读。|
||[isEntireRow](/javascript/api/excel/excel.rangeareas#isentirerow)|指示此 RangeAreas 对象上的所有区域是否表示整行（例如“1:3, 5:7”）。 只读。|
||[worksheet](/javascript/api/excel/excel.rangeareas#worksheet)|返回当前 RangeAreas 的工作表。 只读。|
||[setDirty()](/javascript/api/excel/excel.rangeareas#setdirty--)|设置要在下一次重新计算时重新进行计算的 RangeAreas。|
||[style](/javascript/api/excel/excel.rangeareas#style)|表示此 RangeAreas 对象中的所有区域的样式。|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[tintAndShade](/javascript/api/excel/excel.rangeborder#tintandshade)|返回或设置一个使区域边框的颜色变亮或变暗的双精度数值，该值介于 -1（最暗）与 1（最亮）之间，初始颜色为 0。|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[tintAndShade](/javascript/api/excel/excel.rangebordercollection#tintandshade)|返回或设置一个使区域边框的颜色变亮或变暗的双精度数值，该值介于 -1（最暗）与 1（最亮）之间，初始颜色为 0。|
|[RangeCollection](/javascript/api/excel/excel.rangecollection)|[getCount()](/javascript/api/excel/excel.rangecollection#getcount--)|返回 RangeCollection 中的区域数量。|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangecollection#getitemat-index-)|根据其在 RangeCollection 中的位置返回 Range 对象。|
||[items](/javascript/api/excel/excel.rangecollection#items)|获取此集合中已加载的子项。|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[pattern](/javascript/api/excel/excel.rangefill#pattern)|获取或设置区域的图案。 有关详细信息，请参阅 Excel.FillPattern。 不支持 LinearGradient 和 RectangularGradient。|
||[patternColor](/javascript/api/excel/excel.rangefill#patterncolor)|设置 HTML 颜色代码，它表示窗体 #RRGGBB（例如“FFA500”）的区域图案颜色或作为已命名的 HTML 颜色（例如“orange”）。|
||[patternTintAndShade](/javascript/api/excel/excel.rangefill#patterntintandshade)|返回或设置一个使区域填充的图案颜色变亮或变暗的双精度数值，该值介于 -1（最暗）与 1（最亮）之间，初始颜色为 0。|
||[tintAndShade](/javascript/api/excel/excel.rangefill#tintandshade)|返回或设置一个使区域填充的颜色变亮或变暗的双精度数值，该值介于 -1（最暗）与 1（最亮）之间，初始颜色为 0。|
|[RangeFont](/javascript/api/excel/excel.rangefont)|[strikethrough](/javascript/api/excel/excel.rangefont#strikethrough)|表示字体的删除线状态。 Null 值表示整个区域没有统一的删除线设置。|
||[subscript](/javascript/api/excel/excel.rangefont#subscript)|表示字体的下标状态。|
||[superscript](/javascript/api/excel/excel.rangefont#superscript)|表示字体的上标状态。|
||[tintAndShade](/javascript/api/excel/excel.rangefont#tintandshade)|返回或设置一个使区域字体的颜色变亮或变暗的双精度数值，该值介于 -1（最暗）与 1（最亮）之间，初始颜色为 0。|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[autoIndent](/javascript/api/excel/excel.rangeformat#autoindent)|指示将文本对齐方式设为相等分布时文本是否会自动缩进。|
||[indentLevel](/javascript/api/excel/excel.rangeformat#indentlevel)|0 到 250 之间的一个整数，指示缩进水平。|
||[readingOrder](/javascript/api/excel/excel.rangeformat#readingorder)|区域的读取顺序。|
||[shrinkToFit](/javascript/api/excel/excel.rangeformat#shrinktofit)|指示文本是否自动缩小以适合可用列宽。|
|[RemoveDuplicatesResult](/javascript/api/excel/excel.removeduplicatesresult)|[removed](/javascript/api/excel/excel.removeduplicatesresult#removed)|由操作删除的重复行数。|
||[uniqueRemaining](/javascript/api/excel/excel.removeduplicatesresult#uniqueremaining)|所生成的区域中存在的剩余唯一行数。|
|[ReplaceCriteria](/javascript/api/excel/excel.replacecriteria)|[completeMatch](/javascript/api/excel/excel.replacecriteria#completematch)|指定匹配必须是完整匹配还是部分匹配。 默认值为 false（部分）。|
||[matchCase](/javascript/api/excel/excel.replacecriteria#matchcase)|指定匹配是否区分大小写。 默认值为 false（不区分大小写）。|
|[RowProperties](/javascript/api/excel/excel.rowproperties)|[address](/javascript/api/excel/excel.rowproperties#address)|表示`address`属性。|
||[addressLocal](/javascript/api/excel/excel.rowproperties#addresslocal)|表示`addressLocal`属性。|
||[rowIndex](/javascript/api/excel/excel.rowproperties#rowindex)|表示`rowIndex`属性。|
|[SearchCriteria](/javascript/api/excel/excel.searchcriteria)|[completeMatch](/javascript/api/excel/excel.searchcriteria#completematch)|指定匹配必须是完整匹配还是部分匹配。 完全匹配的单元格的全部内容。 默认值为 false（部分）。|
||[matchCase](/javascript/api/excel/excel.searchcriteria#matchcase)|指定匹配是否区分大小写。 默认值为 false（不区分大小写）。|
||[searchDirection](/javascript/api/excel/excel.searchcriteria#searchdirection)|指定搜索方向。 默认值为向前。 请参阅 Excel.SearchDirection。|
|[SettableCellProperties](/javascript/api/excel/excel.settablecellproperties)|[format](/javascript/api/excel/excel.settablecellproperties#format)|表示`format`属性。|
||[hyperlink](/javascript/api/excel/excel.settablecellproperties#hyperlink)|表示`hyperlink`属性。|
||[style](/javascript/api/excel/excel.settablecellproperties#style)|表示`style`属性。|
|[SettableColumnProperties](/javascript/api/excel/excel.settablecolumnproperties)|[columnHidden](/javascript/api/excel/excel.settablecolumnproperties#columnhidden)|表示`columnHidden`属性。|
||[columnWidth](/javascript/api/excel/excel.settablecolumnproperties#columnwidth)||
||[格式: CellPropertiesFormat & {
            columnWidth？](/javascript/api/excel/excel.settablecolumnproperties # format)|表示`format`属性。|
|[SettableRowProperties](/javascript/api/excel/excel.settablerowproperties)|[格式: CellPropertiesFormat & {
            rowHeight？](/javascript/api/excel/excel.settablerowproperties # format)|表示`format`属性。|
||[rowHeight](/javascript/api/excel/excel.settablerowproperties#rowheight)||
||[rowHidden](/javascript/api/excel/excel.settablerowproperties#rowhidden)|表示`rowHidden`属性。|
|[Shape](/javascript/api/excel/excel.shape)|[altTextDescription](/javascript/api/excel/excel.shape#alttextdescription)|返回或设置形状对象的可选说明文本。|
||[altTextTitle](/javascript/api/excel/excel.shape#alttexttitle)|返回或设置形状对象的可选标题文本。|
||[delete()](/javascript/api/excel/excel.shape#delete--)|从工作表删除形状。|
||[geometricShapeType](/javascript/api/excel/excel.shape#geometricshapetype)|表示此几何形状的几何形状类型。 有关详细信息，请参阅 Excel.GeometricShapeType。 如果形状类型不是“GeometricShape”，返回 NULL。|
||[getAsImage(format: Excel.PictureFormat)](/javascript/api/excel/excel.shape#getasimage-format-)|将形状转换为图像并将图像返回为 base64 编码字符串。 DPI 为 96。 仅支持格式 `Excel.PictureFormat.BMP`、`Excel.PictureFormat.PNG`、`Excel.PictureFormat.JPEG` 和 `Excel.PictureFormat.GIF`。|
||[height](/javascript/api/excel/excel.shape#height)|表示形状的高度（以磅为单位）。|
||[incrementLeft(increment: number)](/javascript/api/excel/excel.shape#incrementleft-increment-)|以指定磅数水平移动形状。|
||[incrementRotation(increment: number)](/javascript/api/excel/excel.shape#incrementrotation-increment-)|将形状围绕 z 轴旋转特定度数。|
||[incrementTop(increment: number)](/javascript/api/excel/excel.shape#incrementtop-increment-)|以指定磅数垂直移动形状。|
||[left](/javascript/api/excel/excel.shape#left)|从形状左侧到工作表左侧的距离（以磅为单位）。|
||[lockAspectRatio](/javascript/api/excel/excel.shape#lockaspectratio)|指定此形状的纵横比是否锁定。|
||[名称](/javascript/api/excel/excel.shape#name)|表示形状的名称。|
||[connectionSiteCount](/javascript/api/excel/excel.shape#connectionsitecount)|返回此形状上的连接站点数。 只读。|
||[fill](/javascript/api/excel/excel.shape#fill)|返回此形状的填充格式。 只读。|
||[geometricShape](/javascript/api/excel/excel.shape#geometricshape)|返回与形状关联的几何形状。 如果形状类型不是“GeometricShape”，则会引发错误。|
||[组](/javascript/api/excel/excel.shape#group)|返回与形状关联的形状组。 如果形状类型不是“GroupShape”，则会引发错误。|
||[id](/javascript/api/excel/excel.shape#id)|表示形状标识符。 只读。|
||[image](/javascript/api/excel/excel.shape#image)|返回与形状关联的图像。 如果形状类型不是“Image”，则会引发错误。|
||[level](/javascript/api/excel/excel.shape#level)|表示指定形状的级别。 例如，级别 0 表示形状不是任何组的一部分，级别 1 表示形状是顶级组的一部分，级别 2 表示形状是顶级子组的一部分。|
||[line](/javascript/api/excel/excel.shape#line)|返回与形状关联的线条。 如果形状类型不是“Line”，则会引发错误。|
||[lineFormat](/javascript/api/excel/excel.shape#lineformat)|返回此形状的线条格式。 只读。|
||[onActivated](/javascript/api/excel/excel.shape#onactivated)|当激活形状时发生此事件。|
||[onDeactivated](/javascript/api/excel/excel.shape#ondeactivated)|当停用形状时发生此事件。|
||[parentGroup](/javascript/api/excel/excel.shape#parentgroup)|表示此形状的父组。|
||[textFrame](/javascript/api/excel/excel.shape#textframe)|返回此形状的文本框对象。 只读。|
||[type](/javascript/api/excel/excel.shape#type)|返回此形状的类型。 有关详细信息，请参阅 Excel.ShapeType。 只读。|
||[zOrderPosition](/javascript/api/excel/excel.shape#zorderposition)|返回指定形状在 z 顺序中的位置，0 表示顺序堆栈的底部。 只读。|
||[rotation](/javascript/api/excel/excel.shape#rotation)|表示形状的旋转度数。|
||[scaleHeight(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#scaleheight-scalefactor--scaletype--scalefrom-)|按指定因子缩放形状的高度。 对于图像，你可以说明是相对于原始尺寸还是当前尺寸缩放形状。 对于除图片以外的其他形状来说，缩放总是相对于其当前高度而言。|
||[scaleWidth(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#scalewidth-scalefactor--scaletype--scalefrom-)|按指定因子缩放形状的宽度。 对于图像，你可以说明是相对于原始尺寸还是当前尺寸缩放形状。 对于除图片以外的其他形状来说，缩放总是相对于其当前宽度而言。|
||[setZOrder(position: Excel.ShapeZOrder)](/javascript/api/excel/excel.shape#setzorder-position-)|将指定形状沿集合的 z 顺序向上或向下移动，将其移动到其他形状的前面或后面。|
||[top](/javascript/api/excel/excel.shape#top)|从形状上边缘到工作表上边缘之间的距离（以磅为单位）。|
||[visible](/javascript/api/excel/excel.shape#visible)|表示此形状的可视性。|
||[width](/javascript/api/excel/excel.shape#width)|表示形状的宽度（以磅为单位）。|
|[ShapeActivatedEventArgs](/javascript/api/excel/excel.shapeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapeactivatedeventargs#shapeid)|获取已激活的形状的 ID。|
||[type](/javascript/api/excel/excel.shapeactivatedeventargs#type)|获取事件的类型。 有关详细信息，请参阅 Excel.EventType。|
||[worksheetId](/javascript/api/excel/excel.shapeactivatedeventargs#worksheetid)|获取其中的形状已启用的工作表的 ID。|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addGeometricShape(geometricShapeType: Excel.GeometricShapeType)](/javascript/api/excel/excel.shapecollection#addgeometricshape-geometricshapetype-)|将几何形状添加到工作表。 返回一个 Shape 对象，该对象代表新图形。|
||[addGroup(values: Array<string \| Shape>)](/javascript/api/excel/excel.shapecollection#addgroup-values-)|在此集合的工作表中对形状的子集进行分组。 返回表示新形状组的 Shape 对象。|
||[addImage(base64ImageString: string)](/javascript/api/excel/excel.shapecollection#addimage-base64imagestring-)|从 base64 编码的字符串创建图像并将其添加到工作表。 返回表示新图片的 Shape 对象。|
||[addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType)](/javascript/api/excel/excel.shapecollection#addline-startleft--starttop--endleft--endtop--connectortype-)|将线条添加到工作表。 返回表示新线条的 Shape 对象。|
||[addTextBox(text?: string)](/javascript/api/excel/excel.shapecollection#addtextbox-text-)|使用提供的文本作为内容，将文本框添加到工作表。 返回表示新文本框的 Shape 对象。|
||[getCount()](/javascript/api/excel/excel.shapecollection#getcount--)|返回工作表中的形状数。 只读。|
||[getItem(key: string)](/javascript/api/excel/excel.shapecollection#getitem-key-)|按名称或 ID 获取形状。|
||[getItemAt(index: number)](/javascript/api/excel/excel.shapecollection#getitemat-index-)|使用其在集合中的位置获取形状。|
||[items](/javascript/api/excel/excel.shapecollection#items)|获取此集合中已加载的子项。|
|[ShapeDeactivatedEventArgs](/javascript/api/excel/excel.shapedeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapedeactivatedeventargs#shapeid)|获取已停用的形状的 ID。|
||[type](/javascript/api/excel/excel.shapedeactivatedeventargs#type)|获取事件的类型。 有关详细信息，请参阅 Excel.EventType。|
||[worksheetId](/javascript/api/excel/excel.shapedeactivatedeventargs#worksheetid)|获取其中的形状已停用的工作表的 ID。|
|[ShapeFill](/javascript/api/excel/excel.shapefill)|[clear()](/javascript/api/excel/excel.shapefill#clear--)|清除此形状的填充格式。|
||[foregroundColor](/javascript/api/excel/excel.shapefill#foregroundcolor)|表示窗体 #RRGGBB（例如“FFA500”）的形状填充前景色（采用 HTML 颜色格式）或作为已命名的 HTML 颜色（例如“orange”）|
||[type](/javascript/api/excel/excel.shapefill#type)|返回形状的填充类型。 只读。 有关详细信息，请参阅 Excel.ShapeFillType。|
||[setSolidColor(color: string)](/javascript/api/excel/excel.shapefill#setsolidcolor-color-)|将形状的填充格式设置为统一颜色。 这样可将填充类型更改为“Solid”。|
||[transparency](/javascript/api/excel/excel.shapefill#transparency)|返回或设置填充的透明度百分比，取值范围为 0.0（不透明）到 1.0（清晰）之间。 如果形状类型不支持透明度或形状填充透明度不一致（例如使用渐变填充类型），则返回 null。|
|[ShapeFont](/javascript/api/excel/excel.shapefont)|[bold](/javascript/api/excel/excel.shapefont#bold)|表示字体的加粗状态。 如果 TextRange 包含粗体和非粗体文本片段，则返回 null。|
||[color](/javascript/api/excel/excel.shapefont#color)|文本颜色的 HTML 颜色代码表示（例如，#FF0000 表示红色）。 如果 TextRange 包含具有不同颜色的文本片段，则返回 null。|
||[italic](/javascript/api/excel/excel.shapefont#italic)|表示字体的斜体状态。 如果 TextRange 包含斜体和非斜体文本片段，则返回 null。|
||[name](/javascript/api/excel/excel.shapefont#name)|表示字体名称（例如“Calibri”）。 如果文本是复杂脚本或东亚语言，则这是相应的字体名称，否则是拉丁字体名称。|
||[size](/javascript/api/excel/excel.shapefont#size)|表示以磅为单位的字体大小（例如 11）。 如果 TextRange 包含具有不同字体大小的文本片段，则返回 null。|
||[underline](/javascript/api/excel/excel.shapefont#underline)|应用于字体的下划线类型。 如果 TextRange 包含具有不同下划线样式的文本片段，则返回 null。 有关详细信息，请参阅 Excel.ShapeFontUnderlineStyle。|
|[ShapeGroup](/javascript/api/excel/excel.shapegroup)|[id](/javascript/api/excel/excel.shapegroup#id)|表示形状标识符。 只读。|
||[shape](/javascript/api/excel/excel.shapegroup#shape)|返回与组关联的 Shape 对象。 只读。|
||[shapes](/javascript/api/excel/excel.shapegroup#shapes)|返回 Shape 对象的集合。 只读。|
||[ungroup()](/javascript/api/excel/excel.shapegroup#ungroup--)|取消分组指定形状组中的任何已分组形状。|
|[ShapeLineFormat](/javascript/api/excel/excel.shapelineformat)|[color](/javascript/api/excel/excel.shapelineformat#color)|表示窗体 #RRGGBB（例如“FFA500”）的线条颜色（采用 HTML 颜色格式）或作为已命名的 HTML 颜色（例如“orange”）|
||[dashStyle](/javascript/api/excel/excel.shapelineformat#dashstyle)|表示形状的线条样式。 当线条不可见或短划线样式不一致时，返回 null。 有关详细信息，请参阅 Excel.ShapeLineStyle。|
||[style](/javascript/api/excel/excel.shapelineformat#style)|表示形状的线条样式。 当线条不可见或样式不一致时，返回 null。 有关详细信息，请参阅 Excel.ShapeLineStyle。|
||[transparency](/javascript/api/excel/excel.shapelineformat#transparency)|将指定线条的透明度表示为从 0.0（不透明）到 1.0（清晰）的值。 当形状的透明度不一致时，返回 null。|
||[visible](/javascript/api/excel/excel.shapelineformat#visible)|表示形状元素的线条格式是否可见。 当形状的可见性不一致时，返回 null。|
||[weight](/javascript/api/excel/excel.shapelineformat#weight)|表示线条的粗细（以磅为单位）。 当线条不可见或线条粗细不一致时，返回 null。|
|[SortField](/javascript/api/excel/excel.sortfield)|[subField](/javascript/api/excel/excel.sortfield#subfield)|表示子字段，它是要排序的复合值的目标属性名称。|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[getCount()](/javascript/api/excel/excel.stylecollection#getcount--)|获取集合中的样式数量。|
||[getItemAt(index: number)](/javascript/api/excel/excel.stylecollection#getitemat-index-)|根据其在集合中的位置获取样式。|
|[Table](/javascript/api/excel/excel.table)|[autoFilter](/javascript/api/excel/excel.table#autofilter)|表示表格的 AutoFilter 对象。 只读。|
|[TableAddedEventArgs](/javascript/api/excel/excel.tableaddedeventargs)|[source](/javascript/api/excel/excel.tableaddedeventargs#source)|获取事件源。 有关详细信息，请参阅 Excel.EventSource。|
||[tableId](/javascript/api/excel/excel.tableaddedeventargs#tableid)|获取已添加的表格的 ID。|
||[type](/javascript/api/excel/excel.tableaddedeventargs#type)|获取事件的类型。 有关详细信息，请参阅 Excel.EventType。|
||[worksheetId](/javascript/api/excel/excel.tableaddedeventargs#worksheetid)|获取已在其中添加表格的工作表的 ID。|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[details](/javascript/api/excel/excel.tablechangedeventargs#details)|表示更改详情的信息|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onAdded](/javascript/api/excel/excel.tablecollection#onadded)|在工作簿中添加新表格时发生。|
||[onDeleted](/javascript/api/excel/excel.tablecollection#ondeleted)|在工作簿中删除指定的表格时发生。|
|[TableDeletedEventArgs](/javascript/api/excel/excel.tabledeletedeventargs)|[source](/javascript/api/excel/excel.tabledeletedeventargs#source)|指定时间源。 有关详细信息，请参阅 Excel.EventSource。|
||[tableId](/javascript/api/excel/excel.tabledeletedeventargs#tableid)|指定已删除的表格的 ID。|
||[tableName](/javascript/api/excel/excel.tabledeletedeventargs#tablename)|指定已删除的表格的名称。|
||[type](/javascript/api/excel/excel.tabledeletedeventargs#type)|指定事件类型。 有关详细信息，请参阅 Excel.EventType。|
||[worksheetId](/javascript/api/excel/excel.tabledeletedeventargs#worksheetid)|指定已在其内删除表格的工作表的 ID。|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getCount()](/javascript/api/excel/excel.tablescopedcollection#getcount--)|获取集合中的表数量。|
||[getFirst()](/javascript/api/excel/excel.tablescopedcollection#getfirst--)|获取集合中的第一个表格。 集合中的表格按照从上到下、从左到右的顺序排列，因此左上表格是集合中的第一个表格。|
||[getItem(key: string)](/javascript/api/excel/excel.tablescopedcollection#getitem-key-)|按名称或 ID 获取表。|
||[items](/javascript/api/excel/excel.tablescopedcollection#items)|获取此集合中已加载的子项。|
|[TextFrame](/javascript/api/excel/excel.textframe)|[autoSizeSetting](/javascript/api/excel/excel.textframe#autosizesetting)|获取或设置文本框的自动调整大小设置。 可以将文本框设置为自动调整文本大小以适应文本框，或自动调整文本框大小以适应文本，或者不使用自动调整大小设置。|
||[bottomMargin](/javascript/api/excel/excel.textframe#bottommargin)|表示文本框的下边距（以磅为单位）。|
||[deleteText()](/javascript/api/excel/excel.textframe#deletetext--)|删除文本框中的所有文本。|
||[horizontalAlignment](/javascript/api/excel/excel.textframe#horizontalalignment)|表示文本框的水平对齐方式。 有关详细信息，请参阅 Excel.ShapeTextHorizontalAlignment。|
||[horizontalOverflow](/javascript/api/excel/excel.textframe#horizontaloverflow)|表示文本框的水平溢出行为。 有关详细信息，请参阅 Excel.ShapeTextHorizontalOverflow。|
||[leftMargin](/javascript/api/excel/excel.textframe#leftmargin)|表示文本框的左边距（以磅为单位）。|
||[orientation](/javascript/api/excel/excel.textframe#orientation)|表示文本框的文本方向。 有关详细信息，请参阅 Excel.ShapeTextOrientation。|
||[readingOrder](/javascript/api/excel/excel.textframe#readingorder)|表示文本框从左到右或从右到左的读取顺序。 有关详细信息，请参阅 Excel.ShapeTextReadingOrder。|
||[hasText](/javascript/api/excel/excel.textframe#hastext)|指定文本框是否包含文本。|
||[textRange](/javascript/api/excel/excel.textframe#textrange)|表示附加到文本框中形状上的文本，以及用于操作文本的属性和方法。 有关详细信息，请参阅 Excel.TextRange。|
||[rightMargin](/javascript/api/excel/excel.textframe#rightmargin)|表示文本框的右边距（以磅为单位）。|
||[topMargin](/javascript/api/excel/excel.textframe#topmargin)|表示文本框的上边距（以磅为单位）。|
||[verticalAlignment](/javascript/api/excel/excel.textframe#verticalalignment)|表示文本框的垂直对齐方式。 有关详细信息，请参阅 Excel.ShapeTextVerticalAlignment。|
||[verticalOverflow](/javascript/api/excel/excel.textframe#verticaloverflow)|表示文本框的垂直溢出行为。 有关详细信息，请参阅 Excel.ShapeTextVerticalOverflow。|
|[TextRange](/javascript/api/excel/excel.textrange)|[getSubstring(start: number, length?: number)](/javascript/api/excel/excel.textrange#getsubstring-start--length-)|返回给定区域内子字符串的 TextRange 对象。|
||[font](/javascript/api/excel/excel.textrange#font)|返回一个 ShapeFont 对象，该对象表示文本范围的字体属性。 只读。|
||[text](/javascript/api/excel/excel.textrange#text)|表示文本范围的纯文本内容。|
|[Workbook](/javascript/api/excel/excel.workbook)|[chartDataPointTrack](/javascript/api/excel/excel.workbook#chartdatapointtrack)|如果工作簿中的所有图表都跟踪它们所附加的实际数据点，则为 True。|
||[getActiveChart()](/javascript/api/excel/excel.workbook#getactivechart--)|获取工作簿中的当前活动图表。 如果没有活动图表，则在调用此语句时将引发异常|
||[getActiveChartOrNullObject()](/javascript/api/excel/excel.workbook#getactivechartornullobject--)|获取工作簿中的当前活动图表。 如果没有活动图表，则返回 null 对象|
||[getIsActiveCollabSession()](/javascript/api/excel/excel.workbook#getisactivecollabsession--)|如果多个用户正在编辑工作簿（共同创作），则为 True。|
||[getSelectedRanges()](/javascript/api/excel/excel.workbook#getselectedranges--)|从工作簿中获取当前选定的一个或多个区域。 与 getSelectedRange() 不同，此方法返回表示所有选定区域的 RangeAreas 对象。|
||[isDirty](/javascript/api/excel/excel.workbook#isdirty)|指定自上次保存以来是否对指定的工作簿进行任何更改。|
||[autoSave](/javascript/api/excel/excel.workbook#autosave)|指定工作簿是否处于自动保存模式。 只读。|
||[calculationEngineVersion](/javascript/api/excel/excel.workbook#calculationengineversion)|返回有关 Excel 计算引擎的版本号。 只读。|
||[onAutoSaveSettingChanged](/javascript/api/excel/excel.workbook#onautosavesettingchanged)|在工作簿上更改“自动保存”设置时发生。|
||[previouslySaved](/javascript/api/excel/excel.workbook#previouslysaved)|指定工作簿是否已在本地或在线保存。 只读。|
||[usePrecisionAsDisplayed](/javascript/api/excel/excel.workbook#useprecisionasdisplayed)|如果此工作簿中的计算仅使用显示的数字精度来完成，则为 True。|
|[WorkbookAutoSaveSettingChangedEventArgs](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs)|[type](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs#type)|表示事件的类型。 有关详细信息，请参阅 Excel.EventType。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[enableCalculation](/javascript/api/excel/excel.worksheet#enablecalculation)|获取或设置工作表的 enableCalculation 属性。|
||[findAll(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#findall-text--criteria-)|根据指定的条件查找给定字符串的所有匹配项，并将它们作为包含一个或多个矩形区域的 RangeAreas 对象返回。|
||[findAllOrNullObject(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#findallornullobject-text--criteria-)|根据指定的条件查找给定字符串的所有匹配项，并将它们作为包含一个或多个矩形区域的 RangeAreas 对象返回。|
||[getRanges(address?: string)](/javascript/api/excel/excel.worksheet#getranges-address-)|获取按地址或名称指定的 RangeAreas 对象，它表示一个或多个矩形区域块。|
||[autoFilter](/javascript/api/excel/excel.worksheet#autofilter)|表示工作表的 AutoFilter 对象。 只读。|
||[horizontalPageBreaks](/javascript/api/excel/excel.worksheet#horizontalpagebreaks)|获取工作表的水平分页符集合。 此集合仅包含手动分页符。|
||[onFormatChanged](/javascript/api/excel/excel.worksheet#onformatchanged)|在特定工作表上更改格式时发生。|
||[pageLayout](/javascript/api/excel/excel.worksheet#pagelayout)|获取工作表的 PageLayout 对象。|
||[shapes](/javascript/api/excel/excel.worksheet#shapes)|返回工作表上的所有 Shape 对象的集合。 只读。|
||[verticalPageBreaks](/javascript/api/excel/excel.worksheet#verticalpagebreaks)|获取工作表的垂直分页符集合。 此集合仅包含手动分页符。|
||[replaceAll(text: string, replacement: string, criteria: Excel.ReplaceCriteria)](/javascript/api/excel/excel.worksheet#replaceall-text--replacement--criteria-)|根据当前工作表中指定的条件查找并替换给定的字符串。|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[details](/javascript/api/excel/excel.worksheetchangedeventargs#details)|表示更改详情的信息|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onChanged](/javascript/api/excel/excel.worksheetcollection#onchanged)|在更改工作簿中的任何工作表时发生。|
||[onFormatChanged](/javascript/api/excel/excel.worksheetcollection#onformatchanged)|在更改工作簿中的任何工作表的格式时发生。|
||[onSelectionChanged](/javascript/api/excel/excel.worksheetcollection#onselectionchanged)|在任何工作表上更改选择时发生。|
|[WorksheetFormatChangedEventArgs](/javascript/api/excel/excel.worksheetformatchangedeventargs)|[address](/javascript/api/excel/excel.worksheetformatchangedeventargs#address)|获取区域地址，该地址表示特定工作表上的更改区域。|
||[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#getrange-ctx-)|获取区域，该区域表示特定工作表上的更改区域。|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#getrangeornullobject-ctx-)|获取区域，该区域表示特定工作表上的更改区域。 它可能会返回 null 对象。|
||[source](/javascript/api/excel/excel.worksheetformatchangedeventargs#source)|获取事件源。 有关详细信息，请参阅 Excel.EventSource。|
||[type](/javascript/api/excel/excel.worksheetformatchangedeventargs#type)|获取事件的类型。 有关详细信息，请参阅 Excel.EventType。|
||[worksheetId](/javascript/api/excel/excel.worksheetformatchangedeventargs#worksheetid)|获取其中的数据发生更改的工作表的 ID。|
|[WorksheetSearchCriteria](/javascript/api/excel/excel.worksheetsearchcriteria)|[completeMatch](/javascript/api/excel/excel.worksheetsearchcriteria#completematch)|指定匹配必须是完整匹配还是部分匹配。 完全匹配的单元格的全部内容。 默认值为 false（部分）。|
||[matchCase](/javascript/api/excel/excel.worksheetsearchcriteria#matchcase)|指定匹配是否区分大小写。 默认值为 false（不区分大小写）。|

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 参考文档](/javascript/api/excel?view=excel-js-1.9)
- [Excel JavaScript API 要求集](./excel-api-requirement-sets.md)
