---
title: Excel JavaScript API 要求集1。9
description: 有关 ExcelApi 1.9 要求集的详细信息。
ms.date: 11/09/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: e3878954bca943e1895a44ea9482f1c67cba9211
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996506"
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

下表列出了 Excel JavaScript API 要求集1.9 中的 Api。 若要查看 Excel JavaScript API 要求集1.9 或更早版本支持的所有 Api 的 API 参考文档，请参阅 [要求集1.9 或更早版本中的 Excel api](/javascript/api/excel?view=excel-js-1.9&preserve-view=true)。

| Class | 域 | 说明 |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[calculationEngineVersion](/javascript/api/excel/excel.application#calculationengineversion)|返回用于上次完整重新计算的 Excel 计算引擎版本。|
||[calculationState](/javascript/api/excel/excel.application#calculationstate)|返回应用程序的计算状态。|
||[iterativeCalculation](/javascript/api/excel/excel.application#iterativecalculation)|返回“迭代计算”设置。|
||[suspendScreenUpdatingUntilNextSync()](/javascript/api/excel/excel.application#suspendscreenupdatinguntilnextsync--)|在调用下一步之前挂起屏幕更新 `context.sync()` 。|
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[apply(range: Range \| string, columnIndex?: number, criteria?: Excel.FilterCriteria)](/javascript/api/excel/excel.autofilter#apply-range--columnindex--criteria-)|将自动筛选器应用于区域。|
||[clearCriteria()](/javascript/api/excel/excel.autofilter#clearcriteria--)|清除自动筛选器的筛选条件。|
||[getRange()](/javascript/api/excel/excel.autofilter#getrange--)|返回 Range 对象，该对象表示“自动筛选”应用的区域。|
||[getRangeOrNullObject()](/javascript/api/excel/excel.autofilter#getrangeornullobject--)|返回 Range 对象，该对象表示“自动筛选”应用的区域。|
||[criteria](/javascript/api/excel/excel.autofilter#criteria)|在自动筛选区域中保留所有筛选条件的数组。|
||[enabled](/javascript/api/excel/excel.autofilter#enabled)|指定是否启用自动筛选。|
||[isDataFiltered](/javascript/api/excel/excel.autofilter#isdatafiltered)|指定自动筛选是否具有筛选条件。|
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
|[ChangedEventDetail](/javascript/api/excel/excel.changedeventdetail)|[valueAfter](/javascript/api/excel/excel.changedeventdetail#valueafter)|表示更改之后的值。|
||[valueBefore](/javascript/api/excel/excel.changedeventdetail#valuebefore)|表示更改之前的值。|
||[valueTypeAfter](/javascript/api/excel/excel.changedeventdetail#valuetypeafter)|表示更改之后的值类型。|
||[valueTypeBefore](/javascript/api/excel/excel.changedeventdetail#valuetypebefore)|表示更改之前的值类型。|
|[Chart](/javascript/api/excel/excel.chart)|[activate()](/javascript/api/excel/excel.chart#activate--)|在 Excel UI 中激活图表。|
||[pivotOptions](/javascript/api/excel/excel.chart#pivotoptions)|封装数据透视图的选项。|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[colorScheme](/javascript/api/excel/excel.chartareaformat#colorscheme)|指定图表的配色方案。|
||[roundedCorners](/javascript/api/excel/excel.chartareaformat#roundedcorners)|指定图表的图表区是否有圆角。|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[linkNumberFormat](/javascript/api/excel/excel.chartaxis#linknumberformat)|指定数字格式是否链接到单元格。|
|[ChartBinOptions](/javascript/api/excel/excel.chartbinoptions)|[allowOverflow](/javascript/api/excel/excel.chartbinoptions#allowoverflow)|指定直方图图表或排列图表中是否启用了 "bin 溢出"。|
||[allowUnderflow](/javascript/api/excel/excel.chartbinoptions#allowunderflow)|指定直方图图表或排列图表中是否启用了 bin 下溢。|
||[count](/javascript/api/excel/excel.chartbinoptions#count)|指定直方图图表或排列图表的分类数。|
||[overflowValue](/javascript/api/excel/excel.chartbinoptions#overflowvalue)|指定直方图图表或排列图表的 "bin 溢出" 值。|
||[type](/javascript/api/excel/excel.chartbinoptions#type)|指定用于直方图图表或排列图表的 bin 类型。|
||[underflowValue](/javascript/api/excel/excel.chartbinoptions#underflowvalue)|指定直方图图表或排列图表的 bin 下溢值。|
||[width](/javascript/api/excel/excel.chartbinoptions#width)|指定直方图图表或排列图表的纸盒宽度值。|
|[ChartBoxwhiskerOptions](/javascript/api/excel/excel.chartboxwhiskeroptions)|[quartileCalculation](/javascript/api/excel/excel.chartboxwhiskeroptions#quartilecalculation)|指定 box 和形图的四分四个计算类型。|
||[showInnerPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#showinnerpoints)|指定是否在 box 和形图中显示内部点。|
||[showMeanLine](/javascript/api/excel/excel.chartboxwhiskeroptions#showmeanline)|指定是否在 box 和线形图中显示 "均值" 线。|
||[showMeanMarker](/javascript/api/excel/excel.chartboxwhiskeroptions#showmeanmarker)|指定是否在框和形图中显示 "均值" 标记。|
||[showOutlierPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#showoutlierpoints)|指定框和形图中是否显示了异常点。|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabel#linknumberformat)|指定数字格式是否链接到单元格 (以便在标签中的) 单元格区域发生更改时，数字格式的更改将发生变化。|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabels#linknumberformat)|指定数字格式是否链接到单元格。|
|[ChartErrorBars](/javascript/api/excel/excel.charterrorbars)|[endStyleCap](/javascript/api/excel/excel.charterrorbars#endstylecap)|指定误差线是否具有最终样式帽。|
||[include](/javascript/api/excel/excel.charterrorbars#include)|指定包含误差线的哪些部分。|
||[format](/javascript/api/excel/excel.charterrorbars#format)|指定误差线的格式类型。|
||[type](/javascript/api/excel/excel.charterrorbars#type)|误差线标记的区域类型。|
||[visible](/javascript/api/excel/excel.charterrorbars#visible)|指定是否显示误差线。|
|[ChartErrorBarsFormat](/javascript/api/excel/excel.charterrorbarsformat)|[line](/javascript/api/excel/excel.charterrorbarsformat#line)|表示图表线条格式。|
|[ChartMapOptions](/javascript/api/excel/excel.chartmapoptions)|[labelStrategy](/javascript/api/excel/excel.chartmapoptions#labelstrategy)|指定区域地图图表的系列地图标签策略。|
||[level](/javascript/api/excel/excel.chartmapoptions#level)|指定区域映射图表的系列映射级别。|
||[projectionType](/javascript/api/excel/excel.chartmapoptions#projectiontype)|指定区域地图图表的系列投影类型。|
|[ChartPivotOptions](/javascript/api/excel/excel.chartpivotoptions)|[showAxisFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showaxisfieldbuttons)|指定是否在数据透视图中显示轴字段按钮。|
||[showLegendFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showlegendfieldbuttons)|指定是否在数据透视图上显示图例字段按钮。|
||[showReportFilterFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showreportfilterfieldbuttons)|指定是否在数据透视图上显示报表筛选字段按钮。|
||[showValueFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showvaluefieldbuttons)|指定是否在数据透视图中显示 "显示值" 字段按钮。|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[bubbleScale](/javascript/api/excel/excel.chartseries#bubblescale)|这可以是从 0（零）到 300 的整数值，表示默认大小的百分比。|
||[gradientMaximumColor](/javascript/api/excel/excel.chartseries#gradientmaximumcolor)|指定区域地图图表系列的最大值的颜色。|
||[gradientMaximumType](/javascript/api/excel/excel.chartseries#gradientmaximumtype)|指定区域地图图表系列的最大值的类型。|
||[gradientMaximumValue](/javascript/api/excel/excel.chartseries#gradientmaximumvalue)|指定区域地图图表系列的最大值。|
||[gradientMidpointColor](/javascript/api/excel/excel.chartseries#gradientmidpointcolor)|指定区域地图图表系列的中点值的颜色。|
||[gradientMidpointType](/javascript/api/excel/excel.chartseries#gradientmidpointtype)|指定区域地图图表系列的中点值的类型。|
||[gradientMidpointValue](/javascript/api/excel/excel.chartseries#gradientmidpointvalue)|指定区域地图图表系列的中点值。|
||[gradientMinimumColor](/javascript/api/excel/excel.chartseries#gradientminimumcolor)|指定区域地图图表系列的最小值的颜色。|
||[gradientMinimumType](/javascript/api/excel/excel.chartseries#gradientminimumtype)|指定区域地图图表系列的最小值的类型。|
||[gradientMinimumValue](/javascript/api/excel/excel.chartseries#gradientminimumvalue)|指定区域地图图表系列的最小值。|
||[gradientStyle](/javascript/api/excel/excel.chartseries#gradientstyle)|指定区域地图图表的系列渐变样式。|
||[invertColor](/javascript/api/excel/excel.chartseries#invertcolor)|指定系列中的负数据点的填充颜色。|
||[parentLabelStrategy](/javascript/api/excel/excel.chartseries#parentlabelstrategy)|指定树状图图表的系列父标签策略区域。|
||[binOptions](/javascript/api/excel/excel.chartseries#binoptions)|封装直方图和排列图的容器选项。|
||[boxwhiskerOptions](/javascript/api/excel/excel.chartseries#boxwhiskeroptions)|封装箱形图的选项。|
||[mapOptions](/javascript/api/excel/excel.chartseries#mapoptions)|封装区域地图图表的选项。|
||[xErrorBars](/javascript/api/excel/excel.chartseries#xerrorbars)|表示图表系列的误差线对象。|
||[yErrorBars](/javascript/api/excel/excel.chartseries#yerrorbars)|表示图表系列的误差线对象。|
||[showConnectorLines](/javascript/api/excel/excel.chartseries#showconnectorlines)|指定是否在瀑布图中显示连接符线。|
||[showLeaderLines](/javascript/api/excel/excel.chartseries#showleaderlines)|指定是否为系列中的每个数据标签显示引导线。|
||[splitValue](/javascript/api/excel/excel.chartseries#splitvalue)|指定用于分隔复合饼图或复合条饼图中的两个节的临界值。|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[linkNumberFormat](/javascript/api/excel/excel.charttrendlinelabel#linknumberformat)|指定数字格式是否链接到单元格 (以便在标签中的) 单元格区域发生更改时，数字格式的更改将发生变化。|
|[ColumnProperties](/javascript/api/excel/excel.columnproperties)|[address](/javascript/api/excel/excel.columnproperties#address)|表示`address`属性。|
||[addressLocal](/javascript/api/excel/excel.columnproperties#addresslocal)|表示`addressLocal`属性。|
||[columnIndex](/javascript/api/excel/excel.columnproperties#columnindex)|表示`columnIndex`属性。|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[getRanges()](/javascript/api/excel/excel.conditionalformat#getranges--)|返回将为其应用条件格式的 RangeAreas，它包含一个或多个矩形区域。|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[getInvalidCells()](/javascript/api/excel/excel.datavalidation#getinvalidcells--)|返回包含一个或多个矩形区域的 RangeAreas，它具有无效单元格值。|
||[getInvalidCellsOrNullObject()](/javascript/api/excel/excel.datavalidation#getinvalidcellsornullobject--)|返回包含一个或多个矩形区域的 RangeAreas，它具有无效单元格值。|
|[FilterCriteria](/javascript/api/excel/excel.filtercriteria)|[subField](/javascript/api/excel/excel.filtercriteria#subfield)|筛选器使用该属性对 richvalue 执行丰富的筛选。|
|[GeometricShape](/javascript/api/excel/excel.geometricshape)|[id](/javascript/api/excel/excel.geometricshape#id)|返回形状标识符。|
||[shape](/javascript/api/excel/excel.geometricshape#shape)|返回几何形状的形状对象。|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getCount()](/javascript/api/excel/excel.groupshapecollection#getcount--)|返回形状组中的形状数量。|
||[getItem(key: string)](/javascript/api/excel/excel.groupshapecollection#getitem-key-)|按名称或 ID 获取形状。|
||[getItemAt(index: number)](/javascript/api/excel/excel.groupshapecollection#getitemat-index-)|根据其在集合中的位置获取形状。|
||[items](/javascript/api/excel/excel.groupshapecollection#items)|获取此集合中已加载的子项。|
|[HeaderFooter](/javascript/api/excel/excel.headerfooter)|[centerFooter](/javascript/api/excel/excel.headerfooter#centerfooter)|工作表的中心页脚。|
||[centerHeader](/javascript/api/excel/excel.headerfooter#centerheader)|工作表的中间标头。|
||[LeftFooter](/javascript/api/excel/excel.headerfooter#leftfooter)|工作表的左页脚。|
||[leftHeader](/javascript/api/excel/excel.headerfooter#leftheader)|工作表的左页眉。|
||[rightFooter](/javascript/api/excel/excel.headerfooter#rightfooter)|工作表的右页脚。|
||[rightHeader](/javascript/api/excel/excel.headerfooter#rightheader)|工作表的右侧标头。|
|[HeaderFooterGroup](/javascript/api/excel/excel.headerfootergroup)|[defaultForAllPages](/javascript/api/excel/excel.headerfootergroup#defaultforallpages)|常规页眉/页脚，除非指定偶数页/奇数页或首页，否则适用于所有页面，|
||[evenPages](/javascript/api/excel/excel.headerfootergroup#evenpages)|用于偶数页的页眉/页脚，需要为奇数页指定奇数页页眉/页脚。|
||[firstPage](/javascript/api/excel/excel.headerfootergroup#firstpage)|首页的页眉/页脚，为所有其他页使用常规或偶数页/奇数页页眉/页脚。|
||[oddPages](/javascript/api/excel/excel.headerfootergroup#oddpages)|用于奇数页的页眉/页脚，需要为偶数页指定偶数页页眉/页脚。|
||[state](/javascript/api/excel/excel.headerfootergroup#state)|页眉/页脚的设置状态。|
||[useSheetMargins](/javascript/api/excel/excel.headerfootergroup#usesheetmargins)|获取或设置一个标记，指示页眉/页脚是否与工作表的页面布局选项中设置的页边距对齐。|
||[useSheetScale](/javascript/api/excel/excel.headerfootergroup#usesheetscale)|获取或设置一个标记，指示是否应按照工作表的页面布局选项中设置的页面缩放百分比来缩放页眉/页脚。|
|[Image](/javascript/api/excel/excel.image)|[format](/javascript/api/excel/excel.image#format)|返回图像的格式。|
||[id](/javascript/api/excel/excel.image#id)|指定 image 对象的形状标识符。|
||[shape](/javascript/api/excel/excel.image#shape)|返回与图像关联的形状对象。|
|[IterativeCalculation](/javascript/api/excel/excel.iterativecalculation)|[enabled](/javascript/api/excel/excel.iterativecalculation#enabled)|如果 Excel 使用迭代来处理循环引用，则为 True。|
||[maxChange](/javascript/api/excel/excel.iterativecalculation#maxchange)|指定在 Excel 解析循环引用时，每次迭代之间的最大更改量。|
||[maxIteration](/javascript/api/excel/excel.iterativecalculation#maxiteration)|指定 Excel 可用于解析循环引用的最大迭代次数。|
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
||[beginConnectedShape](/javascript/api/excel/excel.line#beginconnectedshape)|表示指定线条始端所附加到的形状。|
||[beginConnectedSite](/javascript/api/excel/excel.line#beginconnectedsite)|表示连接线始端所连接的连接站点。|
||[endConnectedShape](/javascript/api/excel/excel.line#endconnectedshape)|表示指定线条末端所附加到的形状。|
||[endConnectedSite](/javascript/api/excel/excel.line#endconnectedsite)|表示连接线末端所连接的连接站点。|
||[id](/javascript/api/excel/excel.line#id)|指定形状标识符。|
||[isBeginConnected](/javascript/api/excel/excel.line#isbeginconnected)|指定是否将指定线条的起点连接到形状。|
||[isEndConnected](/javascript/api/excel/excel.line#isendconnected)|指定是否将指定线条的终点连接到形状。|
||[shape](/javascript/api/excel/excel.line#shape)|返回与线条关联的形状对象。|
|[PageBreak](/javascript/api/excel/excel.pagebreak)|[delete()](/javascript/api/excel/excel.pagebreak#delete--)|删除分页符对象。|
||[getCellAfterBreak()](/javascript/api/excel/excel.pagebreak#getcellafterbreak--)|获取分页符后的第一个单元格。|
||[columnIndex](/javascript/api/excel/excel.pagebreak#columnindex)|指定分页符的列索引|
||[rowIndex](/javascript/api/excel/excel.pagebreak#rowindex)|指定分页符的行索引|
|[PageBreakCollection](/javascript/api/excel/excel.pagebreakcollection)|[add(pageBreakRange: Range \| string)](/javascript/api/excel/excel.pagebreakcollection#add-pagebreakrange-)|在指定区域的左上角单元格之前添加分页符。|
||[getCount()](/javascript/api/excel/excel.pagebreakcollection#getcount--)|获取集合中的分页符数量。|
||[getItem(index: number)](/javascript/api/excel/excel.pagebreakcollection#getitem-index-)|通过索引获取分页符对象。|
||[items](/javascript/api/excel/excel.pagebreakcollection#items)|获取此集合中已加载的子项。|
||[removePageBreaks()](/javascript/api/excel/excel.pagebreakcollection#removepagebreaks--)|重置集合中的所有手动分页符。|
|[PageLayout](/javascript/api/excel/excel.pagelayout)|[blackAndWhite](/javascript/api/excel/excel.pagelayout#blackandwhite)|工作表的黑色和白色打印选项。|
||[bottomMargin](/javascript/api/excel/excel.pagelayout#bottommargin)|工作表的下一页边距，用于以磅为单位打印。|
||[centerHorizontally](/javascript/api/excel/excel.pagelayout#centerhorizontally)|工作表的水平标记。|
||[centerVertically](/javascript/api/excel/excel.pagelayout#centervertically)|工作表的垂直居中标志。|
||[draftMode](/javascript/api/excel/excel.pagelayout#draftmode)|工作表的草稿模式选项。|
||[firstPageNumber](/javascript/api/excel/excel.pagelayout#firstpagenumber)|要打印的工作表的第一个页码。|
||[footerMargin](/javascript/api/excel/excel.pagelayout#footermargin)|工作表的页脚边距（以磅为单位）在打印时使用。|
||[getPrintArea()](/javascript/api/excel/excel.pagelayout#getprintarea--)|获取包含一个或多个矩形区域的 RangeAreas 对象，它表示工作表的打印区域。|
||[getPrintAreaOrNullObject()](/javascript/api/excel/excel.pagelayout#getprintareaornullobject--)|获取包含一个或多个矩形区域的 RangeAreas 对象，它表示工作表的打印区域。|
||[getPrintTitleColumns()](/javascript/api/excel/excel.pagelayout#getprinttitlecolumns--)|获取表示标题列的 Range 对象。|
||[getPrintTitleColumnsOrNullObject()](/javascript/api/excel/excel.pagelayout#getprinttitlecolumnsornullobject--)|获取表示标题列的 Range 对象。|
||[getPrintTitleRows()](/javascript/api/excel/excel.pagelayout#getprinttitlerows--)|获取表示标题行的 Range 对象。|
||[getPrintTitleRowsOrNullObject()](/javascript/api/excel/excel.pagelayout#getprinttitlerowsornullobject--)|获取表示标题行的 Range 对象。|
||[headerMargin](/javascript/api/excel/excel.pagelayout#headermargin)|工作表的页眉边距（以磅为单位）在打印时使用。|
||[leftMargin](/javascript/api/excel/excel.pagelayout#leftmargin)|工作表的左边距（以磅为单位）在打印时使用。|
||[orientation](/javascript/api/excel/excel.pagelayout#orientation)|页面的工作表方向。|
||[paperSize](/javascript/api/excel/excel.pagelayout#papersize)|页面的工作表纸张大小。|
||[printComments](/javascript/api/excel/excel.pagelayout#printcomments)|指定是否应在打印时显示工作表的批注。|
||[printErrors](/javascript/api/excel/excel.pagelayout#printerrors)|工作表的 "打印错误" 选项。|
||[printGridlines](/javascript/api/excel/excel.pagelayout#printgridlines)|指定是否打印工作表的网格线。|
||[printHeadings](/javascript/api/excel/excel.pagelayout#printheadings)|指定是否将打印工作表的标题。|
||[printOrder](/javascript/api/excel/excel.pagelayout#printorder)|工作表的 "页面打印顺序" 选项。|
||[headersFooters](/javascript/api/excel/excel.pagelayout#headersfooters)|工作表的页眉和页脚配置。|
||[rightMargin](/javascript/api/excel/excel.pagelayout#rightmargin)|工作表的右边距（以磅为单位）在打印时使用。|
||[setPrintArea(printArea: Range \| RangeAreas \| string)](/javascript/api/excel/excel.pagelayout#setprintarea-printarea-)|设置工作表的打印区域。|
||[setPrintMargins(unit: Excel.PrintMarginUnit, marginOptions: Excel.PageLayoutMarginOptions)](/javascript/api/excel/excel.pagelayout#setprintmargins-unit--marginoptions-)|设置带单位的工作表的页边距。|
||[setPrintTitleColumns(printTitleColumns: Range \| string)](/javascript/api/excel/excel.pagelayout#setprinttitlecolumns-printtitlecolumns-)|设置列，这些列包含要在打印的工作表的每页左侧重复的单元格。|
||[setPrintTitleRows(printTitleRows: Range \| string)](/javascript/api/excel/excel.pagelayout#setprinttitlerows-printtitlerows-)|设置行，这些行包含要在打印的工作表的每页顶部重复的单元格。|
||[topMargin](/javascript/api/excel/excel.pagelayout#topmargin)|工作表的上边距（以磅为单位）在打印时使用。|
||[zoom](/javascript/api/excel/excel.pagelayout#zoom)|工作表的打印缩放选项。|
|[PageLayoutMarginOptions](/javascript/api/excel/excel.pagelayoutmarginoptions)|[bottom](/javascript/api/excel/excel.pagelayoutmarginoptions#bottom)|指定要用于打印的单位的页面下边距（单位为单位）。|
||[footer](/javascript/api/excel/excel.pagelayoutmarginoptions#footer)|指定要用于打印的单位的页面布局页脚边距。|
||[header](/javascript/api/excel/excel.pagelayoutmarginoptions#header)|以指定要用于打印的单位指定页面布局标题边距。|
||[left](/javascript/api/excel/excel.pagelayoutmarginoptions#left)|按指定要用于打印的单位指定页面布局的左边距。|
||[right](/javascript/api/excel/excel.pagelayoutmarginoptions#right)|指定要用于打印的单位的页面（以指定单位为右边距）。|
||[top](/javascript/api/excel/excel.pagelayoutmarginoptions#top)|指定要用于打印的单位的页面上边距（按指定单位）。|
|[PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions)|[horizontalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#horizontalfittopages)|水平放置的页数。|
||[scale](/javascript/api/excel/excel.pagelayoutzoomoptions#scale)|打印页面缩放值可以介于 10 至 400 之间。|
||[verticalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#verticalfittopages)|垂直放置的页数。|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[sortByValues(sortBy: Excel.SortBy, valuesHierarchy: Excel.DataPivotHierarchy, pivotItemScope?: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotfield#sortbyvalues-sortby--valueshierarchy--pivotitemscope-)|按给定范围中的指定值对 PivotField 进行排序。|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[autoFormat](/javascript/api/excel/excel.pivotlayout#autoformat)|指定在刷新或移动域时是否自动设置格式。|
||[getDataHierarchy(cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#getdatahierarchy-cell-)|获取 DataHierarchy，它用于计算数据透视表中指定区域内的值。|
||[getPivotItems(axis: Excel.PivotAxis, cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#getpivotitems-axis--cell-)|从构成数据透视表中指定区域内的值的轴获取 PivotItems。|
||[preserveFormatting](/javascript/api/excel/excel.pivotlayout#preserveformatting)|指定当通过透视、排序或更改页字段项等操作刷新或重新计算报表时，是否保留格式设置。|
||[setAutoSortOnCell(cell: Range \| string, sortBy: Excel.SortBy)](/javascript/api/excel/excel.pivotlayout#setautosortoncell-cell--sortby-)|将数据透视表设置为使用指定的单元格设置自动排序，以自动选择排序的所有条件和上下文。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[enableDataValueEditing](/javascript/api/excel/excel.pivottable#enabledatavalueediting)|指定数据透视表是否允许用户对数据正文中的值进行编辑。|
||[useCustomSortLists](/javascript/api/excel/excel.pivottable#usecustomsortlists)|指定在排序时数据透视表是否使用自定义列表。|
|[Range](/javascript/api/excel/excel.range)|[自动填充 (destinationRange？： Range \| string，autoFillType？： autoFillType) ](/javascript/api/excel/excel.range#autofill-destinationrange--autofilltype-)|使用指定的自动填充逻辑将范围从当前范围填充到目标区域。|
||[convertDataTypeToText()](/javascript/api/excel/excel.range#convertdatatypetotext--)|将具有数据类型的区域单元格转换为文本。|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.range#converttolinkeddatatype-serviceid--languageculture-)|将区域单元格转换为工作表中的链接数据类型。|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-)|将单元格数据或格式从源区域或 RangeAreas 复制到当前区域。|
||[find(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#find-text--criteria-)|根据指定的条件查找给定的字符串。|
||[findOrNullObject(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#findornullobject-text--criteria-)|根据指定的条件查找给定的字符串。|
||[flashFill()](/javascript/api/excel/excel.range#flashfill--)|对当前区域进行快速填充。快速填充在感知到模式时可自动填充数据，因此该区域必须是单列区域且周围有数据以便查找模式。|
||[getCellProperties(cellPropertiesLoadOptions: CellPropertiesLoadOptions)](/javascript/api/excel/excel.range#getcellproperties-cellpropertiesloadoptions-)|返回一个 2D 数组，其中封装了每个单元格的字体、填充、边框、对齐方式和其他属性数据。|
||[getColumnProperties(columnPropertiesLoadOptions: ColumnPropertiesLoadOptions)](/javascript/api/excel/excel.range#getcolumnproperties-columnpropertiesloadoptions-)|返回一个一维数组，其中封装了每个列的字体、填充、边框、对齐方式和其他属性数据。|
||[getRowProperties(rowPropertiesLoadOptions: RowPropertiesLoadOptions)](/javascript/api/excel/excel.range#getrowproperties-rowpropertiesloadoptions-)|返回一个一维数组，其中封装了每个行的字体、填充、边框、对齐方式和其他属性数据。|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-)|获取包含一个或多个矩形区域的 RangeAreas 对象，它表示匹配指定类型和值的所有单元格。|
||[getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#getspecialcellsornullobject-celltype--cellvaluetype-)|获取包含一个或多个区域的 RangeAreas 对象，它表示匹配指定类型和值的所有单元格。|
||[getTables(fullyContained?: boolean)](/javascript/api/excel/excel.range#gettables-fullycontained-)|获取与区域重叠的限定范围的表格集合。|
||[linkedDataTypeState](/javascript/api/excel/excel.range#linkeddatatypestate)|表示每个单元格的数据类型状态。|
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
||[getIntersection(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#getintersection-anotherrange-)|返回 RangeAreas 对象，它表示给定区域或 RangeAreas 的交集。|
||[getIntersectionOrNullObject(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#getintersectionornullobject-anotherrange-)|返回 RangeAreas 对象，它表示给定区域或 RangeAreas 的交集。|
||[getOffsetRangeAreas(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.rangeareas#getoffsetrangeareas-rowoffset--columnoffset-)|返回 RangeAreas 对象，它按特定的行和列偏移量进行移动。|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#getspecialcells-celltype--cellvaluetype-)|返回一个 RangeAreas 对象，它表示匹配指定类型和值的所有单元格。|
||[getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#getspecialcellsornullobject-celltype--cellvaluetype-)|返回一个 RangeAreas 对象，它表示匹配指定类型和值的所有单元格。|
||[getTables(fullyContained?: boolean)](/javascript/api/excel/excel.rangeareas#gettables-fullycontained-)|返回与此 RangeAreas 对象中的任何区域重叠的限定范围的表格集合。|
||[getUsedRangeAreas(valuesOnly?: boolean)](/javascript/api/excel/excel.rangeareas#getusedrangeareas-valuesonly-)|返回使用的 RangeAreas，它包含 RangeAreas 对象中的各个矩形区域的所有已用区域。|
||[getUsedRangeAreasOrNullObject(valuesOnly?: boolean)](/javascript/api/excel/excel.rangeareas#getusedrangeareasornullobject-valuesonly-)|返回使用的 RangeAreas，它包含 RangeAreas 对象中的各个矩形区域的所有已用区域。|
||[address](/javascript/api/excel/excel.rangeareas#address)|返回 A1 样式中的 RangeAreas 引用。|
||[addressLocal](/javascript/api/excel/excel.rangeareas#addresslocal)|返回用户区域设置中的 RangeAreas 引用。|
||[areaCount](/javascript/api/excel/excel.rangeareas#areacount)|返回包含此 RangeAreas 对象的矩形区域的数量。|
||[areas](/javascript/api/excel/excel.rangeareas#areas)|返回包含此 RangeAreas 对象的矩形区域的集合。|
||[cellCount](/javascript/api/excel/excel.rangeareas#cellcount)|返回 RangeAreas 对象中的单元格数量，即总计各个矩形区域的单元格计数。|
||[conditionalFormats](/javascript/api/excel/excel.rangeareas#conditionalformats)|返回与此 RangeAreas 对象中的任何单元格相交的 ConditionalFormats 集合。|
||[dataValidation](/javascript/api/excel/excel.rangeareas#datavalidation)|返回 RangeAreas 中的所有区域的 dataValidation 对象。|
||[format](/javascript/api/excel/excel.rangeareas#format)|返回一个 RangeFormat 对象，封装 RangeAreas 对象中所有区域的字体、填充、边框、对齐方式和其他属性。|
||[isEntireColumn](/javascript/api/excel/excel.rangeareas#isentirecolumn)|指定此 RangeAreas 对象上的所有区域是否都代表整列 (例如，"A:C，Q:Z" ) 。|
||[isEntireRow](/javascript/api/excel/excel.rangeareas#isentirerow)|指定此 RangeAreas 对象上的所有区域是否代表整行 (例如，"1:3，5:7" ) 。|
||[worksheet](/javascript/api/excel/excel.rangeareas#worksheet)|返回当前 RangeAreas 的工作表。|
||[setDirty()](/javascript/api/excel/excel.rangeareas#setdirty--)|设置要在下一次重新计算时重新进行计算的 RangeAreas。|
||[style](/javascript/api/excel/excel.rangeareas#style)|表示此 RangeAreas 对象中的所有区域的样式。|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[tintAndShade](/javascript/api/excel/excel.rangeborder#tintandshade)|指定将区域边框的颜色变浅或变暗的双精度值，该值介于-1 (最暗) 和 1 (最明亮的) ，0表示原始颜色。|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[tintAndShade](/javascript/api/excel/excel.rangebordercollection#tintandshade)|指定将区域边框的颜色变浅或变暗的双精度值，该值介于-1 (最暗) 和 1 (最明亮的) ，0表示原始颜色。|
|[RangeCollection](/javascript/api/excel/excel.rangecollection)|[getCount()](/javascript/api/excel/excel.rangecollection#getcount--)|返回 RangeCollection 中的区域数量。|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangecollection#getitemat-index-)|根据其在 RangeCollection 中的位置返回 Range 对象。|
||[items](/javascript/api/excel/excel.rangecollection#items)|获取此集合中已加载的子项。|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[pattern](/javascript/api/excel/excel.rangefill#pattern)|区域的图案。|
||[patternColor](/javascript/api/excel/excel.rangefill#patterncolor)|表示范围模式的颜色的 HTML 颜色代码，格式 #RRGGBB (如 "FFA500" ) 或命名的 HTML 颜色 (例如 "橙色" ) 。|
||[patternTintAndShade](/javascript/api/excel/excel.rangefill#patterntintandshade)|指定用于使区域填充的图案颜色变浅或变暗的双精度值，该值介于-1 (最暗) 和 1 (明亮的) ，0表示原始颜色。|
||[tintAndShade](/javascript/api/excel/excel.rangefill#tintandshade)|指定用于将区域填充的颜色变浅或变暗的双精度值，该值介于-1 (最暗) 和 1 (明亮的) 之间，0表示原始颜色。|
|[RangeFont](/javascript/api/excel/excel.rangefont)|[strikethrough](/javascript/api/excel/excel.rangefont#strikethrough)|指定字体的删除线状态。|
||[subscript](/javascript/api/excel/excel.rangefont#subscript)|指定字体的下标状态。|
||[superscript](/javascript/api/excel/excel.rangefont#superscript)|指定字体的上标状态。|
||[tintAndShade](/javascript/api/excel/excel.rangefont#tintandshade)|指定将区域字体的颜色变浅或变暗的双精度值，该值介于-1 (最暗) 和 1 (最明亮的) ，0表示原始颜色。|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[autoIndent](/javascript/api/excel/excel.rangeformat#autoindent)|指定当文本对齐方式设置为相等分布时，文本是否自动缩进。|
||[indentLevel](/javascript/api/excel/excel.rangeformat#indentlevel)|0 到 250 之间的一个整数，指示缩进水平。|
||[readingOrder](/javascript/api/excel/excel.rangeformat#readingorder)|区域的读取顺序。|
||[shrinkToFit](/javascript/api/excel/excel.rangeformat#shrinktofit)|指定文本是否自动收缩以显示可用列宽。|
|[RemoveDuplicatesResult](/javascript/api/excel/excel.removeduplicatesresult)|[removed](/javascript/api/excel/excel.removeduplicatesresult#removed)|由操作删除的重复行数。|
||[uniqueRemaining](/javascript/api/excel/excel.removeduplicatesresult#uniqueremaining)|所生成的区域中存在的剩余唯一行数。|
|[ReplaceCriteria](/javascript/api/excel/excel.replacecriteria)|[completeMatch](/javascript/api/excel/excel.replacecriteria#completematch)|指定是否需要完成匹配或部分匹配。|
||[matchCase](/javascript/api/excel/excel.replacecriteria#matchcase)|指定匹配是否区分大小写。|
|[RowProperties](/javascript/api/excel/excel.rowproperties)|[address](/javascript/api/excel/excel.rowproperties#address)|表示`address`属性。|
||[addressLocal](/javascript/api/excel/excel.rowproperties#addresslocal)|表示`addressLocal`属性。|
||[rowIndex](/javascript/api/excel/excel.rowproperties#rowindex)|表示`rowIndex`属性。|
|[SearchCriteria](/javascript/api/excel/excel.searchcriteria)|[completeMatch](/javascript/api/excel/excel.searchcriteria#completematch)|指定是否需要完成匹配或部分匹配。|
||[matchCase](/javascript/api/excel/excel.searchcriteria#matchcase)|指定匹配是否区分大小写。|
||[searchDirection](/javascript/api/excel/excel.searchcriteria#searchdirection)|指定搜索方向。|
|[SettableCellProperties](/javascript/api/excel/excel.settablecellproperties)|[format](/javascript/api/excel/excel.settablecellproperties#format)|表示`format`属性。|
||[hyperlink](/javascript/api/excel/excel.settablecellproperties#hyperlink)|表示`hyperlink`属性。|
||[style](/javascript/api/excel/excel.settablecellproperties#style)|表示`style`属性。|
|[SettableColumnProperties](/javascript/api/excel/excel.settablecolumnproperties)|[columnHidden](/javascript/api/excel/excel.settablecolumnproperties#columnhidden)|表示`columnHidden`属性。|
||[columnWidth](/javascript/api/excel/excel.settablecolumnproperties#columnwidth)||
||[格式： CellPropertiesFormat & {columnWidth？](/javascript/api/excel/excel.settablecolumnproperties#format)|表示`format`属性。|
|[SettableRowProperties](/javascript/api/excel/excel.settablerowproperties)|[格式： CellPropertiesFormat & {rowHeight？](/javascript/api/excel/excel.settablerowproperties#format)|表示`format`属性。|
||[rowHeight](/javascript/api/excel/excel.settablerowproperties#rowheight)||
||[rowHidden](/javascript/api/excel/excel.settablerowproperties#rowhidden)|表示`rowHidden`属性。|
|[Shape](/javascript/api/excel/excel.shape)|[altTextDescription](/javascript/api/excel/excel.shape#alttextdescription)|指定 Shape 对象的替代说明文本。|
||[altTextTitle](/javascript/api/excel/excel.shape#alttexttitle)|为 Shape 对象指定可选的标题文本。|
||[delete()](/javascript/api/excel/excel.shape#delete--)|从工作表删除形状。|
||[geometricShapeType](/javascript/api/excel/excel.shape#geometricshapetype)|指定此几何形状的几何形状类型。|
||[getAsImage(format: Excel.PictureFormat)](/javascript/api/excel/excel.shape#getasimage-format-)|将形状转换为图像并将图像返回为 base64 编码字符串。|
||[height](/javascript/api/excel/excel.shape#height)|指定形状的高度（以磅为单位）。|
||[incrementLeft(increment: number)](/javascript/api/excel/excel.shape#incrementleft-increment-)|以指定磅数水平移动形状。|
||[incrementRotation(increment: number)](/javascript/api/excel/excel.shape#incrementrotation-increment-)|将形状围绕 z 轴旋转特定度数。|
||[incrementTop(increment: number)](/javascript/api/excel/excel.shape#incrementtop-increment-)|以指定磅数垂直移动形状。|
||[left](/javascript/api/excel/excel.shape#left)|从形状左侧到工作表左侧的距离（以磅为单位）。|
||[lockAspectRatio](/javascript/api/excel/excel.shape#lockaspectratio)|指定此形状的纵横比是否已锁定。|
||[name](/javascript/api/excel/excel.shape#name)|指定形状的名称。|
||[connectionSiteCount](/javascript/api/excel/excel.shape#connectionsitecount)|返回此形状上的连接站点数。|
||[fill](/javascript/api/excel/excel.shape#fill)|返回此形状的填充格式。|
||[geometricShape](/javascript/api/excel/excel.shape#geometricshape)|返回与形状关联的几何形状。|
||[组](/javascript/api/excel/excel.shape#group)|返回与形状关联的形状组。|
||[id](/javascript/api/excel/excel.shape#id)|指定形状标识符。|
||[image](/javascript/api/excel/excel.shape#image)|返回与形状关联的图像。|
||[level](/javascript/api/excel/excel.shape#level)|指定指定形状的级别。|
||[line](/javascript/api/excel/excel.shape#line)|返回与形状关联的线条。|
||[lineFormat](/javascript/api/excel/excel.shape#lineformat)|返回此形状的线条格式。|
||[onActivated](/javascript/api/excel/excel.shape#onactivated)|当激活形状时发生此事件。|
||[onDeactivated](/javascript/api/excel/excel.shape#ondeactivated)|当停用形状时发生此事件。|
||[parentGroup](/javascript/api/excel/excel.shape#parentgroup)|指定此形状的父组。|
||[textFrame](/javascript/api/excel/excel.shape#textframe)|返回此形状的文本框对象。|
||[type](/javascript/api/excel/excel.shape#type)|返回此形状的类型。|
||[zOrderPosition](/javascript/api/excel/excel.shape#zorderposition)|返回指定形状在 z 顺序中的位置，0 表示顺序堆栈的底部。|
||[rotation](/javascript/api/excel/excel.shape#rotation)|指定形状的旋转角度（以度为单位）。|
||[scaleHeight(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#scaleheight-scalefactor--scaletype--scalefrom-)|按指定因子缩放形状的高度。|
||[scaleWidth(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#scalewidth-scalefactor--scaletype--scalefrom-)|按指定因子缩放形状的宽度。|
||[setZOrder(position: Excel.ShapeZOrder)](/javascript/api/excel/excel.shape#setzorder-position-)|将指定形状沿集合的 z 顺序向上或向下移动，将其移动到其他形状的前面或后面。|
||[top](/javascript/api/excel/excel.shape#top)|从形状上边缘到工作表上边缘之间的距离（以磅为单位）。|
||[visible](/javascript/api/excel/excel.shape#visible)|指定形状是否可见。|
||[width](/javascript/api/excel/excel.shape#width)|指定形状的宽度（以磅为单位）。|
|[ShapeActivatedEventArgs](/javascript/api/excel/excel.shapeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapeactivatedeventargs#shapeid)|获取已激活的形状的 ID。|
||[type](/javascript/api/excel/excel.shapeactivatedeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.shapeactivatedeventargs#worksheetid)|获取其中的形状已启用的工作表的 ID。|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addGeometricShape(geometricShapeType: Excel.GeometricShapeType)](/javascript/api/excel/excel.shapecollection#addgeometricshape-geometricshapetype-)|将几何形状添加到工作表。|
||[addGroup(values: Array<string \| Shape>)](/javascript/api/excel/excel.shapecollection#addgroup-values-)|在此集合的工作表中对形状的子集进行分组。|
||[addImage(base64ImageString: string)](/javascript/api/excel/excel.shapecollection#addimage-base64imagestring-)|从 base64 编码的字符串创建图像并将其添加到工作表。|
||[addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType)](/javascript/api/excel/excel.shapecollection#addline-startleft--starttop--endleft--endtop--connectortype-)|将线条添加到工作表。|
||[addTextBox(text?: string)](/javascript/api/excel/excel.shapecollection#addtextbox-text-)|使用提供的文本作为内容，将文本框添加到工作表。|
||[getCount()](/javascript/api/excel/excel.shapecollection#getcount--)|返回工作表中的形状数。|
||[getItem(key: string)](/javascript/api/excel/excel.shapecollection#getitem-key-)|按名称或 ID 获取形状。|
||[getItemAt(index: number)](/javascript/api/excel/excel.shapecollection#getitemat-index-)|使用其在集合中的位置获取形状。|
||[items](/javascript/api/excel/excel.shapecollection#items)|获取此集合中已加载的子项。|
|[ShapeDeactivatedEventArgs](/javascript/api/excel/excel.shapedeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapedeactivatedeventargs#shapeid)|获取已停用的形状的 ID。|
||[type](/javascript/api/excel/excel.shapedeactivatedeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.shapedeactivatedeventargs#worksheetid)|获取其中的形状已停用的工作表的 ID。|
|[ShapeFill](/javascript/api/excel/excel.shapefill)|[clear()](/javascript/api/excel/excel.shapefill#clear--)|清除此形状的填充格式。|
||[foregroundColor](/javascript/api/excel/excel.shapefill#foregroundcolor)|代表形状填充前景色的 HTML 颜色格式，格式 #RRGGBB (例如，"FFA500" ) 或作为命名的 HTML 颜色 (例如 "橙色" ) |
||[type](/javascript/api/excel/excel.shapefill#type)|返回形状的填充类型。|
||[setSolidColor(color: string)](/javascript/api/excel/excel.shapefill#setsolidcolor-color-)|将形状的填充格式设置为统一颜色。|
||[transparency](/javascript/api/excel/excel.shapefill#transparency)|将 fill 的透明度百分比指定为 0.0 (不透明) 到 1.0 (clear) 中的值。|
|[ShapeFont](/javascript/api/excel/excel.shapefont)|[bold](/javascript/api/excel/excel.shapefont#bold)|表示字体的加粗状态。|
||[color](/javascript/api/excel/excel.shapefont#color)|文本颜色的 HTML 颜色代码表示 (例如，"#FF0000" 表示红色) 。|
||[italic](/javascript/api/excel/excel.shapefont#italic)|表示字体的斜体状态。|
||[name](/javascript/api/excel/excel.shapefont#name)|表示字体名称 (例如，"Calibri" ) 。|
||[size](/javascript/api/excel/excel.shapefont#size)|表示字体大小，以磅为单位 (例如，11) 。|
||[underline](/javascript/api/excel/excel.shapefont#underline)|应用于字体的下划线类型。|
|[ShapeGroup](/javascript/api/excel/excel.shapegroup)|[id](/javascript/api/excel/excel.shapegroup#id)|指定形状标识符。|
||[shape](/javascript/api/excel/excel.shapegroup#shape)|返回与组关联的 Shape 对象。|
||[shapes](/javascript/api/excel/excel.shapegroup#shapes)|返回 Shape 对象的集合。|
||[ungroup()](/javascript/api/excel/excel.shapegroup#ungroup--)|取消分组指定形状组中的任何已分组形状。|
|[ShapeLineFormat](/javascript/api/excel/excel.shapelineformat)|[color](/javascript/api/excel/excel.shapelineformat#color)|代表 HTML 颜色格式的线条颜色，格式 #RRGGBB (例如，"FFA500" ) 或作为命名的 HTML 颜色 (例如，"橙色" ) 。|
||[dashStyle](/javascript/api/excel/excel.shapelineformat#dashstyle)|表示形状的线条样式。|
||[style](/javascript/api/excel/excel.shapelineformat#style)|表示形状的线条样式。|
||[transparency](/javascript/api/excel/excel.shapelineformat#transparency)|将指定线条的透明度表示为从 0.0（不透明）到 1.0（清晰）的值。|
||[visible](/javascript/api/excel/excel.shapelineformat#visible)|指定是否显示 shape 元素的线条格式。|
||[weight](/javascript/api/excel/excel.shapelineformat#weight)|表示线条的粗细（以磅为单位）。|
|[SortField](/javascript/api/excel/excel.sortfield)|[subField](/javascript/api/excel/excel.sortfield#subfield)|指定作为排序依据的格式值的目标属性名称的子字段。|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[getCount()](/javascript/api/excel/excel.stylecollection#getcount--)|获取集合中的样式数量。|
||[getItemAt(index: number)](/javascript/api/excel/excel.stylecollection#getitemat-index-)|根据其在集合中的位置获取样式。|
|[Table](/javascript/api/excel/excel.table)|[autoFilter](/javascript/api/excel/excel.table#autofilter)|表示表格的 AutoFilter 对象。|
|[TableAddedEventArgs](/javascript/api/excel/excel.tableaddedeventargs)|[source](/javascript/api/excel/excel.tableaddedeventargs#source)|获取事件源。|
||[tableId](/javascript/api/excel/excel.tableaddedeventargs#tableid)|获取已添加的表格的 ID。|
||[type](/javascript/api/excel/excel.tableaddedeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.tableaddedeventargs#worksheetid)|获取已在其中添加表格的工作表的 ID。|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[details](/javascript/api/excel/excel.tablechangedeventargs#details)|获取有关更改详细信息的信息。|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onAdded](/javascript/api/excel/excel.tablecollection#onadded)|在工作簿中添加新表格时发生。|
||[onDeleted](/javascript/api/excel/excel.tablecollection#ondeleted)|在工作簿中删除指定的表格时发生。|
|[TableDeletedEventArgs](/javascript/api/excel/excel.tabledeletedeventargs)|[source](/javascript/api/excel/excel.tabledeletedeventargs#source)|获取事件源。|
||[tableId](/javascript/api/excel/excel.tabledeletedeventargs#tableid)|获取已删除的表的 id。|
||[tableName](/javascript/api/excel/excel.tabledeletedeventargs#tablename)|获取已删除的表的名称。|
||[type](/javascript/api/excel/excel.tabledeletedeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.tabledeletedeventargs#worksheetid)|获取在其中删除表的工作表的 id。|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getCount()](/javascript/api/excel/excel.tablescopedcollection#getcount--)|获取集合中的表数量。|
||[getFirst()](/javascript/api/excel/excel.tablescopedcollection#getfirst--)|获取集合中的第一个表格。|
||[getItem(key: string)](/javascript/api/excel/excel.tablescopedcollection#getitem-key-)|按名称或 ID 获取表。|
||[items](/javascript/api/excel/excel.tablescopedcollection#items)|获取此集合中已加载的子项。|
|[TextFrame](/javascript/api/excel/excel.textframe)|[autoSizeSetting](/javascript/api/excel/excel.textframe#autosizesetting)|文本框架的自动调整大小设置。|
||[bottomMargin](/javascript/api/excel/excel.textframe#bottommargin)|表示文本框的下边距（以磅为单位）。|
||[deleteText()](/javascript/api/excel/excel.textframe#deletetext--)|删除文本框中的所有文本。|
||[horizontalAlignment](/javascript/api/excel/excel.textframe#horizontalalignment)|表示文本框的水平对齐方式。|
||[horizontalOverflow](/javascript/api/excel/excel.textframe#horizontaloverflow)|表示文本框的水平溢出行为。|
||[leftMargin](/javascript/api/excel/excel.textframe#leftmargin)|表示文本框的左边距（以磅为单位）。|
||[orientation](/javascript/api/excel/excel.textframe#orientation)|表示文本在文本框架中定向到的角度。|
||[readingOrder](/javascript/api/excel/excel.textframe#readingorder)|表示文本框从左到右或从右到左的读取顺序。|
||[hasText](/javascript/api/excel/excel.textframe#hastext)|指定文本框架是否包含文本。|
||[textRange](/javascript/api/excel/excel.textframe#textrange)|表示附加到文本框中形状上的文本，以及用于操作文本的属性和方法。|
||[rightMargin](/javascript/api/excel/excel.textframe#rightmargin)|表示文本框的右边距（以磅为单位）。|
||[topMargin](/javascript/api/excel/excel.textframe#topmargin)|表示文本框的上边距（以磅为单位）。|
||[verticalAlignment](/javascript/api/excel/excel.textframe#verticalalignment)|表示文本框的垂直对齐方式。|
||[verticalOverflow](/javascript/api/excel/excel.textframe#verticaloverflow)|表示文本框的垂直溢出行为。|
|[TextRange](/javascript/api/excel/excel.textrange)|[getSubstring(start: number, length?: number)](/javascript/api/excel/excel.textrange#getsubstring-start--length-)|返回给定区域内子字符串的 TextRange 对象。|
||[font](/javascript/api/excel/excel.textrange#font)|返回一个 ShapeFont 对象，该对象表示文本范围的字体属性。|
||[text](/javascript/api/excel/excel.textrange#text)|表示文本范围的纯文本内容。|
|[Workbook](/javascript/api/excel/excel.workbook)|[chartDataPointTrack](/javascript/api/excel/excel.workbook#chartdatapointtrack)|如果工作簿中的所有图表都跟踪它们所附加的实际数据点，则为 True。|
||[getActiveChart()](/javascript/api/excel/excel.workbook#getactivechart--)|获取工作簿中的当前活动图表。|
||[getActiveChartOrNullObject()](/javascript/api/excel/excel.workbook#getactivechartornullobject--)|获取工作簿中的当前活动图表。|
||[getIsActiveCollabSession()](/javascript/api/excel/excel.workbook#getisactivecollabsession--)|如果多个用户正在编辑工作簿（共同创作），则为 True。|
||[getSelectedRanges()](/javascript/api/excel/excel.workbook#getselectedranges--)|从工作簿中获取当前选定的一个或多个区域。|
||[isDirty](/javascript/api/excel/excel.workbook#isdirty)|指定自上次保存工作簿后是否进行了更改。|
||[autoSave](/javascript/api/excel/excel.workbook#autosave)|指定工作簿是否处于自动保存模式。|
||[calculationEngineVersion](/javascript/api/excel/excel.workbook#calculationengineversion)|返回有关 Excel 计算引擎的版本号。|
||[onAutoSaveSettingChanged](/javascript/api/excel/excel.workbook#onautosavesettingchanged)|在工作簿上更改“自动保存”设置时发生。|
||[previouslySaved](/javascript/api/excel/excel.workbook#previouslysaved)|指定是否曾在本地或联机保存工作簿。|
||[usePrecisionAsDisplayed](/javascript/api/excel/excel.workbook#useprecisionasdisplayed)|如果此工作簿中的计算仅使用显示的数字精度来完成，则为 True。|
|[WorkbookAutoSaveSettingChangedEventArgs](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs)|[type](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs#type)|获取事件的类型。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[enableCalculation](/javascript/api/excel/excel.worksheet#enablecalculation)|确定 Excel 是否应根据需要重新计算工作表。|
||[findAll(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#findall-text--criteria-)|根据指定的条件查找给定字符串的所有匹配项，并将它们作为包含一个或多个矩形区域的 RangeAreas 对象返回。|
||[findAllOrNullObject(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#findallornullobject-text--criteria-)|根据指定的条件查找给定字符串的所有匹配项，并将它们作为包含一个或多个矩形区域的 RangeAreas 对象返回。|
||[getRanges(address?: string)](/javascript/api/excel/excel.worksheet#getranges-address-)|获取按地址或名称指定的 RangeAreas 对象，它表示一个或多个矩形区域块。|
||[autoFilter](/javascript/api/excel/excel.worksheet#autofilter)|表示工作表的 AutoFilter 对象。|
||[horizontalPageBreaks](/javascript/api/excel/excel.worksheet#horizontalpagebreaks)|获取工作表的水平分页符集合。|
||[onFormatChanged](/javascript/api/excel/excel.worksheet#onformatchanged)|在特定工作表上更改格式时发生。|
||[pageLayout](/javascript/api/excel/excel.worksheet#pagelayout)|获取工作表的 PageLayout 对象。|
||[shapes](/javascript/api/excel/excel.worksheet#shapes)|返回工作表上的所有 Shape 对象的集合。|
||[verticalPageBreaks](/javascript/api/excel/excel.worksheet#verticalpagebreaks)|获取工作表的垂直分页符集合。|
||[replaceAll(text: string, replacement: string, criteria: Excel.ReplaceCriteria)](/javascript/api/excel/excel.worksheet#replaceall-text--replacement--criteria-)|根据当前工作表中指定的条件查找并替换给定的字符串。|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[details](/javascript/api/excel/excel.worksheetchangedeventargs#details)|表示有关更改详细信息的信息。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onChanged](/javascript/api/excel/excel.worksheetcollection#onchanged)|在更改工作簿中的任何工作表时发生。|
||[onFormatChanged](/javascript/api/excel/excel.worksheetcollection#onformatchanged)|在更改工作簿中的任何工作表的格式时发生。|
||[onSelectionChanged](/javascript/api/excel/excel.worksheetcollection#onselectionchanged)|在任何工作表上更改选择时发生。|
|[WorksheetFormatChangedEventArgs](/javascript/api/excel/excel.worksheetformatchangedeventargs)|[address](/javascript/api/excel/excel.worksheetformatchangedeventargs#address)|获取区域地址，该地址表示特定工作表上的更改区域。|
||[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#getrange-ctx-)|获取区域，该区域表示特定工作表上的更改区域。|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#getrangeornullobject-ctx-)|获取区域，该区域表示特定工作表上的更改区域。|
||[source](/javascript/api/excel/excel.worksheetformatchangedeventargs#source)|获取事件源。|
||[type](/javascript/api/excel/excel.worksheetformatchangedeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.worksheetformatchangedeventargs#worksheetid)|获取其中的数据发生更改的工作表的 ID。|
|[WorksheetSearchCriteria](/javascript/api/excel/excel.worksheetsearchcriteria)|[completeMatch](/javascript/api/excel/excel.worksheetsearchcriteria#completematch)|指定是否需要完成匹配或部分匹配。|
||[matchCase](/javascript/api/excel/excel.worksheetsearchcriteria#matchcase)|指定匹配是否区分大小写。|

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 参考文档](/javascript/api/excel?view=excel-js-1.9&preserve-view=true)
- [Excel JavaScript API 要求集](excel-api-requirement-sets.md)
