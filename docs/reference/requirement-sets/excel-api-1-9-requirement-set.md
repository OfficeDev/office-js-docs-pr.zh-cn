---
title: Excel JavaScript API 要求集1。9
description: 有关 ExcelApi 1.9 要求集的详细信息
ms.date: 07/11/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 1c7361debe7ba09c3477d39d9337c35bf5df3066
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/17/2019
ms.locfileid: "35772000"
---
# <a name="whats-new-in-excel-javascript-api-19"></a>Excel JavaScript API 1.9 的最近更新

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

| Class | 域 | 说明 |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[calculationEngineVersion](/javascript/api/excel/excel.application#calculationengineversion)|返回用于上次完整重新计算的 Excel 计算引擎版本。 只读。|
||[calculationState](/javascript/api/excel/excel.application#calculationstate)|返回应用程序的计算状态。 有关详细信息，请参阅 Excel.CalculationState。 只读。|
||[iterativeCalculation](/javascript/api/excel/excel.application#iterativecalculation)|返回“迭代计算”设置。|
||[suspendScreenUpdatingUntilNextSync()](/javascript/api/excel/excel.application#suspendscreenupdatinguntilnextsync--)|在下一次调用“context.sync()”前暂停屏幕更新。|
|[ApplicationData](/javascript/api/excel/excel.applicationdata)|[calculationEngineVersion](/javascript/api/excel/excel.applicationdata#calculationengineversion)|返回用于上次完整重新计算的 Excel 计算引擎版本。 只读。|
||[calculationState](/javascript/api/excel/excel.applicationdata#calculationstate)|返回应用程序的计算状态。 有关详细信息，请参阅 Excel.CalculationState。 只读。|
||[iterativeCalculation](/javascript/api/excel/excel.applicationdata#iterativecalculation)|返回“迭代计算”设置。|
|[ApplicationLoadOptions](/javascript/api/excel/excel.applicationloadoptions)|[calculationEngineVersion](/javascript/api/excel/excel.applicationloadoptions#calculationengineversion)|返回用于上次完整重新计算的 Excel 计算引擎版本。 只读。|
||[calculationState](/javascript/api/excel/excel.applicationloadoptions#calculationstate)|返回应用程序的计算状态。 有关详细信息，请参阅 Excel.CalculationState。 只读。|
||[iterativeCalculation](/javascript/api/excel/excel.applicationloadoptions#iterativecalculation)|返回“迭代计算”设置。|
|[ApplicationUpdateData](/javascript/api/excel/excel.applicationupdatedata)|[iterativeCalculation](/javascript/api/excel/excel.applicationupdatedata#iterativecalculation)|返回“迭代计算”设置。|
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[apply(range: Range \| string, columnIndex?: number, criteria?: Excel.FilterCriteria)](/javascript/api/excel/excel.autofilter#apply-range--columnindex--criteria-)|将自动筛选器应用于区域。 如果指定了列索引和筛选条件，则筛选列。|
||[clearCriteria()](/javascript/api/excel/excel.autofilter#clearcriteria--)|清除自动筛选器的筛选条件。|
||[getRange()](/javascript/api/excel/excel.autofilter#getrange--)|返回 Range 对象，该对象表示“自动筛选”应用的区域。|
||[getRangeOrNullObject()](/javascript/api/excel/excel.autofilter#getrangeornullobject--)|返回 Range 对象，该对象表示“自动筛选”应用的区域。|
||[criteria](/javascript/api/excel/excel.autofilter#criteria)|在自动筛选区域中保留所有筛选条件的数组。 只读。|
||[enabled](/javascript/api/excel/excel.autofilter#enabled)|指示是否启用了自动筛选。 只读。|
||[isDataFiltered](/javascript/api/excel/excel.autofilter#isdatafiltered)|指示自动筛选是否具有筛选条件。 只读。|
||[reapply()](/javascript/api/excel/excel.autofilter#reapply--)|应用当前位于区域上的指定 Autofilter 对象。|
||[remove()](/javascript/api/excel/excel.autofilter#remove--)|删除区域的自动筛选。|
|[AutoFilterData](/javascript/api/excel/excel.autofilterdata)|[criteria](/javascript/api/excel/excel.autofilterdata#criteria)|在自动筛选区域中保留所有筛选条件的数组。 只读。|
||[enabled](/javascript/api/excel/excel.autofilterdata#enabled)|指示是否启用了自动筛选。 只读。|
||[isDataFiltered](/javascript/api/excel/excel.autofilterdata#isdatafiltered)|指示自动筛选是否具有筛选条件。 只读。|
|[AutoFilterLoadOptions](/javascript/api/excel/excel.autofilterloadoptions)|[$all](/javascript/api/excel/excel.autofilterloadoptions#$all)||
||[criteria](/javascript/api/excel/excel.autofilterloadoptions#criteria)|在自动筛选区域中保留所有筛选条件的数组。 只读。|
||[enabled](/javascript/api/excel/excel.autofilterloadoptions#enabled)|指示是否启用了自动筛选。 只读。|
||[isDataFiltered](/javascript/api/excel/excel.autofilterloadoptions#isdatafiltered)|指示自动筛选是否具有筛选条件。 只读。|
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
|[CellPropertiesBorderLoadOptions](/javascript/api/excel/excel.cellpropertiesborderloadoptions)|[color](/javascript/api/excel/excel.cellpropertiesborderloadoptions#color)|指定是否在`color`属性上进行加载。|
||[style](/javascript/api/excel/excel.cellpropertiesborderloadoptions#style)|指定是否在`style`属性上进行加载。|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesborderloadoptions#tintandshade)|指定是否在`tintAndShade`属性上进行加载。|
||[weight](/javascript/api/excel/excel.cellpropertiesborderloadoptions#weight)|指定是否在`weight`属性上进行加载。|
|[CellPropertiesFill](/javascript/api/excel/excel.cellpropertiesfill)|[color](/javascript/api/excel/excel.cellpropertiesfill#color)|表示`format.fill.color`属性。|
||[pattern](/javascript/api/excel/excel.cellpropertiesfill#pattern)|表示`format.fill.pattern`属性。|
||[patternColor](/javascript/api/excel/excel.cellpropertiesfill#patterncolor)|表示`format.fill.patternColor`属性。|
||[patternTintAndShade](/javascript/api/excel/excel.cellpropertiesfill#patterntintandshade)|表示`format.fill.patternTintAndShade`属性。|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfill#tintandshade)|表示`format.fill.tintAndShade`属性。|
|[CellPropertiesFillLoadOptions](/javascript/api/excel/excel.cellpropertiesfillloadoptions)|[color](/javascript/api/excel/excel.cellpropertiesfillloadoptions#color)|指定是否在`color`属性上进行加载。|
||[pattern](/javascript/api/excel/excel.cellpropertiesfillloadoptions#pattern)|指定是否在`pattern`属性上进行加载。|
||[patternColor](/javascript/api/excel/excel.cellpropertiesfillloadoptions#patterncolor)|指定是否在`patternColor`属性上进行加载。|
||[patternTintAndShade](/javascript/api/excel/excel.cellpropertiesfillloadoptions#patterntintandshade)|指定是否在`patternTintAndShade`属性上进行加载。|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfillloadoptions#tintandshade)|指定是否在`tintAndShade`属性上进行加载。|
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
|[CellPropertiesFontLoadOptions](/javascript/api/excel/excel.cellpropertiesfontloadoptions)|[bold](/javascript/api/excel/excel.cellpropertiesfontloadoptions#bold)|指定是否在`bold`属性上进行加载。|
||[color](/javascript/api/excel/excel.cellpropertiesfontloadoptions#color)|指定是否在`color`属性上进行加载。|
||[italic](/javascript/api/excel/excel.cellpropertiesfontloadoptions#italic)|指定是否在`italic`属性上进行加载。|
||[name](/javascript/api/excel/excel.cellpropertiesfontloadoptions#name)|指定是否在`name`属性上进行加载。|
||[size](/javascript/api/excel/excel.cellpropertiesfontloadoptions#size)|指定是否在`size`属性上进行加载。|
||[strikethrough](/javascript/api/excel/excel.cellpropertiesfontloadoptions#strikethrough)|指定是否在`strikethrough`属性上进行加载。|
||[subscript](/javascript/api/excel/excel.cellpropertiesfontloadoptions#subscript)|指定是否在`subscript`属性上进行加载。|
||[superscript](/javascript/api/excel/excel.cellpropertiesfontloadoptions#superscript)|指定是否在`superscript`属性上进行加载。|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfontloadoptions#tintandshade)|指定是否在`tintAndShade`属性上进行加载。|
||[underline](/javascript/api/excel/excel.cellpropertiesfontloadoptions#underline)|指定是否在`underline`属性上进行加载。|
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
|[CellPropertiesFormatLoadOptions](/javascript/api/excel/excel.cellpropertiesformatloadoptions)|[autoIndent](/javascript/api/excel/excel.cellpropertiesformatloadoptions#autoindent)|指定是否在`autoIndent`属性上进行加载。|
||[Borders](/javascript/api/excel/excel.cellpropertiesformatloadoptions#borders)|指定是否在`borders`属性上进行加载。|
||[fill](/javascript/api/excel/excel.cellpropertiesformatloadoptions#fill)|指定是否在`fill`属性上进行加载。|
||[font](/javascript/api/excel/excel.cellpropertiesformatloadoptions#font)|指定是否在`font`属性上进行加载。|
||[horizontalAlignment](/javascript/api/excel/excel.cellpropertiesformatloadoptions#horizontalalignment)|指定是否在`horizontalAlignment`属性上进行加载。|
||[indentLevel](/javascript/api/excel/excel.cellpropertiesformatloadoptions#indentlevel)|指定是否在`indentLevel`属性上进行加载。|
||[protection](/javascript/api/excel/excel.cellpropertiesformatloadoptions#protection)|指定是否在`protection`属性上进行加载。|
||[readingOrder](/javascript/api/excel/excel.cellpropertiesformatloadoptions#readingorder)|指定是否在`readingOrder`属性上进行加载。|
||[shrinkToFit](/javascript/api/excel/excel.cellpropertiesformatloadoptions#shrinktofit)|指定是否在`shrinkToFit`属性上进行加载。|
||[textOrientation](/javascript/api/excel/excel.cellpropertiesformatloadoptions#textorientation)|指定是否在`textOrientation`属性上进行加载。|
||[useStandardHeight](/javascript/api/excel/excel.cellpropertiesformatloadoptions#usestandardheight)|指定是否在`useStandardHeight`属性上进行加载。|
||[useStandardWidth](/javascript/api/excel/excel.cellpropertiesformatloadoptions#usestandardwidth)|指定是否在`useStandardWidth`属性上进行加载。|
||[verticalAlignment](/javascript/api/excel/excel.cellpropertiesformatloadoptions#verticalalignment)|指定是否在`verticalAlignment`属性上进行加载。|
||[wrapText](/javascript/api/excel/excel.cellpropertiesformatloadoptions#wraptext)|指定是否在`wrapText`属性上进行加载。|
|[CellPropertiesLoadOptions](/javascript/api/excel/excel.cellpropertiesloadoptions)|[address](/javascript/api/excel/excel.cellpropertiesloadoptions#address)|指定是否在`address`属性上进行加载。|
||[addressLocal](/javascript/api/excel/excel.cellpropertiesloadoptions#addresslocal)|指定是否在`addressLocal`属性上进行加载。|
||[format](/javascript/api/excel/excel.cellpropertiesloadoptions#format)|指定是否在`format`属性上进行加载。|
||[hidden](/javascript/api/excel/excel.cellpropertiesloadoptions#hidden)|指定是否在`hidden`属性上进行加载。|
||[hyperlink](/javascript/api/excel/excel.cellpropertiesloadoptions#hyperlink)|指定是否在`hyperlink`属性上进行加载。|
||[style](/javascript/api/excel/excel.cellpropertiesloadoptions#style)|指定是否在`style`属性上进行加载。|
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
|[ChartAreaFormatData](/javascript/api/excel/excel.chartareaformatdata)|[colorScheme](/javascript/api/excel/excel.chartareaformatdata#colorscheme)|返回或设置图表的配色方案。 读/写。|
||[roundedCorners](/javascript/api/excel/excel.chartareaformatdata#roundedcorners)|指定图表的图表区域是否有圆角。 读/写。|
|[ChartAreaFormatLoadOptions](/javascript/api/excel/excel.chartareaformatloadoptions)|[colorScheme](/javascript/api/excel/excel.chartareaformatloadoptions#colorscheme)|返回或设置图表的配色方案。 读/写。|
||[roundedCorners](/javascript/api/excel/excel.chartareaformatloadoptions#roundedcorners)|指定图表的图表区域是否有圆角。 读/写。|
|[ChartAreaFormatUpdateData](/javascript/api/excel/excel.chartareaformatupdatedata)|[colorScheme](/javascript/api/excel/excel.chartareaformatupdatedata#colorscheme)|返回或设置图表的配色方案。 读/写。|
||[roundedCorners](/javascript/api/excel/excel.chartareaformatupdatedata#roundedcorners)|指定图表的图表区域是否有圆角。 读/写。|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[linkNumberFormat](/javascript/api/excel/excel.chartaxis#linknumberformat)|表示数字格式是否链接到单元格。 如果为 true，则数字格式会在单元格中更改时在标签中更改。|
|[ChartAxisData](/javascript/api/excel/excel.chartaxisdata)|[linkNumberFormat](/javascript/api/excel/excel.chartaxisdata#linknumberformat)|表示数字格式是否链接到单元格。 如果为 true，则数字格式会在单元格中更改时在标签中更改。|
|[ChartAxisLoadOptions](/javascript/api/excel/excel.chartaxisloadoptions)|[linkNumberFormat](/javascript/api/excel/excel.chartaxisloadoptions#linknumberformat)|表示数字格式是否链接到单元格。 如果为 true，则数字格式会在单元格中更改时在标签中更改。|
|[ChartAxisUpdateData](/javascript/api/excel/excel.chartaxisupdatedata)|[linkNumberFormat](/javascript/api/excel/excel.chartaxisupdatedata#linknumberformat)|表示数字格式是否链接到单元格。 如果为 true，则数字格式会在单元格中更改时在标签中更改。|
|[ChartBinOptions](/javascript/api/excel/excel.chartbinoptions)|[allowOverflow](/javascript/api/excel/excel.chartbinoptions#allowoverflow)|指定是否在直方图或排列图中启用容器溢出。 读/写。|
||[allowUnderflow](/javascript/api/excel/excel.chartbinoptions#allowunderflow)|指定是否在直方图或排列图中启用容器下溢。 读/写。|
||[count](/javascript/api/excel/excel.chartbinoptions#count)|返回或设置直方图或排列图的容器数量。 读/写。|
||[overflowValue](/javascript/api/excel/excel.chartbinoptions#overflowvalue)|返回或设置直方图或排列图的容器溢出值。 读/写。|
||[set (properties: ChartBinOptions)](/javascript/api/excel/excel.chartbinoptions#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: ChartBinOptionsUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.chartbinoptions#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
||[type](/javascript/api/excel/excel.chartbinoptions#type)|返回或设置直方图或排列图的容器类型。 读/写。|
||[underflowValue](/javascript/api/excel/excel.chartbinoptions#underflowvalue)|返回或设置直方图或排列图的容器下溢值。 读/写。|
||[width](/javascript/api/excel/excel.chartbinoptions#width)|返回或设置直方图或排列图的容器宽度值。 读/写。|
|[ChartBinOptionsData](/javascript/api/excel/excel.chartbinoptionsdata)|[allowOverflow](/javascript/api/excel/excel.chartbinoptionsdata#allowoverflow)|指定是否在直方图或排列图中启用容器溢出。 读/写。|
||[allowUnderflow](/javascript/api/excel/excel.chartbinoptionsdata#allowunderflow)|指定是否在直方图或排列图中启用容器下溢。 读/写。|
||[count](/javascript/api/excel/excel.chartbinoptionsdata#count)|返回或设置直方图或排列图的容器数量。 读/写。|
||[overflowValue](/javascript/api/excel/excel.chartbinoptionsdata#overflowvalue)|返回或设置直方图或排列图的容器溢出值。 读/写。|
||[type](/javascript/api/excel/excel.chartbinoptionsdata#type)|返回或设置直方图或排列图的容器类型。 读/写。|
||[underflowValue](/javascript/api/excel/excel.chartbinoptionsdata#underflowvalue)|返回或设置直方图或排列图的容器下溢值。 读/写。|
||[width](/javascript/api/excel/excel.chartbinoptionsdata#width)|返回或设置直方图或排列图的容器宽度值。 读/写。|
|[ChartBinOptionsLoadOptions](/javascript/api/excel/excel.chartbinoptionsloadoptions)|[$all](/javascript/api/excel/excel.chartbinoptionsloadoptions#$all)||
||[allowOverflow](/javascript/api/excel/excel.chartbinoptionsloadoptions#allowoverflow)|指定是否在直方图或排列图中启用容器溢出。 读/写。|
||[allowUnderflow](/javascript/api/excel/excel.chartbinoptionsloadoptions#allowunderflow)|指定是否在直方图或排列图中启用容器下溢。 读/写。|
||[count](/javascript/api/excel/excel.chartbinoptionsloadoptions#count)|返回或设置直方图或排列图的容器数量。 读/写。|
||[overflowValue](/javascript/api/excel/excel.chartbinoptionsloadoptions#overflowvalue)|返回或设置直方图或排列图的容器溢出值。 读/写。|
||[type](/javascript/api/excel/excel.chartbinoptionsloadoptions#type)|返回或设置直方图或排列图的容器类型。 读/写。|
||[underflowValue](/javascript/api/excel/excel.chartbinoptionsloadoptions#underflowvalue)|返回或设置直方图或排列图的容器下溢值。 读/写。|
||[width](/javascript/api/excel/excel.chartbinoptionsloadoptions#width)|返回或设置直方图或排列图的容器宽度值。 读/写。|
|[ChartBinOptionsUpdateData](/javascript/api/excel/excel.chartbinoptionsupdatedata)|[allowOverflow](/javascript/api/excel/excel.chartbinoptionsupdatedata#allowoverflow)|指定是否在直方图或排列图中启用容器溢出。 读/写。|
||[allowUnderflow](/javascript/api/excel/excel.chartbinoptionsupdatedata#allowunderflow)|指定是否在直方图或排列图中启用容器下溢。 读/写。|
||[count](/javascript/api/excel/excel.chartbinoptionsupdatedata#count)|返回或设置直方图或排列图的容器数量。 读/写。|
||[overflowValue](/javascript/api/excel/excel.chartbinoptionsupdatedata#overflowvalue)|返回或设置直方图或排列图的容器溢出值。 读/写。|
||[type](/javascript/api/excel/excel.chartbinoptionsupdatedata#type)|返回或设置直方图或排列图的容器类型。 读/写。|
||[underflowValue](/javascript/api/excel/excel.chartbinoptionsupdatedata#underflowvalue)|返回或设置直方图或排列图的容器下溢值。 读/写。|
||[width](/javascript/api/excel/excel.chartbinoptionsupdatedata#width)|返回或设置直方图或排列图的容器宽度值。 读/写。|
|[ChartBoxwhiskerOptions](/javascript/api/excel/excel.chartboxwhiskeroptions)|[quartileCalculation](/javascript/api/excel/excel.chartboxwhiskeroptions#quartilecalculation)|返回或设置箱形图的四分位点计算类型。 读/写。|
||[set (properties: ChartBoxwhiskerOptions)](/javascript/api/excel/excel.chartboxwhiskeroptions#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: ChartBoxwhiskerOptionsUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.chartboxwhiskeroptions#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
||[showInnerPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#showinnerpoints)|指定在箱形图中是否显示内部点。 读/写。|
||[showMeanLine](/javascript/api/excel/excel.chartboxwhiskeroptions#showmeanline)|指定在箱形图中是否显示中线。 读/写。|
||[showMeanMarker](/javascript/api/excel/excel.chartboxwhiskeroptions#showmeanmarker)|指定在箱形图中是否显示平均值标记。 读/写。|
||[showOutlierPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#showoutlierpoints)|指定在箱形图中是否显示离群值点。 读/写。|
|[ChartBoxwhiskerOptionsData](/javascript/api/excel/excel.chartboxwhiskeroptionsdata)|[quartileCalculation](/javascript/api/excel/excel.chartboxwhiskeroptionsdata#quartilecalculation)|返回或设置箱形图的四分位点计算类型。 读/写。|
||[showInnerPoints](/javascript/api/excel/excel.chartboxwhiskeroptionsdata#showinnerpoints)|指定在箱形图中是否显示内部点。 读/写。|
||[showMeanLine](/javascript/api/excel/excel.chartboxwhiskeroptionsdata#showmeanline)|指定在箱形图中是否显示中线。 读/写。|
||[showMeanMarker](/javascript/api/excel/excel.chartboxwhiskeroptionsdata#showmeanmarker)|指定在箱形图中是否显示平均值标记。 读/写。|
||[showOutlierPoints](/javascript/api/excel/excel.chartboxwhiskeroptionsdata#showoutlierpoints)|指定在箱形图中是否显示离群值点。 读/写。|
|[ChartBoxwhiskerOptionsLoadOptions](/javascript/api/excel/excel.chartboxwhiskeroptionsloadoptions)|[$all](/javascript/api/excel/excel.chartboxwhiskeroptionsloadoptions#$all)||
||[quartileCalculation](/javascript/api/excel/excel.chartboxwhiskeroptionsloadoptions#quartilecalculation)|返回或设置箱形图的四分位点计算类型。 读/写。|
||[showInnerPoints](/javascript/api/excel/excel.chartboxwhiskeroptionsloadoptions#showinnerpoints)|指定在箱形图中是否显示内部点。 读/写。|
||[showMeanLine](/javascript/api/excel/excel.chartboxwhiskeroptionsloadoptions#showmeanline)|指定在箱形图中是否显示中线。 读/写。|
||[showMeanMarker](/javascript/api/excel/excel.chartboxwhiskeroptionsloadoptions#showmeanmarker)|指定在箱形图中是否显示平均值标记。 读/写。|
||[showOutlierPoints](/javascript/api/excel/excel.chartboxwhiskeroptionsloadoptions#showoutlierpoints)|指定在箱形图中是否显示离群值点。 读/写。|
|[ChartBoxwhiskerOptionsUpdateData](/javascript/api/excel/excel.chartboxwhiskeroptionsupdatedata)|[quartileCalculation](/javascript/api/excel/excel.chartboxwhiskeroptionsupdatedata#quartilecalculation)|返回或设置箱形图的四分位点计算类型。 读/写。|
||[showInnerPoints](/javascript/api/excel/excel.chartboxwhiskeroptionsupdatedata#showinnerpoints)|指定在箱形图中是否显示内部点。 读/写。|
||[showMeanLine](/javascript/api/excel/excel.chartboxwhiskeroptionsupdatedata#showmeanline)|指定在箱形图中是否显示中线。 读/写。|
||[showMeanMarker](/javascript/api/excel/excel.chartboxwhiskeroptionsupdatedata#showmeanmarker)|指定在箱形图中是否显示平均值标记。 读/写。|
||[showOutlierPoints](/javascript/api/excel/excel.chartboxwhiskeroptionsupdatedata#showoutlierpoints)|指定在箱形图中是否显示离群值点。 读/写。|
|[ChartCollectionLoadOptions](/javascript/api/excel/excel.chartcollectionloadoptions)|[pivotOptions](/javascript/api/excel/excel.chartcollectionloadoptions#pivotoptions)|对于集合中的每一项: 封装数据透视图表的选项。|
|[ChartData](/javascript/api/excel/excel.chartdata)|[pivotOptions](/javascript/api/excel/excel.chartdata#pivotoptions)|封装数据透视图的选项。 只读。|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabel#linknumberformat)|布尔值，表示数字格式是否链接到单元格（以便在单元格中更改时标签中的数字格式会发生改变）。|
|[ChartDataLabelData](/javascript/api/excel/excel.chartdatalabeldata)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabeldata#linknumberformat)|布尔值，表示数字格式是否链接到单元格（以便在单元格中更改时标签中的数字格式会发生改变）。|
|[ChartDataLabelLoadOptions](/javascript/api/excel/excel.chartdatalabelloadoptions)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabelloadoptions#linknumberformat)|布尔值，表示数字格式是否链接到单元格（以便在单元格中更改时标签中的数字格式会发生改变）。|
|[ChartDataLabelUpdateData](/javascript/api/excel/excel.chartdatalabelupdatedata)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabelupdatedata#linknumberformat)|布尔值，表示数字格式是否链接到单元格（以便在单元格中更改时标签中的数字格式会发生改变）。|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabels#linknumberformat)|表示数字格式是否链接到单元格。 如果为 true，则数字格式会在单元格中更改时在标签中更改。|
|[ChartDataLabelsData](/javascript/api/excel/excel.chartdatalabelsdata)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabelsdata#linknumberformat)|表示数字格式是否链接到单元格。 如果为 true，则数字格式会在单元格中更改时在标签中更改。|
|[ChartDataLabelsLoadOptions](/javascript/api/excel/excel.chartdatalabelsloadoptions)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabelsloadoptions#linknumberformat)|表示数字格式是否链接到单元格。 如果为 true，则数字格式会在单元格中更改时在标签中更改。|
|[ChartDataLabelsUpdateData](/javascript/api/excel/excel.chartdatalabelsupdatedata)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabelsupdatedata#linknumberformat)|表示数字格式是否链接到单元格。 如果为 true，则数字格式会在单元格中更改时在标签中更改。|
|[ChartErrorBars](/javascript/api/excel/excel.charterrorbars)|[endStyleCap](/javascript/api/excel/excel.charterrorbars#endstylecap)|指定误差线是否具有终止端样式。|
||[include](/javascript/api/excel/excel.charterrorbars#include)|指定包含误差线的哪些部分。|
||[format](/javascript/api/excel/excel.charterrorbars#format)|指定误差线的格式类型。|
||[set (properties: ChartErrorBars)](/javascript/api/excel/excel.charterrorbars#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: ChartErrorBarsUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.charterrorbars#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
||[type](/javascript/api/excel/excel.charterrorbars#type)|误差线标记的区域类型。|
||[visible](/javascript/api/excel/excel.charterrorbars#visible)|指定是否显示误差线。|
|[ChartErrorBarsData](/javascript/api/excel/excel.charterrorbarsdata)|[endStyleCap](/javascript/api/excel/excel.charterrorbarsdata#endstylecap)|指定误差线是否具有终止端样式。|
||[format](/javascript/api/excel/excel.charterrorbarsdata#format)|指定误差线的格式类型。|
||[include](/javascript/api/excel/excel.charterrorbarsdata#include)|指定包含误差线的哪些部分。|
||[type](/javascript/api/excel/excel.charterrorbarsdata#type)|误差线标记的区域类型。|
||[visible](/javascript/api/excel/excel.charterrorbarsdata#visible)|指定是否显示误差线。|
|[ChartErrorBarsFormat](/javascript/api/excel/excel.charterrorbarsformat)|[line](/javascript/api/excel/excel.charterrorbarsformat#line)|表示图表线条格式。|
||[set (properties: ChartErrorBarsFormat)](/javascript/api/excel/excel.charterrorbarsformat#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: ChartErrorBarsFormatUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.charterrorbarsformat#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
|[ChartErrorBarsFormatData](/javascript/api/excel/excel.charterrorbarsformatdata)|[line](/javascript/api/excel/excel.charterrorbarsformatdata#line)|表示图表线条格式。|
|[ChartErrorBarsFormatLoadOptions](/javascript/api/excel/excel.charterrorbarsformatloadoptions)|[$all](/javascript/api/excel/excel.charterrorbarsformatloadoptions#$all)||
||[line](/javascript/api/excel/excel.charterrorbarsformatloadoptions#line)|表示图表线条格式。|
|[ChartErrorBarsFormatUpdateData](/javascript/api/excel/excel.charterrorbarsformatupdatedata)|[line](/javascript/api/excel/excel.charterrorbarsformatupdatedata#line)|表示图表线条格式。|
|[ChartErrorBarsLoadOptions](/javascript/api/excel/excel.charterrorbarsloadoptions)|[$all](/javascript/api/excel/excel.charterrorbarsloadoptions#$all)||
||[endStyleCap](/javascript/api/excel/excel.charterrorbarsloadoptions#endstylecap)|指定误差线是否具有终止端样式。|
||[format](/javascript/api/excel/excel.charterrorbarsloadoptions#format)|指定误差线的格式类型。|
||[include](/javascript/api/excel/excel.charterrorbarsloadoptions#include)|指定包含误差线的哪些部分。|
||[type](/javascript/api/excel/excel.charterrorbarsloadoptions#type)|误差线标记的区域类型。|
||[visible](/javascript/api/excel/excel.charterrorbarsloadoptions#visible)|指定是否显示误差线。|
|[ChartErrorBarsUpdateData](/javascript/api/excel/excel.charterrorbarsupdatedata)|[endStyleCap](/javascript/api/excel/excel.charterrorbarsupdatedata#endstylecap)|指定误差线是否具有终止端样式。|
||[format](/javascript/api/excel/excel.charterrorbarsupdatedata#format)|指定误差线的格式类型。|
||[include](/javascript/api/excel/excel.charterrorbarsupdatedata#include)|指定包含误差线的哪些部分。|
||[type](/javascript/api/excel/excel.charterrorbarsupdatedata#type)|误差线标记的区域类型。|
||[visible](/javascript/api/excel/excel.charterrorbarsupdatedata#visible)|指定是否显示误差线。|
|[ChartLoadOptions](/javascript/api/excel/excel.chartloadoptions)|[pivotOptions](/javascript/api/excel/excel.chartloadoptions#pivotoptions)|封装数据透视图的选项。|
|[ChartMapOptions](/javascript/api/excel/excel.chartmapoptions)|[labelStrategy](/javascript/api/excel/excel.chartmapoptions#labelstrategy)|返回或设置区域地图图表的系列地图标签策略。 读/写。|
||[level](/javascript/api/excel/excel.chartmapoptions#level)|返回或设置区域地图图表的系列映射级别。 读/写。|
||[projectionType](/javascript/api/excel/excel.chartmapoptions#projectiontype)|返回或设置区域地图图表的系列投影类型。 读/写。|
||[set (properties: ChartMapOptions)](/javascript/api/excel/excel.chartmapoptions#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: ChartMapOptionsUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.chartmapoptions#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
|[ChartMapOptionsData](/javascript/api/excel/excel.chartmapoptionsdata)|[labelStrategy](/javascript/api/excel/excel.chartmapoptionsdata#labelstrategy)|返回或设置区域地图图表的系列地图标签策略。 读/写。|
||[level](/javascript/api/excel/excel.chartmapoptionsdata#level)|返回或设置区域地图图表的系列映射级别。 读/写。|
||[projectionType](/javascript/api/excel/excel.chartmapoptionsdata#projectiontype)|返回或设置区域地图图表的系列投影类型。 读/写。|
|[ChartMapOptionsLoadOptions](/javascript/api/excel/excel.chartmapoptionsloadoptions)|[$all](/javascript/api/excel/excel.chartmapoptionsloadoptions#$all)||
||[labelStrategy](/javascript/api/excel/excel.chartmapoptionsloadoptions#labelstrategy)|返回或设置区域地图图表的系列地图标签策略。 读/写。|
||[level](/javascript/api/excel/excel.chartmapoptionsloadoptions#level)|返回或设置区域地图图表的系列映射级别。 读/写。|
||[projectionType](/javascript/api/excel/excel.chartmapoptionsloadoptions#projectiontype)|返回或设置区域地图图表的系列投影类型。 读/写。|
|[ChartMapOptionsUpdateData](/javascript/api/excel/excel.chartmapoptionsupdatedata)|[labelStrategy](/javascript/api/excel/excel.chartmapoptionsupdatedata#labelstrategy)|返回或设置区域地图图表的系列地图标签策略。 读/写。|
||[level](/javascript/api/excel/excel.chartmapoptionsupdatedata#level)|返回或设置区域地图图表的系列映射级别。 读/写。|
||[projectionType](/javascript/api/excel/excel.chartmapoptionsupdatedata#projectiontype)|返回或设置区域地图图表的系列投影类型。 读/写。|
|[ChartPivotOptions](/javascript/api/excel/excel.chartpivotoptions)|[set (properties: ChartPivotOptions)](/javascript/api/excel/excel.chartpivotoptions#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: ChartPivotOptionsUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.chartpivotoptions#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
||[showAxisFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showaxisfieldbuttons)|指定是否在数据透视图上显示轴字段按钮。 ShowAxisFieldButtons 属性对应于“分析”选项卡（在选择数据透视图时可用）的“字段按钮”下拉列表中的“显示坐标轴字段按钮”命令。|
||[showLegendFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showlegendfieldbuttons)|指定是否在数据透视图上显示图例字段按钮。|
||[showReportFilterFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showreportfilterfieldbuttons)|指定是否在数据透视图上显示报表筛选字段按钮。|
||[showValueFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showvaluefieldbuttons)|指定是否在数据透视图上显示值字段按钮。|
|[ChartPivotOptionsData](/javascript/api/excel/excel.chartpivotoptionsdata)|[showAxisFieldButtons](/javascript/api/excel/excel.chartpivotoptionsdata#showaxisfieldbuttons)|指定是否在数据透视图上显示轴字段按钮。 ShowAxisFieldButtons 属性对应于“分析”选项卡（在选择数据透视图时可用）的“字段按钮”下拉列表中的“显示坐标轴字段按钮”命令。|
||[showLegendFieldButtons](/javascript/api/excel/excel.chartpivotoptionsdata#showlegendfieldbuttons)|指定是否在数据透视图上显示图例字段按钮。|
||[showReportFilterFieldButtons](/javascript/api/excel/excel.chartpivotoptionsdata#showreportfilterfieldbuttons)|指定是否在数据透视图上显示报表筛选字段按钮。|
||[showValueFieldButtons](/javascript/api/excel/excel.chartpivotoptionsdata#showvaluefieldbuttons)|指定是否在数据透视图上显示值字段按钮。|
|[ChartPivotOptionsLoadOptions](/javascript/api/excel/excel.chartpivotoptionsloadoptions)|[$all](/javascript/api/excel/excel.chartpivotoptionsloadoptions#$all)||
||[showAxisFieldButtons](/javascript/api/excel/excel.chartpivotoptionsloadoptions#showaxisfieldbuttons)|指定是否在数据透视图上显示轴字段按钮。 ShowAxisFieldButtons 属性对应于“分析”选项卡（在选择数据透视图时可用）的“字段按钮”下拉列表中的“显示坐标轴字段按钮”命令。|
||[showLegendFieldButtons](/javascript/api/excel/excel.chartpivotoptionsloadoptions#showlegendfieldbuttons)|指定是否在数据透视图上显示图例字段按钮。|
||[showReportFilterFieldButtons](/javascript/api/excel/excel.chartpivotoptionsloadoptions#showreportfilterfieldbuttons)|指定是否在数据透视图上显示报表筛选字段按钮。|
||[showValueFieldButtons](/javascript/api/excel/excel.chartpivotoptionsloadoptions#showvaluefieldbuttons)|指定是否在数据透视图上显示值字段按钮。|
|[ChartPivotOptionsUpdateData](/javascript/api/excel/excel.chartpivotoptionsupdatedata)|[showAxisFieldButtons](/javascript/api/excel/excel.chartpivotoptionsupdatedata#showaxisfieldbuttons)|指定是否在数据透视图上显示轴字段按钮。 ShowAxisFieldButtons 属性对应于“分析”选项卡（在选择数据透视图时可用）的“字段按钮”下拉列表中的“显示坐标轴字段按钮”命令。|
||[showLegendFieldButtons](/javascript/api/excel/excel.chartpivotoptionsupdatedata#showlegendfieldbuttons)|指定是否在数据透视图上显示图例字段按钮。|
||[showReportFilterFieldButtons](/javascript/api/excel/excel.chartpivotoptionsupdatedata#showreportfilterfieldbuttons)|指定是否在数据透视图上显示报表筛选字段按钮。|
||[showValueFieldButtons](/javascript/api/excel/excel.chartpivotoptionsupdatedata#showvaluefieldbuttons)|指定是否在数据透视图上显示值字段按钮。|
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
|[ChartSeriesCollectionLoadOptions](/javascript/api/excel/excel.chartseriescollectionloadoptions)|[binOptions](/javascript/api/excel/excel.chartseriescollectionloadoptions#binoptions)|对于集合中的每一项: 封装直方图图表和排列图表的 bin 选项。|
||[boxwhiskerOptions](/javascript/api/excel/excel.chartseriescollectionloadoptions#boxwhiskeroptions)|对于集合中的每一项: 封装 box 和流程图的选项。|
||[bubbleScale](/javascript/api/excel/excel.chartseriescollectionloadoptions#bubblescale)|对于集合中的每一项: 此值可以是从 0 (零) 到300的整数值, 表示默认大小的百分比。 该属性仅适用于气泡图。 读/写。|
||[gradientMaximumColor](/javascript/api/excel/excel.chartseriescollectionloadoptions#gradientmaximumcolor)|对于集合中的每个项目: 返回或设置区域地图图表系列的最大值的颜色。 读/写。|
||[gradientMaximumType](/javascript/api/excel/excel.chartseriescollectionloadoptions#gradientmaximumtype)|对于集合中的每个项目: 返回或设置区域地图图表系列的最大值的类型。 读/写。|
||[gradientMaximumValue](/javascript/api/excel/excel.chartseriescollectionloadoptions#gradientmaximumvalue)|对于集合中的每个项目: 返回或设置区域地图图表系列的最大值。 读/写。|
||[gradientMidpointColor](/javascript/api/excel/excel.chartseriescollectionloadoptions#gradientmidpointcolor)|对于集合中的每个项目: 返回或设置区域地图图表系列的中点值的颜色。 读/写。|
||[gradientMidpointType](/javascript/api/excel/excel.chartseriescollectionloadoptions#gradientmidpointtype)|对于集合中的每个项目: 返回或设置区域地图图表系列的中点值的类型。 读/写。|
||[gradientMidpointValue](/javascript/api/excel/excel.chartseriescollectionloadoptions#gradientmidpointvalue)|对于集合中的每个项目: 返回或设置区域地图图表系列的中点值。 读/写。|
||[gradientMinimumColor](/javascript/api/excel/excel.chartseriescollectionloadoptions#gradientminimumcolor)|对于集合中的每个项目: 返回或设置区域地图图表系列的最小值的颜色。 读/写。|
||[gradientMinimumType](/javascript/api/excel/excel.chartseriescollectionloadoptions#gradientminimumtype)|对于集合中的每个项目: 返回或设置区域地图图表系列的最小值的类型。 读/写。|
||[gradientMinimumValue](/javascript/api/excel/excel.chartseriescollectionloadoptions#gradientminimumvalue)|对于集合中的每个项目: 返回或设置区域地图图表系列的最小值。 读/写。|
||[gradientStyle](/javascript/api/excel/excel.chartseriescollectionloadoptions#gradientstyle)|对于集合中的每个项目: 返回或设置区域地图图表的系列渐变样式。 读/写。|
||[invertColor](/javascript/api/excel/excel.chartseriescollectionloadoptions#invertcolor)|对于集合中的每一项: 返回或设置系列中的负数据点的填充颜色。 读/写。|
||[mapOptions](/javascript/api/excel/excel.chartseriescollectionloadoptions#mapoptions)|对于集合中的每一项: 封装区域映射图的选项。|
||[parentLabelStrategy](/javascript/api/excel/excel.chartseriescollectionloadoptions#parentlabelstrategy)|对于集合中的每一项: 返回或设置树状图图表的系列父标签策略区域。 读/写。|
||[showConnectorLines](/javascript/api/excel/excel.chartseriescollectionloadoptions#showconnectorlines)|对于集合中的每一项: 指定在瀑布图中是否显示连接符线。 读/写。|
||[showLeaderLines](/javascript/api/excel/excel.chartseriescollectionloadoptions#showleaderlines)|对于集合中的每一项: 指定是否为系列中的每个数据标签显示引导线。 读/写。|
||[splitValue](/javascript/api/excel/excel.chartseriescollectionloadoptions#splitvalue)|对于集合中的每一项: 返回或设置用于分隔复合饼图或复合条饼图的两个部分的阈值。 读/写。|
||[xErrorBars](/javascript/api/excel/excel.chartseriescollectionloadoptions#xerrorbars)|对于集合中的每一项: 代表图表系列的误差条形图对象。|
||[yErrorBars](/javascript/api/excel/excel.chartseriescollectionloadoptions#yerrorbars)|对于集合中的每一项: 代表图表系列的误差条形图对象。|
|[ChartSeriesData](/javascript/api/excel/excel.chartseriesdata)|[binOptions](/javascript/api/excel/excel.chartseriesdata#binoptions)|封装直方图和排列图的容器选项。 只读。|
||[boxwhiskerOptions](/javascript/api/excel/excel.chartseriesdata#boxwhiskeroptions)|封装箱形图的选项。 只读。|
||[bubbleScale](/javascript/api/excel/excel.chartseriesdata#bubblescale)|这可以是从 0（零）到 300 的整数值，表示默认大小的百分比。 该属性仅适用于气泡图。 读/写。|
||[gradientMaximumColor](/javascript/api/excel/excel.chartseriesdata#gradientmaximumcolor)|返回或设置区域地图图表系列的最大值的颜色。 读/写。|
||[gradientMaximumType](/javascript/api/excel/excel.chartseriesdata#gradientmaximumtype)|返回或设置区域地图图表系列的最大值的类型。 读/写。|
||[gradientMaximumValue](/javascript/api/excel/excel.chartseriesdata#gradientmaximumvalue)|返回或设置区域地图图表系列的最大值。 读/写。|
||[gradientMidpointColor](/javascript/api/excel/excel.chartseriesdata#gradientmidpointcolor)|返回或设置区域地图图表系列的中间值的颜色。 读/写。|
||[gradientMidpointType](/javascript/api/excel/excel.chartseriesdata#gradientmidpointtype)|返回或设置区域地图图表系列的中间值的类型。 读/写。|
||[gradientMidpointValue](/javascript/api/excel/excel.chartseriesdata#gradientmidpointvalue)|返回或设置区域地图图表系列的中间值。 读/写。|
||[gradientMinimumColor](/javascript/api/excel/excel.chartseriesdata#gradientminimumcolor)|返回或设置区域地图图表系列的最小值的颜色。 读/写。|
||[gradientMinimumType](/javascript/api/excel/excel.chartseriesdata#gradientminimumtype)|返回或设置区域地图图表系列的最小值的类型。 读/写。|
||[gradientMinimumValue](/javascript/api/excel/excel.chartseriesdata#gradientminimumvalue)|返回或设置区域地图图表系列的最小值。 读/写。|
||[gradientStyle](/javascript/api/excel/excel.chartseriesdata#gradientstyle)|返回或设置区域地图图表的系列渐变样式。 读/写。|
||[invertColor](/javascript/api/excel/excel.chartseriesdata#invertcolor)|返回或设置系列中负数据点的填充颜色。 读/写。|
||[mapOptions](/javascript/api/excel/excel.chartseriesdata#mapoptions)|封装区域地图图表的选项。 只读。|
||[parentLabelStrategy](/javascript/api/excel/excel.chartseriesdata#parentlabelstrategy)|返回或设置树状图的系列父标签策略区域。 读/写。|
||[showConnectorLines](/javascript/api/excel/excel.chartseriesdata#showconnectorlines)|指定是否在瀑布图中显示连接线。 读/写。|
||[showLeaderLines](/javascript/api/excel/excel.chartseriesdata#showleaderlines)|指定是否在系列中显示每个数据标签的引导线。 读/写。|
||[splitValue](/javascript/api/excel/excel.chartseriesdata#splitvalue)|返回或设置复合饼图或复合条饼图中分隔两部分的阈值。 读/写。|
||[xErrorBars](/javascript/api/excel/excel.chartseriesdata#xerrorbars)|表示图表系列的误差线对象。|
||[yErrorBars](/javascript/api/excel/excel.chartseriesdata#yerrorbars)|表示图表系列的误差线对象。|
|[ChartSeriesLoadOptions](/javascript/api/excel/excel.chartseriesloadoptions)|[binOptions](/javascript/api/excel/excel.chartseriesloadoptions#binoptions)|封装直方图和排列图的容器选项。|
||[boxwhiskerOptions](/javascript/api/excel/excel.chartseriesloadoptions#boxwhiskeroptions)|封装箱形图的选项。|
||[bubbleScale](/javascript/api/excel/excel.chartseriesloadoptions#bubblescale)|这可以是从 0（零）到 300 的整数值，表示默认大小的百分比。 该属性仅适用于气泡图。 读/写。|
||[gradientMaximumColor](/javascript/api/excel/excel.chartseriesloadoptions#gradientmaximumcolor)|返回或设置区域地图图表系列的最大值的颜色。 读/写。|
||[gradientMaximumType](/javascript/api/excel/excel.chartseriesloadoptions#gradientmaximumtype)|返回或设置区域地图图表系列的最大值的类型。 读/写。|
||[gradientMaximumValue](/javascript/api/excel/excel.chartseriesloadoptions#gradientmaximumvalue)|返回或设置区域地图图表系列的最大值。 读/写。|
||[gradientMidpointColor](/javascript/api/excel/excel.chartseriesloadoptions#gradientmidpointcolor)|返回或设置区域地图图表系列的中间值的颜色。 读/写。|
||[gradientMidpointType](/javascript/api/excel/excel.chartseriesloadoptions#gradientmidpointtype)|返回或设置区域地图图表系列的中间值的类型。 读/写。|
||[gradientMidpointValue](/javascript/api/excel/excel.chartseriesloadoptions#gradientmidpointvalue)|返回或设置区域地图图表系列的中间值。 读/写。|
||[gradientMinimumColor](/javascript/api/excel/excel.chartseriesloadoptions#gradientminimumcolor)|返回或设置区域地图图表系列的最小值的颜色。 读/写。|
||[gradientMinimumType](/javascript/api/excel/excel.chartseriesloadoptions#gradientminimumtype)|返回或设置区域地图图表系列的最小值的类型。 读/写。|
||[gradientMinimumValue](/javascript/api/excel/excel.chartseriesloadoptions#gradientminimumvalue)|返回或设置区域地图图表系列的最小值。 读/写。|
||[gradientStyle](/javascript/api/excel/excel.chartseriesloadoptions#gradientstyle)|返回或设置区域地图图表的系列渐变样式。 读/写。|
||[invertColor](/javascript/api/excel/excel.chartseriesloadoptions#invertcolor)|返回或设置系列中负数据点的填充颜色。 读/写。|
||[mapOptions](/javascript/api/excel/excel.chartseriesloadoptions#mapoptions)|封装区域地图图表的选项。|
||[parentLabelStrategy](/javascript/api/excel/excel.chartseriesloadoptions#parentlabelstrategy)|返回或设置树状图的系列父标签策略区域。 读/写。|
||[showConnectorLines](/javascript/api/excel/excel.chartseriesloadoptions#showconnectorlines)|指定是否在瀑布图中显示连接线。 读/写。|
||[showLeaderLines](/javascript/api/excel/excel.chartseriesloadoptions#showleaderlines)|指定是否在系列中显示每个数据标签的引导线。 读/写。|
||[splitValue](/javascript/api/excel/excel.chartseriesloadoptions#splitvalue)|返回或设置复合饼图或复合条饼图中分隔两部分的阈值。 读/写。|
||[xErrorBars](/javascript/api/excel/excel.chartseriesloadoptions#xerrorbars)|表示图表系列的误差线对象。|
||[yErrorBars](/javascript/api/excel/excel.chartseriesloadoptions#yerrorbars)|表示图表系列的误差线对象。|
|[ChartSeriesUpdateData](/javascript/api/excel/excel.chartseriesupdatedata)|[binOptions](/javascript/api/excel/excel.chartseriesupdatedata#binoptions)|封装直方图和排列图的容器选项。|
||[boxwhiskerOptions](/javascript/api/excel/excel.chartseriesupdatedata#boxwhiskeroptions)|封装箱形图的选项。|
||[bubbleScale](/javascript/api/excel/excel.chartseriesupdatedata#bubblescale)|这可以是从 0（零）到 300 的整数值，表示默认大小的百分比。 该属性仅适用于气泡图。 读/写。|
||[gradientMaximumColor](/javascript/api/excel/excel.chartseriesupdatedata#gradientmaximumcolor)|返回或设置区域地图图表系列的最大值的颜色。 读/写。|
||[gradientMaximumType](/javascript/api/excel/excel.chartseriesupdatedata#gradientmaximumtype)|返回或设置区域地图图表系列的最大值的类型。 读/写。|
||[gradientMaximumValue](/javascript/api/excel/excel.chartseriesupdatedata#gradientmaximumvalue)|返回或设置区域地图图表系列的最大值。 读/写。|
||[gradientMidpointColor](/javascript/api/excel/excel.chartseriesupdatedata#gradientmidpointcolor)|返回或设置区域地图图表系列的中间值的颜色。 读/写。|
||[gradientMidpointType](/javascript/api/excel/excel.chartseriesupdatedata#gradientmidpointtype)|返回或设置区域地图图表系列的中间值的类型。 读/写。|
||[gradientMidpointValue](/javascript/api/excel/excel.chartseriesupdatedata#gradientmidpointvalue)|返回或设置区域地图图表系列的中间值。 读/写。|
||[gradientMinimumColor](/javascript/api/excel/excel.chartseriesupdatedata#gradientminimumcolor)|返回或设置区域地图图表系列的最小值的颜色。 读/写。|
||[gradientMinimumType](/javascript/api/excel/excel.chartseriesupdatedata#gradientminimumtype)|返回或设置区域地图图表系列的最小值的类型。 读/写。|
||[gradientMinimumValue](/javascript/api/excel/excel.chartseriesupdatedata#gradientminimumvalue)|返回或设置区域地图图表系列的最小值。 读/写。|
||[gradientStyle](/javascript/api/excel/excel.chartseriesupdatedata#gradientstyle)|返回或设置区域地图图表的系列渐变样式。 读/写。|
||[invertColor](/javascript/api/excel/excel.chartseriesupdatedata#invertcolor)|返回或设置系列中负数据点的填充颜色。 读/写。|
||[mapOptions](/javascript/api/excel/excel.chartseriesupdatedata#mapoptions)|封装区域地图图表的选项。|
||[parentLabelStrategy](/javascript/api/excel/excel.chartseriesupdatedata#parentlabelstrategy)|返回或设置树状图的系列父标签策略区域。 读/写。|
||[showConnectorLines](/javascript/api/excel/excel.chartseriesupdatedata#showconnectorlines)|指定是否在瀑布图中显示连接线。 读/写。|
||[showLeaderLines](/javascript/api/excel/excel.chartseriesupdatedata#showleaderlines)|指定是否在系列中显示每个数据标签的引导线。 读/写。|
||[splitValue](/javascript/api/excel/excel.chartseriesupdatedata#splitvalue)|返回或设置复合饼图或复合条饼图中分隔两部分的阈值。 读/写。|
||[xErrorBars](/javascript/api/excel/excel.chartseriesupdatedata#xerrorbars)|表示图表系列的误差线对象。|
||[yErrorBars](/javascript/api/excel/excel.chartseriesupdatedata#yerrorbars)|表示图表系列的误差线对象。|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[linkNumberFormat](/javascript/api/excel/excel.charttrendlinelabel#linknumberformat)|布尔值，表示数字格式是否链接到单元格（以便在单元格中更改时标签中的数字格式会发生改变）。|
|[ChartTrendlineLabelData](/javascript/api/excel/excel.charttrendlinelabeldata)|[linkNumberFormat](/javascript/api/excel/excel.charttrendlinelabeldata#linknumberformat)|布尔值，表示数字格式是否链接到单元格（以便在单元格中更改时标签中的数字格式会发生改变）。|
|[ChartTrendlineLabelLoadOptions](/javascript/api/excel/excel.charttrendlinelabelloadoptions)|[linkNumberFormat](/javascript/api/excel/excel.charttrendlinelabelloadoptions#linknumberformat)|布尔值，表示数字格式是否链接到单元格（以便在单元格中更改时标签中的数字格式会发生改变）。|
|[ChartTrendlineLabelUpdateData](/javascript/api/excel/excel.charttrendlinelabelupdatedata)|[linkNumberFormat](/javascript/api/excel/excel.charttrendlinelabelupdatedata#linknumberformat)|布尔值，表示数字格式是否链接到单元格（以便在单元格中更改时标签中的数字格式会发生改变）。|
|[ChartUpdateData](/javascript/api/excel/excel.chartupdatedata)|[pivotOptions](/javascript/api/excel/excel.chartupdatedata#pivotoptions)|封装数据透视图的选项。|
|[ColumnProperties](/javascript/api/excel/excel.columnproperties)|[address](/javascript/api/excel/excel.columnproperties#address)|表示`address`属性。|
||[addressLocal](/javascript/api/excel/excel.columnproperties#addresslocal)|表示`addressLocal`属性。|
||[columnIndex](/javascript/api/excel/excel.columnproperties#columnindex)|表示`columnIndex`属性。|
|[ColumnPropertiesLoadOptions](/javascript/api/excel/excel.columnpropertiesloadoptions)|[columnHidden](/javascript/api/excel/excel.columnpropertiesloadoptions#columnhidden)|指定是否在`columnHidden`属性上进行加载。|
||[columnIndex](/javascript/api/excel/excel.columnpropertiesloadoptions#columnindex)|指定是否在`columnIndex`属性上进行加载。|
||[columnWidth](/javascript/api/excel/excel.columnpropertiesloadoptions#columnwidth)||
||[格式: CellPropertiesFormatLoadOptions & {
            columnWidth？](/javascript/api/excel/excel.columnpropertiesloadoptions # format)|指定是否在`format`属性上进行加载。|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[getRanges()](/javascript/api/excel/excel.conditionalformat#getranges--)|返回将为其应用条件格式的 RangeAreas，它包含一个或多个矩形区域。 只读。|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[getInvalidCells()](/javascript/api/excel/excel.datavalidation#getinvalidcells--)|返回包含一个或多个矩形区域的 RangeAreas，它具有无效单元格值。 如果所有单元格值都有效，则此函数将引发 ItemNotFound 错误。|
||[getInvalidCellsOrNullObject()](/javascript/api/excel/excel.datavalidation#getinvalidcellsornullobject--)|返回包含一个或多个矩形区域的 RangeAreas，它具有无效单元格值。 如果所有单元格值都有效，则此函数将返回 null。|
|[FilterCriteria](/javascript/api/excel/excel.filtercriteria)|[subField](/javascript/api/excel/excel.filtercriteria#subfield)|筛选器使用该属性对 richvalue 执行丰富的筛选。|
|[GeometricShape](/javascript/api/excel/excel.geometricshape)|[id](/javascript/api/excel/excel.geometricshape#id)|返回形状标识符。 只读。|
||[shape](/javascript/api/excel/excel.geometricshape#shape)|返回几何形状的形状对象。 只读。|
|[GeometricShapeData](/javascript/api/excel/excel.geometricshapedata)|[id](/javascript/api/excel/excel.geometricshapedata#id)|返回形状标识符。 只读。|
|[GeometricShapeLoadOptions](/javascript/api/excel/excel.geometricshapeloadoptions)|[$all](/javascript/api/excel/excel.geometricshapeloadoptions#$all)||
||[id](/javascript/api/excel/excel.geometricshapeloadoptions#id)|返回形状标识符。 只读。|
||[shape](/javascript/api/excel/excel.geometricshapeloadoptions#shape)|返回几何形状的形状对象。|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getCount()](/javascript/api/excel/excel.groupshapecollection#getcount--)|返回形状组中的形状数量。 只读。|
||[getItem(key: string)](/javascript/api/excel/excel.groupshapecollection#getitem-key-)|按名称或 ID 获取形状。|
||[getItemAt(index: number)](/javascript/api/excel/excel.groupshapecollection#getitemat-index-)|根据其在集合中的位置获取形状。|
||[items](/javascript/api/excel/excel.groupshapecollection#items)|获取此集合中已加载的子项。|
|[GroupShapeCollectionLoadOptions](/javascript/api/excel/excel.groupshapecollectionloadoptions)|[$all](/javascript/api/excel/excel.groupshapecollectionloadoptions#$all)||
||[altTextDescription](/javascript/api/excel/excel.groupshapecollectionloadoptions#alttextdescription)|对于集合中的每一项: 返回或设置 Shape 对象的替代说明文本。|
||[altTextTitle](/javascript/api/excel/excel.groupshapecollectionloadoptions#alttexttitle)|对于集合中的每一项: 返回或设置 Shape 对象的可选标题文本。|
||[connectionSiteCount](/javascript/api/excel/excel.groupshapecollectionloadoptions#connectionsitecount)|对于集合中的每个项目: 返回此形状上的连接结点的数目。 只读。|
||[fill](/javascript/api/excel/excel.groupshapecollectionloadoptions#fill)|对于集合中的每一项: 返回此形状的填充格式。|
||[geometricShape](/javascript/api/excel/excel.groupshapecollectionloadoptions#geometricshape)|对于集合中的每个项目: 返回与形状相关联的几何形状。 如果形状类型不是“GeometricShape”，则会引发错误。|
||[geometricShapeType](/javascript/api/excel/excel.groupshapecollectionloadoptions#geometricshapetype)|对于集合中的每一项: 表示此几何形状的几何形状类型。 有关详细信息，请参阅 Excel.GeometricShapeType。 如果形状类型不是“GeometricShape”，返回 NULL。|
||[组](/javascript/api/excel/excel.groupshapecollectionloadoptions#group)|对于集合中的每个项目: 返回与形状相关联的形状组。 如果形状类型不是“GroupShape”，则会引发错误。|
||[height](/javascript/api/excel/excel.groupshapecollectionloadoptions#height)|对于集合中的每个项目: 表示形状的高度 (以磅为单位)。|
||[id](/javascript/api/excel/excel.groupshapecollectionloadoptions#id)|对于集合中的每一项: 代表形状标识符。 只读。|
||[image](/javascript/api/excel/excel.groupshapecollectionloadoptions#image)|对于集合中的每个项目: 返回与该形状相关联的图像。 如果形状类型不是“Image”，则会引发错误。|
||[left](/javascript/api/excel/excel.groupshapecollectionloadoptions#left)|对于集合中的每个项目: 从形状左侧到工作表左侧的距离 (以磅为单位)。|
||[level](/javascript/api/excel/excel.groupshapecollectionloadoptions#level)|对于集合中的每一项: 代表指定形状的级别。 例如，级别 0 表示形状不是任何组的一部分，级别 1 表示形状是顶级组的一部分，级别 2 表示形状是顶级子组的一部分。|
||[line](/javascript/api/excel/excel.groupshapecollectionloadoptions#line)|对于集合中的每一项: 返回与形状相关联的线条。 如果形状类型不是“Line”，则会引发错误。|
||[lineFormat](/javascript/api/excel/excel.groupshapecollectionloadoptions#lineformat)|对于集合中的每一项: 返回此形状的线条格式。|
||[lockAspectRatio](/javascript/api/excel/excel.groupshapecollectionloadoptions#lockaspectratio)|对于集合中的每一项: 指定是否锁定此形状的纵横比。|
||[name](/javascript/api/excel/excel.groupshapecollectionloadoptions#name)|对于集合中的每一项: 代表形状的名称。|
||[parentGroup](/javascript/api/excel/excel.groupshapecollectionloadoptions#parentgroup)|对于集合中的每个项目: 代表此形状的父组。|
||[rotation](/javascript/api/excel/excel.groupshapecollectionloadoptions#rotation)|对于集合中的每个项目: 表示形状的旋转角度 (以度为单位)。|
||[textFrame](/javascript/api/excel/excel.groupshapecollectionloadoptions#textframe)|对于集合中的每一项: 返回此形状的文本框架对象。 只读。|
||[top](/javascript/api/excel/excel.groupshapecollectionloadoptions#top)|对于集合中的每个项目: 从形状的上边缘到工作表的上边缘之间的距离 (以磅为单位)。|
||[type](/javascript/api/excel/excel.groupshapecollectionloadoptions#type)|对于集合中的每一项: 返回此形状的类型。 有关详细信息，请参阅 Excel.ShapeType。 只读。|
||[visible](/javascript/api/excel/excel.groupshapecollectionloadoptions#visible)|对于集合中的每一项: 表示此形状的可见性。|
||[width](/javascript/api/excel/excel.groupshapecollectionloadoptions#width)|对于集合中的每个项目: 代表形状的宽度 (以磅为单位)。|
||[zOrderPosition](/javascript/api/excel/excel.groupshapecollectionloadoptions#zorderposition)|对于集合中的每一项: 返回指定的形状在 z-顺序中的位置, 0 表示顺序堆栈的底部。 只读。|
|[HeaderFooter](/javascript/api/excel/excel.headerfooter)|[centerFooter](/javascript/api/excel/excel.headerfooter#centerfooter)|获取或设置工作表的中心页脚。|
||[centerHeader](/javascript/api/excel/excel.headerfooter#centerheader)|获取或设置工作表的中心页眉。|
||[LeftFooter](/javascript/api/excel/excel.headerfooter#leftfooter)|获取或设置工作表的左侧页脚。|
||[leftHeader](/javascript/api/excel/excel.headerfooter#leftheader)|获取或设置工作表的左侧页眉。|
||[rightFooter](/javascript/api/excel/excel.headerfooter#rightfooter)|获取或设置工作表的右侧页脚。|
||[rightHeader](/javascript/api/excel/excel.headerfooter#rightheader)|获取或设置工作表的右侧页眉。|
||[set (properties: HeaderFooter)](/javascript/api/excel/excel.headerfooter#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: HeaderFooterUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.headerfooter#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
|[HeaderFooterData](/javascript/api/excel/excel.headerfooterdata)|[centerFooter](/javascript/api/excel/excel.headerfooterdata#centerfooter)|获取或设置工作表的中心页脚。|
||[centerHeader](/javascript/api/excel/excel.headerfooterdata#centerheader)|获取或设置工作表的中心页眉。|
||[LeftFooter](/javascript/api/excel/excel.headerfooterdata#leftfooter)|获取或设置工作表的左侧页脚。|
||[leftHeader](/javascript/api/excel/excel.headerfooterdata#leftheader)|获取或设置工作表的左侧页眉。|
||[rightFooter](/javascript/api/excel/excel.headerfooterdata#rightfooter)|获取或设置工作表的右侧页脚。|
||[rightHeader](/javascript/api/excel/excel.headerfooterdata#rightheader)|获取或设置工作表的右侧页眉。|
|[HeaderFooterGroup](/javascript/api/excel/excel.headerfootergroup)|[defaultForAllPages](/javascript/api/excel/excel.headerfootergroup#defaultforallpages)|常规页眉/页脚，除非指定偶数页/奇数页或首页，否则适用于所有页面，|
||[evenPages](/javascript/api/excel/excel.headerfootergroup#evenpages)|用于偶数页的页眉/页脚，需要为奇数页指定奇数页页眉/页脚。|
||[firstPage](/javascript/api/excel/excel.headerfootergroup#firstpage)|首页的页眉/页脚，为所有其他页使用常规或偶数页/奇数页页眉/页脚。|
||[oddPages](/javascript/api/excel/excel.headerfootergroup#oddpages)|用于奇数页的页眉/页脚，需要为偶数页指定偶数页页眉/页脚。|
||[set (properties: HeaderFooterGroup)](/javascript/api/excel/excel.headerfootergroup#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: HeaderFooterGroupUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.headerfootergroup#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
||[state](/javascript/api/excel/excel.headerfootergroup#state)|获取或设置所设置的页眉/页脚的状态。 有关详细信息，请参阅 Excel.HeaderFooterState。|
||[useSheetMargins](/javascript/api/excel/excel.headerfootergroup#usesheetmargins)|获取或设置一个标记，指示页眉/页脚是否与工作表的页面布局选项中设置的页边距对齐。|
||[useSheetScale](/javascript/api/excel/excel.headerfootergroup#usesheetscale)|获取或设置一个标记，指示是否应按照工作表的页面布局选项中设置的页面缩放百分比来缩放页眉/页脚。|
|[HeaderFooterGroupData](/javascript/api/excel/excel.headerfootergroupdata)|[defaultForAllPages](/javascript/api/excel/excel.headerfootergroupdata#defaultforallpages)|常规页眉/页脚，除非指定偶数页/奇数页或首页，否则适用于所有页面，|
||[evenPages](/javascript/api/excel/excel.headerfootergroupdata#evenpages)|用于偶数页的页眉/页脚，需要为奇数页指定奇数页页眉/页脚。|
||[firstPage](/javascript/api/excel/excel.headerfootergroupdata#firstpage)|首页的页眉/页脚，为所有其他页使用常规或偶数页/奇数页页眉/页脚。|
||[oddPages](/javascript/api/excel/excel.headerfootergroupdata#oddpages)|用于奇数页的页眉/页脚，需要为偶数页指定偶数页页眉/页脚。|
||[state](/javascript/api/excel/excel.headerfootergroupdata#state)|获取或设置所设置的页眉/页脚的状态。 有关详细信息，请参阅 Excel.HeaderFooterState。|
||[useSheetMargins](/javascript/api/excel/excel.headerfootergroupdata#usesheetmargins)|获取或设置一个标记，指示页眉/页脚是否与工作表的页面布局选项中设置的页边距对齐。|
||[useSheetScale](/javascript/api/excel/excel.headerfootergroupdata#usesheetscale)|获取或设置一个标记，指示是否应按照工作表的页面布局选项中设置的页面缩放百分比来缩放页眉/页脚。|
|[HeaderFooterGroupLoadOptions](/javascript/api/excel/excel.headerfootergrouploadoptions)|[$all](/javascript/api/excel/excel.headerfootergrouploadoptions#$all)||
||[defaultForAllPages](/javascript/api/excel/excel.headerfootergrouploadoptions#defaultforallpages)|常规页眉/页脚，除非指定偶数页/奇数页或首页，否则适用于所有页面，|
||[evenPages](/javascript/api/excel/excel.headerfootergrouploadoptions#evenpages)|用于偶数页的页眉/页脚，需要为奇数页指定奇数页页眉/页脚。|
||[firstPage](/javascript/api/excel/excel.headerfootergrouploadoptions#firstpage)|首页的页眉/页脚，为所有其他页使用常规或偶数页/奇数页页眉/页脚。|
||[oddPages](/javascript/api/excel/excel.headerfootergrouploadoptions#oddpages)|用于奇数页的页眉/页脚，需要为偶数页指定偶数页页眉/页脚。|
||[state](/javascript/api/excel/excel.headerfootergrouploadoptions#state)|获取或设置所设置的页眉/页脚的状态。 有关详细信息，请参阅 Excel.HeaderFooterState。|
||[useSheetMargins](/javascript/api/excel/excel.headerfootergrouploadoptions#usesheetmargins)|获取或设置一个标记，指示页眉/页脚是否与工作表的页面布局选项中设置的页边距对齐。|
||[useSheetScale](/javascript/api/excel/excel.headerfootergrouploadoptions#usesheetscale)|获取或设置一个标记，指示是否应按照工作表的页面布局选项中设置的页面缩放百分比来缩放页眉/页脚。|
|[HeaderFooterGroupUpdateData](/javascript/api/excel/excel.headerfootergroupupdatedata)|[defaultForAllPages](/javascript/api/excel/excel.headerfootergroupupdatedata#defaultforallpages)|常规页眉/页脚，除非指定偶数页/奇数页或首页，否则适用于所有页面，|
||[evenPages](/javascript/api/excel/excel.headerfootergroupupdatedata#evenpages)|用于偶数页的页眉/页脚，需要为奇数页指定奇数页页眉/页脚。|
||[firstPage](/javascript/api/excel/excel.headerfootergroupupdatedata#firstpage)|首页的页眉/页脚，为所有其他页使用常规或偶数页/奇数页页眉/页脚。|
||[oddPages](/javascript/api/excel/excel.headerfootergroupupdatedata#oddpages)|用于奇数页的页眉/页脚，需要为偶数页指定偶数页页眉/页脚。|
||[state](/javascript/api/excel/excel.headerfootergroupupdatedata#state)|获取或设置所设置的页眉/页脚的状态。 有关详细信息，请参阅 Excel.HeaderFooterState。|
||[useSheetMargins](/javascript/api/excel/excel.headerfootergroupupdatedata#usesheetmargins)|获取或设置一个标记，指示页眉/页脚是否与工作表的页面布局选项中设置的页边距对齐。|
||[useSheetScale](/javascript/api/excel/excel.headerfootergroupupdatedata#usesheetscale)|获取或设置一个标记，指示是否应按照工作表的页面布局选项中设置的页面缩放百分比来缩放页眉/页脚。|
|[HeaderFooterLoadOptions](/javascript/api/excel/excel.headerfooterloadoptions)|[$all](/javascript/api/excel/excel.headerfooterloadoptions#$all)||
||[centerFooter](/javascript/api/excel/excel.headerfooterloadoptions#centerfooter)|获取或设置工作表的中心页脚。|
||[centerHeader](/javascript/api/excel/excel.headerfooterloadoptions#centerheader)|获取或设置工作表的中心页眉。|
||[LeftFooter](/javascript/api/excel/excel.headerfooterloadoptions#leftfooter)|获取或设置工作表的左侧页脚。|
||[leftHeader](/javascript/api/excel/excel.headerfooterloadoptions#leftheader)|获取或设置工作表的左侧页眉。|
||[rightFooter](/javascript/api/excel/excel.headerfooterloadoptions#rightfooter)|获取或设置工作表的右侧页脚。|
||[rightHeader](/javascript/api/excel/excel.headerfooterloadoptions#rightheader)|获取或设置工作表的右侧页眉。|
|[HeaderFooterUpdateData](/javascript/api/excel/excel.headerfooterupdatedata)|[centerFooter](/javascript/api/excel/excel.headerfooterupdatedata#centerfooter)|获取或设置工作表的中心页脚。|
||[centerHeader](/javascript/api/excel/excel.headerfooterupdatedata#centerheader)|获取或设置工作表的中心页眉。|
||[LeftFooter](/javascript/api/excel/excel.headerfooterupdatedata#leftfooter)|获取或设置工作表的左侧页脚。|
||[leftHeader](/javascript/api/excel/excel.headerfooterupdatedata#leftheader)|获取或设置工作表的左侧页眉。|
||[rightFooter](/javascript/api/excel/excel.headerfooterupdatedata#rightfooter)|获取或设置工作表的右侧页脚。|
||[rightHeader](/javascript/api/excel/excel.headerfooterupdatedata#rightheader)|获取或设置工作表的右侧页眉。|
|[Image](/javascript/api/excel/excel.image)|[format](/javascript/api/excel/excel.image#format)|返回图像的格式。 只读。|
||[id](/javascript/api/excel/excel.image#id)|表示图像对象的形状标识符。 只读。|
||[shape](/javascript/api/excel/excel.image#shape)|返回与图像关联的形状对象。 只读。|
|[ImageData](/javascript/api/excel/excel.imagedata)|[format](/javascript/api/excel/excel.imagedata#format)|返回图像的格式。 只读。|
||[id](/javascript/api/excel/excel.imagedata#id)|表示图像对象的形状标识符。 只读。|
|[ImageLoadOptions](/javascript/api/excel/excel.imageloadoptions)|[$all](/javascript/api/excel/excel.imageloadoptions#$all)||
||[format](/javascript/api/excel/excel.imageloadoptions#format)|返回图像的格式。 只读。|
||[id](/javascript/api/excel/excel.imageloadoptions#id)|表示图像对象的形状标识符。 只读。|
||[shape](/javascript/api/excel/excel.imageloadoptions#shape)|返回与图像关联的形状对象。|
|[IterativeCalculation](/javascript/api/excel/excel.iterativecalculation)|[enabled](/javascript/api/excel/excel.iterativecalculation#enabled)|如果 Excel 使用迭代来处理循环引用，则为 True。|
||[maxChange](/javascript/api/excel/excel.iterativecalculation#maxchange)|返回或设置 Excel 处理循环引用时迭代之间的最大变化值。|
||[maxIteration](/javascript/api/excel/excel.iterativecalculation#maxiteration)|返回或设置 Excel 处理循环引用的最大迭代次数。|
||[set (properties: IterativeCalculation)](/javascript/api/excel/excel.iterativecalculation#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: IterativeCalculationUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.iterativecalculation#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
|[IterativeCalculationData](/javascript/api/excel/excel.iterativecalculationdata)|[enabled](/javascript/api/excel/excel.iterativecalculationdata#enabled)|如果 Excel 使用迭代来处理循环引用，则为 True。|
||[maxChange](/javascript/api/excel/excel.iterativecalculationdata#maxchange)|返回或设置 Excel 处理循环引用时迭代之间的最大变化值。|
||[maxIteration](/javascript/api/excel/excel.iterativecalculationdata#maxiteration)|返回或设置 Excel 处理循环引用的最大迭代次数。|
|[IterativeCalculationLoadOptions](/javascript/api/excel/excel.iterativecalculationloadoptions)|[$all](/javascript/api/excel/excel.iterativecalculationloadoptions#$all)||
||[enabled](/javascript/api/excel/excel.iterativecalculationloadoptions#enabled)|如果 Excel 使用迭代来处理循环引用，则为 True。|
||[maxChange](/javascript/api/excel/excel.iterativecalculationloadoptions#maxchange)|返回或设置 Excel 处理循环引用时迭代之间的最大变化值。|
||[maxIteration](/javascript/api/excel/excel.iterativecalculationloadoptions#maxiteration)|返回或设置 Excel 处理循环引用的最大迭代次数。|
|[IterativeCalculationUpdateData](/javascript/api/excel/excel.iterativecalculationupdatedata)|[enabled](/javascript/api/excel/excel.iterativecalculationupdatedata#enabled)|如果 Excel 使用迭代来处理循环引用，则为 True。|
||[maxChange](/javascript/api/excel/excel.iterativecalculationupdatedata#maxchange)|返回或设置 Excel 处理循环引用时迭代之间的最大变化值。|
||[maxIteration](/javascript/api/excel/excel.iterativecalculationupdatedata#maxiteration)|返回或设置 Excel 处理循环引用的最大迭代次数。|
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
||[set (properties: Excel. 行)](/javascript/api/excel/excel.line#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: LineUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.line#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
|[LineData](/javascript/api/excel/excel.linedata)|[beginArrowheadLength](/javascript/api/excel/excel.linedata#beginarrowheadlength)|表示指定线条始端的箭头长度。|
||[beginArrowheadStyle](/javascript/api/excel/excel.linedata#beginarrowheadstyle)|表示指定线条始端的箭头样式。|
||[beginArrowheadWidth](/javascript/api/excel/excel.linedata#beginarrowheadwidth)|表示指定线条始端的箭头宽度。|
||[beginConnectedSite](/javascript/api/excel/excel.linedata#beginconnectedsite)|表示连接线始端所连接的连接站点。 只读。 当线条的始端没有附加到任何形状时，返回 null。|
||[connectorType](/javascript/api/excel/excel.linedata#connectortype)|表示线条的连接器类型。|
||[endArrowheadLength](/javascript/api/excel/excel.linedata#endarrowheadlength)|表示指定线条末端的箭头长度。|
||[endArrowheadStyle](/javascript/api/excel/excel.linedata#endarrowheadstyle)|表示指定线条末端的箭头样式。|
||[endArrowheadWidth](/javascript/api/excel/excel.linedata#endarrowheadwidth)|表示指定线条末端的箭头宽度。|
||[endConnectedSite](/javascript/api/excel/excel.linedata#endconnectedsite)|表示连接线末端所连接的连接站点。 只读。 当线条的末端没有附加到任何形状时，返回 null。|
||[id](/javascript/api/excel/excel.linedata#id)|表示形状标识符。 只读。|
||[isBeginConnected](/javascript/api/excel/excel.linedata#isbeginconnected)|指定指定线条的始端是否连接到形状。 只读。|
||[isEndConnected](/javascript/api/excel/excel.linedata#isendconnected)|指定指定线条的末端是否连接到形状。 只读。|
|[LineLoadOptions](/javascript/api/excel/excel.lineloadoptions)|[$all](/javascript/api/excel/excel.lineloadoptions#$all)||
||[beginArrowheadLength](/javascript/api/excel/excel.lineloadoptions#beginarrowheadlength)|表示指定线条始端的箭头长度。|
||[beginArrowheadStyle](/javascript/api/excel/excel.lineloadoptions#beginarrowheadstyle)|表示指定线条始端的箭头样式。|
||[beginArrowheadWidth](/javascript/api/excel/excel.lineloadoptions#beginarrowheadwidth)|表示指定线条始端的箭头宽度。|
||[beginConnectedShape](/javascript/api/excel/excel.lineloadoptions#beginconnectedshape)|表示指定线条始端所附加到的形状。|
||[beginConnectedSite](/javascript/api/excel/excel.lineloadoptions#beginconnectedsite)|表示连接线始端所连接的连接站点。 只读。 当线条的始端没有附加到任何形状时，返回 null。|
||[connectorType](/javascript/api/excel/excel.lineloadoptions#connectortype)|表示线条的连接器类型。|
||[endArrowheadLength](/javascript/api/excel/excel.lineloadoptions#endarrowheadlength)|表示指定线条末端的箭头长度。|
||[endArrowheadStyle](/javascript/api/excel/excel.lineloadoptions#endarrowheadstyle)|表示指定线条末端的箭头样式。|
||[endArrowheadWidth](/javascript/api/excel/excel.lineloadoptions#endarrowheadwidth)|表示指定线条末端的箭头宽度。|
||[endConnectedShape](/javascript/api/excel/excel.lineloadoptions#endconnectedshape)|表示指定线条末端所附加到的形状。|
||[endConnectedSite](/javascript/api/excel/excel.lineloadoptions#endconnectedsite)|表示连接线末端所连接的连接站点。 只读。 当线条的末端没有附加到任何形状时，返回 null。|
||[id](/javascript/api/excel/excel.lineloadoptions#id)|表示形状标识符。 只读。|
||[isBeginConnected](/javascript/api/excel/excel.lineloadoptions#isbeginconnected)|指定指定线条的始端是否连接到形状。 只读。|
||[isEndConnected](/javascript/api/excel/excel.lineloadoptions#isendconnected)|指定指定线条的末端是否连接到形状。 只读。|
||[shape](/javascript/api/excel/excel.lineloadoptions#shape)|返回与线条关联的形状对象。|
|[LineUpdateData](/javascript/api/excel/excel.lineupdatedata)|[beginArrowheadLength](/javascript/api/excel/excel.lineupdatedata#beginarrowheadlength)|表示指定线条始端的箭头长度。|
||[beginArrowheadStyle](/javascript/api/excel/excel.lineupdatedata#beginarrowheadstyle)|表示指定线条始端的箭头样式。|
||[beginArrowheadWidth](/javascript/api/excel/excel.lineupdatedata#beginarrowheadwidth)|表示指定线条始端的箭头宽度。|
||[connectorType](/javascript/api/excel/excel.lineupdatedata#connectortype)|表示线条的连接器类型。|
||[endArrowheadLength](/javascript/api/excel/excel.lineupdatedata#endarrowheadlength)|表示指定线条末端的箭头长度。|
||[endArrowheadStyle](/javascript/api/excel/excel.lineupdatedata#endarrowheadstyle)|表示指定线条末端的箭头样式。|
||[endArrowheadWidth](/javascript/api/excel/excel.lineupdatedata#endarrowheadwidth)|表示指定线条末端的箭头宽度。|
|[PageBreak](/javascript/api/excel/excel.pagebreak)|[delete()](/javascript/api/excel/excel.pagebreak#delete--)|删除分页符对象。|
||[getCellAfterBreak()](/javascript/api/excel/excel.pagebreak#getcellafterbreak--)|获取分页符后的第一个单元格。|
||[columnIndex](/javascript/api/excel/excel.pagebreak#columnindex)|表示分页符的列索引|
||[rowIndex](/javascript/api/excel/excel.pagebreak#rowindex)|表示分页符的行索引|
|[PageBreakCollection](/javascript/api/excel/excel.pagebreakcollection)|[add(pageBreakRange: Range \| string)](/javascript/api/excel/excel.pagebreakcollection#add-pagebreakrange-)|在指定区域的左上角单元格之前添加分页符。|
||[getCount()](/javascript/api/excel/excel.pagebreakcollection#getcount--)|获取集合中的分页符数量。|
||[getItem(index: number)](/javascript/api/excel/excel.pagebreakcollection#getitem-index-)|通过索引获取分页符对象。|
||[items](/javascript/api/excel/excel.pagebreakcollection#items)|获取此集合中已加载的子项。|
||[removePageBreaks()](/javascript/api/excel/excel.pagebreakcollection#removepagebreaks--)|重置集合中的所有手动分页符。|
|[PageBreakCollectionLoadOptions](/javascript/api/excel/excel.pagebreakcollectionloadoptions)|[$all](/javascript/api/excel/excel.pagebreakcollectionloadoptions#$all)||
||[columnIndex](/javascript/api/excel/excel.pagebreakcollectionloadoptions#columnindex)|对于集合中的每一项: 代表分页符的列索引|
||[rowIndex](/javascript/api/excel/excel.pagebreakcollectionloadoptions#rowindex)|对于集合中的每一项: 代表分页符的行索引|
|[PageBreakData](/javascript/api/excel/excel.pagebreakdata)|[columnIndex](/javascript/api/excel/excel.pagebreakdata#columnindex)|表示分页符的列索引|
||[rowIndex](/javascript/api/excel/excel.pagebreakdata#rowindex)|表示分页符的行索引|
|[PageBreakLoadOptions](/javascript/api/excel/excel.pagebreakloadoptions)|[$all](/javascript/api/excel/excel.pagebreakloadoptions#$all)||
||[columnIndex](/javascript/api/excel/excel.pagebreakloadoptions#columnindex)|表示分页符的列索引|
||[rowIndex](/javascript/api/excel/excel.pagebreakloadoptions#rowindex)|表示分页符的行索引|
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
||[set (properties: 页面布局)](/javascript/api/excel/excel.pagelayout#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: PageLayoutUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.pagelayout#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
||[setPrintArea(printArea: Range \| RangeAreas \| string)](/javascript/api/excel/excel.pagelayout#setprintarea-printarea-)|设置工作表的打印区域。|
||[setPrintMargins(unit: "Points" \| "Inches" \| "Centimeters", marginOptions: Excel.PageLayoutMarginOptions)](/javascript/api/excel/excel.pagelayout#setprintmargins-unit--marginoptions-)|设置带单位的工作表的页边距。|
||[setPrintMargins(unit: Excel.PrintMarginUnit, marginOptions: Excel.PageLayoutMarginOptions)](/javascript/api/excel/excel.pagelayout#setprintmargins-unit--marginoptions-)|设置带单位的工作表的页边距。|
||[setPrintTitleColumns(printTitleColumns: Range \| string)](/javascript/api/excel/excel.pagelayout#setprinttitlecolumns-printtitlecolumns-)|设置列，这些列包含要在打印的工作表的每页左侧重复的单元格。|
||[setPrintTitleRows(printTitleRows: Range \| string)](/javascript/api/excel/excel.pagelayout#setprinttitlerows-printtitlerows-)|设置行，这些行包含要在打印的工作表的每页顶部重复的单元格。|
||[topMargin](/javascript/api/excel/excel.pagelayout#topmargin)|获取或设置在打印时使用的工作表的上边距（以磅为单位）。|
||[zoom](/javascript/api/excel/excel.pagelayout#zoom)|获取或设置工作表的打印缩放选项。|
|[PageLayoutData](/javascript/api/excel/excel.pagelayoutdata)|[blackAndWhite](/javascript/api/excel/excel.pagelayoutdata#blackandwhite)|获取或设置工作表的黑白打印选项。|
||[bottomMargin](/javascript/api/excel/excel.pagelayoutdata#bottommargin)|获取或设置要用于打印的工作表的底部页边距（以磅为单位）。|
||[centerHorizontally](/javascript/api/excel/excel.pagelayoutdata#centerhorizontally)|获取或设置工作表的中心水平标记。 此标记确定在打印时是否水平居中工作表。|
||[centerVertically](/javascript/api/excel/excel.pagelayoutdata#centervertically)|获取或设置工作表的中心垂直标记。 此标记确定在打印时是否垂直居中工作表。|
||[draftMode](/javascript/api/excel/excel.pagelayoutdata#draftmode)|获取或设置工作表的草稿模式选项。 如果为 True，则将打印没有图形的工作表。|
||[firstPageNumber](/javascript/api/excel/excel.pagelayoutdata#firstpagenumber)|获取或设置要打印的工作表的首页页码。 Null 值表示“自动”页码编号。|
||[footerMargin](/javascript/api/excel/excel.pagelayoutdata#footermargin)|获取或设置在打印时使用的工作表的页脚边距（以磅为单位）。|
||[headerMargin](/javascript/api/excel/excel.pagelayoutdata#headermargin)|获取或设置在打印时使用的工作表的页眉边距（以磅为单位）。|
||[headersFooters](/javascript/api/excel/excel.pagelayoutdata#headersfooters)|工作表的页眉和页脚配置。|
||[leftMargin](/javascript/api/excel/excel.pagelayoutdata#leftmargin)|获取或设置在打印时使用的工作表的左边距（以磅为单位）。|
||[orientation](/javascript/api/excel/excel.pagelayoutdata#orientation)|获取或设置工作表的页面方向。|
||[paperSize](/javascript/api/excel/excel.pagelayoutdata#papersize)|获取或设置工作表的页面纸张大小。|
||[printComments](/javascript/api/excel/excel.pagelayoutdata#printcomments)|获取或设置在打印时是否应该显示工作表的批注。|
||[printErrors](/javascript/api/excel/excel.pagelayoutdata#printerrors)|获取或设置工作表的打印错误选项。|
||[printGridlines](/javascript/api/excel/excel.pagelayoutdata#printgridlines)|获取或设置工作表的打印网格线标记。 此标记确定是否打印网格线。|
||[printHeadings](/javascript/api/excel/excel.pagelayoutdata#printheadings)|获取或设置工作表的打印标题标记。 此标记确定是否打印标题。|
||[printOrder](/javascript/api/excel/excel.pagelayoutdata#printorder)|获取或设置工作表的页面打印顺序选项。 它指定用于处理打印页码的顺序。|
||[rightMargin](/javascript/api/excel/excel.pagelayoutdata#rightmargin)|获取或设置在打印时使用的工作表的右边距（以磅为单位）。|
||[topMargin](/javascript/api/excel/excel.pagelayoutdata#topmargin)|获取或设置在打印时使用的工作表的上边距（以磅为单位）。|
||[zoom](/javascript/api/excel/excel.pagelayoutdata#zoom)|获取或设置工作表的打印缩放选项。|
|[PageLayoutLoadOptions](/javascript/api/excel/excel.pagelayoutloadoptions)|[$all](/javascript/api/excel/excel.pagelayoutloadoptions#$all)||
||[blackAndWhite](/javascript/api/excel/excel.pagelayoutloadoptions#blackandwhite)|获取或设置工作表的黑白打印选项。|
||[bottomMargin](/javascript/api/excel/excel.pagelayoutloadoptions#bottommargin)|获取或设置要用于打印的工作表的底部页边距（以磅为单位）。|
||[centerHorizontally](/javascript/api/excel/excel.pagelayoutloadoptions#centerhorizontally)|获取或设置工作表的中心水平标记。 此标记确定在打印时是否水平居中工作表。|
||[centerVertically](/javascript/api/excel/excel.pagelayoutloadoptions#centervertically)|获取或设置工作表的中心垂直标记。 此标记确定在打印时是否垂直居中工作表。|
||[draftMode](/javascript/api/excel/excel.pagelayoutloadoptions#draftmode)|获取或设置工作表的草稿模式选项。 如果为 True，则将打印没有图形的工作表。|
||[firstPageNumber](/javascript/api/excel/excel.pagelayoutloadoptions#firstpagenumber)|获取或设置要打印的工作表的首页页码。 Null 值表示“自动”页码编号。|
||[footerMargin](/javascript/api/excel/excel.pagelayoutloadoptions#footermargin)|获取或设置在打印时使用的工作表的页脚边距（以磅为单位）。|
||[headerMargin](/javascript/api/excel/excel.pagelayoutloadoptions#headermargin)|获取或设置在打印时使用的工作表的页眉边距（以磅为单位）。|
||[headersFooters](/javascript/api/excel/excel.pagelayoutloadoptions#headersfooters)|工作表的页眉和页脚配置。|
||[leftMargin](/javascript/api/excel/excel.pagelayoutloadoptions#leftmargin)|获取或设置在打印时使用的工作表的左边距（以磅为单位）。|
||[orientation](/javascript/api/excel/excel.pagelayoutloadoptions#orientation)|获取或设置工作表的页面方向。|
||[paperSize](/javascript/api/excel/excel.pagelayoutloadoptions#papersize)|获取或设置工作表的页面纸张大小。|
||[printComments](/javascript/api/excel/excel.pagelayoutloadoptions#printcomments)|获取或设置在打印时是否应该显示工作表的批注。|
||[printErrors](/javascript/api/excel/excel.pagelayoutloadoptions#printerrors)|获取或设置工作表的打印错误选项。|
||[printGridlines](/javascript/api/excel/excel.pagelayoutloadoptions#printgridlines)|获取或设置工作表的打印网格线标记。 此标记确定是否打印网格线。|
||[printHeadings](/javascript/api/excel/excel.pagelayoutloadoptions#printheadings)|获取或设置工作表的打印标题标记。 此标记确定是否打印标题。|
||[printOrder](/javascript/api/excel/excel.pagelayoutloadoptions#printorder)|获取或设置工作表的页面打印顺序选项。 它指定用于处理打印页码的顺序。|
||[rightMargin](/javascript/api/excel/excel.pagelayoutloadoptions#rightmargin)|获取或设置在打印时使用的工作表的右边距（以磅为单位）。|
||[topMargin](/javascript/api/excel/excel.pagelayoutloadoptions#topmargin)|获取或设置在打印时使用的工作表的上边距（以磅为单位）。|
||[zoom](/javascript/api/excel/excel.pagelayoutloadoptions#zoom)|获取或设置工作表的打印缩放选项。|
|[PageLayoutMarginOptions](/javascript/api/excel/excel.pagelayoutmarginoptions)|[bottom](/javascript/api/excel/excel.pagelayoutmarginoptions#bottom)|表示要在打印时使用的页面布局下边距（使用指定的单位）。|
||[footer](/javascript/api/excel/excel.pagelayoutmarginoptions#footer)|表示要在打印时使用的页面布局页脚边距（使用指定的单位）。|
||[header](/javascript/api/excel/excel.pagelayoutmarginoptions#header)|表示要在打印时使用的页面布局页眉边距（使用指定的单位）。|
||[left](/javascript/api/excel/excel.pagelayoutmarginoptions#left)|表示要在打印时使用的页面布局左边距（使用指定的单位）。|
||[right](/javascript/api/excel/excel.pagelayoutmarginoptions#right)|表示要在打印时使用的页面布局右边距（使用指定的单位）。|
||[top](/javascript/api/excel/excel.pagelayoutmarginoptions#top)|表示要在打印时使用的页面布局上边距（使用指定的单位）。|
|[PageLayoutUpdateData](/javascript/api/excel/excel.pagelayoutupdatedata)|[blackAndWhite](/javascript/api/excel/excel.pagelayoutupdatedata#blackandwhite)|获取或设置工作表的黑白打印选项。|
||[bottomMargin](/javascript/api/excel/excel.pagelayoutupdatedata#bottommargin)|获取或设置要用于打印的工作表的底部页边距（以磅为单位）。|
||[centerHorizontally](/javascript/api/excel/excel.pagelayoutupdatedata#centerhorizontally)|获取或设置工作表的中心水平标记。 此标记确定在打印时是否水平居中工作表。|
||[centerVertically](/javascript/api/excel/excel.pagelayoutupdatedata#centervertically)|获取或设置工作表的中心垂直标记。 此标记确定在打印时是否垂直居中工作表。|
||[draftMode](/javascript/api/excel/excel.pagelayoutupdatedata#draftmode)|获取或设置工作表的草稿模式选项。 如果为 True，则将打印没有图形的工作表。|
||[firstPageNumber](/javascript/api/excel/excel.pagelayoutupdatedata#firstpagenumber)|获取或设置要打印的工作表的首页页码。 Null 值表示“自动”页码编号。|
||[footerMargin](/javascript/api/excel/excel.pagelayoutupdatedata#footermargin)|获取或设置在打印时使用的工作表的页脚边距（以磅为单位）。|
||[headerMargin](/javascript/api/excel/excel.pagelayoutupdatedata#headermargin)|获取或设置在打印时使用的工作表的页眉边距（以磅为单位）。|
||[headersFooters](/javascript/api/excel/excel.pagelayoutupdatedata#headersfooters)|工作表的页眉和页脚配置。|
||[leftMargin](/javascript/api/excel/excel.pagelayoutupdatedata#leftmargin)|获取或设置在打印时使用的工作表的左边距（以磅为单位）。|
||[orientation](/javascript/api/excel/excel.pagelayoutupdatedata#orientation)|获取或设置工作表的页面方向。|
||[paperSize](/javascript/api/excel/excel.pagelayoutupdatedata#papersize)|获取或设置工作表的页面纸张大小。|
||[printComments](/javascript/api/excel/excel.pagelayoutupdatedata#printcomments)|获取或设置在打印时是否应该显示工作表的批注。|
||[printErrors](/javascript/api/excel/excel.pagelayoutupdatedata#printerrors)|获取或设置工作表的打印错误选项。|
||[printGridlines](/javascript/api/excel/excel.pagelayoutupdatedata#printgridlines)|获取或设置工作表的打印网格线标记。 此标记确定是否打印网格线。|
||[printHeadings](/javascript/api/excel/excel.pagelayoutupdatedata#printheadings)|获取或设置工作表的打印标题标记。 此标记确定是否打印标题。|
||[printOrder](/javascript/api/excel/excel.pagelayoutupdatedata#printorder)|获取或设置工作表的页面打印顺序选项。 它指定用于处理打印页码的顺序。|
||[rightMargin](/javascript/api/excel/excel.pagelayoutupdatedata#rightmargin)|获取或设置在打印时使用的工作表的右边距（以磅为单位）。|
||[topMargin](/javascript/api/excel/excel.pagelayoutupdatedata#topmargin)|获取或设置在打印时使用的工作表的上边距（以磅为单位）。|
||[zoom](/javascript/api/excel/excel.pagelayoutupdatedata#zoom)|获取或设置工作表的打印缩放选项。|
|[PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions)|[horizontalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#horizontalfittopages)|水平放置的页数。 如果使用百分比缩放，则此值可以为 null。|
||[scale](/javascript/api/excel/excel.pagelayoutzoomoptions#scale)|打印页面缩放值可以介于 10 至 400 之间。 如果已指定适应页面高度或宽度，则此值可以为 null。|
||[verticalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#verticalfittopages)|垂直放置的页数。 如果使用百分比缩放，则此值可以为 null。|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[sortByValues (sortBy: "升序" \| "降序", ValuesHierarchy: DataPivotHierarchy, pivotItemScope？: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotfield#sortbyvalues-sortby--valueshierarchy--pivotitemscope-)|按给定范围中的指定值对 PivotField 进行排序。 该范围定义将使用哪些特定值进行排序|
||[sortByValues(sortBy: Excel.SortBy, valuesHierarchy: Excel.DataPivotHierarchy, pivotItemScope?: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotfield#sortbyvalues-sortby--valueshierarchy--pivotitemscope-)|按给定范围中的指定值对 PivotField 进行排序。 该范围定义将使用哪些特定值进行排序|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[autoFormat](/javascript/api/excel/excel.pivotlayout#autoformat)|指定是否在刷新或移动字段时自动进行格式化|
||[getDataHierarchy(cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#getdatahierarchy-cell-)|获取 DataHierarchy，它用于计算数据透视表中指定区域内的值。|
||[getPivotItems(axis: "Unknown" \| "Row" \| "Column" \| "Data" \| "Filter", cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#getpivotitems-axis--cell-)|从构成数据透视表中指定区域内的值的轴获取 PivotItems。|
||[getPivotItems(axis: Excel.PivotAxis, cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#getpivotitems-axis--cell-)|从构成数据透视表中指定区域内的值的轴获取 PivotItems。|
||[preserveFormatting](/javascript/api/excel/excel.pivotlayout#preserveformatting)|指定是否在通过透视、排序或更改页面字段项等操作来刷新或重新计算报表时保留格式。|
||[setAutoSortOnCell (cell: Range \| String, sortBy: "升序" \| "降序")](/javascript/api/excel/excel.pivotlayout#setautosortoncell-cell--sortby-)|将数据透视表设置为使用指定的单元格设置自动排序，以自动选择排序的所有条件和上下文。 这与从 UI 应用自动排序的行为相同。|
||[setAutoSortOnCell(cell: Range \| string, sortBy: Excel.SortBy)](/javascript/api/excel/excel.pivotlayout#setautosortoncell-cell--sortby-)|将数据透视表设置为使用指定的单元格设置自动排序，以自动选择排序的所有条件和上下文。 这与从 UI 应用自动排序的行为相同。|
|[PivotLayoutData](/javascript/api/excel/excel.pivotlayoutdata)|[autoFormat](/javascript/api/excel/excel.pivotlayoutdata#autoformat)|指定是否在刷新或移动字段时自动进行格式化|
||[preserveFormatting](/javascript/api/excel/excel.pivotlayoutdata#preserveformatting)|指定是否在通过透视、排序或更改页面字段项等操作来刷新或重新计算报表时保留格式。|
|[PivotLayoutLoadOptions](/javascript/api/excel/excel.pivotlayoutloadoptions)|[autoFormat](/javascript/api/excel/excel.pivotlayoutloadoptions#autoformat)|指定是否在刷新或移动字段时自动进行格式化|
||[preserveFormatting](/javascript/api/excel/excel.pivotlayoutloadoptions#preserveformatting)|指定是否在通过透视、排序或更改页面字段项等操作来刷新或重新计算报表时保留格式。|
|[PivotLayoutUpdateData](/javascript/api/excel/excel.pivotlayoutupdatedata)|[autoFormat](/javascript/api/excel/excel.pivotlayoutupdatedata#autoformat)|指定是否在刷新或移动字段时自动进行格式化|
||[preserveFormatting](/javascript/api/excel/excel.pivotlayoutupdatedata#preserveformatting)|指定是否在通过透视、排序或更改页面字段项等操作来刷新或重新计算报表时保留格式。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[enableDataValueEditing](/javascript/api/excel/excel.pivottable#enabledatavalueediting)|指定数据透视表是否允许用户编辑数据体中的值。|
||[useCustomSortLists](/javascript/api/excel/excel.pivottable#usecustomsortlists)|指定排序时，数据透视表是否使用自定义列表。|
|[PivotTableCollectionLoadOptions](/javascript/api/excel/excel.pivottablecollectionloadoptions)|[enableDataValueEditing](/javascript/api/excel/excel.pivottablecollectionloadoptions#enabledatavalueediting)|对于集合中的每一项: 指定数据透视表是否允许用户对数据正文中的值进行编辑。|
||[useCustomSortLists](/javascript/api/excel/excel.pivottablecollectionloadoptions#usecustomsortlists)|对于集合中的每一项: 指定数据透视表在排序时是否使用自定义列表。|
|[PivotTableData](/javascript/api/excel/excel.pivottabledata)|[enableDataValueEditing](/javascript/api/excel/excel.pivottabledata#enabledatavalueediting)|指定数据透视表是否允许用户编辑数据体中的值。|
||[useCustomSortLists](/javascript/api/excel/excel.pivottabledata#usecustomsortlists)|指定排序时，数据透视表是否使用自定义列表。|
|[PivotTableLoadOptions](/javascript/api/excel/excel.pivottableloadoptions)|[enableDataValueEditing](/javascript/api/excel/excel.pivottableloadoptions#enabledatavalueediting)|指定数据透视表是否允许用户编辑数据体中的值。|
||[useCustomSortLists](/javascript/api/excel/excel.pivottableloadoptions#usecustomsortlists)|指定排序时，数据透视表是否使用自定义列表。|
|[PivotTableUpdateData](/javascript/api/excel/excel.pivottableupdatedata)|[enableDataValueEditing](/javascript/api/excel/excel.pivottableupdatedata#enabledatavalueediting)|指定数据透视表是否允许用户编辑数据体中的值。|
||[useCustomSortLists](/javascript/api/excel/excel.pivottableupdatedata#usecustomsortlists)|指定排序时，数据透视表是否使用自定义列表。|
|[Range](/javascript/api/excel/excel.range)|[autoFill(destinationRange: Range \| string, autoFillType?: "FillDefault" \| "FillCopy" \| "FillSeries" \| "FillFormats" \| "FillValues" \| "FillDays" \| "FillWeekdays" \| "FillMonths" \| "FillYears" \| "LinearTrend" \| "GrowthTrend" \| "FlashFill")](/javascript/api/excel/excel.range#autofill-destinationrange--autofilltype-)|填充区域从当前区域到目标区域。|
||[autoFill(destinationRange: Range \| string, autoFillType?: Excel.AutoFillType)](/javascript/api/excel/excel.range#autofill-destinationrange--autofilltype-)|填充区域从当前区域到目标区域。|
||[convertDataTypeToText()](/javascript/api/excel/excel.range#convertdatatypetotext--)|将具有数据类型的区域单元格转换为文本。|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.range#converttolinkeddatatype-serviceid--languageculture-)|将区域单元格转换为工作表中的链接数据类型。|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: "All" \| "Formulas" \| "Values" \| "Formats", skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-)|将单元格数据或格式从源区域或 RangeAreas 复制到当前区域。|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-)|将单元格数据或格式从源区域或 RangeAreas 复制到当前区域。|
||[find(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#find-text--criteria-)|根据指定的条件查找给定的字符串。|
||[findOrNullObject(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#findornullobject-text--criteria-)|根据指定的条件查找给定的字符串。|
||[flashFill()](/javascript/api/excel/excel.range#flashfill--)|对当前区域进行快速填充。快速填充在感知到模式时可自动填充数据，因此该区域必须是单列区域且周围有数据以便查找模式。|
||[getCellProperties(cellPropertiesLoadOptions: CellPropertiesLoadOptions)](/javascript/api/excel/excel.range#getcellproperties-cellpropertiesloadoptions-)|返回一个 2D 数组，其中封装了每个单元格的字体、填充、边框、对齐方式和其他属性数据。|
||[getColumnProperties(columnPropertiesLoadOptions: ColumnPropertiesLoadOptions)](/javascript/api/excel/excel.range#getcolumnproperties-columnpropertiesloadoptions-)|返回一个一维数组，其中封装了每个列的字体、填充、边框、对齐方式和其他属性数据。  对于给定列中每个单元格不一致的属性，将返回 null。|
||[getRowProperties(rowPropertiesLoadOptions: RowPropertiesLoadOptions)](/javascript/api/excel/excel.range#getrowproperties-rowpropertiesloadoptions-)|返回一个一维数组，其中封装了每个行的字体、填充、边框、对齐方式和其他属性数据。  对于给定行中每个单元格不一致的属性，将返回 null。|
||[getSpecialCells(cellType: "ConditionalFormats" \| "DataValidations" \| "Blanks" \| "Constants" \| "Formulas" \| "SameConditionalFormat" \| "SameDataValidation" \| "Visible", cellValueType?: "All" \| "Errors" \| "ErrorsLogical" \| "ErrorsNumbers" \| "ErrorsText" \| "ErrorsLogicalNumber" \| "ErrorsLogicalText" \| "ErrorsNumberText" \| "Logical" \| "LogicalNumbers" \| "LogicalText" \| "LogicalNumbersText" \| "Numbers" \| "NumbersText" \| "Text")](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-)|获取包含一个或多个矩形区域的 RangeAreas 对象，它表示匹配指定类型和值的所有单元格。|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-)|获取包含一个或多个矩形区域的 RangeAreas 对象，它表示匹配指定类型和值的所有单元格。|
||[getSpecialCellsOrNullObject(cellType: "ConditionalFormats" \| "DataValidations" \| "Blanks" \| "Constants" \| "Formulas" \| "SameConditionalFormat" \| "SameDataValidation" \| "Visible", cellValueType?: "All" \| "Errors" \| "ErrorsLogical" \| "ErrorsNumbers" \| "ErrorsText" \| "ErrorsLogicalNumber" \| "ErrorsLogicalText" \| "ErrorsNumberText" \| "Logical" \| "LogicalNumbers" \| "LogicalText" \| "LogicalNumbersText" \| "Numbers" \| "NumbersText" \| "Text")](/javascript/api/excel/excel.range#getspecialcellsornullobject-celltype--cellvaluetype-)|获取包含一个或多个区域的 RangeAreas 对象，它表示匹配指定类型和值的所有单元格。|
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
||[clear(applyTo?: "All" \| "Formats" \| "Contents" \| "Hyperlinks" \| "RemoveHyperlinks")](/javascript/api/excel/excel.rangeareas#clear-applyto-)|清除包含此 RangeAreas 对象的每个区域的值、格式、填充、边框等。|
||[clear(applyTo?: Excel.ClearApplyTo)](/javascript/api/excel/excel.rangeareas#clear-applyto-)|清除包含此 RangeAreas 对象的每个区域的值、格式、填充、边框等。|
||[convertDataTypeToText()](/javascript/api/excel/excel.rangeareas#convertdatatypetotext--)|将 RangeAreas 中具有数据类型的所有单元格转换为文本。|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.rangeareas#converttolinkeddatatype-serviceid--languageculture-)|将 RangeAreas 中的所有单元格转换为链接数据类型。|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: "All" \| "Formulas" \| "Values" \| "Formats", skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.rangeareas#copyfrom-sourcerange--copytype--skipblanks--transpose-)|将单元格数据或格式从源区域或 RangeAreas 复制到当前 RangeAreas。|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.rangeareas#copyfrom-sourcerange--copytype--skipblanks--transpose-)|将单元格数据或格式从源区域或 RangeAreas 复制到当前 RangeAreas。|
||[getEntireColumn()](/javascript/api/excel/excel.rangeareas#getentirecolumn--)|返回表示 RangeAreas 的整个列的 RangeAreas 对象（例如，如果当前 RangeAreas 表示单元格“B4:E11, H2”，它将返回表示列“B:E, H:H”的 RangeAreas）。|
||[getEntireRow()](/javascript/api/excel/excel.rangeareas#getentirerow--)|获取表示 RangeAreas 的整个行的 RangeAreas 对象（例如，如果当前 RangeAreas 表示单元格“B4:E11”，它将返回表示行“4:11”的 RangeAreas）。|
||[getIntersection(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#getintersection-anotherrange-)|返回 RangeAreas 对象，它表示给定区域或 RangeAreas 的交集。 如果未找到任何交集，则将引发 ItemNotFound 错误。|
||[getIntersectionOrNullObject(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#getintersectionornullobject-anotherrange-)|返回 RangeAreas 对象，它表示给定区域或 RangeAreas 的交集。 如果未找到任何交集，将返回 null 对象。|
||[getOffsetRangeAreas(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.rangeareas#getoffsetrangeareas-rowoffset--columnoffset-)|返回 RangeAreas 对象，它按特定的行和列偏移量进行移动。 返回的 RangeAreas 的维度将与原始对象匹配。 如果生成的 RangeAreas 强行超出工作表网格的边界，则将引发错误。|
||[getSpecialCells(cellType: "ConditionalFormats" \| "DataValidations" \| "Blanks" \| "Constants" \| "Formulas" \| "SameConditionalFormat" \| "SameDataValidation" \| "Visible", cellValueType?: "All" \| "Errors" \| "ErrorsLogical" \| "ErrorsNumbers" \| "ErrorsText" \| "ErrorsLogicalNumber" \| "ErrorsLogicalText" \| "ErrorsNumberText" \| "Logical" \| "LogicalNumbers" \| "LogicalText" \| "LogicalNumbersText" \| "Numbers" \| "NumbersText" \| "Text")](/javascript/api/excel/excel.rangeareas#getspecialcells-celltype--cellvaluetype-)|返回一个 RangeAreas 对象，它表示匹配指定类型和值的所有单元格。 如果未找到符合条件的特殊单元格，则会引发错误。|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#getspecialcells-celltype--cellvaluetype-)|返回一个 RangeAreas 对象，它表示匹配指定类型和值的所有单元格。 如果未找到符合条件的特殊单元格，则会引发错误。|
||[getSpecialCellsOrNullObject(cellType: "ConditionalFormats" \| "DataValidations" \| "Blanks" \| "Constants" \| "Formulas" \| "SameConditionalFormat" \| "SameDataValidation" \| "Visible", cellValueType?: "All" \| "Errors" \| "ErrorsLogical" \| "ErrorsNumbers" \| "ErrorsText" \| "ErrorsLogicalNumber" \| "ErrorsLogicalText" \| "ErrorsNumberText" \| "Logical" \| "LogicalNumbers" \| "LogicalText" \| "LogicalNumbersText" \| "Numbers" \| "NumbersText" \| "Text")](/javascript/api/excel/excel.rangeareas#getspecialcellsornullobject-celltype--cellvaluetype-)|返回一个 RangeAreas 对象，它表示匹配指定类型和值的所有单元格。 如果未找到符合条件的特殊单元格，则返回 null 对象。|
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
||[set (properties: RangeAreas)](/javascript/api/excel/excel.rangeareas#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: RangeAreasUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.rangeareas#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
||[setDirty()](/javascript/api/excel/excel.rangeareas#setdirty--)|设置要在下一次重新计算时重新进行计算的 RangeAreas。|
||[style](/javascript/api/excel/excel.rangeareas#style)|表示此 RangeAreas 对象中的所有区域的样式。|
||[track()](/javascript/api/excel/excel.rangeareas#track--)|根据文档中的相应更改来跟踪对象，以便进行自动调整。 此调用是 context.trackedObjects.add(thisObject) 的缩写。 如果你在“.sync”调用之间和按顺序执行“.run”批处理之外使用此对象，并且在对象上设置属性或调用方法时出现“InvalidObjectPath”错误，则需要在首次创建对象时为跟踪的对象集合添加对象。|
||[untrack()](/javascript/api/excel/excel.rangeareas#untrack--)|释放与此对象关联的内存（如果先前已跟踪过）。 此调用是 context.trackedObjects.add(thisObject) 的缩写。 拥有许多跟踪对象会降低主机应用程序的速度，因此请在使用完毕后释放所添加的任何对象。 在内存释放生效之前，你需要调用“context.sync()”。|
|[RangeAreasData](/javascript/api/excel/excel.rangeareasdata)|[address](/javascript/api/excel/excel.rangeareasdata#address)|返回 A1 样式中的 RageAreas 引用。 地址值将包含单元格的每个矩形块的工作表名称（例如“Sheet1!A1:B4, Sheet1!D1:D4”）。 只读。|
||[addressLocal](/javascript/api/excel/excel.rangeareasdata#addresslocal)|返回用户区域设置中的 RageAreas 引用。 只读。|
||[areaCount](/javascript/api/excel/excel.rangeareasdata#areacount)|返回包含此 RangeAreas 对象的矩形区域的数量。|
||[areas](/javascript/api/excel/excel.rangeareasdata#areas)|返回包含此 RangeAreas 对象的矩形区域的集合。|
||[cellCount](/javascript/api/excel/excel.rangeareasdata#cellcount)|返回 RangeAreas 对象中的单元格数量，即总计各个矩形区域的单元格计数。 如果单元格计数超过 2^31-1 (2,147,483,647)，则返回 -1。 只读。|
||[conditionalFormats](/javascript/api/excel/excel.rangeareasdata#conditionalformats)|返回与此 RangeAreas 对象中的任何单元格相交的 ConditionalFormats 集合。 只读。|
||[dataValidation](/javascript/api/excel/excel.rangeareasdata#datavalidation)|返回 RangeAreas 中的所有区域的 dataValidation 对象。|
||[format](/javascript/api/excel/excel.rangeareasdata#format)|返回一个 rangeFormat 对象，其中封装了 RangeAreas 对象中的所有区域的字体、填充、边框、对齐方式和其他属性。 只读。|
||[isEntireColumn](/javascript/api/excel/excel.rangeareasdata#isentirecolumn)|指示此 RangeAreas 对象上的所有区域是否表示整列（例如“A:C, Q:Z”）。 只读。|
||[isEntireRow](/javascript/api/excel/excel.rangeareasdata#isentirerow)|指示此 RangeAreas 对象上的所有区域是否表示整行（例如“1:3, 5:7”）。 只读。|
||[style](/javascript/api/excel/excel.rangeareasdata#style)|表示此 RangeAreas 对象中的所有区域的样式。|
|[RangeAreasLoadOptions](/javascript/api/excel/excel.rangeareasloadoptions)|[$all](/javascript/api/excel/excel.rangeareasloadoptions#$all)||
||[address](/javascript/api/excel/excel.rangeareasloadoptions#address)|返回 A1 样式中的 RageAreas 引用。 地址值将包含单元格的每个矩形块的工作表名称（例如“Sheet1!A1:B4, Sheet1!D1:D4”）。 只读。|
||[addressLocal](/javascript/api/excel/excel.rangeareasloadoptions#addresslocal)|返回用户区域设置中的 RageAreas 引用。 只读。|
||[areaCount](/javascript/api/excel/excel.rangeareasloadoptions#areacount)|返回包含此 RangeAreas 对象的矩形区域的数量。|
||[cellCount](/javascript/api/excel/excel.rangeareasloadoptions#cellcount)|返回 RangeAreas 对象中的单元格数量，即总计各个矩形区域的单元格计数。 如果单元格计数超过 2^31-1 (2,147,483,647)，则返回 -1。 只读。|
||[dataValidation](/javascript/api/excel/excel.rangeareasloadoptions#datavalidation)|返回 RangeAreas 中的所有区域的 dataValidation 对象。|
||[format](/javascript/api/excel/excel.rangeareasloadoptions#format)|返回一个 rangeFormat 对象，其中封装了 RangeAreas 对象中的所有区域的字体、填充、边框、对齐方式和其他属性。|
||[isEntireColumn](/javascript/api/excel/excel.rangeareasloadoptions#isentirecolumn)|指示此 RangeAreas 对象上的所有区域是否表示整列（例如“A:C, Q:Z”）。 只读。|
||[isEntireRow](/javascript/api/excel/excel.rangeareasloadoptions#isentirerow)|指示此 RangeAreas 对象上的所有区域是否表示整行（例如“1:3, 5:7”）。 只读。|
||[style](/javascript/api/excel/excel.rangeareasloadoptions#style)|表示此 RangeAreas 对象中的所有区域的样式。|
||[worksheet](/javascript/api/excel/excel.rangeareasloadoptions#worksheet)|返回当前 RangeAreas 的工作表。|
|[RangeAreasUpdateData](/javascript/api/excel/excel.rangeareasupdatedata)|[dataValidation](/javascript/api/excel/excel.rangeareasupdatedata#datavalidation)|返回 RangeAreas 中的所有区域的 dataValidation 对象。|
||[format](/javascript/api/excel/excel.rangeareasupdatedata#format)|返回一个 rangeFormat 对象，其中封装了 RangeAreas 对象中的所有区域的字体、填充、边框、对齐方式和其他属性。|
||[style](/javascript/api/excel/excel.rangeareasupdatedata#style)|表示此 RangeAreas 对象中的所有区域的样式。|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[tintAndShade](/javascript/api/excel/excel.rangeborder#tintandshade)|返回或设置一个使区域边框的颜色变亮或变暗的双精度数值，该值介于 -1（最暗）与 1（最亮）之间，初始颜色为 0。|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[tintAndShade](/javascript/api/excel/excel.rangebordercollection#tintandshade)|返回或设置一个使区域边框的颜色变亮或变暗的双精度数值，该值介于 -1（最暗）与 1（最亮）之间，初始颜色为 0。|
|[RangeBorderCollectionLoadOptions](/javascript/api/excel/excel.rangebordercollectionloadoptions)|[tintAndShade](/javascript/api/excel/excel.rangebordercollectionloadoptions#tintandshade)|对于集合中的每一项: 返回或设置一个双精度值, 该值为区域边框的颜色变浅或变暗, 值介于-1 (最暗) 和 1 (最亮) 之间, 原始颜色为0。|
|[RangeBorderCollectionUpdateData](/javascript/api/excel/excel.rangebordercollectionupdatedata)|[tintAndShade](/javascript/api/excel/excel.rangebordercollectionupdatedata#tintandshade)|返回或设置一个使区域边框的颜色变亮或变暗的双精度数值，该值介于 -1（最暗）与 1（最亮）之间，初始颜色为 0。|
|[RangeBorderData](/javascript/api/excel/excel.rangeborderdata)|[tintAndShade](/javascript/api/excel/excel.rangeborderdata#tintandshade)|返回或设置一个使区域边框的颜色变亮或变暗的双精度数值，该值介于 -1（最暗）与 1（最亮）之间，初始颜色为 0。|
|[RangeBorderLoadOptions](/javascript/api/excel/excel.rangeborderloadoptions)|[tintAndShade](/javascript/api/excel/excel.rangeborderloadoptions#tintandshade)|返回或设置一个使区域边框的颜色变亮或变暗的双精度数值，该值介于 -1（最暗）与 1（最亮）之间，初始颜色为 0。|
|[RangeBorderUpdateData](/javascript/api/excel/excel.rangeborderupdatedata)|[tintAndShade](/javascript/api/excel/excel.rangeborderupdatedata#tintandshade)|返回或设置一个使区域边框的颜色变亮或变暗的双精度数值，该值介于 -1（最暗）与 1（最亮）之间，初始颜色为 0。|
|[RangeCollection](/javascript/api/excel/excel.rangecollection)|[getCount()](/javascript/api/excel/excel.rangecollection#getcount--)|返回 RangeCollection 中的区域数量。|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangecollection#getitemat-index-)|根据其在 RangeCollection 中的位置返回 Range 对象。|
||[items](/javascript/api/excel/excel.rangecollection#items)|获取此集合中已加载的子项。|
|[RangeCollectionLoadOptions](/javascript/api/excel/excel.rangecollectionloadoptions)|[$all](/javascript/api/excel/excel.rangecollectionloadoptions#$all)||
||[address](/javascript/api/excel/excel.rangecollectionloadoptions#address)|对于集合中的每一项: 代表 A1 样式中的区域引用。 Address 值将包含工作表引用 (例如, "Sheet1!A1: B4 ")。 只读。|
||[addressLocal](/javascript/api/excel/excel.rangecollectionloadoptions#addresslocal)|对于集合中的每一项: 代表指定区域的区域引用 (以用户语言表示)。 只读。|
||[cellCount](/javascript/api/excel/excel.rangecollectionloadoptions#cellcount)|对于集合中的每个项目: 区域中的单元格数量。 如果单元格数超过 2^31-1 (2,147,483,647)，此 API 返回 -1。 只读。|
||[columnCount](/javascript/api/excel/excel.rangecollectionloadoptions#columncount)|对于集合中的每一项: 表示区域中的总列数。 只读。|
||[columnHidden](/javascript/api/excel/excel.rangecollectionloadoptions#columnhidden)|对于集合中的每一项: 表示是否隐藏当前区域中的所有列。|
||[columnIndex](/javascript/api/excel/excel.rangecollectionloadoptions#columnindex)|对于集合中的每一项: 表示区域中的第一个单元格的列号。 从零开始编制索引。 只读。|
||[dataValidation](/javascript/api/excel/excel.rangecollectionloadoptions#datavalidation)|对于集合中的每一项: 返回一个数据验证对象。|
||[format](/javascript/api/excel/excel.rangecollectionloadoptions#format)|对于集合中的每一项: 返回一个格式对象, 封装区域的字体、填充、边框、对齐方式和其他属性。|
||[formulas](/javascript/api/excel/excel.rangecollectionloadoptions#formulas)|对于集合中的每一项: 代表 A1 样式表示法中的公式。|
||[formulasLocal](/javascript/api/excel/excel.rangecollectionloadoptions#formulaslocal)|对于集合中的每一项: 代表 A1 样式表示法中的公式, 位于用户的语言和数字格式设置区域中。  例如，英语中的公式 "=SUM(A1, 1.5)" 在德语中将变为 "=SUMME(A1; 1,5)"。|
||[formulasR1C1](/javascript/api/excel/excel.rangecollectionloadoptions#formulasr1c1)|对于集合中的每一项: 以 R1C1 样式表示法表示的公式。|
||[hidden](/javascript/api/excel/excel.rangecollectionloadoptions#hidden)|对于集合中的每一项: 表示是否隐藏当前区域中的所有单元格。 只读。|
||[hyperlink](/javascript/api/excel/excel.rangecollectionloadoptions#hyperlink)|对于集合中的每一项: 代表当前区域的超链接。|
||[isEntireColumn](/javascript/api/excel/excel.rangecollectionloadoptions#isentirecolumn)|对于集合中的每一项: 表示当前区域是否为整列。 只读。|
||[isEntireRow](/javascript/api/excel/excel.rangecollectionloadoptions#isentirerow)|对于集合中的每一项: 表示当前区域是否为整行。 只读。|
||[linkedDataTypeState](/javascript/api/excel/excel.rangecollectionloadoptions#linkeddatatypestate)|对于集合中的每一项: 表示每个单元格的数据类型状态。 只读。|
||[numberFormat](/javascript/api/excel/excel.rangecollectionloadoptions#numberformat)|对于集合中的每一项: 表示给定范围的 Excel 的编号格式代码。|
||[numberFormatLocal](/javascript/api/excel/excel.rangecollectionloadoptions#numberformatlocal)|对于集合中的每一项: 以用户语言的字符串形式表示给定范围的 Excel 数字格式代码。|
||[rowCount](/javascript/api/excel/excel.rangecollectionloadoptions#rowcount)|对于集合中的每一项: 返回范围中的总行数。 只读。|
||[rowHidden](/javascript/api/excel/excel.rangecollectionloadoptions#rowhidden)|对于集合中的每一项: 表示是否隐藏当前区域中的所有行。|
||[rowIndex](/javascript/api/excel/excel.rangecollectionloadoptions#rowindex)|对于集合中的每一项: 返回区域中第一个单元格的行号。 从零开始编制索引。 只读。|
||[style](/javascript/api/excel/excel.rangecollectionloadoptions#style)|对于集合中的每一项: 代表当前区域的样式。|
||[text](/javascript/api/excel/excel.rangecollectionloadoptions#text)|对于集合中的每一项: 指定区域的文本值。 文本值与单元格宽度无关。 在 Excel UI 中替代 # 符号不会影响 API 返回的文本值。 只读。|
||[valueTypes](/javascript/api/excel/excel.rangecollectionloadoptions#valuetypes)|对于集合中的每一项: 代表每个单元格的数据类型。 只读。|
||[values](/javascript/api/excel/excel.rangecollectionloadoptions#values)|对于集合中的每一项: 代表指定区域的原始值。 返回的数据可能是字符串、数字，也可能是布尔值。 包含错误的单元格将返回错误字符串。|
||[worksheet](/javascript/api/excel/excel.rangecollectionloadoptions#worksheet)|对于集合中的每一项: 包含当前区域的工作表。|
|[RangeData](/javascript/api/excel/excel.rangedata)|[linkedDataTypeState](/javascript/api/excel/excel.rangedata#linkeddatatypestate)|表示每个单元格的数据类型状态。 只读。|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[pattern](/javascript/api/excel/excel.rangefill#pattern)|获取或设置区域的图案。 有关详细信息，请参阅 Excel.FillPattern。 不支持 LinearGradient 和 RectangularGradient。|
||[patternColor](/javascript/api/excel/excel.rangefill#patterncolor)|设置 HTML 颜色代码，它表示窗体 #RRGGBB（例如“FFA500”）的区域图案颜色或作为已命名的 HTML 颜色（例如“orange”）。|
||[patternTintAndShade](/javascript/api/excel/excel.rangefill#patterntintandshade)|返回或设置一个使区域填充的图案颜色变亮或变暗的双精度数值，该值介于 -1（最暗）与 1（最亮）之间，初始颜色为 0。|
||[tintAndShade](/javascript/api/excel/excel.rangefill#tintandshade)|返回或设置一个使区域填充的颜色变亮或变暗的双精度数值，该值介于 -1（最暗）与 1（最亮）之间，初始颜色为 0。|
|[RangeFillData](/javascript/api/excel/excel.rangefilldata)|[pattern](/javascript/api/excel/excel.rangefilldata#pattern)|获取或设置区域的图案。 有关详细信息，请参阅 Excel.FillPattern。 不支持 LinearGradient 和 RectangularGradient。|
||[patternColor](/javascript/api/excel/excel.rangefilldata#patterncolor)|设置 HTML 颜色代码，它表示窗体 #RRGGBB（例如“FFA500”）的区域图案颜色或作为已命名的 HTML 颜色（例如“orange”）。|
||[patternTintAndShade](/javascript/api/excel/excel.rangefilldata#patterntintandshade)|返回或设置一个使区域填充的图案颜色变亮或变暗的双精度数值，该值介于 -1（最暗）与 1（最亮）之间，初始颜色为 0。|
||[tintAndShade](/javascript/api/excel/excel.rangefilldata#tintandshade)|返回或设置一个使区域填充的颜色变亮或变暗的双精度数值，该值介于 -1（最暗）与 1（最亮）之间，初始颜色为 0。|
|[RangeFillLoadOptions](/javascript/api/excel/excel.rangefillloadoptions)|[pattern](/javascript/api/excel/excel.rangefillloadoptions#pattern)|获取或设置区域的图案。 有关详细信息，请参阅 Excel.FillPattern。 不支持 LinearGradient 和 RectangularGradient。|
||[patternColor](/javascript/api/excel/excel.rangefillloadoptions#patterncolor)|设置 HTML 颜色代码，它表示窗体 #RRGGBB（例如“FFA500”）的区域图案颜色或作为已命名的 HTML 颜色（例如“orange”）。|
||[patternTintAndShade](/javascript/api/excel/excel.rangefillloadoptions#patterntintandshade)|返回或设置一个使区域填充的图案颜色变亮或变暗的双精度数值，该值介于 -1（最暗）与 1（最亮）之间，初始颜色为 0。|
||[tintAndShade](/javascript/api/excel/excel.rangefillloadoptions#tintandshade)|返回或设置一个使区域填充的颜色变亮或变暗的双精度数值，该值介于 -1（最暗）与 1（最亮）之间，初始颜色为 0。|
|[RangeFillUpdateData](/javascript/api/excel/excel.rangefillupdatedata)|[pattern](/javascript/api/excel/excel.rangefillupdatedata#pattern)|获取或设置区域的图案。 有关详细信息，请参阅 Excel.FillPattern。 不支持 LinearGradient 和 RectangularGradient。|
||[patternColor](/javascript/api/excel/excel.rangefillupdatedata#patterncolor)|设置 HTML 颜色代码，它表示窗体 #RRGGBB（例如“FFA500”）的区域图案颜色或作为已命名的 HTML 颜色（例如“orange”）。|
||[patternTintAndShade](/javascript/api/excel/excel.rangefillupdatedata#patterntintandshade)|返回或设置一个使区域填充的图案颜色变亮或变暗的双精度数值，该值介于 -1（最暗）与 1（最亮）之间，初始颜色为 0。|
||[tintAndShade](/javascript/api/excel/excel.rangefillupdatedata#tintandshade)|返回或设置一个使区域填充的颜色变亮或变暗的双精度数值，该值介于 -1（最暗）与 1（最亮）之间，初始颜色为 0。|
|[RangeFont](/javascript/api/excel/excel.rangefont)|[strikethrough](/javascript/api/excel/excel.rangefont#strikethrough)|表示字体的删除线状态。 Null 值表示整个区域没有统一的删除线设置。|
||[subscript](/javascript/api/excel/excel.rangefont#subscript)|表示字体的下标状态。|
||[superscript](/javascript/api/excel/excel.rangefont#superscript)|表示字体的上标状态。|
||[tintAndShade](/javascript/api/excel/excel.rangefont#tintandshade)|返回或设置一个使区域字体的颜色变亮或变暗的双精度数值，该值介于 -1（最暗）与 1（最亮）之间，初始颜色为 0。|
|[RangeFontData](/javascript/api/excel/excel.rangefontdata)|[strikethrough](/javascript/api/excel/excel.rangefontdata#strikethrough)|表示字体的删除线状态。 Null 值表示整个区域没有统一的删除线设置。|
||[subscript](/javascript/api/excel/excel.rangefontdata#subscript)|表示字体的下标状态。|
||[superscript](/javascript/api/excel/excel.rangefontdata#superscript)|表示字体的上标状态。|
||[tintAndShade](/javascript/api/excel/excel.rangefontdata#tintandshade)|返回或设置一个使区域字体的颜色变亮或变暗的双精度数值，该值介于 -1（最暗）与 1（最亮）之间，初始颜色为 0。|
|[RangeFontLoadOptions](/javascript/api/excel/excel.rangefontloadoptions)|[strikethrough](/javascript/api/excel/excel.rangefontloadoptions#strikethrough)|表示字体的删除线状态。 Null 值表示整个区域没有统一的删除线设置。|
||[subscript](/javascript/api/excel/excel.rangefontloadoptions#subscript)|表示字体的下标状态。|
||[superscript](/javascript/api/excel/excel.rangefontloadoptions#superscript)|表示字体的上标状态。|
||[tintAndShade](/javascript/api/excel/excel.rangefontloadoptions#tintandshade)|返回或设置一个使区域字体的颜色变亮或变暗的双精度数值，该值介于 -1（最暗）与 1（最亮）之间，初始颜色为 0。|
|[RangeFontUpdateData](/javascript/api/excel/excel.rangefontupdatedata)|[strikethrough](/javascript/api/excel/excel.rangefontupdatedata#strikethrough)|表示字体的删除线状态。 Null 值表示整个区域没有统一的删除线设置。|
||[subscript](/javascript/api/excel/excel.rangefontupdatedata#subscript)|表示字体的下标状态。|
||[superscript](/javascript/api/excel/excel.rangefontupdatedata#superscript)|表示字体的上标状态。|
||[tintAndShade](/javascript/api/excel/excel.rangefontupdatedata#tintandshade)|返回或设置一个使区域字体的颜色变亮或变暗的双精度数值，该值介于 -1（最暗）与 1（最亮）之间，初始颜色为 0。|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[autoIndent](/javascript/api/excel/excel.rangeformat#autoindent)|指示将文本对齐方式设为相等分布时文本是否会自动缩进。|
||[indentLevel](/javascript/api/excel/excel.rangeformat#indentlevel)|0 到 250 之间的一个整数，指示缩进水平。|
||[readingOrder](/javascript/api/excel/excel.rangeformat#readingorder)|区域的读取顺序。|
||[shrinkToFit](/javascript/api/excel/excel.rangeformat#shrinktofit)|指示文本是否自动缩小以适合可用列宽。|
|[RangeFormatData](/javascript/api/excel/excel.rangeformatdata)|[autoIndent](/javascript/api/excel/excel.rangeformatdata#autoindent)|指示将文本对齐方式设为相等分布时文本是否会自动缩进。|
||[indentLevel](/javascript/api/excel/excel.rangeformatdata#indentlevel)|0 到 250 之间的一个整数，指示缩进水平。|
||[readingOrder](/javascript/api/excel/excel.rangeformatdata#readingorder)|区域的读取顺序。|
||[shrinkToFit](/javascript/api/excel/excel.rangeformatdata#shrinktofit)|指示文本是否自动缩小以适合可用列宽。|
|[RangeFormatLoadOptions](/javascript/api/excel/excel.rangeformatloadoptions)|[autoIndent](/javascript/api/excel/excel.rangeformatloadoptions#autoindent)|指示将文本对齐方式设为相等分布时文本是否会自动缩进。|
||[indentLevel](/javascript/api/excel/excel.rangeformatloadoptions#indentlevel)|0 到 250 之间的一个整数，指示缩进水平。|
||[readingOrder](/javascript/api/excel/excel.rangeformatloadoptions#readingorder)|区域的读取顺序。|
||[shrinkToFit](/javascript/api/excel/excel.rangeformatloadoptions#shrinktofit)|指示文本是否自动缩小以适合可用列宽。|
|[RangeFormatUpdateData](/javascript/api/excel/excel.rangeformatupdatedata)|[autoIndent](/javascript/api/excel/excel.rangeformatupdatedata#autoindent)|指示将文本对齐方式设为相等分布时文本是否会自动缩进。|
||[indentLevel](/javascript/api/excel/excel.rangeformatupdatedata#indentlevel)|0 到 250 之间的一个整数，指示缩进水平。|
||[readingOrder](/javascript/api/excel/excel.rangeformatupdatedata#readingorder)|区域的读取顺序。|
||[shrinkToFit](/javascript/api/excel/excel.rangeformatupdatedata#shrinktofit)|指示文本是否自动缩小以适合可用列宽。|
|[RangeLoadOptions](/javascript/api/excel/excel.rangeloadoptions)|[linkedDataTypeState](/javascript/api/excel/excel.rangeloadoptions#linkeddatatypestate)|表示每个单元格的数据类型状态。 只读。|
|[RemoveDuplicatesResult](/javascript/api/excel/excel.removeduplicatesresult)|[removed](/javascript/api/excel/excel.removeduplicatesresult#removed)|由操作删除的重复行数。|
||[uniqueRemaining](/javascript/api/excel/excel.removeduplicatesresult#uniqueremaining)|所生成的区域中存在的剩余唯一行数。|
|[RemoveDuplicatesResultData](/javascript/api/excel/excel.removeduplicatesresultdata)|[removed](/javascript/api/excel/excel.removeduplicatesresultdata#removed)|由操作删除的重复行数。|
||[uniqueRemaining](/javascript/api/excel/excel.removeduplicatesresultdata#uniqueremaining)|所生成的区域中存在的剩余唯一行数。|
|[RemoveDuplicatesResultLoadOptions](/javascript/api/excel/excel.removeduplicatesresultloadoptions)|[$all](/javascript/api/excel/excel.removeduplicatesresultloadoptions#$all)||
||[removed](/javascript/api/excel/excel.removeduplicatesresultloadoptions#removed)|由操作删除的重复行数。|
||[uniqueRemaining](/javascript/api/excel/excel.removeduplicatesresultloadoptions#uniqueremaining)|所生成的区域中存在的剩余唯一行数。|
|[ReplaceCriteria](/javascript/api/excel/excel.replacecriteria)|[completeMatch](/javascript/api/excel/excel.replacecriteria#completematch)|指定匹配必须是完整匹配还是部分匹配。 默认值为 false（部分）。|
||[matchCase](/javascript/api/excel/excel.replacecriteria#matchcase)|指定匹配是否区分大小写。 默认值为 false（不区分大小写）。|
|[RowProperties](/javascript/api/excel/excel.rowproperties)|[address](/javascript/api/excel/excel.rowproperties#address)|表示`address`属性。|
||[addressLocal](/javascript/api/excel/excel.rowproperties#addresslocal)|表示`addressLocal`属性。|
||[rowIndex](/javascript/api/excel/excel.rowproperties#rowindex)|表示`rowIndex`属性。|
|[RowPropertiesLoadOptions](/javascript/api/excel/excel.rowpropertiesloadoptions)|[格式: CellPropertiesFormatLoadOptions & {
            rowHeight？](/javascript/api/excel/excel.rowpropertiesloadoptions # format)|指定是否在`format`属性上进行加载。|
||[rowHeight](/javascript/api/excel/excel.rowpropertiesloadoptions#rowheight)||
||[rowHidden](/javascript/api/excel/excel.rowpropertiesloadoptions#rowhidden)|指定是否在`rowHidden`属性上进行加载。|
||[rowIndex](/javascript/api/excel/excel.rowpropertiesloadoptions#rowindex)|指定是否在`rowIndex`属性上进行加载。|
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
||[getAsImage(format: "UNKNOWN" \| "BMP" \| "JPEG" \| "GIF" \| "PNG" \| "SVG")](/javascript/api/excel/excel.shape#getasimage-format-)|将形状转换为图像并将图像返回为 base64 编码字符串。 DPI 为 96。 仅支持格式 `Excel.PictureFormat.BMP`、`Excel.PictureFormat.PNG`、`Excel.PictureFormat.JPEG` 和 `Excel.PictureFormat.GIF`。|
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
||[scaleHeight(scaleFactor: number, scaleType: "CurrentSize" \| "OriginalSize", scaleFrom?: "ScaleFromTopLeft" \| "ScaleFromMiddle" \| "ScaleFromBottomRight")](/javascript/api/excel/excel.shape#scaleheight-scalefactor--scaletype--scalefrom-)|按指定因子缩放形状的高度。 对于图像，你可以说明是相对于原始尺寸还是当前尺寸缩放形状。 对于除图片以外的其他形状来说，缩放总是相对于其当前高度而言。|
||[scaleHeight(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#scaleheight-scalefactor--scaletype--scalefrom-)|按指定因子缩放形状的高度。 对于图像，你可以说明是相对于原始尺寸还是当前尺寸缩放形状。 对于除图片以外的其他形状来说，缩放总是相对于其当前高度而言。|
||[scaleWidth(scaleFactor: number, scaleType: "CurrentSize" \| "OriginalSize", scaleFrom?: "ScaleFromTopLeft" \| "ScaleFromMiddle" \| "ScaleFromBottomRight")](/javascript/api/excel/excel.shape#scalewidth-scalefactor--scaletype--scalefrom-)|按指定因子缩放形状的宽度。 对于图像，你可以说明是相对于原始尺寸还是当前尺寸缩放形状。 对于除图片以外的其他形状来说，缩放总是相对于其当前宽度而言。|
||[scaleWidth(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#scalewidth-scalefactor--scaletype--scalefrom-)|按指定因子缩放形状的宽度。 对于图像，你可以说明是相对于原始尺寸还是当前尺寸缩放形状。 对于除图片以外的其他形状来说，缩放总是相对于其当前宽度而言。|
||[set (properties: Excel. Shape)](/javascript/api/excel/excel.shape#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: ShapeUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.shape#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
||[setZOrder(position: "BringToFront" \| "BringForward" \| "SendToBack" \| "SendBackward")](/javascript/api/excel/excel.shape#setzorder-position-)|将指定形状沿集合的 z 顺序向上或向下移动，将其移动到其他形状的前面或后面。|
||[setZOrder(position: Excel.ShapeZOrder)](/javascript/api/excel/excel.shape#setzorder-position-)|将指定形状沿集合的 z 顺序向上或向下移动，将其移动到其他形状的前面或后面。|
||[top](/javascript/api/excel/excel.shape#top)|从形状上边缘到工作表上边缘之间的距离（以磅为单位）。|
||[visible](/javascript/api/excel/excel.shape#visible)|表示此形状的可视性。|
||[width](/javascript/api/excel/excel.shape#width)|表示形状的宽度（以磅为单位）。|
|[ShapeActivatedEventArgs](/javascript/api/excel/excel.shapeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapeactivatedeventargs#shapeid)|获取已激活的形状的 ID。|
||[type](/javascript/api/excel/excel.shapeactivatedeventargs#type)|获取事件的类型。 有关详细信息，请参阅 Excel.EventType。|
||[worksheetId](/javascript/api/excel/excel.shapeactivatedeventargs#worksheetid)|获取其中的形状已启用的工作表的 ID。|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addGeometricShape(geometricShapeType: "LineInverse" \| "Triangle" \| "RightTriangle" \| "Rectangle" \| "Diamond" \| "Parallelogram" \| "Trapezoid" \| "NonIsoscelesTrapezoid" \| "Pentagon" \| "Hexagon" \| "Heptagon" \| "Octagon" \| "Decagon" \| "Dodecagon" \| "Star4" \| "Star5" \| "Star6" \| "Star7" \| "Star8" \| "Star10" \| "Star12" \| "Star16" \| "Star24" \| "Star32" \| "RoundRectangle" \| "Round1Rectangle" \| "Round2SameRectangle" \| "Round2DiagonalRectangle" \| "SnipRoundRectangle" \| "Snip1Rectangle" \| "Snip2SameRectangle" \| "Snip2DiagonalRectangle" \| "Plaque" \| "Ellipse" \| "Teardrop" \| "HomePlate" \| "Chevron" \| "PieWedge" \| "Pie" \| "BlockArc" \| "Donut" \| "NoSmoking" \| "RightArrow" \| "LeftArrow" \| "UpArrow" \| "DownArrow" \| "StripedRightArrow" \| "NotchedRightArrow" \| "BentUpArrow" \| "LeftRightArrow" \| "UpDownArrow" \| "LeftUpArrow" \| "LeftRightUpArrow" \| "QuadArrow" \| "LeftArrowCallout" \| "RightArrowCallout" \| "UpArrowCallout" \| "DownArrowCallout" \| "LeftRightArrowCallout" \| "UpDownArrowCallout" \| "QuadArrowCallout" \| "BentArrow" \| "UturnArrow" \| "CircularArrow" \| "LeftCircularArrow" \| "LeftRightCircularArrow" \| "CurvedRightArrow" \| "CurvedLeftArrow" \| "CurvedUpArrow" \| "CurvedDownArrow" \| "SwooshArrow" \| "Cube" \| "Can" \| "LightningBolt" \| "Heart" \| "Sun" \| "Moon" \| "SmileyFace" \| "IrregularSeal1" \| "IrregularSeal2" \| "FoldedCorner" \| "Bevel" \| "Frame" \| "HalfFrame" \| "Corner" \| "DiagonalStripe" \| "Chord" \| "Arc" \| "LeftBracket" \| "RightBracket" \| "LeftBrace" \| "RightBrace" \| "BracketPair" \| "BracePair" \| "Callout1" \| "Callout2" \| "Callout3" \| "AccentCallout1" \| "AccentCallout2" \| "AccentCallout3" \| "BorderCallout1" \| "BorderCallout2" \| "BorderCallout3" \| "AccentBorderCallout1" \| "AccentBorderCallout2" \| "AccentBorderCallout3" \| "WedgeRectCallout" \| "WedgeRRectCallout" \| "WedgeEllipseCallout" \| "CloudCallout" \| "Cloud" \| "Ribbon" \| "Ribbon2" \| "EllipseRibbon" \| "EllipseRibbon2" \| "LeftRightRibbon" \| "VerticalScroll" \| "HorizontalScroll" \| "Wave" \| "DoubleWave" \| "Plus" \| "FlowChartProcess" \| "FlowChartDecision" \| "FlowChartInputOutput" \| "FlowChartPredefinedProcess" \| "FlowChartInternalStorage" \| "FlowChartDocument" \| "FlowChartMultidocument" \| "FlowChartTerminator" \| "FlowChartPreparation" \| "FlowChartManualInput" \| "FlowChartManualOperation" \| "FlowChartConnector" \| "FlowChartPunchedCard" \| "FlowChartPunchedTape" \| "FlowChartSummingJunction" \| "FlowChartOr" \| "FlowChartCollate" \| "FlowChartSort" \| "FlowChartExtract" \| "FlowChartMerge" \| "FlowChartOfflineStorage" \| "FlowChartOnlineStorage" \| "FlowChartMagneticTape" \| "FlowChartMagneticDisk" \| "FlowChartMagneticDrum" \| "FlowChartDisplay" \| "FlowChartDelay" \| "FlowChartAlternateProcess" \| "FlowChartOffpageConnector" \| "ActionButtonBlank" \| "ActionButtonHome" \| "ActionButtonHelp" \| "ActionButtonInformation" \| "ActionButtonForwardNext" \| "ActionButtonBackPrevious" \| "ActionButtonEnd" \| "ActionButtonBeginning" \| "ActionButtonReturn" \| "ActionButtonDocument" \| "ActionButtonSound" \| "ActionButtonMovie" \| "Gear6" \| "Gear9" \| "Funnel" \| "MathPlus" \| "MathMinus" \| "MathMultiply" \| "MathDivide" \| "MathEqual" \| "MathNotEqual" \| "CornerTabs" \| "SquareTabs" \| "PlaqueTabs" \| "ChartX" \| "ChartStar" \| "ChartPlus")](/javascript/api/excel/excel.shapecollection#addgeometricshape-geometricshapetype-)|将几何形状添加到工作表。 返回一个 Shape 对象，该对象代表新图形。|
||[addGeometricShape(geometricShapeType: Excel.GeometricShapeType)](/javascript/api/excel/excel.shapecollection#addgeometricshape-geometricshapetype-)|将几何形状添加到工作表。 返回一个 Shape 对象，该对象代表新图形。|
||[addGroup(values: Array<string \| Shape>)](/javascript/api/excel/excel.shapecollection#addgroup-values-)|在此集合的工作表中对形状的子集进行分组。 返回表示新形状组的 Shape 对象。|
||[addImage(base64ImageString: string)](/javascript/api/excel/excel.shapecollection#addimage-base64imagestring-)|从 base64 编码的字符串创建图像并将其添加到工作表。 返回表示新图片的 Shape 对象。|
||[addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: "Straight" \| "Elbow" \| "Curve")](/javascript/api/excel/excel.shapecollection#addline-startleft--starttop--endleft--endtop--connectortype-)|将线条添加到工作表。 返回表示新线条的 Shape 对象。|
||[addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType)](/javascript/api/excel/excel.shapecollection#addline-startleft--starttop--endleft--endtop--connectortype-)|将线条添加到工作表。 返回表示新线条的 Shape 对象。|
||[addTextBox(text?: string)](/javascript/api/excel/excel.shapecollection#addtextbox-text-)|使用提供的文本作为内容，将文本框添加到工作表。 返回表示新文本框的 Shape 对象。|
||[getCount()](/javascript/api/excel/excel.shapecollection#getcount--)|返回工作表中的形状数。 只读。|
||[getItem(key: string)](/javascript/api/excel/excel.shapecollection#getitem-key-)|按名称或 ID 获取形状。|
||[getItemAt(index: number)](/javascript/api/excel/excel.shapecollection#getitemat-index-)|使用其在集合中的位置获取形状。|
||[items](/javascript/api/excel/excel.shapecollection#items)|获取此集合中已加载的子项。|
|[ShapeCollectionLoadOptions](/javascript/api/excel/excel.shapecollectionloadoptions)|[$all](/javascript/api/excel/excel.shapecollectionloadoptions#$all)||
||[altTextDescription](/javascript/api/excel/excel.shapecollectionloadoptions#alttextdescription)|对于集合中的每一项: 返回或设置 Shape 对象的替代说明文本。|
||[altTextTitle](/javascript/api/excel/excel.shapecollectionloadoptions#alttexttitle)|对于集合中的每一项: 返回或设置 Shape 对象的可选标题文本。|
||[connectionSiteCount](/javascript/api/excel/excel.shapecollectionloadoptions#connectionsitecount)|对于集合中的每个项目: 返回此形状上的连接结点的数目。 只读。|
||[fill](/javascript/api/excel/excel.shapecollectionloadoptions#fill)|对于集合中的每一项: 返回此形状的填充格式。|
||[geometricShape](/javascript/api/excel/excel.shapecollectionloadoptions#geometricshape)|对于集合中的每个项目: 返回与形状相关联的几何形状。 如果形状类型不是“GeometricShape”，则会引发错误。|
||[geometricShapeType](/javascript/api/excel/excel.shapecollectionloadoptions#geometricshapetype)|对于集合中的每一项: 表示此几何形状的几何形状类型。 有关详细信息，请参阅 Excel.GeometricShapeType。 如果形状类型不是“GeometricShape”，返回 NULL。|
||[组](/javascript/api/excel/excel.shapecollectionloadoptions#group)|对于集合中的每个项目: 返回与形状相关联的形状组。 如果形状类型不是“GroupShape”，则会引发错误。|
||[height](/javascript/api/excel/excel.shapecollectionloadoptions#height)|对于集合中的每个项目: 表示形状的高度 (以磅为单位)。|
||[id](/javascript/api/excel/excel.shapecollectionloadoptions#id)|对于集合中的每一项: 代表形状标识符。 只读。|
||[image](/javascript/api/excel/excel.shapecollectionloadoptions#image)|对于集合中的每个项目: 返回与该形状相关联的图像。 如果形状类型不是“Image”，则会引发错误。|
||[left](/javascript/api/excel/excel.shapecollectionloadoptions#left)|对于集合中的每个项目: 从形状左侧到工作表左侧的距离 (以磅为单位)。|
||[level](/javascript/api/excel/excel.shapecollectionloadoptions#level)|对于集合中的每一项: 代表指定形状的级别。 例如，级别 0 表示形状不是任何组的一部分，级别 1 表示形状是顶级组的一部分，级别 2 表示形状是顶级子组的一部分。|
||[line](/javascript/api/excel/excel.shapecollectionloadoptions#line)|对于集合中的每一项: 返回与形状相关联的线条。 如果形状类型不是“Line”，则会引发错误。|
||[lineFormat](/javascript/api/excel/excel.shapecollectionloadoptions#lineformat)|对于集合中的每一项: 返回此形状的线条格式。|
||[lockAspectRatio](/javascript/api/excel/excel.shapecollectionloadoptions#lockaspectratio)|对于集合中的每一项: 指定是否锁定此形状的纵横比。|
||[name](/javascript/api/excel/excel.shapecollectionloadoptions#name)|对于集合中的每一项: 代表形状的名称。|
||[parentGroup](/javascript/api/excel/excel.shapecollectionloadoptions#parentgroup)|对于集合中的每个项目: 代表此形状的父组。|
||[rotation](/javascript/api/excel/excel.shapecollectionloadoptions#rotation)|对于集合中的每个项目: 表示形状的旋转角度 (以度为单位)。|
||[textFrame](/javascript/api/excel/excel.shapecollectionloadoptions#textframe)|对于集合中的每一项: 返回此形状的文本框架对象。 只读。|
||[top](/javascript/api/excel/excel.shapecollectionloadoptions#top)|对于集合中的每个项目: 从形状的上边缘到工作表的上边缘之间的距离 (以磅为单位)。|
||[type](/javascript/api/excel/excel.shapecollectionloadoptions#type)|对于集合中的每一项: 返回此形状的类型。 有关详细信息，请参阅 Excel.ShapeType。 只读。|
||[visible](/javascript/api/excel/excel.shapecollectionloadoptions#visible)|对于集合中的每一项: 表示此形状的可见性。|
||[width](/javascript/api/excel/excel.shapecollectionloadoptions#width)|对于集合中的每个项目: 代表形状的宽度 (以磅为单位)。|
||[zOrderPosition](/javascript/api/excel/excel.shapecollectionloadoptions#zorderposition)|对于集合中的每一项: 返回指定的形状在 z-顺序中的位置, 0 表示顺序堆栈的底部。 只读。|
|[ShapeData](/javascript/api/excel/excel.shapedata)|[altTextDescription](/javascript/api/excel/excel.shapedata#alttextdescription)|返回或设置形状对象的可选说明文本。|
||[altTextTitle](/javascript/api/excel/excel.shapedata#alttexttitle)|返回或设置形状对象的可选标题文本。|
||[connectionSiteCount](/javascript/api/excel/excel.shapedata#connectionsitecount)|返回此形状上的连接站点数。 只读。|
||[fill](/javascript/api/excel/excel.shapedata#fill)|返回此形状的填充格式。 只读。|
||[geometricShapeType](/javascript/api/excel/excel.shapedata#geometricshapetype)|表示此几何形状的几何形状类型。 有关详细信息，请参阅 Excel.GeometricShapeType。 如果形状类型不是“GeometricShape”，返回 NULL。|
||[height](/javascript/api/excel/excel.shapedata#height)|表示形状的高度（以磅为单位）。|
||[id](/javascript/api/excel/excel.shapedata#id)|表示形状标识符。 只读。|
||[left](/javascript/api/excel/excel.shapedata#left)|从形状左侧到工作表左侧的距离（以磅为单位）。|
||[level](/javascript/api/excel/excel.shapedata#level)|表示指定形状的级别。 例如，级别 0 表示形状不是任何组的一部分，级别 1 表示形状是顶级组的一部分，级别 2 表示形状是顶级子组的一部分。|
||[lineFormat](/javascript/api/excel/excel.shapedata#lineformat)|返回此形状的线条格式。 只读。|
||[lockAspectRatio](/javascript/api/excel/excel.shapedata#lockaspectratio)|指定此形状的纵横比是否锁定。|
||[名称](/javascript/api/excel/excel.shapedata#name)|表示形状的名称。|
||[rotation](/javascript/api/excel/excel.shapedata#rotation)|表示形状的旋转度数。|
||[top](/javascript/api/excel/excel.shapedata#top)|从形状上边缘到工作表上边缘之间的距离（以磅为单位）。|
||[type](/javascript/api/excel/excel.shapedata#type)|返回此形状的类型。 有关详细信息，请参阅 Excel.ShapeType。 只读。|
||[visible](/javascript/api/excel/excel.shapedata#visible)|表示此形状的可视性。|
||[width](/javascript/api/excel/excel.shapedata#width)|表示形状的宽度（以磅为单位）。|
||[zOrderPosition](/javascript/api/excel/excel.shapedata#zorderposition)|返回指定形状在 z 顺序中的位置，0 表示顺序堆栈的底部。 只读。|
|[ShapeDeactivatedEventArgs](/javascript/api/excel/excel.shapedeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapedeactivatedeventargs#shapeid)|获取已停用的形状的 ID。|
||[type](/javascript/api/excel/excel.shapedeactivatedeventargs#type)|获取事件的类型。 有关详细信息，请参阅 Excel.EventType。|
||[worksheetId](/javascript/api/excel/excel.shapedeactivatedeventargs#worksheetid)|获取其中的形状已停用的工作表的 ID。|
|[ShapeFill](/javascript/api/excel/excel.shapefill)|[clear()](/javascript/api/excel/excel.shapefill#clear--)|清除此形状的填充格式。|
||[foregroundColor](/javascript/api/excel/excel.shapefill#foregroundcolor)|表示窗体 #RRGGBB（例如“FFA500”）的形状填充前景色（采用 HTML 颜色格式）或作为已命名的 HTML 颜色（例如“orange”）|
||[type](/javascript/api/excel/excel.shapefill#type)|返回形状的填充类型。 只读。 有关详细信息，请参阅 Excel.ShapeFillType。|
||[set (properties: ShapeFill)](/javascript/api/excel/excel.shapefill#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: ShapeFillUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.shapefill#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
||[setSolidColor(color: string)](/javascript/api/excel/excel.shapefill#setsolidcolor-color-)|将形状的填充格式设置为统一颜色。 这样可将填充类型更改为“Solid”。|
||[transparency](/javascript/api/excel/excel.shapefill#transparency)|返回或设置填充的透明度百分比，取值范围为 0.0（不透明）到 1.0（清晰）之间。 如果形状类型不支持透明度或形状填充透明度不一致（例如使用渐变填充类型），则返回 null。|
|[ShapeFillData](/javascript/api/excel/excel.shapefilldata)|[foregroundColor](/javascript/api/excel/excel.shapefilldata#foregroundcolor)|表示窗体 #RRGGBB（例如“FFA500”）的形状填充前景色（采用 HTML 颜色格式）或作为已命名的 HTML 颜色（例如“orange”）|
||[transparency](/javascript/api/excel/excel.shapefilldata#transparency)|返回或设置填充的透明度百分比，取值范围为 0.0（不透明）到 1.0（清晰）之间。 如果形状类型不支持透明度或形状填充透明度不一致（例如使用渐变填充类型），则返回 null。|
||[type](/javascript/api/excel/excel.shapefilldata#type)|返回形状的填充类型。 只读。 有关详细信息，请参阅 Excel.ShapeFillType。|
|[ShapeFillLoadOptions](/javascript/api/excel/excel.shapefillloadoptions)|[$all](/javascript/api/excel/excel.shapefillloadoptions#$all)||
||[foregroundColor](/javascript/api/excel/excel.shapefillloadoptions#foregroundcolor)|表示窗体 #RRGGBB（例如“FFA500”）的形状填充前景色（采用 HTML 颜色格式）或作为已命名的 HTML 颜色（例如“orange”）|
||[transparency](/javascript/api/excel/excel.shapefillloadoptions#transparency)|返回或设置填充的透明度百分比，取值范围为 0.0（不透明）到 1.0（清晰）之间。 如果形状类型不支持透明度或形状填充透明度不一致（例如使用渐变填充类型），则返回 null。|
||[type](/javascript/api/excel/excel.shapefillloadoptions#type)|返回形状的填充类型。 只读。 有关详细信息，请参阅 Excel.ShapeFillType。|
|[ShapeFillUpdateData](/javascript/api/excel/excel.shapefillupdatedata)|[foregroundColor](/javascript/api/excel/excel.shapefillupdatedata#foregroundcolor)|表示窗体 #RRGGBB（例如“FFA500”）的形状填充前景色（采用 HTML 颜色格式）或作为已命名的 HTML 颜色（例如“orange”）|
||[transparency](/javascript/api/excel/excel.shapefillupdatedata#transparency)|返回或设置填充的透明度百分比，取值范围为 0.0（不透明）到 1.0（清晰）之间。 如果形状类型不支持透明度或形状填充透明度不一致（例如使用渐变填充类型），则返回 null。|
|[ShapeFont](/javascript/api/excel/excel.shapefont)|[bold](/javascript/api/excel/excel.shapefont#bold)|表示字体的加粗状态。 如果 TextRange 包含粗体和非粗体文本片段，则返回 null。|
||[color](/javascript/api/excel/excel.shapefont#color)|文本颜色的 HTML 颜色代码表示（例如，#FF0000 表示红色）。 如果 TextRange 包含具有不同颜色的文本片段，则返回 null。|
||[italic](/javascript/api/excel/excel.shapefont#italic)|表示字体的斜体状态。 如果 TextRange 包含斜体和非斜体文本片段，则返回 null。|
||[name](/javascript/api/excel/excel.shapefont#name)|表示字体名称（例如“Calibri”）。 如果文本是复杂脚本或东亚语言，则这是相应的字体名称，否则是拉丁字体名称。|
||[set (properties: ShapeFont)](/javascript/api/excel/excel.shapefont#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: ShapeFontUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.shapefont#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
||[size](/javascript/api/excel/excel.shapefont#size)|表示以磅为单位的字体大小（例如 11）。 如果 TextRange 包含具有不同字体大小的文本片段，则返回 null。|
||[underline](/javascript/api/excel/excel.shapefont#underline)|应用于字体的下划线类型。 如果 TextRange 包含具有不同下划线样式的文本片段，则返回 null。 有关详细信息，请参阅 Excel.ShapeFontUnderlineStyle。|
|[ShapeFontData](/javascript/api/excel/excel.shapefontdata)|[bold](/javascript/api/excel/excel.shapefontdata#bold)|表示字体的加粗状态。 如果 TextRange 包含粗体和非粗体文本片段，则返回 null。|
||[color](/javascript/api/excel/excel.shapefontdata#color)|文本颜色的 HTML 颜色代码表示（例如，#FF0000 表示红色）。 如果 TextRange 包含具有不同颜色的文本片段，则返回 null。|
||[italic](/javascript/api/excel/excel.shapefontdata#italic)|表示字体的斜体状态。 如果 TextRange 包含斜体和非斜体文本片段，则返回 null。|
||[name](/javascript/api/excel/excel.shapefontdata#name)|表示字体名称（例如“Calibri”）。 如果文本是复杂脚本或东亚语言，则这是相应的字体名称，否则是拉丁字体名称。|
||[size](/javascript/api/excel/excel.shapefontdata#size)|表示以磅为单位的字体大小（例如 11）。 如果 TextRange 包含具有不同字体大小的文本片段，则返回 null。|
||[underline](/javascript/api/excel/excel.shapefontdata#underline)|应用于字体的下划线类型。 如果 TextRange 包含具有不同下划线样式的文本片段，则返回 null。 有关详细信息，请参阅 Excel.ShapeFontUnderlineStyle。|
|[ShapeFontLoadOptions](/javascript/api/excel/excel.shapefontloadoptions)|[$all](/javascript/api/excel/excel.shapefontloadoptions#$all)||
||[bold](/javascript/api/excel/excel.shapefontloadoptions#bold)|表示字体的加粗状态。 如果 TextRange 包含粗体和非粗体文本片段，则返回 null。|
||[color](/javascript/api/excel/excel.shapefontloadoptions#color)|文本颜色的 HTML 颜色代码表示（例如，#FF0000 表示红色）。 如果 TextRange 包含具有不同颜色的文本片段，则返回 null。|
||[italic](/javascript/api/excel/excel.shapefontloadoptions#italic)|表示字体的斜体状态。 如果 TextRange 包含斜体和非斜体文本片段，则返回 null。|
||[name](/javascript/api/excel/excel.shapefontloadoptions#name)|表示字体名称（例如“Calibri”）。 如果文本是复杂脚本或东亚语言，则这是相应的字体名称，否则是拉丁字体名称。|
||[size](/javascript/api/excel/excel.shapefontloadoptions#size)|表示以磅为单位的字体大小（例如 11）。 如果 TextRange 包含具有不同字体大小的文本片段，则返回 null。|
||[underline](/javascript/api/excel/excel.shapefontloadoptions#underline)|应用于字体的下划线类型。 如果 TextRange 包含具有不同下划线样式的文本片段，则返回 null。 有关详细信息，请参阅 Excel.ShapeFontUnderlineStyle。|
|[ShapeFontUpdateData](/javascript/api/excel/excel.shapefontupdatedata)|[bold](/javascript/api/excel/excel.shapefontupdatedata#bold)|表示字体的加粗状态。 如果 TextRange 包含粗体和非粗体文本片段，则返回 null。|
||[color](/javascript/api/excel/excel.shapefontupdatedata#color)|文本颜色的 HTML 颜色代码表示（例如，#FF0000 表示红色）。 如果 TextRange 包含具有不同颜色的文本片段，则返回 null。|
||[italic](/javascript/api/excel/excel.shapefontupdatedata#italic)|表示字体的斜体状态。 如果 TextRange 包含斜体和非斜体文本片段，则返回 null。|
||[name](/javascript/api/excel/excel.shapefontupdatedata#name)|表示字体名称（例如“Calibri”）。 如果文本是复杂脚本或东亚语言，则这是相应的字体名称，否则是拉丁字体名称。|
||[size](/javascript/api/excel/excel.shapefontupdatedata#size)|表示以磅为单位的字体大小（例如 11）。 如果 TextRange 包含具有不同字体大小的文本片段，则返回 null。|
||[underline](/javascript/api/excel/excel.shapefontupdatedata#underline)|应用于字体的下划线类型。 如果 TextRange 包含具有不同下划线样式的文本片段，则返回 null。 有关详细信息，请参阅 Excel.ShapeFontUnderlineStyle。|
|[ShapeGroup](/javascript/api/excel/excel.shapegroup)|[id](/javascript/api/excel/excel.shapegroup#id)|表示形状标识符。 只读。|
||[shape](/javascript/api/excel/excel.shapegroup#shape)|返回与组关联的 Shape 对象。 只读。|
||[shapes](/javascript/api/excel/excel.shapegroup#shapes)|返回 Shape 对象的集合。 只读。|
||[ungroup()](/javascript/api/excel/excel.shapegroup#ungroup--)|取消分组指定形状组中的任何已分组形状。|
|[ShapeGroupData](/javascript/api/excel/excel.shapegroupdata)|[id](/javascript/api/excel/excel.shapegroupdata#id)|表示形状标识符。 只读。|
||[shapes](/javascript/api/excel/excel.shapegroupdata#shapes)|返回 Shape 对象的集合。 只读。|
|[ShapeGroupLoadOptions](/javascript/api/excel/excel.shapegrouploadoptions)|[$all](/javascript/api/excel/excel.shapegrouploadoptions#$all)||
||[id](/javascript/api/excel/excel.shapegrouploadoptions#id)|表示形状标识符。 只读。|
||[shape](/javascript/api/excel/excel.shapegrouploadoptions#shape)|返回与组关联的 Shape 对象。|
|[ShapeLineFormat](/javascript/api/excel/excel.shapelineformat)|[color](/javascript/api/excel/excel.shapelineformat#color)|表示窗体 #RRGGBB（例如“FFA500”）的线条颜色（采用 HTML 颜色格式）或作为已命名的 HTML 颜色（例如“orange”）|
||[dashStyle](/javascript/api/excel/excel.shapelineformat#dashstyle)|表示形状的线条样式。 当线条不可见或短划线样式不一致时，返回 null。 有关详细信息，请参阅 Excel.ShapeLineStyle。|
||[set (properties: ShapeLineFormat)](/javascript/api/excel/excel.shapelineformat#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: ShapeLineFormatUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.shapelineformat#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
||[style](/javascript/api/excel/excel.shapelineformat#style)|表示形状的线条样式。 当线条不可见或样式不一致时，返回 null。 有关详细信息，请参阅 Excel.ShapeLineStyle。|
||[transparency](/javascript/api/excel/excel.shapelineformat#transparency)|将指定线条的透明度表示为从 0.0（不透明）到 1.0（清晰）的值。 当形状的透明度不一致时，返回 null。|
||[visible](/javascript/api/excel/excel.shapelineformat#visible)|表示形状元素的线条格式是否可见。 当形状的可见性不一致时，返回 null。|
||[weight](/javascript/api/excel/excel.shapelineformat#weight)|表示线条的粗细（以磅为单位）。 当线条不可见或线条粗细不一致时，返回 null。|
|[ShapeLineFormatData](/javascript/api/excel/excel.shapelineformatdata)|[color](/javascript/api/excel/excel.shapelineformatdata#color)|表示窗体 #RRGGBB（例如“FFA500”）的线条颜色（采用 HTML 颜色格式）或作为已命名的 HTML 颜色（例如“orange”）|
||[dashStyle](/javascript/api/excel/excel.shapelineformatdata#dashstyle)|表示形状的线条样式。 当线条不可见或短划线样式不一致时，返回 null。 有关详细信息，请参阅 Excel.ShapeLineStyle。|
||[style](/javascript/api/excel/excel.shapelineformatdata#style)|表示形状的线条样式。 当线条不可见或样式不一致时，返回 null。 有关详细信息，请参阅 Excel.ShapeLineStyle。|
||[transparency](/javascript/api/excel/excel.shapelineformatdata#transparency)|将指定线条的透明度表示为从 0.0（不透明）到 1.0（清晰）的值。 当形状的透明度不一致时，返回 null。|
||[visible](/javascript/api/excel/excel.shapelineformatdata#visible)|表示形状元素的线条格式是否可见。 当形状的可见性不一致时，返回 null。|
||[weight](/javascript/api/excel/excel.shapelineformatdata#weight)|表示线条的粗细（以磅为单位）。 当线条不可见或线条粗细不一致时，返回 null。|
|[ShapeLineFormatLoadOptions](/javascript/api/excel/excel.shapelineformatloadoptions)|[$all](/javascript/api/excel/excel.shapelineformatloadoptions#$all)||
||[color](/javascript/api/excel/excel.shapelineformatloadoptions#color)|表示窗体 #RRGGBB（例如“FFA500”）的线条颜色（采用 HTML 颜色格式）或作为已命名的 HTML 颜色（例如“orange”）|
||[dashStyle](/javascript/api/excel/excel.shapelineformatloadoptions#dashstyle)|表示形状的线条样式。 当线条不可见或短划线样式不一致时，返回 null。 有关详细信息，请参阅 Excel.ShapeLineStyle。|
||[style](/javascript/api/excel/excel.shapelineformatloadoptions#style)|表示形状的线条样式。 当线条不可见或样式不一致时，返回 null。 有关详细信息，请参阅 Excel.ShapeLineStyle。|
||[transparency](/javascript/api/excel/excel.shapelineformatloadoptions#transparency)|将指定线条的透明度表示为从 0.0（不透明）到 1.0（清晰）的值。 当形状的透明度不一致时，返回 null。|
||[visible](/javascript/api/excel/excel.shapelineformatloadoptions#visible)|表示形状元素的线条格式是否可见。 当形状的可见性不一致时，返回 null。|
||[weight](/javascript/api/excel/excel.shapelineformatloadoptions#weight)|表示线条的粗细（以磅为单位）。 当线条不可见或线条粗细不一致时，返回 null。|
|[ShapeLineFormatUpdateData](/javascript/api/excel/excel.shapelineformatupdatedata)|[color](/javascript/api/excel/excel.shapelineformatupdatedata#color)|表示窗体 #RRGGBB（例如“FFA500”）的线条颜色（采用 HTML 颜色格式）或作为已命名的 HTML 颜色（例如“orange”）|
||[dashStyle](/javascript/api/excel/excel.shapelineformatupdatedata#dashstyle)|表示形状的线条样式。 当线条不可见或短划线样式不一致时，返回 null。 有关详细信息，请参阅 Excel.ShapeLineStyle。|
||[style](/javascript/api/excel/excel.shapelineformatupdatedata#style)|表示形状的线条样式。 当线条不可见或样式不一致时，返回 null。 有关详细信息，请参阅 Excel.ShapeLineStyle。|
||[transparency](/javascript/api/excel/excel.shapelineformatupdatedata#transparency)|将指定线条的透明度表示为从 0.0（不透明）到 1.0（清晰）的值。 当形状的透明度不一致时，返回 null。|
||[visible](/javascript/api/excel/excel.shapelineformatupdatedata#visible)|表示形状元素的线条格式是否可见。 当形状的可见性不一致时，返回 null。|
||[weight](/javascript/api/excel/excel.shapelineformatupdatedata#weight)|表示线条的粗细（以磅为单位）。 当线条不可见或线条粗细不一致时，返回 null。|
|[ShapeLoadOptions](/javascript/api/excel/excel.shapeloadoptions)|[$all](/javascript/api/excel/excel.shapeloadoptions#$all)||
||[altTextDescription](/javascript/api/excel/excel.shapeloadoptions#alttextdescription)|返回或设置形状对象的可选说明文本。|
||[altTextTitle](/javascript/api/excel/excel.shapeloadoptions#alttexttitle)|返回或设置形状对象的可选标题文本。|
||[connectionSiteCount](/javascript/api/excel/excel.shapeloadoptions#connectionsitecount)|返回此形状上的连接站点数。 只读。|
||[fill](/javascript/api/excel/excel.shapeloadoptions#fill)|返回此形状的填充格式。|
||[geometricShape](/javascript/api/excel/excel.shapeloadoptions#geometricshape)|返回与形状关联的几何形状。 如果形状类型不是“GeometricShape”，则会引发错误。|
||[geometricShapeType](/javascript/api/excel/excel.shapeloadoptions#geometricshapetype)|表示此几何形状的几何形状类型。 有关详细信息，请参阅 Excel.GeometricShapeType。 如果形状类型不是“GeometricShape”，返回 NULL。|
||[组](/javascript/api/excel/excel.shapeloadoptions#group)|返回与形状关联的形状组。 如果形状类型不是“GroupShape”，则会引发错误。|
||[height](/javascript/api/excel/excel.shapeloadoptions#height)|表示形状的高度（以磅为单位）。|
||[id](/javascript/api/excel/excel.shapeloadoptions#id)|表示形状标识符。 只读。|
||[image](/javascript/api/excel/excel.shapeloadoptions#image)|返回与形状关联的图像。 如果形状类型不是“Image”，则会引发错误。|
||[left](/javascript/api/excel/excel.shapeloadoptions#left)|从形状左侧到工作表左侧的距离（以磅为单位）。|
||[level](/javascript/api/excel/excel.shapeloadoptions#level)|表示指定形状的级别。 例如，级别 0 表示形状不是任何组的一部分，级别 1 表示形状是顶级组的一部分，级别 2 表示形状是顶级子组的一部分。|
||[line](/javascript/api/excel/excel.shapeloadoptions#line)|返回与形状关联的线条。 如果形状类型不是“Line”，则会引发错误。|
||[lineFormat](/javascript/api/excel/excel.shapeloadoptions#lineformat)|返回此形状的线条格式。|
||[lockAspectRatio](/javascript/api/excel/excel.shapeloadoptions#lockaspectratio)|指定此形状的纵横比是否锁定。|
||[名称](/javascript/api/excel/excel.shapeloadoptions#name)|表示形状的名称。|
||[parentGroup](/javascript/api/excel/excel.shapeloadoptions#parentgroup)|表示此形状的父组。|
||[rotation](/javascript/api/excel/excel.shapeloadoptions#rotation)|表示形状的旋转度数。|
||[textFrame](/javascript/api/excel/excel.shapeloadoptions#textframe)|返回此形状的文本框对象。 只读。|
||[top](/javascript/api/excel/excel.shapeloadoptions#top)|从形状上边缘到工作表上边缘之间的距离（以磅为单位）。|
||[type](/javascript/api/excel/excel.shapeloadoptions#type)|返回此形状的类型。 有关详细信息，请参阅 Excel.ShapeType。 只读。|
||[visible](/javascript/api/excel/excel.shapeloadoptions#visible)|表示此形状的可视性。|
||[width](/javascript/api/excel/excel.shapeloadoptions#width)|表示形状的宽度（以磅为单位）。|
||[zOrderPosition](/javascript/api/excel/excel.shapeloadoptions#zorderposition)|返回指定形状在 z 顺序中的位置，0 表示顺序堆栈的底部。 只读。|
|[ShapeUpdateData](/javascript/api/excel/excel.shapeupdatedata)|[altTextDescription](/javascript/api/excel/excel.shapeupdatedata#alttextdescription)|返回或设置形状对象的可选说明文本。|
||[altTextTitle](/javascript/api/excel/excel.shapeupdatedata#alttexttitle)|返回或设置形状对象的可选标题文本。|
||[fill](/javascript/api/excel/excel.shapeupdatedata#fill)|返回此形状的填充格式。|
||[geometricShapeType](/javascript/api/excel/excel.shapeupdatedata#geometricshapetype)|表示此几何形状的几何形状类型。 有关详细信息，请参阅 Excel.GeometricShapeType。 如果形状类型不是“GeometricShape”，返回 NULL。|
||[height](/javascript/api/excel/excel.shapeupdatedata#height)|表示形状的高度（以磅为单位）。|
||[left](/javascript/api/excel/excel.shapeupdatedata#left)|从形状左侧到工作表左侧的距离（以磅为单位）。|
||[lineFormat](/javascript/api/excel/excel.shapeupdatedata#lineformat)|返回此形状的线条格式。|
||[lockAspectRatio](/javascript/api/excel/excel.shapeupdatedata#lockaspectratio)|指定此形状的纵横比是否锁定。|
||[名称](/javascript/api/excel/excel.shapeupdatedata#name)|表示形状的名称。|
||[rotation](/javascript/api/excel/excel.shapeupdatedata#rotation)|表示形状的旋转度数。|
||[top](/javascript/api/excel/excel.shapeupdatedata#top)|从形状上边缘到工作表上边缘之间的距离（以磅为单位）。|
||[visible](/javascript/api/excel/excel.shapeupdatedata#visible)|表示此形状的可视性。|
||[width](/javascript/api/excel/excel.shapeupdatedata#width)|表示形状的宽度（以磅为单位）。|
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
|[TableCollectionLoadOptions](/javascript/api/excel/excel.tablecollectionloadoptions)|[autoFilter](/javascript/api/excel/excel.tablecollectionloadoptions#autofilter)|对于集合中的每一项: 代表表的自动筛选对象。|
|[TableData](/javascript/api/excel/excel.tabledata)|[autoFilter](/javascript/api/excel/excel.tabledata#autofilter)|表示表格的 AutoFilter 对象。 只读。|
|[TableDeletedEventArgs](/javascript/api/excel/excel.tabledeletedeventargs)|[source](/javascript/api/excel/excel.tabledeletedeventargs#source)|指定时间源。 有关详细信息，请参阅 Excel.EventSource。|
||[tableId](/javascript/api/excel/excel.tabledeletedeventargs#tableid)|指定已删除的表格的 ID。|
||[tableName](/javascript/api/excel/excel.tabledeletedeventargs#tablename)|指定已删除的表格的名称。|
||[type](/javascript/api/excel/excel.tabledeletedeventargs#type)|指定事件类型。 有关详细信息，请参阅 Excel.EventType。|
||[worksheetId](/javascript/api/excel/excel.tabledeletedeventargs#worksheetid)|指定已在其内删除表格的工作表的 ID。|
|[TableLoadOptions](/javascript/api/excel/excel.tableloadoptions)|[autoFilter](/javascript/api/excel/excel.tableloadoptions#autofilter)|表示表格的 AutoFilter 对象。|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getCount()](/javascript/api/excel/excel.tablescopedcollection#getcount--)|获取集合中的表数量。|
||[getFirst()](/javascript/api/excel/excel.tablescopedcollection#getfirst--)|获取集合中的第一个表格。 集合中的表格按照从上到下、从左到右的顺序排列，因此左上表格是集合中的第一个表格。|
||[getItem(key: string)](/javascript/api/excel/excel.tablescopedcollection#getitem-key-)|按名称或 ID 获取表。|
||[items](/javascript/api/excel/excel.tablescopedcollection#items)|获取此集合中已加载的子项。|
|[TableScopedCollectionLoadOptions](/javascript/api/excel/excel.tablescopedcollectionloadoptions)|[$all](/javascript/api/excel/excel.tablescopedcollectionloadoptions#$all)||
||[autoFilter](/javascript/api/excel/excel.tablescopedcollectionloadoptions#autofilter)|对于集合中的每一项: 代表表的自动筛选对象。|
||[列](/javascript/api/excel/excel.tablescopedcollectionloadoptions#columns)|对于集合中的每一项: 代表表中所有列的集合。|
||[highlightFirstColumn](/javascript/api/excel/excel.tablescopedcollectionloadoptions#highlightfirstcolumn)|对于集合中的每一项: 指示第一列是否包含特殊格式。|
||[highlightLastColumn](/javascript/api/excel/excel.tablescopedcollectionloadoptions#highlightlastcolumn)|对于集合中的每一项: 指示最后一列是否包含特殊格式。|
||[id](/javascript/api/excel/excel.tablescopedcollectionloadoptions#id)|对于集合中的每一项: 返回一个值, 该值唯一地标识给定工作簿中的表。 即使表被重命名，标识符的值仍然相同。 只读。|
||[legacyId](/javascript/api/excel/excel.tablescopedcollectionloadoptions#legacyid)|对于集合中的每一项: 返回一个数字 id。|
||[name](/javascript/api/excel/excel.tablescopedcollectionloadoptions#name)|对于集合中的每一项: 表的名称。|
||[rows](/javascript/api/excel/excel.tablescopedcollectionloadoptions#rows)|对于集合中的每一项: 代表表中所有行的集合。|
||[showBandedColumns](/javascript/api/excel/excel.tablescopedcollectionloadoptions#showbandedcolumns)|对于集合中的每一项: 指示列是否显示条带格式, 其中奇数列以不同的方式突出显示, 而不是为了使表更易于阅读。|
||[showBandedRows](/javascript/api/excel/excel.tablescopedcollectionloadoptions#showbandedrows)|对于集合中的每一项: 指示行是否显示条带格式, 其中奇数行以不同的方式突出显示, 即使是偶数行, 也可以使表更易于阅读。|
||[showFilterButton](/javascript/api/excel/excel.tablescopedcollectionloadoptions#showfilterbutton)|对于集合中的每一项: 指示筛选按钮是否显示在每个列标头的顶部。 仅当 table 中包含标题行时，才允许设定此设置。|
||[showHeaders](/javascript/api/excel/excel.tablescopedcollectionloadoptions#showheaders)|对于集合中的每一项: 指示标题行是否可见。 该值可以设置为显示或删除标头行。|
||[showTotals](/javascript/api/excel/excel.tablescopedcollectionloadoptions#showtotals)|对于集合中的每一项: 指示汇总行是否可见。 该值可以设置为显示或删除总计行。|
||[sort](/javascript/api/excel/excel.tablescopedcollectionloadoptions#sort)|对于集合中的每一项: 表示对表的排序。|
||[style](/javascript/api/excel/excel.tablescopedcollectionloadoptions#style)|对于集合中的每一项: 表示表样式的常量值。 可能的值是：TableStyleLight1 thru TableStyleLight21、TableStyleMedium1 thru TableStyleMedium28、TableStyleStyleDark1 thru TableStyleStyleDark11。 还可以指定工作簿中显示的用户定义的自定义样式。|
||[worksheet](/javascript/api/excel/excel.tablescopedcollectionloadoptions#worksheet)|对于集合中的每一项: 包含当前表的工作表。|
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
||[set (properties: TextFrame)](/javascript/api/excel/excel.textframe#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: TextFrameUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.textframe#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
||[topMargin](/javascript/api/excel/excel.textframe#topmargin)|表示文本框的上边距（以磅为单位）。|
||[verticalAlignment](/javascript/api/excel/excel.textframe#verticalalignment)|表示文本框的垂直对齐方式。 有关详细信息，请参阅 Excel.ShapeTextVerticalAlignment。|
||[verticalOverflow](/javascript/api/excel/excel.textframe#verticaloverflow)|表示文本框的垂直溢出行为。 有关详细信息，请参阅 Excel.ShapeTextVerticalOverflow。|
|[TextFrameData](/javascript/api/excel/excel.textframedata)|[autoSizeSetting](/javascript/api/excel/excel.textframedata#autosizesetting)|获取或设置文本框的自动调整大小设置。 可以将文本框设置为自动调整文本大小以适应文本框，或自动调整文本框大小以适应文本，或者不使用自动调整大小设置。|
||[bottomMargin](/javascript/api/excel/excel.textframedata#bottommargin)|表示文本框的下边距（以磅为单位）。|
||[hasText](/javascript/api/excel/excel.textframedata#hastext)|指定文本框是否包含文本。|
||[horizontalAlignment](/javascript/api/excel/excel.textframedata#horizontalalignment)|表示文本框的水平对齐方式。 有关详细信息，请参阅 Excel.ShapeTextHorizontalAlignment。|
||[horizontalOverflow](/javascript/api/excel/excel.textframedata#horizontaloverflow)|表示文本框的水平溢出行为。 有关详细信息，请参阅 Excel.ShapeTextHorizontalOverflow。|
||[leftMargin](/javascript/api/excel/excel.textframedata#leftmargin)|表示文本框的左边距（以磅为单位）。|
||[orientation](/javascript/api/excel/excel.textframedata#orientation)|表示文本框的文本方向。 有关详细信息，请参阅 Excel.ShapeTextOrientation。|
||[readingOrder](/javascript/api/excel/excel.textframedata#readingorder)|表示文本框从左到右或从右到左的读取顺序。 有关详细信息，请参阅 Excel.ShapeTextReadingOrder。|
||[rightMargin](/javascript/api/excel/excel.textframedata#rightmargin)|表示文本框的右边距（以磅为单位）。|
||[topMargin](/javascript/api/excel/excel.textframedata#topmargin)|表示文本框的上边距（以磅为单位）。|
||[verticalAlignment](/javascript/api/excel/excel.textframedata#verticalalignment)|表示文本框的垂直对齐方式。 有关详细信息，请参阅 Excel.ShapeTextVerticalAlignment。|
||[verticalOverflow](/javascript/api/excel/excel.textframedata#verticaloverflow)|表示文本框的垂直溢出行为。 有关详细信息，请参阅 Excel.ShapeTextVerticalOverflow。|
|[TextFrameLoadOptions](/javascript/api/excel/excel.textframeloadoptions)|[$all](/javascript/api/excel/excel.textframeloadoptions#$all)||
||[autoSizeSetting](/javascript/api/excel/excel.textframeloadoptions#autosizesetting)|获取或设置文本框的自动调整大小设置。 可以将文本框设置为自动调整文本大小以适应文本框，或自动调整文本框大小以适应文本，或者不使用自动调整大小设置。|
||[bottomMargin](/javascript/api/excel/excel.textframeloadoptions#bottommargin)|表示文本框的下边距（以磅为单位）。|
||[hasText](/javascript/api/excel/excel.textframeloadoptions#hastext)|指定文本框是否包含文本。|
||[horizontalAlignment](/javascript/api/excel/excel.textframeloadoptions#horizontalalignment)|表示文本框的水平对齐方式。 有关详细信息，请参阅 Excel.ShapeTextHorizontalAlignment。|
||[horizontalOverflow](/javascript/api/excel/excel.textframeloadoptions#horizontaloverflow)|表示文本框的水平溢出行为。 有关详细信息，请参阅 Excel.ShapeTextHorizontalOverflow。|
||[leftMargin](/javascript/api/excel/excel.textframeloadoptions#leftmargin)|表示文本框的左边距（以磅为单位）。|
||[orientation](/javascript/api/excel/excel.textframeloadoptions#orientation)|表示文本框的文本方向。 有关详细信息，请参阅 Excel.ShapeTextOrientation。|
||[readingOrder](/javascript/api/excel/excel.textframeloadoptions#readingorder)|表示文本框从左到右或从右到左的读取顺序。 有关详细信息，请参阅 Excel.ShapeTextReadingOrder。|
||[rightMargin](/javascript/api/excel/excel.textframeloadoptions#rightmargin)|表示文本框的右边距（以磅为单位）。|
||[textRange](/javascript/api/excel/excel.textframeloadoptions#textrange)|表示附加到文本框中形状上的文本，以及用于操作文本的属性和方法。 有关详细信息，请参阅 Excel.TextRange。|
||[topMargin](/javascript/api/excel/excel.textframeloadoptions#topmargin)|表示文本框的上边距（以磅为单位）。|
||[verticalAlignment](/javascript/api/excel/excel.textframeloadoptions#verticalalignment)|表示文本框的垂直对齐方式。 有关详细信息，请参阅 Excel.ShapeTextVerticalAlignment。|
||[verticalOverflow](/javascript/api/excel/excel.textframeloadoptions#verticaloverflow)|表示文本框的垂直溢出行为。 有关详细信息，请参阅 Excel.ShapeTextVerticalOverflow。|
|[TextFrameUpdateData](/javascript/api/excel/excel.textframeupdatedata)|[autoSizeSetting](/javascript/api/excel/excel.textframeupdatedata#autosizesetting)|获取或设置文本框的自动调整大小设置。 可以将文本框设置为自动调整文本大小以适应文本框，或自动调整文本框大小以适应文本，或者不使用自动调整大小设置。|
||[bottomMargin](/javascript/api/excel/excel.textframeupdatedata#bottommargin)|表示文本框的下边距（以磅为单位）。|
||[horizontalAlignment](/javascript/api/excel/excel.textframeupdatedata#horizontalalignment)|表示文本框的水平对齐方式。 有关详细信息，请参阅 Excel.ShapeTextHorizontalAlignment。|
||[horizontalOverflow](/javascript/api/excel/excel.textframeupdatedata#horizontaloverflow)|表示文本框的水平溢出行为。 有关详细信息，请参阅 Excel.ShapeTextHorizontalOverflow。|
||[leftMargin](/javascript/api/excel/excel.textframeupdatedata#leftmargin)|表示文本框的左边距（以磅为单位）。|
||[orientation](/javascript/api/excel/excel.textframeupdatedata#orientation)|表示文本框的文本方向。 有关详细信息，请参阅 Excel.ShapeTextOrientation。|
||[readingOrder](/javascript/api/excel/excel.textframeupdatedata#readingorder)|表示文本框从左到右或从右到左的读取顺序。 有关详细信息，请参阅 Excel.ShapeTextReadingOrder。|
||[rightMargin](/javascript/api/excel/excel.textframeupdatedata#rightmargin)|表示文本框的右边距（以磅为单位）。|
||[topMargin](/javascript/api/excel/excel.textframeupdatedata#topmargin)|表示文本框的上边距（以磅为单位）。|
||[verticalAlignment](/javascript/api/excel/excel.textframeupdatedata#verticalalignment)|表示文本框的垂直对齐方式。 有关详细信息，请参阅 Excel.ShapeTextVerticalAlignment。|
||[verticalOverflow](/javascript/api/excel/excel.textframeupdatedata#verticaloverflow)|表示文本框的垂直溢出行为。 有关详细信息，请参阅 Excel.ShapeTextVerticalOverflow。|
|[TextRange](/javascript/api/excel/excel.textrange)|[getSubstring(start: number, length?: number)](/javascript/api/excel/excel.textrange#getsubstring-start--length-)|返回给定区域内子字符串的 TextRange 对象。|
||[font](/javascript/api/excel/excel.textrange#font)|返回一个 ShapeFont 对象，该对象表示文本范围的字体属性。 只读。|
||[set (properties: Excel. TextRange)](/javascript/api/excel/excel.textrange#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: TextRangeUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.textrange#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
||[text](/javascript/api/excel/excel.textrange#text)|表示文本范围的纯文本内容。|
|[TextRangeData](/javascript/api/excel/excel.textrangedata)|[font](/javascript/api/excel/excel.textrangedata#font)|返回一个 ShapeFont 对象，该对象表示文本范围的字体属性。 只读。|
||[text](/javascript/api/excel/excel.textrangedata#text)|表示文本范围的纯文本内容。|
|[TextRangeLoadOptions](/javascript/api/excel/excel.textrangeloadoptions)|[$all](/javascript/api/excel/excel.textrangeloadoptions#$all)||
||[font](/javascript/api/excel/excel.textrangeloadoptions#font)|返回一个 ShapeFont 对象，该对象表示文本范围的字体属性。|
||[text](/javascript/api/excel/excel.textrangeloadoptions#text)|表示文本范围的纯文本内容。|
|[TextRangeUpdateData](/javascript/api/excel/excel.textrangeupdatedata)|[font](/javascript/api/excel/excel.textrangeupdatedata#font)|返回一个 ShapeFont 对象，该对象表示文本范围的字体属性。|
||[text](/javascript/api/excel/excel.textrangeupdatedata#text)|表示文本范围的纯文本内容。|
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
|[WorkbookData](/javascript/api/excel/excel.workbookdata)|[autoSave](/javascript/api/excel/excel.workbookdata#autosave)|指定工作簿是否处于自动保存模式。 只读。|
||[calculationEngineVersion](/javascript/api/excel/excel.workbookdata#calculationengineversion)|返回有关 Excel 计算引擎的版本号。 只读。|
||[chartDataPointTrack](/javascript/api/excel/excel.workbookdata#chartdatapointtrack)|如果工作簿中的所有图表都跟踪它们所附加的实际数据点，则为 True。|
||[isDirty](/javascript/api/excel/excel.workbookdata#isdirty)|指定自上次保存以来是否对指定的工作簿进行任何更改。|
||[previouslySaved](/javascript/api/excel/excel.workbookdata#previouslysaved)|指定工作簿是否已在本地或在线保存。 只读。|
||[usePrecisionAsDisplayed](/javascript/api/excel/excel.workbookdata#useprecisionasdisplayed)|如果此工作簿中的计算仅使用显示的数字精度来完成，则为 True。|
|[WorkbookLoadOptions](/javascript/api/excel/excel.workbookloadoptions)|[autoSave](/javascript/api/excel/excel.workbookloadoptions#autosave)|指定工作簿是否处于自动保存模式。 只读。|
||[calculationEngineVersion](/javascript/api/excel/excel.workbookloadoptions#calculationengineversion)|返回有关 Excel 计算引擎的版本号。 只读。|
||[chartDataPointTrack](/javascript/api/excel/excel.workbookloadoptions#chartdatapointtrack)|如果工作簿中的所有图表都跟踪它们所附加的实际数据点，则为 True。|
||[isDirty](/javascript/api/excel/excel.workbookloadoptions#isdirty)|指定自上次保存以来是否对指定的工作簿进行任何更改。|
||[previouslySaved](/javascript/api/excel/excel.workbookloadoptions#previouslysaved)|指定工作簿是否已在本地或在线保存。 只读。|
||[usePrecisionAsDisplayed](/javascript/api/excel/excel.workbookloadoptions#useprecisionasdisplayed)|如果此工作簿中的计算仅使用显示的数字精度来完成，则为 True。|
|[WorkbookUpdateData](/javascript/api/excel/excel.workbookupdatedata)|[chartDataPointTrack](/javascript/api/excel/excel.workbookupdatedata#chartdatapointtrack)|如果工作簿中的所有图表都跟踪它们所附加的实际数据点，则为 True。|
||[isDirty](/javascript/api/excel/excel.workbookupdatedata#isdirty)|指定自上次保存以来是否对指定的工作簿进行任何更改。|
||[usePrecisionAsDisplayed](/javascript/api/excel/excel.workbookupdatedata#useprecisionasdisplayed)|如果此工作簿中的计算仅使用显示的数字精度来完成，则为 True。|
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
|[WorksheetCollectionLoadOptions](/javascript/api/excel/excel.worksheetcollectionloadoptions)|[autoFilter](/javascript/api/excel/excel.worksheetcollectionloadoptions#autofilter)|对于集合中的每一项: 代表工作表的自动筛选对象。|
||[enableCalculation](/javascript/api/excel/excel.worksheetcollectionloadoptions#enablecalculation)|对于集合中的每一项: 获取或设置工作表的 enableCalculation 属性。|
||[pageLayout](/javascript/api/excel/excel.worksheetcollectionloadoptions#pagelayout)|对于集合中的每一项: 获取工作表的页面布局对象。|
|[WorksheetData](/javascript/api/excel/excel.worksheetdata)|[autoFilter](/javascript/api/excel/excel.worksheetdata#autofilter)|表示工作表的 AutoFilter 对象。 只读。|
||[enableCalculation](/javascript/api/excel/excel.worksheetdata#enablecalculation)|获取或设置工作表的 enableCalculation 属性。|
||[horizontalPageBreaks](/javascript/api/excel/excel.worksheetdata#horizontalpagebreaks)|获取工作表的水平分页符集合。 此集合仅包含手动分页符。|
||[pageLayout](/javascript/api/excel/excel.worksheetdata#pagelayout)|获取工作表的 PageLayout 对象。|
||[shapes](/javascript/api/excel/excel.worksheetdata#shapes)|返回工作表上的所有 Shape 对象的集合。 只读。|
||[verticalPageBreaks](/javascript/api/excel/excel.worksheetdata#verticalpagebreaks)|获取工作表的垂直分页符集合。 此集合仅包含手动分页符。|
|[WorksheetFormatChangedEventArgs](/javascript/api/excel/excel.worksheetformatchangedeventargs)|[address](/javascript/api/excel/excel.worksheetformatchangedeventargs#address)|获取区域地址，该地址表示特定工作表上的更改区域。|
||[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#getrange-ctx-)|获取区域，该区域表示特定工作表上的更改区域。|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#getrangeornullobject-ctx-)|获取区域，该区域表示特定工作表上的更改区域。 它可能会返回 null 对象。|
||[source](/javascript/api/excel/excel.worksheetformatchangedeventargs#source)|获取事件源。 有关详细信息，请参阅 Excel.EventSource。|
||[type](/javascript/api/excel/excel.worksheetformatchangedeventargs#type)|获取事件的类型。 有关详细信息，请参阅 Excel.EventType。|
||[worksheetId](/javascript/api/excel/excel.worksheetformatchangedeventargs#worksheetid)|获取其中的数据发生更改的工作表的 ID。|
|[WorksheetLoadOptions](/javascript/api/excel/excel.worksheetloadoptions)|[autoFilter](/javascript/api/excel/excel.worksheetloadoptions#autofilter)|表示工作表的 AutoFilter 对象。|
||[enableCalculation](/javascript/api/excel/excel.worksheetloadoptions#enablecalculation)|获取或设置工作表的 enableCalculation 属性。|
||[pageLayout](/javascript/api/excel/excel.worksheetloadoptions#pagelayout)|获取工作表的 PageLayout 对象。|
|[WorksheetSearchCriteria](/javascript/api/excel/excel.worksheetsearchcriteria)|[completeMatch](/javascript/api/excel/excel.worksheetsearchcriteria#completematch)|指定匹配必须是完整匹配还是部分匹配。 完全匹配的单元格的全部内容。 默认值为 false（部分）。|
||[matchCase](/javascript/api/excel/excel.worksheetsearchcriteria#matchcase)|指定匹配是否区分大小写。 默认值为 false（不区分大小写）。|
|[WorksheetUpdateData](/javascript/api/excel/excel.worksheetupdatedata)|[enableCalculation](/javascript/api/excel/excel.worksheetupdatedata#enablecalculation)|获取或设置工作表的 enableCalculation 属性。|
||[pageLayout](/javascript/api/excel/excel.worksheetupdatedata#pagelayout)|获取工作表的 PageLayout 对象。|

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 参考文档](/javascript/api/excel)
- [Excel JavaScript API 要求集](./excel-api-requirement-sets.md)
