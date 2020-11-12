---
title: Excel JavaScript API 要求集1。7
description: 有关 ExcelApi 1.7 要求集的详细信息。
ms.date: 11/09/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: ea1fe7a3d28acce2d1f4e9ff33f7b2bd31758fbd
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996233"
---
# <a name="whats-new-in-excel-javascript-api-17"></a>Excel JavaScript API 1.7 的最近更新

Excel JavaScript API 要求集 1.7 的功能包括用于图表、事件、工作表、区域、文档属性、已命名项目、保护选项和样式的 API。

## <a name="customize-charts"></a>自定义图表

通过新的图表 API，你可以创建其他图表类型、向图表中添加数据系列、设置图表标题、添加轴标题、添加显示单位、添加采用移动平均值的趋势线、将趋势线更改为线性趋势线等。 下面是一些示例：

* 图表轴 - 获取、设置、格式化和删除图表中的轴单位、标签和标题。
* 图表系列 - 添加、设置和删除图表中的某个系列。  更改系列标记、绘制顺序和大小。
* 图表趋势线 - 添加、获取和格式化图表中的趋势线。
* 图表图例 - 设置图表中的图例字体的格式。
* 图表点 - 设置图表点颜色。
* 图表标题子字符串 - 获取和设置图表的标题子字符串。
* 图表类型 - 用于创建更多图表类型的选项。

## <a name="events"></a>事件

Excel 事件 API 提供了多个事件处理程序，以便加载项能够在发生特定事件时自动运行指定的函数。 可以将函数设计为执行方案所需的任何操作。 有关当前可用的事件列表，请参阅[使用 Excel JavaScript API 处理事件](../../excel/excel-add-ins-events.md)。

## <a name="customize-the-appearance-of-worksheets-and-ranges"></a>自定义工作表和区域的外观

使用新的 API 可以通过多种方式自定义工作表的外观：

* 冻结窗格，使特定行或列在你滚动工作表时保持可见。 例如，如果工作表中的第一行包含标题，则可以冻结此行，以便在你向下滚动工作表时列标题保持可见。
* 修改工作表标签颜色。
* 添加工作表标题。

可以通过多种方式自定义区域的外观：

* 设置某个区域的单元格样式，确保该区域内的所有单元格采用一致的格式。 单元格 样式是一组定义的格式特征，例如字体和字号、数字格式、单元格边框和单元格底纹。 使用 Excel 中的任意内置单元格样式，或者使用自己的自定义单元格样式。
* 设置区域的文本方向。
* 添加或修改区域上链接至工作表中的其他位置或外部位置的超链接。

## <a name="manage-document-properties"></a>管理文档属性

使用文档属性 API，你可以访问内置文档属性，并且还可以创建和管理自定义文档属性，以存储工作表的状态和驱动工作流和业务逻辑。

## <a name="copy-worksheets"></a>复制工作表

使用工作表复制 APIs，你可以将一个工作表中的数据和格式复制到相同工作簿中的另一个工作表，从而减少所需的数据传输量。

## <a name="handle-ranges-with-ease"></a>轻松地处理区域

使用各种区域 API，你可以完成诸如获取周围区域、获取大小经过重设的区域之类的任务。 这些 API 可以显著提高诸如区域操作和寻址之类任务的效率。

此外：

* 工作簿和工作表保护选项 - 使用这些 API 可保护工作表和工作簿结构中的数据。
* 更新已命名项目 - 使用此 API 可更新已命名项目。
* 获取活动单元格 - 使用此 API 可获取工作表中的活动单元格。

## <a name="api-list"></a>API 列表

下表列出了 Excel JavaScript API 要求集1.7 中的 Api。 若要查看 Excel JavaScript API 要求集1.7 或更早版本支持的所有 Api 的 API 参考文档，请参阅 [要求集1.7 或更早版本中的 Excel api](/javascript/api/excel?view=excel-js-1.7&preserve-view=true)。

| Class | 域 | 说明 |
|:---|:---|:---|
|[Chart](/javascript/api/excel/excel.chart)|[chartType](/javascript/api/excel/excel.chart#charttype)|指定图表的类型。|
||[id](/javascript/api/excel/excel.chart#id)|图表的唯一 ID。|
||[showAllFieldButtons](/javascript/api/excel/excel.chart#showallfieldbuttons)|指定是否在数据透视图上显示所有字段按钮。|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[边框](/javascript/api/excel/excel.chartareaformat#border)|代表图表区域的边框格式，包括颜色、linestyle 和粗细。|
|[ChartAxes](/javascript/api/excel/excel.chartaxes)|[getItem (类型： ChartAxisType、group？： ChartAxisGroup) ](/javascript/api/excel/excel.chartaxes#getitem-type--group-)|返回通过类型和组标识的特定轴。|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[baseTimeUnit](/javascript/api/excel/excel.chartaxis#basetimeunit)|指定指定分类轴的基本单位。|
||[categoryType](/javascript/api/excel/excel.chartaxis#categorytype)|指定分类轴类型。|
||[displayUnit](/javascript/api/excel/excel.chartaxis#displayunit)|表示轴显示单位。|
||[logBase](/javascript/api/excel/excel.chartaxis#logbase)|指定使用对数刻度时的对数的底数。|
||[majorTickMark](/javascript/api/excel/excel.chartaxis#majortickmark)|指定坐标轴的主要刻度线的类型。|
||[majorTimeUnitScale](/javascript/api/excel/excel.chartaxis#majortimeunitscale)|当 CategoryType 属性设置为 "时间刻度" 时，指定分类轴的主要刻度单位刻度值。|
||[minorTickMark](/javascript/api/excel/excel.chartaxis#minortickmark)|指定指定坐标轴的次要刻度线类型。|
||[minorTimeUnitScale](/javascript/api/excel/excel.chartaxis#minortimeunitscale)|当 CategoryType 属性设置为 "时间刻度" 时，指定分类轴的次要刻度单位刻度值。|
||[axisGroup](/javascript/api/excel/excel.chartaxis#axisgroup)|为指定坐标轴指定组。|
||[customDisplayUnit](/javascript/api/excel/excel.chartaxis#customdisplayunit)|指定自定义轴显示单位值。|
||[height](/javascript/api/excel/excel.chartaxis#height)|指定图表坐标轴的高度（以磅为单位）。|
||[left](/javascript/api/excel/excel.chartaxis#left)|指定从坐标轴的左边缘到图表区左侧的距离（以磅为单位）。|
||[top](/javascript/api/excel/excel.chartaxis#top)|指定从轴的上边缘到图表区域顶部的距离（以磅为单位）。|
||[type](/javascript/api/excel/excel.chartaxis#type)|指定坐标轴类型。|
||[width](/javascript/api/excel/excel.chartaxis#width)|指定图表坐标轴的宽度（以磅为单位）。|
||[reversePlotOrder](/javascript/api/excel/excel.chartaxis#reverseplotorder)|指定 Excel 是否将数据点绘制为从最后一个到第一个。|
||[scaleType](/javascript/api/excel/excel.chartaxis#scaletype)|指定数值轴的刻度类型。|
||[setCategoryNames (sourceData： Range) ](/javascript/api/excel/excel.chartaxis#setcategorynames-sourcedata-)|设置指定轴的所有分类名称。|
||[setCustomDisplayUnit (值： number) ](/javascript/api/excel/excel.chartaxis#setcustomdisplayunit-value-)|将轴显示单位设为自定义值。|
||[showDisplayUnitLabel](/javascript/api/excel/excel.chartaxis#showdisplayunitlabel)|指定轴显示单位标签是否可见。|
||[tickLabelPosition](/javascript/api/excel/excel.chartaxis#ticklabelposition)|在指定坐标轴上指定刻度线标签的位置。|
||[tickLabelSpacing](/javascript/api/excel/excel.chartaxis#ticklabelspacing)|指定分类或系列轴刻度线标签之间的数目。|
||[tickMarkSpacing](/javascript/api/excel/excel.chartaxis#tickmarkspacing)|指定分类或刻度线之间的系列数。|
||[visible](/javascript/api/excel/excel.chartaxis#visible)|指定轴是否可见。|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[color](/javascript/api/excel/excel.chartborder#color)|表示图表中的边框颜色的 HTML 颜色代码。|
||[lineStyle](/javascript/api/excel/excel.chartborder#linestyle)|表示边框的线条样式。|
||[weight](/javascript/api/excel/excel.chartborder#weight)|表示边框的粗细，以磅为单位。|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[position](/javascript/api/excel/excel.chartdatalabel#position)|表示数据标签的位置的 DataLabelPosition 值。|
||[分隔符](/javascript/api/excel/excel.chartdatalabel#separator)|该字符串表示用于图表中数据标签的分隔符。|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabel#showbubblesize)|指定数据标签气泡大小是否可见。|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabel#showcategoryname)|指定数据标签类别名称是否可见。|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabel#showlegendkey)|指定数据标签图例标示是否可见。|
||[showPercentage](/javascript/api/excel/excel.chartdatalabel#showpercentage)|指定数据标签百分比是否可见。|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabel#showseriesname)|指定数据标签系列名称是否可见。|
||[showValue](/javascript/api/excel/excel.chartdatalabel#showvalue)|指定数据标签值是否可见。|
|[ChartFormatString](/javascript/api/excel/excel.chartformatstring)|[font](/javascript/api/excel/excel.chartformatstring#font)|代表字体属性，如字体名称、字体大小、颜色等。|
|[ChartLegend](/javascript/api/excel/excel.chartlegend)|[height](/javascript/api/excel/excel.chartlegend#height)|指定图表上的图例的高度（以磅为单位）。|
||[left](/javascript/api/excel/excel.chartlegend#left)|指定图表上的图例的左（以磅为单位）。|
||[legendEntries](/javascript/api/excel/excel.chartlegend#legendentries)|表示图例中 legendEntries 的集合。|
||[showShadow](/javascript/api/excel/excel.chartlegend#showshadow)|指定图例在图表上是否有阴影。|
||[top](/javascript/api/excel/excel.chartlegend#top)|指定图表图例的顶端。|
||[width](/javascript/api/excel/excel.chartlegend#width)|指定图表上的图例的宽度（以磅为单位）。|
|[ChartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|[visible](/javascript/api/excel/excel.chartlegendentry#visible)|表示图表图例条目可见。|
|[ChartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|[getCount()](/javascript/api/excel/excel.chartlegendentrycollection#getcount--)|返回集合中的 legendEntry 数量。|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartlegendentrycollection#getitemat-index-)|返回给定索引处的 legendEntry。|
||[items](/javascript/api/excel/excel.chartlegendentrycollection#items)|获取此集合中已加载的子项。|
|[ChartLineFormat](/javascript/api/excel/excel.chartlineformat)|[lineStyle](/javascript/api/excel/excel.chartlineformat#linestyle)|代表线条样式。|
||[weight](/javascript/api/excel/excel.chartlineformat#weight)|表示线条的粗细（以磅为单位）。|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|[hasDataLabel](/javascript/api/excel/excel.chartpoint#hasdatalabel)|表示数据点是否具有数据标签。|
||[markerBackgroundColor](/javascript/api/excel/excel.chartpoint#markerbackgroundcolor)|数据点的标记背景色的 HTML 颜色代码表示形式 (例如，#FF0000 代表红色) 。|
||[markerForegroundColor](/javascript/api/excel/excel.chartpoint#markerforegroundcolor)|数据点的标记前景色的 HTML 颜色代码表示形式 (例如，#FF0000 代表红色) 。|
||[markerSize](/javascript/api/excel/excel.chartpoint#markersize)|表示数据点的标记大小。|
||[markerStyle](/javascript/api/excel/excel.chartpoint#markerstyle)|表示图表数据点的标记样式。|
||[dataLabel](/javascript/api/excel/excel.chartpoint#datalabel)|返回图表点的数据标签。|
|[ChartPointFormat](/javascript/api/excel/excel.chartpointformat)|[边框](/javascript/api/excel/excel.chartpointformat#border)|表示图表数据点的边框格式，包括颜色、样式和权重信息。|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[chartType](/javascript/api/excel/excel.chartseries#charttype)|表示系列的图表类型。|
||[delete()](/javascript/api/excel/excel.chartseries#delete--)|删除 chart series 对象。|
||[doughnutHoleSize](/javascript/api/excel/excel.chartseries#doughnutholesize)|表示图表系列的圆环孔大小。|
||[筛选](/javascript/api/excel/excel.chartseries#filtered)|指定是否筛选系列。|
||[gapWidth](/javascript/api/excel/excel.chartseries#gapwidth)|表示图表系列的间隙宽度。|
||[hasDataLabels](/javascript/api/excel/excel.chartseries#hasdatalabels)|指定该系列是否具有数据标签。|
||[markerBackgroundColor](/javascript/api/excel/excel.chartseries#markerbackgroundcolor)|指定图表系列的标记背景色。|
||[markerForegroundColor](/javascript/api/excel/excel.chartseries#markerforegroundcolor)|指定图表系列的标记前景色。|
||[markerSize](/javascript/api/excel/excel.chartseries#markersize)|指定图表系列的标记大小。|
||[markerStyle](/javascript/api/excel/excel.chartseries#markerstyle)|指定图表系列的标记样式。|
||[plotOrder](/javascript/api/excel/excel.chartseries#plotorder)|指定图表系列在图表组中的绘制顺序。|
||[趋势](/javascript/api/excel/excel.chartseries#trendlines)|系列中趋势线的集合。|
||[setBubbleSizes (sourceData： Range) ](/javascript/api/excel/excel.chartseries#setbubblesizes-sourcedata-)|设置图表系列的气泡大小。|
||[setValues (sourceData： Range) ](/javascript/api/excel/excel.chartseries#setvalues-sourcedata-)|设置图表系列的值。|
||[setXAxisValues (sourceData： Range) ](/javascript/api/excel/excel.chartseries#setxaxisvalues-sourcedata-)|设置图表系列的 X 轴的值。|
||[showShadow](/javascript/api/excel/excel.chartseries#showshadow)|指定系列是否具有阴影。|
||[平滑](/javascript/api/excel/excel.chartseries#smooth)|指定系列是否平滑。|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[是否添加 (名称？： string，index？： number) ](/javascript/api/excel/excel.chartseriescollection#add-name--index-)|向集合添加新系列。|
|[ChartTitle](/javascript/api/excel/excel.charttitle)|[getSubstring (start： number，length： number) ](/javascript/api/excel/excel.charttitle#getsubstring-start--length-)|获取图表标题的子字符串。|
||[horizontalAlignment](/javascript/api/excel/excel.charttitle#horizontalalignment)|指定图表标题的水平对齐方式。|
||[left](/javascript/api/excel/excel.charttitle#left)|指定图表标题的左边缘到图表区左边缘的距离（以磅为单位）。|
||[position](/javascript/api/excel/excel.charttitle#position)|表示图表标题的位置。|
||[height](/javascript/api/excel/excel.charttitle#height)|返回图表标题的高度，以磅为单位。|
||[width](/javascript/api/excel/excel.charttitle#width)|指定图表标题的宽度（以磅为单位）。|
||[setFormula (公式： string) ](/javascript/api/excel/excel.charttitle#setformula-formula-)|设置一个字符串值，用于表示采用 A1 表示法的图表标题的公式。|
||[showShadow](/javascript/api/excel/excel.charttitle#showshadow)|表示一个布尔值，用于确定图表标题是否具有阴影。|
||[textOrientation](/javascript/api/excel/excel.charttitle#textorientation)|指定文本面向图表标题的角度。|
||[top](/javascript/api/excel/excel.charttitle#top)|指定图表标题的上边缘到图表区域顶部的距离（以磅为单位）。|
||[verticalAlignment](/javascript/api/excel/excel.charttitle#verticalalignment)|指定图表标题的垂直对齐方式。|
|[ChartTitleFormat](/javascript/api/excel/excel.charttitleformat)|[边框](/javascript/api/excel/excel.charttitleformat#border)|代表图表标题的边框格式，包括颜色、linestyle 和粗细。|
|[ChartTrendline](/javascript/api/excel/excel.charttrendline)|[delete()](/javascript/api/excel/excel.charttrendline#delete--)|删除 Trendline 对象。|
||[intercept](/javascript/api/excel/excel.charttrendline#intercept)|表示趋势线的截距值。|
||[movingAveragePeriod](/javascript/api/excel/excel.charttrendline#movingaverageperiod)|表示图表趋势线的周期。|
||[name](/javascript/api/excel/excel.charttrendline#name)|表示趋势线的名称。|
||[polynomialOrder](/javascript/api/excel/excel.charttrendline#polynomialorder)|表示图表趋势线的顺序。|
||[format](/javascript/api/excel/excel.charttrendline#format)|表示图表趋势线的格式。|
||[type](/javascript/api/excel/excel.charttrendline#type)|表示图表趋势线的类型。|
|[ChartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|[是否添加 (类型？： ChartTrendlineType) ](/javascript/api/excel/excel.charttrendlinecollection#add-type-)|向趋势线集合添加新的趋势线。|
||[getCount()](/javascript/api/excel/excel.charttrendlinecollection#getcount--)|返回集合中的趋势线数量。|
||[getItem(index: number)](/javascript/api/excel/excel.charttrendlinecollection#getitem-index-)|按索引（在项目数组中的插入顺序）获取 Trendline 对象。|
||[items](/javascript/api/excel/excel.charttrendlinecollection#items)|获取此集合中已加载的子项。|
|[ChartTrendlineFormat](/javascript/api/excel/excel.charttrendlineformat)|[line](/javascript/api/excel/excel.charttrendlineformat#line)|表示图表线条格式。|
|[CustomProperty](/javascript/api/excel/excel.customproperty)|[delete()](/javascript/api/excel/excel.customproperty#delete--)|删除 custom property 对象。|
||[key](/javascript/api/excel/excel.customproperty#key)|自定义属性的键。|
||[type](/javascript/api/excel/excel.customproperty#type)|用于自定义属性的值的类型。|
||[value](/javascript/api/excel/excel.customproperty#value)|自定义属性的值。|
|[CustomPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|[add (key： string，value： any) ](/javascript/api/excel/excel.custompropertycollection#add-key--value-)|新建自定义属性或设置现有自定义属性。|
||[deleteAll ( # B1 ](/javascript/api/excel/excel.custompropertycollection#deleteall--)|删除此集合中的所有自定义属性。|
||[getCount()](/javascript/api/excel/excel.custompropertycollection#getcount--)|获取自定义属性的计数。|
||[getItem(key: string)](/javascript/api/excel/excel.custompropertycollection#getitem-key-)|按键获取自定义属性对象（不区分大小写）。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.custompropertycollection#getitemornullobject-key-)|按键获取自定义属性对象（不区分大小写）。|
||[items](/javascript/api/excel/excel.custompropertycollection#items)|获取此集合中已加载的子项。|
|[DataConnectionCollection](/javascript/api/excel/excel.dataconnectioncollection)|[refreshAll ( # B1 ](/javascript/api/excel/excel.dataconnectioncollection#refreshall--)|刷新集合中的所有数据连接。|
|[DocumentProperties](/javascript/api/excel/excel.documentproperties)|[编写](/javascript/api/excel/excel.documentproperties#author)|工作簿的作者。|
||[类别](/javascript/api/excel/excel.documentproperties#category)|工作簿的类别。|
||[comments](/javascript/api/excel/excel.documentproperties#comments)|工作簿的注释。|
||[company](/javascript/api/excel/excel.documentproperties#company)|工作簿的公司。|
||[关键字](/javascript/api/excel/excel.documentproperties#keywords)|工作簿的关键字。|
||[manager](/javascript/api/excel/excel.documentproperties#manager)|工作簿的管理者。|
||[creationDate](/javascript/api/excel/excel.documentproperties#creationdate)|获取工作簿的创建日期。|
||[自](/javascript/api/excel/excel.documentproperties#custom)|获取工作簿的自定义属性的集合。|
||[lastAuthor](/javascript/api/excel/excel.documentproperties#lastauthor)|获取工作簿的最终作者。|
||[revisionNumber](/javascript/api/excel/excel.documentproperties#revisionnumber)|获取工作簿的修订号。|
||[subject](/javascript/api/excel/excel.documentproperties#subject)|工作簿的主题。|
||[title](/javascript/api/excel/excel.documentproperties#title)|工作簿的标题。|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[formula](/javascript/api/excel/excel.nameditem#formula)|已命名项的公式。|
||[arrayValues](/javascript/api/excel/excel.nameditem#arrayvalues)|返回包含已命名项目的值和类型的对象。|
|[NamedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|[types](/javascript/api/excel/excel.nameditemarrayvalues#types)|表示已命名项目数组中每个项目的类型|
||[values](/javascript/api/excel/excel.nameditemarrayvalues#values)|表示已命名项目数组中每个项目的值。|
|[Range](/javascript/api/excel/excel.range)|[getAbsoluteResizedRange (numRows： number，numColumns： number) ](/javascript/api/excel/excel.range#getabsoluteresizedrange-numrows--numcolumns-)|获取一个 Range 对象，该对象的左上单元格与当前 Range 对象相同，但具有指定的行数和列数。|
||[getImage ( # B1 ](/javascript/api/excel/excel.range#getimage--)|将区域呈现为 base64 编码的 png 图像。|
||[getSurroundingRegion ( # B1 ](/javascript/api/excel/excel.range#getsurroundingregion--)|返回一个 Range 对象，该对象表示此区域左上单元格的周围区域。|
||[hyperlink](/javascript/api/excel/excel.range#hyperlink)|表示当前区域的超链接。|
||[numberFormatLocal](/javascript/api/excel/excel.range#numberformatlocal)|根据用户的语言设置，表示给定范围的 Excel 数字格式代码。|
||[isEntireColumn](/javascript/api/excel/excel.range#isentirecolumn)|表示当前区域是否为整列。|
||[isEntireRow](/javascript/api/excel/excel.range#isentirerow)|表示当前区域是否为整行。|
||[showCard ( # B1 ](/javascript/api/excel/excel.range#showcard--)|显示活动单元格的卡片（如果该单元格具有富值内容）。|
||[style](/javascript/api/excel/excel.range#style)|表示当前区域的样式。|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[textOrientation](/javascript/api/excel/excel.rangeformat#textorientation)|区域中所有单元格的文本方向。|
||[useStandardHeight](/javascript/api/excel/excel.rangeformat#usestandardheight)|确定 Range 对象的行高是否等于工作表的标准行高。|
||[useStandardWidth](/javascript/api/excel/excel.rangeformat#usestandardwidth)|指定 Range 对象的列宽是否等于工作表的标准宽度。|
|[RangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|[address](/javascript/api/excel/excel.rangehyperlink#address)|表示超链接的 URL 目标。|
||[documentReference](/javascript/api/excel/excel.rangehyperlink#documentreference)|表示超链接的文档引用目标。|
||[屏幕](/javascript/api/excel/excel.rangehyperlink#screentip)|表示鼠标悬停在超链接上时显示的字符串。|
||[textToDisplay](/javascript/api/excel/excel.rangehyperlink#texttodisplay)|表示区域最左上方单元格中显示的字符串。|
|[Style](/javascript/api/excel/excel.style)|[delete()](/javascript/api/excel/excel.style#delete--)|删除此样式。|
||[formulaHidden](/javascript/api/excel/excel.style#formulahidden)|指定在工作表处于保护状态时是否隐藏公式。|
||[horizontalAlignment](/javascript/api/excel/excel.style#horizontalalignment)|表示样式水平对齐。|
||[includeAlignment](/javascript/api/excel/excel.style#includealignment)|指定样式是否包括 AutoIndent、HorizontalAlignment、VerticalAlignment、WrapText、IndentLevel 和 TextOrientation 属性。|
||[includeBorder](/javascript/api/excel/excel.style#includeborder)|指定样式是否包含 Color、ColorIndex、LineStyle 和权数边框属性。|
||[includeFont](/javascript/api/excel/excel.style#includefont)|指定样式是否包括背景、粗体、颜色、ColorIndex、FontStyle、斜体、名称、大小、删除线、下标、上标和下划线字体属性。|
||[includeNumber](/javascript/api/excel/excel.style#includenumber)|指定样式是否包含 NumberFormat 属性。|
||[includePatterns](/javascript/api/excel/excel.style#includepatterns)|指定样式是否包含 Color、ColorIndex、InvertIfNegative、Pattern、PatternColor 和 PatternColorIndex 内部属性。|
||[includeProtection](/javascript/api/excel/excel.style#includeprotection)|指定样式是否包含 FormulaHidden 和 Locked 保护属性。|
||[indentLevel](/javascript/api/excel/excel.style#indentlevel)|0 到 250 之间的一个整数，指示样式的缩进水平。|
||[locked](/javascript/api/excel/excel.style#locked)|指定在工作表处于保护状态时是否锁定对象。|
||[numberFormat](/javascript/api/excel/excel.style#numberformat)|样式中数字格式的格式代码。|
||[numberFormatLocal](/javascript/api/excel/excel.style#numberformatlocal)|样式中数字格式的本地化格式代码。|
||[readingOrder](/javascript/api/excel/excel.style#readingorder)|样式中的阅读顺序。|
||[Borders](/javascript/api/excel/excel.style#borders)|四个 Border 对象的 Border 集合，表示四个边框的样式。|
||[内置](/javascript/api/excel/excel.style#builtin)|指定样式是否为内置样式。|
||[fill](/javascript/api/excel/excel.style#fill)|样式的填充。|
||[font](/javascript/api/excel/excel.style#font)|该 Font 对象表示样式的字体。|
||[name](/javascript/api/excel/excel.style#name)|样式的名称。|
||[shrinkToFit](/javascript/api/excel/excel.style#shrinktofit)|指定文本是否自动收缩以显示可用列宽。|
||[verticalAlignment](/javascript/api/excel/excel.style#verticalalignment)|指定样式的垂直对齐方式。|
||[wrapText](/javascript/api/excel/excel.style#wraptext)|指定 Excel 是否对对象中的文本进行换行。|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[add(name: string)](/javascript/api/excel/excel.stylecollection#add-name-)|向集合添加新样式。|
||[getItem(name: string)](/javascript/api/excel/excel.stylecollection#getitem-name-)|按名称获取样式。|
||[items](/javascript/api/excel/excel.stylecollection#items)|获取此集合中已加载的子项。|
|[Table](/javascript/api/excel/excel.table)|[onChanged](/javascript/api/excel/excel.table#onchanged)|当特定表上单元格的数据更改时发生。|
||[onSelectionChanged](/javascript/api/excel/excel.table#onselectionchanged)|当特定表格上的所选内容更改时发生。|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[address](/javascript/api/excel/excel.tablechangedeventargs#address)|获取地址，该地址表示特定工作表上的表格的更改区域。|
||[changeType](/javascript/api/excel/excel.tablechangedeventargs#changetype)|获取更改类型，该类型表示 Changed 事件的触发方式。|
||[source](/javascript/api/excel/excel.tablechangedeventargs#source)|获取事件源。|
||[tableId](/javascript/api/excel/excel.tablechangedeventargs#tableid)|获取其中的数据发生更改的表格的 ID。|
||[type](/javascript/api/excel/excel.tablechangedeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.tablechangedeventargs#worksheetid)|获取其中的数据发生更改的工作表的 ID。|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onChanged](/javascript/api/excel/excel.tablecollection#onchanged)|当工作簿中的任何表或工作表上的数据发生更改时发生。|
|[TableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|[address](/javascript/api/excel/excel.tableselectionchangedeventargs#address)|获取区域地址，该地址表示特定工作表上的表格选定区域。|
||[isInsideTable](/javascript/api/excel/excel.tableselectionchangedeventargs#isinsidetable)|指定所选内容是否在表格中，如果 IsInsideTable 为 false，则地址将无用。|
||[tableId](/javascript/api/excel/excel.tableselectionchangedeventargs#tableid)|获取其中的选定区域发生更改的表格 ID。|
||[type](/javascript/api/excel/excel.tableselectionchangedeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.tableselectionchangedeventargs#worksheetid)|获取其中的选定区域发生更改的工作表的 ID。|
|[Workbook](/javascript/api/excel/excel.workbook)|[getActiveCell ( # B1 ](/javascript/api/excel/excel.workbook#getactivecell--)|获取工作簿中当前处于活动状态的单元格。|
||[dataConnections](/javascript/api/excel/excel.workbook#dataconnections)|表示工作簿中的所有数据连接。|
||[name](/javascript/api/excel/excel.workbook#name)|获取工作簿名称。|
||[properties](/javascript/api/excel/excel.workbook#properties)|获取工作簿属性。|
||[protection](/javascript/api/excel/excel.workbook#protection)|返回工作簿的保护对象。|
||[styles](/javascript/api/excel/excel.workbook#styles)|表示与工作簿关联的样式的集合。|
|[WorkbookProtection](/javascript/api/excel/excel.workbookprotection)|[保护 (密码？： string) ](/javascript/api/excel/excel.workbookprotection#protect-password-)|保护工作簿。|
||[受保护](/javascript/api/excel/excel.workbookprotection#protected)|指定工作簿是否受保护。|
||[取消保护 (密码？： string) ](/javascript/api/excel/excel.workbookprotection#unprotect-password-)|解除保护工作簿。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[copy (positionType？： WorksheetPositionType，relativeTo？： Excel. 工作表) ](/javascript/api/excel/excel.worksheet#copy-positiontype--relativeto-)|复制工作表并将其放在指定位置。|
||[getRangeByIndexes (startRow： number，startColumn： number，rowCount： number，columnCount： number) ](/javascript/api/excel/excel.worksheet#getrangebyindexes-startrow--startcolumn--rowcount--columncount-)|获取以特定行索引和列索引开始并跨越了一定数量的行和列的 range 对象。|
||[freezePanes](/javascript/api/excel/excel.worksheet#freezepanes)|获取一个对象，该对象可用于操作工作表上的冻结窗格。|
||[onActivated](/javascript/api/excel/excel.worksheet#onactivated)|当激活工作表时发生此事件。|
||[onChanged](/javascript/api/excel/excel.worksheet#onchanged)|当指定的工作表上的数据发生更改时发生。|
||[onDeactivated](/javascript/api/excel/excel.worksheet#ondeactivated)|停用工作表时发生此事件。|
||[onSelectionChanged](/javascript/api/excel/excel.worksheet#onselectionchanged)|当指定的工作表上的所选内容更改时发生。|
||[standardHeight](/javascript/api/excel/excel.worksheet#standardheight)|返回工作表中所有行的标准（默认）行高，以磅为单位。|
||[standardWidth](/javascript/api/excel/excel.worksheet#standardwidth)|指定工作表中所有列的标准 (默认) 宽度。|
||[tabColor](/javascript/api/excel/excel.worksheet#tabcolor)|工作表的选项卡颜色。|
|[WorksheetActivatedEventArgs](/javascript/api/excel/excel.worksheetactivatedeventargs)|[type](/javascript/api/excel/excel.worksheetactivatedeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.worksheetactivatedeventargs#worksheetid)|获取已启用的工作表的 ID。|
|[WorksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|[source](/javascript/api/excel/excel.worksheetaddedeventargs#source)|获取事件源。|
||[type](/javascript/api/excel/excel.worksheetaddedeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.worksheetaddedeventargs#worksheetid)|获取已添加至工作簿的工作表的 ID。|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[address](/javascript/api/excel/excel.worksheetchangedeventargs#address)|获取区域地址，该地址表示特定工作表上的更改区域。|
||[changeType](/javascript/api/excel/excel.worksheetchangedeventargs#changetype)|获取更改类型，该类型表示 Changed 事件的触发方式。|
||[source](/javascript/api/excel/excel.worksheetchangedeventargs#source)|获取事件源。|
||[type](/javascript/api/excel/excel.worksheetchangedeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.worksheetchangedeventargs#worksheetid)|获取其中的数据发生更改的工作表的 ID。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onActivated](/javascript/api/excel/excel.worksheetcollection#onactivated)|当工作簿中的任何工作表激活时发生此事件。|
||[onAdded](/javascript/api/excel/excel.worksheetcollection#onadded)|将新工作表添加到工作簿时发生此事件。|
||[onDeactivated](/javascript/api/excel/excel.worksheetcollection#ondeactivated)|当工作簿中的任何工作表被停用时发生。|
||[onDeleted](/javascript/api/excel/excel.worksheetcollection#ondeleted)|当从工作簿中删除工作表时发生此事件。|
|[WorksheetDeactivatedEventArgs](/javascript/api/excel/excel.worksheetdeactivatedeventargs)|[type](/javascript/api/excel/excel.worksheetdeactivatedeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.worksheetdeactivatedeventargs#worksheetid)|获取已停用的工作表的 ID。|
|[WorksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|[source](/javascript/api/excel/excel.worksheetdeletedeventargs#source)|获取事件源。|
||[type](/javascript/api/excel/excel.worksheetdeletedeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.worksheetdeletedeventargs#worksheetid)|获取已从工作簿删除的工作表的 ID。|
|[WorksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|[freezeAt (frozenRange： Range \| string) ](/javascript/api/excel/excel.worksheetfreezepanes#freezeat-frozenrange-)|设置活动工作表视图中的冻结单元格。|
||[freezeColumns (count？： number) ](/javascript/api/excel/excel.worksheetfreezepanes#freezecolumns-count-)|就地冻结工作表的第一列。|
||[freezeRows (count？： number) ](/javascript/api/excel/excel.worksheetfreezepanes#freezerows-count-)|就地冻结工作表的顶行。|
||[getLocation()](/javascript/api/excel/excel.worksheetfreezepanes#getlocation--)|获取用于描述活动工作表视图中的冻结单元格的区域。|
||[getLocationOrNullObject ( # B1 ](/javascript/api/excel/excel.worksheetfreezepanes#getlocationornullobject--)|获取用于描述活动工作表视图中的冻结单元格的区域。|
||[取消冻结 ( # B1 ](/javascript/api/excel/excel.worksheetfreezepanes#unfreeze--)|移除工作表中的所有冻结窗格。|
|[WorksheetProtection](/javascript/api/excel/excel.worksheetprotection)|[取消保护 (密码？： string) ](/javascript/api/excel/excel.worksheetprotection#unprotect-password-)|解除对 worksheet 的保护。|
|[WorksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|[allowEditObjects](/javascript/api/excel/excel.worksheetprotectionoptions#alloweditobjects)|表示允许编辑对象的工作表保护选项。|
||[allowEditScenarios](/javascript/api/excel/excel.worksheetprotectionoptions#alloweditscenarios)|表示允许编辑应用场景的工作表保护选项。|
||[selectionMode](/javascript/api/excel/excel.worksheetprotectionoptions#selectionmode)|表示选择模式的工作表保护选项。|
|[WorksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|[address](/javascript/api/excel/excel.worksheetselectionchangedeventargs#address)|获取区域地址，该地址表示特定工作表上的选定区域。|
||[type](/javascript/api/excel/excel.worksheetselectionchangedeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.worksheetselectionchangedeventargs#worksheetid)|获取其中的选定区域发生更改的工作表的 ID。|

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 参考文档](/javascript/api/excel?view=excel-js-1.7&preserve-view=true)
- [Excel JavaScript API 要求集](excel-api-requirement-sets.md)
