---
title: Excel JavaScript API 要求集 1.8
description: 有关 ExcelApi 1.8 要求集的详细信息。
ms.date: 03/19/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 6e5a87741618d8d132bc699e2a5b14c68b4403b6
ms.sourcegitcommit: 7482ab6bc258d98acb9ba9b35c7dd3b5cc5bed21
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/24/2021
ms.locfileid: "51178081"
---
# <a name="whats-new-in-excel-javascript-api-18"></a>Excel JavaScript API 1.8 的新增功能

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

下表列出了 Excel JavaScript API 要求集 1.8 中的 API。 若要查看 Excel JavaScript API 要求集 1.8 或更早版本支持的所有 API 的 API 参考文档，请参阅要求集 [1.8](/javascript/api/excel?view=excel-js-1.8&preserve-view=true)或更早版本中的 Excel API。

| 类 | 域 | 说明 |
|:---|:---|:---|
|[BasicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|[formula1](/javascript/api/excel/excel.basicdatavalidation#formula1)|当运算符属性设置为二进制运算符（如 GreaterThan (）时，指定右侧操作数 (左侧操作数是用户尝试在单元格) 中输入的值。|
||[formula2](/javascript/api/excel/excel.basicdatavalidation#formula2)|使用三元运算符 Between 和 NotBetween 指定上限操作数。|
||[operator](/javascript/api/excel/excel.basicdatavalidation#operator)|用于验证数据有效性的运算符。|
|[Chart](/javascript/api/excel/excel.chart)|[categoryLabelLevel](/javascript/api/excel/excel.chart#categorylabellevel)|指定 ChartCategoryLabelLevel 枚举常量，该常量引用|
||[displayBlanksAs](/javascript/api/excel/excel.chart#displayblanksas)|指定在图表上绘制空白单元格的方式。|
||[plotBy](/javascript/api/excel/excel.chart#plotby)|指定列或行在图表上用作数据系列的方式。|
||[plotVisibleOnly](/javascript/api/excel/excel.chart#plotvisibleonly)|如果仅绘制可见单元格，则为 True。 如果绘制可见单元格和隐藏单元格，则为 False。|
||[onActivated](/javascript/api/excel/excel.chart#onactivated)|激活图表时发生。|
||[onDeactivated](/javascript/api/excel/excel.chart#ondeactivated)|在图表被停用时发生。|
||[plotArea](/javascript/api/excel/excel.chart#plotarea)|表示图表的绘制区域。|
||[seriesNameLevel](/javascript/api/excel/excel.chart#seriesnamelevel)|指定 ChartSeriesNameLevel 枚举常量，该常量引用|
||[showDataLabelsOverMaximum](/javascript/api/excel/excel.chart#showdatalabelsovermaximum)|指定当值大于数值轴上的最大值时是否显示数据标签。|
||[style](/javascript/api/excel/excel.chart#style)|指定图表的图表样式。|
|[ChartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|[chartId](/javascript/api/excel/excel.chartactivatedeventargs#chartid)|获取已启用图表的 ID。|
||[type](/javascript/api/excel/excel.chartactivatedeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.chartactivatedeventargs#worksheetid)|获取其中的图表已启用的工作表的 ID。|
|[ChartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|[chartId](/javascript/api/excel/excel.chartaddedeventargs#chartid)|获取已添加至工作表的图表的 ID。|
||[source](/javascript/api/excel/excel.chartaddedeventargs#source)|获取事件源。|
||[type](/javascript/api/excel/excel.chartaddedeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.chartaddedeventargs#worksheetid)|获取已在其中添加图表的工作表的 ID。|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[alignment](/javascript/api/excel/excel.chartaxis#alignment)|指定指定坐标轴刻度线标签的对齐方式。|
||[isBetweenCategories](/javascript/api/excel/excel.chartaxis#isbetweencategories)|指定数值轴是否与分类之间的分类轴相交。|
||[multiLevel](/javascript/api/excel/excel.chartaxis#multilevel)|指定坐标轴是否多级。|
||[numberFormat](/javascript/api/excel/excel.chartaxis#numberformat)|指定坐标轴刻度线标签的格式代码。|
||[offset](/javascript/api/excel/excel.chartaxis#offset)|指定标签级别之间的距离，以及第一级标签与轴线之间的距离。|
||[position](/javascript/api/excel/excel.chartaxis#position)|指定两轴交叉的指定坐标轴位置。|
||[positionAt](/javascript/api/excel/excel.chartaxis#positionat)|指定两轴交叉的指定坐标轴位置。|
||[setPositionAt (值：number) ](/javascript/api/excel/excel.chartaxis#setpositionat-value-)|设置两轴交叉的指定坐标轴位置。|
||[textOrientation](/javascript/api/excel/excel.chartaxis#textorientation)|指定文本面向图表坐标轴刻度线标签的角度。|
|[ChartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|[fill](/javascript/api/excel/excel.chartaxisformat#fill)|指定图表填充格式。|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[setFormula (公式：string) ](/javascript/api/excel/excel.chartaxistitle#setformula-formula-)|该字符串值表示采用 A1 表示法的图表轴标题的公式。|
|[ChartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|[border](/javascript/api/excel/excel.chartaxistitleformat#border)|指定图表坐标轴标题的边框格式，包括颜色、线条和粗细。|
||[fill](/javascript/api/excel/excel.chartaxistitleformat#fill)|指定图表坐标轴标题的填充格式。|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[clear()](/javascript/api/excel/excel.chartborder#clear--)|清除图表元素的边框格式。|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[onActivated](/javascript/api/excel/excel.chartcollection#onactivated)|激活图表时发生。|
||[onAdded](/javascript/api/excel/excel.chartcollection#onadded)|在将新图表添加到工作表时发生。|
||[onDeactivated](/javascript/api/excel/excel.chartcollection#ondeactivated)|停用图表时发生。|
||[onDeleted](/javascript/api/excel/excel.chartcollection#ondeleted)|删除图表时发生。|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[autoText](/javascript/api/excel/excel.chartdatalabel#autotext)|指定数据标签是否根据上下文自动生成相应的文本。|
||[formula](/javascript/api/excel/excel.chartdatalabel#formula)|该字符串值表示采用 A1 表示法的图表数据标签的公式。|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabel#horizontalalignment)|表示图表数据标签水平对齐。|
||[left](/javascript/api/excel/excel.chartdatalabel#left)|表示图表数据标签左边缘到图表区域左边缘的距离，以磅为单位。|
||[numberFormat](/javascript/api/excel/excel.chartdatalabel#numberformat)|该字符串值表示数据标签的格式代码。|
||[format](/javascript/api/excel/excel.chartdatalabel#format)|表示图表数据标签的格式。|
||[height](/javascript/api/excel/excel.chartdatalabel#height)|返回图表数据标签的高度，以磅为单位。|
||[width](/javascript/api/excel/excel.chartdatalabel#width)|返回图表数据标签的宽度，以磅为单位。|
||[text](/javascript/api/excel/excel.chartdatalabel#text)|该字符串表示图表上的数据标签文本。|
||[textOrientation](/javascript/api/excel/excel.chartdatalabel#textorientation)|表示文本面向图表数据标签的角度。|
||[top](/javascript/api/excel/excel.chartdatalabel#top)|表示图表数据标签上边缘到图表区域顶部的距离，以磅为单位。|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabel#verticalalignment)|表示图表数据标签垂直对齐。|
|[ChartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|[border](/javascript/api/excel/excel.chartdatalabelformat#border)|表示边框格式，包括颜色、线条样式和粗细。|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[autoText](/javascript/api/excel/excel.chartdatalabels#autotext)|指定数据标签是否根据上下文自动生成相应的文本。|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabels#horizontalalignment)|指定图表数据标签的水平对齐方式。|
||[numberFormat](/javascript/api/excel/excel.chartdatalabels#numberformat)|指定数据标签的格式代码。|
||[textOrientation](/javascript/api/excel/excel.chartdatalabels#textorientation)|表示文本面向数据标签的角度。|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabels#verticalalignment)|表示图表数据标签垂直对齐。|
|[ChartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|[chartId](/javascript/api/excel/excel.chartdeactivatedeventargs#chartid)|获取停用图表的 ID。|
||[type](/javascript/api/excel/excel.chartdeactivatedeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.chartdeactivatedeventargs#worksheetid)|获取其中的图表已停用的工作表的 ID。|
|[ChartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|[chartId](/javascript/api/excel/excel.chartdeletedeventargs#chartid)|获取已从工作表删除的图表的 ID。|
||[source](/javascript/api/excel/excel.chartdeletedeventargs#source)|获取事件源。|
||[type](/javascript/api/excel/excel.chartdeletedeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.chartdeletedeventargs#worksheetid)|获取已在其中删除图表的工作表的 ID。|
|[ChartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|[height](/javascript/api/excel/excel.chartlegendentry#height)|指定图表图例上的 legendEntry 的高度。|
||[index](/javascript/api/excel/excel.chartlegendentry#index)|指定图表图例中 legendEntry 的索引。|
||[left](/javascript/api/excel/excel.chartlegendentry#left)|指定图表 legendEntry 的左侧。|
||[top](/javascript/api/excel/excel.chartlegendentry#top)|指定图表 legendEntry 的顶部。|
||[width](/javascript/api/excel/excel.chartlegendentry#width)|表示图表图例上的 legendEntry 的宽度。|
|[ChartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|[border](/javascript/api/excel/excel.chartlegendformat#border)|表示边框格式，包括颜色、线条样式和粗细。|
|[ChartPlotArea](/javascript/api/excel/excel.chartplotarea)|[height](/javascript/api/excel/excel.chartplotarea#height)|指定 plotArea 的高度值。|
||[insideHeight](/javascript/api/excel/excel.chartplotarea#insideheight)|指定 plotArea 的 insideHeight 值。|
||[insideLeft](/javascript/api/excel/excel.chartplotarea#insideleft)|指定 plotArea 的 insideLeft 值。|
||[insideTop](/javascript/api/excel/excel.chartplotarea#insidetop)|指定 plotArea 的 insideTop 值。|
||[insideWidth](/javascript/api/excel/excel.chartplotarea#insidewidth)|指定 plotArea 的 insideWidth 值。|
||[left](/javascript/api/excel/excel.chartplotarea#left)|指定 plotArea 的左侧值。|
||[position](/javascript/api/excel/excel.chartplotarea#position)|指定 plotArea 的位置。|
||[format](/javascript/api/excel/excel.chartplotarea#format)|指定图表 plotArea 的格式。|
||[top](/javascript/api/excel/excel.chartplotarea#top)|指定 plotArea 的上值。|
||[width](/javascript/api/excel/excel.chartplotarea#width)|指定 plotArea 的宽度值。|
|[ChartPlotAreaFormat](/javascript/api/excel/excel.chartplotareaformat)|[border](/javascript/api/excel/excel.chartplotareaformat#border)|指定图表 plotArea 的边框属性。|
||[fill](/javascript/api/excel/excel.chartplotareaformat#fill)|指定对象的填充格式，其中包括背景格式信息。|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[axisGroup](/javascript/api/excel/excel.chartseries#axisgroup)|指定指定系列的组。|
||[explosion](/javascript/api/excel/excel.chartseries#explosion)|指定饼图或圆环图扇区的分解值。|
||[firstSliceAngle](/javascript/api/excel/excel.chartseries#firstsliceangle)|指定第一个饼图或圆环图扇区的角度（以度 (从垂直方向顺) 。|
||[invertIfNegative](/javascript/api/excel/excel.chartseries#invertifnegative)|如果 Excel 在对应于负数时反转项中的图案，则其为 True。|
||[overlap](/javascript/api/excel/excel.chartseries#overlap)|指定条柱的摆放方式。|
||[dataLabels](/javascript/api/excel/excel.chartseries#datalabels)|表示系列中所有数据标签的集合。|
||[secondPlotSize](/javascript/api/excel/excel.chartseries#secondplotsize)|指定复合饼图或复合条饼图的第二部分的大小，以主饼图大小的百分比表示。|
||[splitType](/javascript/api/excel/excel.chartseries#splittype)|指定拆分复合饼图或复合条饼图的两部分的方式。|
||[varyByCategories](/javascript/api/excel/excel.chartseries#varybycategories)|如此 如果 Excel 为每个数据标记分配不同的颜色或图案。|
|[ChartTrendline](/javascript/api/excel/excel.charttrendline)|[backwardPeriod](/javascript/api/excel/excel.charttrendline#backwardperiod)|表示趋势线向后延伸的周期数。|
||[forwardPeriod](/javascript/api/excel/excel.charttrendline#forwardperiod)|表示趋势线向前延伸的周期数。|
||[label](/javascript/api/excel/excel.charttrendline#label)|表示图表趋势线的标签。|
||[showEquation](/javascript/api/excel/excel.charttrendline#showequation)|如果图表上显示趋势线公式，则为 True。|
||[showRSquared](/javascript/api/excel/excel.charttrendline#showrsquared)|如果图表上显示趋势线的 R 平方值，则为 True。|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[autoText](/javascript/api/excel/excel.charttrendlinelabel#autotext)|指定趋势线标签是否根据上下文自动生成相应的文本。|
||[formula](/javascript/api/excel/excel.charttrendlinelabel#formula)|该字符串值表示采用 A1 表示法的图表趋势线标签的公式。|
||[horizontalAlignment](/javascript/api/excel/excel.charttrendlinelabel#horizontalalignment)|表示图表趋势线标签水平对齐。|
||[left](/javascript/api/excel/excel.charttrendlinelabel#left)|表示图表趋势线标签左边缘到图表区域左边缘的距离，以磅为单位。|
||[numberFormat](/javascript/api/excel/excel.charttrendlinelabel#numberformat)|该字符串值表示趋势线标签的格式代码。|
||[format](/javascript/api/excel/excel.charttrendlinelabel#format)|图表趋势线标签的格式。|
||[height](/javascript/api/excel/excel.charttrendlinelabel#height)|返回图表趋势线标签的高度，以磅为单位。|
||[width](/javascript/api/excel/excel.charttrendlinelabel#width)|返回图表趋势线标签的宽度，以磅为单位。|
||[text](/javascript/api/excel/excel.charttrendlinelabel#text)|该字符串表示图表上的趋势线标签文本。|
||[textOrientation](/javascript/api/excel/excel.charttrendlinelabel#textorientation)|表示文本面向图表趋势线标签的角度。|
||[top](/javascript/api/excel/excel.charttrendlinelabel#top)|表示图表趋势线标签上边缘到图表区域顶部的距离，以磅为单位。|
||[verticalAlignment](/javascript/api/excel/excel.charttrendlinelabel#verticalalignment)|表示图表趋势线标签垂直对齐。|
|[ChartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|[border](/javascript/api/excel/excel.charttrendlinelabelformat#border)|指定边框格式，包括颜色、线条和粗细。|
||[fill](/javascript/api/excel/excel.charttrendlinelabelformat#fill)|指定当前图表趋势线标签的填充格式。|
||[font](/javascript/api/excel/excel.charttrendlinelabelformat#font)|指定图表趋势 (的字体属性) 字体名称、字体大小、颜色等。|
|[CustomDataValidation](/javascript/api/excel/excel.customdatavalidation)|[formula](/javascript/api/excel/excel.customdatavalidation#formula)|自定义数据验证公式。|
|[DataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|[name](/javascript/api/excel/excel.datapivothierarchy#name)|DataPivotHierarchy 的名称。|
||[numberFormat](/javascript/api/excel/excel.datapivothierarchy#numberformat)|DataPivotHierarchy 的数字格式。|
||[position](/javascript/api/excel/excel.datapivothierarchy#position)|DataPivotHierarchy 的位置。|
||[field](/javascript/api/excel/excel.datapivothierarchy#field)|返回与 DataPivotHierarchy 相关联的 PivotFields。|
||[id](/javascript/api/excel/excel.datapivothierarchy#id)|DataPivotHierarchy ID。|
||[setToDefault () ](/javascript/api/excel/excel.datapivothierarchy#settodefault--)|将 DataPivotHierarchy 重置回其默认值。|
||[showAs](/javascript/api/excel/excel.datapivothierarchy#showas)|指定数据是否应显示为特定的摘要计算。|
||[summarizeBy](/javascript/api/excel/excel.datapivothierarchy#summarizeby)|指定是否显示 DataPivotHierarchy 的所有项。|
|[DataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|[添加 (pivotHierarchy：Excel.PivotHierarchy) ](/javascript/api/excel/excel.datapivothierarchycollection#add-pivothierarchy-)|将 PivotHierarchy 添加到当前轴。|
||[getCount()](/javascript/api/excel/excel.datapivothierarchycollection#getcount--)|获取集合中的透视层级结构的数量。|
||[getItem(name: string)](/javascript/api/excel/excel.datapivothierarchycollection#getitem-name-)|按名称或 ID 获取 DataPivotHierarchy。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.datapivothierarchycollection#getitemornullobject-name-)|按名称获取 DataPivotHierarchy。|
||[items](/javascript/api/excel/excel.datapivothierarchycollection#items)|获取此集合中已加载的子项。|
||[remove (DataPivotHierarchy： Excel.DataPivotHierarchy) ](/javascript/api/excel/excel.datapivothierarchycollection#remove-datapivothierarchy-)|从当前轴删除 PivotHierarchy。|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[clear()](/javascript/api/excel/excel.datavalidation#clear--)|清除当前区域中的数据有效性。|
||[errorAlert](/javascript/api/excel/excel.datavalidation#erroralert)|用户输入无效数据时，出现错误警报。|
||[ignoreBlanks](/javascript/api/excel/excel.datavalidation#ignoreblanks)|指定是否对空白单元格执行数据验证，默认为 true。|
||[prompt](/javascript/api/excel/excel.datavalidation#prompt)|当用户选择单元格时提示。|
||[类型](/javascript/api/excel/excel.datavalidation#type)|数据有效性类型，有关详细信息，请参阅 Excel.DataValidationType。|
||[valid](/javascript/api/excel/excel.datavalidation#valid)|表示所有单元格值根据数据有效性规则是否全部有效。|
||[rule](/javascript/api/excel/excel.datavalidation#rule)|包含不同类型的数据有效性条件的数据有效性规则。|
|[DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|[邮件](/javascript/api/excel/excel.datavalidationerroralert#message)|表示错误警报消息。|
||[showAlert](/javascript/api/excel/excel.datavalidationerroralert#showalert)|指定在用户输入无效数据时是否显示错误警报对话框。|
||[style](/javascript/api/excel/excel.datavalidationerroralert#style)|数据有效性警报类型，有关详细信息，请参阅 Excel.DataValidationAlertStyle。|
||[title](/javascript/api/excel/excel.datavalidationerroralert#title)|表示错误警报对话框标题。|
|[DataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)|[邮件](/javascript/api/excel/excel.datavalidationprompt#message)|指定提示消息。|
||[showPrompt](/javascript/api/excel/excel.datavalidationprompt#showprompt)|指定当用户选择具有数据有效性的单元格时是否显示提示。|
||[title](/javascript/api/excel/excel.datavalidationprompt#title)|指定提示的标题。|
|[DataValidationRule](/javascript/api/excel/excel.datavalidationrule)|[custom](/javascript/api/excel/excel.datavalidationrule#custom)|自定义数据有效性条件。|
||[date](/javascript/api/excel/excel.datavalidationrule#date)|日期数据有效性条件。|
||[decimal](/javascript/api/excel/excel.datavalidationrule#decimal)|小数数据有效性条件。|
||[list](/javascript/api/excel/excel.datavalidationrule#list)|列表数据有效性条件。|
||[textLength](/javascript/api/excel/excel.datavalidationrule#textlength)|TextLength 数据有效性条件。|
||[time](/javascript/api/excel/excel.datavalidationrule#time)|时间数据有效性条件。|
||[wholeNumber](/javascript/api/excel/excel.datavalidationrule#wholenumber)|WholeNumber 数据有效性条件。|
|[DateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|[formula1](/javascript/api/excel/excel.datetimedatavalidation#formula1)|当运算符属性设置为二进制运算符（如 GreaterThan (）时，指定右侧操作数 (左侧操作数是用户尝试在单元格) 中输入的值。|
||[formula2](/javascript/api/excel/excel.datetimedatavalidation#formula2)|使用三元运算符 Between 和 NotBetween 指定上限操作数。|
||[operator](/javascript/api/excel/excel.datetimedatavalidation#operator)|用于验证数据有效性的运算符。|
|[FilterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|[enableMultipleFilterItems](/javascript/api/excel/excel.filterpivothierarchy#enablemultiplefilteritems)|确定是否允许多个筛选项。|
||[name](/javascript/api/excel/excel.filterpivothierarchy#name)|FilterPivotHierarchy 的名称。|
||[position](/javascript/api/excel/excel.filterpivothierarchy#position)|FilterPivotHierarchy 的位置。|
||[fields](/javascript/api/excel/excel.filterpivothierarchy#fields)|返回与 FilterPivotHierarchy 相关联的 PivotFields。|
||[id](/javascript/api/excel/excel.filterpivothierarchy#id)|FilterPivotHierarchy 的 ID。|
||[setToDefault () ](/javascript/api/excel/excel.filterpivothierarchy#settodefault--)|将 FilterPivotHierarchy 重置回其默认值。|
|[FilterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|[添加 (pivotHierarchy：Excel.PivotHierarchy) ](/javascript/api/excel/excel.filterpivothierarchycollection#add-pivothierarchy-)|将 PivotHierarchy 添加到当前轴。|
||[getCount()](/javascript/api/excel/excel.filterpivothierarchycollection#getcount--)|获取集合中的透视层级结构的数量。|
||[getItem(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection#getitem-name-)|按名称或 ID 获取 FilterPivotHierarchy。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection#getitemornullobject-name-)|按名称获取 FilterPivotHierarchy。|
||[items](/javascript/api/excel/excel.filterpivothierarchycollection#items)|获取此集合中已加载的子项。|
||[remove (filterPivotHierarchy： Excel.FilterPivotHierarchy) ](/javascript/api/excel/excel.filterpivothierarchycollection#remove-filterpivothierarchy-)|从当前轴删除 PivotHierarchy。|
|[ListDataValidation](/javascript/api/excel/excel.listdatavalidation)|[inCellDropDown](/javascript/api/excel/excel.listdatavalidation#incelldropdown)|是否显示单元格下拉菜单中的列表，默认为 true。|
||[source](/javascript/api/excel/excel.listdatavalidation#source)|数据有效性列表源|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[name](/javascript/api/excel/excel.pivotfield#name)|PivotField 的名称。|
||[id](/javascript/api/excel/excel.pivotfield#id)|PivotField 的 ID。|
||[项目](/javascript/api/excel/excel.pivotfield#items)|返回与 PivotField 相关联的 PivotFields。|
||[showAllItems](/javascript/api/excel/excel.pivotfield#showallitems)|确定是否显示 PivotField 的所有项。|
||[sortByLabels (sortBy： SortBy) ](/javascript/api/excel/excel.pivotfield#sortbylabels-sortby-)|PivotField 排序。|
||[subtotals](/javascript/api/excel/excel.pivotfield#subtotals)|PivotField 小计。|
|[PivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|[getCount()](/javascript/api/excel/excel.pivotfieldcollection#getcount--)|获取集合中透视字段的数量。|
||[getItem(name: string)](/javascript/api/excel/excel.pivotfieldcollection#getitem-name-)|按名称或 ID 获取 PivotField。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotfieldcollection#getitemornullobject-name-)|按名称获取 PivotField。|
||[items](/javascript/api/excel/excel.pivotfieldcollection#items)|获取此集合中已加载的子项。|
|[PivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|[name](/javascript/api/excel/excel.pivothierarchy#name)|PivotHierarchy 的名称。|
||[fields](/javascript/api/excel/excel.pivothierarchy#fields)|返回与 PivotHierarchy 相关联的 PivotFields。|
||[id](/javascript/api/excel/excel.pivothierarchy#id)|PivotHierarchy 的 ID。|
|[PivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|[getCount()](/javascript/api/excel/excel.pivothierarchycollection#getcount--)|获取集合中的透视层级结构的数量。|
||[getItem(name: string)](/javascript/api/excel/excel.pivothierarchycollection#getitem-name-)|按名称或 ID 获取 PivotHierarchy。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivothierarchycollection#getitemornullobject-name-)|按名称获取 PivotHierarchy。|
||[items](/javascript/api/excel/excel.pivothierarchycollection#items)|获取此集合中已加载的子项。|
|[PivotItem](/javascript/api/excel/excel.pivotitem)|[isExpanded](/javascript/api/excel/excel.pivotitem#isexpanded)|确定是展开项以显示子项还是折叠项并隐藏子项。|
||[name](/javascript/api/excel/excel.pivotitem#name)|PivotItem 的名称。|
||[id](/javascript/api/excel/excel.pivotitem#id)|PivotItem 的 ID。|
||[visible](/javascript/api/excel/excel.pivotitem#visible)|指定 PivotItem 是否可见。|
|[PivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|[getCount()](/javascript/api/excel/excel.pivotitemcollection#getcount--)|获取集合中 PivotItems 的数量。|
||[getItem(name: string)](/javascript/api/excel/excel.pivotitemcollection#getitem-name-)|按名称或 ID 获取 PivotItem。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotitemcollection#getitemornullobject-name-)|按名称获取 PivotItem。|
||[items](/javascript/api/excel/excel.pivotitemcollection#items)|获取此集合中已加载的子项。|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getColumnLabelRange () ](/javascript/api/excel/excel.pivotlayout#getcolumnlabelrange--)|返回数据透视表列标签所在位置的区域。|
||[getDataBodyRange () ](/javascript/api/excel/excel.pivotlayout#getdatabodyrange--)|返回数据透视表数据值所在位置的区域。|
||[getFilterAxisRange () ](/javascript/api/excel/excel.pivotlayout#getfilteraxisrange--)|返回数据透视表筛选区的区域。|
||[getRange()](/javascript/api/excel/excel.pivotlayout#getrange--)|返回存在数据透视表的区域，不包括筛选区。|
||[getRowLabelRange () ](/javascript/api/excel/excel.pivotlayout#getrowlabelrange--)|返回数据透视表行标签所在位置的区域。|
||[layoutType](/javascript/api/excel/excel.pivotlayout#layouttype)|此属性指示数据透视表上的所有字段的 PivotLayoutType。|
||[showColumnGrandTotals](/javascript/api/excel/excel.pivotlayout#showcolumngrandtotals)|指定数据透视表是否显示列的总计。|
||[showRowGrandTotals](/javascript/api/excel/excel.pivotlayout#showrowgrandtotals)|指定数据透视表是否显示行的总计。|
||[subtotalLocation](/javascript/api/excel/excel.pivotlayout#subtotallocation)|此属性指示数据透视表上的所有字段的 SubtotalLocationType。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[delete()](/javascript/api/excel/excel.pivottable#delete--)|删除 PivotTable 对象。|
||[columnHierarchies](/javascript/api/excel/excel.pivottable#columnhierarchies)|数据透视表的列透视层级结构。|
||[dataHierarchies](/javascript/api/excel/excel.pivottable#datahierarchies)|数据透视表的数据透视层级结构。|
||[filterHierarchies](/javascript/api/excel/excel.pivottable#filterhierarchies)|数据透视表的筛选器透视层级结构。|
||[hierarchies](/javascript/api/excel/excel.pivottable#hierarchies)|数据透视表的透视层级结构。|
||[layout](/javascript/api/excel/excel.pivottable#layout)|PivotLayout，用于说明数据透视表的布局和可视化结构。|
||[rowHierarchies](/javascript/api/excel/excel.pivottable#rowhierarchies)|数据透视表的行透视层级结构。|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[add (name： string， source： Range \| string \| Table， destination： Range \| string) ](/javascript/api/excel/excel.pivottablecollection#add-name--source--destination-)|添加基于指定源数据的数据透视表，并将其插入到目标区域左上方的单元格中。|
|[区域](/javascript/api/excel/excel.range)|[dataValidation](/javascript/api/excel/excel.range#datavalidation)|返回数据有效性对象。|
|[RowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|[name](/javascript/api/excel/excel.rowcolumnpivothierarchy#name)|RowColumnPivotHierarchy 的名称。|
||[position](/javascript/api/excel/excel.rowcolumnpivothierarchy#position)|RowColumnPivotHierarchy 的位置。|
||[fields](/javascript/api/excel/excel.rowcolumnpivothierarchy#fields)|返回与 RowColumnPivotHierarchy 相关联的 PivotFields。|
||[id](/javascript/api/excel/excel.rowcolumnpivothierarchy#id)|RowColumnPivotHierarchy 的 ID。|
||[setToDefault () ](/javascript/api/excel/excel.rowcolumnpivothierarchy#settodefault--)|将 RowColumnPivotHierarchy 重置回其默认值。|
|[RowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|[添加 (pivotHierarchy：Excel.PivotHierarchy) ](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#add-pivothierarchy-)|将 PivotHierarchy 添加到当前轴。|
||[getCount()](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getcount--)|获取集合中的透视层级结构的数量。|
||[getItem(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getitem-name-)|按名称或 ID 获取 RowColumnPivotHierarchy。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getitemornullobject-name-)|按名称获取 RowColumnPivotHierarchy。|
||[items](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#items)|获取此集合中已加载的子项。|
||[remove (rowColumnPivotHierarchy： Excel.RowColumnPivotHierarchy) ](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#remove-rowcolumnpivothierarchy-)|从当前轴删除 PivotHierarchy。|
|[运行时](/javascript/api/excel/excel.runtime)|[enableEvents](/javascript/api/excel/excel.runtime#enableevents)|在当前任务窗格或内容加载项中切换 JavaScript 事件。|
|[ShowAsRule](/javascript/api/excel/excel.showasrule)|[baseField](/javascript/api/excel/excel.showasrule#basefield)|基于 ShowAs 计算的基础 PivotField，如适用，基于 ShowAsCalculation 类型，否则为 null。|
||[baseItem](/javascript/api/excel/excel.showasrule#baseitem)|基于 ShowAs 计算的基础 Item，如适用，基于 ShowAsCalculation 类型，否则为 null。|
||[calculation](/javascript/api/excel/excel.showasrule#calculation)|数据 PivotField 使用的 ShowAs 计算。|
|[Style](/javascript/api/excel/excel.style)|[autoIndent](/javascript/api/excel/excel.style#autoindent)|指定当单元格中的文本对齐方式设置为相等分布时文本是否自动缩进。|
||[textOrientation](/javascript/api/excel/excel.style#textorientation)|此样式中的文本方向。|
|[Subtotals](/javascript/api/excel/excel.subtotals)|[automatic](/javascript/api/excel/excel.subtotals#automatic)|如果将“Automatic”设为 true，则在设置 Subtotals 时，所有其他值均会被忽略。|
||[average](/javascript/api/excel/excel.subtotals#average)||
||[count](/javascript/api/excel/excel.subtotals#count)||
||[countNumbers](/javascript/api/excel/excel.subtotals#countnumbers)||
||[max](/javascript/api/excel/excel.subtotals#max)||
||[min](/javascript/api/excel/excel.subtotals#min)||
||[product](/javascript/api/excel/excel.subtotals#product)||
||[standardDeviation](/javascript/api/excel/excel.subtotals#standarddeviation)||
||[standardDeviationP](/javascript/api/excel/excel.subtotals#standarddeviationp)||
||[sum](/javascript/api/excel/excel.subtotals#sum)||
||[variance](/javascript/api/excel/excel.subtotals#variance)||
||[varianceP](/javascript/api/excel/excel.subtotals#variancep)||
|[Table](/javascript/api/excel/excel.table)|[legacyId](/javascript/api/excel/excel.table#legacyid)|返回数字 ID。|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.tablechangedeventargs#getrange-ctx-)|获取表示特定工作表上表格的已更改区域的区域。|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.tablechangedeventargs#getrangeornullobject-ctx-)|获取表示特定工作表上表格的已更改区域的区域。|
|[Workbook](/javascript/api/excel/excel.workbook)|[readOnly](/javascript/api/excel/excel.workbook#readonly)|如果在只读模式下打开工作簿，则为 True。|
|[WorkbookCreated](/javascript/api/excel/excel.workbookcreated)|||
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onCalculated](/javascript/api/excel/excel.worksheet#oncalculated)|计算工作表时发生。|
||[showGridlines](/javascript/api/excel/excel.worksheet#showgridlines)|指定网格线是否对用户可见。|
||[showHeadings](/javascript/api/excel/excel.worksheet#showheadings)|指定标题是否对用户可见。|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[type](/javascript/api/excel/excel.worksheetcalculatedeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.worksheetcalculatedeventargs#worksheetid)|获取进行计算的工作表的 ID。|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#getrange-ctx-)|获取区域，该区域表示特定工作表上的更改区域。|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#getrangeornullobject-ctx-)|获取区域，该区域表示特定工作表上的更改区域。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onCalculated](/javascript/api/excel/excel.worksheetcollection#oncalculated)|计算工作簿中任何工作表时发生。|

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 参考文档](/javascript/api/excel?view=excel-js-1.8&preserve-view=true)
- [Excel JavaScript API 要求集](excel-api-requirement-sets.md)
