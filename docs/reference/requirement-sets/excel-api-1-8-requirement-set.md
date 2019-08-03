---
title: Excel JavaScript API 要求集1。8
description: 有关 ExcelApi 1.8 要求集的详细信息
ms.date: 07/26/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 6849ccb3dc83275509d26c63054a518d41cb060e
ms.sourcegitcommit: 3f5d7f4794e3d3c8bc3a79fa05c54157613b9376
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/02/2019
ms.locfileid: "36064891"
---
# <a name="whats-new-in-excel-javascript-api-18"></a>Excel JavaScript API 1.8 中的新增功能

Excel JavaScript API 要求集 1.8 的功能包括适用于数据透视表、数据验证、图表、图表事件、性能选项和工作簿创建的 API。

## <a name="pivottable"></a>数据透视表

加载项通过数据透视表 API 的波形 2 设置数据透视表的层次结构。 现在可以控制数据及其聚合方式。 [数据透视表](/office/dev/add-ins/excel/excel-add-ins-pivottables)一文详细介绍了新的数据透视表功能。

## <a name="data-validation"></a>数据有效性

数据有效性可以控制用户在工作表中输入的内容。 可以将单元格限制为预定义的答案集，或者在用户输入无效数据时提供弹出警告。 立即详细了解[向区域添加数据有效性](/office/dev/add-ins/excel/excel-add-ins-data-validation)。

## <a name="charts"></a>图表

另一轮图表 API 可更好地对图表元素进行编程控制。 现在，你对图例、坐标轴、趋势线和绘图区拥有更高的访问权限。

## <a name="events"></a>事件

已为图表添加更多[事件](/office/dev/add-ins/excel/excel-add-ins-events)。 让加载项处理用于与图表的交互。 此外，你还可以在整个工作簿中[触发事件](/office/dev/add-ins/excel/performance#enable-and-disable-events)。

## <a name="api-list"></a>API 列表

下表列出了 Excel JavaScript API 要求集1.8 中的 Api。 若要查看 Excel JavaScript API 要求集1.8 或更早版本支持的所有 Api 的 API 参考文档, 请参阅[要求集1.8 或更早版本中的 Excel api](/javascript/api/excel?view=excel-js-1.8)。

| Class | 域 | 说明 |
|:---|:---|:---|
|[BasicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|[formula1](/javascript/api/excel/excel.basicdatavalidation#formula1)|当 operator 属性设置为二元运算符 (如 GreaterThan (左边的操作数是用户试图在单元格中输入的值) 时, 指定右边的操作数。 使用和 NotBetween 之间的三元运算符指定下界操作数。|
||[formula2](/javascript/api/excel/excel.basicdatavalidation#formula2)|使用和 NotBetween 之间的三元运算符指定上界操作数。 不与二元运算符 (如 GreaterThan) 一起使用。|
||[接线员](/javascript/api/excel/excel.basicdatavalidation#operator)|用于验证数据有效性的运算符。|
|[Chart](/javascript/api/excel/excel.chart)|[categoryLabelLevel](/javascript/api/excel/excel.chart#categorylabellevel)|返回或设置一个 ChartCategoryLabelLevel 枚举常量, 该常量引用|
||[displayBlanksAs](/javascript/api/excel/excel.chart#displayblanksas)|返回或设置图表上的空白单元格的绘制方式。 读/写。|
||[plotBy](/javascript/api/excel/excel.chart#plotby)|返回或设置图表上的列或行用作数据系列的方式。 读/写。|
||[plotVisibleOnly](/javascript/api/excel/excel.chart#plotvisibleonly)|如果仅绘制可见单元格，则为 True。如果绘制可见单元格和隐藏单元格，则为 False。 读/写。|
||[onActivated](/javascript/api/excel/excel.chart#onactivated)|图表激活时发生。|
||[onDeactivated](/javascript/api/excel/excel.chart#ondeactivated)|图表停用时发生。|
||[plotArea](/javascript/api/excel/excel.chart#plotarea)|表示图表的绘制区域。|
||[seriesNameLevel](/javascript/api/excel/excel.chart#seriesnamelevel)|返回或设置一个 ChartSeriesNameLevel 枚举常量, 该常量引用|
||[showDataLabelsOverMaximum](/javascript/api/excel/excel.chart#showdatalabelsovermaximum)|表示当值大于数值轴上的最大值时是否显示数据标签。|
||[style](/javascript/api/excel/excel.chart#style)|返回或设置图表的图表样式。 读/写。|
|[ChartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|[chartId](/javascript/api/excel/excel.chartactivatedeventargs#chartid)|获取已启用图表的 ID。|
||[type](/javascript/api/excel/excel.chartactivatedeventargs#type)|获取事件的类型。 有关详细信息，请参阅 Excel.EventType。|
||[worksheetId](/javascript/api/excel/excel.chartactivatedeventargs#worksheetid)|获取其中的图表已启用的工作表的 ID。|
|[ChartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|[chartId](/javascript/api/excel/excel.chartaddedeventargs#chartid)|获取已添加至工作表的图表的 ID。|
||[source](/javascript/api/excel/excel.chartaddedeventargs#source)|获取事件源。 有关详细信息，请参阅 Excel.EventSource。|
||[type](/javascript/api/excel/excel.chartaddedeventargs#type)|获取事件的类型。 有关详细信息，请参阅 Excel.EventType。|
||[worksheetId](/javascript/api/excel/excel.chartaddedeventargs#worksheetid)|获取已在其中添加图表的工作表的 ID。|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[对齐方式](/javascript/api/excel/excel.chartaxis#alignment)|表示指定轴刻度线标签的对齐方式。 有关详细信息, 请参阅 ChartTextHorizontalAlignment。|
||[isBetweenCategories](/javascript/api/excel/excel.chartaxis#isbetweencategories)|表示数值轴是否与分类之间的分类轴交叉。|
||[符号](/javascript/api/excel/excel.chartaxis#multilevel)|表示是否为多级轴。|
||[numberFormat](/javascript/api/excel/excel.chartaxis#numberformat)|表示轴刻度线标签的格式代码。|
||[一定](/javascript/api/excel/excel.chartaxis#offset)|表示不同标签级别之间的距离以及一级标签和轴线之间的距离。 此值应该是 0 到 1000 之间的整数。|
||[position](/javascript/api/excel/excel.chartaxis#position)|表示两轴交叉的特定轴位置。 有关详细信息, 请参阅 ChartAxisPosition。|
||[positionAt](/javascript/api/excel/excel.chartaxis#positionat)|表示两轴交叉的特定轴位置。 应使用 SetPositionAt(double) 方法设置此属性。|
||[setPositionAt (value: 数字)](/javascript/api/excel/excel.chartaxis#setpositionat-value-)|设置两轴交叉的特定轴位置。|
||[textOrientation](/javascript/api/excel/excel.chartaxis#textorientation)|表示轴刻度线标签的文本方向。 此值应是 -90 到 90 或 180（垂直文本）之间的整数。|
|[ChartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|[fill](/javascript/api/excel/excel.chartaxisformat#fill)|表示图表填充格式。 只读。|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[setFormula (formula: string)](/javascript/api/excel/excel.chartaxistitle#setformula-formula-)|该字符串值表示采用 A1 表示法的图表轴标题的公式。|
|[ChartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|[边缘](/javascript/api/excel/excel.chartaxistitleformat#border)|表示边框格式，包括颜色、线条样式和粗细。|
||[fill](/javascript/api/excel/excel.chartaxistitleformat#fill)|表示图表填充格式。|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[clear()](/javascript/api/excel/excel.chartborder#clear--)|清除图表元素的边框格式。|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[onActivated](/javascript/api/excel/excel.chartcollection#onactivated)|在激活图表时发生。|
||[onAdded](/javascript/api/excel/excel.chartcollection#onadded)|将新图表添加到工作表时发生此事件。|
||[onDeactivated](/javascript/api/excel/excel.chartcollection#ondeactivated)|当停用图表时发生此事件。|
||[onDeleted](/javascript/api/excel/excel.chartcollection#ondeleted)|在删除图表时发生。|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[自动图文集](/javascript/api/excel/excel.chartdatalabel#autotext)|该布尔值表示数据标签是否根据上下文自动生成相应的文本。|
||[formula](/javascript/api/excel/excel.chartdatalabel#formula)|该字符串值表示采用 A1 表示法的图表数据标签的公式。|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabel#horizontalalignment)|表示图表数据标签水平对齐。 有关详细信息, 请参阅 ChartTextHorizontalAlignment。|
||[left](/javascript/api/excel/excel.chartdatalabel#left)|表示图表数据标签左边缘到图表区域左边缘的距离，以磅为单位。 如果图表数据标签不可见，则为 Null。|
||[numberFormat](/javascript/api/excel/excel.chartdatalabel#numberformat)|该字符串值表示数据标签的格式代码。|
||[format](/javascript/api/excel/excel.chartdatalabel#format)|表示图表数据标签的格式。|
||[height](/javascript/api/excel/excel.chartdatalabel#height)|返回图表数据标签的高度，以磅为单位。 只读。 如果图表数据标签不可见，则为 Null。|
||[width](/javascript/api/excel/excel.chartdatalabel#width)|返回图表数据标签的宽度，以磅为单位。 只读。 如果图表数据标签不可见，则为 Null。|
||[text](/javascript/api/excel/excel.chartdatalabel#text)|该字符串表示图表上的数据标签文本。|
||[textOrientation](/javascript/api/excel/excel.chartdatalabel#textorientation)|表示图表数据标签的文本方向。 此值应是 -90 到 90 或 180（垂直文本）之间的整数。|
||[top](/javascript/api/excel/excel.chartdatalabel#top)|表示图表数据标签上边缘到图表区域顶部的距离，以磅为单位。 如果图表数据标签不可见，则为 Null。|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabel#verticalalignment)|表示图表数据标签垂直对齐。 有关详细信息, 请参阅 ChartTextVerticalAlignment。|
|[ChartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|[边缘](/javascript/api/excel/excel.chartdatalabelformat#border)|表示边框格式，包括颜色、线条样式和粗细。 只读。|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[自动图文集](/javascript/api/excel/excel.chartdatalabels#autotext)|表示数据标签是否根据上下文自动生成相应的文本。|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabels#horizontalalignment)|表示图表数据标签水平对齐。 有关详细信息, 请参阅 ChartTextHorizontalAlignment。|
||[numberFormat](/javascript/api/excel/excel.chartdatalabels#numberformat)|表示数据标签的格式代码。|
||[textOrientation](/javascript/api/excel/excel.chartdatalabels#textorientation)|表示数据标签的文本方向。 此值应是 -90 到 90 或 180（垂直文本）之间的整数。|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabels#verticalalignment)|表示图表数据标签垂直对齐。 有关详细信息, 请参阅 ChartTextVerticalAlignment。|
|[ChartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|[chartId](/javascript/api/excel/excel.chartdeactivatedeventargs#chartid)|获取停用图表的 ID。|
||[type](/javascript/api/excel/excel.chartdeactivatedeventargs#type)|获取事件的类型。 有关详细信息，请参阅 Excel.EventType。|
||[worksheetId](/javascript/api/excel/excel.chartdeactivatedeventargs#worksheetid)|获取其中的图表已停用的工作表的 ID。|
|[ChartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|[chartId](/javascript/api/excel/excel.chartdeletedeventargs#chartid)|获取已从工作表删除的图表的 ID。|
||[source](/javascript/api/excel/excel.chartdeletedeventargs#source)|获取事件源。 有关详细信息，请参阅 Excel.EventSource。|
||[type](/javascript/api/excel/excel.chartdeletedeventargs#type)|获取事件的类型。 有关详细信息，请参阅 Excel.EventType。|
||[worksheetId](/javascript/api/excel/excel.chartdeletedeventargs#worksheetid)|获取已在其中删除图表的工作表的 ID。|
|[ChartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|[height](/javascript/api/excel/excel.chartlegendentry#height)|表示图表图例上的 legendEntry 的高度。|
||[index](/javascript/api/excel/excel.chartlegendentry#index)|表示图表图例中的 legendEntry 的索引。|
||[left](/javascript/api/excel/excel.chartlegendentry#left)|表示图表 legendEntry 的左侧。|
||[top](/javascript/api/excel/excel.chartlegendentry#top)|表示图表 legendEntry 的顶部。|
||[width](/javascript/api/excel/excel.chartlegendentry#width)|表示图表图例上的 legendEntry 的宽度。|
|[ChartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|[边缘](/javascript/api/excel/excel.chartlegendformat#border)|表示边框格式，包括颜色、线条样式和粗细。 只读。|
|[ChartPlotArea](/javascript/api/excel/excel.chartplotarea)|[height](/javascript/api/excel/excel.chartplotarea#height)|表示 plotArea 的高度值。|
||[insideHeight](/javascript/api/excel/excel.chartplotarea#insideheight)|表示 plotArea 的 insideHeight 值。|
||[insideLeft](/javascript/api/excel/excel.chartplotarea#insideleft)|表示 plotArea 的 insideLeft 值。|
||[insideTop](/javascript/api/excel/excel.chartplotarea#insidetop)|表示 plotArea 的 insideTop 值。|
||[insideWidth](/javascript/api/excel/excel.chartplotarea#insidewidth)|表示 plotArea 的 insideWidth 值。|
||[left](/javascript/api/excel/excel.chartplotarea#left)|表示 plotArea 的 left 值。|
||[position](/javascript/api/excel/excel.chartplotarea#position)|表示 plotArea 的位置。|
||[format](/javascript/api/excel/excel.chartplotarea#format)|表示图表 plotArea 的格式。|
||[top](/javascript/api/excel/excel.chartplotarea#top)|表示 plotArea 的 top 值。|
||[width](/javascript/api/excel/excel.chartplotarea#width)|表示 plotArea 的宽度值。|
|[ChartPlotAreaFormat](/javascript/api/excel/excel.chartplotareaformat)|[边缘](/javascript/api/excel/excel.chartplotareaformat#border)|表示图表 plotArea 的边框属性。|
||[fill](/javascript/api/excel/excel.chartplotareaformat#fill)|表示对象的填充格式，包括背景格式信息。|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[axisGroup](/javascript/api/excel/excel.chartseries#axisgroup)|返回或设置指定系列的组。 读/写|
||[分离](/javascript/api/excel/excel.chartseries#explosion)|返回或设置饼图或圆环图切片的分解程度值。 如果未分解（切片尖端位于饼图中心），则返回 0（零）。 读/写。|
||[firstSliceAngle](/javascript/api/excel/excel.chartseries#firstsliceangle)|返回或设置第一个饼图或圆环图切片的角度，以度为单位（从垂直方向起为顺时针）。 只适用于饼图、三维饼图和圆环图。 可以是 0 到 360 之间的值。 读/写|
||[invertIfNegative](/javascript/api/excel/excel.chartseries#invertifnegative)|如果 Excel 在值为负数时反转项目中的图案，则为 True。 读/写。|
||[比例](/javascript/api/excel/excel.chartseries#overlap)|指定条柱的摆放方式。 可以是–100到100之间的值。 只适用于二维条形图和二维柱形图。 读/写。|
||[dataLabels](/javascript/api/excel/excel.chartseries#datalabels)|表示系列中所有数据标签的集合。|
||[secondPlotSize](/javascript/api/excel/excel.chartseries#secondplotsize)|返回或设置复合饼图或复合条饼图的辅助分区的大小，以占主饼图大小的百分比表示。 可以是 5 到 200 之间的值。 读/写。|
||[splitType](/javascript/api/excel/excel.chartseries#splittype)|返回或设置复合饼图或复合条饼图中两个分区的拆分方式。 读/写。|
||[varyByCategories](/javascript/api/excel/excel.chartseries#varybycategories)|如果 Excel 为每个数据标记分配不同的颜色或图案，则为 True。 图表必须只包含一个系列。 读/写。|
|[ChartTrendline](/javascript/api/excel/excel.charttrendline)|[backwardPeriod](/javascript/api/excel/excel.charttrendline#backwardperiod)|表示趋势线向后延伸的周期数。|
||[forwardPeriod](/javascript/api/excel/excel.charttrendline#forwardperiod)|表示趋势线向前延伸的周期数。|
||[标志](/javascript/api/excel/excel.charttrendline#label)|表示图表趋势线的标签。|
||[showEquation](/javascript/api/excel/excel.charttrendline#showequation)|如果图表上显示趋势线公式，则为 True。|
||[showRSquared](/javascript/api/excel/excel.charttrendline#showrsquared)|如果图表上显示趋势线的 R 平方值，则为 True。|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[自动图文集](/javascript/api/excel/excel.charttrendlinelabel#autotext)|该布尔值表示趋势线标签是否根据上下文自动生成相应的文本。|
||[formula](/javascript/api/excel/excel.charttrendlinelabel#formula)|该字符串值表示采用 A1 表示法的图表趋势线标签的公式。|
||[horizontalAlignment](/javascript/api/excel/excel.charttrendlinelabel#horizontalalignment)|表示图表趋势线标签水平对齐。 有关详细信息, 请参阅 ChartTextHorizontalAlignment。|
||[left](/javascript/api/excel/excel.charttrendlinelabel#left)|表示图表趋势线标签左边缘到图表区域左边缘的距离，以磅为单位。 如果图表趋势线标签不可见，则为 Null。|
||[numberFormat](/javascript/api/excel/excel.charttrendlinelabel#numberformat)|该字符串值表示趋势线标签的格式代码。|
||[format](/javascript/api/excel/excel.charttrendlinelabel#format)|表示图表趋势线标签的格式。|
||[height](/javascript/api/excel/excel.charttrendlinelabel#height)|返回图表趋势线标签的高度，以磅为单位。 只读。 如果图表趋势线标签不可见，则为 Null。|
||[width](/javascript/api/excel/excel.charttrendlinelabel#width)|返回图表趋势线标签的宽度，以磅为单位。 只读。 如果图表趋势线标签不可见，则为 Null。|
||[text](/javascript/api/excel/excel.charttrendlinelabel#text)|该字符串表示图表上的趋势线标签文本。|
||[textOrientation](/javascript/api/excel/excel.charttrendlinelabel#textorientation)|表示图表趋势线标签的文本方向。 此值应是 -90 到 90 或 180（垂直文本）之间的整数。|
||[top](/javascript/api/excel/excel.charttrendlinelabel#top)|表示图表趋势线标签上边缘到图表区域顶部的距离，以磅为单位。 如果图表趋势线标签不可见，则为 Null。|
||[verticalAlignment](/javascript/api/excel/excel.charttrendlinelabel#verticalalignment)|表示图表趋势线标签垂直对齐。 有关详细信息, 请参阅 ChartTextVerticalAlignment。|
|[ChartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|[边缘](/javascript/api/excel/excel.charttrendlinelabelformat#border)|表示边框格式，包括颜色、线条样式和粗细。|
||[fill](/javascript/api/excel/excel.charttrendlinelabelformat#fill)|表示当前图表趋势线标签的填充格式。|
||[font](/javascript/api/excel/excel.charttrendlinelabelformat#font)|表示图表趋势线标签的字体属性（字体名称、字体大小、颜色等）。|
|[CustomDataValidation](/javascript/api/excel/excel.customdatavalidation)|[formula](/javascript/api/excel/excel.customdatavalidation#formula)|自定义数据验证公式。 这将创建特殊的输入规则, 如阻止重复项或限制单元格范围中的总计。|
|[DataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|[name](/javascript/api/excel/excel.datapivothierarchy#name)|DataPivotHierarchy 的名称。|
||[numberFormat](/javascript/api/excel/excel.datapivothierarchy#numberformat)|DataPivotHierarchy 的数字格式。|
||[position](/javascript/api/excel/excel.datapivothierarchy#position)|DataPivotHierarchy 的位置。|
||[field](/javascript/api/excel/excel.datapivothierarchy#field)|返回与 DataPivotHierarchy 相关联的 PivotFields。|
||[id](/javascript/api/excel/excel.datapivothierarchy#id)|DataPivotHierarchy ID。|
||[setToDefault()](/javascript/api/excel/excel.datapivothierarchy#settodefault--)|将 DataPivotHierarchy 重置回其默认值。|
||[showAs](/javascript/api/excel/excel.datapivothierarchy#showas)|确定数据是否应显示为特定计算汇总。|
||[summarizeBy](/javascript/api/excel/excel.datapivothierarchy#summarizeby)|确定是否显示 DataPivotHierarchy 的所有项。|
|[DataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|[add (pivotHierarchy: PivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection#add-pivothierarchy-)|将 PivotHierarchy 添加到当前轴。|
||[getCount()](/javascript/api/excel/excel.datapivothierarchycollection#getcount--)|获取集合中的透视层级结构的数量。|
||[getItem(name: string)](/javascript/api/excel/excel.datapivothierarchycollection#getitem-name-)|按名称或 ID 获取 DataPivotHierarchy。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.datapivothierarchycollection#getitemornullobject-name-)|按名称获取 DataPivotHierarchy。 如果 DataPivotHierarchy 不存在，则返回 Null 对象。|
||[items](/javascript/api/excel/excel.datapivothierarchycollection#items)|获取此集合中已加载的子项。|
||[remove (DataPivotHierarchy: DataPivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection#remove-datapivothierarchy-)|从当前轴删除 PivotHierarchy。|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[clear()](/javascript/api/excel/excel.datavalidation#clear--)|清除当前区域中的数据有效性。|
||[errorAlert](/javascript/api/excel/excel.datavalidation#erroralert)|用户输入无效数据时，出现错误警报。|
||[ignoreBlanks](/javascript/api/excel/excel.datavalidation#ignoreblanks)|忽略空白：不会对空白单元格执行数据严重，默认为 true。|
||[prompt](/javascript/api/excel/excel.datavalidation#prompt)|当用户选择单元格时提示。|
||[type](/javascript/api/excel/excel.datavalidation#type)|数据有效性类型，有关详细信息，请参阅 Excel.DataValidationType。|
||[有效](/javascript/api/excel/excel.datavalidation#valid)|表示所有单元格值根据数据有效性规则是否全部有效。|
||[标尺](/javascript/api/excel/excel.datavalidation#rule)|包含不同类型的数据验证条件的数据有效性规则。|
|[DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|[邮件](/javascript/api/excel/excel.datavalidationerroralert#message)|表示错误警报消息。|
||[showAlert](/javascript/api/excel/excel.datavalidationerroralert#showalert)|确定在用户输入无效数据时是否显示错误警报对话框。 默认值为 true。|
||[style](/javascript/api/excel/excel.datavalidationerroralert#style)|表示数据有效性警报类型，有关详细信息，请参阅 Excel.DataValidationAlertStyle。|
||[title](/javascript/api/excel/excel.datavalidationerroralert#title)|表示错误警报对话框标题。|
|[DataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)|[邮件](/javascript/api/excel/excel.datavalidationprompt#message)|表示提示消息。|
||[showPrompt](/javascript/api/excel/excel.datavalidationprompt#showprompt)|确定在用户选择具有数据有效性的单元格时是否显示提示。|
||[title](/javascript/api/excel/excel.datavalidationprompt#title)|表示提示标题。|
|[DataValidationRule](/javascript/api/excel/excel.datavalidationrule)|[自](/javascript/api/excel/excel.datavalidationrule#custom)|自定义数据有效性条件。|
||[date](/javascript/api/excel/excel.datavalidationrule#date)|日期数据有效性条件。|
||[数位](/javascript/api/excel/excel.datavalidationrule#decimal)|小数数据有效性条件。|
||[list](/javascript/api/excel/excel.datavalidationrule#list)|列表数据有效性条件。|
||[textLength](/javascript/api/excel/excel.datavalidationrule#textlength)|TextLength 数据有效性条件。|
||[time](/javascript/api/excel/excel.datavalidationrule#time)|时间数据有效性条件。|
||[wholeNumber](/javascript/api/excel/excel.datavalidationrule#wholenumber)|WholeNumber 数据有效性条件。|
|[DateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|[formula1](/javascript/api/excel/excel.datetimedatavalidation#formula1)|当 operator 属性设置为二元运算符 (如 GreaterThan (左边的操作数是用户试图在单元格中输入的值) 时, 指定右边的操作数。 使用和 NotBetween 之间的三元运算符指定下界操作数。|
||[formula2](/javascript/api/excel/excel.datetimedatavalidation#formula2)|使用和 NotBetween 之间的三元运算符指定上界操作数。 不与二元运算符 (如 GreaterThan) 一起使用。|
||[接线员](/javascript/api/excel/excel.datetimedatavalidation#operator)|用于验证数据有效性的运算符。|
|[FilterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|[enableMultipleFilterItems](/javascript/api/excel/excel.filterpivothierarchy#enablemultiplefilteritems)|确定是否允许多个筛选项。|
||[name](/javascript/api/excel/excel.filterpivothierarchy#name)|FilterPivotHierarchy 的名称。|
||[position](/javascript/api/excel/excel.filterpivothierarchy#position)|FilterPivotHierarchy 的位置。|
||[fields](/javascript/api/excel/excel.filterpivothierarchy#fields)|返回与 FilterPivotHierarchy 相关联的 PivotFields。|
||[id](/javascript/api/excel/excel.filterpivothierarchy#id)|FilterPivotHierarchy 的 ID。|
||[setToDefault()](/javascript/api/excel/excel.filterpivothierarchy#settodefault--)|将 FilterPivotHierarchy 重置回其默认值。|
|[FilterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|[add (pivotHierarchy: PivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection#add-pivothierarchy-)|将 PivotHierarchy 添加到当前轴。 行和列上的其他位置是否存在层次结构。|
||[getCount()](/javascript/api/excel/excel.filterpivothierarchycollection#getcount--)|获取集合中的透视层级结构的数量。|
||[getItem(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection#getitem-name-)|按名称或 ID 获取 FilterPivotHierarchy。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection#getitemornullobject-name-)|按名称获取 FilterPivotHierarchy。 如果 FilterPivotHierarchy 不存在，则返回 Null 对象。|
||[items](/javascript/api/excel/excel.filterpivothierarchycollection#items)|获取此集合中已加载的子项。|
||[remove (filterPivotHierarchy: FilterPivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection#remove-filterpivothierarchy-)|从当前轴删除 PivotHierarchy。|
|[ListDataValidation](/javascript/api/excel/excel.listdatavalidation)|[inCellDropDown](/javascript/api/excel/excel.listdatavalidation#incelldropdown)|是否显示单元格下拉菜单中的列表，默认为 true。|
||[source](/javascript/api/excel/excel.listdatavalidation#source)|数据有效性列表源|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[name](/javascript/api/excel/excel.pivotfield#name)|PivotField 的名称。|
||[id](/javascript/api/excel/excel.pivotfield#id)|PivotField 的 ID。|
||[items](/javascript/api/excel/excel.pivotfield#items)|返回包含透视字段的 PivotItems。|
||[showAllItems](/javascript/api/excel/excel.pivotfield#showallitems)|确定是否显示 PivotField 的所有项。|
||[sortByLabels (sortBy: SortBy)](/javascript/api/excel/excel.pivotfield#sortbylabels-sortby-)|PivotField 排序。 如果指定 DataPivotHierarchy，则会基于它进行排序，如果未指定，则会基于 PivotField 本身进行排序。|
||[subtotals](/javascript/api/excel/excel.pivotfield#subtotals)|PivotField 小计。|
|[PivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|[getCount()](/javascript/api/excel/excel.pivotfieldcollection#getcount--)|获取集合中的数据透视字段数。|
||[getItem(name: string)](/javascript/api/excel/excel.pivotfieldcollection#getitem-name-)|按其名称或 id 获取透视字段。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotfieldcollection#getitemornullobject-name-)|按名称获取透视字段。 如果透视字段不存在, 则将返回 null 对象。|
||[items](/javascript/api/excel/excel.pivotfieldcollection#items)|获取此集合中已加载的子项。|
|[PivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|[name](/javascript/api/excel/excel.pivothierarchy#name)|PivotHierarchy 的名称。|
||[fields](/javascript/api/excel/excel.pivothierarchy#fields)|返回与 PivotHierarchy 相关联的 PivotFields。|
||[id](/javascript/api/excel/excel.pivothierarchy#id)|PivotHierarchy 的 ID。|
|[PivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|[getCount()](/javascript/api/excel/excel.pivothierarchycollection#getcount--)|获取集合中的透视层级结构的数量。|
||[getItem(name: string)](/javascript/api/excel/excel.pivothierarchycollection#getitem-name-)|按名称或 ID 获取 PivotHierarchy。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivothierarchycollection#getitemornullobject-name-)|按名称获取 PivotHierarchy。 如果 PivotHierarchy 不存在，则返回 Null 对象。|
||[items](/javascript/api/excel/excel.pivothierarchycollection#items)|获取此集合中已加载的子项。|
|[PivotItem](/javascript/api/excel/excel.pivotitem)|[isExpanded](/javascript/api/excel/excel.pivotitem#isexpanded)|确定是展开项以显示子项还是折叠项并隐藏子项。|
||[name](/javascript/api/excel/excel.pivotitem#name)|PivotItem 的名称。|
||[id](/javascript/api/excel/excel.pivotitem#id)|PivotItem 的 ID。|
||[visible](/javascript/api/excel/excel.pivotitem#visible)|确定 PivotItem 是否可见。|
|[PivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|[getCount()](/javascript/api/excel/excel.pivotitemcollection#getcount--)|获取集合中的数据透视项的数目。|
||[getItem(name: string)](/javascript/api/excel/excel.pivotitemcollection#getitem-name-)|按其名称或 id 获取 PivotItem。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotitemcollection#getitemornullobject-name-)|按名称获取 PivotItem。 如果 PivotItem 不存在, 则将返回一个 null 对象。|
||[items](/javascript/api/excel/excel.pivotitemcollection#items)|获取此集合中已加载的子项。|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getColumnLabelRange()](/javascript/api/excel/excel.pivotlayout#getcolumnlabelrange--)|返回数据透视表列标签所在位置的区域。|
||[getDataBodyRange()](/javascript/api/excel/excel.pivotlayout#getdatabodyrange--)|返回数据透视表数据值所在位置的区域。|
||[getFilterAxisRange()](/javascript/api/excel/excel.pivotlayout#getfilteraxisrange--)|返回数据透视表筛选区的区域。|
||[getRange()](/javascript/api/excel/excel.pivotlayout#getrange--)|返回存在数据透视表的区域，不包括筛选区。|
||[getRowLabelRange()](/javascript/api/excel/excel.pivotlayout#getrowlabelrange--)|返回数据透视表行标签所在位置的区域。|
||[layoutType](/javascript/api/excel/excel.pivotlayout#layouttype)|此属性指示数据透视表上的所有字段的 PivotLayoutType。 如果字段状态不同，则为 null。|
||[showColumnGrandTotals](/javascript/api/excel/excel.pivotlayout#showcolumngrandtotals)|指定数据透视表报表是否显示列总计。|
||[showRowGrandTotals](/javascript/api/excel/excel.pivotlayout#showrowgrandtotals)|指定数据透视表报表是否显示行总计。|
||[subtotalLocation](/javascript/api/excel/excel.pivotlayout#subtotallocation)|此属性指示数据透视表上的所有字段的 SubtotalLocationType。 如果字段状态不同，则为 null。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[delete()](/javascript/api/excel/excel.pivottable#delete--)|删除 PivotTable 对象。|
||[columnHierarchies](/javascript/api/excel/excel.pivottable#columnhierarchies)|数据透视表的列透视层级结构。|
||[dataHierarchies](/javascript/api/excel/excel.pivottable#datahierarchies)|数据透视表的数据透视层级结构。|
||[filterHierarchies](/javascript/api/excel/excel.pivottable#filterhierarchies)|数据透视表的筛选器透视层级结构。|
||[层次结构](/javascript/api/excel/excel.pivottable#hierarchies)|数据透视表的透视层级结构。|
||[布局](/javascript/api/excel/excel.pivottable#layout)|PivotLayout，用于说明数据透视表的布局和可视化结构。|
||[rowHierarchies](/javascript/api/excel/excel.pivottable#rowhierarchies)|数据透视表的行透视层级结构。|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[add (name: string, source: Range \| string \| Table, destination: Range \| string)](/javascript/api/excel/excel.pivottablecollection#add-name--source--destination-)|基于指定的数据源添加数据透视表，并将其插入到目标区域的左上单元格。|
|[Range](/javascript/api/excel/excel.range)|[dataValidation](/javascript/api/excel/excel.range#datavalidation)|返回数据有效性对象。|
|[RowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|[name](/javascript/api/excel/excel.rowcolumnpivothierarchy#name)|RowColumnPivotHierarchy 的名称。|
||[position](/javascript/api/excel/excel.rowcolumnpivothierarchy#position)|RowColumnPivotHierarchy 的位置。|
||[fields](/javascript/api/excel/excel.rowcolumnpivothierarchy#fields)|返回与 RowColumnPivotHierarchy 相关联的 PivotFields。|
||[id](/javascript/api/excel/excel.rowcolumnpivothierarchy#id)|RowColumnPivotHierarchy 的 ID。|
||[setToDefault()](/javascript/api/excel/excel.rowcolumnpivothierarchy#settodefault--)|将 RowColumnPivotHierarchy 重置回其默认值。|
|[RowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|[add (pivotHierarchy: PivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#add-pivothierarchy-)|将 PivotHierarchy 添加到当前轴。 行和列上的其他位置是否存在层次结构。|
||[getCount()](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getcount--)|获取集合中的透视层级结构的数量。|
||[getItem(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getitem-name-)|按名称或 ID 获取 RowColumnPivotHierarchy。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getitemornullobject-name-)|按名称获取 RowColumnPivotHierarchy。 如果 RowColumnPivotHierarchy 不存在，则返回 Null 对象。|
||[items](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#items)|获取此集合中已加载的子项。|
||[remove (rowColumnPivotHierarchy: RowColumnPivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#remove-rowcolumnpivothierarchy-)|从当前轴删除 PivotHierarchy。|
|[语言](/javascript/api/excel/excel.runtime)|[enableEvents](/javascript/api/excel/excel.runtime#enableevents)|切换当前任务窗格或内容加载项中的 JavaScript 事件。|
|[ShowAsRule](/javascript/api/excel/excel.showasrule)|[baseField](/javascript/api/excel/excel.showasrule#basefield)|基于 ShowAs 计算的基础 PivotField，如适用，基于 ShowAsCalculation 类型，否则为 null。|
||[baseItem](/javascript/api/excel/excel.showasrule#baseitem)|基于 ShowAs 计算的基础 Item，如适用，基于 ShowAsCalculation 类型，否则为 null。|
||[结果](/javascript/api/excel/excel.showasrule#calculation)|数据 PivotField 使用的 ShowAs 计算。 有关详细信息, 请参阅 ShowAsCalculation。|
|[Style](/javascript/api/excel/excel.style)|[autoIndent](/javascript/api/excel/excel.style#autoindent)|指示将单元格中的文本对齐方式设为相等分布时文本是否会自动缩进。|
||[textOrientation](/javascript/api/excel/excel.style#textorientation)|此样式中的文本方向。|
|[Subtotals](/javascript/api/excel/excel.subtotals)|[automatic](/javascript/api/excel/excel.subtotals#automatic)|如果将“Automatic”设为 true，则在设置 Subtotals 时，所有其他值均会被忽略。|
||[平均](/javascript/api/excel/excel.subtotals#average)||
||[count](/javascript/api/excel/excel.subtotals#count)||
||[countNumbers](/javascript/api/excel/excel.subtotals#countnumbers)||
||[max](/javascript/api/excel/excel.subtotals#max)||
||[min](/javascript/api/excel/excel.subtotals#min)||
||[装箱](/javascript/api/excel/excel.subtotals#product)||
||[standardDeviation](/javascript/api/excel/excel.subtotals#standarddeviation)||
||[standardDeviationP](/javascript/api/excel/excel.subtotals#standarddeviationp)||
||[总值](/javascript/api/excel/excel.subtotals#sum)||
||[差](/javascript/api/excel/excel.subtotals#variance)||
||[varianceP](/javascript/api/excel/excel.subtotals#variancep)||
|[Table](/javascript/api/excel/excel.table)|[legacyId](/javascript/api/excel/excel.table#legacyid)|返回一个数字 id。|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.tablechangedeventargs#getrange-ctx-)|获取表示特定工作表上的表的更改区域的范围。|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.tablechangedeventargs#getrangeornullobject-ctx-)|获取表示特定工作表上的表的更改区域的范围。 它可能会返回 null 对象。|
|[Workbook](/javascript/api/excel/excel.workbook)|[readOnly](/javascript/api/excel/excel.workbook#readonly)|如果在只读模式下打开工作簿，则为 True。 只读。|
|[WorkbookCreated](/javascript/api/excel/excel.workbookcreated)||[Worksheet](/javascript/api/excel/excel.worksheet)|[onCalculated](/javascript/api/excel/excel.worksheet#oncalculated)|在计算工作表时发生。|
||[showGridlines](/javascript/api/excel/excel.worksheet#showgridlines)|获取或设置工作表的网格线标志。|
||[showHeadings](/javascript/api/excel/excel.worksheet#showheadings)|获取或设置工作表的标题标志。|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[type](/javascript/api/excel/excel.worksheetcalculatedeventargs#type)|获取事件的类型。 有关详细信息，请参阅 Excel.EventType。|
||[worksheetId](/javascript/api/excel/excel.worksheetcalculatedeventargs#worksheetid)|获取计算的工作表的 ID。|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#getrange-ctx-)|获取区域，该区域表示特定工作表上的更改区域。|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#getrangeornullobject-ctx-)|获取区域，该区域表示特定工作表上的更改区域。 它可能会返回 null 对象。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onCalculated](/javascript/api/excel/excel.worksheetcollection#oncalculated)|计算工作簿中的任何工作表时发生。|

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 参考文档](/javascript/api/excel?view=excel-js-1.8)
- [Excel JavaScript API 要求集](./excel-api-requirement-sets.md)
