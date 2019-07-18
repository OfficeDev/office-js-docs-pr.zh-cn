---
title: Excel JavaScript API 要求集1。8
description: 有关 ExcelApi 1.8 要求集的详细信息
ms.date: 07/11/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: a5adcf56654070ca2a8336385f73062c34e90e1d
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/17/2019
ms.locfileid: "35772007"
---
# <a name="whats-new-in-excel-javascript-api-18"></a>Excel JavaScript API 1.8 的最近更新

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
|[ChartAxisData](/javascript/api/excel/excel.chartaxisdata)|[对齐方式](/javascript/api/excel/excel.chartaxisdata#alignment)|表示指定轴刻度线标签的对齐方式。 有关详细信息, 请参阅 ChartTextHorizontalAlignment。|
||[isBetweenCategories](/javascript/api/excel/excel.chartaxisdata#isbetweencategories)|表示数值轴是否与分类之间的分类轴交叉。|
||[符号](/javascript/api/excel/excel.chartaxisdata#multilevel)|表示是否为多级轴。|
||[numberFormat](/javascript/api/excel/excel.chartaxisdata#numberformat)|表示轴刻度线标签的格式代码。|
||[一定](/javascript/api/excel/excel.chartaxisdata#offset)|表示不同标签级别之间的距离以及一级标签和轴线之间的距离。 此值应该是 0 到 1000 之间的整数。|
||[position](/javascript/api/excel/excel.chartaxisdata#position)|表示两轴交叉的特定轴位置。 有关详细信息, 请参阅 ChartAxisPosition。|
||[positionAt](/javascript/api/excel/excel.chartaxisdata#positionat)|表示两轴交叉的特定轴位置。 应使用 SetPositionAt(double) 方法设置此属性。|
||[textOrientation](/javascript/api/excel/excel.chartaxisdata#textorientation)|表示轴刻度线标签的文本方向。 此值应是 -90 到 90 或 180（垂直文本）之间的整数。|
|[ChartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|[fill](/javascript/api/excel/excel.chartaxisformat#fill)|表示图表填充格式。 只读。|
|[ChartAxisLoadOptions](/javascript/api/excel/excel.chartaxisloadoptions)|[对齐方式](/javascript/api/excel/excel.chartaxisloadoptions#alignment)|表示指定轴刻度线标签的对齐方式。 有关详细信息, 请参阅 ChartTextHorizontalAlignment。|
||[isBetweenCategories](/javascript/api/excel/excel.chartaxisloadoptions#isbetweencategories)|表示数值轴是否与分类之间的分类轴交叉。|
||[符号](/javascript/api/excel/excel.chartaxisloadoptions#multilevel)|表示是否为多级轴。|
||[numberFormat](/javascript/api/excel/excel.chartaxisloadoptions#numberformat)|表示轴刻度线标签的格式代码。|
||[一定](/javascript/api/excel/excel.chartaxisloadoptions#offset)|表示不同标签级别之间的距离以及一级标签和轴线之间的距离。 此值应该是 0 到 1000 之间的整数。|
||[position](/javascript/api/excel/excel.chartaxisloadoptions#position)|表示两轴交叉的特定轴位置。 有关详细信息, 请参阅 ChartAxisPosition。|
||[positionAt](/javascript/api/excel/excel.chartaxisloadoptions#positionat)|表示两轴交叉的特定轴位置。 应使用 SetPositionAt(double) 方法设置此属性。|
||[textOrientation](/javascript/api/excel/excel.chartaxisloadoptions#textorientation)|表示轴刻度线标签的文本方向。 此值应是 -90 到 90 或 180（垂直文本）之间的整数。|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[setFormula (formula: string)](/javascript/api/excel/excel.chartaxistitle#setformula-formula-)|该字符串值表示采用 A1 表示法的图表轴标题的公式。|
|[ChartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|[边缘](/javascript/api/excel/excel.chartaxistitleformat#border)|表示边框格式，包括颜色、线条样式和粗细。|
||[fill](/javascript/api/excel/excel.chartaxistitleformat#fill)|表示图表填充格式。|
|[ChartAxisTitleFormatData](/javascript/api/excel/excel.chartaxistitleformatdata)|[边缘](/javascript/api/excel/excel.chartaxistitleformatdata#border)|表示边框格式，包括颜色、线条样式和粗细。|
|[ChartAxisTitleFormatLoadOptions](/javascript/api/excel/excel.chartaxistitleformatloadoptions)|[边缘](/javascript/api/excel/excel.chartaxistitleformatloadoptions#border)|表示边框格式，包括颜色、线条样式和粗细。|
|[ChartAxisTitleFormatUpdateData](/javascript/api/excel/excel.chartaxistitleformatupdatedata)|[边缘](/javascript/api/excel/excel.chartaxistitleformatupdatedata#border)|表示边框格式，包括颜色、线条样式和粗细。|
|[ChartAxisUpdateData](/javascript/api/excel/excel.chartaxisupdatedata)|[对齐方式](/javascript/api/excel/excel.chartaxisupdatedata#alignment)|表示指定轴刻度线标签的对齐方式。 有关详细信息, 请参阅 ChartTextHorizontalAlignment。|
||[isBetweenCategories](/javascript/api/excel/excel.chartaxisupdatedata#isbetweencategories)|表示数值轴是否与分类之间的分类轴交叉。|
||[符号](/javascript/api/excel/excel.chartaxisupdatedata#multilevel)|表示是否为多级轴。|
||[numberFormat](/javascript/api/excel/excel.chartaxisupdatedata#numberformat)|表示轴刻度线标签的格式代码。|
||[一定](/javascript/api/excel/excel.chartaxisupdatedata#offset)|表示不同标签级别之间的距离以及一级标签和轴线之间的距离。 此值应该是 0 到 1000 之间的整数。|
||[position](/javascript/api/excel/excel.chartaxisupdatedata#position)|表示两轴交叉的特定轴位置。 有关详细信息, 请参阅 ChartAxisPosition。|
||[textOrientation](/javascript/api/excel/excel.chartaxisupdatedata#textorientation)|表示轴刻度线标签的文本方向。 此值应是 -90 到 90 或 180（垂直文本）之间的整数。|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[clear()](/javascript/api/excel/excel.chartborder#clear--)|清除图表元素的边框格式。|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[onActivated](/javascript/api/excel/excel.chartcollection#onactivated)|在激活图表时发生。|
||[onAdded](/javascript/api/excel/excel.chartcollection#onadded)|将新图表添加到工作表时发生此事件。|
||[onDeactivated](/javascript/api/excel/excel.chartcollection#ondeactivated)|当停用图表时发生此事件。|
||[onDeleted](/javascript/api/excel/excel.chartcollection#ondeleted)|在删除图表时发生。|
|[ChartCollectionLoadOptions](/javascript/api/excel/excel.chartcollectionloadoptions)|[categoryLabelLevel](/javascript/api/excel/excel.chartcollectionloadoptions#categorylabellevel)|对于集合中的每一项: 返回或设置一个 ChartCategoryLabelLevel 枚举常量, 该常量引用|
||[displayBlanksAs](/javascript/api/excel/excel.chartcollectionloadoptions#displayblanksas)|对于集合中的每一项: 返回或设置在图表上绘制空白单元格的方式。 读/写。|
||[plotArea](/javascript/api/excel/excel.chartcollectionloadoptions#plotarea)|对于集合中的每一项: 代表图表的 plotArea。|
||[plotBy](/javascript/api/excel/excel.chartcollectionloadoptions#plotby)|对于集合中的每一项: 返回或设置在图表上将列或行用作数据系列的方式。 读/写。|
||[plotVisibleOnly](/javascript/api/excel/excel.chartcollectionloadoptions#plotvisibleonly)|对于集合中的每一项: 如果只绘制可见单元格, 则为 True。如果绘制可见单元格和隐藏单元格，则为 False。 读/写。|
||[seriesNameLevel](/javascript/api/excel/excel.chartcollectionloadoptions#seriesnamelevel)|对于集合中的每一项: 返回或设置一个 ChartSeriesNameLevel 枚举常量, 该常量引用|
||[showDataLabelsOverMaximum](/javascript/api/excel/excel.chartcollectionloadoptions#showdatalabelsovermaximum)|对于集合中的每一项: 表示在值大于数值轴上的最大值时是否显示数据标签。|
||[style](/javascript/api/excel/excel.chartcollectionloadoptions#style)|对于集合中的每一项: 返回或设置图表的图表样式。 读/写。|
|[ChartData](/javascript/api/excel/excel.chartdata)|[categoryLabelLevel](/javascript/api/excel/excel.chartdata#categorylabellevel)|返回或设置一个 ChartCategoryLabelLevel 枚举常量, 该常量引用|
||[displayBlanksAs](/javascript/api/excel/excel.chartdata#displayblanksas)|返回或设置图表上的空白单元格的绘制方式。 读/写。|
||[plotArea](/javascript/api/excel/excel.chartdata#plotarea)|表示图表的绘制区域。|
||[plotBy](/javascript/api/excel/excel.chartdata#plotby)|返回或设置图表上的列或行用作数据系列的方式。 读/写。|
||[plotVisibleOnly](/javascript/api/excel/excel.chartdata#plotvisibleonly)|如果仅绘制可见单元格，则为 True。如果绘制可见单元格和隐藏单元格，则为 False。 读/写。|
||[seriesNameLevel](/javascript/api/excel/excel.chartdata#seriesnamelevel)|返回或设置一个 ChartSeriesNameLevel 枚举常量, 该常量引用|
||[showDataLabelsOverMaximum](/javascript/api/excel/excel.chartdata#showdatalabelsovermaximum)|表示当值大于数值轴上的最大值时是否显示数据标签。|
||[style](/javascript/api/excel/excel.chartdata#style)|返回或设置图表的图表样式。 读/写。|
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
|[ChartDataLabelData](/javascript/api/excel/excel.chartdatalabeldata)|[自动图文集](/javascript/api/excel/excel.chartdatalabeldata#autotext)|该布尔值表示数据标签是否根据上下文自动生成相应的文本。|
||[format](/javascript/api/excel/excel.chartdatalabeldata#format)|表示图表数据标签的格式。|
||[formula](/javascript/api/excel/excel.chartdatalabeldata#formula)|该字符串值表示采用 A1 表示法的图表数据标签的公式。|
||[height](/javascript/api/excel/excel.chartdatalabeldata#height)|返回图表数据标签的高度，以磅为单位。 只读。 如果图表数据标签不可见，则为 Null。|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabeldata#horizontalalignment)|表示图表数据标签水平对齐。 有关详细信息, 请参阅 ChartTextHorizontalAlignment。|
||[left](/javascript/api/excel/excel.chartdatalabeldata#left)|表示图表数据标签左边缘到图表区域左边缘的距离，以磅为单位。 如果图表数据标签不可见，则为 Null。|
||[numberFormat](/javascript/api/excel/excel.chartdatalabeldata#numberformat)|该字符串值表示数据标签的格式代码。|
||[text](/javascript/api/excel/excel.chartdatalabeldata#text)|该字符串表示图表上的数据标签文本。|
||[textOrientation](/javascript/api/excel/excel.chartdatalabeldata#textorientation)|表示图表数据标签的文本方向。 此值应是 -90 到 90 或 180（垂直文本）之间的整数。|
||[top](/javascript/api/excel/excel.chartdatalabeldata#top)|表示图表数据标签上边缘到图表区域顶部的距离，以磅为单位。 如果图表数据标签不可见，则为 Null。|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabeldata#verticalalignment)|表示图表数据标签垂直对齐。 有关详细信息, 请参阅 ChartTextVerticalAlignment。|
||[width](/javascript/api/excel/excel.chartdatalabeldata#width)|返回图表数据标签的宽度，以磅为单位。 只读。 如果图表数据标签不可见，则为 Null。|
|[ChartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|[边缘](/javascript/api/excel/excel.chartdatalabelformat#border)|表示边框格式，包括颜色、线条样式和粗细。 只读。|
|[ChartDataLabelFormatData](/javascript/api/excel/excel.chartdatalabelformatdata)|[边缘](/javascript/api/excel/excel.chartdatalabelformatdata#border)|表示边框格式，包括颜色、线条样式和粗细。 只读。|
|[ChartDataLabelFormatLoadOptions](/javascript/api/excel/excel.chartdatalabelformatloadoptions)|[边缘](/javascript/api/excel/excel.chartdatalabelformatloadoptions#border)|表示边框格式，包括颜色、线条样式和粗细。|
|[ChartDataLabelFormatUpdateData](/javascript/api/excel/excel.chartdatalabelformatupdatedata)|[边缘](/javascript/api/excel/excel.chartdatalabelformatupdatedata#border)|表示边框格式，包括颜色、线条样式和粗细。|
|[ChartDataLabelLoadOptions](/javascript/api/excel/excel.chartdatalabelloadoptions)|[自动图文集](/javascript/api/excel/excel.chartdatalabelloadoptions#autotext)|该布尔值表示数据标签是否根据上下文自动生成相应的文本。|
||[format](/javascript/api/excel/excel.chartdatalabelloadoptions#format)|表示图表数据标签的格式。|
||[formula](/javascript/api/excel/excel.chartdatalabelloadoptions#formula)|该字符串值表示采用 A1 表示法的图表数据标签的公式。|
||[height](/javascript/api/excel/excel.chartdatalabelloadoptions#height)|返回图表数据标签的高度，以磅为单位。 只读。 如果图表数据标签不可见，则为 Null。|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabelloadoptions#horizontalalignment)|表示图表数据标签水平对齐。 有关详细信息, 请参阅 ChartTextHorizontalAlignment。|
||[left](/javascript/api/excel/excel.chartdatalabelloadoptions#left)|表示图表数据标签左边缘到图表区域左边缘的距离，以磅为单位。 如果图表数据标签不可见，则为 Null。|
||[numberFormat](/javascript/api/excel/excel.chartdatalabelloadoptions#numberformat)|该字符串值表示数据标签的格式代码。|
||[text](/javascript/api/excel/excel.chartdatalabelloadoptions#text)|该字符串表示图表上的数据标签文本。|
||[textOrientation](/javascript/api/excel/excel.chartdatalabelloadoptions#textorientation)|表示图表数据标签的文本方向。 此值应是 -90 到 90 或 180（垂直文本）之间的整数。|
||[top](/javascript/api/excel/excel.chartdatalabelloadoptions#top)|表示图表数据标签上边缘到图表区域顶部的距离，以磅为单位。 如果图表数据标签不可见，则为 Null。|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabelloadoptions#verticalalignment)|表示图表数据标签垂直对齐。 有关详细信息, 请参阅 ChartTextVerticalAlignment。|
||[width](/javascript/api/excel/excel.chartdatalabelloadoptions#width)|返回图表数据标签的宽度，以磅为单位。 只读。 如果图表数据标签不可见，则为 Null。|
|[ChartDataLabelUpdateData](/javascript/api/excel/excel.chartdatalabelupdatedata)|[自动图文集](/javascript/api/excel/excel.chartdatalabelupdatedata#autotext)|该布尔值表示数据标签是否根据上下文自动生成相应的文本。|
||[format](/javascript/api/excel/excel.chartdatalabelupdatedata#format)|表示图表数据标签的格式。|
||[formula](/javascript/api/excel/excel.chartdatalabelupdatedata#formula)|该字符串值表示采用 A1 表示法的图表数据标签的公式。|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabelupdatedata#horizontalalignment)|表示图表数据标签水平对齐。 有关详细信息, 请参阅 ChartTextHorizontalAlignment。|
||[left](/javascript/api/excel/excel.chartdatalabelupdatedata#left)|表示图表数据标签左边缘到图表区域左边缘的距离，以磅为单位。 如果图表数据标签不可见，则为 Null。|
||[numberFormat](/javascript/api/excel/excel.chartdatalabelupdatedata#numberformat)|该字符串值表示数据标签的格式代码。|
||[text](/javascript/api/excel/excel.chartdatalabelupdatedata#text)|该字符串表示图表上的数据标签文本。|
||[textOrientation](/javascript/api/excel/excel.chartdatalabelupdatedata#textorientation)|表示图表数据标签的文本方向。 此值应是 -90 到 90 或 180（垂直文本）之间的整数。|
||[top](/javascript/api/excel/excel.chartdatalabelupdatedata#top)|表示图表数据标签上边缘到图表区域顶部的距离，以磅为单位。 如果图表数据标签不可见，则为 Null。|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabelupdatedata#verticalalignment)|表示图表数据标签垂直对齐。 有关详细信息, 请参阅 ChartTextVerticalAlignment。|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[自动图文集](/javascript/api/excel/excel.chartdatalabels#autotext)|表示数据标签是否根据上下文自动生成相应的文本。|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabels#horizontalalignment)|表示图表数据标签水平对齐。 有关详细信息, 请参阅 ChartTextHorizontalAlignment。|
||[numberFormat](/javascript/api/excel/excel.chartdatalabels#numberformat)|表示数据标签的格式代码。|
||[textOrientation](/javascript/api/excel/excel.chartdatalabels#textorientation)|表示数据标签的文本方向。 此值应是 -90 到 90 或 180（垂直文本）之间的整数。|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabels#verticalalignment)|表示图表数据标签垂直对齐。 有关详细信息, 请参阅 ChartTextVerticalAlignment。|
|[ChartDataLabelsData](/javascript/api/excel/excel.chartdatalabelsdata)|[自动图文集](/javascript/api/excel/excel.chartdatalabelsdata#autotext)|表示数据标签是否根据上下文自动生成相应的文本。|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabelsdata#horizontalalignment)|表示图表数据标签水平对齐。 有关详细信息, 请参阅 ChartTextHorizontalAlignment。|
||[numberFormat](/javascript/api/excel/excel.chartdatalabelsdata#numberformat)|表示数据标签的格式代码。|
||[textOrientation](/javascript/api/excel/excel.chartdatalabelsdata#textorientation)|表示数据标签的文本方向。 此值应是 -90 到 90 或 180（垂直文本）之间的整数。|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabelsdata#verticalalignment)|表示图表数据标签垂直对齐。 有关详细信息, 请参阅 ChartTextVerticalAlignment。|
|[ChartDataLabelsLoadOptions](/javascript/api/excel/excel.chartdatalabelsloadoptions)|[自动图文集](/javascript/api/excel/excel.chartdatalabelsloadoptions#autotext)|表示数据标签是否根据上下文自动生成相应的文本。|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabelsloadoptions#horizontalalignment)|表示图表数据标签水平对齐。 有关详细信息, 请参阅 ChartTextHorizontalAlignment。|
||[numberFormat](/javascript/api/excel/excel.chartdatalabelsloadoptions#numberformat)|表示数据标签的格式代码。|
||[textOrientation](/javascript/api/excel/excel.chartdatalabelsloadoptions#textorientation)|表示数据标签的文本方向。 此值应是 -90 到 90 或 180（垂直文本）之间的整数。|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabelsloadoptions#verticalalignment)|表示图表数据标签垂直对齐。 有关详细信息, 请参阅 ChartTextVerticalAlignment。|
|[ChartDataLabelsUpdateData](/javascript/api/excel/excel.chartdatalabelsupdatedata)|[自动图文集](/javascript/api/excel/excel.chartdatalabelsupdatedata#autotext)|表示数据标签是否根据上下文自动生成相应的文本。|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabelsupdatedata#horizontalalignment)|表示图表数据标签水平对齐。 有关详细信息, 请参阅 ChartTextHorizontalAlignment。|
||[numberFormat](/javascript/api/excel/excel.chartdatalabelsupdatedata#numberformat)|表示数据标签的格式代码。|
||[textOrientation](/javascript/api/excel/excel.chartdatalabelsupdatedata#textorientation)|表示数据标签的文本方向。 此值应是 -90 到 90 或 180（垂直文本）之间的整数。|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabelsupdatedata#verticalalignment)|表示图表数据标签垂直对齐。 有关详细信息, 请参阅 ChartTextVerticalAlignment。|
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
|[ChartLegendEntryCollectionLoadOptions](/javascript/api/excel/excel.chartlegendentrycollectionloadoptions)|[height](/javascript/api/excel/excel.chartlegendentrycollectionloadoptions#height)|对于集合中的每一项: 表示图表图例上的 legendEntry 的高度。|
||[index](/javascript/api/excel/excel.chartlegendentrycollectionloadoptions#index)|对于集合中的每一项: 代表图表图例中的 legendEntry 的索引。|
||[left](/javascript/api/excel/excel.chartlegendentrycollectionloadoptions#left)|对于集合中的每一项: 表示图表 legendEntry 的左侧。|
||[top](/javascript/api/excel/excel.chartlegendentrycollectionloadoptions#top)|对于集合中的每一项: 表示图表 legendEntry 的顶部。|
||[width](/javascript/api/excel/excel.chartlegendentrycollectionloadoptions#width)|对于集合中的每一项: 表示图表图例上的 legendEntry 的宽度。|
|[ChartLegendEntryData](/javascript/api/excel/excel.chartlegendentrydata)|[height](/javascript/api/excel/excel.chartlegendentrydata#height)|表示图表图例上的 legendEntry 的高度。|
||[index](/javascript/api/excel/excel.chartlegendentrydata#index)|表示图表图例中的 legendEntry 的索引。|
||[left](/javascript/api/excel/excel.chartlegendentrydata#left)|表示图表 legendEntry 的左侧。|
||[top](/javascript/api/excel/excel.chartlegendentrydata#top)|表示图表 legendEntry 的顶部。|
||[width](/javascript/api/excel/excel.chartlegendentrydata#width)|表示图表图例上的 legendEntry 的宽度。|
|[ChartLegendEntryLoadOptions](/javascript/api/excel/excel.chartlegendentryloadoptions)|[height](/javascript/api/excel/excel.chartlegendentryloadoptions#height)|表示图表图例上的 legendEntry 的高度。|
||[index](/javascript/api/excel/excel.chartlegendentryloadoptions#index)|表示图表图例中的 legendEntry 的索引。|
||[left](/javascript/api/excel/excel.chartlegendentryloadoptions#left)|表示图表 legendEntry 的左侧。|
||[top](/javascript/api/excel/excel.chartlegendentryloadoptions#top)|表示图表 legendEntry 的顶部。|
||[width](/javascript/api/excel/excel.chartlegendentryloadoptions#width)|表示图表图例上的 legendEntry 的宽度。|
|[ChartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|[边缘](/javascript/api/excel/excel.chartlegendformat#border)|表示边框格式，包括颜色、线条样式和粗细。 只读。|
|[ChartLegendFormatData](/javascript/api/excel/excel.chartlegendformatdata)|[边缘](/javascript/api/excel/excel.chartlegendformatdata#border)|表示边框格式，包括颜色、线条样式和粗细。 只读。|
|[ChartLegendFormatLoadOptions](/javascript/api/excel/excel.chartlegendformatloadoptions)|[边缘](/javascript/api/excel/excel.chartlegendformatloadoptions#border)|表示边框格式，包括颜色、线条样式和粗细。|
|[ChartLegendFormatUpdateData](/javascript/api/excel/excel.chartlegendformatupdatedata)|[边缘](/javascript/api/excel/excel.chartlegendformatupdatedata#border)|表示边框格式，包括颜色、线条样式和粗细。|
|[ChartLoadOptions](/javascript/api/excel/excel.chartloadoptions)|[categoryLabelLevel](/javascript/api/excel/excel.chartloadoptions#categorylabellevel)|返回或设置一个 ChartCategoryLabelLevel 枚举常量, 该常量引用|
||[displayBlanksAs](/javascript/api/excel/excel.chartloadoptions#displayblanksas)|返回或设置图表上的空白单元格的绘制方式。 读/写。|
||[plotArea](/javascript/api/excel/excel.chartloadoptions#plotarea)|表示图表的绘制区域。|
||[plotBy](/javascript/api/excel/excel.chartloadoptions#plotby)|返回或设置图表上的列或行用作数据系列的方式。 读/写。|
||[plotVisibleOnly](/javascript/api/excel/excel.chartloadoptions#plotvisibleonly)|如果仅绘制可见单元格，则为 True。如果绘制可见单元格和隐藏单元格，则为 False。 读/写。|
||[seriesNameLevel](/javascript/api/excel/excel.chartloadoptions#seriesnamelevel)|返回或设置一个 ChartSeriesNameLevel 枚举常量, 该常量引用|
||[showDataLabelsOverMaximum](/javascript/api/excel/excel.chartloadoptions#showdatalabelsovermaximum)|表示当值大于数值轴上的最大值时是否显示数据标签。|
||[style](/javascript/api/excel/excel.chartloadoptions#style)|返回或设置图表的图表样式。 读/写。|
|[ChartPlotArea](/javascript/api/excel/excel.chartplotarea)|[height](/javascript/api/excel/excel.chartplotarea#height)|表示 plotArea 的高度值。|
||[insideHeight](/javascript/api/excel/excel.chartplotarea#insideheight)|表示 plotArea 的 insideHeight 值。|
||[insideLeft](/javascript/api/excel/excel.chartplotarea#insideleft)|表示 plotArea 的 insideLeft 值。|
||[insideTop](/javascript/api/excel/excel.chartplotarea#insidetop)|表示 plotArea 的 insideTop 值。|
||[insideWidth](/javascript/api/excel/excel.chartplotarea#insidewidth)|表示 plotArea 的 insideWidth 值。|
||[left](/javascript/api/excel/excel.chartplotarea#left)|表示 plotArea 的 left 值。|
||[position](/javascript/api/excel/excel.chartplotarea#position)|表示 plotArea 的位置。|
||[format](/javascript/api/excel/excel.chartplotarea#format)|表示图表 plotArea 的格式。|
||[set (properties: ChartPlotArea)](/javascript/api/excel/excel.chartplotarea#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: ChartPlotAreaUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.chartplotarea#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
||[top](/javascript/api/excel/excel.chartplotarea#top)|表示 plotArea 的 top 值。|
||[width](/javascript/api/excel/excel.chartplotarea#width)|表示 plotArea 的宽度值。|
|[ChartPlotAreaData](/javascript/api/excel/excel.chartplotareadata)|[format](/javascript/api/excel/excel.chartplotareadata#format)|表示图表 plotArea 的格式。|
||[height](/javascript/api/excel/excel.chartplotareadata#height)|表示 plotArea 的高度值。|
||[insideHeight](/javascript/api/excel/excel.chartplotareadata#insideheight)|表示 plotArea 的 insideHeight 值。|
||[insideLeft](/javascript/api/excel/excel.chartplotareadata#insideleft)|表示 plotArea 的 insideLeft 值。|
||[insideTop](/javascript/api/excel/excel.chartplotareadata#insidetop)|表示 plotArea 的 insideTop 值。|
||[insideWidth](/javascript/api/excel/excel.chartplotareadata#insidewidth)|表示 plotArea 的 insideWidth 值。|
||[left](/javascript/api/excel/excel.chartplotareadata#left)|表示 plotArea 的 left 值。|
||[position](/javascript/api/excel/excel.chartplotareadata#position)|表示 plotArea 的位置。|
||[top](/javascript/api/excel/excel.chartplotareadata#top)|表示 plotArea 的 top 值。|
||[width](/javascript/api/excel/excel.chartplotareadata#width)|表示 plotArea 的宽度值。|
|[ChartPlotAreaFormat](/javascript/api/excel/excel.chartplotareaformat)|[边缘](/javascript/api/excel/excel.chartplotareaformat#border)|表示图表 plotArea 的边框属性。|
||[fill](/javascript/api/excel/excel.chartplotareaformat#fill)|表示对象的填充格式，包括背景格式信息。|
||[set (properties: ChartPlotAreaFormat)](/javascript/api/excel/excel.chartplotareaformat#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: ChartPlotAreaFormatUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.chartplotareaformat#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
|[ChartPlotAreaFormatData](/javascript/api/excel/excel.chartplotareaformatdata)|[边缘](/javascript/api/excel/excel.chartplotareaformatdata#border)|表示图表 plotArea 的边框属性。|
|[ChartPlotAreaFormatLoadOptions](/javascript/api/excel/excel.chartplotareaformatloadoptions)|[$all](/javascript/api/excel/excel.chartplotareaformatloadoptions#$all)||
||[边缘](/javascript/api/excel/excel.chartplotareaformatloadoptions#border)|表示图表 plotArea 的边框属性。|
|[ChartPlotAreaFormatUpdateData](/javascript/api/excel/excel.chartplotareaformatupdatedata)|[边缘](/javascript/api/excel/excel.chartplotareaformatupdatedata#border)|表示图表 plotArea 的边框属性。|
|[ChartPlotAreaLoadOptions](/javascript/api/excel/excel.chartplotarealoadoptions)|[$all](/javascript/api/excel/excel.chartplotarealoadoptions#$all)||
||[format](/javascript/api/excel/excel.chartplotarealoadoptions#format)|表示图表 plotArea 的格式。|
||[height](/javascript/api/excel/excel.chartplotarealoadoptions#height)|表示 plotArea 的高度值。|
||[insideHeight](/javascript/api/excel/excel.chartplotarealoadoptions#insideheight)|表示 plotArea 的 insideHeight 值。|
||[insideLeft](/javascript/api/excel/excel.chartplotarealoadoptions#insideleft)|表示 plotArea 的 insideLeft 值。|
||[insideTop](/javascript/api/excel/excel.chartplotarealoadoptions#insidetop)|表示 plotArea 的 insideTop 值。|
||[insideWidth](/javascript/api/excel/excel.chartplotarealoadoptions#insidewidth)|表示 plotArea 的 insideWidth 值。|
||[left](/javascript/api/excel/excel.chartplotarealoadoptions#left)|表示 plotArea 的 left 值。|
||[position](/javascript/api/excel/excel.chartplotarealoadoptions#position)|表示 plotArea 的位置。|
||[top](/javascript/api/excel/excel.chartplotarealoadoptions#top)|表示 plotArea 的 top 值。|
||[width](/javascript/api/excel/excel.chartplotarealoadoptions#width)|表示 plotArea 的宽度值。|
|[ChartPlotAreaUpdateData](/javascript/api/excel/excel.chartplotareaupdatedata)|[format](/javascript/api/excel/excel.chartplotareaupdatedata#format)|表示图表 plotArea 的格式。|
||[height](/javascript/api/excel/excel.chartplotareaupdatedata#height)|表示 plotArea 的高度值。|
||[insideHeight](/javascript/api/excel/excel.chartplotareaupdatedata#insideheight)|表示 plotArea 的 insideHeight 值。|
||[insideLeft](/javascript/api/excel/excel.chartplotareaupdatedata#insideleft)|表示 plotArea 的 insideLeft 值。|
||[insideTop](/javascript/api/excel/excel.chartplotareaupdatedata#insidetop)|表示 plotArea 的 insideTop 值。|
||[insideWidth](/javascript/api/excel/excel.chartplotareaupdatedata#insidewidth)|表示 plotArea 的 insideWidth 值。|
||[left](/javascript/api/excel/excel.chartplotareaupdatedata#left)|表示 plotArea 的 left 值。|
||[position](/javascript/api/excel/excel.chartplotareaupdatedata#position)|表示 plotArea 的位置。|
||[top](/javascript/api/excel/excel.chartplotareaupdatedata#top)|表示 plotArea 的 top 值。|
||[width](/javascript/api/excel/excel.chartplotareaupdatedata#width)|表示 plotArea 的宽度值。|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[axisGroup](/javascript/api/excel/excel.chartseries#axisgroup)|返回或设置指定系列的组。 读/写|
||[分离](/javascript/api/excel/excel.chartseries#explosion)|返回或设置饼图或圆环图切片的分解程度值。 如果未分解（切片尖端位于饼图中心），则返回 0（零）。 读/写。|
||[firstSliceAngle](/javascript/api/excel/excel.chartseries#firstsliceangle)|返回或设置第一个饼图或圆环图切片的角度，以度为单位（从垂直方向起为顺时针）。 只适用于饼图、三维饼图和圆环图。 可以是 0 到 360 之间的值。 读/写|
||[invertIfNegative](/javascript/api/excel/excel.chartseries#invertifnegative)|如果 Excel 在值为负数时反转项目中的图案，则为 True。 读/写。|
||[比例](/javascript/api/excel/excel.chartseries#overlap)|指定条柱的摆放方式。 可以是–100到100之间的值。 只适用于二维条形图和二维柱形图。 读/写。|
||[dataLabels](/javascript/api/excel/excel.chartseries#datalabels)|表示系列中所有数据标签的集合。|
||[secondPlotSize](/javascript/api/excel/excel.chartseries#secondplotsize)|返回或设置复合饼图或复合条饼图的辅助分区的大小，以占主饼图大小的百分比表示。 可以是 5 到 200 之间的值。 读/写。|
||[splitType](/javascript/api/excel/excel.chartseries#splittype)|返回或设置复合饼图或复合条饼图中两个分区的拆分方式。 读/写。|
||[varyByCategories](/javascript/api/excel/excel.chartseries#varybycategories)|如果 Excel 为每个数据标记分配不同的颜色或图案，则为 True。 图表必须只包含一个系列。 读/写。|
|[ChartSeriesCollectionLoadOptions](/javascript/api/excel/excel.chartseriescollectionloadoptions)|[axisGroup](/javascript/api/excel/excel.chartseriescollectionloadoptions#axisgroup)|对于集合中的每一项: 返回或设置指定系列的组。 读/写|
||[dataLabels](/javascript/api/excel/excel.chartseriescollectionloadoptions#datalabels)|对于集合中的每一项: 代表系列中所有 dataLabels 的集合。|
||[分离](/javascript/api/excel/excel.chartseriescollectionloadoptions#explosion)|对于集合中的每一项: 返回或设置饼图或圆环图切片的分解值。 如果未分解（切片尖端位于饼图中心），则返回 0（零）。 读/写。|
||[firstSliceAngle](/javascript/api/excel/excel.chartseriescollectionloadoptions#firstsliceangle)|对于集合中的每一项: 返回或设置第一个饼图或圆环图的扇区的角度 (以度为单位) (从垂直方向顺时针方向)。 只适用于饼图、三维饼图和圆环图。 可以是 0 到 360 之间的值。 读/写|
||[invertIfNegative](/javascript/api/excel/excel.chartseriescollectionloadoptions#invertifnegative)|对于集合中的每一项: True 如果 Microsoft Excel 在对应于负数时反转项中的模式。 读/写。|
||[比例](/javascript/api/excel/excel.chartseriescollectionloadoptions#overlap)|对于集合中的每一项: 指定栏和列的放置方式。 可以是–100到100之间的值。 只适用于二维条形图和二维柱形图。 读/写。|
||[secondPlotSize](/javascript/api/excel/excel.chartseriescollectionloadoptions#secondplotsize)|对于集合中的每一项: 返回或设置复合饼图或复合条饼图中的第二部分的大小, 以主饼图大小的百分比表示。 可以是 5 到 200 之间的值。 读/写。|
||[splitType](/javascript/api/excel/excel.chartseriescollectionloadoptions#splittype)|对于集合中的每一项: 返回或设置一个复合饼图或复合条饼图中的两个部分的分割方式。 读/写。|
||[varyByCategories](/javascript/api/excel/excel.chartseriescollectionloadoptions#varybycategories)|对于集合中的每个项目: 如果 Microsoft Excel 为每个数据标记分配不同的颜色或图案, 则为 True。 图表必须只包含一个系列。 读/写。|
|[ChartSeriesData](/javascript/api/excel/excel.chartseriesdata)|[axisGroup](/javascript/api/excel/excel.chartseriesdata#axisgroup)|返回或设置指定系列的组。 读/写|
||[dataLabels](/javascript/api/excel/excel.chartseriesdata#datalabels)|表示系列中所有数据标签的集合。|
||[分离](/javascript/api/excel/excel.chartseriesdata#explosion)|返回或设置饼图或圆环图切片的分解程度值。 如果未分解（切片尖端位于饼图中心），则返回 0（零）。 读/写。|
||[firstSliceAngle](/javascript/api/excel/excel.chartseriesdata#firstsliceangle)|返回或设置第一个饼图或圆环图切片的角度，以度为单位（从垂直方向起为顺时针）。 只适用于饼图、三维饼图和圆环图。 可以是 0 到 360 之间的值。 读/写|
||[invertIfNegative](/javascript/api/excel/excel.chartseriesdata#invertifnegative)|如果 Excel 在值为负数时反转项目中的图案，则为 True。 读/写。|
||[比例](/javascript/api/excel/excel.chartseriesdata#overlap)|指定条柱的摆放方式。 可以是–100到100之间的值。 只适用于二维条形图和二维柱形图。 读/写。|
||[secondPlotSize](/javascript/api/excel/excel.chartseriesdata#secondplotsize)|返回或设置复合饼图或复合条饼图的辅助分区的大小，以占主饼图大小的百分比表示。 可以是 5 到 200 之间的值。 读/写。|
||[splitType](/javascript/api/excel/excel.chartseriesdata#splittype)|返回或设置复合饼图或复合条饼图中两个分区的拆分方式。 读/写。|
||[varyByCategories](/javascript/api/excel/excel.chartseriesdata#varybycategories)|如果 Excel 为每个数据标记分配不同的颜色或图案，则为 True。 图表必须只包含一个系列。 读/写。|
|[ChartSeriesLoadOptions](/javascript/api/excel/excel.chartseriesloadoptions)|[axisGroup](/javascript/api/excel/excel.chartseriesloadoptions#axisgroup)|返回或设置指定系列的组。 读/写|
||[dataLabels](/javascript/api/excel/excel.chartseriesloadoptions#datalabels)|表示系列中所有数据标签的集合。|
||[分离](/javascript/api/excel/excel.chartseriesloadoptions#explosion)|返回或设置饼图或圆环图切片的分解程度值。 如果未分解（切片尖端位于饼图中心），则返回 0（零）。 读/写。|
||[firstSliceAngle](/javascript/api/excel/excel.chartseriesloadoptions#firstsliceangle)|返回或设置第一个饼图或圆环图切片的角度，以度为单位（从垂直方向起为顺时针）。 只适用于饼图、三维饼图和圆环图。 可以是 0 到 360 之间的值。 读/写|
||[invertIfNegative](/javascript/api/excel/excel.chartseriesloadoptions#invertifnegative)|如果 Excel 在值为负数时反转项目中的图案，则为 True。 读/写。|
||[比例](/javascript/api/excel/excel.chartseriesloadoptions#overlap)|指定条柱的摆放方式。 可以是–100到100之间的值。 只适用于二维条形图和二维柱形图。 读/写。|
||[secondPlotSize](/javascript/api/excel/excel.chartseriesloadoptions#secondplotsize)|返回或设置复合饼图或复合条饼图的辅助分区的大小，以占主饼图大小的百分比表示。 可以是 5 到 200 之间的值。 读/写。|
||[splitType](/javascript/api/excel/excel.chartseriesloadoptions#splittype)|返回或设置复合饼图或复合条饼图中两个分区的拆分方式。 读/写。|
||[varyByCategories](/javascript/api/excel/excel.chartseriesloadoptions#varybycategories)|如果 Excel 为每个数据标记分配不同的颜色或图案，则为 True。 图表必须只包含一个系列。 读/写。|
|[ChartSeriesUpdateData](/javascript/api/excel/excel.chartseriesupdatedata)|[axisGroup](/javascript/api/excel/excel.chartseriesupdatedata#axisgroup)|返回或设置指定系列的组。 读/写|
||[dataLabels](/javascript/api/excel/excel.chartseriesupdatedata#datalabels)|表示系列中所有数据标签的集合。|
||[分离](/javascript/api/excel/excel.chartseriesupdatedata#explosion)|返回或设置饼图或圆环图切片的分解程度值。 如果未分解（切片尖端位于饼图中心），则返回 0（零）。 读/写。|
||[firstSliceAngle](/javascript/api/excel/excel.chartseriesupdatedata#firstsliceangle)|返回或设置第一个饼图或圆环图切片的角度，以度为单位（从垂直方向起为顺时针）。 只适用于饼图、三维饼图和圆环图。 可以是 0 到 360 之间的值。 读/写|
||[invertIfNegative](/javascript/api/excel/excel.chartseriesupdatedata#invertifnegative)|如果 Excel 在值为负数时反转项目中的图案，则为 True。 读/写。|
||[比例](/javascript/api/excel/excel.chartseriesupdatedata#overlap)|指定条柱的摆放方式。 可以是–100到100之间的值。 只适用于二维条形图和二维柱形图。 读/写。|
||[secondPlotSize](/javascript/api/excel/excel.chartseriesupdatedata#secondplotsize)|返回或设置复合饼图或复合条饼图的辅助分区的大小，以占主饼图大小的百分比表示。 可以是 5 到 200 之间的值。 读/写。|
||[splitType](/javascript/api/excel/excel.chartseriesupdatedata#splittype)|返回或设置复合饼图或复合条饼图中两个分区的拆分方式。 读/写。|
||[varyByCategories](/javascript/api/excel/excel.chartseriesupdatedata#varybycategories)|如果 Excel 为每个数据标记分配不同的颜色或图案，则为 True。 图表必须只包含一个系列。 读/写。|
|[ChartTrendline](/javascript/api/excel/excel.charttrendline)|[backwardPeriod](/javascript/api/excel/excel.charttrendline#backwardperiod)|表示趋势线向后延伸的周期数。|
||[forwardPeriod](/javascript/api/excel/excel.charttrendline#forwardperiod)|表示趋势线向前延伸的周期数。|
||[标志](/javascript/api/excel/excel.charttrendline#label)|表示图表趋势线的标签。|
||[showEquation](/javascript/api/excel/excel.charttrendline#showequation)|如果图表上显示趋势线公式，则为 True。|
||[showRSquared](/javascript/api/excel/excel.charttrendline#showrsquared)|如果图表上显示趋势线的 R 平方值，则为 True。|
|[ChartTrendlineCollectionLoadOptions](/javascript/api/excel/excel.charttrendlinecollectionloadoptions)|[backwardPeriod](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#backwardperiod)|对于集合中的每一项: 表示趋势线向后延伸的周期数。|
||[forwardPeriod](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#forwardperiod)|对于集合中的每一项: 表示趋势线向前延伸的周期数。|
||[标志](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#label)|对于集合中的每一项: 代表图表趋势线的标签。|
||[showEquation](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#showequation)|对于集合中的每个项目: 如果图表上显示趋势线的公式, 则为 True。|
||[showRSquared](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#showrsquared)|对于集合中的每一项: 如果趋势线的 R 平方值显示在图表上, 则为 True。|
|[ChartTrendlineData](/javascript/api/excel/excel.charttrendlinedata)|[backwardPeriod](/javascript/api/excel/excel.charttrendlinedata#backwardperiod)|表示趋势线向后延伸的周期数。|
||[forwardPeriod](/javascript/api/excel/excel.charttrendlinedata#forwardperiod)|表示趋势线向前延伸的周期数。|
||[标志](/javascript/api/excel/excel.charttrendlinedata#label)|表示图表趋势线的标签。|
||[showEquation](/javascript/api/excel/excel.charttrendlinedata#showequation)|如果图表上显示趋势线公式，则为 True。|
||[showRSquared](/javascript/api/excel/excel.charttrendlinedata#showrsquared)|如果图表上显示趋势线的 R 平方值，则为 True。|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[自动图文集](/javascript/api/excel/excel.charttrendlinelabel#autotext)|该布尔值表示趋势线标签是否根据上下文自动生成相应的文本。|
||[formula](/javascript/api/excel/excel.charttrendlinelabel#formula)|该字符串值表示采用 A1 表示法的图表趋势线标签的公式。|
||[horizontalAlignment](/javascript/api/excel/excel.charttrendlinelabel#horizontalalignment)|表示图表趋势线标签水平对齐。 有关详细信息, 请参阅 ChartTextHorizontalAlignment。|
||[left](/javascript/api/excel/excel.charttrendlinelabel#left)|表示图表趋势线标签左边缘到图表区域左边缘的距离，以磅为单位。 如果图表趋势线标签不可见，则为 Null。|
||[numberFormat](/javascript/api/excel/excel.charttrendlinelabel#numberformat)|该字符串值表示趋势线标签的格式代码。|
||[format](/javascript/api/excel/excel.charttrendlinelabel#format)|表示图表趋势线标签的格式。|
||[height](/javascript/api/excel/excel.charttrendlinelabel#height)|返回图表趋势线标签的高度，以磅为单位。 只读。 如果图表趋势线标签不可见，则为 Null。|
||[width](/javascript/api/excel/excel.charttrendlinelabel#width)|返回图表趋势线标签的宽度，以磅为单位。 只读。 如果图表趋势线标签不可见，则为 Null。|
||[set (properties: ChartTrendlineLabel)](/javascript/api/excel/excel.charttrendlinelabel#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: ChartTrendlineLabelUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.charttrendlinelabel#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
||[text](/javascript/api/excel/excel.charttrendlinelabel#text)|该字符串表示图表上的趋势线标签文本。|
||[textOrientation](/javascript/api/excel/excel.charttrendlinelabel#textorientation)|表示图表趋势线标签的文本方向。 此值应是 -90 到 90 或 180（垂直文本）之间的整数。|
||[top](/javascript/api/excel/excel.charttrendlinelabel#top)|表示图表趋势线标签上边缘到图表区域顶部的距离，以磅为单位。 如果图表趋势线标签不可见，则为 Null。|
||[verticalAlignment](/javascript/api/excel/excel.charttrendlinelabel#verticalalignment)|表示图表趋势线标签垂直对齐。 有关详细信息, 请参阅 ChartTextVerticalAlignment。|
|[ChartTrendlineLabelData](/javascript/api/excel/excel.charttrendlinelabeldata)|[自动图文集](/javascript/api/excel/excel.charttrendlinelabeldata#autotext)|该布尔值表示趋势线标签是否根据上下文自动生成相应的文本。|
||[format](/javascript/api/excel/excel.charttrendlinelabeldata#format)|表示图表趋势线标签的格式。|
||[formula](/javascript/api/excel/excel.charttrendlinelabeldata#formula)|该字符串值表示采用 A1 表示法的图表趋势线标签的公式。|
||[height](/javascript/api/excel/excel.charttrendlinelabeldata#height)|返回图表趋势线标签的高度，以磅为单位。 只读。 如果图表趋势线标签不可见，则为 Null。|
||[horizontalAlignment](/javascript/api/excel/excel.charttrendlinelabeldata#horizontalalignment)|表示图表趋势线标签水平对齐。 有关详细信息, 请参阅 ChartTextHorizontalAlignment。|
||[left](/javascript/api/excel/excel.charttrendlinelabeldata#left)|表示图表趋势线标签左边缘到图表区域左边缘的距离，以磅为单位。 如果图表趋势线标签不可见，则为 Null。|
||[numberFormat](/javascript/api/excel/excel.charttrendlinelabeldata#numberformat)|该字符串值表示趋势线标签的格式代码。|
||[text](/javascript/api/excel/excel.charttrendlinelabeldata#text)|该字符串表示图表上的趋势线标签文本。|
||[textOrientation](/javascript/api/excel/excel.charttrendlinelabeldata#textorientation)|表示图表趋势线标签的文本方向。 此值应是 -90 到 90 或 180（垂直文本）之间的整数。|
||[top](/javascript/api/excel/excel.charttrendlinelabeldata#top)|表示图表趋势线标签上边缘到图表区域顶部的距离，以磅为单位。 如果图表趋势线标签不可见，则为 Null。|
||[verticalAlignment](/javascript/api/excel/excel.charttrendlinelabeldata#verticalalignment)|表示图表趋势线标签垂直对齐。 有关详细信息, 请参阅 ChartTextVerticalAlignment。|
||[width](/javascript/api/excel/excel.charttrendlinelabeldata#width)|返回图表趋势线标签的宽度，以磅为单位。 只读。 如果图表趋势线标签不可见，则为 Null。|
|[ChartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|[边缘](/javascript/api/excel/excel.charttrendlinelabelformat#border)|表示边框格式，包括颜色、线条样式和粗细。|
||[fill](/javascript/api/excel/excel.charttrendlinelabelformat#fill)|表示当前图表趋势线标签的填充格式。|
||[font](/javascript/api/excel/excel.charttrendlinelabelformat#font)|表示图表趋势线标签的字体属性（字体名称、字体大小、颜色等）。|
||[set (properties: ChartTrendlineLabelFormat)](/javascript/api/excel/excel.charttrendlinelabelformat#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: ChartTrendlineLabelFormatUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.charttrendlinelabelformat#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
|[ChartTrendlineLabelFormatData](/javascript/api/excel/excel.charttrendlinelabelformatdata)|[边缘](/javascript/api/excel/excel.charttrendlinelabelformatdata#border)|表示边框格式，包括颜色、线条样式和粗细。|
||[font](/javascript/api/excel/excel.charttrendlinelabelformatdata#font)|表示图表趋势线标签的字体属性（字体名称、字体大小、颜色等）。|
|[ChartTrendlineLabelFormatLoadOptions](/javascript/api/excel/excel.charttrendlinelabelformatloadoptions)|[$all](/javascript/api/excel/excel.charttrendlinelabelformatloadoptions#$all)||
||[边缘](/javascript/api/excel/excel.charttrendlinelabelformatloadoptions#border)|表示边框格式，包括颜色、线条样式和粗细。|
||[font](/javascript/api/excel/excel.charttrendlinelabelformatloadoptions#font)|表示图表趋势线标签的字体属性（字体名称、字体大小、颜色等）。|
|[ChartTrendlineLabelFormatUpdateData](/javascript/api/excel/excel.charttrendlinelabelformatupdatedata)|[边缘](/javascript/api/excel/excel.charttrendlinelabelformatupdatedata#border)|表示边框格式，包括颜色、线条样式和粗细。|
||[font](/javascript/api/excel/excel.charttrendlinelabelformatupdatedata#font)|表示图表趋势线标签的字体属性（字体名称、字体大小、颜色等）。|
|[ChartTrendlineLabelLoadOptions](/javascript/api/excel/excel.charttrendlinelabelloadoptions)|[$all](/javascript/api/excel/excel.charttrendlinelabelloadoptions#$all)||
||[自动图文集](/javascript/api/excel/excel.charttrendlinelabelloadoptions#autotext)|该布尔值表示趋势线标签是否根据上下文自动生成相应的文本。|
||[format](/javascript/api/excel/excel.charttrendlinelabelloadoptions#format)|表示图表趋势线标签的格式。|
||[formula](/javascript/api/excel/excel.charttrendlinelabelloadoptions#formula)|该字符串值表示采用 A1 表示法的图表趋势线标签的公式。|
||[height](/javascript/api/excel/excel.charttrendlinelabelloadoptions#height)|返回图表趋势线标签的高度，以磅为单位。 只读。 如果图表趋势线标签不可见，则为 Null。|
||[horizontalAlignment](/javascript/api/excel/excel.charttrendlinelabelloadoptions#horizontalalignment)|表示图表趋势线标签水平对齐。 有关详细信息, 请参阅 ChartTextHorizontalAlignment。|
||[left](/javascript/api/excel/excel.charttrendlinelabelloadoptions#left)|表示图表趋势线标签左边缘到图表区域左边缘的距离，以磅为单位。 如果图表趋势线标签不可见，则为 Null。|
||[numberFormat](/javascript/api/excel/excel.charttrendlinelabelloadoptions#numberformat)|该字符串值表示趋势线标签的格式代码。|
||[text](/javascript/api/excel/excel.charttrendlinelabelloadoptions#text)|该字符串表示图表上的趋势线标签文本。|
||[textOrientation](/javascript/api/excel/excel.charttrendlinelabelloadoptions#textorientation)|表示图表趋势线标签的文本方向。 此值应是 -90 到 90 或 180（垂直文本）之间的整数。|
||[top](/javascript/api/excel/excel.charttrendlinelabelloadoptions#top)|表示图表趋势线标签上边缘到图表区域顶部的距离，以磅为单位。 如果图表趋势线标签不可见，则为 Null。|
||[verticalAlignment](/javascript/api/excel/excel.charttrendlinelabelloadoptions#verticalalignment)|表示图表趋势线标签垂直对齐。 有关详细信息, 请参阅 ChartTextVerticalAlignment。|
||[width](/javascript/api/excel/excel.charttrendlinelabelloadoptions#width)|返回图表趋势线标签的宽度，以磅为单位。 只读。 如果图表趋势线标签不可见，则为 Null。|
|[ChartTrendlineLabelUpdateData](/javascript/api/excel/excel.charttrendlinelabelupdatedata)|[自动图文集](/javascript/api/excel/excel.charttrendlinelabelupdatedata#autotext)|该布尔值表示趋势线标签是否根据上下文自动生成相应的文本。|
||[format](/javascript/api/excel/excel.charttrendlinelabelupdatedata#format)|表示图表趋势线标签的格式。|
||[formula](/javascript/api/excel/excel.charttrendlinelabelupdatedata#formula)|该字符串值表示采用 A1 表示法的图表趋势线标签的公式。|
||[horizontalAlignment](/javascript/api/excel/excel.charttrendlinelabelupdatedata#horizontalalignment)|表示图表趋势线标签水平对齐。 有关详细信息, 请参阅 ChartTextHorizontalAlignment。|
||[left](/javascript/api/excel/excel.charttrendlinelabelupdatedata#left)|表示图表趋势线标签左边缘到图表区域左边缘的距离，以磅为单位。 如果图表趋势线标签不可见，则为 Null。|
||[numberFormat](/javascript/api/excel/excel.charttrendlinelabelupdatedata#numberformat)|该字符串值表示趋势线标签的格式代码。|
||[text](/javascript/api/excel/excel.charttrendlinelabelupdatedata#text)|该字符串表示图表上的趋势线标签文本。|
||[textOrientation](/javascript/api/excel/excel.charttrendlinelabelupdatedata#textorientation)|表示图表趋势线标签的文本方向。 此值应是 -90 到 90 或 180（垂直文本）之间的整数。|
||[top](/javascript/api/excel/excel.charttrendlinelabelupdatedata#top)|表示图表趋势线标签上边缘到图表区域顶部的距离，以磅为单位。 如果图表趋势线标签不可见，则为 Null。|
||[verticalAlignment](/javascript/api/excel/excel.charttrendlinelabelupdatedata#verticalalignment)|表示图表趋势线标签垂直对齐。 有关详细信息, 请参阅 ChartTextVerticalAlignment。|
|[ChartTrendlineLoadOptions](/javascript/api/excel/excel.charttrendlineloadoptions)|[backwardPeriod](/javascript/api/excel/excel.charttrendlineloadoptions#backwardperiod)|表示趋势线向后延伸的周期数。|
||[forwardPeriod](/javascript/api/excel/excel.charttrendlineloadoptions#forwardperiod)|表示趋势线向前延伸的周期数。|
||[标志](/javascript/api/excel/excel.charttrendlineloadoptions#label)|表示图表趋势线的标签。|
||[showEquation](/javascript/api/excel/excel.charttrendlineloadoptions#showequation)|如果图表上显示趋势线公式，则为 True。|
||[showRSquared](/javascript/api/excel/excel.charttrendlineloadoptions#showrsquared)|如果图表上显示趋势线的 R 平方值，则为 True。|
|[ChartTrendlineUpdateData](/javascript/api/excel/excel.charttrendlineupdatedata)|[backwardPeriod](/javascript/api/excel/excel.charttrendlineupdatedata#backwardperiod)|表示趋势线向后延伸的周期数。|
||[forwardPeriod](/javascript/api/excel/excel.charttrendlineupdatedata#forwardperiod)|表示趋势线向前延伸的周期数。|
||[标志](/javascript/api/excel/excel.charttrendlineupdatedata#label)|表示图表趋势线的标签。|
||[showEquation](/javascript/api/excel/excel.charttrendlineupdatedata#showequation)|如果图表上显示趋势线公式，则为 True。|
||[showRSquared](/javascript/api/excel/excel.charttrendlineupdatedata#showrsquared)|如果图表上显示趋势线的 R 平方值，则为 True。|
|[ChartUpdateData](/javascript/api/excel/excel.chartupdatedata)|[categoryLabelLevel](/javascript/api/excel/excel.chartupdatedata#categorylabellevel)|返回或设置一个 ChartCategoryLabelLevel 枚举常量, 该常量引用|
||[displayBlanksAs](/javascript/api/excel/excel.chartupdatedata#displayblanksas)|返回或设置图表上的空白单元格的绘制方式。 读/写。|
||[plotArea](/javascript/api/excel/excel.chartupdatedata#plotarea)|表示图表的绘制区域。|
||[plotBy](/javascript/api/excel/excel.chartupdatedata#plotby)|返回或设置图表上的列或行用作数据系列的方式。 读/写。|
||[plotVisibleOnly](/javascript/api/excel/excel.chartupdatedata#plotvisibleonly)|如果仅绘制可见单元格，则为 True。如果绘制可见单元格和隐藏单元格，则为 False。 读/写。|
||[seriesNameLevel](/javascript/api/excel/excel.chartupdatedata#seriesnamelevel)|返回或设置一个 ChartSeriesNameLevel 枚举常量, 该常量引用|
||[showDataLabelsOverMaximum](/javascript/api/excel/excel.chartupdatedata#showdatalabelsovermaximum)|表示当值大于数值轴上的最大值时是否显示数据标签。|
||[style](/javascript/api/excel/excel.chartupdatedata#style)|返回或设置图表的图表样式。 读/写。|
|[CustomDataValidation](/javascript/api/excel/excel.customdatavalidation)|[formula](/javascript/api/excel/excel.customdatavalidation#formula)|自定义数据验证公式。 这将创建特殊的输入规则, 如阻止重复项或限制单元格范围中的总计。|
|[DataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|[name](/javascript/api/excel/excel.datapivothierarchy#name)|DataPivotHierarchy 的名称。|
||[numberFormat](/javascript/api/excel/excel.datapivothierarchy#numberformat)|DataPivotHierarchy 的数字格式。|
||[position](/javascript/api/excel/excel.datapivothierarchy#position)|DataPivotHierarchy 的位置。|
||[field](/javascript/api/excel/excel.datapivothierarchy#field)|返回与 DataPivotHierarchy 相关联的 PivotFields。|
||[id](/javascript/api/excel/excel.datapivothierarchy#id)|DataPivotHierarchy ID。|
||[set (properties: DataPivotHierarchy)](/javascript/api/excel/excel.datapivothierarchy#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: DataPivotHierarchyUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.datapivothierarchy#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
||[setToDefault()](/javascript/api/excel/excel.datapivothierarchy#settodefault--)|将 DataPivotHierarchy 重置回其默认值。|
||[showAs](/javascript/api/excel/excel.datapivothierarchy#showas)|确定数据是否应显示为特定计算汇总。|
||[summarizeBy](/javascript/api/excel/excel.datapivothierarchy#summarizeby)|确定是否显示 DataPivotHierarchy 的所有项。|
|[DataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|[add (pivotHierarchy: PivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection#add-pivothierarchy-)|将 PivotHierarchy 添加到当前轴。|
||[getCount()](/javascript/api/excel/excel.datapivothierarchycollection#getcount--)|获取集合中的透视层级结构的数量。|
||[getItem(name: string)](/javascript/api/excel/excel.datapivothierarchycollection#getitem-name-)|按名称或 ID 获取 DataPivotHierarchy。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.datapivothierarchycollection#getitemornullobject-name-)|按名称获取 DataPivotHierarchy。 如果 DataPivotHierarchy 不存在，则返回 Null 对象。|
||[items](/javascript/api/excel/excel.datapivothierarchycollection#items)|获取此集合中已加载的子项。|
||[remove (DataPivotHierarchy: DataPivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection#remove-datapivothierarchy-)|从当前轴删除 PivotHierarchy。|
|[DataPivotHierarchyCollectionLoadOptions](/javascript/api/excel/excel.datapivothierarchycollectionloadoptions)|[$all](/javascript/api/excel/excel.datapivothierarchycollectionloadoptions#$all)||
||[field](/javascript/api/excel/excel.datapivothierarchycollectionloadoptions#field)|对于集合中的每一项: 返回与 DataPivotHierarchy 相关联的透视字段。|
||[id](/javascript/api/excel/excel.datapivothierarchycollectionloadoptions#id)|对于集合中的每一项: DataPivotHierarchy 的 Id。|
||[name](/javascript/api/excel/excel.datapivothierarchycollectionloadoptions#name)|对于集合中的每一项: DataPivotHierarchy 的名称。|
||[numberFormat](/javascript/api/excel/excel.datapivothierarchycollectionloadoptions#numberformat)|对于集合中的每一项: DataPivotHierarchy 的数字格式。|
||[position](/javascript/api/excel/excel.datapivothierarchycollectionloadoptions#position)|对于集合中的每一项: DataPivotHierarchy 的位置。|
||[showAs](/javascript/api/excel/excel.datapivothierarchycollectionloadoptions#showas)|对于集合中的每一项: 确定是否应将数据显示为特定的汇总计算。|
||[summarizeBy](/javascript/api/excel/excel.datapivothierarchycollectionloadoptions#summarizeby)|对于集合中的每一项: 确定是否显示 DataPivotHierarchy 的所有项。|
|[DataPivotHierarchyData](/javascript/api/excel/excel.datapivothierarchydata)|[field](/javascript/api/excel/excel.datapivothierarchydata#field)|返回与 DataPivotHierarchy 相关联的 PivotFields。|
||[id](/javascript/api/excel/excel.datapivothierarchydata#id)|DataPivotHierarchy ID。|
||[name](/javascript/api/excel/excel.datapivothierarchydata#name)|DataPivotHierarchy 的名称。|
||[numberFormat](/javascript/api/excel/excel.datapivothierarchydata#numberformat)|DataPivotHierarchy 的数字格式。|
||[position](/javascript/api/excel/excel.datapivothierarchydata#position)|DataPivotHierarchy 的位置。|
||[showAs](/javascript/api/excel/excel.datapivothierarchydata#showas)|确定数据是否应显示为特定计算汇总。|
||[summarizeBy](/javascript/api/excel/excel.datapivothierarchydata#summarizeby)|确定是否显示 DataPivotHierarchy 的所有项。|
|[DataPivotHierarchyLoadOptions](/javascript/api/excel/excel.datapivothierarchyloadoptions)|[$all](/javascript/api/excel/excel.datapivothierarchyloadoptions#$all)||
||[field](/javascript/api/excel/excel.datapivothierarchyloadoptions#field)|返回与 DataPivotHierarchy 相关联的 PivotFields。|
||[id](/javascript/api/excel/excel.datapivothierarchyloadoptions#id)|DataPivotHierarchy ID。|
||[name](/javascript/api/excel/excel.datapivothierarchyloadoptions#name)|DataPivotHierarchy 的名称。|
||[numberFormat](/javascript/api/excel/excel.datapivothierarchyloadoptions#numberformat)|DataPivotHierarchy 的数字格式。|
||[position](/javascript/api/excel/excel.datapivothierarchyloadoptions#position)|DataPivotHierarchy 的位置。|
||[showAs](/javascript/api/excel/excel.datapivothierarchyloadoptions#showas)|确定数据是否应显示为特定计算汇总。|
||[summarizeBy](/javascript/api/excel/excel.datapivothierarchyloadoptions#summarizeby)|确定是否显示 DataPivotHierarchy 的所有项。|
|[DataPivotHierarchyUpdateData](/javascript/api/excel/excel.datapivothierarchyupdatedata)|[field](/javascript/api/excel/excel.datapivothierarchyupdatedata#field)|返回与 DataPivotHierarchy 相关联的 PivotFields。|
||[name](/javascript/api/excel/excel.datapivothierarchyupdatedata#name)|DataPivotHierarchy 的名称。|
||[numberFormat](/javascript/api/excel/excel.datapivothierarchyupdatedata#numberformat)|DataPivotHierarchy 的数字格式。|
||[position](/javascript/api/excel/excel.datapivothierarchyupdatedata#position)|DataPivotHierarchy 的位置。|
||[showAs](/javascript/api/excel/excel.datapivothierarchyupdatedata#showas)|确定数据是否应显示为特定计算汇总。|
||[summarizeBy](/javascript/api/excel/excel.datapivothierarchyupdatedata#summarizeby)|确定是否显示 DataPivotHierarchy 的所有项。|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[clear()](/javascript/api/excel/excel.datavalidation#clear--)|清除当前区域中的数据有效性。|
||[errorAlert](/javascript/api/excel/excel.datavalidation#erroralert)|用户输入无效数据时，出现错误警报。|
||[ignoreBlanks](/javascript/api/excel/excel.datavalidation#ignoreblanks)|忽略空白：不会对空白单元格执行数据严重，默认为 true。|
||[prompt](/javascript/api/excel/excel.datavalidation#prompt)|当用户选择单元格时提示。|
||[type](/javascript/api/excel/excel.datavalidation#type)|数据有效性类型，有关详细信息，请参阅 Excel.DataValidationType。|
||[有效](/javascript/api/excel/excel.datavalidation#valid)|表示所有单元格值根据数据有效性规则是否全部有效。|
||[标尺](/javascript/api/excel/excel.datavalidation#rule)|包含不同类型的数据验证条件的数据有效性规则。|
||[set (properties: DataValidation)](/javascript/api/excel/excel.datavalidation#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: DataValidationUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.datavalidation#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
|[DataValidationData](/javascript/api/excel/excel.datavalidationdata)|[errorAlert](/javascript/api/excel/excel.datavalidationdata#erroralert)|用户输入无效数据时，出现错误警报。|
||[ignoreBlanks](/javascript/api/excel/excel.datavalidationdata#ignoreblanks)|忽略空白：不会对空白单元格执行数据严重，默认为 true。|
||[prompt](/javascript/api/excel/excel.datavalidationdata#prompt)|当用户选择单元格时提示。|
||[标尺](/javascript/api/excel/excel.datavalidationdata#rule)|包含不同类型的数据验证条件的数据有效性规则。|
||[type](/javascript/api/excel/excel.datavalidationdata#type)|数据有效性类型，有关详细信息，请参阅 Excel.DataValidationType。|
||[有效](/javascript/api/excel/excel.datavalidationdata#valid)|表示所有单元格值根据数据有效性规则是否全部有效。|
|[DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|[邮件](/javascript/api/excel/excel.datavalidationerroralert#message)|表示错误警报消息。|
||[showAlert](/javascript/api/excel/excel.datavalidationerroralert#showalert)|确定在用户输入无效数据时是否显示错误警报对话框。 默认值为 true。|
||[style](/javascript/api/excel/excel.datavalidationerroralert#style)|表示数据有效性警报类型，有关详细信息，请参阅 Excel.DataValidationAlertStyle。|
||[title](/javascript/api/excel/excel.datavalidationerroralert#title)|表示错误警报对话框标题。|
|[DataValidationLoadOptions](/javascript/api/excel/excel.datavalidationloadoptions)|[$all](/javascript/api/excel/excel.datavalidationloadoptions#$all)||
||[errorAlert](/javascript/api/excel/excel.datavalidationloadoptions#erroralert)|用户输入无效数据时，出现错误警报。|
||[ignoreBlanks](/javascript/api/excel/excel.datavalidationloadoptions#ignoreblanks)|忽略空白：不会对空白单元格执行数据严重，默认为 true。|
||[prompt](/javascript/api/excel/excel.datavalidationloadoptions#prompt)|当用户选择单元格时提示。|
||[标尺](/javascript/api/excel/excel.datavalidationloadoptions#rule)|包含不同类型的数据验证条件的数据有效性规则。|
||[type](/javascript/api/excel/excel.datavalidationloadoptions#type)|数据有效性类型，有关详细信息，请参阅 Excel.DataValidationType。|
||[有效](/javascript/api/excel/excel.datavalidationloadoptions#valid)|表示所有单元格值根据数据有效性规则是否全部有效。|
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
|[DataValidationUpdateData](/javascript/api/excel/excel.datavalidationupdatedata)|[errorAlert](/javascript/api/excel/excel.datavalidationupdatedata#erroralert)|用户输入无效数据时，出现错误警报。|
||[ignoreBlanks](/javascript/api/excel/excel.datavalidationupdatedata#ignoreblanks)|忽略空白：不会对空白单元格执行数据严重，默认为 true。|
||[prompt](/javascript/api/excel/excel.datavalidationupdatedata#prompt)|当用户选择单元格时提示。|
||[标尺](/javascript/api/excel/excel.datavalidationupdatedata#rule)|包含不同类型的数据验证条件的数据有效性规则。|
|[DateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|[formula1](/javascript/api/excel/excel.datetimedatavalidation#formula1)|当 operator 属性设置为二元运算符 (如 GreaterThan (左边的操作数是用户试图在单元格中输入的值) 时, 指定右边的操作数。 使用和 NotBetween 之间的三元运算符指定下界操作数。|
||[formula2](/javascript/api/excel/excel.datetimedatavalidation#formula2)|使用和 NotBetween 之间的三元运算符指定上界操作数。 不与二元运算符 (如 GreaterThan) 一起使用。|
||[接线员](/javascript/api/excel/excel.datetimedatavalidation#operator)|用于验证数据有效性的运算符。|
|[FilterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|[enableMultipleFilterItems](/javascript/api/excel/excel.filterpivothierarchy#enablemultiplefilteritems)|确定是否允许多个筛选项。|
||[name](/javascript/api/excel/excel.filterpivothierarchy#name)|FilterPivotHierarchy 的名称。|
||[position](/javascript/api/excel/excel.filterpivothierarchy#position)|FilterPivotHierarchy 的位置。|
||[fields](/javascript/api/excel/excel.filterpivothierarchy#fields)|返回与 FilterPivotHierarchy 相关联的 PivotFields。|
||[id](/javascript/api/excel/excel.filterpivothierarchy#id)|FilterPivotHierarchy 的 ID。|
||[set (properties: FilterPivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchy#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: FilterPivotHierarchyUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.filterpivothierarchy#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
||[setToDefault()](/javascript/api/excel/excel.filterpivothierarchy#settodefault--)|将 FilterPivotHierarchy 重置回其默认值。|
|[FilterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|[add (pivotHierarchy: PivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection#add-pivothierarchy-)|将 PivotHierarchy 添加到当前轴。 行和列上的其他位置是否存在层次结构。|
||[getCount()](/javascript/api/excel/excel.filterpivothierarchycollection#getcount--)|获取集合中的透视层级结构的数量。|
||[getItem(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection#getitem-name-)|按名称或 ID 获取 FilterPivotHierarchy。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection#getitemornullobject-name-)|按名称获取 FilterPivotHierarchy。 如果 FilterPivotHierarchy 不存在，则返回 Null 对象。|
||[items](/javascript/api/excel/excel.filterpivothierarchycollection#items)|获取此集合中已加载的子项。|
||[remove (filterPivotHierarchy: FilterPivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection#remove-filterpivothierarchy-)|从当前轴删除 PivotHierarchy。|
|[FilterPivotHierarchyCollectionLoadOptions](/javascript/api/excel/excel.filterpivothierarchycollectionloadoptions)|[$all](/javascript/api/excel/excel.filterpivothierarchycollectionloadoptions#$all)||
||[enableMultipleFilterItems](/javascript/api/excel/excel.filterpivothierarchycollectionloadoptions#enablemultiplefilteritems)|对于集合中的每个项目: 确定是否允许多个筛选器项目。|
||[id](/javascript/api/excel/excel.filterpivothierarchycollectionloadoptions#id)|对于集合中的每一项: FilterPivotHierarchy 的 Id。|
||[name](/javascript/api/excel/excel.filterpivothierarchycollectionloadoptions#name)|对于集合中的每一项: FilterPivotHierarchy 的名称。|
||[position](/javascript/api/excel/excel.filterpivothierarchycollectionloadoptions#position)|对于集合中的每一项: FilterPivotHierarchy 的位置。|
|[FilterPivotHierarchyData](/javascript/api/excel/excel.filterpivothierarchydata)|[enableMultipleFilterItems](/javascript/api/excel/excel.filterpivothierarchydata#enablemultiplefilteritems)|确定是否允许多个筛选项。|
||[fields](/javascript/api/excel/excel.filterpivothierarchydata#fields)|返回与 FilterPivotHierarchy 相关联的 PivotFields。|
||[id](/javascript/api/excel/excel.filterpivothierarchydata#id)|FilterPivotHierarchy 的 ID。|
||[name](/javascript/api/excel/excel.filterpivothierarchydata#name)|FilterPivotHierarchy 的名称。|
||[position](/javascript/api/excel/excel.filterpivothierarchydata#position)|FilterPivotHierarchy 的位置。|
|[FilterPivotHierarchyLoadOptions](/javascript/api/excel/excel.filterpivothierarchyloadoptions)|[$all](/javascript/api/excel/excel.filterpivothierarchyloadoptions#$all)||
||[enableMultipleFilterItems](/javascript/api/excel/excel.filterpivothierarchyloadoptions#enablemultiplefilteritems)|确定是否允许多个筛选项。|
||[id](/javascript/api/excel/excel.filterpivothierarchyloadoptions#id)|FilterPivotHierarchy 的 ID。|
||[name](/javascript/api/excel/excel.filterpivothierarchyloadoptions#name)|FilterPivotHierarchy 的名称。|
||[position](/javascript/api/excel/excel.filterpivothierarchyloadoptions#position)|FilterPivotHierarchy 的位置。|
|[FilterPivotHierarchyUpdateData](/javascript/api/excel/excel.filterpivothierarchyupdatedata)|[enableMultipleFilterItems](/javascript/api/excel/excel.filterpivothierarchyupdatedata#enablemultiplefilteritems)|确定是否允许多个筛选项。|
||[name](/javascript/api/excel/excel.filterpivothierarchyupdatedata#name)|FilterPivotHierarchy 的名称。|
||[position](/javascript/api/excel/excel.filterpivothierarchyupdatedata#position)|FilterPivotHierarchy 的位置。|
|[ListDataValidation](/javascript/api/excel/excel.listdatavalidation)|[inCellDropDown](/javascript/api/excel/excel.listdatavalidation#incelldropdown)|是否显示单元格下拉菜单中的列表，默认为 true。|
||[source](/javascript/api/excel/excel.listdatavalidation#source)|数据有效性列表源|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[name](/javascript/api/excel/excel.pivotfield#name)|PivotField 的名称。|
||[id](/javascript/api/excel/excel.pivotfield#id)|PivotField 的 ID。|
||[items](/javascript/api/excel/excel.pivotfield#items)|返回包含透视字段的 PivotItems。|
||[set (properties: Excel. 透视字段)](/javascript/api/excel/excel.pivotfield#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: PivotFieldUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.pivotfield#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
||[showAllItems](/javascript/api/excel/excel.pivotfield#showallitems)|确定是否显示 PivotField 的所有项。|
||[sortByLabels (sortBy: SortBy)](/javascript/api/excel/excel.pivotfield#sortbylabels-sortby-)|PivotField 排序。 如果指定 DataPivotHierarchy，则会基于它进行排序，如果未指定，则会基于 PivotField 本身进行排序。|
||[subtotals](/javascript/api/excel/excel.pivotfield#subtotals)|PivotField 小计。|
|[PivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|[getCount()](/javascript/api/excel/excel.pivotfieldcollection#getcount--)|获取集合中的数据透视字段数。|
||[getItem(name: string)](/javascript/api/excel/excel.pivotfieldcollection#getitem-name-)|按其名称或 id 获取透视字段。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotfieldcollection#getitemornullobject-name-)|按名称获取透视字段。 如果透视字段不存在, 则将返回 null 对象。|
||[items](/javascript/api/excel/excel.pivotfieldcollection#items)|获取此集合中已加载的子项。|
|[PivotFieldCollectionLoadOptions](/javascript/api/excel/excel.pivotfieldcollectionloadoptions)|[$all](/javascript/api/excel/excel.pivotfieldcollectionloadoptions#$all)||
||[id](/javascript/api/excel/excel.pivotfieldcollectionloadoptions#id)|对于集合中的每一项: 透视字段的 Id。|
||[name](/javascript/api/excel/excel.pivotfieldcollectionloadoptions#name)|对于集合中的每一项: 透视字段的名称。|
||[showAllItems](/javascript/api/excel/excel.pivotfieldcollectionloadoptions#showallitems)|对于集合中的每一项: 确定是否显示透视字段的所有项。|
||[subtotals](/javascript/api/excel/excel.pivotfieldcollectionloadoptions#subtotals)|对于集合中的每一项: 透视字段的小计。|
|[PivotFieldData](/javascript/api/excel/excel.pivotfielddata)|[id](/javascript/api/excel/excel.pivotfielddata#id)|PivotField 的 ID。|
||[items](/javascript/api/excel/excel.pivotfielddata#items)|返回与 PivotField 相关联的 PivotFields。|
||[name](/javascript/api/excel/excel.pivotfielddata#name)|PivotField 的名称。|
||[showAllItems](/javascript/api/excel/excel.pivotfielddata#showallitems)|确定是否显示 PivotField 的所有项。|
||[subtotals](/javascript/api/excel/excel.pivotfielddata#subtotals)|PivotField 小计。|
|[PivotFieldLoadOptions](/javascript/api/excel/excel.pivotfieldloadoptions)|[$all](/javascript/api/excel/excel.pivotfieldloadoptions#$all)||
||[id](/javascript/api/excel/excel.pivotfieldloadoptions#id)|PivotField 的 ID。|
||[name](/javascript/api/excel/excel.pivotfieldloadoptions#name)|PivotField 的名称。|
||[showAllItems](/javascript/api/excel/excel.pivotfieldloadoptions#showallitems)|确定是否显示 PivotField 的所有项。|
||[subtotals](/javascript/api/excel/excel.pivotfieldloadoptions#subtotals)|PivotField 小计。|
|[PivotFieldUpdateData](/javascript/api/excel/excel.pivotfieldupdatedata)|[name](/javascript/api/excel/excel.pivotfieldupdatedata#name)|PivotField 的名称。|
||[showAllItems](/javascript/api/excel/excel.pivotfieldupdatedata#showallitems)|确定是否显示 PivotField 的所有项。|
||[subtotals](/javascript/api/excel/excel.pivotfieldupdatedata#subtotals)|PivotField 小计。|
|[PivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|[name](/javascript/api/excel/excel.pivothierarchy#name)|PivotHierarchy 的名称。|
||[fields](/javascript/api/excel/excel.pivothierarchy#fields)|返回与 PivotHierarchy 相关联的 PivotFields。|
||[id](/javascript/api/excel/excel.pivothierarchy#id)|PivotHierarchy 的 ID。|
||[set (properties: PivotHierarchy)](/javascript/api/excel/excel.pivothierarchy#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: PivotHierarchyUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.pivothierarchy#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
|[PivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|[getCount()](/javascript/api/excel/excel.pivothierarchycollection#getcount--)|获取集合中的透视层级结构的数量。|
||[getItem(name: string)](/javascript/api/excel/excel.pivothierarchycollection#getitem-name-)|按名称或 ID 获取 PivotHierarchy。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivothierarchycollection#getitemornullobject-name-)|按名称获取 PivotHierarchy。 如果 PivotHierarchy 不存在，则返回 Null 对象。|
||[items](/javascript/api/excel/excel.pivothierarchycollection#items)|获取此集合中已加载的子项。|
|[PivotHierarchyCollectionLoadOptions](/javascript/api/excel/excel.pivothierarchycollectionloadoptions)|[$all](/javascript/api/excel/excel.pivothierarchycollectionloadoptions#$all)||
||[id](/javascript/api/excel/excel.pivothierarchycollectionloadoptions#id)|对于集合中的每一项: PivotHierarchy 的 Id。|
||[name](/javascript/api/excel/excel.pivothierarchycollectionloadoptions#name)|对于集合中的每一项: PivotHierarchy 的名称。|
|[PivotHierarchyData](/javascript/api/excel/excel.pivothierarchydata)|[fields](/javascript/api/excel/excel.pivothierarchydata#fields)|返回与 PivotHierarchy 相关联的 PivotFields。|
||[id](/javascript/api/excel/excel.pivothierarchydata#id)|PivotHierarchy 的 ID。|
||[name](/javascript/api/excel/excel.pivothierarchydata#name)|PivotHierarchy 的名称。|
|[PivotHierarchyLoadOptions](/javascript/api/excel/excel.pivothierarchyloadoptions)|[$all](/javascript/api/excel/excel.pivothierarchyloadoptions#$all)||
||[id](/javascript/api/excel/excel.pivothierarchyloadoptions#id)|PivotHierarchy 的 ID。|
||[name](/javascript/api/excel/excel.pivothierarchyloadoptions#name)|PivotHierarchy 的名称。|
|[PivotHierarchyUpdateData](/javascript/api/excel/excel.pivothierarchyupdatedata)|[name](/javascript/api/excel/excel.pivothierarchyupdatedata#name)|PivotHierarchy 的名称。|
|[PivotItem](/javascript/api/excel/excel.pivotitem)|[isExpanded](/javascript/api/excel/excel.pivotitem#isexpanded)|确定是展开项以显示子项还是折叠项并隐藏子项。|
||[name](/javascript/api/excel/excel.pivotitem#name)|PivotItem 的名称。|
||[id](/javascript/api/excel/excel.pivotitem#id)|PivotItem 的 ID。|
||[set (properties: PivotItem)](/javascript/api/excel/excel.pivotitem#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: PivotItemUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.pivotitem#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
||[visible](/javascript/api/excel/excel.pivotitem#visible)|确定 PivotItem 是否可见。|
|[PivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|[getCount()](/javascript/api/excel/excel.pivotitemcollection#getcount--)|获取集合中的数据透视项的数目。|
||[getItem(name: string)](/javascript/api/excel/excel.pivotitemcollection#getitem-name-)|按其名称或 id 获取 PivotItem。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotitemcollection#getitemornullobject-name-)|按名称获取 PivotItem。 如果 PivotItem 不存在, 则将返回一个 null 对象。|
||[items](/javascript/api/excel/excel.pivotitemcollection#items)|获取此集合中已加载的子项。|
|[PivotItemCollectionLoadOptions](/javascript/api/excel/excel.pivotitemcollectionloadoptions)|[$all](/javascript/api/excel/excel.pivotitemcollectionloadoptions#$all)||
||[id](/javascript/api/excel/excel.pivotitemcollectionloadoptions#id)|对于集合中的每一项: PivotItem 的 Id。|
||[isExpanded](/javascript/api/excel/excel.pivotitemcollectionloadoptions#isexpanded)|对于集合中的每个项目: 确定是否展开项目以显示子项目, 或者是否折叠和子项被隐藏。|
||[name](/javascript/api/excel/excel.pivotitemcollectionloadoptions#name)|对于集合中的每一项: PivotItem 的名称。|
||[visible](/javascript/api/excel/excel.pivotitemcollectionloadoptions#visible)|对于集合中的每一项: 确定 PivotItem 是否可见。|
|[PivotItemData](/javascript/api/excel/excel.pivotitemdata)|[id](/javascript/api/excel/excel.pivotitemdata#id)|PivotItem 的 ID。|
||[isExpanded](/javascript/api/excel/excel.pivotitemdata#isexpanded)|确定是展开项以显示子项还是折叠项并隐藏子项。|
||[name](/javascript/api/excel/excel.pivotitemdata#name)|PivotItem 的名称。|
||[visible](/javascript/api/excel/excel.pivotitemdata#visible)|确定 PivotItem 是否可见。|
|[PivotItemLoadOptions](/javascript/api/excel/excel.pivotitemloadoptions)|[$all](/javascript/api/excel/excel.pivotitemloadoptions#$all)||
||[id](/javascript/api/excel/excel.pivotitemloadoptions#id)|PivotItem 的 ID。|
||[isExpanded](/javascript/api/excel/excel.pivotitemloadoptions#isexpanded)|确定是展开项以显示子项还是折叠项并隐藏子项。|
||[name](/javascript/api/excel/excel.pivotitemloadoptions#name)|PivotItem 的名称。|
||[visible](/javascript/api/excel/excel.pivotitemloadoptions#visible)|确定 PivotItem 是否可见。|
|[PivotItemUpdateData](/javascript/api/excel/excel.pivotitemupdatedata)|[isExpanded](/javascript/api/excel/excel.pivotitemupdatedata#isexpanded)|确定是展开项以显示子项还是折叠项并隐藏子项。|
||[name](/javascript/api/excel/excel.pivotitemupdatedata#name)|PivotItem 的名称。|
||[visible](/javascript/api/excel/excel.pivotitemupdatedata#visible)|确定 PivotItem 是否可见。|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getColumnLabelRange()](/javascript/api/excel/excel.pivotlayout#getcolumnlabelrange--)|返回数据透视表列标签所在位置的区域。|
||[getDataBodyRange()](/javascript/api/excel/excel.pivotlayout#getdatabodyrange--)|返回数据透视表数据值所在位置的区域。|
||[getFilterAxisRange()](/javascript/api/excel/excel.pivotlayout#getfilteraxisrange--)|返回数据透视表筛选区的区域。|
||[getRange()](/javascript/api/excel/excel.pivotlayout#getrange--)|返回存在数据透视表的区域，不包括筛选区。|
||[getRowLabelRange()](/javascript/api/excel/excel.pivotlayout#getrowlabelrange--)|返回数据透视表行标签所在位置的区域。|
||[layoutType](/javascript/api/excel/excel.pivotlayout#layouttype)|此属性指示数据透视表上的所有字段的 PivotLayoutType。 如果字段状态不同，则为 null。|
||[set (properties: PivotLayout)](/javascript/api/excel/excel.pivotlayout#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: PivotLayoutUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.pivotlayout#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
||[showColumnGrandTotals](/javascript/api/excel/excel.pivotlayout#showcolumngrandtotals)|指定数据透视表报表是否显示列总计。|
||[showRowGrandTotals](/javascript/api/excel/excel.pivotlayout#showrowgrandtotals)|指定数据透视表报表是否显示行总计。|
||[subtotalLocation](/javascript/api/excel/excel.pivotlayout#subtotallocation)|此属性指示数据透视表上的所有字段的 SubtotalLocationType。 如果字段状态不同，则为 null。|
|[PivotLayoutData](/javascript/api/excel/excel.pivotlayoutdata)|[layoutType](/javascript/api/excel/excel.pivotlayoutdata#layouttype)|此属性指示数据透视表上的所有字段的 PivotLayoutType。 如果字段状态不同，则为 null。|
||[showColumnGrandTotals](/javascript/api/excel/excel.pivotlayoutdata#showcolumngrandtotals)|指定数据透视表报表是否显示列总计。|
||[showRowGrandTotals](/javascript/api/excel/excel.pivotlayoutdata#showrowgrandtotals)|指定数据透视表报表是否显示行总计。|
||[subtotalLocation](/javascript/api/excel/excel.pivotlayoutdata#subtotallocation)|此属性指示数据透视表上的所有字段的 SubtotalLocationType。 如果字段状态不同，则为 null。|
|[PivotLayoutLoadOptions](/javascript/api/excel/excel.pivotlayoutloadoptions)|[$all](/javascript/api/excel/excel.pivotlayoutloadoptions#$all)||
||[layoutType](/javascript/api/excel/excel.pivotlayoutloadoptions#layouttype)|此属性指示数据透视表上的所有字段的 PivotLayoutType。 如果字段状态不同，则为 null。|
||[showColumnGrandTotals](/javascript/api/excel/excel.pivotlayoutloadoptions#showcolumngrandtotals)|指定数据透视表报表是否显示列总计。|
||[showRowGrandTotals](/javascript/api/excel/excel.pivotlayoutloadoptions#showrowgrandtotals)|指定数据透视表报表是否显示行总计。|
||[subtotalLocation](/javascript/api/excel/excel.pivotlayoutloadoptions#subtotallocation)|此属性指示数据透视表上的所有字段的 SubtotalLocationType。 如果字段状态不同，则为 null。|
|[PivotLayoutUpdateData](/javascript/api/excel/excel.pivotlayoutupdatedata)|[layoutType](/javascript/api/excel/excel.pivotlayoutupdatedata#layouttype)|此属性指示数据透视表上的所有字段的 PivotLayoutType。 如果字段状态不同，则为 null。|
||[showColumnGrandTotals](/javascript/api/excel/excel.pivotlayoutupdatedata#showcolumngrandtotals)|指定数据透视表报表是否显示列总计。|
||[showRowGrandTotals](/javascript/api/excel/excel.pivotlayoutupdatedata#showrowgrandtotals)|指定数据透视表报表是否显示行总计。|
||[subtotalLocation](/javascript/api/excel/excel.pivotlayoutupdatedata#subtotallocation)|此属性指示数据透视表上的所有字段的 SubtotalLocationType。 如果字段状态不同，则为 null。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[delete()](/javascript/api/excel/excel.pivottable#delete--)|删除 PivotTable 对象。|
||[columnHierarchies](/javascript/api/excel/excel.pivottable#columnhierarchies)|数据透视表的列透视层级结构。|
||[dataHierarchies](/javascript/api/excel/excel.pivottable#datahierarchies)|数据透视表的数据透视层级结构。|
||[filterHierarchies](/javascript/api/excel/excel.pivottable#filterhierarchies)|数据透视表的筛选器透视层级结构。|
||[层次结构](/javascript/api/excel/excel.pivottable#hierarchies)|数据透视表的透视层级结构。|
||[布局](/javascript/api/excel/excel.pivottable#layout)|PivotLayout，用于说明数据透视表的布局和可视化结构。|
||[rowHierarchies](/javascript/api/excel/excel.pivottable#rowhierarchies)|数据透视表的行透视层级结构。|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[add (name: string, source: Range \| string \| Table, destination: Range \| string)](/javascript/api/excel/excel.pivottablecollection#add-name--source--destination-)|基于指定的数据源添加数据透视表，并将其插入到目标区域的左上单元格。|
|[PivotTableCollectionLoadOptions](/javascript/api/excel/excel.pivottablecollectionloadoptions)|[布局](/javascript/api/excel/excel.pivottablecollectionloadoptions#layout)|对于集合中的每一项: 描述数据透视表的布局和可视结构的 PivotLayout。|
|[PivotTableData](/javascript/api/excel/excel.pivottabledata)|[columnHierarchies](/javascript/api/excel/excel.pivottabledata#columnhierarchies)|数据透视表的列透视层级结构。|
||[dataHierarchies](/javascript/api/excel/excel.pivottabledata#datahierarchies)|数据透视表的数据透视层级结构。|
||[filterHierarchies](/javascript/api/excel/excel.pivottabledata#filterhierarchies)|数据透视表的筛选器透视层级结构。|
||[层次结构](/javascript/api/excel/excel.pivottabledata#hierarchies)|数据透视表的透视层级结构。|
||[rowHierarchies](/javascript/api/excel/excel.pivottabledata#rowhierarchies)|数据透视表的行透视层级结构。|
|[PivotTableLoadOptions](/javascript/api/excel/excel.pivottableloadoptions)|[布局](/javascript/api/excel/excel.pivottableloadoptions#layout)|PivotLayout，用于说明数据透视表的布局和可视化结构。|
|[Range](/javascript/api/excel/excel.range)|[dataValidation](/javascript/api/excel/excel.range#datavalidation)|返回数据有效性对象。|
|[RangeData](/javascript/api/excel/excel.rangedata)|[dataValidation](/javascript/api/excel/excel.rangedata#datavalidation)|返回数据有效性对象。|
|[RangeLoadOptions](/javascript/api/excel/excel.rangeloadoptions)|[dataValidation](/javascript/api/excel/excel.rangeloadoptions#datavalidation)|返回数据有效性对象。|
|[RangeUpdateData](/javascript/api/excel/excel.rangeupdatedata)|[dataValidation](/javascript/api/excel/excel.rangeupdatedata#datavalidation)|返回数据有效性对象。|
|[RowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|[name](/javascript/api/excel/excel.rowcolumnpivothierarchy#name)|RowColumnPivotHierarchy 的名称。|
||[position](/javascript/api/excel/excel.rowcolumnpivothierarchy#position)|RowColumnPivotHierarchy 的位置。|
||[fields](/javascript/api/excel/excel.rowcolumnpivothierarchy#fields)|返回与 RowColumnPivotHierarchy 相关联的 PivotFields。|
||[id](/javascript/api/excel/excel.rowcolumnpivothierarchy#id)|RowColumnPivotHierarchy 的 ID。|
||[set (properties: RowColumnPivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchy#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: RowColumnPivotHierarchyUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.rowcolumnpivothierarchy#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
||[setToDefault()](/javascript/api/excel/excel.rowcolumnpivothierarchy#settodefault--)|将 RowColumnPivotHierarchy 重置回其默认值。|
|[RowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|[add (pivotHierarchy: PivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#add-pivothierarchy-)|将 PivotHierarchy 添加到当前轴。 行和列上的其他位置是否存在层次结构。|
||[getCount()](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getcount--)|获取集合中的透视层级结构的数量。|
||[getItem(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getitem-name-)|按名称或 ID 获取 RowColumnPivotHierarchy。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getitemornullobject-name-)|按名称获取 RowColumnPivotHierarchy。 如果 RowColumnPivotHierarchy 不存在，则返回 Null 对象。|
||[items](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#items)|获取此集合中已加载的子项。|
||[remove (rowColumnPivotHierarchy: RowColumnPivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#remove-rowcolumnpivothierarchy-)|从当前轴删除 PivotHierarchy。|
|[RowColumnPivotHierarchyCollectionLoadOptions](/javascript/api/excel/excel.rowcolumnpivothierarchycollectionloadoptions)|[$all](/javascript/api/excel/excel.rowcolumnpivothierarchycollectionloadoptions#$all)||
||[id](/javascript/api/excel/excel.rowcolumnpivothierarchycollectionloadoptions#id)|对于集合中的每一项: RowColumnPivotHierarchy 的 Id。|
||[name](/javascript/api/excel/excel.rowcolumnpivothierarchycollectionloadoptions#name)|对于集合中的每一项: RowColumnPivotHierarchy 的名称。|
||[position](/javascript/api/excel/excel.rowcolumnpivothierarchycollectionloadoptions#position)|对于集合中的每一项: RowColumnPivotHierarchy 的位置。|
|[RowColumnPivotHierarchyData](/javascript/api/excel/excel.rowcolumnpivothierarchydata)|[fields](/javascript/api/excel/excel.rowcolumnpivothierarchydata#fields)|返回与 RowColumnPivotHierarchy 相关联的 PivotFields。|
||[id](/javascript/api/excel/excel.rowcolumnpivothierarchydata#id)|RowColumnPivotHierarchy 的 ID。|
||[name](/javascript/api/excel/excel.rowcolumnpivothierarchydata#name)|RowColumnPivotHierarchy 的名称。|
||[position](/javascript/api/excel/excel.rowcolumnpivothierarchydata#position)|RowColumnPivotHierarchy 的位置。|
|[RowColumnPivotHierarchyLoadOptions](/javascript/api/excel/excel.rowcolumnpivothierarchyloadoptions)|[$all](/javascript/api/excel/excel.rowcolumnpivothierarchyloadoptions#$all)||
||[id](/javascript/api/excel/excel.rowcolumnpivothierarchyloadoptions#id)|RowColumnPivotHierarchy 的 ID。|
||[name](/javascript/api/excel/excel.rowcolumnpivothierarchyloadoptions#name)|RowColumnPivotHierarchy 的名称。|
||[position](/javascript/api/excel/excel.rowcolumnpivothierarchyloadoptions#position)|RowColumnPivotHierarchy 的位置。|
|[RowColumnPivotHierarchyUpdateData](/javascript/api/excel/excel.rowcolumnpivothierarchyupdatedata)|[name](/javascript/api/excel/excel.rowcolumnpivothierarchyupdatedata#name)|RowColumnPivotHierarchy 的名称。|
||[position](/javascript/api/excel/excel.rowcolumnpivothierarchyupdatedata#position)|RowColumnPivotHierarchy 的位置。|
|[语言](/javascript/api/excel/excel.runtime)|[enableEvents](/javascript/api/excel/excel.runtime#enableevents)|切换当前任务窗格或内容加载项中的 JavaScript 事件。|
|[RuntimeData](/javascript/api/excel/excel.runtimedata)|[enableEvents](/javascript/api/excel/excel.runtimedata#enableevents)|切换当前任务窗格或内容加载项中的 JavaScript 事件。|
|[RuntimeLoadOptions](/javascript/api/excel/excel.runtimeloadoptions)|[enableEvents](/javascript/api/excel/excel.runtimeloadoptions#enableevents)|切换当前任务窗格或内容加载项中的 JavaScript 事件。|
|[RuntimeUpdateData](/javascript/api/excel/excel.runtimeupdatedata)|[enableEvents](/javascript/api/excel/excel.runtimeupdatedata#enableevents)|切换当前任务窗格或内容加载项中的 JavaScript 事件。|
|[ShowAsRule](/javascript/api/excel/excel.showasrule)|[baseField](/javascript/api/excel/excel.showasrule#basefield)|基于 ShowAs 计算的基础 PivotField，如适用，基于 ShowAsCalculation 类型，否则为 null。|
||[baseItem](/javascript/api/excel/excel.showasrule#baseitem)|基于 ShowAs 计算的基础 Item，如适用，基于 ShowAsCalculation 类型，否则为 null。|
||[结果](/javascript/api/excel/excel.showasrule#calculation)|数据 PivotField 使用的 ShowAs 计算。 有关详细信息, 请参阅 ShowAsCalculation。|
|[Style](/javascript/api/excel/excel.style)|[autoIndent](/javascript/api/excel/excel.style#autoindent)|指示将单元格中的文本对齐方式设为相等分布时文本是否会自动缩进。|
||[textOrientation](/javascript/api/excel/excel.style#textorientation)|此样式中的文本方向。|
|[StyleCollectionLoadOptions](/javascript/api/excel/excel.stylecollectionloadoptions)|[autoIndent](/javascript/api/excel/excel.stylecollectionloadoptions#autoindent)|对于集合中的每一项: 指示当单元格中文本的对齐方式设置为相等分布时, 文本是否自动缩进。|
||[textOrientation](/javascript/api/excel/excel.stylecollectionloadoptions#textorientation)|对于集合中的每一项: 样式的文本方向。|
|[StyleData](/javascript/api/excel/excel.styledata)|[autoIndent](/javascript/api/excel/excel.styledata#autoindent)|指示将单元格中的文本对齐方式设为相等分布时文本是否会自动缩进。|
||[textOrientation](/javascript/api/excel/excel.styledata#textorientation)|此样式中的文本方向。|
|[StyleLoadOptions](/javascript/api/excel/excel.styleloadoptions)|[autoIndent](/javascript/api/excel/excel.styleloadoptions#autoindent)|指示将单元格中的文本对齐方式设为相等分布时文本是否会自动缩进。|
||[textOrientation](/javascript/api/excel/excel.styleloadoptions#textorientation)|此样式中的文本方向。|
|[StyleUpdateData](/javascript/api/excel/excel.styleupdatedata)|[autoIndent](/javascript/api/excel/excel.styleupdatedata#autoindent)|指示将单元格中的文本对齐方式设为相等分布时文本是否会自动缩进。|
||[textOrientation](/javascript/api/excel/excel.styleupdatedata#textorientation)|此样式中的文本方向。|
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
|[TableCollectionLoadOptions](/javascript/api/excel/excel.tablecollectionloadoptions)|[legacyId](/javascript/api/excel/excel.tablecollectionloadoptions#legacyid)|对于集合中的每一项: 返回一个数字 id。|
|[TableData](/javascript/api/excel/excel.tabledata)|[legacyId](/javascript/api/excel/excel.tabledata#legacyid)|返回一个数字 id。|
|[TableLoadOptions](/javascript/api/excel/excel.tableloadoptions)|[legacyId](/javascript/api/excel/excel.tableloadoptions#legacyid)|返回一个数字 id。|
|[Workbook](/javascript/api/excel/excel.workbook)|[readOnly](/javascript/api/excel/excel.workbook#readonly)|如果在只读模式下打开工作簿，则为 True。 只读。|
|[WorkbookCreated](/javascript/api/excel/excel.workbookcreated)||[WorkbookData](/javascript/api/excel/excel.workbookdata)|[readOnly](/javascript/api/excel/excel.workbookdata#readonly)|如果在只读模式下打开工作簿，则为 True。 只读。|
|[WorkbookLoadOptions](/javascript/api/excel/excel.workbookloadoptions)|[readOnly](/javascript/api/excel/excel.workbookloadoptions#readonly)|如果在只读模式下打开工作簿，则为 True。 只读。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onCalculated](/javascript/api/excel/excel.worksheet#oncalculated)|在计算工作表时发生。|
||[showGridlines](/javascript/api/excel/excel.worksheet#showgridlines)|获取或设置工作表的网格线标志。|
||[showHeadings](/javascript/api/excel/excel.worksheet#showheadings)|获取或设置工作表的标题标志。|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[type](/javascript/api/excel/excel.worksheetcalculatedeventargs#type)|获取事件的类型。 有关详细信息，请参阅 Excel.EventType。|
||[worksheetId](/javascript/api/excel/excel.worksheetcalculatedeventargs#worksheetid)|获取计算的工作表的 ID。|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#getrange-ctx-)|获取区域，该区域表示特定工作表上的更改区域。|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#getrangeornullobject-ctx-)|获取区域，该区域表示特定工作表上的更改区域。 它可能会返回 null 对象。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onCalculated](/javascript/api/excel/excel.worksheetcollection#oncalculated)|计算工作簿中的任何工作表时发生。|
|[WorksheetCollectionLoadOptions](/javascript/api/excel/excel.worksheetcollectionloadoptions)|[showGridlines](/javascript/api/excel/excel.worksheetcollectionloadoptions#showgridlines)|对于集合中的每一项: 获取或设置工作表的网格线标志。|
||[showHeadings](/javascript/api/excel/excel.worksheetcollectionloadoptions#showheadings)|对于集合中的每一项: 获取或设置工作表的标题标志。|
|[WorksheetData](/javascript/api/excel/excel.worksheetdata)|[showGridlines](/javascript/api/excel/excel.worksheetdata#showgridlines)|获取或设置工作表的网格线标志。|
||[showHeadings](/javascript/api/excel/excel.worksheetdata#showheadings)|获取或设置工作表的标题标志。|
|[WorksheetLoadOptions](/javascript/api/excel/excel.worksheetloadoptions)|[showGridlines](/javascript/api/excel/excel.worksheetloadoptions#showgridlines)|获取或设置工作表的网格线标志。|
||[showHeadings](/javascript/api/excel/excel.worksheetloadoptions#showheadings)|获取或设置工作表的标题标志。|
|[WorksheetUpdateData](/javascript/api/excel/excel.worksheetupdatedata)|[showGridlines](/javascript/api/excel/excel.worksheetupdatedata#showgridlines)|获取或设置工作表的网格线标志。|
||[showHeadings](/javascript/api/excel/excel.worksheetupdatedata#showheadings)|获取或设置工作表的标题标志。|

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 参考文档](/javascript/api/excel)
- [Excel JavaScript API 要求集](./excel-api-requirement-sets.md)
