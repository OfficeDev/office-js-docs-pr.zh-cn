---
title: Excel JavaScript API 要求集
description: ''
ms.date: 10/09/2018
localization_priority: Priority
ms.openlocfilehash: fdcbee0374851f0f88130ae8afe28eec3a0fe77c
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/23/2019
ms.locfileid: "29388722"
---
# <a name="excel-javascript-api-requirement-sets"></a>Excel JavaScript API 要求集

要求集是指各组已命名的 API 成员。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 主机是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)。

Excel 加载项在多个 Office 版本中运行，包括 Office 2016 for Windows 或更高版本、Office for iPad、Office for Mac 和 Office Online。下表列出了 Excel 要求集、支持各个要求集的 Office 主机应用程序，以及这些应用程序的内部版本或版本号。

> [!NOTE]
> 标记为 **Beta** 的所有 API 均不可供最终用户使用。我们发布这些 API 的目的在于，让开发人员可以在测试和开发环境中试用。这并不意味着它们可用于生产/业务关键型文档。
> 
> 对于标记为 **Beta** 的要求集，请使用指定版本（或更高版本）的 Office 软件并使用 CDN 上的 Beta 库：https://appsforoffice.microsoft.com/lib/beta/hosted/office.js。未标记为 **Beta** 的条目一般可用，并且可以使用 CDN 上的生产库：https://appsforoffice.microsoft.com/lib/1/hosted/office.js。

|  要求集  |  Office 365 for Windows\*  |  Office 365 for iPad  |  Office 365 for Mac  | Office Online  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|
| Beta  | 请访问 [Excel JavaScript API 开放性规范页面](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)！ |
| ExcelApi1.8  | 版本 1808（内部版本 10730.20102）或更高版本 | 2.17 或更高版本 | 16.17 或更高版本 | 2018 年 9 月 | 即将推出 |
| ExcelApi1.7  | 版本 1801（内部版本 9001.2171）或更高版本   | 2.9 或更高版本 | 16.9 或更高版本 | 2018 年 4 月 | 即将推出 |
| ExcelApi1.6  | 版本 1704（生成号 8201.2001）或更高版本   | 2.2 或更高版本 |15.36 或更高版本| 2017 年 4 月 | 即将推出|
| ExcelApi1.5  | 版本 1703（内部版本 8067.2070）或更高版本   | 2.2 或更高版本 |15.36 或更高版本| 2017 年 3 月 | 即将推出|
| ExcelApi1.4  | 版本 1701（内部版本 7870.2024）或更高版本   | 2.2 或更高版本 |15.36 或更高版本| 2017 年 1 月 | 即将推出|
| ExcelApi1.3  | 版本 1608（内部版本 7369.2055）或更高版本 | 1.27 或更高版本 |  15.27 或更高版本| 2016 年 9 月 | 版本 1608（内部版本 7601.6800）或更高版本|
| ExcelApi1.2  | 版本 1601（内部版本 6741.2088）或更高版本 | 1.21 或更高版本 | 15.22 或更高版本| 2016 年 1 月 ||
| ExcelApi1.1  | 版本 1509（内部版本 4266.1001）或更高版本 | 1.19 或更高版本 | 15.20 或更高版本| 2016 年 1 月 ||

> [!NOTE]
> 通过 MSI 安装的 Office 2016 的内部版本号为 16.0.4266.1001。 此版本只包含 ExcelApi 1.1 要求集。

有关版本、内部版本号和 Office Online Server 的详细信息，请参阅：

- [更新频道发布的 Office 365 客户端版本号和内部版本号](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [使用的是哪一版 Office？](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [在哪里可以找到 Office 365 客户端应用程序的版本号和内部版本号](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Office Online Server 概述](https://docs.microsoft.com/officeonlineserver/office-online-server-overview)

## <a name="whats-new-in-excel-javascript-api-18"></a>Excel JavaScript API 1.8 的最近更新

Excel JavaScript API 要求集 1.8 的功能包括适用于数据透视表、数据验证、图表、图表事件、性能选项和工作簿创建的 API。

### <a name="pivottable"></a>数据透视表

加载项通过数据透视表 API 的波形 2 设置数据透视表的层次结构。 现在可以控制数据及其聚合方式。 [数据透视表](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-pivottables)一文详细介绍了新的数据透视表功能。

### <a name="data-validation"></a>数据有效性

数据有效性可以控制用户在工作表中输入的内容。 可以将单元格限制为预定义的答案集，或者在用户输入无效数据时提供弹出警告。 立即详细了解[向区域添加数据有效性](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-data-validation)。

### <a name="charts"></a>图表

另一轮图表 API 可更好地对图表元素进行编程控制。 现在，你对图例、坐标轴、趋势线和绘图区拥有更高的访问权限。

### <a name="events"></a>事件

已为图表添加更多[事件](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-events)。 让加载项处理用于与图表的交互。 此外，你还可以在整个工作簿中[触发事件](https://docs.microsoft.com/office/dev/add-ins/excel/performance#enable-and-disable-events)。


|对象| 最近更新| 说明|要求集|
|:----|:----|:----|:----|
|[application](/javascript/api/excel/excel.application)|_方法_ > [createWorkbook(base64File: string)](/javascript/api/excel/excel.application)|使用可选 base64 编码的 .xlsx 文件创建新的隐藏工作簿。|1.8|
|[basicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|_属性_ > formula1|获取或设置 Formula1，即最小值或取决于运算符的值。|1.8|
|[basicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|_属性_ > formula2|获取或设置 Formula2，即最大值或取决于运算符的值。|1.8|
|[basicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|_关系_ > operator|用于验证数据有效性的运算符。|1.8|
|[chart](/javascript/api/excel/excel.chart)|_属性_ > categoryLabelLevel|返回或设置 ChartCategoryLabelLevel 枚举常量，该常量代表分类标签源自位置的级别。 读/写。|1.8|
|[chart](/javascript/api/excel/excel.chart)|_属性_ > plotVisibleOnly|如果仅绘制可见单元格，则为 True。 如果绘制可见单元格和隐藏单元格，则为 False。 读写。|1.8|
|[chart](/javascript/api/excel/excel.chart)|_属性_ > seriesNameLevel|返回或设置 ChartSeriesNameLevel 枚举常量，该常量代表系列名称源自位置的级别。 读/写。|1.8|
|[chart](/javascript/api/excel/excel.chart)|_属性_ > showDataLabelsOverMaximum|表示当值大于数值轴上的最大值时是否显示数据标签。|1.8|
|[chart](/javascript/api/excel/excel.chart)|_属性_ > style|返回或设置图表的图表样式。 读写。|1.8|
|[chart](/javascript/api/excel/excel.chart)|_关系_ > displayBlanksAs|返回或设置图表上的空白单元格的绘制方式。 读写。|1.8|
|[chart](/javascript/api/excel/excel.chart)|_关系_ > plotArea|表示图表的绘制区域。 只读。|1.8|
|[chart](/javascript/api/excel/excel.chart)|_关系_ > plotBy|返回或设置图表上的列或行用作数据系列的方式。 读写。|1.8|
|[chartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|_属性_ > chartId|获取已启用图表的 ID。|1.8|
|[chartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|_属性_ > type|获取事件的类型。|1.8|
|[chartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|_属性_ > worksheetId|获取其中的图表已启用的工作表的 ID。|1.8|
|[chartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|_属性_ > chartId|获取已添加至工作表的图表的 ID。|1.8|
|[chartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|_属性_ > type|获取事件的类型。|1.8|
|[chartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|_属性_ > worksheetId|获取已在其中添加图表的工作表的 ID。|1.8|
|[chartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|_关系_ > source|获取事件源。|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_ > isBetweenCategories|表示数值轴是否与分类之间的分类轴交叉。|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_ > multiLevel|表示是否为多级轴。|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_ > numberFormat|表示轴刻度线标签的格式代码。|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_ > offset|表示不同标签级别之间的距离以及一级标签和轴线之间的距离。 此值应该是 0 到 1000 之间的整数。|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_ > positionAt|表示两轴交叉的特定轴位置。 应使用 SetPositionAt(double) 方法设置此属性。 只读。|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_ > textOrientation|表示轴刻度线标签的文本方向。 此值应是 -90 到 90 或 180（垂直文本）之间的整数。|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_关系_ > alignment|表示指定轴刻度线标签的对齐方式。|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_关系_ > position|表示两轴交叉的特定轴位置。|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_方法_ > [setPositionAt(value: double)](/javascript/api/excel/excel.chartaxis)|设置两轴交叉的特定轴位置。|1.8|
|[chartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|_关系_ > fill|表示图表填充格式。 只读。|1.8|
|[chartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|_方法_ > [setFormula(formula: string)](/javascript/api/excel/excel.chartaxistitle)|该字符串值表示采用 A1 表示法的图表轴标题的公式。|1.8|
|[chartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|_关系_ > border|表示边框格式，包括颜色、线条样式和粗细。 只读。|1.8|
|[chartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|_关系_ > fill|表示图表填充格式。 只读。|1.8|
|[chartBorder](/javascript/api/excel/excel.chartborder)|_方法_ > [clear()](/javascript/api/excel/excel.chartborder)|清除图表元素的边框格式。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_属性_ > autoText|该布尔值表示数据标签是否根据上下文自动生成相应的文本。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_属性_ > formula|该字符串值表示采用 A1 表示法的图表数据标签的公式。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_属性_ > height|返回图表数据标签的高度，以磅为单位。 只读。 如果图表数据标签不可见，则为 Null。 只读。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_属性_ > left|表示图表数据标签左边缘到图表区域左边缘的距离，以磅为单位。 如果图表数据标签不可见，则为 Null。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_属性_ > numberFormat|该字符串值表示数据标签的格式代码。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_属性_ > text|该字符串表示图表上的数据标签文本。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_属性_ > textOrientation|表示图表数据标签的文本方向。 此值应是 -90 到 90 或 180（垂直文本）之间的整数。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_属性_ > top|表示图表数据标签上边缘到图表区域顶部的距离，以磅为单位。 如果图表数据标签不可见，则为 Null。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_属性_ > width|返回图表数据标签的宽度，以磅为单位。 只读。 如果图表数据标签不可见，则为 Null。 只读。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_关系_ > format|表示图表数据标签的格式。 只读。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_关系_ > horizontalAlignment|表示图表数据标签水平对齐。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_关系_ > verticalAlignment|表示图表数据标签垂直对齐。|1.8|
|[chartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|_关系_ > border|表示边框格式，包括颜色、线条样式和粗细。 只读。|1.8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_属性_ > autoText|表示数据标签是否根据上下文自动生成相应的文本。|1.8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_属性_ > numberFormat|表示数据标签的格式代码。|1.8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_属性_ > textOrientation|表示数据标签的文本方向。 此值应是 -90 到 90 或 0 到 180（垂直文本）之间的整数。|1.8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_关系_ > horizontalAlignment|表示图表数据标签水平对齐。|1.8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_关系_ > verticalAlignment|表示图表数据标签垂直对齐。|1.8|
|[chartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|_属性_ > chartId|获取停用图表的 ID。|1.8|
|[chartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|_属性_ > type|获取事件的类型。|1.8|
|[chartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|_属性_ > worksheetId|获取其中的图表已停用的工作表的 ID。|1.8|
|[chartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|_属性_ > chartId|获取已从工作表删除的图表的 ID。|1.8|
|[chartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|_属性_ > type|获取事件的类型。|1.8|
|[chartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|_属性_ > worksheetId|获取已在其中删除图表的工作表的 ID。|1.8|
|[chartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|_关系_ > source|获取事件源。|1.8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_属性_ > height|表示图表图例上的 legendEntry 的高度。 只读。|1.8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_属性_ > index|表示图表图例中的 legendEntry 的索引。 只读。|1.8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_属性_ > left|表示图表 legendEntry 的左侧。 只读。|1.8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_属性_ > top|表示图表 legendEntry 的顶部。 只读。|1.8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_属性_ > width|表示图表图例上的 legendEntry 的宽度。 只读。|1.8|
|[chartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|_关系_ > border|表示边框格式，包括颜色、线条样式和粗细。 只读。|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_属性_ > height|表示 plotArea 的高度值。|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_属性_ > insideHeight|表示 plotArea 的 insideHeight 值。|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_属性_ > insideLeft|表示 plotArea 的 insideLeft 值。|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_属性_ > insideTop|表示 plotArea 的 insideTop 值。|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_属性_ > insideWidth|表示 plotArea 的 insideWidth 值。|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_属性_ > left|表示 plotArea 的 left 值。|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_属性_ > top|表示 plotArea 的 top 值。|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_属性_ > width|表示 plotArea 的宽度值。|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_关系_ > format|表示图表 plotArea 的格式。 只读。|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_关系_ > position|表示 plotArea 的位置。|1.8|
|[chartPlotAreaFormat](/javascript/api/excel/excel.chartplotareaformat)|_关系_ > border|表示图表 plotArea 的边框属性。 只读。|1.8|
|[chartPlotAreaFormat](/javascript/api/excel/excel.chartplotareaformat)|_关系_ > fill|表示对象的填充格式，包括背景格式信息。只读。|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_属性_ > explosion|返回或设置饼图或圆环图切片的分解程度值。 如果未分解（切片尖端位于饼图中心），则返回 0（零）。 读写。|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_属性_ > firstSliceAngle|返回或设置第一个饼图或圆环图切片的角度，以度为单位（从垂直方向起为顺时针）。 只适用于饼图、三维饼图和圆环图。 可以是 0 到 360 之间的值。 读写|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_属性_ > invertIfNegative|如果 Excel 在值为负数时反转项目中的图案，则为 True。 读写。|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_属性_ > overlap|指定条柱的摆放方式。 可以是 -100 到 100 之间的值。 只适用于二维条形图和二维柱形图。 读写。|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_属性_ > secondPlotSize|返回或设置复合饼图或复合条饼图的辅助分区的大小，以占主饼图大小的百分比表示。 可以是 5 到 200 之间的值。 读写。|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_属性_ > varyByCategories|如果 Excel 为每个数据标记分配不同的颜色或图案，则为 True。 图表必须只包含一个系列。 读写。|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_关系_ > axisGroup|返回或设置指定系列的组。读写|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_关系_ > dataLabels|表示系列中所有数据标签的集合。 只读。|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_关系_ > splitType|返回或设置复合饼图或复合条饼图中两个分区的拆分方式。 读写。|1.8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_属性_ > backwardPeriod|表示趋势线向后延伸的周期数。|1.8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_属性_ > forwardPeriod|表示趋势线向前延伸的周期数。|1.8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_属性_ > showEquation|如果图表上显示趋势线公式，则为 True。|1.8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_属性_ > showRSquared|如果图表上显示趋势线的 R 平方值，则为 True。|1.8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_关系_ > label|表示图表趋势线的标签。 只读。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_属性_ > autoText|该布尔值表示趋势线标签是否根据上下文自动生成相应的文本。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_属性_ > formula|该字符串值表示采用 A1 表示法的图表趋势线标签的公式。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_属性_ > height|返回图表趋势线标签的高度，以磅为单位。 只读。 如果图表趋势线标签不可见，则为 Null。 只读。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_属性_ > left|表示图表趋势线标签左边缘到图表区域左边缘的距离，以磅为单位。 如果图表趋势线标签不可见，则为 Null。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_属性_ > numberFormat|该字符串值表示趋势线标签的格式代码。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_属性_ > text|该字符串表示图表上的趋势线标签文本。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_属性_ > textOrientation|表示图表趋势线标签的文本方向。 此值应是 -90 到 90 或 180（垂直文本）之间的整数。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_属性_ > top|表示图表趋势线标签上边缘到图表区域顶部的距离，以磅为单位。 如果图表趋势线标签不可见，则为 Null。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_属性_ > width|返回图表趋势线标签的宽度，以磅为单位。 只读。 如果图表趋势线标签不可见，则为 Null。 只读。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_关系_ > format|表示图表趋势线标签的格式。 只读。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_关系_ > horizontalAlignment|表示图表趋势线标签水平对齐。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_关系_ > verticalAlignment|表示图表趋势线标签垂直对齐。|1.8|
|[chartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|_关系_ > border|表示边框格式，包括颜色、线条样式和粗细。 只读。|1.8|
|[chartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|_关系_ > fill|表示当前图表趋势线标签的填充格式。 只读。|1.8|
|[chartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|_关系_ > font|表示图表趋势线标签的字体属性（字体名称、字体大小、颜色等）。 只读。|1.8|
|[customDataValidation](/javascript/api/excel/excel.customdatavalidation)|_属性_ > formula| 自定义数据验证公式。 这将创建特殊输入规则，例如阻止重复或显示单元格区域的总值。|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_属性_ > id|DataPivotHierarchy ID。 只读。|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_属性_ > name|DataPivotHierarchy 的名称。|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_属性_ > numberFormat|DataPivotHierarchy 的数字格式。|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_属性_ > position|DataPivotHierarchy 的位置。|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_关系_ > field|返回与 DataPivotHierarchy 相关联的 PivotFields。 只读。|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_关系_ > showAs|确定数据是否应显示为特定计算汇总。|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_关系_ > summarizeBy|确定是否显示 DataPivotHierarchy 的所有项。|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_方法_ > [setToDefault()](/javascript/api/excel/excel.datapivothierarchy#settodefault--)|将 DataPivotHierarchy 重置回其默认值。|1.8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_属性_ > items|DataPivotHierarchy 对象的集合。 只读。|1.8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_方法_ > [add(pivotHierarchy: PivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection)|将 PivotHierarchy 添加到当前轴。|1.8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_方法_ > [getCount()](/javascript/api/excel/excel.datapivothierarchycollection)|获取集合中的透视层级结构的数量。|1.8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_方法_ > [getItem(name: string)](/javascript/api/excel/excel.datapivothierarchycollection)|按名称或 ID 获取 DataPivotHierarchy。|1.8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_方法_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.datapivothierarchycollection)|按名称获取 DataPivotHierarchy。 如果 DataPivotHierarchy 不存在，则返回 Null 对象。|1.8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_方法_ > [remove(DataPivotHierarchy: DataPivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection)|从当前轴删除 PivotHierarchy。|1.8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_属性_ > ignoreBlanks|忽略空白：不会对空白单元格执行数据严重，默认为 true。|1.8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_属性_ > valid|表示所有单元格值根据数据有效性规则是否全部有效。 只读。|1.8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_关系_ > errorAlert|用户输入无效数据时，出现错误警报。|1.8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_关系_ > prompt|用户选择某个单元格时进行提示。|1.8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_关系_ > rule|包含不同类型的数据有效性标准的数据有效性规则。|1.8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_关系_ > type|数据有效性类型，有关详细信息，请参阅 [Excel.DataValidationType](/javascript/api/excel/excel.datavalidationtype)。 只读。|1.8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_方法_ > [clear()](/javascript/api/excel/excel.datavalidation)|清除当前区域中的数据有效性。|1.8|
|[dataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|_属性_ > message|表示错误警报消息。|1.8|
|[dataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|_属性_ > showAlert|确定在用户输入无效数据时是否显示错误警报对话框。 默认值为 true。|1.8|
|[dataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|_属性_ > title|表示错误警报对话框标题。|1.8|
|[dataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|_关系_ > style|表示数据有效性警报类型，有关详细信息，请参阅 [Excel.DataValidationAlertStyle](/javascript/api/excel/excel.datavalidationalertstyle)。|1.8|
|[dataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)|_属性_ > message|表示提示消息。|1.8|
|[dataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)|_属性_ > showPrompt|确定在用户选择具有数据有效性的单元格时是否显示提示。|1.8|
|[dataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)|_属性_ > title|表示提示标题。|1.8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_关系_ > custom|自定义数据有效性条件。|1.8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_关系_ > date|日期数据有效性条件。|1.8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_关系_ > decimal|小数数据有效性条件。|1.8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_关系_ > list|列表数据有效性条件。|1.8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_关系_ > textLength|TextLength 数据有效性条件。|1.8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_关系_ > time|时间数据有效性条件。|1.8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_关系_ > wholeNumber|WholeNumber 数据有效性条件。|1.8|
|[dateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|_属性_ > formula1|获取或设置 Formula1，即最小值或取决于运算符的值。|1.8|
|[dateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|_属性_ > formula2|获取或设置 Formula2，即最大值或取决于运算符的值。|1.8|
|[dateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|_关系_ > operator|用于验证数据有效性的运算符。|1.8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_属性_ > enableMultipleFilterItems|确定是否允许多个筛选项。|1.8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_属性_ > id|FilterPivotHierarchy 的 ID。 只读。|1.8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_属性_ > name|FilterPivotHierarchy 的名称。|1.8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_属性_ > position|FilterPivotHierarchy 的位置。|1.8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_关系_ > fields|返回与 FilterPivotHierarchy 相关联的 PivotFields。 只读。|1.8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_方法_ > [setToDefault()](/javascript/api/excel/excel.filterpivothierarchy)|将 FilterPivotHierarchy 重置回其默认值。|1.8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_属性_ > items|FilterPivotHierarchy 对象的集合。 只读。|1.8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_方法_ > [add(pivotHierarchy: PivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection)|将 PivotHierarchy 添加到当前轴。 如果行、列或筛选轴上的其他位置存在层次结构，则会将该层次结构从相应的位置移除。|1.8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_方法_ > [getCount()](/javascript/api/excel/excel.filterpivothierarchycollection)|获取集合中的透视层级结构的数量。|1.8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_方法_ > [getItem(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection)|按名称或 ID 获取 FilterPivotHierarchy。|1.8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_方法_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection)|按名称获取 FilterPivotHierarchy。 如果 FilterPivotHierarchy 不存在，则返回 Null 对象。|1.8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_方法_ > [remove(filterPivotHierarchy: FilterPivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection)|从当前轴删除 PivotHierarchy。|1.8|
|[listDataValidation](/javascript/api/excel/excel.listdatavalidation)|_属性_ > inCellDropDown|是否显示单元格下拉菜单中的列表，默认为 true。|1.8|
|[listDataValidation](/javascript/api/excel/excel.listdatavalidation)|_属性_ > source|数据有效性列表源|1.8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_属性_ > id|PivotField 的 ID。 只读。|1.8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_属性_ > name|PivotField 的名称。|1.8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_属性_ > showAllItems|确定是否显示 PivotField 的所有项。|1.8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_关系_ > items|返回与 PivotField 相关联的 PivotFields。 只读。|1.8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_关系_ > subtotals|PivotField 小计。|1.8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_方法_ > [sortByLabels(sortby: SortBy)](/javascript/api/excel/excel.pivotfield)|PivotField 排序。 如果指定 DataPivotHierarchy，则会基于它进行排序，如果未指定，则会基于 PivotField 本身进行排序。|1.8|
|[pivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|_属性_ > items|pivotField 对象的集合。 只读。|1.8|
|[pivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|_方法_ > [getCount()](/javascript/api/excel/excel.pivotfieldcollection)|获取集合中的透视层级结构的数量。|1.8|
|[pivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|_方法_ > [getItem(name: string)](/javascript/api/excel/excel.pivotfieldcollection)|按名称或 ID 获取 PivotHierarchy。|1.8|
|[pivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|_方法_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotfieldcollection)|按名称获取 PivotHierarchy。 如果 PivotHierarchy 不存在，则返回 Null 对象。|1.8|
|[pivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|_属性_ > id|PivotHierarchy 的 ID。 只读。|1.8|
|[pivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|_属性_ > name|PivotHierarchy 的名称。|1.8|
|[pivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|_关系_ > fields|返回与 PivotHierarchy 相关联的 PivotFields。 只读。|1.8|
|[pivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|_属性_ > items|PivotHierarchy 对象的集合。 只读。|1.8|
|[pivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|_方法_ > [getCount()](/javascript/api/excel/excel.pivothierarchycollection)|获取集合中的透视层级结构的数量。|1.8|
|[pivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|_方法_ > [getItem(name: string)](/javascript/api/excel/excel.pivothierarchycollection)|按名称或 ID 获取 PivotHierarchy。|1.8|
|[pivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|_方法_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivothierarchycollection)|按名称获取 PivotHierarchy。 如果 PivotHierarchy 不存在，则返回 Null 对象。|1.8|
|[pivotItem](/javascript/api/excel/excel.pivotitem)|_属性_ > id|PivotItem 的 ID。 只读。|1.8|
|[pivotItem](/javascript/api/excel/excel.pivotitem)|_属性_ > isExpanded|确定是展开项以显示子项还是折叠项并隐藏子项。|1.8|
|[pivotItem](/javascript/api/excel/excel.pivotitem)|_属性_ > name|PivotItem 的名称。|1.8|
|[pivotItem](/javascript/api/excel/excel.pivotitem)|_属性_ > visible|确定 PivotItem 是否可见。|1.8|
|[pivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|_属性_ > items|pivotItem 对象的集合。 只读。|1.8|
|[pivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|_方法_ > [getCount()](/javascript/api/excel/excel.pivotitemcollection)|获取集合中的透视层级结构的数量。|1.8|
|[pivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|_方法_ > [getItem(name: string)](/javascript/api/excel/excel.pivotitemcollection)|按名称或 ID 获取 PivotHierarchy。|1.8|
|[pivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|_方法_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotitemcollection)|按名称获取 PivotHierarchy。 如果 PivotHierarchy 不存在，则返回 Null 对象。|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_属性_ > showColumnGrandTotals|如果数据透视表显示列总计，则为 True。|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_属性_ > showRowGrandTotals|如果数据透视表显示行总计，则为 True。|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_属性_ > subtotalLocation|此属性指示数据透视表上的所有字段的 SubtotalLocationType。 如果字段状态不同，则为 null。 可能的值包括：AtTop、AtBottom。|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_关系_ > layoutType|此属性指示数据透视表上的所有字段的 PivotLayoutType。 如果字段状态不同，则为 null。|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_方法_ > [getColumnLabelRange()](/javascript/api/excel/excel.pivotlayout)|返回数据透视表列标签所在位置的区域。|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_方法_ > [getDataBodyRange()](/javascript/api/excel/excel.pivotlayout)|返回数据透视表数据值所在位置的区域。|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_方法_ > [getFilterAxisRange()](/javascript/api/excel/excel.pivotlayout)|返回数据透视表筛选区的区域。|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_方法_ > [getRange()](/javascript/api/excel/excel.pivotlayout)|返回存在数据透视表的区域，不包括筛选区。|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_方法_ > [getRowLabelRange()](/javascript/api/excel/excel.pivotlayout)|返回数据透视表行标签所在位置的区域。|1.8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_关系_ > columnHierarchies|数据透视表的列透视层级结构。 只读。|1.8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_关系_ > dataHierarchies|数据透视表的数据透视层级结构。 只读。|1.8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_关系_ > filterHierarchies|数据透视表的筛选器透视层级结构。 只读。|1.8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_关系_ > hierarchies|数据透视表的透视层级结构。 只读。|1.8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_关系_ > layout|PivotLayout，用于说明数据透视表的布局和可视化结构。 只读。|1.8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_关系_ > rowHierarchies|数据透视表的行透视层级结构。 只读。|1.8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_方法_  >  [delete()](/javascript/api/excel/excel.pivottable)|删除 PivotTable 对象。|1.8|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_方法_ > [add(name: string, source: object, destination: object)](/javascript/api/excel/excel.pivottablecollection)|基于指定的数据源添加数据透视表，并将其插入到目标区域的左上单元格。|1.8|
|[range](/javascript/api/excel/excel.range)|_关系_ > dataValidation|返回数据有效性对象。 只读。|1.8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_属性_ > id|RowColumnPivotHierarchy 的 ID。 只读。|1.8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_属性_ > name|RowColumnPivotHierarchy 的名称。|1.8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_属性_ > position|RowColumnPivotHierarchy 的位置。|1.8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_关系_ > fields|返回与 RowColumnPivotHierarchy 相关联的 PivotFields。 只读。|1.8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_方法_ > [setToDefault()](/javascript/api/excel/excel.rowcolumnpivothierarchy)|将 RowColumnPivotHierarchy 重置回其默认值。|1.8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_属性_ > items|RowColumnPivotHierarchy 对象的集合。 只读。|1.8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_方法_ > [add(pivotHierarchy: PivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|将 PivotHierarchy 添加到当前轴。 行和列上的其他位置是否存在层次结构。|1.8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_方法_ > [getCount()](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|获取集合中的透视层级结构的数量。|1.8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_方法_ > [getItem(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|按名称或 ID 获取 RowColumnPivotHierarchy。|1.8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_方法_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|按名称获取 RowColumnPivotHierarchy。 如果 RowColumnPivotHierarchy 不存在，则返回 Null 对象。|1.8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_方法_ > [remove(rowColumnPivotHierarchy: RowColumnPivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|从当前轴删除 PivotHierarchy。|1.8|
|[runtime](/javascript/api/excel/excel.runtime)|_属性_ > enableEvents|切换当前任务窗格或内容加载项中的 JavaScript 事件。|1.8|
|[showAsRule](/javascript/api/excel/excel.showasrule)|_关系_ > baseField|基于 ShowAs 计算的基础 PivotField，如适用，基于 ShowAsCalculation 类型，否则为 null。|1.8|
|[showAsRule](/javascript/api/excel/excel.showasrule)|_关系_ > baseItem|基于 ShowAs 计算的基础 Item，如适用，基于 ShowAsCalculation 类型，否则为 null。|1.8|
|[showAsRule](/javascript/api/excel/excel.showasrule)|_关系_ > calculation|数据 PivotField 使用的 ShowAs 计算。|1.8|
|[style](/javascript/api/excel/excel.style)|_属性_ > autoIndent|指示将单元格中的文本对齐方式设为相等分布时文本是否会自动缩进。|1.8|
|[style](/javascript/api/excel/excel.style)|_属性_ > textOrientation|此样式中的文本方向。|1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_属性_ > automatic|如果将“Automatic”设为 true，则在设置 Subtotals 时，所有其他值均会被忽略。|1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_属性_ > average| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_属性_ > count| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_属性_> countNumbers| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_属性_ > max| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_属性_ > min| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_属性_ > product| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_属性_ > standardDeviation| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_属性_ > standardDeviationP| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_属性_ > sum| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_属性_ > variance| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_属性_ > varianceP| |1.8|
|[table](/javascript/api/excel/excel.table)|_属性_ > legacyId|返回数字 ID。只读。|1.8|
|[workbook](/javascript/api/excel/excel.workbook)|_属性_ > readOnly|如果在只读模式下打开工作簿，则为 True。 只读。|1.8|
|[workbookCreated](/javascript/api/excel/excel.workbookcreated)|_属性_ > id|返回用于唯一标识 WorkbookCreated 对象的值。 只读。|1.8|
|[workbookCreated](/javascript/api/excel/excel.workbookcreated)|_方法_ > [open()](/javascript/api/excel/excel.workbookcreated)|打开工作表。|1.8|
|[worksheet](/javascript/api/excel/excel.worksheet)|_属性_ > showGridlines|获取或设置工作表的网格线标志。|1.8|
|[worksheet](/javascript/api/excel/excel.worksheet)|_属性_ > showHeadings|获取或设置工作表的标题标志。|1.8|
|[worksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|_属性_ > type|获取事件的类型。|1.8|
|[worksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|_属性_ > worksheetId|获取计算的工作表的 ID。|1.8|

## <a name="whats-new-in-excel-javascript-api-17"></a>Excel JavaScript API 1.7 的最近更新

Excel JavaScript API 要求集 1.7 的功能包括用于图表、事件、工作表、区域、文档属性、已命名项目、保护选项和样式的 API。

### <a name="customize-charts"></a>自定义图表

通过新的图表 API，你可以创建其他图表类型、向图表中添加数据系列、设置图表标题、添加轴标题、添加显示单位、添加采用移动平均值的趋势线、将趋势线更改为线性趋势线等。 下面是一些示例：

* 图表轴 - 获取、设置、格式化和删除图表中的轴单位、标签和标题。
* 图表系列 - 添加、设置和删除图表中的某个系列。  更改系列标记、绘制顺序和大小。
* 图表趋势线 - 添加、获取和格式化图表中的趋势线。
* 图表图例 - 设置图表中的图例字体的格式。
* 图表点 - 设置图表点颜色。
* 图表标题子字符串 - 获取和设置图表的标题子字符串。
* 图表类型 - 用于创建更多图表类型的选项。

### <a name="events"></a>事件

Excel 事件 API 提供了多个事件处理程序，以便加载项能够在发生特定事件时自动运行指定的函数。 可以将函数设计为执行方案所需的任何操作。 有关当前可用的事件列表，请参阅[使用 Excel JavaScript API 处理事件](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-events)。

### <a name="customize-the-appearance-of-worksheets-and-ranges"></a>自定义工作表和区域的外观

使用新的 API 可以通过多种方式自定义工作表的外观：

* 冻结窗格，使特定行或列在你滚动工作表时保持可见。 例如，如果工作表中的第一行包含标题，则可以冻结此行，以便在你向下滚动工作表时列标题保持可见。
* 修改工作表标签颜色。
* 添加工作表标题。


可以通过多种方式自定义区域的外观：

* 设置某个区域的单元格样式，确保该区域内的所有单元格采用一致的格式。 单元格 样式是一组定义的格式特征，例如字体和字号、数字格式、单元格边框和单元格底纹。 使用 Excel 中的任意内置单元格样式，或者使用自己的自定义单元格样式。
* 设置区域的文本方向。
* 添加或修改区域上链接至工作表中的其他位置或外部位置的超链接。

### <a name="manage-document-properties"></a>管理文档属性

使用文档属性 API，你可以访问内置文档属性，并且还可以创建和管理自定义文档属性，以存储工作表的状态和驱动工作流和业务逻辑。

### <a name="copy-worksheets"></a>复制工作表

使用工作表复制 APIs，你可以将一个工作表中的数据和格式复制到相同工作簿中的另一个工作表，从而减少所需的数据传输量。

### <a name="handle-ranges-with-ease"></a>轻松地处理区域

使用各种区域 API，你可以完成诸如获取周围区域、获取大小经过重设的区域之类的任务。 这些 API 可以显著提高诸如区域操作和寻址之类任务的效率。

此外：

* 工作簿和工作表保护选项 - 使用这些 API 可保护工作表和工作簿结构中的数据。
* 更新已命名项目 - 使用此 API 可更新已命名项目。
* 获取活动单元格 - 使用此 API 可获取工作表中的活动单元格。

|对象| 最近更新| 说明|要求集|
|:----|:----|:----|:----|
|[chart](/javascript/api/excel/excel.chart)|_属性_ > chartType|表示图表的类型。 可能的值包括：ColumnClustered、ColumnStacked、ColumnStacked100、BarClustered、BarStacked、BarStacked100、LineStacked、LineStacked100、LineMarkers、LineMarkersStacked、LineMarkersStacked100、PieOfPie 等。|1.7|
|[chart](/javascript/api/excel/excel.chart)|_属性_ > id|图表的唯一 ID。 只读。|1.7|
|[chart](/javascript/api/excel/excel.chart)|_属性_ > showAllFieldButtons|表示是否在数据透视图上显示所有字段按钮。|1.7|
|[chartAreaFormat](/javascript/api/excel/excel.chartareaformat)|_关系_ > border|表示图表区域的边框格式，包括颜色、线条样式和粗细。 只读。|1.7|
|[chartAxes](/javascript/api/excel/excel.chartaxes)|_方法_ > getItem(type: string, group: string)|返回通过类型和组标识的特定轴。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_ > axisBetweenCategories|表示数值轴是否与分类之间的分类轴交叉。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_ > axisGroup|返回或设置指定轴的组。 只读。 可能的值包括：Primary、Secondary。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_ > categoryType|返回或设置分类轴类型。 可能的值包括：Automatic、TextAxis、DateAxis。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_ > crosses|表示两轴交叉处的特定轴。 可能的值包括：Automatic、Maximum、Minimum、Custom。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_ > crossesAt|表示两轴交叉处的特定轴。 只读。 设置此属性应使用 SetCrossesAt(double) 方法。 只读。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_ > customDisplayUnit|标识自定义轴显示单位值。 只读。 要设置此属性，请使用 SetCustomDisplayUnit(double) 方法。 只读。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_ > displayUnit|表示轴显示单位。 可能的值包括：None、Hundreds、Thousands、TenThousands、HundredThousands、Millions、TenMillions、HundredMillions、Billions、Trillions、Custom。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_ > height|表示图表轴的高度，以磅为单位。 如果轴不可见，则为 Null。 只读。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_ > left|表示轴的左边缘到图表区域左侧的距离，以磅为单位。 如果轴不可见，则为 Null。 只读。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_ > logBase|表示使用对数刻度时对数的底数。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_ > reversePlotOrder|表示 Microsoft Excel 是否按照最后一个到第一个的顺序绘制数据点。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_ > scaleType|表示数值轴刻度类型。 可能的值包括：Linear、Logarithmic。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_ > showDisplayUnitLabel|表示轴显示单位标签是否可见。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_ > tickLabelSpacing|表示刻度线标签之间的分类或系列数。 可以是 1 到 31999 的值或空字符串（自动设置）。 返回的值始终为数字。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_ > tickMarkSpacing|表示刻度线之间的分类或系列数。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_ > top|表示轴的上边缘到图表区域顶部的距离，以磅为单位。 如果轴不可见，则为 Null。 只读。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_ > type|表示轴类型。 只读。 可能的值包括：Invalid、Category、Value、Series。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_ > visible|该布尔值表示轴的可见性。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_ > width|表示图表轴的宽度，以磅为单位。 如果轴不可见，则为 Null。 只读。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_关系_ > baseTimeUnit|返回或设置指定分类轴的基本单位。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_关系_ > majorTickMark|表示特定轴的主要刻度线类型。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_关系_ > majorTimeUnitScale|返回或设置当 CategoryType 属性设为 TimeScale 时分类轴的主要单位刻度值。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_关系_ > minorTickMark|表示指定轴的次要刻度线类型。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_关系_ > minorTimeUnitScale|返回或设置当 CategoryType 属性设为 TimeScale 时分类轴的次要单位刻度值。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_关系_ > tickLabelPosition|表示特定轴上的刻度线标签位置。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_方法_ > setCategoryNames(sourceData: Range)|设置指定轴的所有分类名称。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_方法_ > setCrossesAt(value: double)|设置两轴交叉处的特定轴。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_方法_ > setCustomDisplayUnit(value: double)|将轴显示单位设为自定义值。|1.7|
|[chartBorder](/javascript/api/excel/excel.chartborder)|_属性_ > color|表示图表中的边框颜色的 HTML 颜色代码。|1.7|
|[chartBorder](/javascript/api/excel/excel.chartborder)|_属性_ > weight|表示边框的粗细，以磅为单位。|1.7|
|[chartBorder](/javascript/api/excel/excel.chartborder)|_关系_ > lineStyle|表示边框的线条样式。|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_属性_ > position|表示数据标签的位置的 DataLabelPosition 值。可能的值是：None、Center、InsideEnd、InsideBase、OutsideEnd、Left、Right、Top、Bottom、BestFit、Callout。|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_属性_ > separator|该字符串表示用于图表中数据标签的分隔符。|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_属性_ > showBubbleSize|该布尔值表示数据标签气泡大小是否可见。|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_属性_ > showCategoryName|表示数据标签分类名称是否可见的布尔值。|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_属性_ > showLegendKey|该布尔值表示数据标签图例标示是否可见。|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_属性_ > showPercentage|该布尔值表示数据标签百分比是否可见。|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_属性_ > showSeriesName|该布尔值表示数据标签系列名称是否可见。|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_属性_ > showValue|该布尔值表示数据标签值是否可见。|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_属性_ > height|表示图表上的图例高度。|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_属性_ > left|表示图表图例左侧。|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_属性_ > showShadow|表示图表上的图例是否有阴影。|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_属性_ > top|表示图表图例顶部。|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_属性_ > width|表示图表上的图例宽度。|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_关系_ > legendEntries|表示图例中 legendEntries 的集合。 只读。|1.7|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_属性_ > visible|表示图表图例条目可见。|1.7|
|[chartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|_属性_ > items|ChartLegendEntry 对象的集合。 只读。|1.7|
|[chartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|_方法_ > getCount()|返回集合中的 legendEntry 数量。|1.7|
|[chartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|_方法_ > getItemAt(index: number)|返回给定索引处的 legendEntry。|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_属性_ > hasDataLabel|表示数据点是否具有数据标签。 不适用于曲面图。|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_属性_ > markerBackgroundColor|表示数据点的标记背景色的 HTML 颜色代码。 例如， #FF0000 代表红色。|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_属性_ > markerForegroundColor|表示数据点的标记前景色的 HTML 颜色代码。 例如， #FF0000 代表红色。|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_属性_ > markerSize|表示数据点的标记大小。|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_属性_ > markerStyle|表示图表数据点的标记样式。 可能的值包括：Invalid、Automatic、None、Square、Diamond、Triangle、X、Star、Dot、Dash、Circle、Plus、Picture。|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_关系_ > dataLabel|返回图表点的数据标签。 只读。|1.7|
|[chartPointFormat](/javascript/api/excel/excel.chartpointformat)|_关系_ > border|表示图表数据点的边框格式，包括颜色、样式和粗细信息。 只读。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_属性_ > chartType|表示系列的图表类型。 可能的值包括：ColumnClustered、ColumnStacked、ColumnStacked100、BarClustered、BarStacked、BarStacked100、LineStacked、LineStacked100、LineMarkers、LineMarkersStacked、LineMarkersStacked100、PieOfPie 等。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_属性_ > doughnutHoleSize|表示图表系列的圆环孔大小。  仅对圆环图和分离型圆环图有效。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_属性_ > filtered|该布尔值表示是否筛选系列。 不适用于曲面图。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_属性_ > gapWidth|表示图表系列的间隙宽度。  有效对象：条形图和柱形图，以及|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_属性_ > hasDataLabels|该布尔值表示系列是否具有数据标签。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_属性_ > markerBackgroundColor|表示图表系列的标记背景色。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_属性_ > markerForegroundColor|表示图表系列的标记前景色。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_属性_ > markerSize|表示图表系列的标记大小。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_属性_ > markerStyle|表示图表系列的标记类型。 可能的值包括：Invalid、Automatic、None、Square、Diamond、Triangle、X、Star、Dot、Dash、Circle、Plus、Picture。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_属性_ > plotOrder|表示图表组中某个图表系列的绘制顺序。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_属性_ > showShadow|该布尔值表示系列是否具有阴影。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_属性_ > smooth|该布尔值表示系列是否平滑。 仅适用于折线图和散点图。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_关系_ > dataLabels|表示系列中所有数据标签的集合。 只读。|ApiSet.InProgressFeatures.ChartingAPI|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_关系_ > trendlines|表示系列中趋势线的集合。 只读。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_方法_ > delete()|删除 chart series 对象。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_方法_ > setBubbleSizes(sourceData: Range)|设置图表系列的气泡大小。 仅适用于气泡图。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_方法_ > setValues(sourceData: Range)|设置图表系列的值。 对于散点图，它表示 Y 轴的值。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_方法_ > setXAxisValues(sourceData: Range)|设置图表系列 X 轴的值。 仅适用于散点图。|1.7|
|[chartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|_方法_ > add(name: string, index: number)|向集合添加新系列。|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_属性_ > height|返回图表标题的高度，以磅为单位。 只读。 如果图表标题不可见，则为 Null。 只读。|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_属性_ > horizontalAlignment|表示图表标题水平对齐。 可能的值包括：Center、Left、Justify、Distributed、Right。|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_属性_ > left|表示图表标题左边缘到图表区域左边缘的距离，以磅为单位。 如果图表标题不可见，则为 Null。|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_属性_ > position|表示图表标题的位置。 可能的值包括：Top、Automatic、Bottom、Right、Left。|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_属性_ > showShadow|表示一个布尔值，用于确定图表标题是否具有阴影。|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_属性_ > textOrientation|表示图表标题的文本方向。 此值应是 -90 到 90 或 180（垂直文本）之间的整数。|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_属性_ > top|表示图表标题上边缘到图表区域顶部的距离，以磅为单位。 如果图表标题不可见，则为 Null。|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_属性_ > verticalAlignment|表示图表标题垂直对齐。 可能的值包括：Center、Bottom、Top、Justify、Distributed。|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_属性_ > width|返回图表标题的宽度，以磅为单位。 只读。 如果图表标题不可见，则为 Null。 只读。|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_方法_ > setFormula(formula: string)|设置一个字符串值，用于表示采用 A1 表示法的图表标题的公式。|1.7|
|[chartTitleFormat](/javascript/api/excel/excel.charttitleformat)|_关系_ > border|表示图表标题的边框格式，包括颜色、线条样式和粗细。 只读。|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_属性_ > backward|表示趋势线向后延伸的周期数。|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_属性_ > displayEquation|如果图表上显示趋势线公式，则为 True。|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_属性_ > displayRSquared|如果图表上显示趋势线的 R 平方值，则为 True。|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_属性_ > forward|表示趋势线向前延伸的周期数。|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_属性_ > intercept|表示趋势线的截距值。 可以设置为数字值或空字符串（对于自动值）。 返回的值始终为数字。|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_属性_ > movingAveragePeriod|表示图表趋势线的周期，仅适用于 MovingAverage 类型的趋势线。|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_属性_ > name|表示趋势线的名称。 可设为字符串值，或者设为 Null 值（表示自动值）。 返回的值始终为字符串|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_属性_ > polynomialOrder|表示图表趋势线的顺序，仅适用于 Polynomial 类型的趋势线。|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_属性_ > type|表示图表趋势线的类型。 可能的值包括：Linear、Exponential、Logarithmic、MovingAverage、Polynomial、Power。|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_关系_ > format|表示图表趋势线的格式。 只读。|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_方法_ > delete()|删除 Trendline 对象。|1.7|
|[chartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|_属性_ > items|chartTrendline 对象的集合。 只读。|1.7|
|[chartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|_方法_ > add(type: string)|向趋势线集合添加新的趋势线。|1.7|
|[chartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|_方法_ > getCount()|返回集合中的趋势线数量。|1.7|
|[chartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|_方法_ > getItem(index: number)|按索引（在项目数组中的插入顺序）获取 Trendline 对象。|1.7|
|[chartTrendlineFormat](/javascript/api/excel/excel.charttrendlineformat)|_关系_ > line|表示图表线条格式。只读。|1.7|
|[customProperty](/javascript/api/excel/excel.customproperty)|_属性_ > key|获取 customProperty 的键。只读。只读。|1.7|
|[customProperty](/javascript/api/excel/excel.customproperty)|_属性_ > type|获取自定义属性的值类型。 只读。 只读。 可能的值包括：Number、Boolean、Date、String、Float。|1.7|
|[customProperty](/javascript/api/excel/excel.customproperty)|_属性_ > value|获取或设置自定义属性的值。|1.7|
|[customProperty](/javascript/api/excel/excel.customproperty)|_方法_ > delete()|删除 custom property 对象。|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_属性_ > items|customProperty 对象的集合。只读。|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_方法_ > add(key: string, value: object)|新建自定义属性或设置现有自定义属性。|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_方法_ > deleteAll()|删除此集合中的所有自定义属性。|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_方法_ > getCount()|获取自定义属性的计数。|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_方法_ > getItem(key: string)|按键获取 custom property 对象（不区分大小写）。当不存在自定义属性时引发。|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_方法_ > getItemOrNullObject(key: string)|按键获取 custom property 对象（不区分大小写）。如果不存在自定义属性，则返回 null 对象。|1.7|
|[dataConnectionCollection](/javascript/api/excel/excel.dataconnectioncollection)|_属性_ > items|dataConnection 对象的集合。 只读。|1.7|
|[dataConnectionCollection](/javascript/api/excel/excel.dataconnectioncollection)|_方法_ > refreshAll()|刷新集合中的所有数据连接。|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_属性_ > author|获取或设置工作簿的作者。|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_属性_ > category|获取或设置工作簿的类别。|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_属性_ > comments|获取或设置工作簿的注释。|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_属性_ > company|获取或设置工作簿的公司。|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_属性_ > keywords|获取或设置工作簿的关键字。|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_属性_ > lastAuthor|获取工作簿的最终作者。 只读。 只读。|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_属性_ > manager|获取或设置工作簿的管理者。|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_属性_ > revisionNumber|获取工作簿的修订号。 只读。|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_属性_ > subject|获取或设置工作簿的主题。|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_属性_ > title|获取或设置工作簿的标题。|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_关系_ > creationDate|获取工作簿的创建日期。 只读。 只读。|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_关系_ > custom|获取工作簿的自定义属性的集合。 只读。 只读。|1.7|
|[namedItem](/javascript/api/excel/excel.nameditem)|_属性_ > formula|获取或设置的已命名项目的公式。  公式始终以等号 (=) 开头。|1.7|
|[namedItem](/javascript/api/excel/excel.nameditem)|_关系_ > arrayValues|返回包含已命名项目的值和类型的对象。 只读。|1.7|
|[namedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|_属性_ > types|表示已命名项目数组中每个项目的类型。只读。 可能的值包括：Unknown、Empty、String、Integer、Double、Boolean、Error。|1.7|
|[namedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|_属性_ > values|表示已命名项目数组中每个项目的值。 只读。|1.7|
|[range](/javascript/api/excel/excel.range)|_属性_ > isEntireColumn|表示当前区域是否为整列。 只读。|1.7|
|[range](/javascript/api/excel/excel.range)|_属性_ > isEntireRow|表示当前区域是否为整行。 只读。|1.7|
|[range](/javascript/api/excel/excel.range)|_属性_ > numberFormatLocal|表示 Excel 中的给定区域的数字格式代码，以用户语言的字符串表示。|1.7|
|[range](/javascript/api/excel/excel.range)|_属性_ > style|表示当前区域的样式。 它将返回 null 或字符串。|1.7|
|[range](/javascript/api/excel/excel.range)|_方法_ > getAbsoluteResizedRange(numRows: number, numColumns: number)|获取一个 Range 对象，该对象的左上单元格与当前 Range 对象相同，但具有指定的行数和列数。|1.7|
|[range](/javascript/api/excel/excel.range)|_方法_ > getImage()|将区域呈现为 base64 编码图像。|1.7|
|[range](/javascript/api/excel/excel.range)|_方法_ > getSurroundingRegion()|返回一个 Range 对象，该对象表示此区域左上单元格的周围区域。 周围区域是由相对于该区域的空白行和空白列的任何组合所限定的区域。|1.7|
|[range](/javascript/api/excel/excel.range)|_方法_ > showCard()|显示活动单元格的卡片（如果该单元格具有富值内容）。|1.7|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_属性_ > textOrientation|获取或设置区域内的所有单元格的文本方向。|1.7|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_属性_ > useStandardHeight|确定 Range 对象的行高是否等于工作表的标准行高。|1.7|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_属性_ > useStandardWidth|确定 Range 对象的列宽是否等于工作表的标准列宽。|1.7|
|[rangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|_属性_ > address|表示超链接的 URL 目标。|1.7|
|[rangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|_属性_ > document..|表示超链接的 文档目标。|1.7|
|[rangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|_属性_ > screenTip|表示鼠标悬停在超链接上时显示的字符串。|1.7|
|[rangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|_属性_ > textToDisplay|表示区域最左上方单元格中显示的字符串。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_ > addIndent|指示将单元格中的文本对齐方式设为相等分布时文本是否会自动缩进。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_ > autoIndent|指示将单元格中的文本对齐方式设为相等分布时文本是否会自动缩进。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_ > builtIn|指示样式是否为内置样式。 只读。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_ > formulaHidden|指示工作表受保护时是否隐藏公式。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_ > horizontalAlignment|表示样式水平对齐。 可能的值包括：General、Left、Center、Right、Fill、Justify、CenterAcrossSelection、Distributed。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_ > includeAlignment|指示样式是否包含 AutoIndent、HorizontalAlignment、VerticalAlignment、WrapText、IndentLevel 和 TextOrientation 属性。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_ > includeBorder|指示样式是否包含 Color、ColorIndex、LineStyle 和 Weight 边框属性。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_ > includeFont|指示样式是否包含 Background、Bold、Color、ColorIndex、FontStyle、Italic、Name、Size、Strikethrough、Subscrip、Superscript 和 Underline 字体属性。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_ > includeNumber|指示样式是否包含 NumberFormat 属性。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_ > includePatterns|指示样式是否包含 Color、ColorIndex、InvertIfNegative、Pattern、PatternColor 和 PatternColorIndex 内部属性。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_ > includeProtection|指示样式是否包含 FormulaHidden 和 Locked 保护属性。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_ > indentLevel|0 到 250 之间的一个整数，指示样式的缩进水平。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_ > locked|指示工作表受保护时是否锁定对象。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_ > name|样式的名称。 只读。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_ > numberFormat|样式中数字格式的格式代码。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_ > numberFormatLocal|样式中数字格式的本地化格式代码。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_ > orientation|此样式中的文本方向。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_ > readingOrder|样式中的阅读顺序。 可能的值包括：Context、LeftToRight、RightToLeft。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_ > shrinkToFit|指示文本是否自动缩小以适合可用列宽。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_ > textOrientation|此样式中的文本方向。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_ > verticalAlignment|表示样式的垂直对齐方式。 可能的值包括：Top、Center、Bottom、Justify、Distributed。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_ > wrapText|指示 Microsoft Excel 是否将对象中的文本换行。|1.7|
|[style](/javascript/api/excel/excel.style)|_关系_ > borders|四个 Border 对象的 Border 集合，表示四个边框的样式。 只读。|1.7|
|[style](/javascript/api/excel/excel.style)|_关系_ > fill|样式的填充。 只读。|1.7|
|[style](/javascript/api/excel/excel.style)|_关系_ > font|该 Font 对象表示样式的字体。 只读。|1.7|
|[style](/javascript/api/excel/excel.style)|_方法_ > delete()|删除此样式。|1.7|
|[styleCollection](/javascript/api/excel/excel.stylecollection)|_属性_ > items|style 对象的集合。 只读。|1.7|
|[styleCollection](/javascript/api/excel/excel.stylecollection)|_方法_ > add(name: string)]|向集合添加新样式。|1.7|
|[styleCollection](/javascript/api/excel/excel.stylecollection)|_方法_ > getItem(name: string)|按名称获取样式。|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_属性_ > address|获取地址，该地址表示特定工作表上的表格的更改区域。|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_属性_ > changeType|获取更改类型，该类型表示 Changed 事件的触发方式。 可能的值包括：Others、RangeEdited、RowInserted、RowDeleted、ColumnInserted、ColumnDeleted、CellInserted、CellDeleted。|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_属性_ > source|获取事件源。 可能的值包括：Local、Remote。|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_属性_ > tableId|获取其中的数据发生更改的表格的 ID。|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_属性_ > type|获取事件的类型。 可能的值包括：WorksheetDataChanged、WorksheetSelectionChanged、WorksheetAdded、WorksheetActivated、WorksheetDeactivated、TableDataChanged、TableSelectionChanged、WorksheetDeleted。|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_属性_ > worksheetId|获取其中的数据发生更改的工作表的 ID。|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_属性_ > address|获取区域地址，该地址表示特定工作表上的表格选定区域。|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_属性_ > isInsideTable|指示选定区域是否在表格内，如果 IsInsideTable 为 false，则地址无效。|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_属性_ > tableId|获取其中的选定区域发生更改的表格 ID。|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_属性_ > type|获取事件的类型。 可能的值包括：WorksheetDataChanged、WorksheetSelectionChanged、WorksheetAdded、WorksheetActivated、WorksheetDeactivated、TableDataChanged、TableSelectionChanged、WorksheetDeleted。|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_属性_ > worksheetId|获取其中的选定区域发生更改的工作表的 ID。|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_属性_ > name|获取工作簿名称。 只读。|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_关系_ > dataConnections|刷新工作簿中的所有数据连接。 只读。|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_关系_ > properties|获取工作簿属性。 只读。|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_关系_ > protection|返回工作簿的工作簿保护对象。 只读。|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_关系_ > styles|表示与工作簿关联的样式的集合。 只读。|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_方法_ > getActiveCell()|获取工作簿中当前处于活动状态的单元格。|1.7|
|[workbookProtection](/javascript/api/excel/excel.workbookprotection)|_属性_ > protected|指示工作簿是否受保护。 只读。 只读。|1.7|
|[workbookProtection](/javascript/api/excel/excel.workbookprotection)|_方法_ > protect(password: string)|保护工作簿。 如果工作簿处于受保护状态，则无法执行此方法。|1.7|
|[workbookProtection](/javascript/api/excel/excel.workbookprotection)|_方法_ > unprotect(password: string)|解除保护工作簿。|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_属性_ > gridlines|获取或设置工作表的网格线标志。|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_属性_ > headings|获取或设置工作表的标题标志。|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_属性_ > showHeadings|获取或设置工作表的标题标志。|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_属性_ > standardHeight|返回工作表中所有行的标准（默认）行高，以磅为单位。 只读。|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_属性_ > standardWidth|返回或设置工作表中所有列的标准（默认）列宽。|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_属性_ > tabColor|获取或设置工作表标签颜色。|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_关系_ > freezePanes|获取可用于控制工作表上的冻结窗格的对象。只读。|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_方法_ > copy(positionType: WorksheetPositionType, relativeTo: Worksheet)|复制工作表并将其置于指定位置。 返回复制的工作表。|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_方法_ > getRangeByIndexes(startRow: number, startColumn: number, rowCount: number, columnCount: number)|获取以特定行索引和列索引开始并跨越了一定数量的行和列的 range 对象。|1.7|
|[worksheetActivatedEventArgs](/javascript/api/excel/excel.worksheetactivatedeventargs)|_属性_ > type|获取事件的类型。 可能的值包括：WorksheetDataChanged、WorksheetSelectionChanged、WorksheetAdded、WorksheetActivated、WorksheetDeactivated、TableDataChanged、TableSelectionChanged、WorksheetDeleted。|1.7|
|[worksheetActivatedEventArgs](/javascript/api/excel/excel.worksheetactivatedeventargs)|_属性_ > worksheetId|获取已启用的工作表的 ID。|1.7|
|[worksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|_属性_ > source|获取事件源。 可能的值包括：Local、Remote。|1.7|
|[worksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|_属性_ > type|获取事件的类型。 可能的值包括：WorksheetDataChanged、WorksheetSelectionChanged、WorksheetAdded、WorksheetActivated、WorksheetDeactivated、TableDataChanged、TableSelectionChanged、WorksheetDeleted。|1.7|
|[worksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|_属性_ > worksheetId|获取已添加至工作簿的工作表的 ID。|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_属性_ > address|获取区域地址，该地址表示特定工作表上的更改区域。|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_属性_ > changeType|获取更改类型，该类型表示 Changed 事件的触发方式。 可能的值包括：Others、RangeEdited、RowInserted、RowDeleted、ColumnInserted、ColumnDeleted、CellInserted、CellDeleted。|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_属性_ > source|获取事件源。 可能的值包括：Local、Remote。|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_属性_ > type|获取事件的类型。 可能的值包括：WorksheetDataChanged、WorksheetSelectionChanged、WorksheetAdded、WorksheetActivated、WorksheetDeactivated、TableDataChanged、TableSelectionChanged、WorksheetDeleted。|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_属性_ > worksheetId|获取其中的数据发生更改的工作表的 ID。|1.7|
|[worksheetDeactivatedEventArgs](/javascript/api/excel/excel.worksheetdeactivatedeventargs)|_属性_ > type|获取事件的类型。 可能的值包括：WorksheetDataChanged、WorksheetSelectionChanged、WorksheetAdded、WorksheetActivated、WorksheetDeactivated、TableDataChanged、TableSelectionChanged、WorksheetDeleted。|1.7|
|[worksheetDeactivatedEventArgs](/javascript/api/excel/excel.worksheetdeactivatedeventargs)|_属性_ > worksheetId|获取已停用的工作表的 ID。|1.7|
|[worksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|_属性_ > source|获取事件源。 可能的值包括：Local、Remote。|1.7|
|[worksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|_属性_ > type|获取事件的类型。 可能的值包括：WorksheetDataChanged、WorksheetSelectionChanged、WorksheetAdded、WorksheetActivated、WorksheetDeactivated、TableDataChanged、TableSelectionChanged、WorksheetDeleted。|1.7|
|[worksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|_属性_ > worksheetId|获取已从工作簿删除的工作表的 ID。|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_方法_ > freezeAt(frozenRange: Range or string)|设置活动工作表视图中的冻结单元格。|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_方法_ > freezeColumns(count: number)|就地冻结工作表的第一列。|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_方法_ > freezeRows(count: number)|就地冻结工作表的顶行。|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_方法_ > getLocation()|获取用于描述活动工作表视图中的冻结单元格的区域。|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_方法_ > getLocationOrNullObject()|获取用于描述活动工作表视图中的冻结单元格的区域。|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_方法_ > unfreeze()|移除工作表中的所有冻结窗格。|1.7|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_属性_ > allowEditObjects|表示允许编辑对象的工作表保护选项。|1.7|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_属性_ > allowEditScenarios|表示允许编辑应用场景的工作表保护选项。|1.7|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_关系_ > selectionMode|表示选择模式的工作表保护选项。|1.7|
|[worksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|_属性_ > address|获取区域地址，该地址表示特定工作表上的选定区域。|1.7|
|[worksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|_属性_ > type|获取事件的类型。 可能的值包括：WorksheetDataChanged、WorksheetSelectionChanged、WorksheetAdded、WorksheetActivated、WorksheetDeactivated、TableDataChanged、TableSelectionChanged、WorksheetDeleted。|1.7|
|[worksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|_属性_ > worksheetId|获取其中的选定区域发生更改的工作表的 ID。|1.7|


## <a name="whats-new-in-excel-javascript-api-16"></a>Excel JavaScript API 1.6 的最近更新 

### <a name="conditional-formatting"></a>条件格式

引入了区域的条件格式。 允许以下条件格式类型：

* 色阶
* 数据栏
* 图标集
* 自定义

此外：

* 返回应用条件格式的区域。 
* 删除条件格式。 
* 提供优先级和 stopifTrue 功能。 
* 获取给定区域内所有条件格式的集合。 
* 清除当前指定区域中处于活动状态的所有条件格式。 

|对象| 最近更新| 说明|要求集|
|:----|:----|:----|:----|
|[application](/javascript/api/excel/excel.application)|_方法_ > suspendApiCalculationUntilNextSync()|在下一次调用“context.sync()”前暂停计算。设置后，开发者负责重新计算工作簿，以确保传播所有依赖项。|1.6|
|[cellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)|_关系_ > format|返回 format 对象，该对象用于封装条件格式字体、填充、边框和其他属性。只读。|1.6|
|[cellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)|_关系_ > rule|表示此条件格式中的 Rule 对象。|1.6|
|[colorScaleConditionalFormat](/javascript/api/excel/excel.colorscaleconditionalformat)|_属性_ > threeColorScale|如果为 true，则色阶有三个点（最小、中点、最大），否则将有两个点（最小、最大）。只读。|1.6|
|[colorScaleConditionalFormat](/javascript/api/excel/excel.colorscaleconditionalformat)|_关系_ > criteria|色阶的条件。使用两点色阶时，中点可选。|1.6|
|[conditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|_属性_ > formula1|如果需要，公式可对条件格式规则进行求值。|1.6|
|[conditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|_属性_ > formula2|如果需要，公式可对条件格式规则进行求值。|1.6|
|[conditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|_属性_ > operator|文本条件格式的运算符。可能的值包括：Invalid、Between、NotBetween、EqualTo、NotEqualTo、GreaterThan、LessThan、GreaterThanOrEqual、LessThanOrEqual。|1.6|
|[conditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|_关系_ > maximum|最大点色阶条件。|1.6|
|[conditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|_关系_ > midpoint|色阶为 3 色阶时的中点色阶条件。|1.6|
|[conditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|_关系_ > minimum|最小点色阶条件。|1.6|
|[conditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|_属性_ > color|色阶颜色的 HTML 颜色代码表示。例如，#FF0000 代表红色。|1.6|
|[conditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|_属性_ > formula|数字、公式或 null（如果类型为 LowestValue）。|1.6|
|[conditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|_属性_ > type|图标条件公式的依据。可能的值包括：Invalid、LowestValue、HighestValue、Number、Percent、Formula、Percentile。|1.6|
|[conditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|_属性_ > borderColor|表示窗体 #RRGGBB（例如“FFA500”）的边框线条颜色或作为已命名的 HTML 颜色（例如“orange”）的 HTML 颜色代码。|1.6|
|[conditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|_属性_ > fillColor|表示窗体 #RRGGBB（例如“FFA500”）的填充颜色或已命名 HTML 颜色（例如“orange”）的 HTML 颜色代码。|1.6|
|[conditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|_属性_ > matchPositiveBorderColor|该布尔值表示负 DataBar 是否与正 DataBar 具有相同边框颜色。|1.6|
|[conditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|_属性_ > matchPositiveFillColor|该布尔值表示负 DataBar 是否与正 DataBar 具有相同填充颜色。|1.6|
|[conditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|_属性_ > borderColor|表示窗体 #RRGGBB（例如“FFA500”）的边框线条颜色或作为已命名的 HTML 颜色（例如“orange”）的 HTML 颜色代码。|1.6|
|[conditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|_属性_ > fillColor|表示窗体 #RRGGBB（例如“FFA500”）的填充颜色或已命名 HTML 颜色（例如“orange”）的 HTML 颜色代码。|1.6|
|[conditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|_属性_ > gradientFill|该布尔值表示 DataBar 是否具有渐变。|1.6|
|[conditionalDataBarRule](/javascript/api/excel/excel.conditionaldatabarrule)|_属性_ > formula|如果需要，公式可对 databar 规则进行求值。|1.6|
|[conditionalDataBarRule](/javascript/api/excel/excel.conditionaldatabarrule)|_属性_ > type|数据栏的规则类型。可能的值包括：LowestValue、HighestValue、Number、Percent、Formula、Percentile、Automatic。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_属性_ > id|当前 ConditionalFormatCollection 内的条件格式的优先级。 只读。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_属性_ > priority|优先级（或索引）位于此条件格式当前存在的条件格式集合内。更改此属性也会|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_属性_ > stopIfTrue|如果满足此条件格式的条件，则不会有任何低优先级格式应在此单元格上生效。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_属性_ > type|条件格式的类型。一次只能设置一种类型。只读。只读。可能的值包括：Custom、DataBar、ColorScale、IconSet。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_关系_ > cellValue|如果当前的条件格式是 CellValue 类型，则返回单元值条件格式属性。只读。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_关系_ > cellValueOrNullObject|如果当前的条件格式是 CellValue 类型，则返回单元值条件格式属性。只读。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_关系_ > colorScale|如果当前的条件格式为 ColorScale 类型，返回 ColorScale 条件格式属性。只读。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_关系_ > colorScaleOrNullObject|如果当前的条件格式为 ColorScale 类型，返回 ColorScale 条件格式属性。只读。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_关系_ > custom|如果当前的条件格式是自定义类型，则返回自定义条件格式属性。只读。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_关系_ > customOrNullObject|如果当前的条件格式是自定义类型，则返回自定义条件格式属性。只读。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_关系_ > dataBar|如果当前的条件格式是数据栏，则返回数据栏属性。只读。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_关系_ > dataBarOrNullObject|如果当前的条件格式是数据栏，则返回数据栏属性。只读。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_关系_ > iconSet|如果当前的条件格式是 IconSet 类型，则返回 IconSet 条件格式属性。只读。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_关系_ > iconSetOrNullObject|如果当前的条件格式是 IconSet 类型，则返回 IconSet 条件格式属性。只读。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_关系_ > preset|返回预设条件的条件格式，如上述 averagebelow averageunique valuescontains blanknonblankerrornoerror 属性。只读。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_关系_ > presetOrNullObject|返回预设条件的条件格式，如上述 averagebelow averageunique valuescontains blanknonblankerrornoerror 属性。只读。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_关系_ > textComparison|如果当前的条件格式是文本类型，则返回特定文本条件格式属性。只读。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_关系_ > textComparisonOrNullObject|如果当前的条件格式是文本类型，则返回特定文本条件格式属性。只读。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_关系_ > topBottom|如果当前的条件格式是 TopBottom 类型，则返回 TopBottom 条件格式属性。只读。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_关系_ > topBottomOrNullObject|如果当前的条件格式是 TopBottom 类型，则返回 TopBottom 条件格式属性。只读。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_方法_ > delete()|删除此条件格式。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_方法_ > getRange()|返回条件格式应用的区域，如果区域不连续，则返回 NULL 对象。只读。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_方法_ > getRangeOrNullObject()|返回条件格式应用的区域，如果区域不连续，则返回 NULL 对象。只读。|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_属性_ > items|ConditionalFormat 对象的集合。只读。|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_方法_ > add(type: string)|向 firsttop 优先级的集合添加新的条件格式。|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_方法_ > clearAll()|清除当前指定区域中处于活动状态的所有条件格式。|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_方法_ > getCount()|返回工作簿中的条件格式数。只读。|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_方法_ > getItem(id: string)|返回给定 ID 的条件格式。|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_方法_ > getItemAt(index: number)|返回给定索引处的条件格式。|1.6|
|[conditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|_属性_ > formula|如果需要，公式可对条件格式规则进行求值。|1.6|
|[conditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|_属性_ > formulaLocal|如果需要，公式可采用用户的语言对条件格式规则进行求值。|1.6|
|[conditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|_属性_ > formulaR1C1|如果需要，公式可采用 R1C1 表示法对条件格式规则进行求值。|1.6|
|[conditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|_属性_ > formula|取决于类型的数字或公式。|1.6|
|[conditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|_属性_ > operator|Icon 条件格式的每个规则类型的 GreaterThan 或 GreaterThanOrEqual。可能的值包括：Invalid、GreaterThan、GreaterThanOrEqual。|1.6|
|[conditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|_关系_ > customIcon|如果与默认 IconSet 不同，返回当前条件的自定义图标，否则将返回 null。|1.6|
|[conditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|_关系_ > type|应基于的图标条件公式。|1.6|
|[conditionalPresetCriteriaRule](/javascript/api/excel/excel.conditionalpresetcriteriarule)|_属性_ > criterion|条件格式的条件。可能的值是：Invalid、Blanks、NonBlanks、Errors、NonErrors、Yesterday、Today、Tomorrow、LastSevenDays、LastWeek、ThisWeek、NextWeek、LastMonth、ThisMonth、NextMonth、AboveAverage、BelowAverage、EqualOrAboveAverage、EqualOrBelowAverage、OneStdDevAboveAverage、OneStdDevBelowAverage、TwoStdDevAboveAverage、TwoStdDevBelowAverage、ThreeStdDevAboveAverage、ThreeStdDevBelowAverage、UniqueValues、DuplicateValues。|1.6|
|[conditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|_属性_ > color|表示窗体 #RRGGBB（例如“FFA500”）的边框线条颜色或作为已命名的 HTML 颜色（例如“orange”）的 HTML 颜色代码。|1.6|
|[conditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|_属性_ > id|表示边框标识符。只读。可能的值是：EdgeTop、EdgeBottom、EdgeLeft、EdgeRight。|1.6|
|[conditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|_属性_ > sideIndex|指示边框的特定边的常量值。只读。可能的值是：EdgeTop、EdgeBottom、EdgeLeft、EdgeRight。|1.6|
|[conditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|_属性_ > style|线条样式的常量之一，指定边框的线条样式。可能的值是：None、Continuous、Dash、DashDot、DashDotDot、Dot、Double、SlantDashDot。|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_属性_ > count|集合中的 border 对象数量。只读。|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_属性_ > items|conditionalRangeBorder 对象的集合。只读。|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_关系_ > bottom|获取只读的顶部边框。|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Relationship_ > left|获取只读的顶部边框。|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_关系_ > right|获取只读的顶部边框。|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_关系_ > top|获取只读的顶部边框。|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_方法_ > getItem(index: string)|使用其名称获取 border 对象|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_方法_ > getItemAt(index: number)|使用其索引获取 border 对象|1.6|
|[conditionalRangeFill](/javascript/api/excel/excel.conditionalrangefill)|_属性_ > color|表示窗体 #RRGGBB（例如“FFA500”）的填充颜色或已命名 HTML 颜色（例如“orange”）的 HTML 颜色代码。|1.6|
|[conditionalRangeFill](/javascript/api/excel/excel.conditionalrangefill)|_方法_ > clear()|重置填充。|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_属性_ > bold|表示字体的加粗状态。|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_属性_ > color|文本颜色的 HTML 颜色代码表示。例如，#FF0000 表示红色。|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_属性_ > italic|表示字体的斜体状态。|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_属性_ > strikethrough|表示字体的删除线状态。|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_属性_ > underline|应用于字体的下划线类型。可能的值是：None、Single、Double。|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_方法_ > clear()|重置字体格式。|1.6|
|[conditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|_属性_ > numberFormat|表示 Excel 中指定范围的数字格式代码。当传递 null 时清除。|1.6|
|[conditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|_关系_ > borders|应用于整个条件格式范围的 border 对象的集合。只读。|1.6|
|[conditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|_关系_ > fill|返回在整个条件格式范围内定义的 fill 对象。只读。|1.6|
|[conditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|_关系_ > font|返回在整个条件格式范围内定义的 font 对象。只读。|1.6|
|[conditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule)|_属性_ > operator|本文条件格式的运算符。可能的值包括：Invalid、Contains、NotContains、BeginsWith、EndsWith。|1.6|
|[conditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule)|_属性_ > text|条件格式的文本值。|1.6|
|[conditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|_属性_ > rank|1 和 1000 之间的数字排名或 1 和 100 之间的百分比排名。|1.6|
|[conditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|_属性_ > type|基于排名第一或排名最后的格式值。可能的值包括：Invalid、TopItems、TopPercent、BottomItems、BottomPercent。|1.6|
|[customConditionalFormat](/javascript/api/excel/excel.customconditionalformat)|_关系_ > format|返回 format 对象，该对象用于封装条件格式字体、填充、边框和其他属性。只读。|1.6|
|[customConditionalFormat](/javascript/api/excel/excel.customconditionalformat)|_关系_ > rule|表示此条件格式中的 Rule 对象。只读。|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_属性_ > axisColor|表示窗体 #RRGGBB（例如 "FFA500"）的轴行颜色或作为已命名的 HTML 颜色（例如 "orange"）的 HTML 颜色代码。|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_属性_ > axisFormat|如何确定 Excel 数据栏的轴的表示形式。可能的值包括：Automatic、None、CellMidPoint。|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_属性_ > barDirection|表示数据栏图形应遵循的方向。可能的值包括：Context、LeftToRight、RightToLeft。|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_属性_ > showDataBarOnly|如果为 true，则对应用数据栏的单元格隐藏值。|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_关系_ > lowerBoundRule|构成数据栏的下限（以及如何计算，如果适用）的规则。|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_关系_ > negativeFormat|Excel 数据栏中轴左侧的所有值的表示形式。只读。。|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_关系_ > positiveFormat|Excel 数据栏中轴右侧的所有值的表示形式。只读。。|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_关系_ > upperBoundRule|构成数据栏的上限（以及如何计算，如果适用）的规则。|1.6|
|[iconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|_属性_ > reverseIconOrder|如果为 true，则反转 IconSet 的图标顺序。注意，如果使用自定义图标，则不能进行设置。|1.6|
|[iconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|_属性_ > showIconOnly|如果为 true，则隐藏值并仅显示图标。|1.6|
|[iconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|_属性_ > style|如果设置，则显示条件格式的 IconSet 选项。可能的值是：Invalid、ThreeArrows、ThreeArrowsGray、ThreeFlags、ThreeTrafficLights1、ThreeTrafficLights2、ThreeSigns、ThreeSymbols、ThreeSymbols2、FourArrows、FourArrowsGray、FourRedToBlack、FourRating、FourTrafficLights、FiveArrows、FiveArrowsGray、FiveRating、FiveQuarters、ThreeStars、ThreeTriangles、FiveBoxes。|1.6|
|[iconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|_关系_ > criteria|规则的 Criteria 和 IconSet 数组，以及条件图标的潜在自定义图标。注意，对于第一个条件，只能修改自定义图标，类型、公式和运算符在设置时将忽略。|1.6|
|[presetCriteriaConditionalFormat](/javascript/api/excel/excel.presetcriteriaconditionalformat)|_关系_ > format|返回 format 对象，该对象用于封装条件格式字体、填充、边框和其他属性。只读。|1.6|
|[presetCriteriaConditionalFormat](/javascript/api/excel/excel.presetcriteriaconditionalformat)|_关系_ > rule|条件格式的规则。|1.6|
|[range](/javascript/api/excel/excel.range)|_关系_ > conditionalFormats|区域交叉的 ConditionalFormats 的集合。只读。|1.6|
|[range](/javascript/api/excel/excel.range)|_方法_ > calculate()|计算工作表上的单元格区域。|1.6|
|[textConditionalFormat](/javascript/api/excel/excel.textconditionalformat)|_关系_ > format|返回 format 对象，该对象用于封装条件格式字体、填充、边框和其他属性。只读。|1.6|
|[textConditionalFormat](/javascript/api/excel/excel.textconditionalformat)|_关系_ > rule|条件格式的规则。|1.6|
|[topBottomConditionalFormat](/javascript/api/excel/excel.topbottomconditionalformat)|_关系_ > format|返回 format 对象，该对象用于封装条件格式字体、填充、边框和其他属性。只读。|1.6|
|[topBottomConditionalFormat](/javascript/api/excel/excel.topbottomconditionalformat)|_关系_ > rule|表示 TopBottom 条件格式的条件。|1.6|
|[workbook](/javascript/api/excel/excel.workbook)|_关系_ > internalTest|仅供内部使用。只读。|1.6|
|[worksheet](/javascript/api/excel/excel.worksheet)|_方法_ > calculate(markAllDirty: bool)|计算工作表上的所有单元格。|1.6|

##  <a name="whats-new-in-excel-javascript-api-15"></a>Excel JavaScript API 1.5 的最近更新

### <a name="custom-xml-part"></a>自定义 XML 部件

* 将自定义 XML 部件集合添加到工作簿对象中。
* 使用 ID 获取自定义 XML 部件
* 获取其命名空间匹配给定命名空间的自定义 XML 部件的新作用域内集合。
* 获取与某个部件关联的 XML 字符串。
* 提供部件的 ID 和命名空间。
* 向工作簿添加新的自定义 XML 部件。
* 设置整个 XML 部件。
* 删除自定义 XML 部件。
* 删除其给定名称来自由 xpath 标识的元素的属性。
* 按 xpath 查询 XML 内容。
* 插入、更新和删除属性。

**参考实现：** 请参阅[此处](https://github.com/mandren/Excel-CustomXMLPart-Demo)，了解说明如何在外接程序中使用自定义 XML 部件的参考实现。

### <a name="others"></a>其他
* `range.getSurroundingRegion()` 返回一个 Range 对象，该对象表示此范围的周围区域。周围区域是由相对于该范围的空白行和空白列的任何组合所限定的范围。
* 对表列执行 `getNextColumn()`、`getPreviousColumn()` 以及 `getLast() 操作。
* 对工作簿执行 `getActiveWorksheet()` 操作。
* 工作簿的 `getRange(address: string)` 关闭。
* `getBoundingRange(ranges: )` 获取包含提供的范围的最小 range 对象。例如，介于 “B2:C5” 和 “D10:E15” 之间的边界范围为 “B2:E15”。
* 对各种集合（例如已命名项目、工作表、表等）执行 `getCount()` 操作以获取集合中的项目数。 `workbook.worksheets.getCount()`
* 对各种集合（如工作表、表列、图标点、范围视图集合）执行 `getFirst()` 和 `getLast()` 以及 get last 操作。
* 对工作表、表列集合执行 `getNext()` 和 `getPrevious()` 操作。
* `getRangeR1C1()` 获取以特定行索引和列索引开始并跨越了一定数量的行和列的 range 对象。

|对象| 最近更新| 说明|要求集|
|:----|:----|:----|:----|
|[customXmlPart](/javascript/api/excel/excel.customxmlpart)|_属性_ > id|自定义 XML 部件的 ID。只读。|1.5|
|[customXmlPart](/javascript/api/excel/excel.customxmlpart)|_属性_ > namespaceUri|自定义 XML 部件的命名空间 URI。只读。|1.5|
|[customXmlPart](/javascript/api/excel/excel.customxmlpart)|_方法_ > delete()|删除自定义 XML 部件。|1.5|
|[customXmlPart](/javascript/api/excel/excel.customxmlpart)|_方法_ > getXml()|获取自定义 XML 部件的完整 XML 内容。|1.5|
|[customXmlPart](/javascript/api/excel/excel.customxmlpart)|_方法_ > setXml(xml: string)|设置自定义 XML 部件的完整 XML 内容。|1.5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_属性_ > items|customXmlPart 对象的集合。只读。|1.5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_方法_ > add(xml: string)|向工作簿添加新的自定义 XML 部件。|1.5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_方法_ > getByNamespace(namespaceUri: string)|获取其命名空间匹配给定命名空间的自定义 XML 部件的新作用域内集合。|1.5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_方法_ > getCount()|获取此集合中 CustomXml 部件的数量。|1.5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_方法_ > getItem(id: string)|获取基于其 ID 的自定义 XML 部件。|1.5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_方法_ > getItemOrNullObject(id: string)|获取基于其 ID 的自定义 XML 部件。|1.5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_属性_ > items|CustomXmlPartScoped 对象的集合。只读。|1.5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_方法_ > getCount()|获取此集合中 CustomXML 部件的数量。|1.5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_方法_ > getItem(id: string)|获取基于其 ID 的自定义 XML 部件。|1.5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_方法_ > getItemOrNullObject(id: string)|获取基于其 ID 的自定义 XML 部件。|1.5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_方法_ > getOnlyItem()|如果集合仅包含一个项，则此方法返回该项。|1.5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_方法_ > getOnlyItemOrNullObject()|如果集合仅包含一个项，则此方法返回该项。|1.5|
|[workbook](/javascript/api/excel/excel.workbook)|_关系_ > customXmlParts|代表此工作簿包含的自定义 XML 部件的集合。只读。|1.5|
|[worksheet](/javascript/api/excel/excel.worksheet)|_方法_ > getNext(visibleOnly: bool)|获取该工作表之后的工作表。如果该工作表后没有工作表，此方法将引发错误。|1.5|
|[worksheet](/javascript/api/excel/excel.worksheet)|_方法_ > getNextOrNullObject(visibleOnly: bool)|获取该工作表之后的工作表。如果该工作表后没有工作表，此方法将返回 null 值。|1.5|
|[worksheet](/javascript/api/excel/excel.worksheet)|_方法_ > getPrevious(visibleOnly: bool)|获取该工作表之前的工作表。如果该工作表前没有工作表，此方法将引发错误。|1.5|
|[worksheet](/javascript/api/excel/excel.worksheet)|_方法_ > getPreviousOrNullObject(visibleOnly: bool)|获取该工作表之前的工作表。如果该工作表前没有工作表，此方法将返回 null 值。|1.5|
|[worksheetCollection](/javascript/api/excel/excel.worksheetcollection)|_方法_ > getFirst(visibleOnly: bool)|获取集合中的第一个工作表。|1.5|
|[worksheetCollection](/javascript/api/excel/excel.worksheetcollection)|_方法_ > getLast(visibleOnly: bool)|获取集合中的最后一个工作表。|1.5|

## <a name="whats-new-in-excel-javascript-api-14"></a>Excel JavaScript API 1.4 的最近更新
下面介绍了要求集 1.4 中 Excel JavaScript API 的新增内容。

### <a name="named-item-add-and-new-properties"></a>添加了已命名项和新属性

新属性：

* `comment`
* `scope`：限定到工作表或工作簿的项。
* `worksheet`：返回已命名项限定到的工作表。

新方法：

* `add(name: string, reference: Range or string, comment: string)`：将新名称添加到给定范围的集合。
* `addFormulaLocal(name: string, formula: string, comment: string)`：使用用户的公式区域设置，将新名称添加到给定范围的集合。

### <a name="settings-api-in-the-excel-namespace"></a>Excel 命名空间中的设置 API

[Setting](/javascript/api/excel/excel.setting) 对象表示文档保留设置的键值对。 `Excel.Setting` 的功能等同于 `Office.Settings`，但使用批处理 API 语法，而不是通用 API 的回调模型。

API 包括通过键获取设置条目的 `getItem()`，以及将指定键值设置对添加到工作簿的 `add()`。

### <a name="others"></a>其他

* 设置表列名称（旧版只允许读取）。
* 将表列添加到表的末尾（旧版只允许添加到除末尾之外的其他任何位置）。
* 一次性向表中添加多行（旧版只允许一次添加 1 行）。
* `range.getColumnsAfter(count: number)` 和 `range.getColumnsBefore(count: number)` 分别用于获取当前 Range 对象的右/左侧的一定数量的列。
* 获取项或 NULL 对象函数：此功能允许使用键获取对象。如果没有对象，返回的对象的 isNullObject 属性为 true。这样一来，开发者可以检查对象是否存在，而无需通过异常处理来进行处理。适用于工作表、已命名项、绑定、图表系列等

    ```javascript
    worksheet.GetItemOrNullObject()
    ```

|对象| 最近更新| 说明|要求集|
|:----|:----|:----|:----|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_方法_ > getCount()|获取集合中的绑定数量。|1.4|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_方法_ > getItemOrNullObject(id: string)|按 ID 获取 Binding 对象。如果没有 Binding 对象，将返回 NULL 对象。|1.4|
|[chartCollection](/javascript/api/excel/excel.chartcollection)|_方法_ > getCount()|返回工作表中的图表数。|1.4|
|[chartCollection](/javascript/api/excel/excel.chartcollection)|_方法_ > getItemOrNullObject(name: string)|使用图表名称获取图表。如果存在多个名称相同的图表，将返回第一个图表。|1.4|
|[chartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|_方法_ > getCount()|返回系列中的图表点数。|1.4|
|[chartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|_方法_ > getCount()|返回集合中的系列数量。|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_属性_ > comment|表示与此名称相关联的注释。|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_属性_ > scope|指明是否将 name 限定到工作簿或特定工作表。只读。可取值为：Equal、Greater、GreaterEqual、Less、LessEqual、NotEqual。|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_关系_ > worksheet|返回已命名项限定到的工作表。如果项改为限定到工作簿，将引发错误。只读。|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_关系_ > worksheetOrNullObject|返回已命名项限定到的工作表。如果项改为限定到工作簿，将返回 NULL 对象。只读。|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_方法_ > delete()|删除给定的名称。|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_方法_ > getRangeOrNullObject()|返回与名称相关联的 Range 对象。如果已命名项的类型不是 Range，将返回 NULL 对象。|1.4|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_方法_ > add(name: string, reference: Range or string, comment: string)|将新名称添加到给定范围的集合。|1.4|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_方法_ > addFormulaLocal(name: string, formula: string, comment: string)|使用用户的公式区域设置，将新名称添加到给定范围的集合。|1.4|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_方法_ > getCount()|获取集合中已命名项的数量。|1.4|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_方法_ > getItemOrNullObject(name: string)|按 NamedItem 对象的名称获取此对象。如果没有 NamedItem 对象，将返回 NULL 对象。|1.4|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_方法_ > getCount()|获取集合中的数据透视表的数量。|1.4|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_方法_ > getItemOrNullObject(name: string)|按 PivotTable 对象的名称获取此对象。如果没有 PivotTable 对象，将返回 NULL 对象。|1.4|
|[range](/javascript/api/excel/excel.range)|_方法_ > getIntersectionOrNullObject(anotherRange: Range or string)|获取表示指定区域的矩形交集的 range 对象。如果找不到任何交集，则此方法返回空对象。|1.4|
|[range](/javascript/api/excel/excel.range)|_方法_ > getUsedRangeOrNullObject(valuesOnly: bool)|返回指定 Range 对象的所用区域。如果区域内没有使用单元格，此函数将返回 NULL 对象。|1.4|
|[rangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|_方法_ > getCount()|获取集合中 RangeView 对象的数量。|1.4|
|[setting](/javascript/api/excel/excel.setting)|_属性_ > key|返回表示 setting 对象的 ID 的键。只读。|1.4|
|[setting](/javascript/api/excel/excel.setting)|_属性_ > value|表示为此设置存储的值。|1.4|
|[setting](/javascript/api/excel/excel.setting)|_方法_ > delete()|删除 Setting 对象。|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_属性_ > items|一组 setting 对象。只读。|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_方法_ > add(key: string, value: (any))|设置指定的 Setting 对象，或将其添加到工作簿中。|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_方法_ > getCount()|获取集合中的 Setting 对象的数量。|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_方法_ > getItem(key: string)|按键获取 Setting 项。|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_方法_ > getItemOrNullObject(key: string)|按键获取 Setting 项。如果没有 Setting 项，将返回 NULL 对象。|1.4|
|[settingsChangedEventArgs](/javascript/api/excel/excel.settingschangedeventargs)|_关系_ > settings|获取表示引发了 SettingsChanged 事件的绑定的 Setting 对象。|1.4|
|[tableCollection](/javascript/api/excel/excel.tablecollection)|_方法_ > getCount()]|获取集合中的表数量。|1.4|
|[tableCollection](/javascript/api/excel/excel.tablecollection)|_方法_ > getItemOrNullObject(key: number or string)|按名称或 ID 获取表。如果没有表，将返回 NULL 对象。|1.4|
|[tableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|_方法_ > getCount()|获取表中的列数。|1.4|
|[tableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|_方法_ > getItemOrNullObject(key: number or string)|按名称或 ID 获取 column 对象。如果没有 column 对象，将返回 NULL 对象。|1.4|
|[tableRowCollection](/javascript/api/excel/excel.tablerowcollection)|_方法_ > getCount()|获取表格中的行数。|1.4|
|[workbook](/javascript/api/excel/excel.workbook)|_关系_ > settings|表示一组与 workbook 相关联的 setting 对象。只读。|1.4|
|[worksheet](/javascript/api/excel/excel.worksheet)|_关系_ > names|一组范围限定到当前工作表的名称。只读。|1.4|
|[worksheet](/javascript/api/excel/excel.worksheet)|_方法_ > getUsedRangeOrNullObject(valuesOnly: bool)|使用的区域是包含分配了值或格式的任意单元格的最小区域。如果整个工作表为空，此函数将返回 NULL 对象。|1.4|
|[worksheetCollection](/javascript/api/excel/excel.worksheetcollection)|_方法_ > getCount(visibleOnly: bool)|获取集合中的工作表数量。|1.4|
|[worksheetCollection](/javascript/api/excel/excel.worksheetcollection)|_方法_ > getItemOrNullObject(key: string)|按 Worksheet 对象的名称或 ID 获取此对象。如果没有 Worksheet 对象，将返回 NULL 对象。|1.4|

## <a name="whats-new-in-excel-javascript-api-13"></a>Excel JavaScript API 1.3 的最近更新

下面介绍了要求集 1.3 中 Excel JavaScript API 的新增内容。

|对象| 最近更新| 说明|要求集|
|:----|:----|:----|:----|
|[binding](/javascript/api/excel/excel.binding)|_方法_ > delete()|删除 binding 对象。|1.3|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_方法_ > add(range: Range or string, bindingType: string, id: string)|将新的 binding 对象添加到特定区域。|1.3|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_方法_ > addFromNamedItem(name: string, bindingType: string, id: string)|根据工作簿中的命名项添加新的 binding 对象。|1.3|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_方法_ > addFromSelection(bindingType: string, id: string)|根据当前选择的内容添加新的 binding 对象。|1.3|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_方法_ > getItemOrNull(id: string)|按 ID 获取 binding 对象。如果 binding 对象不存在，则返回对象的 isNull 属性为 true。|1.3|
|[chartCollection](/javascript/api/excel/excel.chartcollection)|_方法_ > getItemOrNull(name: string)|使用图表名称获取图表。如果存在多个名称相同的图表，将返回第一个图表。|1.3|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_方法_ > getItemOrNull(name: string)|按 nameditem 对象的名称获取此对象。如果 nameditem 对象不存在，则返回对象的 isNull 属性为 true。|1.3|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_属性_ > name|PivotTable 对象的名称。|1.3|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_关系_ > worksheet|包含当前 PivotTable 对象的工作表。只读。|1.3|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_方法_ > refresh()|刷新 PivotTable 对象。|1.3|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_属性_ > items|一组 PivotTable 对象。只读。|1.3|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_方法_ > getItem(name: string)|按名称获取 PivotTable 对象。|1.3|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_方法_ > getItemOrNull(name: string)|按名称获取 PivotTable 对象。如果 PivotTable 对象不存在，则返回对象的 isNull 属性为 true。|1.3|
|[range](/javascript/api/excel/excel.range)|_方法_ > getIntersectionOrNull(anotherRange: Range or string)|获取表示指定区域的矩形交集的 range 对象。如果找不到任何交集，则此方法返回空对象。|1.3|
|[range](/javascript/api/excel/excel.range)|_方法_ > getVisibleView()|表示当前 range 对象的可见行。|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_属性_ > cellAddresses|表示 RangeView 的单元格地址。只读。|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_属性_ > columnCount|返回可见列数。只读。|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_属性_ > formulas|表示采用 A1 表示法的公式。|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_属性_ > formulasLocal|使用用户语言和数字格式区域设置表示采用 A1 表示法的公式。例如，用英语表示的公式 "=SUM(A1, introduced in 1.5)" 在德语中将变为 "=SUMME(A1; 1,5)"。|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_属性_ > formulasR1C1|表示采用 R1C1 表示法的公式。|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_属性_ > index|返回表示 RangeView 的索引的值。只读。|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_属性_ > numberFormat|表示 Excel 中指定单元格的数字格式代码。|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_属性_ > rowCount|返回可见行数。只读。|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_属性_ > text|指定区域的文本值。文本值与单元格宽度无关。在 Excel UI 中替代 # 符号不会影响 API 返回的文本值。只读。|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_属性_ > valueTypes|表示每个单元格的数据类型。只读。可能的值是：Unknown、Empty、String、Integer、Double、Boolean、Error。|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_属性_ > values|表示指定的 RangeView 的原始值。返回的数据可能是字符串、数字，也可能是布尔值。包含错误的单元格将返回错误字符串。|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_关系_ > rows|表示一组与 range 相关联的 RangeView。只读。|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_方法_ > getRange()|获取与当前 RangeView 相关联的父 range。|1.3|
|[rangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|_属性_ > items|一组 rangeView 对象。只读。|1.3|
|[rangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|_方法_ > getItemAt(index: number)|按索引获取 RangeView 行。从零开始编制索引。|1.3|
|[setting](/javascript/api/excel/excel.setting)|_属性_ > key|返回表示 setting 对象的 ID 的键。只读。|1.3|
|[setting](/javascript/api/excel/excel.setting)|_方法_ > delete()|删除 Setting 对象。|1.3|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_属性_ > items|一组 setting 对象。只读。|1.3|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_方法_ > getItem(key: string)|按键获取 Setting 项。|1.3|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_方法_ > getItemOrNull(key: string)|按键获取 setting 项。如果 setting 对象不存在，则返回对象的 isNull 属性为 true。|1.3|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_方法_ > set(key: string, value: string)|设置指定的 Setting 对象，或将其添加到工作簿中。|1.3|
|[settingsChangedEventArgs](/javascript/api/excel/excel.settingschangedeventargs)|_关系_ > settingCollection|获取表示引发了 SettingsChanged 事件的 binding 的 setting 对象。|1.3|
|[table](/javascript/api/excel/excel.table)|_属性_ > highlightFirstColumn|指明第一列是否包含特殊格式。|1.3|
|[table](/javascript/api/excel/excel.table)|_属性_ > highlightLastColumn|指明最后一列是否包含特殊格式。|1.3|
|[table](/javascript/api/excel/excel.table)|_属性_ > showBandedColumns|指明列是否采用镶边格式来以不同的方式突出显示奇数列与偶数列，让表更易于阅读。|1.3|
|[table](/javascript/api/excel/excel.table)|_属性_ > showBandedRows|指明行是否采用镶边格式来以不同的方式突出显示奇数行与偶数行，让表更易于阅读。|1.3|
|[table](/javascript/api/excel/excel.table)|_属性_ > showFilterButton|指明是否在每个列标题的顶部显示筛选器按钮。仅当 table 中包含标题行时，才允许设定此设置。|1.3|
|[tableCollection](/javascript/api/excel/excel.tablecollection)|_方法_ > getItemOrNull(key: number or string)|按名称或 ID 获取 table 对象。如果 table 对象不存在，则返回对象的 isNull 属性为 true。|1.3|
|[tableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|_方法_ > getItemOrNull(key: number or string)|按名称或 ID 获取 column 对象。如果 column 对象不存在，则返回对象的 isNull 属性为 true。|1.3|
|[workbook](/javascript/api/excel/excel.workbook)|_关系_ > pivotTables|表示一组与 workbook 相关联的 PivotTable 对象。只读。|1.3|
|[workbook](/javascript/api/excel/excel.workbook)|_关系_ > settings|表示一组与 workbook 相关联的 setting 对象。只读。|1.3|
|[worksheet](/javascript/api/excel/excel.worksheet)|_关系_ > pivotTables|一组属于 worksheet 的 PivotTable 对象。只读。|1.3|

## <a name="whats-new-in-excel-javascript-api-12"></a>Excel JavaScript API 1.2 的最近更新

下面介绍了要求集 1.2 中 Excel JavaScript API 的新增内容。

|对象| 最近更新| 说明|要求集|
|:----|:----|:----|:----|
|[chart](/javascript/api/excel/excel.chart)|_属性_ > id|根据其在集合中的位置获取图表。只读。|1.2|
|[chart](/javascript/api/excel/excel.chart)|_关系_ > worksheet|包含当前 chart 的 worksheet 对象。只读。|1.2|
|[chart](/javascript/api/excel/excel.chart)|_方法_ > getImage(height: number, width: number, fittingMode: string)|通过缩放图表以适应指定的尺寸，将图表呈现为 base64 编码的图像。|1.2|
|[filter](/javascript/api/excel/excel.filter)|_关系_ > criteria|给定列上当前应用的筛选器。只读。|1.2|
|[filter](/javascript/api/excel/excel.filter)|_方法_ > apply(criteria: FilterCriteria)|在给定列中应用给定的筛选条件。|1.2|
|[filter](/javascript/api/excel/excel.filter)|_方法_ > applyBottomItemsFilter(count: number)|将“Bottom Item”筛选器应用于列，以获取给定数量的元素。|1.2|
|[filter](/javascript/api/excel/excel.filter)|_方法_ > applyBottomPercentFilter(percent: number)]|将“Bottom Percent”筛选器应用于列，以获取给定比例的元素。|1.2|
|[filter](/javascript/api/excel/excel.filter)|_方法_ > applyCellColorFilter(color: string)|将“Cell Color”筛选器应用于列，以获取给定颜色。|1.2|
|[filter](/javascript/api/excel/excel.filter)|_方法_ > applyCustomFilter(criteria1: string, criteria2: string, oper: string)|将“Icon”筛选器应用于列，以获取给定的条件字符串。|1.2|
|[filter](/javascript/api/excel/excel.filter)|_方法_ > applyDynamicFilter(criteria: string)|将“Dynamic”筛选器应用于列。|1.2|
|[filter](/javascript/api/excel/excel.filter)|_方法_ > applyFontColorFilter(color: string)|将“Font Color”筛选器应用于列，以获取给定颜色。|1.2|
|[filter](/javascript/api/excel/excel.filter)|_方法_ > applyIconFilter(icon: Icon)|将“Icon”筛选器应用于列，以获取给定图标。|1.2|
|[filter](/javascript/api/excel/excel.filter)|_方法_ > applyTopItemsFilter(count: number)|将“Top Item”筛选器应用于列，以获取给定数量的元素。|1.2|
|[filter](/javascript/api/excel/excel.filter)|_方法_ > applyTopPercentFilter(percent: number)|将“Top Percent”筛选器应用于列，以获取给定比例的元素。|1.2|
|[filter](/javascript/api/excel/excel.filter)|_方法_ > applyValuesFilter(values: ())|将“Values”筛选器应用于列，获取给定值。|1.2|
|[filter](/javascript/api/excel/excel.filter)|_方法_ > clear()|清除给定列上的 filter。|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_属性_ > color|用于筛选单元格的 HTML 颜色字符串。与“cellColor”和“fontColor”筛选一起使用。|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_属性_ > criterion1|第一个条件用于筛选数据。在“自定义”筛选中用作运算符。|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_属性_ > criterion2|第二个条件用于筛选数据。在“自定义”筛选中仅用作运算符。|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_属性_ > dynamicCriteria|Excel.DynamicFilterCriteria 集中的动态条件将应用于此列。与“动态”筛选一起使用。可能的值是：Unknown、AboveAverage、AllDatesInPeriodApril、AllDatesInPeriodAugust、AllDatesInPeriodDecember、AllDatesInPeriodFebruray、AllDatesInPeriodJanuary、AllDatesInPeriodJuly、AllDatesInPeriodJune、AllDatesInPeriodMarch、AllDatesInPeriodMay、AllDatesInPeriodNovember、AllDatesInPeriodOctober、AllDatesInPeriodQuarter1、AllDatesInPeriodQuarter2、AllDatesInPeriodQuarter3、AllDatesInPeriodQuarter4、AllDatesInPeriodSeptember、BelowAverage、LastMonth、LastQuarter、LastWeek、LastYear、NextMonth、NextQuarter、NextWeek、NextYear、ThisMonth、ThisQuarter、ThisWeek、ThisYear、Today、Tomorrow、YearToDate、Yesterday。|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_属性_ > filterOn|filter 使用的属性，用于确定值是否应一直可见。可取值为：BottomItems、BottomPercent、CellColor、Dynamic、FontColor、Values、TopItems、TopPercent、Icon、Custom。|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_属性_ > operator|使用“自定义”筛选器时，用于组合条件 1 和 2 的运算符。可取值为：And、Or。|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_属性_ > values|一组用于“values”筛选器的值。|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_关系_ > icon|用于筛选单元格的图标。与“图标”筛选一起使用。|1.2|
|[filterDatetime](/javascript/api/excel/excel.filterdatetime)|_属性_ > date|用于筛选数据的采用 ISO8601 格式的日期。|1.2|
|[filterDatetime](/javascript/api/excel/excel.filterdatetime)|_属性_ > specificity|用于保留数据的日期的具体程度。例如，如果日期是 2005-04-02 并且将特殊性设置为“月”，则筛选操作将保留包含 2009 年 4 月日期的所有行。可能的值是：Year、Monday、Day、Hour、Minute、Second。|1.2|
|[formatProtection](/javascript/api/excel/excel.formatprotection)|_属性_ > formulaHidden|表示 Excel 是否隐藏区域中的单元格公式。指示整个区域不具有统一公式隐藏设置的空值。|1.2|
|[formatProtection](/javascript/api/excel/excel.formatprotection)|_属性_ > locked|指示 Excel 是否锁定对象中的单元格。指示整个区域不具有统一锁定设置的空值。|1.2|
|[icon](/javascript/api/excel/excel.icon)|_属性_ > index|表示 icon 在给定集中的索引。|1.2|
|[icon](/javascript/api/excel/excel.icon)|_属性_ > set|表示图标所属的集合。可能的值是：Invalid、ThreeArrows、ThreeArrowsGray、ThreeFlags、ThreeTrafficLights1、ThreeTrafficLights2、ThreeSigns、ThreeSymbols、ThreeSymbols2、FourArrows、FourArrowsGray、FourRedToBlack、FourRating、FourTrafficLights、FiveArrows、FiveArrowsGray、FiveRating、FiveQuarters、ThreeStars、ThreeTriangles、FiveBoxes。|1.2|
|[range](/javascript/api/excel/excel.range)|_属性_ > columnHidden|表示当前 range 的所有列均已隐藏。|1.2|
|[range](/javascript/api/excel/excel.range)|_属性_ > formulasR1C1|表示采用 R1C1 表示法的公式。|1.2|
|[range](/javascript/api/excel/excel.range)|_属性_ > hidden|表示当前区域中的所有单元格是否隐藏。只读。|1.2|
|[range](/javascript/api/excel/excel.range)|_属性_ > rowHidden|表示当前 range 的所有行均已隐藏。|1.2|
|[range](/javascript/api/excel/excel.range)|_关系_ > sort|表示当前 range 的区域排序。只读。|1.2|
|[range](/javascript/api/excel/excel.range)|_方法_ > merge(across: bool)|在工作表中，将 range 单元格合并到一个区域中。|1.2|
|[range](/javascript/api/excel/excel.range)|_方法_ > unmerge()|将范围单元格取消合并为各个单元格。|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_属性_ > columnWidth|获取或设置区域内的所有列的宽度。如果列宽不统一，则返回 NULL。|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_属性_ > rowHeight|获取或设置区域中所有行的高度。如果行高不统一，则返回 NULL。|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_关系_ > protection|返回某一区域的格式保护对象。只读。|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_方法_ > autofitColumns()|根据列中的当前数据更改当前范围的列宽，以达到最佳宽度。|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_方法_ > autofitRows()|根据列中的当前数据，更改当前范围的行高以达到最佳高度。|1.2|
|[rangeReference](/javascript/api/excel/excel.rangereference)|_属性_ > address|表示当前 range 对象的可见行。|1.2|
|[rangeSort](/javascript/api/excel/excel.rangesort)|_方法_ > apply(fields: SortField, matchCase: bool, hasHeaders: bool, orientation: string, method: string)|执行排序操作。|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_属性_ > ascending|表示是否执行升序排序。|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_属性_ > color|表示按字体或单元格颜色进行排序时，条件的目标颜色。|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_属性_ > dataOption|表示此字段的其他排序选项。可能的值是：Normal、TextAsNumber。|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_属性_ > key|表示条件所在的列（或行，具体取决于排序方向）。表示与第一列（或行）的偏移量。|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_属性_ > sortOn|表示此条件的排序类型。可能的值是：Value、CellColor、FontColor、Icon。|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_关系_ > icon|表示对单元格图标进行排序时，条件的目标图标。|1.2|
|[table](/javascript/api/excel/excel.table)|_关系_ > sort|表示表的排序。只读。|1.2|
|[table](/javascript/api/excel/excel.table)|_关系_ > worksheet|包含当前表的工作表。只读。|1.2|
|[table](/javascript/api/excel/excel.table)|_方法_ > clearFilters()|清除当前表上应用的所有筛选器。|1.2|
|[table](/javascript/api/excel/excel.table)|_方法_ > convertToRange()|将表转换为普通单元格区域。保留所有数据。|1.2|
|[table](/javascript/api/excel/excel.table)|_方法_ > reapplyFilters()|重新应用当前在 table 上应用的所有 filter。|1.2|
|[tableColumn](/javascript/api/excel/excel.tablecolumn)|_关系_ > filter|检索应用于列的筛选器。只读。|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_属性_ > matchCase|表示最后一次对表进行排序时大小写是否有影响。只读。|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_属性_ > method|表示最后一次对表排序所使用的中文字符排序方法。只读。可能的值是：PinYin、StrokeCount。|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_关系_ > fields|表示最后一次对表排序所使用的当前条件。只读。|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_方法_ > apply(fields: SortField, matchCase: bool, method: string)|执行排序操作。|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_方法_ > clear()|清除表上的当前排序。尽管这不能修改表的排序，但它会清除标题按钮的状态。|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_方法_ > reapply()|对 table 重新应用当前的排序参数。|1.2|
|[workbook](/javascript/api/excel/excel.workbook)|_关系_ > functions|表示包含此工作簿的 Excel 应用程序实例。只读。|1.2|
|[worksheet](/javascript/api/excel/excel.worksheet)|_关系_ > protection|返回表工作表的工作表保护对象。只读。|1.2|
|[worksheetProtection](/javascript/api/excel/excel.worksheetprotection)|_属性_ > protected|指明 worksheet 是否受保护。只读。只读。|1.2|
|[worksheetProtection](/javascript/api/excel/excel.worksheetprotection)|_关系_ > options|工作表保护选项。只读。|1.2|
|[worksheetProtection](/javascript/api/excel/excel.worksheetprotection)|_方法_ > protect(options: WorksheetProtectionOptions)|保护 worksheet。如果 worksheet 处于受保护状态，则无法执行此方法。|1.2|
|[worksheetProtection](/javascript/api/excel/excel.worksheetprotection)|_方法_ > unprotect()|解除对 worksheet 的保护。|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_属性_ > allowAutoFilter|表示允许使用自动筛选功能的工作表保护选项。|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_属性_ > allowDeleteColumns|表示允许删除列的工作表保护选项。|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_属性_ > allowDeleteRows|表示允许删除行的工作表保护选项。|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_属性_ > allowFormatCells|表示允许格式化单元格的工作表保护选项。|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_属性_ > allowFormatColumns|表示允许格式化列的工作表保护选项。|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_属性_ > allowFormatRows|表示允许格式化行的工作表保护选项。|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_属性_ > allowInsertColumns|表示允许插入列的工作表保护选项。|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_属性_ > allowInsertHyperlinks|表示允许插入超链接的工作表保护选项。|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_属性_ > allowInsertRows|表示允许插入行的工作表保护选项。|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_属性_ > allowPivotTables|表示允许使用数据透视表功能的工作表保护选项。|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_属性_ > allowSort|表示允许使用排序功能的工作表保护选项。|1.2|

## <a name="excel-javascript-api-11"></a>Excel JavaScript API 1.1

Excel JavaScript API 1.1 是首版 API。有关 API 的详细信息，请参阅 [Excel JavaScript API](/javascript/api/excel) 参考主题。

## <a name="see-also"></a>另请参阅

- [Office 版本和要求集](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [指定 Office 主机和 API 要求](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office 外接程序 XML 清单](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
