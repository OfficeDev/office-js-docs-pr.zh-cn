---
title: Excel JavaScript API 要求集1。1
description: 有关 ExcelApi 1.1 要求集的详细信息
ms.date: 07/11/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 921a67b4242150d767fdac057d21c6fc510d98b3
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/17/2019
ms.locfileid: "35772049"
---
# <a name="excel-javascript-api-requirement-set-11"></a>Excel JavaScript API 要求集1。1

Excel JavaScript API 1.1 是首版 API。 它是 Excel 2016 支持的唯一特定于 Excel 的要求集。

## <a name="api-list"></a>API 列表

| Class | 域 | 说明 |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[计算 (calculationType: "重新计算\| " "Full \| " "FullRebuild")](/javascript/api/excel/excel.application#calculate-calculationtype-)|重新计算 Excel 中当前打开的所有工作簿。|
||[计算 (calculationType: CalculationType)](/javascript/api/excel/excel.application#calculate-calculationtype-)|重新计算 Excel 中当前打开的所有工作簿。|
||[calculationMode](/javascript/api/excel/excel.application#calculationmode)|返回工作簿中使用的计算模式, 如 CalculationMode 中的常量所定义。 可能的值为`Automatic`:, Excel 控制重新计算的位置。`AutomaticExceptTables`, Excel 在其中控制重新计算, 但忽略表中的更改。`Manual`, 在用户请求计算时完成计算。|
||[set (properties: Excel. Application)](/javascript/api/excel/excel.application#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: ApplicationUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.application#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
|[ApplicationData](/javascript/api/excel/excel.applicationdata)|[calculationMode](/javascript/api/excel/excel.applicationdata#calculationmode)|返回工作簿中使用的计算模式, 如 CalculationMode 中的常量所定义。 可能的值为`Automatic`:, Excel 控制重新计算的位置。`AutomaticExceptTables`, Excel 在其中控制重新计算, 但忽略表中的更改。`Manual`, 在用户请求计算时完成计算。|
|[ApplicationLoadOptions](/javascript/api/excel/excel.applicationloadoptions)|[$all](/javascript/api/excel/excel.applicationloadoptions#$all)||
||[calculationMode](/javascript/api/excel/excel.applicationloadoptions#calculationmode)|返回工作簿中使用的计算模式, 如 CalculationMode 中的常量所定义。 可能的值为`Automatic`:, Excel 控制重新计算的位置。`AutomaticExceptTables`, Excel 在其中控制重新计算, 但忽略表中的更改。`Manual`, 在用户请求计算时完成计算。|
|[ApplicationUpdateData](/javascript/api/excel/excel.applicationupdatedata)|[calculationMode](/javascript/api/excel/excel.applicationupdatedata#calculationmode)|返回工作簿中使用的计算模式, 如 CalculationMode 中的常量所定义。 可能的值为`Automatic`:, Excel 控制重新计算的位置。`AutomaticExceptTables`, Excel 在其中控制重新计算, 但忽略表中的更改。`Manual`, 在用户请求计算时完成计算。|
|[Binding](/javascript/api/excel/excel.binding)|[getRange()](/javascript/api/excel/excel.binding#getrange--)|返回绑定表示的区域。如果绑定类型不正确，将引发错误。|
||[getTable()](/javascript/api/excel/excel.binding#gettable--)|返回绑定表示的表。如果绑定类型不正确，将引发错误。|
||[getText()](/javascript/api/excel/excel.binding#gettext--)|返回绑定表示的文本。 如果绑定类型不正确，将引发错误。|
||[id](/javascript/api/excel/excel.binding#id)|表示绑定标识符。 只读。|
||[type](/javascript/api/excel/excel.binding#type)|返回绑定的类型。 有关详细信息, 请参阅 BindingType。 只读。|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[getItem(id: string)](/javascript/api/excel/excel.bindingcollection#getitem-id-)|按 ID 获取绑定对象。|
||[getItemAt(index: number)](/javascript/api/excel/excel.bindingcollection#getitemat-index-)|根据其在项目数组中的位置获取绑定对象。|
||[count](/javascript/api/excel/excel.bindingcollection#count)|返回集合中绑定的数量。 只读。|
||[items](/javascript/api/excel/excel.bindingcollection#items)|获取此集合中已加载的子项。|
|[BindingCollectionLoadOptions](/javascript/api/excel/excel.bindingcollectionloadoptions)|[$all](/javascript/api/excel/excel.bindingcollectionloadoptions#$all)||
||[id](/javascript/api/excel/excel.bindingcollectionloadoptions#id)|对于集合中的每一项: 表示绑定标识符。 只读。|
||[type](/javascript/api/excel/excel.bindingcollectionloadoptions#type)|对于集合中的每一项: 返回绑定的类型。 有关详细信息, 请参阅 BindingType。 只读。|
|[BindingData](/javascript/api/excel/excel.bindingdata)|[id](/javascript/api/excel/excel.bindingdata#id)|表示绑定标识符。 只读。|
||[type](/javascript/api/excel/excel.bindingdata#type)|返回绑定的类型。 有关详细信息, 请参阅 BindingType。 只读。|
|[BindingLoadOptions](/javascript/api/excel/excel.bindingloadoptions)|[$all](/javascript/api/excel/excel.bindingloadoptions#$all)||
||[id](/javascript/api/excel/excel.bindingloadoptions#id)|表示绑定标识符。 只读。|
||[type](/javascript/api/excel/excel.bindingloadoptions#type)|返回绑定的类型。 有关详细信息, 请参阅 BindingType。 只读。|
|[Chart](/javascript/api/excel/excel.chart)|[delete()](/javascript/api/excel/excel.chart#delete--)|删除 chart 对象。|
||[height](/javascript/api/excel/excel.chart#height)|表示 chart 对象的高度，以磅值表示。|
||[left](/javascript/api/excel/excel.chart#left)|从图表左侧到工作表原点的距离，以磅值表示。|
||[name](/javascript/api/excel/excel.chart#name)|表示 chart 对象的名称。|
||[根](/javascript/api/excel/excel.chart#axes)|表示图表坐标轴。 只读。|
||[dataLabels](/javascript/api/excel/excel.chart#datalabels)|表示图表上的数据标签。 只读。|
||[format](/javascript/api/excel/excel.chart#format)|封装图表区域的格式属性。 只读。|
||[图例](/javascript/api/excel/excel.chart#legend)|表示图表的图例。 只读。|
||[series](/javascript/api/excel/excel.chart#series)|表示单个系列或图表中的系列集合。 只读。|
||[title](/javascript/api/excel/excel.chart#title)|表示指定图表的标题，包括标题的文本、可见性、位置和格式。 只读。|
||[set (properties: Excel. Chart)](/javascript/api/excel/excel.chart#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: ChartUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.chart#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
||[setData (sourceData: Range, seriesBy？: "Auto" \| "Columns \| " "Rows")](/javascript/api/excel/excel.chart#setdata-sourcedata--seriesby-)|重置图表的源数据。|
||[setData (sourceData: Range, seriesBy？: ChartSeriesBy)](/javascript/api/excel/excel.chart#setdata-sourcedata--seriesby-)|重置图表的源数据。|
||[setPosition (startCell: Range \| String, endCell？: range \| string)](/javascript/api/excel/excel.chart#setposition-startcell--endcell-)|相对于工作表上的单元格放置图表。|
||[top](/javascript/api/excel/excel.chart#top)|表示从对象左边界至第 1 行顶部（在工作表上）或图表区域顶部（在图表上）的距离，以磅值表示。|
||[width](/javascript/api/excel/excel.chart#width)|表示 chart 对象的宽度，以磅值表示。|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[fill](/javascript/api/excel/excel.chartareaformat#fill)|表示对象的填充格式，包括背景格式信息。 只读。|
||[font](/javascript/api/excel/excel.chartareaformat#font)|表示当前对象的字体属性（字体名称、字体大小、颜色等）。 只读。|
||[set (properties: ChartAreaFormat)](/javascript/api/excel/excel.chartareaformat#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: ChartAreaFormatUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.chartareaformat#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
|[ChartAreaFormatData](/javascript/api/excel/excel.chartareaformatdata)|[font](/javascript/api/excel/excel.chartareaformatdata#font)|表示当前对象的字体属性（字体名称、字体大小、颜色等）。 只读。|
|[ChartAreaFormatLoadOptions](/javascript/api/excel/excel.chartareaformatloadoptions)|[$all](/javascript/api/excel/excel.chartareaformatloadoptions#$all)||
||[font](/javascript/api/excel/excel.chartareaformatloadoptions#font)|表示当前对象的字体属性（字体名称、字体大小、颜色等）。|
|[ChartAreaFormatUpdateData](/javascript/api/excel/excel.chartareaformatupdatedata)|[font](/javascript/api/excel/excel.chartareaformatupdatedata#font)|表示当前对象的字体属性（字体名称、字体大小、颜色等）。|
|[ChartAxes](/javascript/api/excel/excel.chartaxes)|[categoryAxis](/javascript/api/excel/excel.chartaxes#categoryaxis)|表示图表中的类别轴。 只读。|
||[seriesAxis](/javascript/api/excel/excel.chartaxes#seriesaxis)|表示三维图表的系列轴。 只读。|
||[值坐标轴](/javascript/api/excel/excel.chartaxes#valueaxis)|表示坐标轴中的数值轴。 只读。|
||[set (properties: ChartAxes)](/javascript/api/excel/excel.chartaxes#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: ChartAxesUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.chartaxes#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
|[ChartAxesData](/javascript/api/excel/excel.chartaxesdata)|[categoryAxis](/javascript/api/excel/excel.chartaxesdata#categoryaxis)|表示图表中的类别轴。 只读。|
||[seriesAxis](/javascript/api/excel/excel.chartaxesdata#seriesaxis)|表示三维图表的系列轴。 只读。|
||[值坐标轴](/javascript/api/excel/excel.chartaxesdata#valueaxis)|表示坐标轴中的数值轴。 只读。|
|[ChartAxesLoadOptions](/javascript/api/excel/excel.chartaxesloadoptions)|[$all](/javascript/api/excel/excel.chartaxesloadoptions#$all)||
||[categoryAxis](/javascript/api/excel/excel.chartaxesloadoptions#categoryaxis)|表示图表中的类别轴。|
||[seriesAxis](/javascript/api/excel/excel.chartaxesloadoptions#seriesaxis)|表示三维图表的系列轴。|
||[值坐标轴](/javascript/api/excel/excel.chartaxesloadoptions#valueaxis)|表示坐标轴中的数值轴。|
|[ChartAxesUpdateData](/javascript/api/excel/excel.chartaxesupdatedata)|[categoryAxis](/javascript/api/excel/excel.chartaxesupdatedata#categoryaxis)|表示图表中的类别轴。|
||[seriesAxis](/javascript/api/excel/excel.chartaxesupdatedata#seriesaxis)|表示三维图表的系列轴。|
||[值坐标轴](/javascript/api/excel/excel.chartaxesupdatedata#valueaxis)|表示坐标轴中的数值轴。|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[majorUnit](/javascript/api/excel/excel.chartaxis#majorunit)|表示两个主要刻度标记之间的间隔。可以设置为数字值或空字符串。返回的值始终为数字。|
||[maximum](/javascript/api/excel/excel.chartaxis#maximum)|表示数值轴上的最大值。可以设置为数字值或空字符串（对于自动坐标轴值）。返回的值始终为数字。|
||[minimum](/javascript/api/excel/excel.chartaxis#minimum)|表示数值轴上的最小值。可以设置为数字值或空字符串（对于自动坐标轴值）。返回的值始终为数字。|
||[minorUnit](/javascript/api/excel/excel.chartaxis#minorunit)|表示两个次要刻度标记之间的间隔。 可以设置为数字值或空字符串（对于自动坐标轴值）。 返回的值始终为数字。|
||[format](/javascript/api/excel/excel.chartaxis#format)|表示 chart 对象的格式，包括线条和字体格式。 只读。|
||[majorGridlines](/javascript/api/excel/excel.chartaxis#majorgridlines)|返回一个表示指定坐标轴的主要网格线的网格线对象。 只读。|
||[minorGridlines](/javascript/api/excel/excel.chartaxis#minorgridlines)|返回一个表示指定坐标轴的次要网格线的网格线对象。 只读。|
||[title](/javascript/api/excel/excel.chartaxis#title)|表示坐标轴标题。 只读。|
||[set (properties: ChartAxis)](/javascript/api/excel/excel.chartaxis#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: ChartAxisUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.chartaxis#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
|[ChartAxisData](/javascript/api/excel/excel.chartaxisdata)|[format](/javascript/api/excel/excel.chartaxisdata#format)|表示 chart 对象的格式，包括线条和字体格式。 只读。|
||[majorGridlines](/javascript/api/excel/excel.chartaxisdata#majorgridlines)|返回一个表示指定坐标轴的主要网格线的网格线对象。 只读。|
||[majorUnit](/javascript/api/excel/excel.chartaxisdata#majorunit)|表示两个主要刻度标记之间的间隔。可以设置为数字值或空字符串。返回的值始终为数字。|
||[maximum](/javascript/api/excel/excel.chartaxisdata#maximum)|表示数值轴上的最大值。可以设置为数字值或空字符串（对于自动坐标轴值）。返回的值始终为数字。|
||[minimum](/javascript/api/excel/excel.chartaxisdata#minimum)|表示数值轴上的最小值。可以设置为数字值或空字符串（对于自动坐标轴值）。返回的值始终为数字。|
||[minorGridlines](/javascript/api/excel/excel.chartaxisdata#minorgridlines)|返回一个表示指定坐标轴的次要网格线的网格线对象。 只读。|
||[minorUnit](/javascript/api/excel/excel.chartaxisdata#minorunit)|表示两个次要刻度标记之间的间隔。 可以设置为数字值或空字符串（对于自动坐标轴值）。 返回的值始终为数字。|
||[title](/javascript/api/excel/excel.chartaxisdata#title)|表示坐标轴标题。 只读。|
|[ChartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|[font](/javascript/api/excel/excel.chartaxisformat#font)|表示图表坐标轴元素的字体属性（字体名称、字体大小、颜色等）。 只读。|
||[line](/javascript/api/excel/excel.chartaxisformat#line)|表示图表线条格式。 只读。|
||[set (properties: ChartAxisFormat)](/javascript/api/excel/excel.chartaxisformat#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: ChartAxisFormatUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.chartaxisformat#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
|[ChartAxisFormatData](/javascript/api/excel/excel.chartaxisformatdata)|[font](/javascript/api/excel/excel.chartaxisformatdata#font)|表示图表坐标轴元素的字体属性（字体名称、字体大小、颜色等）。 只读。|
||[line](/javascript/api/excel/excel.chartaxisformatdata#line)|表示图表线条格式。 只读。|
|[ChartAxisFormatLoadOptions](/javascript/api/excel/excel.chartaxisformatloadoptions)|[$all](/javascript/api/excel/excel.chartaxisformatloadoptions#$all)||
||[font](/javascript/api/excel/excel.chartaxisformatloadoptions#font)|表示图表坐标轴元素的字体属性（字体名称、字体大小、颜色等）。|
||[line](/javascript/api/excel/excel.chartaxisformatloadoptions#line)|表示图表线条格式。|
|[ChartAxisFormatUpdateData](/javascript/api/excel/excel.chartaxisformatupdatedata)|[font](/javascript/api/excel/excel.chartaxisformatupdatedata#font)|表示图表坐标轴元素的字体属性（字体名称、字体大小、颜色等）。|
||[line](/javascript/api/excel/excel.chartaxisformatupdatedata#line)|表示图表线条格式。|
|[ChartAxisLoadOptions](/javascript/api/excel/excel.chartaxisloadoptions)|[$all](/javascript/api/excel/excel.chartaxisloadoptions#$all)||
||[format](/javascript/api/excel/excel.chartaxisloadoptions#format)|表示 chart 对象的格式，包括线条和字体格式。|
||[majorGridlines](/javascript/api/excel/excel.chartaxisloadoptions#majorgridlines)|返回一个表示指定坐标轴的主要网格线的网格线对象。|
||[majorUnit](/javascript/api/excel/excel.chartaxisloadoptions#majorunit)|表示两个主要刻度标记之间的间隔。可以设置为数字值或空字符串。返回的值始终为数字。|
||[maximum](/javascript/api/excel/excel.chartaxisloadoptions#maximum)|表示数值轴上的最大值。可以设置为数字值或空字符串（对于自动坐标轴值）。返回的值始终为数字。|
||[minimum](/javascript/api/excel/excel.chartaxisloadoptions#minimum)|表示数值轴上的最小值。可以设置为数字值或空字符串（对于自动坐标轴值）。返回的值始终为数字。|
||[minorGridlines](/javascript/api/excel/excel.chartaxisloadoptions#minorgridlines)|返回一个表示指定坐标轴的次要网格线的网格线对象。|
||[minorUnit](/javascript/api/excel/excel.chartaxisloadoptions#minorunit)|表示两个次要刻度标记之间的间隔。 可以设置为数字值或空字符串（对于自动坐标轴值）。 返回的值始终为数字。|
||[title](/javascript/api/excel/excel.chartaxisloadoptions#title)|表示坐标轴标题。|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[format](/javascript/api/excel/excel.chartaxistitle#format)|表示图表坐标轴标题的格式。 只读。|
||[set (properties: ChartAxisTitle)](/javascript/api/excel/excel.chartaxistitle#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: ChartAxisTitleUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.chartaxistitle#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
||[text](/javascript/api/excel/excel.chartaxistitle#text)|表示坐标轴标题。|
||[visible](/javascript/api/excel/excel.chartaxistitle#visible)|指定坐标轴标题是否可见的布尔值。|
|[ChartAxisTitleData](/javascript/api/excel/excel.chartaxistitledata)|[format](/javascript/api/excel/excel.chartaxistitledata#format)|表示图表坐标轴标题的格式。 只读。|
||[text](/javascript/api/excel/excel.chartaxistitledata#text)|表示坐标轴标题。|
||[visible](/javascript/api/excel/excel.chartaxistitledata#visible)|指定坐标轴标题是否可见的布尔值。|
|[ChartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|[font](/javascript/api/excel/excel.chartaxistitleformat#font)|表示图表坐标轴标题对象的字体属性，例如字体名称、字体大小、颜色等。 只读。|
||[set (properties: ChartAxisTitleFormat)](/javascript/api/excel/excel.chartaxistitleformat#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: ChartAxisTitleFormatUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.chartaxistitleformat#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
|[ChartAxisTitleFormatData](/javascript/api/excel/excel.chartaxistitleformatdata)|[font](/javascript/api/excel/excel.chartaxistitleformatdata#font)|表示图表坐标轴标题对象的字体属性，例如字体名称、字体大小、颜色等。 只读。|
|[ChartAxisTitleFormatLoadOptions](/javascript/api/excel/excel.chartaxistitleformatloadoptions)|[$all](/javascript/api/excel/excel.chartaxistitleformatloadoptions#$all)||
||[font](/javascript/api/excel/excel.chartaxistitleformatloadoptions#font)|表示图表坐标轴标题对象的字体属性，例如字体名称、字体大小、颜色等。|
|[ChartAxisTitleFormatUpdateData](/javascript/api/excel/excel.chartaxistitleformatupdatedata)|[font](/javascript/api/excel/excel.chartaxistitleformatupdatedata#font)|表示图表坐标轴标题对象的字体属性，例如字体名称、字体大小、颜色等。|
|[ChartAxisTitleLoadOptions](/javascript/api/excel/excel.chartaxistitleloadoptions)|[$all](/javascript/api/excel/excel.chartaxistitleloadoptions#$all)||
||[format](/javascript/api/excel/excel.chartaxistitleloadoptions#format)|表示图表坐标轴标题的格式。|
||[text](/javascript/api/excel/excel.chartaxistitleloadoptions#text)|表示坐标轴标题。|
||[visible](/javascript/api/excel/excel.chartaxistitleloadoptions#visible)|指定坐标轴标题是否可见的布尔值。|
|[ChartAxisTitleUpdateData](/javascript/api/excel/excel.chartaxistitleupdatedata)|[format](/javascript/api/excel/excel.chartaxistitleupdatedata#format)|表示图表坐标轴标题的格式。|
||[text](/javascript/api/excel/excel.chartaxistitleupdatedata#text)|表示坐标轴标题。|
||[visible](/javascript/api/excel/excel.chartaxistitleupdatedata#visible)|指定坐标轴标题是否可见的布尔值。|
|[ChartAxisUpdateData](/javascript/api/excel/excel.chartaxisupdatedata)|[format](/javascript/api/excel/excel.chartaxisupdatedata#format)|表示 chart 对象的格式，包括线条和字体格式。|
||[majorGridlines](/javascript/api/excel/excel.chartaxisupdatedata#majorgridlines)|返回一个表示指定坐标轴的主要网格线的网格线对象。|
||[majorUnit](/javascript/api/excel/excel.chartaxisupdatedata#majorunit)|表示两个主要刻度标记之间的间隔。可以设置为数字值或空字符串。返回的值始终为数字。|
||[maximum](/javascript/api/excel/excel.chartaxisupdatedata#maximum)|表示数值轴上的最大值。可以设置为数字值或空字符串（对于自动坐标轴值）。返回的值始终为数字。|
||[minimum](/javascript/api/excel/excel.chartaxisupdatedata#minimum)|表示数值轴上的最小值。可以设置为数字值或空字符串（对于自动坐标轴值）。返回的值始终为数字。|
||[minorGridlines](/javascript/api/excel/excel.chartaxisupdatedata#minorgridlines)|返回一个表示指定坐标轴的次要网格线的网格线对象。|
||[minorUnit](/javascript/api/excel/excel.chartaxisupdatedata#minorunit)|表示两个次要刻度标记之间的间隔。 可以设置为数字值或空字符串（对于自动坐标轴值）。 返回的值始终为数字。|
||[title](/javascript/api/excel/excel.chartaxisupdatedata#title)|表示坐标轴标题。|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[add (type: "无效" \| "ColumnClustered" \| "ColumnStacked" \| "ColumnStacked100" \| "3DColumnClustered" \| "3DColumnStacked" \| "3DColumnStacked100" \| "BarClustered" \| "BarStacked" ""\| "BarStacked100" \| "3DBarClustered" \| "3DBarStacked" \| "3DBarStacked100" \| "LineStacked" \| "LineStacked100" \| "LineMarkers" \| "LineMarkersStacked" \| "" "LineMarkersStacked100 " \| " PieOfPie " \| " PieExploded " \| " 3DPieExploded " \| " BarOfPie " \| " XYScatterSmooth " \| " XYScatterSmoothNoMarkers " \| " XYScatterLines \| ""XYScatterLinesNoMarkers " \| " AreaStacked " \| " AreaStacked100 " \| " 3DAreaStacked " \| " 3DAreaStacked100 " \| " DoughnutExploded " \| " RadarMarkers " \| " RadarFilled \| ""Surface " \| " SurfaceWireframe " \| " SurfaceTopView " \| " SurfaceTopViewWireframe " \| " 气泡图\| "" Bubble3DEffect \| "" StockHLC \| "" StockOHLC \| "" StockVHLC \| "" ""StockVOHLC " \| " CylinderColClustered " \| " CylinderColStacked " \| " CylinderColStacked100 " \| " CylinderBarClustered " \| " CylinderBarStacked " \| " CylinderBarStacked100 " \| "CylinderCol " \| " ConeColClustered " \| " ConeColStacked " \| " ConeColStacked100 " \| " ConeBarClustered " \| " ConeBarStacked " \| " ConeBarStacked100 " \| " ConeCol \| ""PyramidColClustered " \| " PyramidColStacked " \| " PyramidColStacked100 " \| " PyramidBarClustered " \| " PyramidBarStacked " \| " PyramidBarStacked100 " \| " PyramidCol " \| " 3DColumn "" "\| "行" \| "3DLine" \| "3DPie" \| "饼图" \| "XYScatter" \| "3DArea" \| "区域" \| "圆环\|图" " \|雷达" " \|柱状图" " \| Boxwhisker" "图表 " \| " RegionMap " \| " 树状图 " \| " 瀑布式\| "" 旭日\| "" 漏斗图 ", sourceData: Range, seriesBy？:" \| Auto "" \| Columns "" Rows ")](/javascript/api/excel/excel.chartcollection#add-type--sourcedata--seriesby-)|创建新图表。|
||[add (type: ChartType, sourceData: Range, seriesBy？: Excel. ChartSeriesBy)](/javascript/api/excel/excel.chartcollection#add-type--sourcedata--seriesby-)|创建新图表。|
||[getItem(name: string)](/javascript/api/excel/excel.chartcollection#getitem-name-)|使用图表名称获取图表。 如果存在多个名称相同的图表，将返回第一个图表。|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartcollection#getitemat-index-)|根据其在集合中的位置获取图表。|
||[count](/javascript/api/excel/excel.chartcollection#count)|返回工作表中的图表数。 只读。|
||[items](/javascript/api/excel/excel.chartcollection#items)|获取此集合中已加载的子项。|
|[ChartCollectionLoadOptions](/javascript/api/excel/excel.chartcollectionloadoptions)|[$all](/javascript/api/excel/excel.chartcollectionloadoptions#$all)||
||[根](/javascript/api/excel/excel.chartcollectionloadoptions#axes)|对于集合中的每一项: 表示图表坐标轴。|
||[dataLabels](/javascript/api/excel/excel.chartcollectionloadoptions#datalabels)|对于集合中的每一项: 代表图表上的 datalabels。|
||[format](/javascript/api/excel/excel.chartcollectionloadoptions#format)|对于集合中的每一项: 封装图表区域的格式属性。|
||[height](/javascript/api/excel/excel.chartcollectionloadoptions#height)|对于集合中的每一项: 代表图表对象的高度 (以磅为单位)。|
||[left](/javascript/api/excel/excel.chartcollectionloadoptions#left)|对于集合中的每一项: 从图表左侧到工作表原点的距离 (以磅为单位)。|
||[图例](/javascript/api/excel/excel.chartcollectionloadoptions#legend)|对于集合中的每一项: 代表图表的图例。|
||[name](/javascript/api/excel/excel.chartcollectionloadoptions#name)|对于集合中的每一项: 代表 chart 对象的名称。|
||[series](/javascript/api/excel/excel.chartcollectionloadoptions#series)|对于集合中的每一项: 代表图表中的单个系列或系列集合。|
||[title](/javascript/api/excel/excel.chartcollectionloadoptions#title)|对于集合中的每一项: 代表指定图表的标题, 包括标题的文本、可见性、位置和格式。|
||[top](/javascript/api/excel/excel.chartcollectionloadoptions#top)|对于集合中的每一项: 代表从对象的上边缘到工作表第1行顶部的距离, 或图表上的图表区顶部的距离 (以磅为单位)。|
||[width](/javascript/api/excel/excel.chartcollectionloadoptions#width)|对于集合中的每一项: 代表 chart 对象的宽度 (以磅为单位)。|
|[ChartData](/javascript/api/excel/excel.chartdata)|[根](/javascript/api/excel/excel.chartdata#axes)|表示图表坐标轴。 只读。|
||[dataLabels](/javascript/api/excel/excel.chartdata#datalabels)|表示图表上的数据标签。 只读。|
||[format](/javascript/api/excel/excel.chartdata#format)|封装图表区域的格式属性。 只读。|
||[height](/javascript/api/excel/excel.chartdata#height)|表示 chart 对象的高度，以磅值表示。|
||[left](/javascript/api/excel/excel.chartdata#left)|从图表左侧到工作表原点的距离，以磅值表示。|
||[图例](/javascript/api/excel/excel.chartdata#legend)|表示图表的图例。 只读。|
||[name](/javascript/api/excel/excel.chartdata#name)|表示 chart 对象的名称。|
||[series](/javascript/api/excel/excel.chartdata#series)|表示单个系列或图表中的系列集合。 只读。|
||[title](/javascript/api/excel/excel.chartdata#title)|表示指定图表的标题，包括标题的文本、可见性、位置和格式。 只读。|
||[top](/javascript/api/excel/excel.chartdata#top)|表示从对象左边界至第 1 行顶部（在工作表上）或图表区域顶部（在图表上）的距离，以磅值表示。|
||[width](/javascript/api/excel/excel.chartdata#width)|表示 chart 对象的宽度，以磅值表示。|
|[ChartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|[fill](/javascript/api/excel/excel.chartdatalabelformat#fill)|表示当前图表数据标签的填充格式。 只读。|
||[font](/javascript/api/excel/excel.chartdatalabelformat#font)|表示图表数据标签的字体属性（字体名称、字体大小、颜色等）。 只读。|
||[set (properties: ChartDataLabelFormat)](/javascript/api/excel/excel.chartdatalabelformat#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: ChartDataLabelFormatUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.chartdatalabelformat#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
|[ChartDataLabelFormatData](/javascript/api/excel/excel.chartdatalabelformatdata)|[font](/javascript/api/excel/excel.chartdatalabelformatdata#font)|表示图表数据标签的字体属性（字体名称、字体大小、颜色等）。 只读。|
|[ChartDataLabelFormatLoadOptions](/javascript/api/excel/excel.chartdatalabelformatloadoptions)|[$all](/javascript/api/excel/excel.chartdatalabelformatloadoptions#$all)||
||[font](/javascript/api/excel/excel.chartdatalabelformatloadoptions#font)|表示图表数据标签的字体属性（字体名称、字体大小、颜色等）。|
|[ChartDataLabelFormatUpdateData](/javascript/api/excel/excel.chartdatalabelformatupdatedata)|[font](/javascript/api/excel/excel.chartdatalabelformatupdatedata#font)|表示图表数据标签的字体属性（字体名称、字体大小、颜色等）。|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[position](/javascript/api/excel/excel.chartdatalabels#position)|表示数据标签的位置的 DataLabelPosition 值。 有关详细信息, 请参阅 ChartDataLabelPosition。|
||[format](/javascript/api/excel/excel.chartdatalabels#format)|表示图表数据标签的格式，包括填充和字体格式。 只读。|
||[分隔符](/javascript/api/excel/excel.chartdatalabels#separator)|表示用于图表中数据标签的分隔符的字符串。|
||[set (properties: ChartDataLabels)](/javascript/api/excel/excel.chartdatalabels#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: ChartDataLabelsUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.chartdatalabels#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabels#showbubblesize)|该布尔值表示数据标签气泡大小是否可见。|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabels#showcategoryname)|表示数据标签分类名称是否可见的布尔值。|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabels#showlegendkey)|该布尔值表示数据标签图例标示是否可见。|
||[showPercentage](/javascript/api/excel/excel.chartdatalabels#showpercentage)|该布尔值表示数据标签百分比是否可见。|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabels#showseriesname)|该布尔值表示数据标签系列名称是否可见。|
||[showValue](/javascript/api/excel/excel.chartdatalabels#showvalue)|该布尔值表示数据标签值是否可见。|
|[ChartDataLabelsData](/javascript/api/excel/excel.chartdatalabelsdata)|[format](/javascript/api/excel/excel.chartdatalabelsdata#format)|表示图表数据标签的格式，包括填充和字体格式。 只读。|
||[position](/javascript/api/excel/excel.chartdatalabelsdata#position)|表示数据标签的位置的 DataLabelPosition 值。 有关详细信息, 请参阅 ChartDataLabelPosition。|
||[分隔符](/javascript/api/excel/excel.chartdatalabelsdata#separator)|表示用于图表中数据标签的分隔符的字符串。|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabelsdata#showbubblesize)|该布尔值表示数据标签气泡大小是否可见。|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabelsdata#showcategoryname)|表示数据标签分类名称是否可见的布尔值。|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabelsdata#showlegendkey)|该布尔值表示数据标签图例标示是否可见。|
||[showPercentage](/javascript/api/excel/excel.chartdatalabelsdata#showpercentage)|该布尔值表示数据标签百分比是否可见。|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabelsdata#showseriesname)|该布尔值表示数据标签系列名称是否可见。|
||[showValue](/javascript/api/excel/excel.chartdatalabelsdata#showvalue)|该布尔值表示数据标签值是否可见。|
|[ChartDataLabelsLoadOptions](/javascript/api/excel/excel.chartdatalabelsloadoptions)|[$all](/javascript/api/excel/excel.chartdatalabelsloadoptions#$all)||
||[format](/javascript/api/excel/excel.chartdatalabelsloadoptions#format)|表示图表数据标签的格式，包括填充和字体格式。|
||[position](/javascript/api/excel/excel.chartdatalabelsloadoptions#position)|表示数据标签的位置的 DataLabelPosition 值。 有关详细信息, 请参阅 ChartDataLabelPosition。|
||[分隔符](/javascript/api/excel/excel.chartdatalabelsloadoptions#separator)|表示用于图表中数据标签的分隔符的字符串。|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabelsloadoptions#showbubblesize)|该布尔值表示数据标签气泡大小是否可见。|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabelsloadoptions#showcategoryname)|表示数据标签分类名称是否可见的布尔值。|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabelsloadoptions#showlegendkey)|该布尔值表示数据标签图例标示是否可见。|
||[showPercentage](/javascript/api/excel/excel.chartdatalabelsloadoptions#showpercentage)|该布尔值表示数据标签百分比是否可见。|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabelsloadoptions#showseriesname)|该布尔值表示数据标签系列名称是否可见。|
||[showValue](/javascript/api/excel/excel.chartdatalabelsloadoptions#showvalue)|该布尔值表示数据标签值是否可见。|
|[ChartDataLabelsUpdateData](/javascript/api/excel/excel.chartdatalabelsupdatedata)|[format](/javascript/api/excel/excel.chartdatalabelsupdatedata#format)|表示图表数据标签的格式，包括填充和字体格式。|
||[position](/javascript/api/excel/excel.chartdatalabelsupdatedata#position)|表示数据标签的位置的 DataLabelPosition 值。 有关详细信息, 请参阅 ChartDataLabelPosition。|
||[分隔符](/javascript/api/excel/excel.chartdatalabelsupdatedata#separator)|表示用于图表中数据标签的分隔符的字符串。|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabelsupdatedata#showbubblesize)|该布尔值表示数据标签气泡大小是否可见。|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabelsupdatedata#showcategoryname)|表示数据标签分类名称是否可见的布尔值。|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabelsupdatedata#showlegendkey)|该布尔值表示数据标签图例标示是否可见。|
||[showPercentage](/javascript/api/excel/excel.chartdatalabelsupdatedata#showpercentage)|该布尔值表示数据标签百分比是否可见。|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabelsupdatedata#showseriesname)|该布尔值表示数据标签系列名称是否可见。|
||[showValue](/javascript/api/excel/excel.chartdatalabelsupdatedata#showvalue)|该布尔值表示数据标签值是否可见。|
|[ChartFill](/javascript/api/excel/excel.chartfill)|[clear()](/javascript/api/excel/excel.chartfill#clear--)|清除图表元素的填充颜色。|
||[setSolidColor(color: string)](/javascript/api/excel/excel.chartfill#setsolidcolor-color-)|将图表元素的填充格式设置为统一颜色。|
|[ChartFont](/javascript/api/excel/excel.chartfont)|[bold](/javascript/api/excel/excel.chartfont#bold)|表示字体的加粗状态。|
||[color](/javascript/api/excel/excel.chartfont#color)|文本颜色的 HTML 颜色代码表示。 例如， #FF0000 代表红色。|
||[italic](/javascript/api/excel/excel.chartfont#italic)|表示字体的斜体状态。|
||[name](/javascript/api/excel/excel.chartfont#name)|字体名称（例如"Calibri"）|
||[set (properties: ChartFont)](/javascript/api/excel/excel.chartfont#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: ChartFontUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.chartfont#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
||[size](/javascript/api/excel/excel.chartfont#size)|字号（例如，11）|
||[underline](/javascript/api/excel/excel.chartfont#underline)|应用于字体的下划线类型。 有关详细信息, 请参阅 ChartUnderlineStyle。|
|[ChartFontData](/javascript/api/excel/excel.chartfontdata)|[bold](/javascript/api/excel/excel.chartfontdata#bold)|表示字体的加粗状态。|
||[color](/javascript/api/excel/excel.chartfontdata#color)|文本颜色的 HTML 颜色代码表示。 例如， #FF0000 代表红色。|
||[italic](/javascript/api/excel/excel.chartfontdata#italic)|表示字体的斜体状态。|
||[name](/javascript/api/excel/excel.chartfontdata#name)|字体名称（例如"Calibri"）|
||[size](/javascript/api/excel/excel.chartfontdata#size)|字号（例如，11）|
||[underline](/javascript/api/excel/excel.chartfontdata#underline)|应用于字体的下划线类型。 有关详细信息, 请参阅 ChartUnderlineStyle。|
|[ChartFontLoadOptions](/javascript/api/excel/excel.chartfontloadoptions)|[$all](/javascript/api/excel/excel.chartfontloadoptions#$all)||
||[bold](/javascript/api/excel/excel.chartfontloadoptions#bold)|表示字体的加粗状态。|
||[color](/javascript/api/excel/excel.chartfontloadoptions#color)|文本颜色的 HTML 颜色代码表示。 例如， #FF0000 代表红色。|
||[italic](/javascript/api/excel/excel.chartfontloadoptions#italic)|表示字体的斜体状态。|
||[name](/javascript/api/excel/excel.chartfontloadoptions#name)|字体名称（例如"Calibri"）|
||[size](/javascript/api/excel/excel.chartfontloadoptions#size)|字号（例如，11）|
||[underline](/javascript/api/excel/excel.chartfontloadoptions#underline)|应用于字体的下划线类型。 有关详细信息, 请参阅 ChartUnderlineStyle。|
|[ChartFontUpdateData](/javascript/api/excel/excel.chartfontupdatedata)|[bold](/javascript/api/excel/excel.chartfontupdatedata#bold)|表示字体的加粗状态。|
||[color](/javascript/api/excel/excel.chartfontupdatedata#color)|文本颜色的 HTML 颜色代码表示。 例如， #FF0000 代表红色。|
||[italic](/javascript/api/excel/excel.chartfontupdatedata#italic)|表示字体的斜体状态。|
||[name](/javascript/api/excel/excel.chartfontupdatedata#name)|字体名称（例如"Calibri"）|
||[size](/javascript/api/excel/excel.chartfontupdatedata#size)|字号（例如，11）|
||[underline](/javascript/api/excel/excel.chartfontupdatedata#underline)|应用于字体的下划线类型。 有关详细信息, 请参阅 ChartUnderlineStyle。|
|[ChartGridlines](/javascript/api/excel/excel.chartgridlines)|[format](/javascript/api/excel/excel.chartgridlines#format)|表示图表网格线的格式。 只读。|
||[set (properties: ChartGridlines)](/javascript/api/excel/excel.chartgridlines#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: ChartGridlinesUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.chartgridlines#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
||[visible](/javascript/api/excel/excel.chartgridlines#visible)|表示坐标轴网格线是否可见的布尔值。|
|[ChartGridlinesData](/javascript/api/excel/excel.chartgridlinesdata)|[format](/javascript/api/excel/excel.chartgridlinesdata#format)|表示图表网格线的格式。 只读。|
||[visible](/javascript/api/excel/excel.chartgridlinesdata#visible)|表示坐标轴网格线是否可见的布尔值。|
|[ChartGridlinesFormat](/javascript/api/excel/excel.chartgridlinesformat)|[line](/javascript/api/excel/excel.chartgridlinesformat#line)|表示图表线条格式。 只读。|
||[set (properties: ChartGridlinesFormat)](/javascript/api/excel/excel.chartgridlinesformat#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: ChartGridlinesFormatUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.chartgridlinesformat#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
|[ChartGridlinesFormatData](/javascript/api/excel/excel.chartgridlinesformatdata)|[line](/javascript/api/excel/excel.chartgridlinesformatdata#line)|表示图表线条格式。 只读。|
|[ChartGridlinesFormatLoadOptions](/javascript/api/excel/excel.chartgridlinesformatloadoptions)|[$all](/javascript/api/excel/excel.chartgridlinesformatloadoptions#$all)||
||[line](/javascript/api/excel/excel.chartgridlinesformatloadoptions#line)|表示图表线条格式。|
|[ChartGridlinesFormatUpdateData](/javascript/api/excel/excel.chartgridlinesformatupdatedata)|[line](/javascript/api/excel/excel.chartgridlinesformatupdatedata#line)|表示图表线条格式。|
|[ChartGridlinesLoadOptions](/javascript/api/excel/excel.chartgridlinesloadoptions)|[$all](/javascript/api/excel/excel.chartgridlinesloadoptions#$all)||
||[format](/javascript/api/excel/excel.chartgridlinesloadoptions#format)|表示图表网格线的格式。|
||[visible](/javascript/api/excel/excel.chartgridlinesloadoptions#visible)|表示坐标轴网格线是否可见的布尔值。|
|[ChartGridlinesUpdateData](/javascript/api/excel/excel.chartgridlinesupdatedata)|[format](/javascript/api/excel/excel.chartgridlinesupdatedata#format)|表示图表网格线的格式。|
||[visible](/javascript/api/excel/excel.chartgridlinesupdatedata#visible)|表示坐标轴网格线是否可见的布尔值。|
|[ChartLegend](/javascript/api/excel/excel.chartlegend)|[重叠](/javascript/api/excel/excel.chartlegend#overlay)|表示图表图例是否应该与图表的主体重叠的布尔值。|
||[position](/javascript/api/excel/excel.chartlegend#position)|表示图例在图表上的位置。 有关详细信息, 请参阅 ChartLegendPosition。|
||[format](/javascript/api/excel/excel.chartlegend#format)|表示图表图例的格式，包括填充和字体格式。 只读。|
||[set (properties: ChartLegend)](/javascript/api/excel/excel.chartlegend#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: ChartLegendUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.chartlegend#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
||[visible](/javascript/api/excel/excel.chartlegend#visible)|表示 ChartLegend 对象是否可见的布尔值。|
|[ChartLegendData](/javascript/api/excel/excel.chartlegenddata)|[format](/javascript/api/excel/excel.chartlegenddata#format)|表示图表图例的格式，包括填充和字体格式。 只读。|
||[重叠](/javascript/api/excel/excel.chartlegenddata#overlay)|表示图表图例是否应该与图表的主体重叠的布尔值。|
||[position](/javascript/api/excel/excel.chartlegenddata#position)|表示图例在图表上的位置。 有关详细信息, 请参阅 ChartLegendPosition。|
||[visible](/javascript/api/excel/excel.chartlegenddata#visible)|表示 ChartLegend 对象是否可见的布尔值。|
|[ChartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|[fill](/javascript/api/excel/excel.chartlegendformat#fill)|表示对象的填充格式，包括背景格式信息。 只读。|
||[font](/javascript/api/excel/excel.chartlegendformat#font)|表示图表图例的字体属性，例如字体名称、字体大小、颜色等。 只读。|
||[set (properties: ChartLegendFormat)](/javascript/api/excel/excel.chartlegendformat#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: ChartLegendFormatUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.chartlegendformat#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
|[ChartLegendFormatData](/javascript/api/excel/excel.chartlegendformatdata)|[font](/javascript/api/excel/excel.chartlegendformatdata#font)|表示图表图例的字体属性，例如字体名称、字体大小、颜色等。 只读。|
|[ChartLegendFormatLoadOptions](/javascript/api/excel/excel.chartlegendformatloadoptions)|[$all](/javascript/api/excel/excel.chartlegendformatloadoptions#$all)||
||[font](/javascript/api/excel/excel.chartlegendformatloadoptions#font)|表示图表图例的字体属性，例如字体名称、字体大小、颜色等。|
|[ChartLegendFormatUpdateData](/javascript/api/excel/excel.chartlegendformatupdatedata)|[font](/javascript/api/excel/excel.chartlegendformatupdatedata#font)|表示图表图例的字体属性，例如字体名称、字体大小、颜色等。|
|[ChartLegendLoadOptions](/javascript/api/excel/excel.chartlegendloadoptions)|[$all](/javascript/api/excel/excel.chartlegendloadoptions#$all)||
||[format](/javascript/api/excel/excel.chartlegendloadoptions#format)|表示图表图例的格式，包括填充和字体格式。|
||[重叠](/javascript/api/excel/excel.chartlegendloadoptions#overlay)|表示图表图例是否应该与图表的主体重叠的布尔值。|
||[position](/javascript/api/excel/excel.chartlegendloadoptions#position)|表示图例在图表上的位置。 有关详细信息, 请参阅 ChartLegendPosition。|
||[visible](/javascript/api/excel/excel.chartlegendloadoptions#visible)|表示 ChartLegend 对象是否可见的布尔值。|
|[ChartLegendUpdateData](/javascript/api/excel/excel.chartlegendupdatedata)|[format](/javascript/api/excel/excel.chartlegendupdatedata#format)|表示图表图例的格式，包括填充和字体格式。|
||[重叠](/javascript/api/excel/excel.chartlegendupdatedata#overlay)|表示图表图例是否应该与图表的主体重叠的布尔值。|
||[position](/javascript/api/excel/excel.chartlegendupdatedata#position)|表示图例在图表上的位置。 有关详细信息, 请参阅 ChartLegendPosition。|
||[visible](/javascript/api/excel/excel.chartlegendupdatedata#visible)|表示 ChartLegend 对象是否可见的布尔值。|
|[ChartLineFormat](/javascript/api/excel/excel.chartlineformat)|[clear()](/javascript/api/excel/excel.chartlineformat#clear--)|清除图表元素的线条格式。|
||[color](/javascript/api/excel/excel.chartlineformat#color)|表示图表中的线条颜色的 HTML 颜色代码。|
||[set (properties: ChartLineFormat)](/javascript/api/excel/excel.chartlineformat#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: ChartLineFormatUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.chartlineformat#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
|[ChartLineFormatData](/javascript/api/excel/excel.chartlineformatdata)|[color](/javascript/api/excel/excel.chartlineformatdata#color)|表示图表中的线条颜色的 HTML 颜色代码。|
|[ChartLineFormatLoadOptions](/javascript/api/excel/excel.chartlineformatloadoptions)|[$all](/javascript/api/excel/excel.chartlineformatloadoptions#$all)||
||[color](/javascript/api/excel/excel.chartlineformatloadoptions#color)|表示图表中的线条颜色的 HTML 颜色代码。|
|[ChartLineFormatUpdateData](/javascript/api/excel/excel.chartlineformatupdatedata)|[color](/javascript/api/excel/excel.chartlineformatupdatedata#color)|表示图表中的线条颜色的 HTML 颜色代码。|
|[ChartLoadOptions](/javascript/api/excel/excel.chartloadoptions)|[$all](/javascript/api/excel/excel.chartloadoptions#$all)||
||[根](/javascript/api/excel/excel.chartloadoptions#axes)|表示图表坐标轴。|
||[dataLabels](/javascript/api/excel/excel.chartloadoptions#datalabels)|表示图表上的数据标签。|
||[format](/javascript/api/excel/excel.chartloadoptions#format)|封装图表区域的格式属性。|
||[height](/javascript/api/excel/excel.chartloadoptions#height)|表示 chart 对象的高度，以磅值表示。|
||[left](/javascript/api/excel/excel.chartloadoptions#left)|从图表左侧到工作表原点的距离，以磅值表示。|
||[图例](/javascript/api/excel/excel.chartloadoptions#legend)|表示图表的图例。|
||[name](/javascript/api/excel/excel.chartloadoptions#name)|表示 chart 对象的名称。|
||[series](/javascript/api/excel/excel.chartloadoptions#series)|表示单个系列或图表中的系列集合。|
||[title](/javascript/api/excel/excel.chartloadoptions#title)|表示指定图表的标题，包括标题的文本、可见性、位置和格式。|
||[top](/javascript/api/excel/excel.chartloadoptions#top)|表示从对象左边界至第 1 行顶部（在工作表上）或图表区域顶部（在图表上）的距离，以磅值表示。|
||[width](/javascript/api/excel/excel.chartloadoptions#width)|表示 chart 对象的宽度，以磅值表示。|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|[format](/javascript/api/excel/excel.chartpoint#format)|封装图表点的格式属性。 只读。|
||[value](/javascript/api/excel/excel.chartpoint#value)|返回图表点的值。 只读。|
||[set (properties: ChartPoint)](/javascript/api/excel/excel.chartpoint#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: ChartPointUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.chartpoint#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
|[ChartPointData](/javascript/api/excel/excel.chartpointdata)|[format](/javascript/api/excel/excel.chartpointdata#format)|封装图表点的格式属性。 只读。|
||[value](/javascript/api/excel/excel.chartpointdata#value)|返回图表点的值。 只读。|
|[ChartPointFormat](/javascript/api/excel/excel.chartpointformat)|[fill](/javascript/api/excel/excel.chartpointformat#fill)|代表图表的填充格式, 其中包括背景格式信息。 只读。|
||[set (properties: ChartPointFormat)](/javascript/api/excel/excel.chartpointformat#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: ChartPointFormatUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.chartpointformat#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
|[ChartPointFormatLoadOptions](/javascript/api/excel/excel.chartpointformatloadoptions)|[$all](/javascript/api/excel/excel.chartpointformatloadoptions#$all)||
|[ChartPointLoadOptions](/javascript/api/excel/excel.chartpointloadoptions)|[$all](/javascript/api/excel/excel.chartpointloadoptions#$all)||
||[format](/javascript/api/excel/excel.chartpointloadoptions#format)|封装图表点的格式属性。|
||[value](/javascript/api/excel/excel.chartpointloadoptions#value)|返回图表点的值。 只读。|
|[ChartPointUpdateData](/javascript/api/excel/excel.chartpointupdatedata)|[format](/javascript/api/excel/excel.chartpointupdatedata#format)|封装图表点的格式属性。|
|[ChartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|[getItemAt(index: number)](/javascript/api/excel/excel.chartpointscollection#getitemat-index-)|根据其在系列中的位置检索点。|
||[count](/javascript/api/excel/excel.chartpointscollection#count)|返回系列中的图表点数。 只读。|
||[items](/javascript/api/excel/excel.chartpointscollection#items)|获取此集合中已加载的子项。|
|[ChartPointsCollectionLoadOptions](/javascript/api/excel/excel.chartpointscollectionloadoptions)|[$all](/javascript/api/excel/excel.chartpointscollectionloadoptions#$all)||
||[format](/javascript/api/excel/excel.chartpointscollectionloadoptions#format)|对于集合中的每一项: 封装格式属性图表点。|
||[value](/javascript/api/excel/excel.chartpointscollectionloadoptions#value)|对于集合中的每一项: 返回图表点的值。 只读。|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[name](/javascript/api/excel/excel.chartseries#name)|表示图表中某个系列的名称。|
||[format](/javascript/api/excel/excel.chartseries#format)|表示图表系列的格式，包括填充和线条格式。 只读。|
||[点](/javascript/api/excel/excel.chartseries#points)|表示系列中所有数据点的集合。 只读。|
||[set (properties: ChartSeries)](/javascript/api/excel/excel.chartseries#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: ChartSeriesUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.chartseries#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[getItemAt(index: number)](/javascript/api/excel/excel.chartseriescollection#getitemat-index-)|根据其在集合中的位置检索系列|
||[count](/javascript/api/excel/excel.chartseriescollection#count)|返回集合中的系列数量。 只读。|
||[items](/javascript/api/excel/excel.chartseriescollection#items)|获取此集合中已加载的子项。|
|[ChartSeriesCollectionLoadOptions](/javascript/api/excel/excel.chartseriescollectionloadoptions)|[$all](/javascript/api/excel/excel.chartseriescollectionloadoptions#$all)||
||[format](/javascript/api/excel/excel.chartseriescollectionloadoptions#format)|对于集合中的每一项: 代表图表系列的格式, 其中包括填充和线条格式。|
||[name](/javascript/api/excel/excel.chartseriescollectionloadoptions#name)|对于集合中的每一项: 代表图表中系列的名称。|
||[点](/javascript/api/excel/excel.chartseriescollectionloadoptions#points)|对于集合中的每一项: 代表系列中所有数据点的集合。|
|[ChartSeriesData](/javascript/api/excel/excel.chartseriesdata)|[format](/javascript/api/excel/excel.chartseriesdata#format)|表示图表系列的格式，包括填充和线条格式。 只读。|
||[name](/javascript/api/excel/excel.chartseriesdata#name)|表示图表中某个系列的名称。|
||[点](/javascript/api/excel/excel.chartseriesdata#points)|表示系列中所有数据点的集合。 只读。|
|[ChartSeriesFormat](/javascript/api/excel/excel.chartseriesformat)|[fill](/javascript/api/excel/excel.chartseriesformat#fill)|表示图表系列的填充格式，包括背景格式信息。 只读。|
||[line](/javascript/api/excel/excel.chartseriesformat#line)|表示线条格式。 只读。|
||[set (properties: ChartSeriesFormat)](/javascript/api/excel/excel.chartseriesformat#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: ChartSeriesFormatUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.chartseriesformat#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
|[ChartSeriesFormatData](/javascript/api/excel/excel.chartseriesformatdata)|[line](/javascript/api/excel/excel.chartseriesformatdata#line)|表示线条格式。 只读。|
|[ChartSeriesFormatLoadOptions](/javascript/api/excel/excel.chartseriesformatloadoptions)|[$all](/javascript/api/excel/excel.chartseriesformatloadoptions#$all)||
||[line](/javascript/api/excel/excel.chartseriesformatloadoptions#line)|表示线条格式。|
|[ChartSeriesFormatUpdateData](/javascript/api/excel/excel.chartseriesformatupdatedata)|[line](/javascript/api/excel/excel.chartseriesformatupdatedata#line)|表示线条格式。|
|[ChartSeriesLoadOptions](/javascript/api/excel/excel.chartseriesloadoptions)|[$all](/javascript/api/excel/excel.chartseriesloadoptions#$all)||
||[format](/javascript/api/excel/excel.chartseriesloadoptions#format)|表示图表系列的格式，包括填充和线条格式。|
||[name](/javascript/api/excel/excel.chartseriesloadoptions#name)|表示图表中某个系列的名称。|
||[点](/javascript/api/excel/excel.chartseriesloadoptions#points)|表示系列中所有数据点的集合。|
|[ChartSeriesUpdateData](/javascript/api/excel/excel.chartseriesupdatedata)|[format](/javascript/api/excel/excel.chartseriesupdatedata#format)|表示图表系列的格式，包括填充和线条格式。|
||[name](/javascript/api/excel/excel.chartseriesupdatedata#name)|表示图表中某个系列的名称。|
|[ChartTitle](/javascript/api/excel/excel.charttitle)|[重叠](/javascript/api/excel/excel.charttitle#overlay)|表示图表标题是否将叠加在图表上的布尔值。|
||[format](/javascript/api/excel/excel.charttitle#format)|表示图表标题的格式，包括填充和字体格式。 只读。|
||[set (properties: ChartTitle)](/javascript/api/excel/excel.charttitle#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: ChartTitleUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.charttitle#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
||[text](/javascript/api/excel/excel.charttitle#text)|表示图表的标题文本。|
||[visible](/javascript/api/excel/excel.charttitle#visible)|表示图表标题对象是否可见的布尔值。|
|[ChartTitleData](/javascript/api/excel/excel.charttitledata)|[format](/javascript/api/excel/excel.charttitledata#format)|表示图表标题的格式，包括填充和字体格式。 只读。|
||[重叠](/javascript/api/excel/excel.charttitledata#overlay)|表示图表标题是否将叠加在图表上的布尔值。|
||[text](/javascript/api/excel/excel.charttitledata#text)|表示图表的标题文本。|
||[visible](/javascript/api/excel/excel.charttitledata#visible)|表示图表标题对象是否可见的布尔值。|
|[ChartTitleFormat](/javascript/api/excel/excel.charttitleformat)|[fill](/javascript/api/excel/excel.charttitleformat#fill)|表示对象的填充格式，包括背景格式信息。 只读。|
||[font](/javascript/api/excel/excel.charttitleformat#font)|表示对象的字体属性（字体名称、字体大小、颜色等）。 只读。|
||[set (properties: ChartTitleFormat)](/javascript/api/excel/excel.charttitleformat#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: ChartTitleFormatUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.charttitleformat#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
|[ChartTitleFormatData](/javascript/api/excel/excel.charttitleformatdata)|[font](/javascript/api/excel/excel.charttitleformatdata#font)|表示对象的字体属性（字体名称、字体大小、颜色等）。 只读。|
|[ChartTitleFormatLoadOptions](/javascript/api/excel/excel.charttitleformatloadoptions)|[$all](/javascript/api/excel/excel.charttitleformatloadoptions#$all)||
||[font](/javascript/api/excel/excel.charttitleformatloadoptions#font)|表示对象的字体属性（字体名称、字体大小、颜色等）。|
|[ChartTitleFormatUpdateData](/javascript/api/excel/excel.charttitleformatupdatedata)|[font](/javascript/api/excel/excel.charttitleformatupdatedata#font)|表示对象的字体属性（字体名称、字体大小、颜色等）。|
|[ChartTitleLoadOptions](/javascript/api/excel/excel.charttitleloadoptions)|[$all](/javascript/api/excel/excel.charttitleloadoptions#$all)||
||[format](/javascript/api/excel/excel.charttitleloadoptions#format)|表示图表标题的格式，包括填充和字体格式。|
||[重叠](/javascript/api/excel/excel.charttitleloadoptions#overlay)|表示图表标题是否将叠加在图表上的布尔值。|
||[text](/javascript/api/excel/excel.charttitleloadoptions#text)|表示图表的标题文本。|
||[visible](/javascript/api/excel/excel.charttitleloadoptions#visible)|表示图表标题对象是否可见的布尔值。|
|[ChartTitleUpdateData](/javascript/api/excel/excel.charttitleupdatedata)|[format](/javascript/api/excel/excel.charttitleupdatedata#format)|表示图表标题的格式，包括填充和字体格式。|
||[重叠](/javascript/api/excel/excel.charttitleupdatedata#overlay)|表示图表标题是否将叠加在图表上的布尔值。|
||[text](/javascript/api/excel/excel.charttitleupdatedata#text)|表示图表的标题文本。|
||[visible](/javascript/api/excel/excel.charttitleupdatedata#visible)|表示图表标题对象是否可见的布尔值。|
|[ChartUpdateData](/javascript/api/excel/excel.chartupdatedata)|[根](/javascript/api/excel/excel.chartupdatedata#axes)|表示图表坐标轴。|
||[dataLabels](/javascript/api/excel/excel.chartupdatedata#datalabels)|表示图表上的数据标签。|
||[format](/javascript/api/excel/excel.chartupdatedata#format)|封装图表区域的格式属性。|
||[height](/javascript/api/excel/excel.chartupdatedata#height)|表示 chart 对象的高度，以磅值表示。|
||[left](/javascript/api/excel/excel.chartupdatedata#left)|从图表左侧到工作表原点的距离，以磅值表示。|
||[图例](/javascript/api/excel/excel.chartupdatedata#legend)|表示图表的图例。|
||[name](/javascript/api/excel/excel.chartupdatedata#name)|表示 chart 对象的名称。|
||[title](/javascript/api/excel/excel.chartupdatedata#title)|表示指定图表的标题，包括标题的文本、可见性、位置和格式。|
||[top](/javascript/api/excel/excel.chartupdatedata#top)|表示从对象左边界至第 1 行顶部（在工作表上）或图表区域顶部（在图表上）的距离，以磅值表示。|
||[width](/javascript/api/excel/excel.chartupdatedata#width)|表示 chart 对象的宽度，以磅值表示。|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[getRange()](/javascript/api/excel/excel.nameditem#getrange--)|返回与名称相关联的 Range 对象。 如果已命名项的类型不是 Range，将引发错误。|
||[name](/javascript/api/excel/excel.nameditem#name)|对象的名称。 只读。|
||[type](/javascript/api/excel/excel.nameditem#type)|指明 name 公式返回的值的类型。 有关详细信息, 请参阅 NamedItemType。 只读。|
||[value](/javascript/api/excel/excel.nameditem#value)|表示 name 公式计算出的值。 对于已命名区域，将返回区域地址。 只读。|
||[set (properties: NamedItem)](/javascript/api/excel/excel.nameditem#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: NamedItemUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.nameditem#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
||[visible](/javascript/api/excel/excel.nameditem#visible)|指定对象是否可见。|
|[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection)|[getItem(name: string)](/javascript/api/excel/excel.nameditemcollection#getitem-name-)|使用其名称获取 NamedItem 对象。|
||[items](/javascript/api/excel/excel.nameditemcollection#items)|获取此集合中已加载的子项。|
|[NamedItemCollectionLoadOptions](/javascript/api/excel/excel.nameditemcollectionloadoptions)|[$all](/javascript/api/excel/excel.nameditemcollectionloadoptions#$all)||
||[name](/javascript/api/excel/excel.nameditemcollectionloadoptions#name)|对于集合中的每一项: 对象的名称。 只读。|
||[type](/javascript/api/excel/excel.nameditemcollectionloadoptions#type)|对于集合中的每一项: 指示由名称的公式返回的值的类型。 有关详细信息, 请参阅 NamedItemType。 只读。|
||[value](/javascript/api/excel/excel.nameditemcollectionloadoptions#value)|对于集合中的每一项: 代表由名称的公式计算的值。 对于已命名区域，将返回区域地址。 只读。|
||[visible](/javascript/api/excel/excel.nameditemcollectionloadoptions#visible)|对于集合中的每一项: 指定对象是否可见。|
|[NamedItemData](/javascript/api/excel/excel.nameditemdata)|[name](/javascript/api/excel/excel.nameditemdata#name)|对象的名称。 只读。|
||[type](/javascript/api/excel/excel.nameditemdata#type)|指明 name 公式返回的值的类型。 有关详细信息, 请参阅 NamedItemType。 只读。|
||[value](/javascript/api/excel/excel.nameditemdata#value)|表示 name 公式计算出的值。 对于已命名区域，将返回区域地址。 只读。|
||[visible](/javascript/api/excel/excel.nameditemdata#visible)|指定对象是否可见。|
|[NamedItemLoadOptions](/javascript/api/excel/excel.nameditemloadoptions)|[$all](/javascript/api/excel/excel.nameditemloadoptions#$all)||
||[name](/javascript/api/excel/excel.nameditemloadoptions#name)|对象的名称。 只读。|
||[type](/javascript/api/excel/excel.nameditemloadoptions#type)|指明 name 公式返回的值的类型。 有关详细信息, 请参阅 NamedItemType。 只读。|
||[value](/javascript/api/excel/excel.nameditemloadoptions#value)|表示 name 公式计算出的值。 对于已命名区域，将返回区域地址。 只读。|
||[visible](/javascript/api/excel/excel.nameditemloadoptions#visible)|指定对象是否可见。|
|[NamedItemUpdateData](/javascript/api/excel/excel.nameditemupdatedata)|[visible](/javascript/api/excel/excel.nameditemupdatedata#visible)|指定对象是否可见。|
|[Range](/javascript/api/excel/excel.range)|[clear(applyTo?: "All" \| "Formats" \| "Contents" \| "Hyperlinks" \| "RemoveHyperlinks")](/javascript/api/excel/excel.range#clear-applyto-)|清除区域值、格式、填充、边框等。|
||[clear(applyTo?: Excel.ClearApplyTo)](/javascript/api/excel/excel.range#clear-applyto-)|清除区域值、格式、填充、边框等。|
||[删除 (shift: "Up" \| "" 左 ")](/javascript/api/excel/excel.range#delete-shift-)|删除与区域相关的单元格。|
||[delete (shift: DeleteShiftDirection)](/javascript/api/excel/excel.range#delete-shift-)|删除与区域相关的单元格。|
||[formulas](/javascript/api/excel/excel.range#formulas)|表示采用 A1 表示法的公式。|
||[formulasLocal](/javascript/api/excel/excel.range#formulaslocal)|表示采用 A1 样式表示法的公式，使用用户的语言和数字格式区域设置。例如，英语中的公式 "=SUM(A1, 1.5)" 在德语中将变为 "=SUMME(A1; 1,5)"。|
||[getBoundingRect (anotherRange: Range \|字符串)](/javascript/api/excel/excel.range#getboundingrect-anotherrange-)|获取包含指定区域的最小 range 对象。 例如，“B2:C5”和“D10:E15”的 GetBoundingRect 为“B2:E15”。|
||[getCell(row: number, column: number)](/javascript/api/excel/excel.range#getcell-row--column-)|根据行和列编号获取包含单个单元格的 range 对象。 单元格可以位于其父区域的边界之外, 但前提是它停留在工作表网格中。 返回的单元格位于相对于区域左上角的单元格的位置。|
||[getColumn(column: number)](/javascript/api/excel/excel.range#getcolumn-column-)|获取区域中包含的列。|
||[getEntireColumn()](/javascript/api/excel/excel.range#getentirecolumn--)|获取一个对象, 该对象代表区域的整列 (例如, 如果当前区域表示单元格 "B4: E11", 则它`getEntireColumn`是表示列 "B:E" 的区域)。|
||[getEntireRow()](/javascript/api/excel/excel.range#getentirerow--)|获取一个对象, 该对象表示区域的整行 (例如, 如果当前区域表示单元格 "B4: E11", 则它`GetEntireRow`是表示行 "4:11" 的区域)。|
||[getIntersection (anotherRange: Range \|字符串)](/javascript/api/excel/excel.range#getintersection-anotherrange-)|获取表示指定区域的矩形交集的 range 对象。|
||[getLastCell()](/javascript/api/excel/excel.range#getlastcell--)|获取区域内的最后一个单元格。例如，“B2:D5”的最后一个单元格是“D5”。|
||[getLastColumn()](/javascript/api/excel/excel.range#getlastcolumn--)|获取区域内的最后一列。例如，“B2:D5”的最后一列是“D2:D5”。|
||[getLastRow()](/javascript/api/excel/excel.range#getlastrow--)|获取区域内的最后一行。例如，“B2:D5”的最后一行是“B5:D5”。|
||[getOffsetRange(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-)|获取表示与指定区域偏移的区域的对象。返回的区域的尺寸将与此区域一致。如果强制在工作表网格的边界之外生成区域，将引发错误。|
||[getRow(row: number)](/javascript/api/excel/excel.range#getrow-row-)|获取范围中包含的行。|
||[插入 (shift: "向下\| " "" 向右 ")](/javascript/api/excel/excel.range#insert-shift-)|将单个单元格或一系列单元格插入到工作表中取代此区域，并移动其他单元格以留出空间。在现在空白的空间返回新的 Range 对象。|
||[insert (shift: InsertShiftDirection)](/javascript/api/excel/excel.range#insert-shift-)|将单个单元格或一系列单元格插入到工作表中取代此区域，并移动其他单元格以留出空间。在现在空白的空间返回新的 Range 对象。|
||[numberFormat](/javascript/api/excel/excel.range#numberformat)|表示给定范围的 Excel 数字格式代码。|
||[address](/javascript/api/excel/excel.range#address)|表示 A1 样式的区域引用。 Address 值将包含工作表引用 (例如, "Sheet1!A1: B4 ")。 只读。|
||[addressLocal](/javascript/api/excel/excel.range#addresslocal)|以用户语言表示对指定区域的区域引用。 只读。|
||[cellCount](/javascript/api/excel/excel.range#cellcount)|范围中的单元格数。 如果单元格数超过 2^31-1 (2,147,483,647)，此 API 返回 -1。 只读。|
||[columnCount](/javascript/api/excel/excel.range#columncount)|表示区域中的列总数。 只读。|
||[columnIndex](/javascript/api/excel/excel.range#columnindex)|表示区域中第一个单元格的列编号。 从零开始编制索引。 只读。|
||[format](/javascript/api/excel/excel.range#format)|返回一个格式对象，其中封装了区域的字体、填充、边框、对齐方式和其他属性。 只读。|
||[rowCount](/javascript/api/excel/excel.range#rowcount)|返回区域中的总行数。 只读。|
||[rowIndex](/javascript/api/excel/excel.range#rowindex)|返回区域中第一个单元格的行编号。 从零开始编制索引。 只读。|
||[text](/javascript/api/excel/excel.range#text)|指定区域的文本值。文本值与单元格宽度无关。在 Excel UI 中替代 # 符号不会影响 API 返回的文本值。只读。|
||[valueTypes](/javascript/api/excel/excel.range#valuetypes)|表示每个单元格的数据类型。 只读。|
||[worksheet](/javascript/api/excel/excel.range#worksheet)|包含当前区域的工作表。 只读。|
||[select()](/javascript/api/excel/excel.range#select--)|在 Excel UI 中选择指定的区域。|
||[set (properties: Excel Range)](/javascript/api/excel/excel.range#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: RangeUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.range#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
||[track()](/javascript/api/excel/excel.range#track--)|根据文档中的相应更改来跟踪对象，以便进行自动调整。 此调用是 context.trackedObjects.add(thisObject) 的缩写。 如果你在“.sync”调用之间和按顺序执行“.run”批处理之外使用此对象，并且在对象上设置属性或调用方法时出现“InvalidObjectPath”错误，则需要在首次创建对象时为跟踪的对象集合添加对象。|
||[untrack()](/javascript/api/excel/excel.range#untrack--)|释放与此对象关联的内存（如果先前已跟踪过）。 此调用是 context.trackedObjects.add(thisObject) 的缩写。 拥有许多跟踪对象会降低主机应用程序的速度，因此请在使用完毕后释放所添加的任何对象。 在内存释放生效之前，你需要调用“context.sync()”。|
||[values](/javascript/api/excel/excel.range#values)|表示指定区域的原始值。 返回的数据可能是字符串、数字，也可能是布尔值。 包含错误的单元格将返回错误字符串。|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[color](/javascript/api/excel/excel.rangeborder#color)|表示窗体 #RRGGBB（例如“FFA500”）的边框线条颜色或作为已命名的 HTML 颜色（例如“orange”）的 HTML 颜色代码。|
||[sideIndex](/javascript/api/excel/excel.rangeborder#sideindex)|指示边框的特定边的常量值。 有关详细信息, 请参阅 BorderIndex。 只读。|
||[set (properties: RangeBorder)](/javascript/api/excel/excel.rangeborder#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: RangeBorderUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.rangeborder#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
||[style](/javascript/api/excel/excel.rangeborder#style)|线条样式的常量之一，指定边框的线条样式。 有关详细信息, 请参阅 BorderLineStyle。|
||[weight](/javascript/api/excel/excel.rangeborder#weight)|指定区域周围的边框的粗细。 有关详细信息, 请参阅 BorderWeight。|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[getItem (index: "EdgeTop" \| "EdgeBottom" \| "EdgeLeft" \| "EdgeRight" \| "InsideVertical" \| "InsideHorizontal" \| "" DiagonalDown \| "" DiagonalUp ")](/javascript/api/excel/excel.rangebordercollection#getitem-index-)|使用其名称获取 border 对象|
||[getItem (index: BorderIndex)](/javascript/api/excel/excel.rangebordercollection#getitem-index-)|使用其名称获取 border 对象|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangebordercollection#getitemat-index-)|使用其索引获取 border 对象|
||[count](/javascript/api/excel/excel.rangebordercollection#count)|集合中的 border 对象数量。 只读。|
||[items](/javascript/api/excel/excel.rangebordercollection#items)|获取此集合中已加载的子项。|
|[RangeBorderCollectionLoadOptions](/javascript/api/excel/excel.rangebordercollectionloadoptions)|[$all](/javascript/api/excel/excel.rangebordercollectionloadoptions#$all)||
||[color](/javascript/api/excel/excel.rangebordercollectionloadoptions#color)|对于集合中的每一项: HTML 颜色代码, 表示边框线的颜色, 格式 #RRGGBB (例如 "FFA500") 或命名的 HTML 颜色 (例如 "橙色")。|
||[sideIndex](/javascript/api/excel/excel.rangebordercollectionloadoptions#sideindex)|对于集合中的每一项: 常量值, 用于指示边框的特定侧。 有关详细信息, 请参阅 BorderIndex。 只读。|
||[style](/javascript/api/excel/excel.rangebordercollectionloadoptions#style)|对于集合中的每个项目: 指定边框线条样式的线条样式的常量之一。 有关详细信息, 请参阅 BorderLineStyle。|
||[weight](/javascript/api/excel/excel.rangebordercollectionloadoptions#weight)|对于集合中的每个项目: 指定某一范围周围的边框的粗细。 有关详细信息, 请参阅 BorderWeight。|
|[RangeBorderData](/javascript/api/excel/excel.rangeborderdata)|[color](/javascript/api/excel/excel.rangeborderdata#color)|表示窗体 #RRGGBB（例如“FFA500”）的边框线条颜色或作为已命名的 HTML 颜色（例如“orange”）的 HTML 颜色代码。|
||[sideIndex](/javascript/api/excel/excel.rangeborderdata#sideindex)|指示边框的特定边的常量值。 有关详细信息, 请参阅 BorderIndex。 只读。|
||[style](/javascript/api/excel/excel.rangeborderdata#style)|线条样式的常量之一，指定边框的线条样式。 有关详细信息, 请参阅 BorderLineStyle。|
||[weight](/javascript/api/excel/excel.rangeborderdata#weight)|指定区域周围的边框的粗细。 有关详细信息, 请参阅 BorderWeight。|
|[RangeBorderLoadOptions](/javascript/api/excel/excel.rangeborderloadoptions)|[$all](/javascript/api/excel/excel.rangeborderloadoptions#$all)||
||[color](/javascript/api/excel/excel.rangeborderloadoptions#color)|表示窗体 #RRGGBB（例如“FFA500”）的边框线条颜色或作为已命名的 HTML 颜色（例如“orange”）的 HTML 颜色代码。|
||[sideIndex](/javascript/api/excel/excel.rangeborderloadoptions#sideindex)|指示边框的特定边的常量值。 有关详细信息, 请参阅 BorderIndex。 只读。|
||[style](/javascript/api/excel/excel.rangeborderloadoptions#style)|线条样式的常量之一，指定边框的线条样式。 有关详细信息, 请参阅 BorderLineStyle。|
||[weight](/javascript/api/excel/excel.rangeborderloadoptions#weight)|指定区域周围的边框的粗细。 有关详细信息, 请参阅 BorderWeight。|
|[RangeBorderUpdateData](/javascript/api/excel/excel.rangeborderupdatedata)|[color](/javascript/api/excel/excel.rangeborderupdatedata#color)|表示窗体 #RRGGBB（例如“FFA500”）的边框线条颜色或作为已命名的 HTML 颜色（例如“orange”）的 HTML 颜色代码。|
||[style](/javascript/api/excel/excel.rangeborderupdatedata#style)|线条样式的常量之一，指定边框的线条样式。 有关详细信息, 请参阅 BorderLineStyle。|
||[weight](/javascript/api/excel/excel.rangeborderupdatedata#weight)|指定区域周围的边框的粗细。 有关详细信息, 请参阅 BorderWeight。|
|[RangeData](/javascript/api/excel/excel.rangedata)|[address](/javascript/api/excel/excel.rangedata#address)|表示 A1 样式的区域引用。 Address 值将包含工作表引用 (例如, "Sheet1!A1: B4 ")。 只读。|
||[addressLocal](/javascript/api/excel/excel.rangedata#addresslocal)|以用户语言表示对指定区域的区域引用。 只读。|
||[cellCount](/javascript/api/excel/excel.rangedata#cellcount)|范围中的单元格数。 如果单元格数超过 2^31-1 (2,147,483,647)，此 API 返回 -1。 只读。|
||[columnCount](/javascript/api/excel/excel.rangedata#columncount)|表示区域中的列总数。 只读。|
||[columnIndex](/javascript/api/excel/excel.rangedata#columnindex)|表示区域中第一个单元格的列编号。 从零开始编制索引。 只读。|
||[format](/javascript/api/excel/excel.rangedata#format)|返回一个格式对象，其中封装了区域的字体、填充、边框、对齐方式和其他属性。 只读。|
||[formulas](/javascript/api/excel/excel.rangedata#formulas)|表示采用 A1 表示法的公式。|
||[formulasLocal](/javascript/api/excel/excel.rangedata#formulaslocal)|表示采用 A1 样式表示法的公式，使用用户的语言和数字格式区域设置。例如，英语中的公式 "=SUM(A1, 1.5)" 在德语中将变为 "=SUMME(A1; 1,5)"。|
||[numberFormat](/javascript/api/excel/excel.rangedata#numberformat)|表示给定范围的 Excel 数字格式代码。|
||[rowCount](/javascript/api/excel/excel.rangedata#rowcount)|返回区域中的总行数。 只读。|
||[rowIndex](/javascript/api/excel/excel.rangedata#rowindex)|返回区域中第一个单元格的行编号。 从零开始编制索引。 只读。|
||[text](/javascript/api/excel/excel.rangedata#text)|指定区域的文本值。文本值与单元格宽度无关。在 Excel UI 中替代 # 符号不会影响 API 返回的文本值。只读。|
||[valueTypes](/javascript/api/excel/excel.rangedata#valuetypes)|表示每个单元格的数据类型。 只读。|
||[values](/javascript/api/excel/excel.rangedata#values)|表示指定区域的原始值。 返回的数据可能是字符串、数字，也可能是布尔值。 包含错误的单元格将返回错误字符串。|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[clear()](/javascript/api/excel/excel.rangefill#clear--)|重置区域背景。|
||[color](/javascript/api/excel/excel.rangefill#color)|表示窗体 #RRGGBB（例如“FFA500”）的边框线条颜色或作为已命名的 HTML 颜色（例如“orange”）的 HTML 颜色代码。|
||[set (properties: RangeFill)](/javascript/api/excel/excel.rangefill#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: RangeFillUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.rangefill#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
|[RangeFillData](/javascript/api/excel/excel.rangefilldata)|[color](/javascript/api/excel/excel.rangefilldata#color)|表示窗体 #RRGGBB（例如“FFA500”）的边框线条颜色或作为已命名的 HTML 颜色（例如“orange”）的 HTML 颜色代码。|
|[RangeFillLoadOptions](/javascript/api/excel/excel.rangefillloadoptions)|[$all](/javascript/api/excel/excel.rangefillloadoptions#$all)||
||[color](/javascript/api/excel/excel.rangefillloadoptions#color)|表示窗体 #RRGGBB（例如“FFA500”）的边框线条颜色或作为已命名的 HTML 颜色（例如“orange”）的 HTML 颜色代码。|
|[RangeFillUpdateData](/javascript/api/excel/excel.rangefillupdatedata)|[color](/javascript/api/excel/excel.rangefillupdatedata#color)|表示窗体 #RRGGBB（例如“FFA500”）的边框线条颜色或作为已命名的 HTML 颜色（例如“orange”）的 HTML 颜色代码。|
|[RangeFont](/javascript/api/excel/excel.rangefont)|[bold](/javascript/api/excel/excel.rangefont#bold)|表示字体的加粗状态。|
||[color](/javascript/api/excel/excel.rangefont#color)|文本颜色的 HTML 颜色代码表示。 例如， #FF0000 代表红色。|
||[italic](/javascript/api/excel/excel.rangefont#italic)|表示字体的斜体状态。|
||[name](/javascript/api/excel/excel.rangefont#name)|字体名称（例如"Calibri"）|
||[set (properties: RangeFont)](/javascript/api/excel/excel.rangefont#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: RangeFontUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.rangefont#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
||[size](/javascript/api/excel/excel.rangefont#size)|字号|
||[underline](/javascript/api/excel/excel.rangefont#underline)|应用于字体的下划线类型。 有关详细信息, 请参阅 RangeUnderlineStyle。|
|[RangeFontData](/javascript/api/excel/excel.rangefontdata)|[bold](/javascript/api/excel/excel.rangefontdata#bold)|表示字体的加粗状态。|
||[color](/javascript/api/excel/excel.rangefontdata#color)|文本颜色的 HTML 颜色代码表示。 例如， #FF0000 代表红色。|
||[italic](/javascript/api/excel/excel.rangefontdata#italic)|表示字体的斜体状态。|
||[name](/javascript/api/excel/excel.rangefontdata#name)|字体名称（例如"Calibri"）|
||[size](/javascript/api/excel/excel.rangefontdata#size)|字号|
||[underline](/javascript/api/excel/excel.rangefontdata#underline)|应用于字体的下划线类型。 有关详细信息, 请参阅 RangeUnderlineStyle。|
|[RangeFontLoadOptions](/javascript/api/excel/excel.rangefontloadoptions)|[$all](/javascript/api/excel/excel.rangefontloadoptions#$all)||
||[bold](/javascript/api/excel/excel.rangefontloadoptions#bold)|表示字体的加粗状态。|
||[color](/javascript/api/excel/excel.rangefontloadoptions#color)|文本颜色的 HTML 颜色代码表示。 例如， #FF0000 代表红色。|
||[italic](/javascript/api/excel/excel.rangefontloadoptions#italic)|表示字体的斜体状态。|
||[name](/javascript/api/excel/excel.rangefontloadoptions#name)|字体名称（例如"Calibri"）|
||[size](/javascript/api/excel/excel.rangefontloadoptions#size)|字号|
||[underline](/javascript/api/excel/excel.rangefontloadoptions#underline)|应用于字体的下划线类型。 有关详细信息, 请参阅 RangeUnderlineStyle。|
|[RangeFontUpdateData](/javascript/api/excel/excel.rangefontupdatedata)|[bold](/javascript/api/excel/excel.rangefontupdatedata#bold)|表示字体的加粗状态。|
||[color](/javascript/api/excel/excel.rangefontupdatedata#color)|文本颜色的 HTML 颜色代码表示。 例如， #FF0000 代表红色。|
||[italic](/javascript/api/excel/excel.rangefontupdatedata#italic)|表示字体的斜体状态。|
||[name](/javascript/api/excel/excel.rangefontupdatedata#name)|字体名称（例如"Calibri"）|
||[size](/javascript/api/excel/excel.rangefontupdatedata#size)|字号|
||[underline](/javascript/api/excel/excel.rangefontupdatedata#underline)|应用于字体的下划线类型。 有关详细信息, 请参阅 RangeUnderlineStyle。|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[horizontalAlignment](/javascript/api/excel/excel.rangeformat#horizontalalignment)|表示指定对象的水平对齐方式。 有关详细信息, 请参阅 HorizontalAlignment。|
||[Borders](/javascript/api/excel/excel.rangeformat#borders)|应用于整个区域的 Border 对象的集合。 只读。|
||[fill](/javascript/api/excel/excel.rangeformat#fill)|返回在整个区域内定义的 fill 对象。 只读。|
||[font](/javascript/api/excel/excel.rangeformat#font)|返回在整个区域内定义的 Font 对象。 只读。|
||[set (properties: RangeFormat)](/javascript/api/excel/excel.rangeformat#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: RangeFormatUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.rangeformat#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
||[verticalAlignment](/javascript/api/excel/excel.rangeformat#verticalalignment)|表示指定对象的垂直对齐方式。 有关详细信息, 请参阅 VerticalAlignment。|
||[wrapText](/javascript/api/excel/excel.rangeformat#wraptext)|指示 Excel 是否将对象中的文本换行。 指示整个区域不具有统一换行设置的空值|
|[RangeFormatData](/javascript/api/excel/excel.rangeformatdata)|[Borders](/javascript/api/excel/excel.rangeformatdata#borders)|应用于整个区域的 Border 对象的集合。 只读。|
||[fill](/javascript/api/excel/excel.rangeformatdata#fill)|返回在整个区域内定义的 fill 对象。 只读。|
||[font](/javascript/api/excel/excel.rangeformatdata#font)|返回在整个区域内定义的 Font 对象。 只读。|
||[horizontalAlignment](/javascript/api/excel/excel.rangeformatdata#horizontalalignment)|表示指定对象的水平对齐方式。 有关详细信息, 请参阅 HorizontalAlignment。|
||[verticalAlignment](/javascript/api/excel/excel.rangeformatdata#verticalalignment)|表示指定对象的垂直对齐方式。 有关详细信息, 请参阅 VerticalAlignment。|
||[wrapText](/javascript/api/excel/excel.rangeformatdata#wraptext)|指示 Excel 是否将对象中的文本换行。 指示整个区域不具有统一换行设置的空值|
|[RangeFormatLoadOptions](/javascript/api/excel/excel.rangeformatloadoptions)|[$all](/javascript/api/excel/excel.rangeformatloadoptions#$all)||
||[Borders](/javascript/api/excel/excel.rangeformatloadoptions#borders)|应用于整个区域的 Border 对象的集合。|
||[fill](/javascript/api/excel/excel.rangeformatloadoptions#fill)|返回在整个区域内定义的 fill 对象。|
||[font](/javascript/api/excel/excel.rangeformatloadoptions#font)|返回在整个区域内定义的 Font 对象。|
||[horizontalAlignment](/javascript/api/excel/excel.rangeformatloadoptions#horizontalalignment)|表示指定对象的水平对齐方式。 有关详细信息, 请参阅 HorizontalAlignment。|
||[verticalAlignment](/javascript/api/excel/excel.rangeformatloadoptions#verticalalignment)|表示指定对象的垂直对齐方式。 有关详细信息, 请参阅 VerticalAlignment。|
||[wrapText](/javascript/api/excel/excel.rangeformatloadoptions#wraptext)|指示 Excel 是否将对象中的文本换行。 指示整个区域不具有统一换行设置的空值|
|[RangeFormatUpdateData](/javascript/api/excel/excel.rangeformatupdatedata)|[Borders](/javascript/api/excel/excel.rangeformatupdatedata#borders)|应用于整个区域的 Border 对象的集合。|
||[fill](/javascript/api/excel/excel.rangeformatupdatedata#fill)|返回在整个区域内定义的 fill 对象。|
||[font](/javascript/api/excel/excel.rangeformatupdatedata#font)|返回在整个区域内定义的 Font 对象。|
||[horizontalAlignment](/javascript/api/excel/excel.rangeformatupdatedata#horizontalalignment)|表示指定对象的水平对齐方式。 有关详细信息, 请参阅 HorizontalAlignment。|
||[verticalAlignment](/javascript/api/excel/excel.rangeformatupdatedata#verticalalignment)|表示指定对象的垂直对齐方式。 有关详细信息, 请参阅 VerticalAlignment。|
||[wrapText](/javascript/api/excel/excel.rangeformatupdatedata#wraptext)|指示 Excel 是否将对象中的文本换行。 指示整个区域不具有统一换行设置的空值|
|[RangeLoadOptions](/javascript/api/excel/excel.rangeloadoptions)|[$all](/javascript/api/excel/excel.rangeloadoptions#$all)||
||[address](/javascript/api/excel/excel.rangeloadoptions#address)|表示 A1 样式的区域引用。 Address 值将包含工作表引用 (例如, "Sheet1!A1: B4 ")。 只读。|
||[addressLocal](/javascript/api/excel/excel.rangeloadoptions#addresslocal)|以用户语言表示对指定区域的区域引用。 只读。|
||[cellCount](/javascript/api/excel/excel.rangeloadoptions#cellcount)|范围中的单元格数。 如果单元格数超过 2^31-1 (2,147,483,647)，此 API 返回 -1。 只读。|
||[columnCount](/javascript/api/excel/excel.rangeloadoptions#columncount)|表示区域中的列总数。 只读。|
||[columnIndex](/javascript/api/excel/excel.rangeloadoptions#columnindex)|表示区域中第一个单元格的列编号。 从零开始编制索引。 只读。|
||[format](/javascript/api/excel/excel.rangeloadoptions#format)|返回一个格式对象，其中封装了区域的字体、填充、边框、对齐方式和其他属性。|
||[formulas](/javascript/api/excel/excel.rangeloadoptions#formulas)|表示采用 A1 表示法的公式。|
||[formulasLocal](/javascript/api/excel/excel.rangeloadoptions#formulaslocal)|表示采用 A1 样式表示法的公式，使用用户的语言和数字格式区域设置。例如，英语中的公式 "=SUM(A1, 1.5)" 在德语中将变为 "=SUMME(A1; 1,5)"。|
||[numberFormat](/javascript/api/excel/excel.rangeloadoptions#numberformat)|表示给定范围的 Excel 数字格式代码。|
||[rowCount](/javascript/api/excel/excel.rangeloadoptions#rowcount)|返回区域中的总行数。 只读。|
||[rowIndex](/javascript/api/excel/excel.rangeloadoptions#rowindex)|返回区域中第一个单元格的行编号。 从零开始编制索引。 只读。|
||[text](/javascript/api/excel/excel.rangeloadoptions#text)|指定区域的文本值。文本值与单元格宽度无关。在 Excel UI 中替代 # 符号不会影响 API 返回的文本值。只读。|
||[valueTypes](/javascript/api/excel/excel.rangeloadoptions#valuetypes)|表示每个单元格的数据类型。 只读。|
||[values](/javascript/api/excel/excel.rangeloadoptions#values)|表示指定区域的原始值。 返回的数据可能是字符串、数字，也可能是布尔值。 包含错误的单元格将返回错误字符串。|
||[worksheet](/javascript/api/excel/excel.rangeloadoptions#worksheet)|包含当前区域的工作表。|
|[RangeUpdateData](/javascript/api/excel/excel.rangeupdatedata)|[format](/javascript/api/excel/excel.rangeupdatedata#format)|返回一个格式对象，其中封装了区域的字体、填充、边框、对齐方式和其他属性。|
||[formulas](/javascript/api/excel/excel.rangeupdatedata#formulas)|表示采用 A1 表示法的公式。|
||[formulasLocal](/javascript/api/excel/excel.rangeupdatedata#formulaslocal)|表示采用 A1 样式表示法的公式，使用用户的语言和数字格式区域设置。例如，英语中的公式 "=SUM(A1, 1.5)" 在德语中将变为 "=SUMME(A1; 1,5)"。|
||[numberFormat](/javascript/api/excel/excel.rangeupdatedata#numberformat)|表示给定范围的 Excel 数字格式代码。|
||[values](/javascript/api/excel/excel.rangeupdatedata#values)|表示指定区域的原始值。 返回的数据可能是字符串、数字，也可能是布尔值。 包含错误的单元格将返回错误字符串。|
|[Table](/javascript/api/excel/excel.table)|[delete()](/javascript/api/excel/excel.table#delete--)|删除表。|
||[getDataBodyRange()](/javascript/api/excel/excel.table#getdatabodyrange--)|获取与表的数据体相关的 range 对象。|
||[getHeaderRowRange()](/javascript/api/excel/excel.table#getheaderrowrange--)|获取与表的标题行相关的 range 对象。|
||[getRange()](/javascript/api/excel/excel.table#getrange--)|获取与整个表相关的 range 对象。|
||[getTotalRowRange()](/javascript/api/excel/excel.table#gettotalrowrange--)|获取与表的总计行相关的 range 对象。|
||[name](/javascript/api/excel/excel.table#name)|表的名称。|
||[列](/javascript/api/excel/excel.table#columns)|表示表中所有列的集合。 只读。|
||[id](/javascript/api/excel/excel.table#id)|返回用于唯一标识指定工作簿中表的值。即使表被重命名，标识符的值仍然相同。只读。|
||[rows](/javascript/api/excel/excel.table#rows)|表示表中所有行的集合。 只读。|
||[set (properties: Excel. Table)](/javascript/api/excel/excel.table#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: TableUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.table#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
||[showHeaders](/javascript/api/excel/excel.table#showheaders)|指示标头行是否可见。 该值可以设置为显示或删除标头行。|
||[showTotals](/javascript/api/excel/excel.table#showtotals)|指示总计行是否可见。 该值可以设置为显示或删除总计行。|
||[style](/javascript/api/excel/excel.table#style)|表示表格样式的常量值。可能的值是：TableStyleLight1 thru TableStyleLight21、TableStyleMedium1 thru TableStyleMedium28、TableStyleStyleDark1 thru TableStyleStyleDark11。还可以指定工作簿中显示的用户定义的自定义样式。|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[add (address: Range \| String, hasHeaders: boolean)](/javascript/api/excel/excel.tablecollection#add-address--hasheaders-)|新建表。范围对象或源地址决定了在哪个工作表下添加表。如果无法添加表（例如，由于地址无效，或者表与另一个表重叠），则会引发错误。|
||[getItem(key: string)](/javascript/api/excel/excel.tablecollection#getitem-key-)|按名称或 ID 获取表。|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablecollection#getitemat-index-)|根据其在集合中的位置获取表。|
||[count](/javascript/api/excel/excel.tablecollection#count)|返回工作簿中的表数目。 只读。|
||[items](/javascript/api/excel/excel.tablecollection#items)|获取此集合中已加载的子项。|
|[TableCollectionLoadOptions](/javascript/api/excel/excel.tablecollectionloadoptions)|[$all](/javascript/api/excel/excel.tablecollectionloadoptions#$all)||
||[列](/javascript/api/excel/excel.tablecollectionloadoptions#columns)|对于集合中的每一项: 代表表中所有列的集合。|
||[id](/javascript/api/excel/excel.tablecollectionloadoptions#id)|对于集合中的每一项: 返回一个值, 该值唯一地标识给定工作簿中的表。 即使表被重命名，标识符的值仍然相同。 只读。|
||[name](/javascript/api/excel/excel.tablecollectionloadoptions#name)|对于集合中的每一项: 表的名称。|
||[rows](/javascript/api/excel/excel.tablecollectionloadoptions#rows)|对于集合中的每一项: 代表表中所有行的集合。|
||[showHeaders](/javascript/api/excel/excel.tablecollectionloadoptions#showheaders)|对于集合中的每一项: 指示标题行是否可见。 该值可以设置为显示或删除标头行。|
||[showTotals](/javascript/api/excel/excel.tablecollectionloadoptions#showtotals)|对于集合中的每一项: 指示汇总行是否可见。 该值可以设置为显示或删除总计行。|
||[style](/javascript/api/excel/excel.tablecollectionloadoptions#style)|对于集合中的每一项: 表示表样式的常量值。 可能的值是：TableStyleLight1 thru TableStyleLight21、TableStyleMedium1 thru TableStyleMedium28、TableStyleStyleDark1 thru TableStyleStyleDark11。 还可以指定工作簿中显示的用户定义的自定义样式。|
|[TableColumn](/javascript/api/excel/excel.tablecolumn)|[delete()](/javascript/api/excel/excel.tablecolumn#delete--)|从表中删除列。|
||[getDataBodyRange()](/javascript/api/excel/excel.tablecolumn#getdatabodyrange--)|获取与列的数据体相关的 range 对象。|
||[getHeaderRowRange()](/javascript/api/excel/excel.tablecolumn#getheaderrowrange--)|获取与列的标头行相关的 range 对象。|
||[getRange()](/javascript/api/excel/excel.tablecolumn#getrange--)|获取与整个列相关的 range 对象。|
||[getTotalRowRange()](/javascript/api/excel/excel.tablecolumn#gettotalrowrange--)|获取与列的总计行相关的 range 对象。|
||[name](/javascript/api/excel/excel.tablecolumn#name)|表示表列的名称。|
||[id](/javascript/api/excel/excel.tablecolumn#id)|返回标识表内的列的唯一键。 只读。|
||[index](/javascript/api/excel/excel.tablecolumn#index)|返回表的列集合内列的索引编号。 从零开始编制索引。 只读。|
||[set (properties: TableColumn)](/javascript/api/excel/excel.tablecolumn#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: TableColumnUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.tablecolumn#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
||[values](/javascript/api/excel/excel.tablecolumn#values)|表示指定区域的原始值。 返回的数据可能是字符串、数字，也可能是布尔值。 包含错误的单元格将返回错误字符串。|
|[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|[add (index？: number, values？: Array<Array<>> \| \| \| \|布尔字符串\|数字, name？: string)](/javascript/api/excel/excel.tablecolumncollection#add-index--values--name-)|向表中添加新列。|
||[getItem (key: 数字\|字符串)](/javascript/api/excel/excel.tablecolumncollection#getitem-key-)|按名称或 ID 获取 column 对象。|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablecolumncollection#getitemat-index-)|根据其在集合中的位置获取列。|
||[count](/javascript/api/excel/excel.tablecolumncollection#count)|返回表中的列数。 只读。|
||[items](/javascript/api/excel/excel.tablecolumncollection#items)|获取此集合中已加载的子项。|
|[TableColumnCollectionLoadOptions](/javascript/api/excel/excel.tablecolumncollectionloadoptions)|[$all](/javascript/api/excel/excel.tablecolumncollectionloadoptions#$all)||
||[id](/javascript/api/excel/excel.tablecolumncollectionloadoptions#id)|对于集合中的每一项: 返回标识表中的列的唯一键。 只读。|
||[index](/javascript/api/excel/excel.tablecolumncollectionloadoptions#index)|对于集合中的每一项: 返回表的列集合中的列的索引号。 从零开始编制索引。 只读。|
||[name](/javascript/api/excel/excel.tablecolumncollectionloadoptions#name)|对于集合中的每个项目: 代表表格列的名称。|
||[values](/javascript/api/excel/excel.tablecolumncollectionloadoptions#values)|对于集合中的每一项: 代表指定区域的原始值。 返回的数据可能是字符串、数字，也可能是布尔值。 包含错误的单元格将返回错误字符串。|
|[TableColumnData](/javascript/api/excel/excel.tablecolumndata)|[id](/javascript/api/excel/excel.tablecolumndata#id)|返回标识表内的列的唯一键。 只读。|
||[index](/javascript/api/excel/excel.tablecolumndata#index)|返回表的列集合内列的索引编号。 从零开始编制索引。 只读。|
||[name](/javascript/api/excel/excel.tablecolumndata#name)|表示表列的名称。|
||[values](/javascript/api/excel/excel.tablecolumndata#values)|表示指定区域的原始值。 返回的数据可能是字符串、数字，也可能是布尔值。 包含错误的单元格将返回错误字符串。|
|[TableColumnLoadOptions](/javascript/api/excel/excel.tablecolumnloadoptions)|[$all](/javascript/api/excel/excel.tablecolumnloadoptions#$all)||
||[id](/javascript/api/excel/excel.tablecolumnloadoptions#id)|返回标识表内的列的唯一键。 只读。|
||[index](/javascript/api/excel/excel.tablecolumnloadoptions#index)|返回表的列集合内列的索引编号。 从零开始编制索引。 只读。|
||[name](/javascript/api/excel/excel.tablecolumnloadoptions#name)|表示表列的名称。|
||[values](/javascript/api/excel/excel.tablecolumnloadoptions#values)|表示指定区域的原始值。 返回的数据可能是字符串、数字，也可能是布尔值。 包含错误的单元格将返回错误字符串。|
|[TableColumnUpdateData](/javascript/api/excel/excel.tablecolumnupdatedata)|[name](/javascript/api/excel/excel.tablecolumnupdatedata#name)|表示表列的名称。|
||[values](/javascript/api/excel/excel.tablecolumnupdatedata#values)|表示指定区域的原始值。 返回的数据可能是字符串、数字，也可能是布尔值。 包含错误的单元格将返回错误字符串。|
|[TableData](/javascript/api/excel/excel.tabledata)|[列](/javascript/api/excel/excel.tabledata#columns)|表示表中所有列的集合。 只读。|
||[id](/javascript/api/excel/excel.tabledata#id)|返回用于唯一标识指定工作簿中表的值。即使表被重命名，标识符的值仍然相同。只读。|
||[name](/javascript/api/excel/excel.tabledata#name)|表的名称。|
||[rows](/javascript/api/excel/excel.tabledata#rows)|表示表中所有行的集合。 只读。|
||[showHeaders](/javascript/api/excel/excel.tabledata#showheaders)|指示标头行是否可见。 该值可以设置为显示或删除标头行。|
||[showTotals](/javascript/api/excel/excel.tabledata#showtotals)|指示总计行是否可见。 该值可以设置为显示或删除总计行。|
||[style](/javascript/api/excel/excel.tabledata#style)|表示表格样式的常量值。可能的值是：TableStyleLight1 thru TableStyleLight21、TableStyleMedium1 thru TableStyleMedium28、TableStyleStyleDark1 thru TableStyleStyleDark11。还可以指定工作簿中显示的用户定义的自定义样式。|
|[TableLoadOptions](/javascript/api/excel/excel.tableloadoptions)|[$all](/javascript/api/excel/excel.tableloadoptions#$all)||
||[列](/javascript/api/excel/excel.tableloadoptions#columns)|表示表中所有列的集合。|
||[id](/javascript/api/excel/excel.tableloadoptions#id)|返回用于唯一标识指定工作簿中表的值。即使表被重命名，标识符的值仍然相同。只读。|
||[name](/javascript/api/excel/excel.tableloadoptions#name)|表的名称。|
||[rows](/javascript/api/excel/excel.tableloadoptions#rows)|表示表中所有行的集合。|
||[showHeaders](/javascript/api/excel/excel.tableloadoptions#showheaders)|指示标头行是否可见。 该值可以设置为显示或删除标头行。|
||[showTotals](/javascript/api/excel/excel.tableloadoptions#showtotals)|指示总计行是否可见。 该值可以设置为显示或删除总计行。|
||[style](/javascript/api/excel/excel.tableloadoptions#style)|表示表格样式的常量值。可能的值是：TableStyleLight1 thru TableStyleLight21、TableStyleMedium1 thru TableStyleMedium28、TableStyleStyleDark1 thru TableStyleStyleDark11。还可以指定工作簿中显示的用户定义的自定义样式。|
|[TableRow](/javascript/api/excel/excel.tablerow)|[delete()](/javascript/api/excel/excel.tablerow#delete--)|从表中删除行。|
||[getRange()](/javascript/api/excel/excel.tablerow#getrange--)|返回与整个行相关的 range 对象。|
||[index](/javascript/api/excel/excel.tablerow#index)|返回表的行集合内行的索引编号。 从零开始编制索引。 只读。|
||[set (properties: TableRow)](/javascript/api/excel/excel.tablerow#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: TableRowUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.tablerow#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
||[values](/javascript/api/excel/excel.tablerow#values)|表示指定区域的原始值。 返回的数据可能是字符串、数字，也可能是布尔值。 包含错误的单元格将返回错误字符串。|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[add (index？: number, values？: Array<Array<布尔\| \|字符串数字>> \|布尔\|字符串\|数字)](/javascript/api/excel/excel.tablerowcollection#add-index--values-)|向表中添加一行或多行。 返回对象是新添加的首行。|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablerowcollection#getitemat-index-)|根据其在集合中的位置获取行。|
||[count](/javascript/api/excel/excel.tablerowcollection#count)|返回表中的行数。 只读。|
||[items](/javascript/api/excel/excel.tablerowcollection#items)|获取此集合中已加载的子项。|
|[TableRowCollectionLoadOptions](/javascript/api/excel/excel.tablerowcollectionloadoptions)|[$all](/javascript/api/excel/excel.tablerowcollectionloadoptions#$all)||
||[index](/javascript/api/excel/excel.tablerowcollectionloadoptions#index)|对于集合中的每一项: 返回表的 rows 集合中的行的索引号。 从零开始编制索引。 只读。|
||[values](/javascript/api/excel/excel.tablerowcollectionloadoptions#values)|对于集合中的每一项: 代表指定区域的原始值。 返回的数据可能是字符串、数字，也可能是布尔值。 包含错误的单元格将返回错误字符串。|
|[TableRowData](/javascript/api/excel/excel.tablerowdata)|[index](/javascript/api/excel/excel.tablerowdata#index)|返回表的行集合内行的索引编号。 从零开始编制索引。 只读。|
||[values](/javascript/api/excel/excel.tablerowdata#values)|表示指定区域的原始值。 返回的数据可能是字符串、数字，也可能是布尔值。 包含错误的单元格将返回错误字符串。|
|[TableRowLoadOptions](/javascript/api/excel/excel.tablerowloadoptions)|[$all](/javascript/api/excel/excel.tablerowloadoptions#$all)||
||[index](/javascript/api/excel/excel.tablerowloadoptions#index)|返回表的行集合内行的索引编号。 从零开始编制索引。 只读。|
||[values](/javascript/api/excel/excel.tablerowloadoptions#values)|表示指定区域的原始值。 返回的数据可能是字符串、数字，也可能是布尔值。 包含错误的单元格将返回错误字符串。|
|[TableRowUpdateData](/javascript/api/excel/excel.tablerowupdatedata)|[values](/javascript/api/excel/excel.tablerowupdatedata#values)|表示指定区域的原始值。 返回的数据可能是字符串、数字，也可能是布尔值。 包含错误的单元格将返回错误字符串。|
|[TableUpdateData](/javascript/api/excel/excel.tableupdatedata)|[name](/javascript/api/excel/excel.tableupdatedata#name)|表的名称。|
||[showHeaders](/javascript/api/excel/excel.tableupdatedata#showheaders)|指示标头行是否可见。 该值可以设置为显示或删除标头行。|
||[showTotals](/javascript/api/excel/excel.tableupdatedata#showtotals)|指示总计行是否可见。 该值可以设置为显示或删除总计行。|
||[style](/javascript/api/excel/excel.tableupdatedata#style)|表示表格样式的常量值。可能的值是：TableStyleLight1 thru TableStyleLight21、TableStyleMedium1 thru TableStyleMedium28、TableStyleStyleDark1 thru TableStyleStyleDark11。还可以指定工作簿中显示的用户定义的自定义样式。|
|[Workbook](/javascript/api/excel/excel.workbook)|[getSelectedRange ()](/javascript/api/excel/excel.workbook#getselectedrange--)|从工作簿中获取当前选定的单个区域。 如果选择了多个区域, 则此方法将引发错误。|
||[application](/javascript/api/excel/excel.workbook#application)|表示包含此工作簿的 Excel 应用程序实例。 只读。|
||[bindings](/javascript/api/excel/excel.workbook#bindings)|表示属于工作簿的绑定的集合。 只读。|
||[names](/javascript/api/excel/excel.workbook#names)|表示工作簿范围内的已命名项目（称为区域和常量）的集合。 只读。|
||[表](/javascript/api/excel/excel.workbook#tables)|表示与工作簿关联的表的集合。 只读。|
||[单](/javascript/api/excel/excel.workbook#worksheets)|表示与工作簿关联的工作表的集合。 只读。|
||[set (properties: Excel. 工作簿)](/javascript/api/excel/excel.workbook#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: WorkbookUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.workbook#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
|[WorkbookData](/javascript/api/excel/excel.workbookdata)|[bindings](/javascript/api/excel/excel.workbookdata#bindings)|表示属于工作簿的绑定的集合。 只读。|
||[names](/javascript/api/excel/excel.workbookdata#names)|表示工作簿范围内的已命名项目（称为区域和常量）的集合。 只读。|
||[表](/javascript/api/excel/excel.workbookdata#tables)|表示与工作簿关联的表的集合。 只读。|
||[单](/javascript/api/excel/excel.workbookdata#worksheets)|表示与工作簿关联的工作表的集合。 只读。|
|[WorkbookLoadOptions](/javascript/api/excel/excel.workbookloadoptions)|[$all](/javascript/api/excel/excel.workbookloadoptions#$all)||
||[application](/javascript/api/excel/excel.workbookloadoptions#application)|表示包含此工作簿的 Excel 应用程序实例。|
||[bindings](/javascript/api/excel/excel.workbookloadoptions#bindings)|表示属于工作簿的绑定的集合。|
||[表](/javascript/api/excel/excel.workbookloadoptions#tables)|表示与工作簿关联的表的集合。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[activate()](/javascript/api/excel/excel.worksheet#activate--)|在 Excel UI 中激活工作表。|
||[delete()](/javascript/api/excel/excel.worksheet#delete--)|从工作簿中删除工作表。 请注意, 如果工作表的可见性设置为 "VeryHidden", 则删除操作将失败, 并出现 GeneralException。|
||[getCell(row: number, column: number)](/javascript/api/excel/excel.worksheet#getcell-row--column-)|根据行和列编号获取包含单个单元格的 range 对象。 单元格可以位于其父区域的边界之外, 但前提是它停留在工作表网格中。|
||[getRange (address？: string)](/javascript/api/excel/excel.worksheet#getrange-address-)|获取一个 range 对象, 该对象代表由地址或名称指定的单个矩形单元格块。|
||[name](/javascript/api/excel/excel.worksheet#name)|工作表的显示名称。|
||[position](/javascript/api/excel/excel.worksheet#position)|工作表在工作簿中的位置，从零开始。|
||[直方图](/javascript/api/excel/excel.worksheet#charts)|返回属于工作表的图表的集合。 只读。|
||[id](/javascript/api/excel/excel.worksheet#id)|返回用于唯一标识指定工作簿中工作表的值。即使工作表被重命名或移动，标识符的值仍然相同。只读。|
||[表](/javascript/api/excel/excel.worksheet#tables)|属于工作表的表的集合。 只读。|
||[set (properties: Excel. 工作表)](/javascript/api/excel/excel.worksheet#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: WorksheetUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.worksheet#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
||[深入](/javascript/api/excel/excel.worksheet#visibility)|工作表的可见性。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[add (name？: string)](/javascript/api/excel/excel.worksheetcollection#add-name-)|向工作簿添加新工作表。工作表将添加到现有工作表的末尾。如果您想要激活新添加的工作表，请对其调用 ".activate()。|
||[getActiveWorksheet()](/javascript/api/excel/excel.worksheetcollection#getactiveworksheet--)|获取工作簿中当前处于活动状态的工作表。|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcollection#getitem-key-)|按 Worksheet 对象的名称或 ID 获取此对象。|
||[items](/javascript/api/excel/excel.worksheetcollection#items)|获取此集合中已加载的子项。|
|[WorksheetCollectionLoadOptions](/javascript/api/excel/excel.worksheetcollectionloadoptions)|[$all](/javascript/api/excel/excel.worksheetcollectionloadoptions#$all)||
||[直方图](/javascript/api/excel/excel.worksheetcollectionloadoptions#charts)|对于集合中的每一项: 返回工作表的一部分的图表的集合。|
||[id](/javascript/api/excel/excel.worksheetcollectionloadoptions#id)|对于集合中的每一项: 返回一个值, 该值唯一地标识给定工作簿中的工作表。 即使工作表被重命名或移动，标识符的值仍然相同。 只读。|
||[name](/javascript/api/excel/excel.worksheetcollectionloadoptions#name)|对于集合中的每一项: 工作表的显示名称。|
||[position](/javascript/api/excel/excel.worksheetcollectionloadoptions#position)|对于集合中的每个项目: 工作表在工作簿中的位置 (从零开始)。|
||[表](/javascript/api/excel/excel.worksheetcollectionloadoptions#tables)|对于集合中的每一项: 工作表中的表的集合。|
||[深入](/javascript/api/excel/excel.worksheetcollectionloadoptions#visibility)|对于集合中的每一项: 工作表的可见性。|
|[WorksheetData](/javascript/api/excel/excel.worksheetdata)|[直方图](/javascript/api/excel/excel.worksheetdata#charts)|返回属于工作表的图表的集合。 只读。|
||[id](/javascript/api/excel/excel.worksheetdata#id)|返回用于唯一标识指定工作簿中工作表的值。即使工作表被重命名或移动，标识符的值仍然相同。只读。|
||[name](/javascript/api/excel/excel.worksheetdata#name)|工作表的显示名称。|
||[position](/javascript/api/excel/excel.worksheetdata#position)|工作表在工作簿中的位置，从零开始。|
||[表](/javascript/api/excel/excel.worksheetdata#tables)|属于工作表的表的集合。 只读。|
||[深入](/javascript/api/excel/excel.worksheetdata#visibility)|工作表的可见性。|
|[WorksheetLoadOptions](/javascript/api/excel/excel.worksheetloadoptions)|[$all](/javascript/api/excel/excel.worksheetloadoptions#$all)||
||[直方图](/javascript/api/excel/excel.worksheetloadoptions#charts)|返回属于工作表的图表的集合。|
||[id](/javascript/api/excel/excel.worksheetloadoptions#id)|返回用于唯一标识指定工作簿中工作表的值。即使工作表被重命名或移动，标识符的值仍然相同。只读。|
||[name](/javascript/api/excel/excel.worksheetloadoptions#name)|工作表的显示名称。|
||[position](/javascript/api/excel/excel.worksheetloadoptions#position)|工作表在工作簿中的位置，从零开始。|
||[表](/javascript/api/excel/excel.worksheetloadoptions#tables)|属于工作表的表的集合。|
||[深入](/javascript/api/excel/excel.worksheetloadoptions#visibility)|工作表的可见性。|
|[WorksheetUpdateData](/javascript/api/excel/excel.worksheetupdatedata)|[name](/javascript/api/excel/excel.worksheetupdatedata#name)|工作表的显示名称。|
||[position](/javascript/api/excel/excel.worksheetupdatedata#position)|工作表在工作簿中的位置，从零开始。|
||[visibility](/javascript/api/excel/excel.worksheetupdatedata#visibility)|工作表的可见性。|

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 参考文档](/javascript/api/excel)
- [Excel JavaScript API 要求集](./excel-api-requirement-sets.md)
