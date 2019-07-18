---
title: Excel JavaScript API 要求集1。7
description: 有关 ExcelApi 1.7 要求集的详细信息
ms.date: 07/11/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: c84d099982225bae11cb3deba8a0503da0695aed
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771986"
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

Excel 事件 API 提供了多个事件处理程序，以便加载项能够在发生特定事件时自动运行指定的函数。 可以将函数设计为执行方案所需的任何操作。 有关当前可用的事件列表，请参阅[使用 Excel JavaScript API 处理事件](/office/dev/add-ins/excel/excel-add-ins-events)。

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

| Class | 域 | 说明 |
|:---|:---|:---|
|[Chart](/javascript/api/excel/excel.chart)|[chartType](/javascript/api/excel/excel.chart#charttype)|表示图表的类型。 有关详细信息, 请参阅 ChartType。|
||[id](/javascript/api/excel/excel.chart#id)|图表的唯一 ID。 只读。|
||[showAllFieldButtons](/javascript/api/excel/excel.chart#showallfieldbuttons)|表示是否在数据透视图上显示所有字段按钮。|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[边缘](/javascript/api/excel/excel.chartareaformat#border)|代表图表区域的边框格式, 包括颜色、linestyle 和粗细。 只读。|
|[ChartAreaFormatData](/javascript/api/excel/excel.chartareaformatdata)|[边缘](/javascript/api/excel/excel.chartareaformatdata#border)|代表图表区域的边框格式, 包括颜色、linestyle 和粗细。 只读。|
|[ChartAreaFormatLoadOptions](/javascript/api/excel/excel.chartareaformatloadoptions)|[边缘](/javascript/api/excel/excel.chartareaformatloadoptions#border)|代表图表区域的边框格式, 包括颜色、linestyle 和粗细。|
|[ChartAreaFormatUpdateData](/javascript/api/excel/excel.chartareaformatupdatedata)|[边缘](/javascript/api/excel/excel.chartareaformatupdatedata#border)|代表图表区域的边框格式, 包括颜色、linestyle 和粗细。|
|[ChartAxes](/javascript/api/excel/excel.chartaxes)|[getItem (类型: \| "' Category" "类别\| " "值\| " "系列", 组？: "Primary \| " "辅助")](/javascript/api/excel/excel.chartaxes#getitem-type--group-)|返回通过类型和组标识的特定轴。|
||[getItem (type: ChartAxisType, group？: ChartAxisGroup)](/javascript/api/excel/excel.chartaxes#getitem-type--group-)|返回通过类型和组标识的特定轴。|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[baseTimeUnit](/javascript/api/excel/excel.chartaxis#basetimeunit)|返回或设置指定分类轴的基本单位。|
||[categoryType](/javascript/api/excel/excel.chartaxis#categorytype)|返回或设置分类轴类型。|
||[displayUnit](/javascript/api/excel/excel.chartaxis#displayunit)|表示轴显示单位。 有关详细信息, 请参阅 ChartAxisDisplayUnit。|
||[logBase](/javascript/api/excel/excel.chartaxis#logbase)|表示使用对数刻度时对数的底数。|
||[majorTickMark](/javascript/api/excel/excel.chartaxis#majortickmark)|表示特定轴的主要刻度线类型。 有关详细信息, 请参阅 ChartAxisTickMark。|
||[majorTimeUnitScale](/javascript/api/excel/excel.chartaxis#majortimeunitscale)|返回或设置当 CategoryType 属性设为 TimeScale 时分类轴的主要单位刻度值。|
||[minorTickMark](/javascript/api/excel/excel.chartaxis#minortickmark)|表示指定轴的次要刻度线类型。 有关详细信息, 请参阅 ChartAxisTickMark。|
||[minorTimeUnitScale](/javascript/api/excel/excel.chartaxis#minortimeunitscale)|返回或设置当 CategoryType 属性设为 TimeScale 时分类轴的次要单位刻度值。|
||[axisGroup](/javascript/api/excel/excel.chartaxis#axisgroup)|返回或设置指定轴的组。 有关详细信息, 请参阅 ChartAxisGroup。 只读。|
||[customDisplayUnit](/javascript/api/excel/excel.chartaxis#customdisplayunit)|标识自定义轴显示单位值。 只读。 要设置此属性，请使用 SetCustomDisplayUnit(double) 方法。|
||[height](/javascript/api/excel/excel.chartaxis#height)|表示图表轴的高度，以磅为单位。 如果轴不可见, 则为 Null。 只读。|
||[left](/javascript/api/excel/excel.chartaxis#left)|表示轴的左边缘到图表区域左侧的距离，以磅为单位。 如果轴不可见, 则为 Null。 只读。|
||[top](/javascript/api/excel/excel.chartaxis#top)|表示轴的上边缘到图表区域顶部的距离，以磅为单位。 如果轴不可见, 则为 Null。 只读。|
||[type](/javascript/api/excel/excel.chartaxis#type)|表示轴类型。 有关详细信息, 请参阅 ChartAxisType。|
||[width](/javascript/api/excel/excel.chartaxis#width)|表示图表轴的宽度，以磅为单位。 如果轴不可见, 则为 Null。 只读。|
||[reversePlotOrder](/javascript/api/excel/excel.chartaxis#reverseplotorder)|表示 Microsoft Excel 是否按照最后一个到第一个的顺序绘制数据点。|
||[scaleType](/javascript/api/excel/excel.chartaxis#scaletype)|表示数值轴刻度类型。 有关详细信息, 请参阅 ChartAxisScaleType。|
||[setCategoryNames (sourceData: Range)](/javascript/api/excel/excel.chartaxis#setcategorynames-sourcedata-)|设置指定轴的所有分类名称。|
||[setCustomDisplayUnit (value: 数字)](/javascript/api/excel/excel.chartaxis#setcustomdisplayunit-value-)|将轴显示单位设为自定义值。|
||[showDisplayUnitLabel](/javascript/api/excel/excel.chartaxis#showdisplayunitlabel)|表示轴显示单位标签是否可见。|
||[tickLabelPosition](/javascript/api/excel/excel.chartaxis#ticklabelposition)|表示特定轴上的刻度线标签位置。 有关详细信息, 请参阅 ChartAxisTickLabelPosition。|
||[tickLabelSpacing](/javascript/api/excel/excel.chartaxis#ticklabelspacing)|表示刻度线标签之间的分类或系列数。 可以是 1 到 31999 的值或空字符串（自动设置）。 返回的值始终为数字。|
||[tickMarkSpacing](/javascript/api/excel/excel.chartaxis#tickmarkspacing)|表示刻度线之间的分类或系列数。|
||[visible](/javascript/api/excel/excel.chartaxis#visible)|该布尔值表示轴的可见性。|
|[ChartAxisData](/javascript/api/excel/excel.chartaxisdata)|[axisGroup](/javascript/api/excel/excel.chartaxisdata#axisgroup)|返回或设置指定轴的组。 有关详细信息, 请参阅 ChartAxisGroup。 只读。|
||[baseTimeUnit](/javascript/api/excel/excel.chartaxisdata#basetimeunit)|返回或设置指定分类轴的基本单位。|
||[categoryType](/javascript/api/excel/excel.chartaxisdata#categorytype)|返回或设置分类轴类型。|
||[customDisplayUnit](/javascript/api/excel/excel.chartaxisdata#customdisplayunit)|标识自定义轴显示单位值。 只读。 要设置此属性，请使用 SetCustomDisplayUnit(double) 方法。|
||[displayUnit](/javascript/api/excel/excel.chartaxisdata#displayunit)|表示轴显示单位。 有关详细信息, 请参阅 ChartAxisDisplayUnit。|
||[height](/javascript/api/excel/excel.chartaxisdata#height)|表示图表轴的高度，以磅为单位。 如果轴不可见, 则为 Null。 只读。|
||[left](/javascript/api/excel/excel.chartaxisdata#left)|表示轴的左边缘到图表区域左侧的距离，以磅为单位。 如果轴不可见, 则为 Null。 只读。|
||[logBase](/javascript/api/excel/excel.chartaxisdata#logbase)|表示使用对数刻度时对数的底数。|
||[majorTickMark](/javascript/api/excel/excel.chartaxisdata#majortickmark)|表示特定轴的主要刻度线类型。 有关详细信息, 请参阅 ChartAxisTickMark。|
||[majorTimeUnitScale](/javascript/api/excel/excel.chartaxisdata#majortimeunitscale)|返回或设置当 CategoryType 属性设为 TimeScale 时分类轴的主要单位刻度值。|
||[minorTickMark](/javascript/api/excel/excel.chartaxisdata#minortickmark)|表示指定轴的次要刻度线类型。 有关详细信息, 请参阅 ChartAxisTickMark。|
||[minorTimeUnitScale](/javascript/api/excel/excel.chartaxisdata#minortimeunitscale)|返回或设置当 CategoryType 属性设为 TimeScale 时分类轴的次要单位刻度值。|
||[reversePlotOrder](/javascript/api/excel/excel.chartaxisdata#reverseplotorder)|表示 Microsoft Excel 是否按照最后一个到第一个的顺序绘制数据点。|
||[scaleType](/javascript/api/excel/excel.chartaxisdata#scaletype)|表示数值轴刻度类型。 有关详细信息, 请参阅 ChartAxisScaleType。|
||[showDisplayUnitLabel](/javascript/api/excel/excel.chartaxisdata#showdisplayunitlabel)|表示轴显示单位标签是否可见。|
||[tickLabelPosition](/javascript/api/excel/excel.chartaxisdata#ticklabelposition)|表示特定轴上的刻度线标签位置。 有关详细信息, 请参阅 ChartAxisTickLabelPosition。|
||[tickLabelSpacing](/javascript/api/excel/excel.chartaxisdata#ticklabelspacing)|表示刻度线标签之间的分类或系列数。 可以是 1 到 31999 的值或空字符串（自动设置）。 返回的值始终为数字。|
||[tickMarkSpacing](/javascript/api/excel/excel.chartaxisdata#tickmarkspacing)|表示刻度线之间的分类或系列数。|
||[top](/javascript/api/excel/excel.chartaxisdata#top)|表示轴的上边缘到图表区域顶部的距离，以磅为单位。 如果轴不可见, 则为 Null。 只读。|
||[type](/javascript/api/excel/excel.chartaxisdata#type)|表示轴类型。 有关详细信息, 请参阅 ChartAxisType。|
||[visible](/javascript/api/excel/excel.chartaxisdata#visible)|该布尔值表示轴的可见性。|
||[width](/javascript/api/excel/excel.chartaxisdata#width)|表示图表轴的宽度，以磅为单位。 如果轴不可见, 则为 Null。 只读。|
|[ChartAxisLoadOptions](/javascript/api/excel/excel.chartaxisloadoptions)|[axisGroup](/javascript/api/excel/excel.chartaxisloadoptions#axisgroup)|返回或设置指定轴的组。 有关详细信息, 请参阅 ChartAxisGroup。 只读。|
||[baseTimeUnit](/javascript/api/excel/excel.chartaxisloadoptions#basetimeunit)|返回或设置指定分类轴的基本单位。|
||[categoryType](/javascript/api/excel/excel.chartaxisloadoptions#categorytype)|返回或设置分类轴类型。|
||[跨过](/javascript/api/excel/excel.chartaxisloadoptions#crosses)|[已弃用; 保留用于与现有第一方解决方案的后端兼容]。 请`Position`改用。|
||[crossesAt](/javascript/api/excel/excel.chartaxisloadoptions#crossesat)|[已弃用; 保留用于与现有第一方解决方案的后端兼容]。 请`PositionAt`改用。|
||[customDisplayUnit](/javascript/api/excel/excel.chartaxisloadoptions#customdisplayunit)|标识自定义轴显示单位值。 只读。 要设置此属性，请使用 SetCustomDisplayUnit(double) 方法。|
||[displayUnit](/javascript/api/excel/excel.chartaxisloadoptions#displayunit)|表示轴显示单位。 有关详细信息, 请参阅 ChartAxisDisplayUnit。|
||[height](/javascript/api/excel/excel.chartaxisloadoptions#height)|表示图表轴的高度，以磅为单位。 如果轴不可见, 则为 Null。 只读。|
||[left](/javascript/api/excel/excel.chartaxisloadoptions#left)|表示轴的左边缘到图表区域左侧的距离，以磅为单位。 如果轴不可见, 则为 Null。 只读。|
||[logBase](/javascript/api/excel/excel.chartaxisloadoptions#logbase)|表示使用对数刻度时对数的底数。|
||[majorTickMark](/javascript/api/excel/excel.chartaxisloadoptions#majortickmark)|表示特定轴的主要刻度线类型。 有关详细信息, 请参阅 ChartAxisTickMark。|
||[majorTimeUnitScale](/javascript/api/excel/excel.chartaxisloadoptions#majortimeunitscale)|返回或设置当 CategoryType 属性设为 TimeScale 时分类轴的主要单位刻度值。|
||[minorTickMark](/javascript/api/excel/excel.chartaxisloadoptions#minortickmark)|表示指定轴的次要刻度线类型。 有关详细信息, 请参阅 ChartAxisTickMark。|
||[minorTimeUnitScale](/javascript/api/excel/excel.chartaxisloadoptions#minortimeunitscale)|返回或设置当 CategoryType 属性设为 TimeScale 时分类轴的次要单位刻度值。|
||[reversePlotOrder](/javascript/api/excel/excel.chartaxisloadoptions#reverseplotorder)|表示 Microsoft Excel 是否按照最后一个到第一个的顺序绘制数据点。|
||[scaleType](/javascript/api/excel/excel.chartaxisloadoptions#scaletype)|表示数值轴刻度类型。 有关详细信息, 请参阅 ChartAxisScaleType。|
||[showDisplayUnitLabel](/javascript/api/excel/excel.chartaxisloadoptions#showdisplayunitlabel)|表示轴显示单位标签是否可见。|
||[tickLabelPosition](/javascript/api/excel/excel.chartaxisloadoptions#ticklabelposition)|表示特定轴上的刻度线标签位置。 有关详细信息, 请参阅 ChartAxisTickLabelPosition。|
||[tickLabelSpacing](/javascript/api/excel/excel.chartaxisloadoptions#ticklabelspacing)|表示刻度线标签之间的分类或系列数。 可以是 1 到 31999 的值或空字符串（自动设置）。 返回的值始终为数字。|
||[tickMarkSpacing](/javascript/api/excel/excel.chartaxisloadoptions#tickmarkspacing)|表示刻度线之间的分类或系列数。|
||[top](/javascript/api/excel/excel.chartaxisloadoptions#top)|表示轴的上边缘到图表区域顶部的距离，以磅为单位。 如果轴不可见, 则为 Null。 只读。|
||[type](/javascript/api/excel/excel.chartaxisloadoptions#type)|表示轴类型。 有关详细信息, 请参阅 ChartAxisType。|
||[visible](/javascript/api/excel/excel.chartaxisloadoptions#visible)|该布尔值表示轴的可见性。|
||[width](/javascript/api/excel/excel.chartaxisloadoptions#width)|表示图表轴的宽度，以磅为单位。 如果轴不可见, 则为 Null。 只读。|
|[ChartAxisUpdateData](/javascript/api/excel/excel.chartaxisupdatedata)|[baseTimeUnit](/javascript/api/excel/excel.chartaxisupdatedata#basetimeunit)|返回或设置指定分类轴的基本单位。|
||[categoryType](/javascript/api/excel/excel.chartaxisupdatedata#categorytype)|返回或设置分类轴类型。|
||[displayUnit](/javascript/api/excel/excel.chartaxisupdatedata#displayunit)|表示轴显示单位。 有关详细信息, 请参阅 ChartAxisDisplayUnit。|
||[logBase](/javascript/api/excel/excel.chartaxisupdatedata#logbase)|表示使用对数刻度时对数的底数。|
||[majorTickMark](/javascript/api/excel/excel.chartaxisupdatedata#majortickmark)|表示特定轴的主要刻度线类型。 有关详细信息, 请参阅 ChartAxisTickMark。|
||[majorTimeUnitScale](/javascript/api/excel/excel.chartaxisupdatedata#majortimeunitscale)|返回或设置当 CategoryType 属性设为 TimeScale 时分类轴的主要单位刻度值。|
||[minorTickMark](/javascript/api/excel/excel.chartaxisupdatedata#minortickmark)|表示指定轴的次要刻度线类型。 有关详细信息, 请参阅 ChartAxisTickMark。|
||[minorTimeUnitScale](/javascript/api/excel/excel.chartaxisupdatedata#minortimeunitscale)|返回或设置当 CategoryType 属性设为 TimeScale 时分类轴的次要单位刻度值。|
||[reversePlotOrder](/javascript/api/excel/excel.chartaxisupdatedata#reverseplotorder)|表示 Microsoft Excel 是否按照最后一个到第一个的顺序绘制数据点。|
||[scaleType](/javascript/api/excel/excel.chartaxisupdatedata#scaletype)|表示数值轴刻度类型。 有关详细信息, 请参阅 ChartAxisScaleType。|
||[showDisplayUnitLabel](/javascript/api/excel/excel.chartaxisupdatedata#showdisplayunitlabel)|表示轴显示单位标签是否可见。|
||[tickLabelPosition](/javascript/api/excel/excel.chartaxisupdatedata#ticklabelposition)|表示特定轴上的刻度线标签位置。 有关详细信息, 请参阅 ChartAxisTickLabelPosition。|
||[tickLabelSpacing](/javascript/api/excel/excel.chartaxisupdatedata#ticklabelspacing)|表示刻度线标签之间的分类或系列数。 可以是 1 到 31999 的值或空字符串（自动设置）。 返回的值始终为数字。|
||[tickMarkSpacing](/javascript/api/excel/excel.chartaxisupdatedata#tickmarkspacing)|表示刻度线之间的分类或系列数。|
||[visible](/javascript/api/excel/excel.chartaxisupdatedata#visible)|该布尔值表示轴的可见性。|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[color](/javascript/api/excel/excel.chartborder#color)|表示图表中的边框颜色的 HTML 颜色代码。|
||[lineStyle](/javascript/api/excel/excel.chartborder#linestyle)|表示边框的线条样式。 有关详细信息, 请参阅 ChartLineStyle。|
||[set (properties: ChartBorder)](/javascript/api/excel/excel.chartborder#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: ChartBorderUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.chartborder#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
||[weight](/javascript/api/excel/excel.chartborder#weight)|表示边框的粗细，以磅为单位。|
|[ChartBorderData](/javascript/api/excel/excel.chartborderdata)|[color](/javascript/api/excel/excel.chartborderdata#color)|表示图表中的边框颜色的 HTML 颜色代码。|
||[lineStyle](/javascript/api/excel/excel.chartborderdata#linestyle)|表示边框的线条样式。 有关详细信息, 请参阅 ChartLineStyle。|
||[weight](/javascript/api/excel/excel.chartborderdata#weight)|表示边框的粗细，以磅为单位。|
|[ChartBorderLoadOptions](/javascript/api/excel/excel.chartborderloadoptions)|[$all](/javascript/api/excel/excel.chartborderloadoptions#$all)||
||[color](/javascript/api/excel/excel.chartborderloadoptions#color)|表示图表中的边框颜色的 HTML 颜色代码。|
||[lineStyle](/javascript/api/excel/excel.chartborderloadoptions#linestyle)|表示边框的线条样式。 有关详细信息, 请参阅 ChartLineStyle。|
||[weight](/javascript/api/excel/excel.chartborderloadoptions#weight)|表示边框的粗细，以磅为单位。|
|[ChartBorderUpdateData](/javascript/api/excel/excel.chartborderupdatedata)|[color](/javascript/api/excel/excel.chartborderupdatedata#color)|表示图表中的边框颜色的 HTML 颜色代码。|
||[lineStyle](/javascript/api/excel/excel.chartborderupdatedata#linestyle)|表示边框的线条样式。 有关详细信息, 请参阅 ChartLineStyle。|
||[weight](/javascript/api/excel/excel.chartborderupdatedata#weight)|表示边框的粗细，以磅为单位。|
|[ChartCollectionLoadOptions](/javascript/api/excel/excel.chartcollectionloadoptions)|[chartType](/javascript/api/excel/excel.chartcollectionloadoptions#charttype)|对于集合中的每一项: 代表图表的类型。 有关详细信息, 请参阅 ChartType。|
||[id](/javascript/api/excel/excel.chartcollectionloadoptions#id)|对于集合中的每一项: 图表的唯一 id。 只读。|
||[showAllFieldButtons](/javascript/api/excel/excel.chartcollectionloadoptions#showallfieldbuttons)|对于集合中的每一项: 表示是否在数据透视图上显示所有字段按钮。|
|[ChartData](/javascript/api/excel/excel.chartdata)|[chartType](/javascript/api/excel/excel.chartdata#charttype)|表示图表的类型。 有关详细信息, 请参阅 ChartType。|
||[id](/javascript/api/excel/excel.chartdata#id)|图表的唯一 ID。 只读。|
||[showAllFieldButtons](/javascript/api/excel/excel.chartdata#showallfieldbuttons)|表示是否在数据透视图上显示所有字段按钮。|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[position](/javascript/api/excel/excel.chartdatalabel#position)|表示数据标签的位置的 DataLabelPosition 值。 有关详细信息, 请参阅 ChartDataLabelPosition。|
||[分隔符](/javascript/api/excel/excel.chartdatalabel#separator)|该字符串表示用于图表中数据标签的分隔符。|
||[set (properties: ChartDataLabel)](/javascript/api/excel/excel.chartdatalabel#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: ChartDataLabelUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.chartdatalabel#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabel#showbubblesize)|该布尔值表示数据标签气泡大小是否可见。|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabel#showcategoryname)|表示数据标签分类名称是否可见的布尔值。|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabel#showlegendkey)|该布尔值表示数据标签图例标示是否可见。|
||[showPercentage](/javascript/api/excel/excel.chartdatalabel#showpercentage)|该布尔值表示数据标签百分比是否可见。|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabel#showseriesname)|该布尔值表示数据标签系列名称是否可见。|
||[showValue](/javascript/api/excel/excel.chartdatalabel#showvalue)|该布尔值表示数据标签值是否可见。|
|[ChartDataLabelData](/javascript/api/excel/excel.chartdatalabeldata)|[position](/javascript/api/excel/excel.chartdatalabeldata#position)|表示数据标签的位置的 DataLabelPosition 值。 有关详细信息, 请参阅 ChartDataLabelPosition。|
||[分隔符](/javascript/api/excel/excel.chartdatalabeldata#separator)|该字符串表示用于图表中数据标签的分隔符。|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabeldata#showbubblesize)|该布尔值表示数据标签气泡大小是否可见。|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabeldata#showcategoryname)|表示数据标签分类名称是否可见的布尔值。|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabeldata#showlegendkey)|该布尔值表示数据标签图例标示是否可见。|
||[showPercentage](/javascript/api/excel/excel.chartdatalabeldata#showpercentage)|该布尔值表示数据标签百分比是否可见。|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabeldata#showseriesname)|该布尔值表示数据标签系列名称是否可见。|
||[showValue](/javascript/api/excel/excel.chartdatalabeldata#showvalue)|该布尔值表示数据标签值是否可见。|
|[ChartDataLabelLoadOptions](/javascript/api/excel/excel.chartdatalabelloadoptions)|[$all](/javascript/api/excel/excel.chartdatalabelloadoptions#$all)||
||[position](/javascript/api/excel/excel.chartdatalabelloadoptions#position)|表示数据标签的位置的 DataLabelPosition 值。 有关详细信息, 请参阅 ChartDataLabelPosition。|
||[分隔符](/javascript/api/excel/excel.chartdatalabelloadoptions#separator)|该字符串表示用于图表中数据标签的分隔符。|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabelloadoptions#showbubblesize)|该布尔值表示数据标签气泡大小是否可见。|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabelloadoptions#showcategoryname)|表示数据标签分类名称是否可见的布尔值。|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabelloadoptions#showlegendkey)|该布尔值表示数据标签图例标示是否可见。|
||[showPercentage](/javascript/api/excel/excel.chartdatalabelloadoptions#showpercentage)|该布尔值表示数据标签百分比是否可见。|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabelloadoptions#showseriesname)|该布尔值表示数据标签系列名称是否可见。|
||[showValue](/javascript/api/excel/excel.chartdatalabelloadoptions#showvalue)|该布尔值表示数据标签值是否可见。|
|[ChartDataLabelUpdateData](/javascript/api/excel/excel.chartdatalabelupdatedata)|[position](/javascript/api/excel/excel.chartdatalabelupdatedata#position)|表示数据标签的位置的 DataLabelPosition 值。 有关详细信息, 请参阅 ChartDataLabelPosition。|
||[分隔符](/javascript/api/excel/excel.chartdatalabelupdatedata#separator)|该字符串表示用于图表中数据标签的分隔符。|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabelupdatedata#showbubblesize)|该布尔值表示数据标签气泡大小是否可见。|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabelupdatedata#showcategoryname)|表示数据标签分类名称是否可见的布尔值。|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabelupdatedata#showlegendkey)|该布尔值表示数据标签图例标示是否可见。|
||[showPercentage](/javascript/api/excel/excel.chartdatalabelupdatedata#showpercentage)|该布尔值表示数据标签百分比是否可见。|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabelupdatedata#showseriesname)|该布尔值表示数据标签系列名称是否可见。|
||[showValue](/javascript/api/excel/excel.chartdatalabelupdatedata#showvalue)|该布尔值表示数据标签值是否可见。|
|[ChartFormatString](/javascript/api/excel/excel.chartformatstring)|[font](/javascript/api/excel/excel.chartformatstring#font)|代表 "图表字符" 对象的字体属性, 如字体名称、字体大小、颜色等。|
||[set (properties: ChartFormatString)](/javascript/api/excel/excel.chartformatstring#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: ChartFormatStringUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.chartformatstring#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
|[ChartFormatStringData](/javascript/api/excel/excel.chartformatstringdata)|[font](/javascript/api/excel/excel.chartformatstringdata#font)|代表 "图表字符" 对象的字体属性, 如字体名称、字体大小、颜色等。|
|[ChartFormatStringLoadOptions](/javascript/api/excel/excel.chartformatstringloadoptions)|[$all](/javascript/api/excel/excel.chartformatstringloadoptions#$all)||
||[font](/javascript/api/excel/excel.chartformatstringloadoptions#font)|代表 "图表字符" 对象的字体属性, 如字体名称、字体大小、颜色等。|
|[ChartFormatStringUpdateData](/javascript/api/excel/excel.chartformatstringupdatedata)|[font](/javascript/api/excel/excel.chartformatstringupdatedata#font)|代表 "图表字符" 对象的字体属性, 如字体名称、字体大小、颜色等。|
|[ChartLegend](/javascript/api/excel/excel.chartlegend)|[height](/javascript/api/excel/excel.chartlegend#height)|表示图表上的图例的高度 (以磅为单位)。 如果图例不可见, 则为 Null。|
||[left](/javascript/api/excel/excel.chartlegend#left)|代表图表图例的左侧 (以磅为单位)。 如果图例不可见, 则为 Null。|
||[legendEntries](/javascript/api/excel/excel.chartlegend#legendentries)|表示图例中 legendEntries 的集合。 只读。|
||[showShadow](/javascript/api/excel/excel.chartlegend#showshadow)|表示图例在图表上是否有阴影。|
||[top](/javascript/api/excel/excel.chartlegend#top)|表示图表图例顶部。|
||[width](/javascript/api/excel/excel.chartlegend#width)|表示图表上的图例的宽度 (以磅为单位)。 如果图例不可见, 则为 Null。|
|[ChartLegendData](/javascript/api/excel/excel.chartlegenddata)|[height](/javascript/api/excel/excel.chartlegenddata#height)|表示图表上的图例的高度 (以磅为单位)。 如果图例不可见, 则为 Null。|
||[left](/javascript/api/excel/excel.chartlegenddata#left)|代表图表图例的左侧 (以磅为单位)。 如果图例不可见, 则为 Null。|
||[legendEntries](/javascript/api/excel/excel.chartlegenddata#legendentries)|表示图例中 legendEntries 的集合。 只读。|
||[showShadow](/javascript/api/excel/excel.chartlegenddata#showshadow)|表示图例在图表上是否有阴影。|
||[top](/javascript/api/excel/excel.chartlegenddata#top)|表示图表图例顶部。|
||[width](/javascript/api/excel/excel.chartlegenddata#width)|表示图表上的图例的宽度 (以磅为单位)。 如果图例不可见, 则为 Null。|
|[ChartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|[set (properties: ChartLegendEntry)](/javascript/api/excel/excel.chartlegendentry#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: ChartLegendEntryUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.chartlegendentry#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
||[visible](/javascript/api/excel/excel.chartlegendentry#visible)|表示图表图例条目可见。|
|[ChartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|[getCount()](/javascript/api/excel/excel.chartlegendentrycollection#getcount--)|返回集合中的 legendEntry 数量。|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartlegendentrycollection#getitemat-index-)|返回给定索引处的 legendEntry。|
||[items](/javascript/api/excel/excel.chartlegendentrycollection#items)|获取此集合中已加载的子项。|
|[ChartLegendEntryCollectionLoadOptions](/javascript/api/excel/excel.chartlegendentrycollectionloadoptions)|[$all](/javascript/api/excel/excel.chartlegendentrycollectionloadoptions#$all)||
||[visible](/javascript/api/excel/excel.chartlegendentrycollectionloadoptions#visible)|对于集合中的每一项: 表示图表图例项的可见项。|
|[ChartLegendEntryData](/javascript/api/excel/excel.chartlegendentrydata)|[visible](/javascript/api/excel/excel.chartlegendentrydata#visible)|表示图表图例条目可见。|
|[ChartLegendEntryLoadOptions](/javascript/api/excel/excel.chartlegendentryloadoptions)|[$all](/javascript/api/excel/excel.chartlegendentryloadoptions#$all)||
||[visible](/javascript/api/excel/excel.chartlegendentryloadoptions#visible)|表示图表图例条目可见。|
|[ChartLegendEntryUpdateData](/javascript/api/excel/excel.chartlegendentryupdatedata)|[visible](/javascript/api/excel/excel.chartlegendentryupdatedata#visible)|表示图表图例条目可见。|
|[ChartLegendLoadOptions](/javascript/api/excel/excel.chartlegendloadoptions)|[height](/javascript/api/excel/excel.chartlegendloadoptions#height)|表示图表上的图例的高度 (以磅为单位)。 如果图例不可见, 则为 Null。|
||[left](/javascript/api/excel/excel.chartlegendloadoptions#left)|代表图表图例的左侧 (以磅为单位)。 如果图例不可见, 则为 Null。|
||[showShadow](/javascript/api/excel/excel.chartlegendloadoptions#showshadow)|表示图例在图表上是否有阴影。|
||[top](/javascript/api/excel/excel.chartlegendloadoptions#top)|表示图表图例顶部。|
||[width](/javascript/api/excel/excel.chartlegendloadoptions#width)|表示图表上的图例的宽度 (以磅为单位)。 如果图例不可见, 则为 Null。|
|[ChartLegendUpdateData](/javascript/api/excel/excel.chartlegendupdatedata)|[height](/javascript/api/excel/excel.chartlegendupdatedata#height)|表示图表上的图例的高度 (以磅为单位)。 如果图例不可见, 则为 Null。|
||[left](/javascript/api/excel/excel.chartlegendupdatedata#left)|代表图表图例的左侧 (以磅为单位)。 如果图例不可见, 则为 Null。|
||[showShadow](/javascript/api/excel/excel.chartlegendupdatedata#showshadow)|表示图例在图表上是否有阴影。|
||[top](/javascript/api/excel/excel.chartlegendupdatedata#top)|表示图表图例顶部。|
||[width](/javascript/api/excel/excel.chartlegendupdatedata#width)|表示图表上的图例的宽度 (以磅为单位)。 如果图例不可见, 则为 Null。|
|[ChartLineFormat](/javascript/api/excel/excel.chartlineformat)|[lineStyle](/javascript/api/excel/excel.chartlineformat#linestyle)|代表线条样式。 有关详细信息, 请参阅 ChartLineStyle。|
||[weight](/javascript/api/excel/excel.chartlineformat#weight)|表示线条的粗细（以磅为单位）。|
|[ChartLineFormatData](/javascript/api/excel/excel.chartlineformatdata)|[lineStyle](/javascript/api/excel/excel.chartlineformatdata#linestyle)|代表线条样式。 有关详细信息, 请参阅 ChartLineStyle。|
||[weight](/javascript/api/excel/excel.chartlineformatdata#weight)|表示线条的粗细（以磅为单位）。|
|[ChartLineFormatLoadOptions](/javascript/api/excel/excel.chartlineformatloadoptions)|[lineStyle](/javascript/api/excel/excel.chartlineformatloadoptions#linestyle)|代表线条样式。 有关详细信息, 请参阅 ChartLineStyle。|
||[weight](/javascript/api/excel/excel.chartlineformatloadoptions#weight)|表示线条的粗细（以磅为单位）。|
|[ChartLineFormatUpdateData](/javascript/api/excel/excel.chartlineformatupdatedata)|[lineStyle](/javascript/api/excel/excel.chartlineformatupdatedata#linestyle)|代表线条样式。 有关详细信息, 请参阅 ChartLineStyle。|
||[weight](/javascript/api/excel/excel.chartlineformatupdatedata#weight)|表示线条的粗细（以磅为单位）。|
|[ChartLoadOptions](/javascript/api/excel/excel.chartloadoptions)|[chartType](/javascript/api/excel/excel.chartloadoptions#charttype)|表示图表的类型。 有关详细信息, 请参阅 ChartType。|
||[id](/javascript/api/excel/excel.chartloadoptions#id)|图表的唯一 ID。 只读。|
||[showAllFieldButtons](/javascript/api/excel/excel.chartloadoptions#showallfieldbuttons)|表示是否在数据透视图上显示所有字段按钮。|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|[hasDataLabel](/javascript/api/excel/excel.chartpoint#hasdatalabel)|表示数据点是否具有数据标签。 不适用于曲面图。|
||[markerBackgroundColor](/javascript/api/excel/excel.chartpoint#markerbackgroundcolor)|表示数据点的标记背景色的 HTML 颜色代码。 例如， #FF0000 代表红色。|
||[markerForegroundColor](/javascript/api/excel/excel.chartpoint#markerforegroundcolor)|表示数据点的标记前景色的 HTML 颜色代码。 例如， #FF0000 代表红色。|
||[markerSize](/javascript/api/excel/excel.chartpoint#markersize)|表示数据点的标记大小。|
||[markerStyle](/javascript/api/excel/excel.chartpoint#markerstyle)|表示图表数据点的标记样式。 有关详细信息, 请参阅 ChartMarkerStyle。|
||[dataLabel](/javascript/api/excel/excel.chartpoint#datalabel)|返回图表点的数据标签。 只读。|
|[ChartPointData](/javascript/api/excel/excel.chartpointdata)|[dataLabel](/javascript/api/excel/excel.chartpointdata#datalabel)|返回图表点的数据标签。 只读。|
||[hasDataLabel](/javascript/api/excel/excel.chartpointdata#hasdatalabel)|表示数据点是否具有数据标签。 不适用于曲面图。|
||[markerBackgroundColor](/javascript/api/excel/excel.chartpointdata#markerbackgroundcolor)|表示数据点的标记背景色的 HTML 颜色代码。 例如， #FF0000 代表红色。|
||[markerForegroundColor](/javascript/api/excel/excel.chartpointdata#markerforegroundcolor)|表示数据点的标记前景色的 HTML 颜色代码。 例如， #FF0000 代表红色。|
||[markerSize](/javascript/api/excel/excel.chartpointdata#markersize)|表示数据点的标记大小。|
||[markerStyle](/javascript/api/excel/excel.chartpointdata#markerstyle)|表示图表数据点的标记样式。 有关详细信息, 请参阅 ChartMarkerStyle。|
|[ChartPointFormat](/javascript/api/excel/excel.chartpointformat)|[边缘](/javascript/api/excel/excel.chartpointformat#border)|表示图表数据点的边框格式, 包括颜色、样式和权重信息。 只读。|
|[ChartPointFormatData](/javascript/api/excel/excel.chartpointformatdata)|[边缘](/javascript/api/excel/excel.chartpointformatdata#border)|表示图表数据点的边框格式, 包括颜色、样式和权重信息。 只读。|
|[ChartPointFormatLoadOptions](/javascript/api/excel/excel.chartpointformatloadoptions)|[边缘](/javascript/api/excel/excel.chartpointformatloadoptions#border)|表示图表数据点的边框格式, 包括颜色、样式和权重信息。|
|[ChartPointFormatUpdateData](/javascript/api/excel/excel.chartpointformatupdatedata)|[边缘](/javascript/api/excel/excel.chartpointformatupdatedata#border)|表示图表数据点的边框格式, 包括颜色、样式和权重信息。|
|[ChartPointLoadOptions](/javascript/api/excel/excel.chartpointloadoptions)|[dataLabel](/javascript/api/excel/excel.chartpointloadoptions#datalabel)|返回图表点的数据标签。|
||[hasDataLabel](/javascript/api/excel/excel.chartpointloadoptions#hasdatalabel)|表示数据点是否具有数据标签。 不适用于曲面图。|
||[markerBackgroundColor](/javascript/api/excel/excel.chartpointloadoptions#markerbackgroundcolor)|表示数据点的标记背景色的 HTML 颜色代码。 例如， #FF0000 代表红色。|
||[markerForegroundColor](/javascript/api/excel/excel.chartpointloadoptions#markerforegroundcolor)|表示数据点的标记前景色的 HTML 颜色代码。 例如， #FF0000 代表红色。|
||[markerSize](/javascript/api/excel/excel.chartpointloadoptions#markersize)|表示数据点的标记大小。|
||[markerStyle](/javascript/api/excel/excel.chartpointloadoptions#markerstyle)|表示图表数据点的标记样式。 有关详细信息, 请参阅 ChartMarkerStyle。|
|[ChartPointUpdateData](/javascript/api/excel/excel.chartpointupdatedata)|[dataLabel](/javascript/api/excel/excel.chartpointupdatedata#datalabel)|返回图表点的数据标签。|
||[hasDataLabel](/javascript/api/excel/excel.chartpointupdatedata#hasdatalabel)|表示数据点是否具有数据标签。 不适用于曲面图。|
||[markerBackgroundColor](/javascript/api/excel/excel.chartpointupdatedata#markerbackgroundcolor)|表示数据点的标记背景色的 HTML 颜色代码。 例如， #FF0000 代表红色。|
||[markerForegroundColor](/javascript/api/excel/excel.chartpointupdatedata#markerforegroundcolor)|表示数据点的标记前景色的 HTML 颜色代码。 例如， #FF0000 代表红色。|
||[markerSize](/javascript/api/excel/excel.chartpointupdatedata#markersize)|表示数据点的标记大小。|
||[markerStyle](/javascript/api/excel/excel.chartpointupdatedata#markerstyle)|表示图表数据点的标记样式。 有关详细信息, 请参阅 ChartMarkerStyle。|
|[ChartPointsCollectionLoadOptions](/javascript/api/excel/excel.chartpointscollectionloadoptions)|[dataLabel](/javascript/api/excel/excel.chartpointscollectionloadoptions#datalabel)|对于集合中的每一项: 返回图表点的数据标签。|
||[hasDataLabel](/javascript/api/excel/excel.chartpointscollectionloadoptions#hasdatalabel)|对于集合中的每一项: 表示数据点是否具有数据标签。 不适用于曲面图。|
||[markerBackgroundColor](/javascript/api/excel/excel.chartpointscollectionloadoptions#markerbackgroundcolor)|对于集合中的每一项: 数据点的标记背景色的 HTML 颜色代码表示形式。 例如， #FF0000 代表红色。|
||[markerForegroundColor](/javascript/api/excel/excel.chartpointscollectionloadoptions#markerforegroundcolor)|对于集合中的每一项: 数据点的标记前景色的 HTML 颜色代码表示形式。 例如， #FF0000 代表红色。|
||[markerSize](/javascript/api/excel/excel.chartpointscollectionloadoptions#markersize)|对于集合中的每一项: 表示数据点的标记大小。|
||[markerStyle](/javascript/api/excel/excel.chartpointscollectionloadoptions#markerstyle)|对于集合中的每一项: 表示图表数据点的标记样式。 有关详细信息, 请参阅 ChartMarkerStyle。|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[chartType](/javascript/api/excel/excel.chartseries#charttype)|表示系列的图表类型。 有关详细信息, 请参阅 ChartType。|
||[delete()](/javascript/api/excel/excel.chartseries#delete--)|删除 chart series 对象。|
||[doughnutHoleSize](/javascript/api/excel/excel.chartseries#doughnutholesize)|表示图表系列的圆环孔大小。  仅对圆环图和分离型圆环图有效。|
||[筛选](/javascript/api/excel/excel.chartseries#filtered)|该布尔值表示是否筛选系列。 不适用于曲面图。|
||[gapWidth](/javascript/api/excel/excel.chartseries#gapwidth)|表示图表系列的间隙宽度。  有效对象：条形图和柱形图，以及|
||[hasDataLabels](/javascript/api/excel/excel.chartseries#hasdatalabels)|该布尔值表示系列是否具有数据标签。|
||[markerBackgroundColor](/javascript/api/excel/excel.chartseries#markerbackgroundcolor)|表示图表系列的标记背景色。|
||[markerForegroundColor](/javascript/api/excel/excel.chartseries#markerforegroundcolor)|表示图表系列的标记前景色。|
||[markerSize](/javascript/api/excel/excel.chartseries#markersize)|表示图表系列的标记大小。|
||[markerStyle](/javascript/api/excel/excel.chartseries#markerstyle)|表示图表系列的标记类型。 有关详细信息, 请参阅 ChartMarkerStyle。|
||[plotOrder](/javascript/api/excel/excel.chartseries#plotorder)|表示图表组中某个图表系列的绘制顺序。|
||[趋势](/javascript/api/excel/excel.chartseries#trendlines)|表示系列中趋势线的集合。 只读。|
||[setBubbleSizes (sourceData: Range)](/javascript/api/excel/excel.chartseries#setbubblesizes-sourcedata-)|设置图表系列的气泡大小。 仅适用于气泡图。|
||[setValues (sourceData: Range)](/javascript/api/excel/excel.chartseries#setvalues-sourcedata-)|设置图表系列的值。 对于散点图，它表示 Y 轴的值。|
||[setXAxisValues (sourceData: Range)](/javascript/api/excel/excel.chartseries#setxaxisvalues-sourcedata-)|设置图表系列 X 轴的值。 仅适用于散点图。|
||[showShadow](/javascript/api/excel/excel.chartseries#showshadow)|表示系列是否具有阴影的布尔值。|
||[平滑](/javascript/api/excel/excel.chartseries#smooth)|该布尔值表示系列是否平滑。 仅适用于折线图和散点图。|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[add (name？: string, index？: number)](/javascript/api/excel/excel.chartseriescollection#add-name--index-)|向集合添加新系列。 在设置值/x 轴值/气泡大小 (具体取决于图表类型) 之前, 新添加的系列将不可见。|
|[ChartSeriesCollectionLoadOptions](/javascript/api/excel/excel.chartseriescollectionloadoptions)|[chartType](/javascript/api/excel/excel.chartseriescollectionloadoptions#charttype)|对于集合中的每一项: 代表系列的图表类型。 有关详细信息, 请参阅 ChartType。|
||[doughnutHoleSize](/javascript/api/excel/excel.chartseriescollectionloadoptions#doughnutholesize)|对于集合中的每一项: 表示图表系列的圆环图内径大小。  仅对圆环图和分离型圆环图有效。|
||[筛选](/javascript/api/excel/excel.chartseriescollectionloadoptions#filtered)|对于集合中的每一项: 布尔值, 表示是否筛选系列。 不适用于曲面图。|
||[gapWidth](/javascript/api/excel/excel.chartseriescollectionloadoptions#gapwidth)|对于集合中的每一项: 表示图表系列的间隙宽度。  有效对象：条形图和柱形图，以及|
||[hasDataLabels](/javascript/api/excel/excel.chartseriescollectionloadoptions#hasdatalabels)|对于集合中的每一项: 布尔值, 表示该系列是否具有数据标签。|
||[markerBackgroundColor](/javascript/api/excel/excel.chartseriescollectionloadoptions#markerbackgroundcolor)|对于集合中的每一项: 表示图表系列的标记背景色。|
||[markerForegroundColor](/javascript/api/excel/excel.chartseriescollectionloadoptions#markerforegroundcolor)|对于集合中的每一项: 表示图表系列的标记前景色。|
||[markerSize](/javascript/api/excel/excel.chartseriescollectionloadoptions#markersize)|对于集合中的每一项: 表示图表系列的标记大小。|
||[markerStyle](/javascript/api/excel/excel.chartseriescollectionloadoptions#markerstyle)|对于集合中的每一项: 表示图表系列的标记样式。 有关详细信息, 请参阅 ChartMarkerStyle。|
||[plotOrder](/javascript/api/excel/excel.chartseriescollectionloadoptions#plotorder)|对于集合中的每一项: 代表图表系列在图表组中的绘制顺序。|
||[showShadow](/javascript/api/excel/excel.chartseriescollectionloadoptions#showshadow)|对于集合中的每一项: 布尔值, 表示系列是否具有阴影。|
||[平滑](/javascript/api/excel/excel.chartseriescollectionloadoptions#smooth)|对于集合中的每一项: 布尔值, 表示系列是否平滑。 仅适用于折线图和散点图。|
|[ChartSeriesData](/javascript/api/excel/excel.chartseriesdata)|[chartType](/javascript/api/excel/excel.chartseriesdata#charttype)|表示系列的图表类型。 有关详细信息, 请参阅 ChartType。|
||[doughnutHoleSize](/javascript/api/excel/excel.chartseriesdata#doughnutholesize)|表示图表系列的圆环孔大小。  仅对圆环图和分离型圆环图有效。|
||[筛选](/javascript/api/excel/excel.chartseriesdata#filtered)|该布尔值表示是否筛选系列。 不适用于曲面图。|
||[gapWidth](/javascript/api/excel/excel.chartseriesdata#gapwidth)|表示图表系列的间隙宽度。  有效对象：条形图和柱形图，以及|
||[hasDataLabels](/javascript/api/excel/excel.chartseriesdata#hasdatalabels)|该布尔值表示系列是否具有数据标签。|
||[markerBackgroundColor](/javascript/api/excel/excel.chartseriesdata#markerbackgroundcolor)|表示图表系列的标记背景色。|
||[markerForegroundColor](/javascript/api/excel/excel.chartseriesdata#markerforegroundcolor)|表示图表系列的标记前景色。|
||[markerSize](/javascript/api/excel/excel.chartseriesdata#markersize)|表示图表系列的标记大小。|
||[markerStyle](/javascript/api/excel/excel.chartseriesdata#markerstyle)|表示图表系列的标记类型。 有关详细信息, 请参阅 ChartMarkerStyle。|
||[plotOrder](/javascript/api/excel/excel.chartseriesdata#plotorder)|表示图表组中某个图表系列的绘制顺序。|
||[showShadow](/javascript/api/excel/excel.chartseriesdata#showshadow)|表示系列是否具有阴影的布尔值。|
||[平滑](/javascript/api/excel/excel.chartseriesdata#smooth)|该布尔值表示系列是否平滑。 仅适用于折线图和散点图。|
||[趋势](/javascript/api/excel/excel.chartseriesdata#trendlines)|表示系列中趋势线的集合。 只读。|
|[ChartSeriesLoadOptions](/javascript/api/excel/excel.chartseriesloadoptions)|[chartType](/javascript/api/excel/excel.chartseriesloadoptions#charttype)|表示系列的图表类型。 有关详细信息, 请参阅 ChartType。|
||[doughnutHoleSize](/javascript/api/excel/excel.chartseriesloadoptions#doughnutholesize)|表示图表系列的圆环孔大小。  仅对圆环图和分离型圆环图有效。|
||[筛选](/javascript/api/excel/excel.chartseriesloadoptions#filtered)|该布尔值表示是否筛选系列。 不适用于曲面图。|
||[gapWidth](/javascript/api/excel/excel.chartseriesloadoptions#gapwidth)|表示图表系列的间隙宽度。  有效对象：条形图和柱形图，以及|
||[hasDataLabels](/javascript/api/excel/excel.chartseriesloadoptions#hasdatalabels)|该布尔值表示系列是否具有数据标签。|
||[markerBackgroundColor](/javascript/api/excel/excel.chartseriesloadoptions#markerbackgroundcolor)|表示图表系列的标记背景色。|
||[markerForegroundColor](/javascript/api/excel/excel.chartseriesloadoptions#markerforegroundcolor)|表示图表系列的标记前景色。|
||[markerSize](/javascript/api/excel/excel.chartseriesloadoptions#markersize)|表示图表系列的标记大小。|
||[markerStyle](/javascript/api/excel/excel.chartseriesloadoptions#markerstyle)|表示图表系列的标记类型。 有关详细信息, 请参阅 ChartMarkerStyle。|
||[plotOrder](/javascript/api/excel/excel.chartseriesloadoptions#plotorder)|表示图表组中某个图表系列的绘制顺序。|
||[showShadow](/javascript/api/excel/excel.chartseriesloadoptions#showshadow)|表示系列是否具有阴影的布尔值。|
||[平滑](/javascript/api/excel/excel.chartseriesloadoptions#smooth)|该布尔值表示系列是否平滑。 仅适用于折线图和散点图。|
|[ChartSeriesUpdateData](/javascript/api/excel/excel.chartseriesupdatedata)|[chartType](/javascript/api/excel/excel.chartseriesupdatedata#charttype)|表示系列的图表类型。 有关详细信息, 请参阅 ChartType。|
||[doughnutHoleSize](/javascript/api/excel/excel.chartseriesupdatedata#doughnutholesize)|表示图表系列的圆环孔大小。  仅对圆环图和分离型圆环图有效。|
||[筛选](/javascript/api/excel/excel.chartseriesupdatedata#filtered)|该布尔值表示是否筛选系列。 不适用于曲面图。|
||[gapWidth](/javascript/api/excel/excel.chartseriesupdatedata#gapwidth)|表示图表系列的间隙宽度。  有效对象：条形图和柱形图，以及|
||[hasDataLabels](/javascript/api/excel/excel.chartseriesupdatedata#hasdatalabels)|该布尔值表示系列是否具有数据标签。|
||[markerBackgroundColor](/javascript/api/excel/excel.chartseriesupdatedata#markerbackgroundcolor)|表示图表系列的标记背景色。|
||[markerForegroundColor](/javascript/api/excel/excel.chartseriesupdatedata#markerforegroundcolor)|表示图表系列的标记前景色。|
||[markerSize](/javascript/api/excel/excel.chartseriesupdatedata#markersize)|表示图表系列的标记大小。|
||[markerStyle](/javascript/api/excel/excel.chartseriesupdatedata#markerstyle)|表示图表系列的标记类型。 有关详细信息, 请参阅 ChartMarkerStyle。|
||[plotOrder](/javascript/api/excel/excel.chartseriesupdatedata#plotorder)|表示图表组中某个图表系列的绘制顺序。|
||[showShadow](/javascript/api/excel/excel.chartseriesupdatedata#showshadow)|表示系列是否具有阴影的布尔值。|
||[平滑](/javascript/api/excel/excel.chartseriesupdatedata#smooth)|该布尔值表示系列是否平滑。 仅适用于折线图和散点图。|
|[ChartTitle](/javascript/api/excel/excel.charttitle)|[getSubstring (start: number, length: 数字)](/javascript/api/excel/excel.charttitle#getsubstring-start--length-)|获取图表标题的子字符串。 换行符 "\n" 也计算一个字符。|
||[horizontalAlignment](/javascript/api/excel/excel.charttitle#horizontalalignment)|表示图表标题水平对齐。|
||[left](/javascript/api/excel/excel.charttitle#left)|表示图表标题左边缘到图表区域左边缘的距离，以磅为单位。 如果图表标题不可见, 则为 Null。|
||[position](/javascript/api/excel/excel.charttitle#position)|表示图表标题的位置。 有关详细信息, 请参阅 ChartTitlePosition。|
||[height](/javascript/api/excel/excel.charttitle#height)|返回图表标题的高度，以磅为单位。 如果图表标题不可见, 则为 Null。 只读。|
||[width](/javascript/api/excel/excel.charttitle#width)|返回图表标题的宽度，以磅为单位。 如果图表标题不可见, 则为 Null。 只读。|
||[setFormula (formula: string)](/javascript/api/excel/excel.charttitle#setformula-formula-)|设置一个字符串值，用于表示采用 A1 表示法的图表标题的公式。|
||[showShadow](/javascript/api/excel/excel.charttitle#showshadow)|表示一个布尔值，用于确定图表标题是否具有阴影。|
||[textOrientation](/javascript/api/excel/excel.charttitle#textorientation)|表示图表标题的文本方向。 此值应是 -90 到 90 或 180（垂直文本）之间的整数。|
||[top](/javascript/api/excel/excel.charttitle#top)|表示图表标题上边缘到图表区域顶部的距离，以磅为单位。 如果图表标题不可见, 则为 Null。|
||[verticalAlignment](/javascript/api/excel/excel.charttitle#verticalalignment)|表示图表标题垂直对齐。 有关详细信息, 请参阅 ChartTextVerticalAlignment。|
|[ChartTitleData](/javascript/api/excel/excel.charttitledata)|[height](/javascript/api/excel/excel.charttitledata#height)|返回图表标题的高度，以磅为单位。 如果图表标题不可见, 则为 Null。 只读。|
||[horizontalAlignment](/javascript/api/excel/excel.charttitledata#horizontalalignment)|表示图表标题水平对齐。|
||[left](/javascript/api/excel/excel.charttitledata#left)|表示图表标题左边缘到图表区域左边缘的距离，以磅为单位。 如果图表标题不可见, 则为 Null。|
||[position](/javascript/api/excel/excel.charttitledata#position)|表示图表标题的位置。 有关详细信息, 请参阅 ChartTitlePosition。|
||[showShadow](/javascript/api/excel/excel.charttitledata#showshadow)|表示一个布尔值，用于确定图表标题是否具有阴影。|
||[textOrientation](/javascript/api/excel/excel.charttitledata#textorientation)|表示图表标题的文本方向。 此值应是 -90 到 90 或 180（垂直文本）之间的整数。|
||[top](/javascript/api/excel/excel.charttitledata#top)|表示图表标题上边缘到图表区域顶部的距离，以磅为单位。 如果图表标题不可见, 则为 Null。|
||[verticalAlignment](/javascript/api/excel/excel.charttitledata#verticalalignment)|表示图表标题垂直对齐。 有关详细信息, 请参阅 ChartTextVerticalAlignment。|
||[width](/javascript/api/excel/excel.charttitledata#width)|返回图表标题的宽度，以磅为单位。 如果图表标题不可见, 则为 Null。 只读。|
|[ChartTitleFormat](/javascript/api/excel/excel.charttitleformat)|[边缘](/javascript/api/excel/excel.charttitleformat#border)|代表图表标题的边框格式, 包括颜色、linestyle 和粗细。 只读。|
|[ChartTitleFormatData](/javascript/api/excel/excel.charttitleformatdata)|[边缘](/javascript/api/excel/excel.charttitleformatdata#border)|代表图表标题的边框格式, 包括颜色、linestyle 和粗细。 只读。|
|[ChartTitleFormatLoadOptions](/javascript/api/excel/excel.charttitleformatloadoptions)|[边缘](/javascript/api/excel/excel.charttitleformatloadoptions#border)|代表图表标题的边框格式, 包括颜色、linestyle 和粗细。|
|[ChartTitleFormatUpdateData](/javascript/api/excel/excel.charttitleformatupdatedata)|[边缘](/javascript/api/excel/excel.charttitleformatupdatedata#border)|代表图表标题的边框格式, 包括颜色、linestyle 和粗细。|
|[ChartTitleLoadOptions](/javascript/api/excel/excel.charttitleloadoptions)|[height](/javascript/api/excel/excel.charttitleloadoptions#height)|返回图表标题的高度，以磅为单位。 如果图表标题不可见, 则为 Null。 只读。|
||[horizontalAlignment](/javascript/api/excel/excel.charttitleloadoptions#horizontalalignment)|表示图表标题水平对齐。|
||[left](/javascript/api/excel/excel.charttitleloadoptions#left)|表示图表标题左边缘到图表区域左边缘的距离，以磅为单位。 如果图表标题不可见, 则为 Null。|
||[position](/javascript/api/excel/excel.charttitleloadoptions#position)|表示图表标题的位置。 有关详细信息, 请参阅 ChartTitlePosition。|
||[showShadow](/javascript/api/excel/excel.charttitleloadoptions#showshadow)|表示一个布尔值，用于确定图表标题是否具有阴影。|
||[textOrientation](/javascript/api/excel/excel.charttitleloadoptions#textorientation)|表示图表标题的文本方向。 此值应是 -90 到 90 或 180（垂直文本）之间的整数。|
||[top](/javascript/api/excel/excel.charttitleloadoptions#top)|表示图表标题上边缘到图表区域顶部的距离，以磅为单位。 如果图表标题不可见, 则为 Null。|
||[verticalAlignment](/javascript/api/excel/excel.charttitleloadoptions#verticalalignment)|表示图表标题垂直对齐。 有关详细信息, 请参阅 ChartTextVerticalAlignment。|
||[width](/javascript/api/excel/excel.charttitleloadoptions#width)|返回图表标题的宽度，以磅为单位。 如果图表标题不可见, 则为 Null。 只读。|
|[ChartTitleUpdateData](/javascript/api/excel/excel.charttitleupdatedata)|[horizontalAlignment](/javascript/api/excel/excel.charttitleupdatedata#horizontalalignment)|表示图表标题水平对齐。|
||[left](/javascript/api/excel/excel.charttitleupdatedata#left)|表示图表标题左边缘到图表区域左边缘的距离，以磅为单位。 如果图表标题不可见, 则为 Null。|
||[position](/javascript/api/excel/excel.charttitleupdatedata#position)|表示图表标题的位置。 有关详细信息, 请参阅 ChartTitlePosition。|
||[showShadow](/javascript/api/excel/excel.charttitleupdatedata#showshadow)|表示一个布尔值，用于确定图表标题是否具有阴影。|
||[textOrientation](/javascript/api/excel/excel.charttitleupdatedata#textorientation)|表示图表标题的文本方向。 此值应是 -90 到 90 或 180（垂直文本）之间的整数。|
||[top](/javascript/api/excel/excel.charttitleupdatedata#top)|表示图表标题上边缘到图表区域顶部的距离，以磅为单位。 如果图表标题不可见, 则为 Null。|
||[verticalAlignment](/javascript/api/excel/excel.charttitleupdatedata#verticalalignment)|表示图表标题垂直对齐。 有关详细信息, 请参阅 ChartTextVerticalAlignment。|
|[ChartTrendline](/javascript/api/excel/excel.charttrendline)|[delete()](/javascript/api/excel/excel.charttrendline#delete--)|删除 Trendline 对象。|
||[intercept](/javascript/api/excel/excel.charttrendline#intercept)|表示趋势线的截距值。 可以设置为数字值或空字符串（对于自动值）。 返回的值始终为数字。|
||[movingAveragePeriod](/javascript/api/excel/excel.charttrendline#movingaverageperiod)|表示图表趋势线的周期。 仅适用于 MovingAverage 类型的趋势线。|
||[name](/javascript/api/excel/excel.charttrendline#name)|表示趋势线的名称。 可设为字符串值，或者设为 Null 值（表示自动值）。 返回的值始终为字符串|
||[polynomialOrder](/javascript/api/excel/excel.charttrendline#polynomialorder)|表示图表趋势线的顺序。 仅适用于具有多项式类型的趋势线。|
||[format](/javascript/api/excel/excel.charttrendline#format)|表示图表趋势线的格式。|
||[set (properties: ChartTrendline)](/javascript/api/excel/excel.charttrendline#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: ChartTrendlineUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.charttrendline#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
||[type](/javascript/api/excel/excel.charttrendline#type)|表示图表趋势线的类型。|
|[ChartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|[add (type？: "线性" \| "指数" \| "对数" \| "MovingAverage" \| "多项式\| " "Power")](/javascript/api/excel/excel.charttrendlinecollection#add-type-)|向趋势线集合添加新的趋势线。|
||[add (type？: ChartTrendlineType)](/javascript/api/excel/excel.charttrendlinecollection#add-type-)|向趋势线集合添加新的趋势线。|
||[getCount()](/javascript/api/excel/excel.charttrendlinecollection#getcount--)|返回集合中的趋势线数量。|
||[getItem(index: number)](/javascript/api/excel/excel.charttrendlinecollection#getitem-index-)|按索引（在项目数组中的插入顺序）获取 Trendline 对象。|
||[items](/javascript/api/excel/excel.charttrendlinecollection#items)|获取此集合中已加载的子项。|
|[ChartTrendlineCollectionLoadOptions](/javascript/api/excel/excel.charttrendlinecollectionloadoptions)|[$all](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#$all)||
||[format](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#format)|对于集合中的每一项: 代表图表趋势线的格式。|
||[intercept](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#intercept)|对于集合中的每一项: 代表趋势线的截距值。 可以设置为数字值或空字符串（对于自动值）。 返回的值始终为数字。|
||[movingAveragePeriod](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#movingaverageperiod)|对于集合中的每一项: 代表图表趋势线的周期。 仅适用于 MovingAverage 类型的趋势线。|
||[name](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#name)|对于集合中的每一项: 代表趋势线的名称。 可设为字符串值，或者设为 Null 值（表示自动值）。 返回的值始终为字符串|
||[polynomialOrder](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#polynomialorder)|对于集合中的每一项: 表示图表趋势线的顺序。 仅适用于具有多项式类型的趋势线。|
||[type](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#type)|对于集合中的每一项: 代表图表趋势线的类型。|
|[ChartTrendlineData](/javascript/api/excel/excel.charttrendlinedata)|[format](/javascript/api/excel/excel.charttrendlinedata#format)|表示图表趋势线的格式。|
||[intercept](/javascript/api/excel/excel.charttrendlinedata#intercept)|表示趋势线的截距值。 可以设置为数字值或空字符串（对于自动值）。 返回的值始终为数字。|
||[movingAveragePeriod](/javascript/api/excel/excel.charttrendlinedata#movingaverageperiod)|表示图表趋势线的周期。 仅适用于 MovingAverage 类型的趋势线。|
||[name](/javascript/api/excel/excel.charttrendlinedata#name)|表示趋势线的名称。 可设为字符串值，或者设为 Null 值（表示自动值）。 返回的值始终为字符串|
||[polynomialOrder](/javascript/api/excel/excel.charttrendlinedata#polynomialorder)|表示图表趋势线的顺序。 仅适用于具有多项式类型的趋势线。|
||[type](/javascript/api/excel/excel.charttrendlinedata#type)|表示图表趋势线的类型。|
|[ChartTrendlineFormat](/javascript/api/excel/excel.charttrendlineformat)|[line](/javascript/api/excel/excel.charttrendlineformat#line)|表示图表线条格式。 只读。|
||[set (properties: ChartTrendlineFormat)](/javascript/api/excel/excel.charttrendlineformat#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: ChartTrendlineFormatUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.charttrendlineformat#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
|[ChartTrendlineFormatData](/javascript/api/excel/excel.charttrendlineformatdata)|[line](/javascript/api/excel/excel.charttrendlineformatdata#line)|表示图表线条格式。 只读。|
|[ChartTrendlineFormatLoadOptions](/javascript/api/excel/excel.charttrendlineformatloadoptions)|[$all](/javascript/api/excel/excel.charttrendlineformatloadoptions#$all)||
||[line](/javascript/api/excel/excel.charttrendlineformatloadoptions#line)|表示图表线条格式。|
|[ChartTrendlineFormatUpdateData](/javascript/api/excel/excel.charttrendlineformatupdatedata)|[line](/javascript/api/excel/excel.charttrendlineformatupdatedata#line)|表示图表线条格式。|
|[ChartTrendlineLoadOptions](/javascript/api/excel/excel.charttrendlineloadoptions)|[$all](/javascript/api/excel/excel.charttrendlineloadoptions#$all)||
||[format](/javascript/api/excel/excel.charttrendlineloadoptions#format)|表示图表趋势线的格式。|
||[intercept](/javascript/api/excel/excel.charttrendlineloadoptions#intercept)|表示趋势线的截距值。 可以设置为数字值或空字符串（对于自动值）。 返回的值始终为数字。|
||[movingAveragePeriod](/javascript/api/excel/excel.charttrendlineloadoptions#movingaverageperiod)|表示图表趋势线的周期。 仅适用于 MovingAverage 类型的趋势线。|
||[name](/javascript/api/excel/excel.charttrendlineloadoptions#name)|表示趋势线的名称。 可设为字符串值，或者设为 Null 值（表示自动值）。 返回的值始终为字符串|
||[polynomialOrder](/javascript/api/excel/excel.charttrendlineloadoptions#polynomialorder)|表示图表趋势线的顺序。 仅适用于具有多项式类型的趋势线。|
||[type](/javascript/api/excel/excel.charttrendlineloadoptions#type)|表示图表趋势线的类型。|
|[ChartTrendlineUpdateData](/javascript/api/excel/excel.charttrendlineupdatedata)|[format](/javascript/api/excel/excel.charttrendlineupdatedata#format)|表示图表趋势线的格式。|
||[intercept](/javascript/api/excel/excel.charttrendlineupdatedata#intercept)|表示趋势线的截距值。 可以设置为数字值或空字符串（对于自动值）。 返回的值始终为数字。|
||[movingAveragePeriod](/javascript/api/excel/excel.charttrendlineupdatedata#movingaverageperiod)|表示图表趋势线的周期。 仅适用于 MovingAverage 类型的趋势线。|
||[name](/javascript/api/excel/excel.charttrendlineupdatedata#name)|表示趋势线的名称。 可设为字符串值，或者设为 Null 值（表示自动值）。 返回的值始终为字符串|
||[polynomialOrder](/javascript/api/excel/excel.charttrendlineupdatedata#polynomialorder)|表示图表趋势线的顺序。 仅适用于具有多项式类型的趋势线。|
||[type](/javascript/api/excel/excel.charttrendlineupdatedata#type)|表示图表趋势线的类型。|
|[ChartUpdateData](/javascript/api/excel/excel.chartupdatedata)|[chartType](/javascript/api/excel/excel.chartupdatedata#charttype)|表示图表的类型。 有关详细信息, 请参阅 ChartType。|
||[showAllFieldButtons](/javascript/api/excel/excel.chartupdatedata#showallfieldbuttons)|表示是否在数据透视图上显示所有字段按钮。|
|[CustomProperty](/javascript/api/excel/excel.customproperty)|[delete()](/javascript/api/excel/excel.customproperty#delete--)|删除 custom property 对象。|
||[key](/javascript/api/excel/excel.customproperty#key)|获取 customProperty 的键。 只读。|
||[type](/javascript/api/excel/excel.customproperty#type)|获取自定义属性的值类型。 只读。|
||[set (properties: CustomProperty)](/javascript/api/excel/excel.customproperty#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: CustomPropertyUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.customproperty#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
||[value](/javascript/api/excel/excel.customproperty#value)|获取或设置自定义属性的值。|
|[CustomPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|[add (key: string, value: any)](/javascript/api/excel/excel.custompropertycollection#add-key--value-)|新建自定义属性或设置现有自定义属性。|
||[deleteAll ()](/javascript/api/excel/excel.custompropertycollection#deleteall--)|删除此集合中的所有自定义属性。|
||[getCount()](/javascript/api/excel/excel.custompropertycollection#getcount--)|获取自定义属性的计数。|
||[getItem(key: string)](/javascript/api/excel/excel.custompropertycollection#getitem-key-)|按键获取自定义属性对象（不区分大小写）。 如果自定义属性不存在, 则引发此异常。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.custompropertycollection#getitemornullobject-key-)|按键获取自定义属性对象（不区分大小写）。 如果自定义属性不存在, 则返回 null 对象。|
||[items](/javascript/api/excel/excel.custompropertycollection#items)|获取此集合中已加载的子项。|
|[CustomPropertyCollectionLoadOptions](/javascript/api/excel/excel.custompropertycollectionloadoptions)|[$all](/javascript/api/excel/excel.custompropertycollectionloadoptions#$all)||
||[key](/javascript/api/excel/excel.custompropertycollectionloadoptions#key)|对于集合中的每一项: 获取自定义属性的键。 只读。|
||[type](/javascript/api/excel/excel.custompropertycollectionloadoptions#type)|对于集合中的每一项: 获取自定义属性的值类型。 只读。|
||[value](/javascript/api/excel/excel.custompropertycollectionloadoptions#value)|对于集合中的每一项: 获取或设置自定义属性的值。|
|[CustomPropertyData](/javascript/api/excel/excel.custompropertydata)|[key](/javascript/api/excel/excel.custompropertydata#key)|获取 customProperty 的键。 只读。|
||[type](/javascript/api/excel/excel.custompropertydata#type)|获取自定义属性的值类型。 只读。|
||[value](/javascript/api/excel/excel.custompropertydata#value)|获取或设置自定义属性的值。|
|[CustomPropertyLoadOptions](/javascript/api/excel/excel.custompropertyloadoptions)|[$all](/javascript/api/excel/excel.custompropertyloadoptions#$all)||
||[key](/javascript/api/excel/excel.custompropertyloadoptions#key)|获取 customProperty 的键。 只读。|
||[type](/javascript/api/excel/excel.custompropertyloadoptions#type)|获取自定义属性的值类型。 只读。|
||[value](/javascript/api/excel/excel.custompropertyloadoptions#value)|获取或设置自定义属性的值。|
|[CustomPropertyUpdateData](/javascript/api/excel/excel.custompropertyupdatedata)|[value](/javascript/api/excel/excel.custompropertyupdatedata#value)|获取或设置自定义属性的值。|
|[DataConnectionCollection](/javascript/api/excel/excel.dataconnectioncollection)|[refreshAll ()](/javascript/api/excel/excel.dataconnectioncollection#refreshall--)|刷新集合中的所有数据连接。|
|[DocumentProperties](/javascript/api/excel/excel.documentproperties)|[编写](/javascript/api/excel/excel.documentproperties#author)|获取或设置工作簿的作者。|
||[类别](/javascript/api/excel/excel.documentproperties#category)|获取或设置工作簿的类别。|
||[comments](/javascript/api/excel/excel.documentproperties#comments)|获取或设置工作簿的注释。|
||[company](/javascript/api/excel/excel.documentproperties#company)|获取或设置工作簿的公司。|
||[关键字](/javascript/api/excel/excel.documentproperties#keywords)|获取或设置工作簿的关键字。|
||[管理器](/javascript/api/excel/excel.documentproperties#manager)|获取或设置工作簿的管理者。|
||[creationDate](/javascript/api/excel/excel.documentproperties#creationdate)|获取工作簿的创建日期。 只读。|
||[自](/javascript/api/excel/excel.documentproperties#custom)|获取工作簿的自定义属性的集合。 只读。|
||[lastAuthor](/javascript/api/excel/excel.documentproperties#lastauthor)|获取工作簿的最终作者。 只读。|
||[revisionNumber](/javascript/api/excel/excel.documentproperties#revisionnumber)|获取工作簿的修订号。 只读。|
||[set (properties: DocumentProperties)](/javascript/api/excel/excel.documentproperties#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: DocumentPropertiesUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.documentproperties#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
||[subject](/javascript/api/excel/excel.documentproperties#subject)|获取或设置工作簿的主题。|
||[title](/javascript/api/excel/excel.documentproperties#title)|获取或设置工作簿的标题。|
|[DocumentPropertiesData](/javascript/api/excel/excel.documentpropertiesdata)|[编写](/javascript/api/excel/excel.documentpropertiesdata#author)|获取或设置工作簿的作者。|
||[类别](/javascript/api/excel/excel.documentpropertiesdata#category)|获取或设置工作簿的类别。|
||[comments](/javascript/api/excel/excel.documentpropertiesdata#comments)|获取或设置工作簿的注释。|
||[company](/javascript/api/excel/excel.documentpropertiesdata#company)|获取或设置工作簿的公司。|
||[creationDate](/javascript/api/excel/excel.documentpropertiesdata#creationdate)|获取工作簿的创建日期。 只读。|
||[自](/javascript/api/excel/excel.documentpropertiesdata#custom)|获取工作簿的自定义属性的集合。 只读。|
||[关键字](/javascript/api/excel/excel.documentpropertiesdata#keywords)|获取或设置工作簿的关键字。|
||[lastAuthor](/javascript/api/excel/excel.documentpropertiesdata#lastauthor)|获取工作簿的最终作者。 只读。|
||[管理器](/javascript/api/excel/excel.documentpropertiesdata#manager)|获取或设置工作簿的管理者。|
||[revisionNumber](/javascript/api/excel/excel.documentpropertiesdata#revisionnumber)|获取工作簿的修订号。 只读。|
||[subject](/javascript/api/excel/excel.documentpropertiesdata#subject)|获取或设置工作簿的主题。|
||[title](/javascript/api/excel/excel.documentpropertiesdata#title)|获取或设置工作簿的标题。|
|[DocumentPropertiesLoadOptions](/javascript/api/excel/excel.documentpropertiesloadoptions)|[$all](/javascript/api/excel/excel.documentpropertiesloadoptions#$all)||
||[编写](/javascript/api/excel/excel.documentpropertiesloadoptions#author)|获取或设置工作簿的作者。|
||[类别](/javascript/api/excel/excel.documentpropertiesloadoptions#category)|获取或设置工作簿的类别。|
||[comments](/javascript/api/excel/excel.documentpropertiesloadoptions#comments)|获取或设置工作簿的注释。|
||[company](/javascript/api/excel/excel.documentpropertiesloadoptions#company)|获取或设置工作簿的公司。|
||[creationDate](/javascript/api/excel/excel.documentpropertiesloadoptions#creationdate)|获取工作簿的创建日期。 只读。|
||[关键字](/javascript/api/excel/excel.documentpropertiesloadoptions#keywords)|获取或设置工作簿的关键字。|
||[lastAuthor](/javascript/api/excel/excel.documentpropertiesloadoptions#lastauthor)|获取工作簿的最终作者。 只读。|
||[管理器](/javascript/api/excel/excel.documentpropertiesloadoptions#manager)|获取或设置工作簿的管理者。|
||[revisionNumber](/javascript/api/excel/excel.documentpropertiesloadoptions#revisionnumber)|获取工作簿的修订号。 只读。|
||[subject](/javascript/api/excel/excel.documentpropertiesloadoptions#subject)|获取或设置工作簿的主题。|
||[title](/javascript/api/excel/excel.documentpropertiesloadoptions#title)|获取或设置工作簿的标题。|
|[DocumentPropertiesUpdateData](/javascript/api/excel/excel.documentpropertiesupdatedata)|[编写](/javascript/api/excel/excel.documentpropertiesupdatedata#author)|获取或设置工作簿的作者。|
||[类别](/javascript/api/excel/excel.documentpropertiesupdatedata#category)|获取或设置工作簿的类别。|
||[comments](/javascript/api/excel/excel.documentpropertiesupdatedata#comments)|获取或设置工作簿的注释。|
||[company](/javascript/api/excel/excel.documentpropertiesupdatedata#company)|获取或设置工作簿的公司。|
||[关键字](/javascript/api/excel/excel.documentpropertiesupdatedata#keywords)|获取或设置工作簿的关键字。|
||[管理器](/javascript/api/excel/excel.documentpropertiesupdatedata#manager)|获取或设置工作簿的管理者。|
||[revisionNumber](/javascript/api/excel/excel.documentpropertiesupdatedata#revisionnumber)|获取工作簿的修订号。 只读。|
||[subject](/javascript/api/excel/excel.documentpropertiesupdatedata#subject)|获取或设置工作簿的主题。|
||[title](/javascript/api/excel/excel.documentpropertiesupdatedata#title)|获取或设置工作簿的标题。|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[formula](/javascript/api/excel/excel.nameditem#formula)|获取或设置的已命名项目的公式。  公式始终以等号 (=) 开头。|
||[arrayValues](/javascript/api/excel/excel.nameditem#arrayvalues)|返回包含已命名项目的值和类型的对象。 只读。|
|[NamedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|[types](/javascript/api/excel/excel.nameditemarrayvalues#types)|表示已命名项目数组中每个项目的类型|
||[values](/javascript/api/excel/excel.nameditemarrayvalues#values)|表示已命名项目数组中每个项目的值。|
|[NamedItemArrayValuesData](/javascript/api/excel/excel.nameditemarrayvaluesdata)|[types](/javascript/api/excel/excel.nameditemarrayvaluesdata#types)|表示已命名项目数组中每个项目的类型|
||[values](/javascript/api/excel/excel.nameditemarrayvaluesdata#values)|表示已命名项目数组中每个项目的值。|
|[NamedItemArrayValuesLoadOptions](/javascript/api/excel/excel.nameditemarrayvaluesloadoptions)|[$all](/javascript/api/excel/excel.nameditemarrayvaluesloadoptions#$all)||
||[types](/javascript/api/excel/excel.nameditemarrayvaluesloadoptions#types)|表示已命名项目数组中每个项目的类型|
||[values](/javascript/api/excel/excel.nameditemarrayvaluesloadoptions#values)|表示已命名项目数组中每个项目的值。|
|[NamedItemCollectionLoadOptions](/javascript/api/excel/excel.nameditemcollectionloadoptions)|[arrayValues](/javascript/api/excel/excel.nameditemcollectionloadoptions#arrayvalues)|对于集合中的每一项: 返回一个对象, 其中包含已命名项目的值和类型。|
||[formula](/javascript/api/excel/excel.nameditemcollectionloadoptions#formula)|对于集合中的每一项: 获取或设置已命名项的公式。  公式始终以等号 (=) 开头。|
|[NamedItemData](/javascript/api/excel/excel.nameditemdata)|[arrayValues](/javascript/api/excel/excel.nameditemdata#arrayvalues)|返回包含已命名项目的值和类型的对象。 只读。|
||[formula](/javascript/api/excel/excel.nameditemdata#formula)|获取或设置的已命名项目的公式。  公式始终以等号 (=) 开头。|
|[NamedItemLoadOptions](/javascript/api/excel/excel.nameditemloadoptions)|[arrayValues](/javascript/api/excel/excel.nameditemloadoptions#arrayvalues)|返回包含已命名项目的值和类型的对象。|
||[formula](/javascript/api/excel/excel.nameditemloadoptions#formula)|获取或设置的已命名项目的公式。  公式始终以等号 (=) 开头。|
|[NamedItemUpdateData](/javascript/api/excel/excel.nameditemupdatedata)|[formula](/javascript/api/excel/excel.nameditemupdatedata#formula)|获取或设置的已命名项目的公式。  公式始终以等号 (=) 开头。|
|[Range](/javascript/api/excel/excel.range)|[getAbsoluteResizedRange (numRows: 数字, numColumns: 数字)](/javascript/api/excel/excel.range#getabsoluteresizedrange-numrows--numcolumns-)|获取一个 Range 对象，该对象的左上单元格与当前 Range 对象相同，但具有指定的行数和列数。|
||[getImage ()](/javascript/api/excel/excel.range#getimage--)|将区域呈现为 base64 编码的 png 图像。|
||[getSurroundingRegion()](/javascript/api/excel/excel.range#getsurroundingregion--)|返回一个 Range 对象，该对象表示此区域左上单元格的周围区域。 周围区域是由相对于该区域的空白行和空白列的任何组合所限定的区域。|
||[hyperlink](/javascript/api/excel/excel.range#hyperlink)|表示当前区域的超链接。|
||[numberFormatLocal](/javascript/api/excel/excel.range#numberformatlocal)|表示 Excel 中的给定区域的数字格式代码，以用户语言的字符串表示。|
||[isEntireColumn](/javascript/api/excel/excel.range#isentirecolumn)|表示当前区域是否为整列。 只读。|
||[isEntireRow](/javascript/api/excel/excel.range#isentirerow)|表示当前区域是否为整行。 只读。|
||[showCard()](/javascript/api/excel/excel.range#showcard--)|显示活动单元格的卡片（如果该单元格具有富值内容）。|
||[style](/javascript/api/excel/excel.range#style)|表示当前区域的样式。|
|[RangeData](/javascript/api/excel/excel.rangedata)|[hyperlink](/javascript/api/excel/excel.rangedata#hyperlink)|表示当前区域的超链接。|
||[isEntireColumn](/javascript/api/excel/excel.rangedata#isentirecolumn)|表示当前区域是否为整列。 只读。|
||[isEntireRow](/javascript/api/excel/excel.rangedata#isentirerow)|表示当前区域是否为整行。 只读。|
||[numberFormatLocal](/javascript/api/excel/excel.rangedata#numberformatlocal)|表示 Excel 中的给定区域的数字格式代码，以用户语言的字符串表示。|
||[style](/javascript/api/excel/excel.rangedata#style)|表示当前区域的样式。|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[textOrientation](/javascript/api/excel/excel.rangeformat#textorientation)|获取或设置区域内的所有单元格的文本方向。|
||[useStandardHeight](/javascript/api/excel/excel.rangeformat#usestandardheight)|确定 Range 对象的行高是否等于工作表的标准行高。|
||[useStandardWidth](/javascript/api/excel/excel.rangeformat#usestandardwidth)|指示 Range 对象的列宽是否等于工作表的标准宽度。|
|[RangeFormatData](/javascript/api/excel/excel.rangeformatdata)|[textOrientation](/javascript/api/excel/excel.rangeformatdata#textorientation)|获取或设置区域内的所有单元格的文本方向。|
||[useStandardHeight](/javascript/api/excel/excel.rangeformatdata#usestandardheight)|确定 Range 对象的行高是否等于工作表的标准行高。|
||[useStandardWidth](/javascript/api/excel/excel.rangeformatdata#usestandardwidth)|指示 Range 对象的列宽是否等于工作表的标准宽度。|
|[RangeFormatLoadOptions](/javascript/api/excel/excel.rangeformatloadoptions)|[textOrientation](/javascript/api/excel/excel.rangeformatloadoptions#textorientation)|获取或设置区域内的所有单元格的文本方向。|
||[useStandardHeight](/javascript/api/excel/excel.rangeformatloadoptions#usestandardheight)|确定 Range 对象的行高是否等于工作表的标准行高。|
||[useStandardWidth](/javascript/api/excel/excel.rangeformatloadoptions#usestandardwidth)|指示 Range 对象的列宽是否等于工作表的标准宽度。|
|[RangeFormatUpdateData](/javascript/api/excel/excel.rangeformatupdatedata)|[textOrientation](/javascript/api/excel/excel.rangeformatupdatedata#textorientation)|获取或设置区域内的所有单元格的文本方向。|
||[useStandardHeight](/javascript/api/excel/excel.rangeformatupdatedata#usestandardheight)|确定 Range 对象的行高是否等于工作表的标准行高。|
||[useStandardWidth](/javascript/api/excel/excel.rangeformatupdatedata#usestandardwidth)|指示 Range 对象的列宽是否等于工作表的标准宽度。|
|[RangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|[address](/javascript/api/excel/excel.rangehyperlink#address)|表示超链接的 URL 目标。|
||[documentReference](/javascript/api/excel/excel.rangehyperlink#documentreference)|表示超链接的文档引用目标。|
||[屏幕](/javascript/api/excel/excel.rangehyperlink#screentip)|表示鼠标悬停在超链接上时显示的字符串。|
||[textToDisplay](/javascript/api/excel/excel.rangehyperlink#texttodisplay)|表示区域最左上方单元格中显示的字符串。|
|[RangeLoadOptions](/javascript/api/excel/excel.rangeloadoptions)|[hyperlink](/javascript/api/excel/excel.rangeloadoptions#hyperlink)|表示当前区域的超链接。|
||[isEntireColumn](/javascript/api/excel/excel.rangeloadoptions#isentirecolumn)|表示当前区域是否为整列。 只读。|
||[isEntireRow](/javascript/api/excel/excel.rangeloadoptions#isentirerow)|表示当前区域是否为整行。 只读。|
||[numberFormatLocal](/javascript/api/excel/excel.rangeloadoptions#numberformatlocal)|表示 Excel 中的给定区域的数字格式代码，以用户语言的字符串表示。|
||[style](/javascript/api/excel/excel.rangeloadoptions#style)|表示当前区域的样式。|
|[RangeUpdateData](/javascript/api/excel/excel.rangeupdatedata)|[hyperlink](/javascript/api/excel/excel.rangeupdatedata#hyperlink)|表示当前区域的超链接。|
||[numberFormatLocal](/javascript/api/excel/excel.rangeupdatedata#numberformatlocal)|表示 Excel 中的给定区域的数字格式代码，以用户语言的字符串表示。|
||[style](/javascript/api/excel/excel.rangeupdatedata#style)|表示当前区域的样式。|
|[Style](/javascript/api/excel/excel.style)|[delete()](/javascript/api/excel/excel.style#delete--)|删除此样式。|
||[formulaHidden](/javascript/api/excel/excel.style#formulahidden)|指示工作表受保护时是否隐藏公式。|
||[horizontalAlignment](/javascript/api/excel/excel.style#horizontalalignment)|表示样式水平对齐。 有关详细信息, 请参阅 HorizontalAlignment。|
||[includeAlignment](/javascript/api/excel/excel.style#includealignment)|指示样式是否包含 AutoIndent、HorizontalAlignment、VerticalAlignment、WrapText、IndentLevel 和 TextOrientation 属性。|
||[includeBorder](/javascript/api/excel/excel.style#includeborder)|指示样式是否包含 Color、ColorIndex、LineStyle 和 Weight 边框属性。|
||[includeFont](/javascript/api/excel/excel.style#includefont)|指示样式是否包含 Background、Bold、Color、ColorIndex、FontStyle、Italic、Name、Size、Strikethrough、Subscrip、Superscript 和 Underline 字体属性。|
||[includeNumber](/javascript/api/excel/excel.style#includenumber)|指示样式是否包含 NumberFormat 属性。|
||[includePatterns](/javascript/api/excel/excel.style#includepatterns)|指示样式是否包含 Color、ColorIndex、InvertIfNegative、Pattern、PatternColor 和 PatternColorIndex 内部属性。|
||[includeProtection](/javascript/api/excel/excel.style#includeprotection)|指示样式是否包含 FormulaHidden 和 Locked 保护属性。|
||[indentLevel](/javascript/api/excel/excel.style#indentlevel)|0 到 250 之间的一个整数，指示样式的缩进水平。|
||[locked](/javascript/api/excel/excel.style#locked)|指示工作表受保护时是否锁定对象。|
||[numberFormat](/javascript/api/excel/excel.style#numberformat)|样式中数字格式的格式代码。|
||[numberFormatLocal](/javascript/api/excel/excel.style#numberformatlocal)|样式中数字格式的本地化格式代码。|
||[readingOrder](/javascript/api/excel/excel.style#readingorder)|样式中的阅读顺序。|
||[Borders](/javascript/api/excel/excel.style#borders)|四个 Border 对象的 Border 集合，表示四个边框的样式。|
||[内置](/javascript/api/excel/excel.style#builtin)|指示样式是否为内置样式。|
||[fill](/javascript/api/excel/excel.style#fill)|样式的填充。|
||[font](/javascript/api/excel/excel.style#font)|该 Font 对象表示样式的字体。|
||[name](/javascript/api/excel/excel.style#name)|样式的名称。|
||[set (properties: Excel 样式)](/javascript/api/excel/excel.style#set-properties-)|基于现有加载的对象同时设置该对象的多个属性。|
||[set (properties: StyleUpdateData, options？: Officeextension.error)](/javascript/api/excel/excel.style#set-properties--options-)|同时设置一个对象的多个属性。 您可以传递具有相应属性的纯对象或相同类型的其他 API 对象。|
||[shrinkToFit](/javascript/api/excel/excel.style#shrinktofit)|指示文本是否自动缩小以适合可用列宽。|
||[verticalAlignment](/javascript/api/excel/excel.style#verticalalignment)|表示样式的垂直对齐方式。 有关详细信息, 请参阅 VerticalAlignment。|
||[wrapText](/javascript/api/excel/excel.style#wraptext)|指示 Microsoft Excel 是否将对象中的文本换行。|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[add(name: string)](/javascript/api/excel/excel.stylecollection#add-name-)|向集合添加新样式。|
||[getItem(name: string)](/javascript/api/excel/excel.stylecollection#getitem-name-)|按名称获取样式。|
||[items](/javascript/api/excel/excel.stylecollection#items)|获取此集合中已加载的子项。|
|[StyleCollectionLoadOptions](/javascript/api/excel/excel.stylecollectionloadoptions)|[$all](/javascript/api/excel/excel.stylecollectionloadoptions#$all)||
||[Borders](/javascript/api/excel/excel.stylecollectionloadoptions#borders)|对于集合中的每个项目: 代表四个边框的样式的四个 Border 对象的边框集合。|
||[内置](/javascript/api/excel/excel.stylecollectionloadoptions#builtin)|对于集合中的每一项: 指示样式是否为内置样式。|
||[fill](/javascript/api/excel/excel.stylecollectionloadoptions#fill)|对于集合中的每个项目: 样式的填充。|
||[font](/javascript/api/excel/excel.stylecollectionloadoptions#font)|对于集合中的每一项: 一个代表样式的字体的 Font 对象。|
||[formulaHidden](/javascript/api/excel/excel.stylecollectionloadoptions#formulahidden)|对于集合中的每一项: 指示工作表受保护时是否隐藏公式。|
||[horizontalAlignment](/javascript/api/excel/excel.stylecollectionloadoptions#horizontalalignment)|对于集合中的每一项: 表示样式的水平对齐方式。 有关详细信息, 请参阅 HorizontalAlignment。|
||[includeAlignment](/javascript/api/excel/excel.stylecollectionloadoptions#includealignment)|对于集合中的每一项: 指示样式是否包括 AutoIndent、HorizontalAlignment、VerticalAlignment、WrapText、IndentLevel 和 TextOrientation 属性。|
||[includeBorder](/javascript/api/excel/excel.stylecollectionloadoptions#includeborder)|对于集合中的每一项: 指示样式是否包含 Color、ColorIndex、LineStyle 和权数边框属性。|
||[includeFont](/javascript/api/excel/excel.stylecollectionloadoptions#includefont)|对于集合中的每一项: 指示样式是否包括背景、粗体、颜色、ColorIndex、FontStyle、斜体、名称、大小、删除线、下标、上标和下划线字体属性。|
||[includeNumber](/javascript/api/excel/excel.stylecollectionloadoptions#includenumber)|对于集合中的每一项: 指示样式是否包含 NumberFormat 属性。|
||[includePatterns](/javascript/api/excel/excel.stylecollectionloadoptions#includepatterns)|对于集合中的每一项: 指示样式是否包含 Color、ColorIndex、InvertIfNegative、Pattern、PatternColor 和 PatternColorIndex 内部属性。|
||[includeProtection](/javascript/api/excel/excel.stylecollectionloadoptions#includeprotection)|对于集合中的每一项: 指示样式是否包括 FormulaHidden 和锁定保护属性。|
||[indentLevel](/javascript/api/excel/excel.stylecollectionloadoptions#indentlevel)|对于集合中的每一项: 一个介于0到250之间的整数, 它指示样式的缩进量级别。|
||[locked](/javascript/api/excel/excel.stylecollectionloadoptions#locked)|对于集合中的每一项: 指示在工作表处于保护状态时是否锁定对象。|
||[name](/javascript/api/excel/excel.stylecollectionloadoptions#name)|对于集合中的每个项目: 样式的名称。|
||[numberFormat](/javascript/api/excel/excel.stylecollectionloadoptions#numberformat)|对于集合中的每一项: 样式的数字格式的格式代码。|
||[numberFormatLocal](/javascript/api/excel/excel.stylecollectionloadoptions#numberformatlocal)|对于集合中的每一项: 样式的数字格式的本地化格式代码。|
||[readingOrder](/javascript/api/excel/excel.stylecollectionloadoptions#readingorder)|对于集合中的每一项: 样式的阅读顺序。|
||[shrinkToFit](/javascript/api/excel/excel.stylecollectionloadoptions#shrinktofit)|对于集合中的每一项: 指示文本是否自动缩小以适合可用列宽。|
||[verticalAlignment](/javascript/api/excel/excel.stylecollectionloadoptions#verticalalignment)|对于集合中的每一项: 表示样式的垂直对齐方式。 有关详细信息, 请参阅 VerticalAlignment。|
||[wrapText](/javascript/api/excel/excel.stylecollectionloadoptions#wraptext)|对于集合中的每一项: 指示 Microsoft Excel 是否对对象中的文本进行换行。|
|[StyleData](/javascript/api/excel/excel.styledata)|[Borders](/javascript/api/excel/excel.styledata#borders)|四个 Border 对象的 Border 集合，表示四个边框的样式。|
||[内置](/javascript/api/excel/excel.styledata#builtin)|指示样式是否为内置样式。|
||[fill](/javascript/api/excel/excel.styledata#fill)|样式的填充。|
||[font](/javascript/api/excel/excel.styledata#font)|该 Font 对象表示样式的字体。|
||[formulaHidden](/javascript/api/excel/excel.styledata#formulahidden)|指示工作表受保护时是否隐藏公式。|
||[horizontalAlignment](/javascript/api/excel/excel.styledata#horizontalalignment)|表示样式水平对齐。 有关详细信息, 请参阅 HorizontalAlignment。|
||[includeAlignment](/javascript/api/excel/excel.styledata#includealignment)|指示样式是否包含 AutoIndent、HorizontalAlignment、VerticalAlignment、WrapText、IndentLevel 和 TextOrientation 属性。|
||[includeBorder](/javascript/api/excel/excel.styledata#includeborder)|指示样式是否包含 Color、ColorIndex、LineStyle 和 Weight 边框属性。|
||[includeFont](/javascript/api/excel/excel.styledata#includefont)|指示样式是否包含 Background、Bold、Color、ColorIndex、FontStyle、Italic、Name、Size、Strikethrough、Subscrip、Superscript 和 Underline 字体属性。|
||[includeNumber](/javascript/api/excel/excel.styledata#includenumber)|指示样式是否包含 NumberFormat 属性。|
||[includePatterns](/javascript/api/excel/excel.styledata#includepatterns)|指示样式是否包含 Color、ColorIndex、InvertIfNegative、Pattern、PatternColor 和 PatternColorIndex 内部属性。|
||[includeProtection](/javascript/api/excel/excel.styledata#includeprotection)|指示样式是否包含 FormulaHidden 和 Locked 保护属性。|
||[indentLevel](/javascript/api/excel/excel.styledata#indentlevel)|0 到 250 之间的一个整数，指示样式的缩进水平。|
||[locked](/javascript/api/excel/excel.styledata#locked)|指示工作表受保护时是否锁定对象。|
||[name](/javascript/api/excel/excel.styledata#name)|样式的名称。|
||[numberFormat](/javascript/api/excel/excel.styledata#numberformat)|样式中数字格式的格式代码。|
||[numberFormatLocal](/javascript/api/excel/excel.styledata#numberformatlocal)|样式中数字格式的本地化格式代码。|
||[readingOrder](/javascript/api/excel/excel.styledata#readingorder)|样式中的阅读顺序。|
||[shrinkToFit](/javascript/api/excel/excel.styledata#shrinktofit)|指示文本是否自动缩小以适合可用列宽。|
||[verticalAlignment](/javascript/api/excel/excel.styledata#verticalalignment)|表示样式的垂直对齐方式。 有关详细信息, 请参阅 VerticalAlignment。|
||[wrapText](/javascript/api/excel/excel.styledata#wraptext)|指示 Microsoft Excel 是否将对象中的文本换行。|
|[StyleLoadOptions](/javascript/api/excel/excel.styleloadoptions)|[$all](/javascript/api/excel/excel.styleloadoptions#$all)||
||[Borders](/javascript/api/excel/excel.styleloadoptions#borders)|四个 Border 对象的 Border 集合，表示四个边框的样式。|
||[内置](/javascript/api/excel/excel.styleloadoptions#builtin)|指示样式是否为内置样式。|
||[fill](/javascript/api/excel/excel.styleloadoptions#fill)|样式的填充。|
||[font](/javascript/api/excel/excel.styleloadoptions#font)|该 Font 对象表示样式的字体。|
||[formulaHidden](/javascript/api/excel/excel.styleloadoptions#formulahidden)|指示工作表受保护时是否隐藏公式。|
||[horizontalAlignment](/javascript/api/excel/excel.styleloadoptions#horizontalalignment)|表示样式水平对齐。 有关详细信息, 请参阅 HorizontalAlignment。|
||[includeAlignment](/javascript/api/excel/excel.styleloadoptions#includealignment)|指示样式是否包含 AutoIndent、HorizontalAlignment、VerticalAlignment、WrapText、IndentLevel 和 TextOrientation 属性。|
||[includeBorder](/javascript/api/excel/excel.styleloadoptions#includeborder)|指示样式是否包含 Color、ColorIndex、LineStyle 和 Weight 边框属性。|
||[includeFont](/javascript/api/excel/excel.styleloadoptions#includefont)|指示样式是否包含 Background、Bold、Color、ColorIndex、FontStyle、Italic、Name、Size、Strikethrough、Subscrip、Superscript 和 Underline 字体属性。|
||[includeNumber](/javascript/api/excel/excel.styleloadoptions#includenumber)|指示样式是否包含 NumberFormat 属性。|
||[includePatterns](/javascript/api/excel/excel.styleloadoptions#includepatterns)|指示样式是否包含 Color、ColorIndex、InvertIfNegative、Pattern、PatternColor 和 PatternColorIndex 内部属性。|
||[includeProtection](/javascript/api/excel/excel.styleloadoptions#includeprotection)|指示样式是否包含 FormulaHidden 和 Locked 保护属性。|
||[indentLevel](/javascript/api/excel/excel.styleloadoptions#indentlevel)|0 到 250 之间的一个整数，指示样式的缩进水平。|
||[locked](/javascript/api/excel/excel.styleloadoptions#locked)|指示工作表受保护时是否锁定对象。|
||[name](/javascript/api/excel/excel.styleloadoptions#name)|样式的名称。|
||[numberFormat](/javascript/api/excel/excel.styleloadoptions#numberformat)|样式中数字格式的格式代码。|
||[numberFormatLocal](/javascript/api/excel/excel.styleloadoptions#numberformatlocal)|样式中数字格式的本地化格式代码。|
||[readingOrder](/javascript/api/excel/excel.styleloadoptions#readingorder)|样式中的阅读顺序。|
||[shrinkToFit](/javascript/api/excel/excel.styleloadoptions#shrinktofit)|指示文本是否自动缩小以适合可用列宽。|
||[verticalAlignment](/javascript/api/excel/excel.styleloadoptions#verticalalignment)|表示样式的垂直对齐方式。 有关详细信息, 请参阅 VerticalAlignment。|
||[wrapText](/javascript/api/excel/excel.styleloadoptions#wraptext)|指示 Microsoft Excel 是否将对象中的文本换行。|
|[StyleUpdateData](/javascript/api/excel/excel.styleupdatedata)|[Borders](/javascript/api/excel/excel.styleupdatedata#borders)|四个 Border 对象的 Border 集合，表示四个边框的样式。|
||[fill](/javascript/api/excel/excel.styleupdatedata#fill)|样式的填充。|
||[font](/javascript/api/excel/excel.styleupdatedata#font)|该 Font 对象表示样式的字体。|
||[formulaHidden](/javascript/api/excel/excel.styleupdatedata#formulahidden)|指示工作表受保护时是否隐藏公式。|
||[horizontalAlignment](/javascript/api/excel/excel.styleupdatedata#horizontalalignment)|表示样式水平对齐。 有关详细信息, 请参阅 HorizontalAlignment。|
||[includeAlignment](/javascript/api/excel/excel.styleupdatedata#includealignment)|指示样式是否包含 AutoIndent、HorizontalAlignment、VerticalAlignment、WrapText、IndentLevel 和 TextOrientation 属性。|
||[includeBorder](/javascript/api/excel/excel.styleupdatedata#includeborder)|指示样式是否包含 Color、ColorIndex、LineStyle 和 Weight 边框属性。|
||[includeFont](/javascript/api/excel/excel.styleupdatedata#includefont)|指示样式是否包含 Background、Bold、Color、ColorIndex、FontStyle、Italic、Name、Size、Strikethrough、Subscrip、Superscript 和 Underline 字体属性。|
||[includeNumber](/javascript/api/excel/excel.styleupdatedata#includenumber)|指示样式是否包含 NumberFormat 属性。|
||[includePatterns](/javascript/api/excel/excel.styleupdatedata#includepatterns)|指示样式是否包含 Color、ColorIndex、InvertIfNegative、Pattern、PatternColor 和 PatternColorIndex 内部属性。|
||[includeProtection](/javascript/api/excel/excel.styleupdatedata#includeprotection)|指示样式是否包含 FormulaHidden 和 Locked 保护属性。|
||[indentLevel](/javascript/api/excel/excel.styleupdatedata#indentlevel)|0 到 250 之间的一个整数，指示样式的缩进水平。|
||[locked](/javascript/api/excel/excel.styleupdatedata#locked)|指示工作表受保护时是否锁定对象。|
||[numberFormat](/javascript/api/excel/excel.styleupdatedata#numberformat)|样式中数字格式的格式代码。|
||[numberFormatLocal](/javascript/api/excel/excel.styleupdatedata#numberformatlocal)|样式中数字格式的本地化格式代码。|
||[readingOrder](/javascript/api/excel/excel.styleupdatedata#readingorder)|样式中的阅读顺序。|
||[shrinkToFit](/javascript/api/excel/excel.styleupdatedata#shrinktofit)|指示文本是否自动缩小以适合可用列宽。|
||[verticalAlignment](/javascript/api/excel/excel.styleupdatedata#verticalalignment)|表示样式的垂直对齐方式。 有关详细信息, 请参阅 VerticalAlignment。|
||[wrapText](/javascript/api/excel/excel.styleupdatedata#wraptext)|指示 Microsoft Excel 是否将对象中的文本换行。|
|[Table](/javascript/api/excel/excel.table)|[onChanged](/javascript/api/excel/excel.table#onchanged)|当特定表上单元格的数据更改时发生。|
||[onSelectionChanged](/javascript/api/excel/excel.table#onselectionchanged)|当特定表格上的所选内容更改时发生。|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[address](/javascript/api/excel/excel.tablechangedeventargs#address)|获取地址，该地址表示特定工作表上的表格的更改区域。|
||[changeType](/javascript/api/excel/excel.tablechangedeventargs#changetype)|获取更改类型，该类型表示 Changed 事件的触发方式。 有关详细信息, 请参阅 DataChangeType。|
||[source](/javascript/api/excel/excel.tablechangedeventargs#source)|获取事件源。 有关详细信息，请参阅 Excel.EventSource。|
||[tableId](/javascript/api/excel/excel.tablechangedeventargs#tableid)|获取其中的数据发生更改的表格的 ID。|
||[type](/javascript/api/excel/excel.tablechangedeventargs#type)|获取事件的类型。 有关详细信息，请参阅 Excel.EventType。|
||[worksheetId](/javascript/api/excel/excel.tablechangedeventargs#worksheetid)|获取其中的数据发生更改的工作表的 ID。|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onChanged](/javascript/api/excel/excel.tablecollection#onchanged)|当工作簿中的任何表或工作表上的数据发生更改时发生。|
|[TableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|[address](/javascript/api/excel/excel.tableselectionchangedeventargs#address)|获取区域地址，该地址表示特定工作表上的表格选定区域。|
||[isInsideTable](/javascript/api/excel/excel.tableselectionchangedeventargs#isinsidetable)|指示选定区域是否在表格内，如果 IsInsideTable 为 false，则地址无效。|
||[tableId](/javascript/api/excel/excel.tableselectionchangedeventargs#tableid)|获取其中的选定区域发生更改的表格 ID。|
||[type](/javascript/api/excel/excel.tableselectionchangedeventargs#type)|获取事件的类型。 有关详细信息，请参阅 Excel.EventType。 只读。|
||[worksheetId](/javascript/api/excel/excel.tableselectionchangedeventargs#worksheetid)|获取其中的选定区域发生更改的工作表的 ID。|
|[Workbook](/javascript/api/excel/excel.workbook)|[getActiveCell()](/javascript/api/excel/excel.workbook#getactivecell--)|获取工作簿中当前处于活动状态的单元格。|
||[dataConnections](/javascript/api/excel/excel.workbook#dataconnections)|表示工作簿中的所有数据连接。 只读。|
||[name](/javascript/api/excel/excel.workbook#name)|获取工作簿名称。 只读。|
||[properties](/javascript/api/excel/excel.workbook#properties)|获取工作簿属性。 只读。|
||[protection](/javascript/api/excel/excel.workbook#protection)|返回工作簿的工作簿保护对象。 只读。|
||[styles](/javascript/api/excel/excel.workbook#styles)|表示与工作簿关联的样式的集合。 只读。|
|[WorkbookData](/javascript/api/excel/excel.workbookdata)|[name](/javascript/api/excel/excel.workbookdata#name)|获取工作簿名称。 只读。|
||[properties](/javascript/api/excel/excel.workbookdata#properties)|获取工作簿属性。 只读。|
||[protection](/javascript/api/excel/excel.workbookdata#protection)|返回工作簿的工作簿保护对象。 只读。|
||[styles](/javascript/api/excel/excel.workbookdata#styles)|表示与工作簿关联的样式的集合。 只读。|
|[WorkbookLoadOptions](/javascript/api/excel/excel.workbookloadoptions)|[name](/javascript/api/excel/excel.workbookloadoptions#name)|获取工作簿名称。 只读。|
||[properties](/javascript/api/excel/excel.workbookloadoptions#properties)|获取工作簿属性。|
||[protection](/javascript/api/excel/excel.workbookloadoptions#protection)|返回工作簿的工作簿保护对象。|
|[WorkbookProtection](/javascript/api/excel/excel.workbookprotection)|[保护 (password？: string)](/javascript/api/excel/excel.workbookprotection#protect-password-)|保护工作簿。 如果工作簿处于受保护状态，则无法执行此方法。|
||[受保护](/javascript/api/excel/excel.workbookprotection#protected)|指示工作簿是否受保护。 只读。|
||[取消保护 (password？: string)](/javascript/api/excel/excel.workbookprotection#unprotect-password-)|解除保护工作簿。|
|[WorkbookProtectionData](/javascript/api/excel/excel.workbookprotectiondata)|[受保护](/javascript/api/excel/excel.workbookprotectiondata#protected)|指示工作簿是否受保护。 只读。|
|[WorkbookProtectionLoadOptions](/javascript/api/excel/excel.workbookprotectionloadoptions)|[$all](/javascript/api/excel/excel.workbookprotectionloadoptions#$all)||
||[受保护](/javascript/api/excel/excel.workbookprotectionloadoptions#protected)|指示工作簿是否受保护。 只读。|
|[WorkbookUpdateData](/javascript/api/excel/excel.workbookupdatedata)|[properties](/javascript/api/excel/excel.workbookupdatedata#properties)|获取工作簿属性。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[副本 (positionType？: \| "在\| " "开始" \| \|结束 "之后的" "结束", relativeTo？: Excel. 工作表)](/javascript/api/excel/excel.worksheet#copy-positiontype--relativeto-)|复制工作表并将其置于指定位置。 返回复制的工作表。|
||[copy (positionType？: WorksheetPositionType, relativeTo？: Excel. 工作表)](/javascript/api/excel/excel.worksheet#copy-positiontype--relativeto-)|复制工作表并将其置于指定位置。 返回复制的工作表。|
||[getRangeByIndexes (startRow: number, startColumn: number, rowCount: number, columnCount: number)](/javascript/api/excel/excel.worksheet#getrangebyindexes-startrow--startcolumn--rowcount--columncount-)|获取以特定行索引和列索引开始并跨越了一定数量的行和列的 range 对象。|
||[freezePanes](/javascript/api/excel/excel.worksheet#freezepanes)|获取一个对象, 该对象可用于操作工作表上的冻结窗格。 只读。|
||[onActivated](/javascript/api/excel/excel.worksheet#onactivated)|当激活工作表时发生此事件。|
||[onChanged](/javascript/api/excel/excel.worksheet#onchanged)|当指定的工作表上的数据发生更改时发生。|
||[onDeactivated](/javascript/api/excel/excel.worksheet#ondeactivated)|停用工作表时发生此事件。|
||[onSelectionChanged](/javascript/api/excel/excel.worksheet#onselectionchanged)|当指定的工作表上的所选内容更改时发生。|
||[standardHeight](/javascript/api/excel/excel.worksheet#standardheight)|返回工作表中所有行的标准（默认）行高，以磅为单位。 只读。|
||[standardWidth](/javascript/api/excel/excel.worksheet#standardwidth)|返回或设置工作表中所有列的标准（默认）列宽。|
||[tabColor](/javascript/api/excel/excel.worksheet#tabcolor)|获取或设置工作表标签颜色。|
|[WorksheetActivatedEventArgs](/javascript/api/excel/excel.worksheetactivatedeventargs)|[type](/javascript/api/excel/excel.worksheetactivatedeventargs#type)|获取事件的类型。 有关详细信息，请参阅 Excel.EventType。|
||[worksheetId](/javascript/api/excel/excel.worksheetactivatedeventargs#worksheetid)|获取已启用的工作表的 ID。|
|[WorksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|[source](/javascript/api/excel/excel.worksheetaddedeventargs#source)|获取事件源。 有关详细信息，请参阅 Excel.EventSource。|
||[type](/javascript/api/excel/excel.worksheetaddedeventargs#type)|获取事件的类型。 有关详细信息，请参阅 Excel.EventType。|
||[worksheetId](/javascript/api/excel/excel.worksheetaddedeventargs#worksheetid)|获取已添加至工作簿的工作表的 ID。|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[address](/javascript/api/excel/excel.worksheetchangedeventargs#address)|获取区域地址，该地址表示特定工作表上的更改区域。|
||[changeType](/javascript/api/excel/excel.worksheetchangedeventargs#changetype)|获取更改类型，该类型表示 Changed 事件的触发方式。 有关详细信息, 请参阅 DataChangeType。|
||[source](/javascript/api/excel/excel.worksheetchangedeventargs#source)|获取事件源。 有关详细信息，请参阅 Excel.EventSource。|
||[type](/javascript/api/excel/excel.worksheetchangedeventargs#type)|获取事件的类型。 有关详细信息，请参阅 Excel.EventType。|
||[worksheetId](/javascript/api/excel/excel.worksheetchangedeventargs#worksheetid)|获取其中的数据发生更改的工作表的 ID。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onActivated](/javascript/api/excel/excel.worksheetcollection#onactivated)|当工作簿中的任何工作表激活时发生此事件。|
||[onAdded](/javascript/api/excel/excel.worksheetcollection#onadded)|将新工作表添加到工作簿时发生此事件。|
||[onDeactivated](/javascript/api/excel/excel.worksheetcollection#ondeactivated)|当工作簿中的任何工作表被停用时发生。|
||[onDeleted](/javascript/api/excel/excel.worksheetcollection#ondeleted)|当从工作簿中删除工作表时发生此事件。|
|[WorksheetCollectionLoadOptions](/javascript/api/excel/excel.worksheetcollectionloadoptions)|[standardHeight](/javascript/api/excel/excel.worksheetcollectionloadoptions#standardheight)|对于集合中的每一项: 返回工作表中所有行的标准 (默认) 高度 (以磅为单位)。 只读。|
||[standardWidth](/javascript/api/excel/excel.worksheetcollectionloadoptions#standardwidth)|对于集合中的每一项: 返回或设置工作表中所有列的标准 (默认) 宽度。|
||[tabColor](/javascript/api/excel/excel.worksheetcollectionloadoptions#tabcolor)|对于集合中的每一项: 获取或设置工作表标签颜色。|
|[WorksheetData](/javascript/api/excel/excel.worksheetdata)|[standardHeight](/javascript/api/excel/excel.worksheetdata#standardheight)|返回工作表中所有行的标准（默认）行高，以磅为单位。 只读。|
||[standardWidth](/javascript/api/excel/excel.worksheetdata#standardwidth)|返回或设置工作表中所有列的标准（默认）列宽。|
||[tabColor](/javascript/api/excel/excel.worksheetdata#tabcolor)|获取或设置工作表标签颜色。|
|[WorksheetDeactivatedEventArgs](/javascript/api/excel/excel.worksheetdeactivatedeventargs)|[type](/javascript/api/excel/excel.worksheetdeactivatedeventargs#type)|获取事件的类型。 有关详细信息，请参阅 Excel.EventType。|
||[worksheetId](/javascript/api/excel/excel.worksheetdeactivatedeventargs#worksheetid)|获取已停用的工作表的 ID。|
|[WorksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|[source](/javascript/api/excel/excel.worksheetdeletedeventargs#source)|获取事件源。 有关详细信息，请参阅 Excel.EventSource。|
||[type](/javascript/api/excel/excel.worksheetdeletedeventargs#type)|获取事件的类型。 有关详细信息，请参阅 Excel.EventType。|
||[worksheetId](/javascript/api/excel/excel.worksheetdeletedeventargs#worksheetid)|获取已从工作簿删除的工作表的 ID。|
|[WorksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|[freezeAt (frozenRange: Range \|字符串)](/javascript/api/excel/excel.worksheetfreezepanes#freezeat-frozenrange-)|设置活动工作表视图中的冻结单元格。|
||[freezeColumns (count？: 数字)](/javascript/api/excel/excel.worksheetfreezepanes#freezecolumns-count-)|就地冻结工作表的第一列。|
||[freezeRows (count？: 数字)](/javascript/api/excel/excel.worksheetfreezepanes#freezerows-count-)|就地冻结工作表的顶行。|
||[getLocation()](/javascript/api/excel/excel.worksheetfreezepanes#getlocation--)|获取用于描述活动工作表视图中的冻结单元格的区域。|
||[getLocationOrNullObject()](/javascript/api/excel/excel.worksheetfreezepanes#getlocationornullobject--)|获取用于描述活动工作表视图中的冻结单元格的区域。|
||[解冻 ()](/javascript/api/excel/excel.worksheetfreezepanes#unfreeze--)|移除工作表中的所有冻结窗格。|
|[WorksheetLoadOptions](/javascript/api/excel/excel.worksheetloadoptions)|[standardHeight](/javascript/api/excel/excel.worksheetloadoptions#standardheight)|返回工作表中所有行的标准（默认）行高，以磅为单位。 只读。|
||[standardWidth](/javascript/api/excel/excel.worksheetloadoptions#standardwidth)|返回或设置工作表中所有列的标准（默认）列宽。|
||[tabColor](/javascript/api/excel/excel.worksheetloadoptions#tabcolor)|获取或设置工作表标签颜色。|
|[WorksheetProtection](/javascript/api/excel/excel.worksheetprotection)|[取消保护 (password？: string)](/javascript/api/excel/excel.worksheetprotection#unprotect-password-)|解除对 worksheet 的保护。|
|[WorksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|[allowEditObjects](/javascript/api/excel/excel.worksheetprotectionoptions#alloweditobjects)|表示允许编辑对象的工作表保护选项。|
||[allowEditScenarios](/javascript/api/excel/excel.worksheetprotectionoptions#alloweditscenarios)|表示允许编辑应用场景的工作表保护选项。|
||[selectionMode](/javascript/api/excel/excel.worksheetprotectionoptions#selectionmode)|表示选择模式的工作表保护选项。|
|[WorksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|[address](/javascript/api/excel/excel.worksheetselectionchangedeventargs#address)|获取区域地址，该地址表示特定工作表上的选定区域。|
||[type](/javascript/api/excel/excel.worksheetselectionchangedeventargs#type)|获取事件的类型。 有关详细信息，请参阅 Excel.EventType。|
||[worksheetId](/javascript/api/excel/excel.worksheetselectionchangedeventargs#worksheetid)|获取其中的选定区域发生更改的工作表的 ID。|
|[WorksheetUpdateData](/javascript/api/excel/excel.worksheetupdatedata)|[standardWidth](/javascript/api/excel/excel.worksheetupdatedata#standardwidth)|返回或设置工作表中所有列的标准（默认）列宽。|
||[tabColor](/javascript/api/excel/excel.worksheetupdatedata#tabcolor)|获取或设置工作表标签颜色。|

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 参考文档](/javascript/api/excel)
- [Excel JavaScript API 要求集](./excel-api-requirement-sets.md)
