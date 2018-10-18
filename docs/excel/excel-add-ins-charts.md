---
title: 使用 Excel JavaScript API 处理图表
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 80b537ec66caf6e173dfe4453a257c5963156e6f
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459299"
---
# <a name="work-with-charts-using-the-excel-javascript-api"></a>使用 Excel JavaScript API 处理图表

本文提供了显示如何使用 Excel JavaScript API 的图表使用执行常见任务的代码示例。有关的属性和方法的 **图表** 和 **ChartCollection** 对象支持的完整列表，请参阅 [Chart 对象 (Excel 的 JavaScript API)](https://docs.microsoft.com/javascript/api/excel/excel.chart?view=office-js) 和 [图表集合对象 (Excel 的 JavaScript API)](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection?view=office-js)。

## <a name="create-a-chart"></a>创建图表

下面的代码示例创建名为 **示例**工作表中的图表。 **取决于** A1:B13 **范围中的数据的折线图**图表。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var dataRange = sheet.getRange("A1:B13");
    var chart = sheet.charts.add("Line", dataRange, "auto");

    chart.title.text = "Sales Data";
    chart.legend.position = "right"
    chart.legend.format.fill.setSolidColor("white");
    chart.dataLabels.format.font.size = 15;
    chart.dataLabels.format.font.color = "black";

    return context.sync();
}).catch(errorHandlerFunction);
```

**新建折线图**

![Excel 中的新折线图](../images/excel-charts-create-line.png)


## <a name="add-a-data-series-to-a-chart"></a>向图表添加数据系列

下面的代码示例向工作表中的第一个图表的数据系列。新数据系列对应于名为 **2016年** 的列，并取决于 **D2:D5**范围中的数据。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var chart = sheet.charts.getItemAt(0);
    var dataRange = sheet.getRange("D2:D5");

    var newSeries = chart.series.add("2016");
    newSeries.setValues(dataRange);

    return context.sync();
}).catch(errorHandlerFunction);
```

**添加 2016 数据系列之前的图表**

![Excel 中添加 2016 数据系列之前的图表](../images/excel-charts-data-series-before.png)

**添加 2016 数据系列之后的图表**

![Excel 中添加 2016 数据系列之后的图表](../images/excel-charts-data-series-after.png)

## <a name="set-chart-title"></a>设置图表标题

下面的代码示例将工作表中的第一个图表标题设置为**年度销售数据**。 

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.title.text = "Sales Data by Year";

    return context.sync();
}).catch(errorHandlerFunction);
```

**设置标题后的图表**

![Excel 中带标题的图表](../images/excel-charts-title-set.png)

## <a name="set-properties-of-an-axis-in-a-chart"></a>在图表中设置轴属性

使用[笛卡儿坐标系统](https://en.wikipedia.org/wiki/Cartesian_coordinate_system)的图表（如柱形图、条形图和散点图）包含分类轴和数值轴。 以下示例介绍如何设置图表中轴的标题和显示单位。

### <a name="set-axis-title"></a>设置轴标题

下面的代码示例将工作表中第一个图表的分类轴标题设置为**产品**。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.axes.categoryAxis.title.text = "Product";

    return context.sync();
}).catch(errorHandlerFunction);
```

**设置分类轴标题后的图表**

![Excel 中带轴标题的图表](../images/excel-charts-axis-title-set.png)

### <a name="set-axis-display-unit"></a>设置轴的显示单位

下面的代码示例将工作表中首个图表的数值轴显示单位设置为“百”****。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.axes.valueAxis.displayUnit = "Hundreds";

    return context.sync();
}).catch(errorHandlerFunction);
```

**设置数值轴显示单位后的图表**

![Excel 中带轴显示单位的图表](../images/excel-charts-axis-display-unit-set.png)

## <a name="set-visibility-of-gridlines-in-a-chart"></a>在图表中设置网格线的可见性

以下代码示例隐藏工作表中第一个图表数值轴的主要网格线。 可以通过将 `chart.axes.valueAxis.majorGridlines.visible` 设置为 **true**，显示图表数值轴的主要网格线。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.axes.valueAxis.majorGridlines.visible = false;

    return context.sync();
}).catch(errorHandlerFunction);
```

**隐藏了网格线的图表**

![Excel 中隐藏了网格线的图表](../images/excel-charts-gridlines-removed.png)

## <a name="chart-trendlines"></a>图表趋势线

### <a name="add-a-trendline"></a>添加趋势线

下面的代码示例向 **Sample** 工作表中首个图表的第一个系列添加移动均线。趋势线显示超过 5 个周期的移动平均。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    var seriesCollection = chart.series;
    seriesCollection.getItemAt(0).trendlines.add("MovingAverage").movingAveragePeriod = 5;

    return context.sync();
}).catch(errorHandlerFunction);
```

**带移动均线的图表**

![Excel 中带移动均线的图表](../images/excel-charts-create-trendline.png)

### <a name="update-a-trendline"></a>更新趋势线

下面的代码示例将 **Sample** 工作表中首个图表的第一个系列的趋势线设置为“线性”**** 类型。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    var seriesCollection = chart.series;
    var series = seriesCollection.getItemAt(0);
    series.trendlines.getItem(0).type = "Linear";

    return context.sync();
}).catch(errorHandlerFunction);
```

**带线性趋势线的图表**

![Excel 中带线性趋势线的图表](../images/excel-charts-trendline-linear.png)

## <a name="see-also"></a>另请参阅

- [使用 Excel JavaScript API 的基本编程概念](excel-add-ins-core-concepts.md)
