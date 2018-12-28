---
title: 使用 Excel JavaScript API 处理图表
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: b804e2130e30626a9caf21bca1f3955c57a3f94c
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/28/2018
ms.locfileid: "27457549"
---
# <a name="work-with-charts-using-the-excel-javascript-api"></a><span data-ttu-id="72599-102">使用 Excel JavaScript API 处理图表</span><span class="sxs-lookup"><span data-stu-id="72599-102">Work with Charts using the Excel JavaScript API</span></span>

<span data-ttu-id="72599-103">本文提供了代码示例，介绍如何使用 Excel JavaScript API 对图表执行常见任务。</span><span class="sxs-lookup"><span data-stu-id="72599-103">This article provides code samples that show how to perform common tasks with charts using the Excel JavaScript API.</span></span> <span data-ttu-id="72599-104">有关 **Chart** 和 **ChartCollection** 对象支持的属性和方法的完整列表，请参阅 [Chart 对象 (Excel JavaScript API)](https://docs.microsoft.com/javascript/api/excel/excel.chart) 和 [Chart Collection 对象 (Excel JavaScript API)](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection)。</span><span class="sxs-lookup"><span data-stu-id="72599-104">For the complete list of properties and methods that the **Chart** and **ChartCollection** objects support, see [Chart Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.chart) and [Chart Collection Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection).</span></span>

## <a name="create-a-chart"></a><span data-ttu-id="72599-105">创建图表</span><span class="sxs-lookup"><span data-stu-id="72599-105">Create a chart</span></span>

<span data-ttu-id="72599-106">下面的代码示例在名为 **Sample** 的工作表中创建一个图表。</span><span class="sxs-lookup"><span data-stu-id="72599-106">The following code sample creates a chart in the worksheet named **Sample**.</span></span> <span data-ttu-id="72599-107">该图表是基于区域 **A1:B13** 的数据的**折线**图。</span><span class="sxs-lookup"><span data-stu-id="72599-107">The chart is a **Line** chart that is based upon data in the range **A1:B13**.</span></span>

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

<span data-ttu-id="72599-108">**新建折线图**</span><span class="sxs-lookup"><span data-stu-id="72599-108">**New line chart**</span></span>

![Excel 中的新折线图](../images/excel-charts-create-line.png)


## <a name="add-a-data-series-to-a-chart"></a><span data-ttu-id="72599-110">向图表添加数据系列</span><span class="sxs-lookup"><span data-stu-id="72599-110">Add a data series to a chart</span></span>

<span data-ttu-id="72599-111">下面的代码示例向工作表中的第一个图表添加数据系列。</span><span class="sxs-lookup"><span data-stu-id="72599-111">The following code sample adds a data series to the first chart in the worksheet.</span></span> <span data-ttu-id="72599-112">新的数据系列对应于“2016 年”\*\*\*\* 列，并以区域 **D2:D5** 中的数据为依据。</span><span class="sxs-lookup"><span data-stu-id="72599-112">The new data series corresponds to the column named **2016** and is based upon data in the range **D2:D5**.</span></span>

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

<span data-ttu-id="72599-113">**添加 2016 数据系列之前的图表**</span><span class="sxs-lookup"><span data-stu-id="72599-113">**Chart before the 2016 data series is added**</span></span>

![Excel 中添加 2016 数据系列之前的图表](../images/excel-charts-data-series-before.png)

<span data-ttu-id="72599-115">**添加 2016 数据系列之后的图表**</span><span class="sxs-lookup"><span data-stu-id="72599-115">**Chart after the 2016 data series is added**</span></span>

![Excel 中添加 2016 数据系列之后的图表](../images/excel-charts-data-series-after.png)

## <a name="set-chart-title"></a><span data-ttu-id="72599-117">设置图表标题</span><span class="sxs-lookup"><span data-stu-id="72599-117">Set chart title</span></span>

<span data-ttu-id="72599-118">下面的代码示例将工作表中的第一个图表标题设置为**年度销售数据**。</span><span class="sxs-lookup"><span data-stu-id="72599-118">The following code sample sets the title of the first chart in the worksheet to **Sales Data by Year**.</span></span> 

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.title.text = "Sales Data by Year";

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="72599-119">**设置标题后的图表**</span><span class="sxs-lookup"><span data-stu-id="72599-119">**Chart after title is set**</span></span>

![Excel 中带标题的图表](../images/excel-charts-title-set.png)

## <a name="set-properties-of-an-axis-in-a-chart"></a><span data-ttu-id="72599-121">在图表中设置轴属性</span><span class="sxs-lookup"><span data-stu-id="72599-121">Set properties of an axis in a chart</span></span>

<span data-ttu-id="72599-122">使用[笛卡儿坐标系统](https://en.wikipedia.org/wiki/Cartesian_coordinate_system)的图表（如柱形图、条形图和散点图）包含分类轴和数值轴。</span><span class="sxs-lookup"><span data-stu-id="72599-122">Charts that use the [Cartesian coordinate system](https://en.wikipedia.org/wiki/Cartesian_coordinate_system) such as column charts, bar charts, and scatter charts contain a category axis and a value axis.</span></span> <span data-ttu-id="72599-123">以下示例介绍如何设置图表中轴的标题和显示单位。</span><span class="sxs-lookup"><span data-stu-id="72599-123">These examples show how to set the title and display unit of an axis in a chart.</span></span>

### <a name="set-axis-title"></a><span data-ttu-id="72599-124">设置轴标题</span><span class="sxs-lookup"><span data-stu-id="72599-124">Set axis title</span></span>

<span data-ttu-id="72599-125">下面的代码示例将工作表中第一个图表的分类轴标题设置为**产品**。</span><span class="sxs-lookup"><span data-stu-id="72599-125">The following code sample sets the title of the category axis for the first chart in the worksheet to **Product**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.axes.categoryAxis.title.text = "Product";

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="72599-126">**设置分类轴标题后的图表**</span><span class="sxs-lookup"><span data-stu-id="72599-126">**Chart after title of category axis is set**</span></span>

![Excel 中带轴标题的图表](../images/excel-charts-axis-title-set.png)

### <a name="set-axis-display-unit"></a><span data-ttu-id="72599-128">设置轴的显示单位</span><span class="sxs-lookup"><span data-stu-id="72599-128">Set axis display unit</span></span>

<span data-ttu-id="72599-129">下面的代码示例将工作表中首个图表的数值轴显示单位设置为“百”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="72599-129">The following code sample sets the display unit of the value axis for the first chart in the worksheet to **Hundreds**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.axes.valueAxis.displayUnit = "Hundreds";

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="72599-130">**设置数值轴显示单位后的图表**</span><span class="sxs-lookup"><span data-stu-id="72599-130">**Chart after display unit of value axis is set**</span></span>

![Excel 中带轴显示单位的图表](../images/excel-charts-axis-display-unit-set.png)

## <a name="set-visibility-of-gridlines-in-a-chart"></a><span data-ttu-id="72599-132">在图表中设置网格线的可见性</span><span class="sxs-lookup"><span data-stu-id="72599-132">Set visibility of gridlines in a chart</span></span>

<span data-ttu-id="72599-133">以下代码示例隐藏工作表中第一个图表数值轴的主要网格线。</span><span class="sxs-lookup"><span data-stu-id="72599-133">The following code sample hides the major gridlines for the value axis of the first chart in the worksheet.</span></span> <span data-ttu-id="72599-134">可以通过将 `chart.axes.valueAxis.majorGridlines.visible` 设置为 **true**，显示图表数值轴的主要网格线。</span><span class="sxs-lookup"><span data-stu-id="72599-134">You can show the major gridlines for the value axis of the chart, by setting `chart.axes.valueAxis.majorGridlines.visible` to **true**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.axes.valueAxis.majorGridlines.visible = false;

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="72599-135">**隐藏了网格线的图表**</span><span class="sxs-lookup"><span data-stu-id="72599-135">**Chart with gridlines hidden**</span></span>

![Excel 中隐藏了网格线的图表](../images/excel-charts-gridlines-removed.png)

## <a name="chart-trendlines"></a><span data-ttu-id="72599-137">图表趋势线</span><span class="sxs-lookup"><span data-stu-id="72599-137">Chart trendlines</span></span>

### <a name="add-a-trendline"></a><span data-ttu-id="72599-138">添加趋势线</span><span class="sxs-lookup"><span data-stu-id="72599-138">Add a trendline</span></span>

<span data-ttu-id="72599-p106">下面的代码示例向 **Sample** 工作表中首个图表的第一个系列添加移动均线。趋势线显示超过 5 个周期的移动平均。</span><span class="sxs-lookup"><span data-stu-id="72599-p106">The following code sample adds a moving average trendline to the first series in the first chart in the worksheet named **Sample**. The trendline shows a moving average over 5 periods.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    var seriesCollection = chart.series;
    seriesCollection.getItemAt(0).trendlines.add("MovingAverage").movingAveragePeriod = 5;

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="72599-141">**带移动均线的图表**</span><span class="sxs-lookup"><span data-stu-id="72599-141">**Chart with moving average trendline**</span></span>

![Excel 中带移动均线的图表](../images/excel-charts-create-trendline.png)

### <a name="update-a-trendline"></a><span data-ttu-id="72599-143">更新趋势线</span><span class="sxs-lookup"><span data-stu-id="72599-143">Update a trendline</span></span>

<span data-ttu-id="72599-144">下面的代码示例将 **Sample** 工作表中首个图表的第一个系列的趋势线设置为“线性”\*\*\*\* 类型。</span><span class="sxs-lookup"><span data-stu-id="72599-144">The following code sample sets the trendline to type **Linear** for the first series in the first chart in the worksheet named **Sample**.</span></span>

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

<span data-ttu-id="72599-145">**带线性趋势线的图表**</span><span class="sxs-lookup"><span data-stu-id="72599-145">**Chart with linear trendline**</span></span>

![Excel 中带线性趋势线的图表](../images/excel-charts-trendline-linear.png)

## <a name="see-also"></a><span data-ttu-id="72599-147">另请参阅</span><span class="sxs-lookup"><span data-stu-id="72599-147">See also</span></span>

- [<span data-ttu-id="72599-148">Excel JavaScript API 基本编程概念</span><span class="sxs-lookup"><span data-stu-id="72599-148">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
