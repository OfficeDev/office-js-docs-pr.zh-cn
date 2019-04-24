---
title: 使用 Excel JavaScript API 处理数据透视表
description: 使用 Excel JavaScript API 创建数据透视表并与其组件进行交互。
ms.date: 03/21/2019
localization_priority: Normal
ms.openlocfilehash: b53d734e676417a6438f1008bac720a38a244d1f
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32449346"
---
# <a name="work-with-pivottables-using-the-excel-javascript-api"></a><span data-ttu-id="dbe55-103">使用 Excel JavaScript API 处理数据透视表</span><span class="sxs-lookup"><span data-stu-id="dbe55-103">Work with PivotTables using the Excel JavaScript API</span></span>

<span data-ttu-id="dbe55-104">数据透视表精简了较大的数据集。</span><span class="sxs-lookup"><span data-stu-id="dbe55-104">PivotTables streamline larger data sets.</span></span> <span data-ttu-id="dbe55-105">它们允许快速操作分组数据。</span><span class="sxs-lookup"><span data-stu-id="dbe55-105">They allow the quick manipulation of grouped data.</span></span> <span data-ttu-id="dbe55-106">Excel JavaScript API 允许你的外接程序创建数据透视表并与其组件进行交互。</span><span class="sxs-lookup"><span data-stu-id="dbe55-106">The Excel JavaScript API lets your add-in create PivotTables and interact with their components.</span></span>

<span data-ttu-id="dbe55-107">如果您对数据透视表的功能不熟悉, 请考虑将其作为最终用户来浏览。</span><span class="sxs-lookup"><span data-stu-id="dbe55-107">If you are unfamiliar with the functionality of PivotTables, consider exploring them as an end user.</span></span> <span data-ttu-id="dbe55-108">有关这些工具的最佳入门知识, 请参阅[创建数据透视表以分析工作表数据](https://support.office.com/en-us/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables)。</span><span class="sxs-lookup"><span data-stu-id="dbe55-108">See [Create a PivotTable to analyze worksheet data](https://support.office.com/en-us/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) for a good primer on these tools.</span></span> 

<span data-ttu-id="dbe55-109">本文提供了常见方案的代码示例。</span><span class="sxs-lookup"><span data-stu-id="dbe55-109">This article provides code samples for common scenarios.</span></span> <span data-ttu-id="dbe55-110">若要进一步了解数据透视表 API, 请参阅[**数据透视表**](/javascript/api/excel/excel.pivottable)和[**PivotTableCollection**](/javascript/api/excel/excel.pivottable)。</span><span class="sxs-lookup"><span data-stu-id="dbe55-110">To further your understanding of the PivotTable API, see [**PivotTable**](/javascript/api/excel/excel.pivottable) and [**PivotTableCollection**](/javascript/api/excel/excel.pivottable).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="dbe55-111">目前不支持使用 OLAP 创建的数据透视表。</span><span class="sxs-lookup"><span data-stu-id="dbe55-111">PivotTables created with OLAP are not currently supported.</span></span> <span data-ttu-id="dbe55-112">此外, 也不支持 Power Pivot。</span><span class="sxs-lookup"><span data-stu-id="dbe55-112">There is also no support for Power Pivot.</span></span>

## <a name="hierarchies"></a><span data-ttu-id="dbe55-113">Hierarchies</span><span class="sxs-lookup"><span data-stu-id="dbe55-113">Hierarchies</span></span>

<span data-ttu-id="dbe55-114">数据透视表基于四种层次结构类别进行组织: 行、列、数据和筛选器。</span><span class="sxs-lookup"><span data-stu-id="dbe55-114">PivotTables are organized based on four hierarchy categories: row, column, data, and filter.</span></span> <span data-ttu-id="dbe55-115">在本文中, 将使用从各个服务器场中描述水果销售的以下数据。</span><span class="sxs-lookup"><span data-stu-id="dbe55-115">The following data describing fruit sales from various farms will be used throughout this article.</span></span>

![来自不同服务器场的不同类型的水果销售的集合。](../images/excel-pivots-raw-data.png)

<span data-ttu-id="dbe55-117">此数据具有五个层次\*\*\*\* 结构: 服务器场、**类型**、**分类**、**服务器场中销售的 Crates**和**Crates 销售批发**。</span><span class="sxs-lookup"><span data-stu-id="dbe55-117">This data has five hierarchies: **Farms**, **Type**, **Classification**, **Crates Sold at Farm**, and **Crates Sold Wholesale**.</span></span> <span data-ttu-id="dbe55-118">每个层次结构只能存在于四个类别之一中。</span><span class="sxs-lookup"><span data-stu-id="dbe55-118">Each hierarchy can only exist in one of the four categories.</span></span> <span data-ttu-id="dbe55-119">如果**Type**添加到列层次结构中, 然后添加到行层次结构中, 则它仅保留在后者中。</span><span class="sxs-lookup"><span data-stu-id="dbe55-119">If **Type** is added to column hierarchies and then added to row hierarchies, it only remains in the latter.</span></span>

<span data-ttu-id="dbe55-120">行和列层次结构定义数据的分组方式。</span><span class="sxs-lookup"><span data-stu-id="dbe55-120">Row and column hierarchies define how data will be grouped.</span></span> <span data-ttu-id="dbe55-121">例如,**服务器场**的行层次结构将把来自同一个服务器场的所有数据集组合在一起。</span><span class="sxs-lookup"><span data-stu-id="dbe55-121">For example, a row hierarchy of **Farms** will group together all the data sets from the same farm.</span></span> <span data-ttu-id="dbe55-122">在行和列层次结构之间进行选择, 以定义数据透视表的方向。</span><span class="sxs-lookup"><span data-stu-id="dbe55-122">The choice between row and column hierarchy defines the orientation of the PivotTable.</span></span>

<span data-ttu-id="dbe55-123">数据层次结构是要根据行和列层次结构聚合的值。</span><span class="sxs-lookup"><span data-stu-id="dbe55-123">Data hierarchies are the values to be aggregated based on the row and column hierarchies.</span></span> <span data-ttu-id="dbe55-124">具有**服务器场**的行层次结构和**Crates 销售**的数据层次结构的数据透视表显示每个服务器场的所有不同 fruits 的总和总计 (默认值)。</span><span class="sxs-lookup"><span data-stu-id="dbe55-124">A PivotTable with a row hierarchy of **Farms** and a data hierarchy of **Crates Sold Wholesale** shows the sum total (by default) of all the different fruits for each farm.</span></span>

<span data-ttu-id="dbe55-125">筛选器层次结构基于该筛选类型中的值包括或排除数据透视表中的数据。</span><span class="sxs-lookup"><span data-stu-id="dbe55-125">Filter hierarchies include or exclude data from the pivot based on values within that filtered type.</span></span> <span data-ttu-id="dbe55-126">选定类型为 "**有机**" 的**分类**筛选器层次结构仅显示用于随机水果的数据。</span><span class="sxs-lookup"><span data-stu-id="dbe55-126">A filter hierarchy of **Classification** with the type **Organic** selected only shows data for organic fruit.</span></span>

<span data-ttu-id="dbe55-127">下面是数据透视表旁边的服务器场数据。</span><span class="sxs-lookup"><span data-stu-id="dbe55-127">Here is the farm data again, alongside a PivotTable.</span></span> <span data-ttu-id="dbe55-128">数据透视表使用**服务器场**和**类型**作为行层次结构,**在服务器场中售出的 Crates**和**Crates 销售批发**作为数据层次结构 (具有 sum 的默认聚合函数) 和**分类**作为筛选器层次结构 (选择了**随机**选择的层次结构)。</span><span class="sxs-lookup"><span data-stu-id="dbe55-128">The PivotTable is using **Farm** and **Type** as the row hierarchies, **Crates Sold at Farm** and **Crates Sold Wholesale** as the data hierarchies (with the default aggregation function of sum), and **Classification** as a filter hierarchy (with **Organic** selected).</span></span> 

![选择了具有行、数据和筛选器层次结构的数据透视表旁边的水果销售数据。](../images/excel-pivot-table-and-data.png)

<span data-ttu-id="dbe55-130">此数据透视表可通过 JavaScript API 或 Excel UI 生成。</span><span class="sxs-lookup"><span data-stu-id="dbe55-130">This PivotTable could be generated through the JavaScript API or through the Excel UI.</span></span> <span data-ttu-id="dbe55-131">这两个选项都允许通过外接程序进行进一步操作。</span><span class="sxs-lookup"><span data-stu-id="dbe55-131">Both options allow for further manipulation through add-ins.</span></span>

## <a name="create-a-pivottable"></a><span data-ttu-id="dbe55-132">创建数据透视表</span><span class="sxs-lookup"><span data-stu-id="dbe55-132">Create a PivotTable</span></span>

<span data-ttu-id="dbe55-133">数据透视表需要名称、源和目标。</span><span class="sxs-lookup"><span data-stu-id="dbe55-133">PivotTables need a name, source, and destination.</span></span> <span data-ttu-id="dbe55-134">源可以是区域地址或表名称 (作为`Range`、 `string`或`Table`类型传递)。</span><span class="sxs-lookup"><span data-stu-id="dbe55-134">The source can be a range address or table name (passed as a `Range`, `string`, or `Table` type).</span></span> <span data-ttu-id="dbe55-135">目标是区域地址 (指定为`Range`或`string`)。</span><span class="sxs-lookup"><span data-stu-id="dbe55-135">The destination is a range address (given as either a `Range` or `string`).</span></span> <span data-ttu-id="dbe55-136">下面的示例展示了各种数据透视表创建技术。</span><span class="sxs-lookup"><span data-stu-id="dbe55-136">The following samples show various PivotTable creation techniques.</span></span>

### <a name="create-a-pivottable-with-range-addresses"></a><span data-ttu-id="dbe55-137">创建包含区域地址的数据透视表</span><span class="sxs-lookup"><span data-stu-id="dbe55-137">Create a PivotTable with range addresses</span></span>

```typescript
await Excel.run(async (context) => {
    // creating a PivotTable named "Farm Sales" on the current worksheet at cell A22 with data from the range A1:E21
    context.workbook.worksheets.getActiveWorksheet().pivotTables.add("Farm Sales", "A1:E21", "A22");

    await context.sync();
});
```

### <a name="create-a-pivottable-with-range-objects"></a><span data-ttu-id="dbe55-138">创建包含 Range 对象的数据透视表</span><span class="sxs-lookup"><span data-stu-id="dbe55-138">Create a PivotTable with Range objects</span></span>

```typescript
await Excel.run(async (context) => {
    // creating a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data comes from the worksheet "DataWorksheet" across the range A1:E21
    const rangeToAnalyze = context.workbook.worksheets.getItem("DataWorksheet").getRange("A1:E21");
    const rangeToPlacePivot = context.workbook.worksheets.getItem("PivotWorksheet").getRange("A2");
    context.workbook.worksheets.getItem("PivotWorksheet").pivotTables.add(
        "Farm Sales", rangeToAnalyze, rangeToPlacePivot);

    await context.sync();
});
```

### <a name="create-a-pivottable-at-the-workbook-level"></a><span data-ttu-id="dbe55-139">在工作簿级别创建数据透视表</span><span class="sxs-lookup"><span data-stu-id="dbe55-139">Create a PivotTable at the workbook level</span></span>

```typescript
await Excel.run(async (context) => {
    // creating a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data is from the worksheet "DataWorksheet" across the range A1:E21
    context.workbook.pivotTables.add("Farm Sales", "DataWorksheet!A1:E21", "PivotWorksheet!A2");

    await context.sync();
});
```

## <a name="use-an-existing-pivottable"></a><span data-ttu-id="dbe55-140">使用现有数据透视表</span><span class="sxs-lookup"><span data-stu-id="dbe55-140">Use an existing PivotTable</span></span>

<span data-ttu-id="dbe55-141">手动创建的数据透视表也可通过工作簿或单个工作表的数据透视表集合进行访问。</span><span class="sxs-lookup"><span data-stu-id="dbe55-141">Manually created PivotTables are also accessible through the PivotTable collection of the workbook or of individual worksheets.</span></span> 

<span data-ttu-id="dbe55-142">下面的代码获取工作簿中的第一个数据透视表。</span><span class="sxs-lookup"><span data-stu-id="dbe55-142">The following code gets the first PivotTable in the workbook.</span></span> <span data-ttu-id="dbe55-143">然后, 它将为表提供一个名称, 以便日后参考。</span><span class="sxs-lookup"><span data-stu-id="dbe55-143">It then gives the table a name for easy future reference.</span></span>

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.pivotTables.getItem("My Pivot");
    await context.sync();
});
```

## <a name="add-rows-and-columns-to-a-pivottable"></a><span data-ttu-id="dbe55-144">向数据透视表添加行和列</span><span class="sxs-lookup"><span data-stu-id="dbe55-144">Add rows and columns to a PivotTable</span></span>

<span data-ttu-id="dbe55-145">数据透视这些字段的值周围的行和列。</span><span class="sxs-lookup"><span data-stu-id="dbe55-145">Rows and columns pivot the data around those fields’ values.</span></span>

<span data-ttu-id="dbe55-146">添加 "**服务器场**" 列将每个服务器场的所有销售额枢轴分布。</span><span class="sxs-lookup"><span data-stu-id="dbe55-146">Adding the **Farm** column pivots all the sales around each farm.</span></span> <span data-ttu-id="dbe55-147">添加 "**类型**" 和 "**分类**" 行会根据所售的水果和是否为 "有随机" 来进一步细分数据。</span><span class="sxs-lookup"><span data-stu-id="dbe55-147">Adding the **Type** and **Classification** rows further breaks down the data based on what fruit was sold and whether it was organic or not.</span></span>

![具有服务器场列和类型和分类行的数据透视表。](../images/excel-pivots-table-rows-and-columns.png)

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));

    pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem("Farm"));

    await context.sync();
});
```

<span data-ttu-id="dbe55-149">您还可以拥有仅包含行或列的数据透视表。</span><span class="sxs-lookup"><span data-stu-id="dbe55-149">You can also have a PivotTable with only rows or columns.</span></span>

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));

    await context.sync();
});
```

## <a name="add-data-hierarchies-to-the-pivottable"></a><span data-ttu-id="dbe55-150">向数据透视表中添加数据层次结构</span><span class="sxs-lookup"><span data-stu-id="dbe55-150">Add data hierarchies to the PivotTable</span></span>

<span data-ttu-id="dbe55-151">数据层次结构使用要基于行和列进行组合的信息填充数据透视表。</span><span class="sxs-lookup"><span data-stu-id="dbe55-151">Data hierarchies fill the PivotTable with information to combine based on the rows and columns.</span></span> <span data-ttu-id="dbe55-152">添加 Crates 的数据层次结构**在服务器场**和**Crates**售出销售批发为每个行和列提供这些数字的总和。</span><span class="sxs-lookup"><span data-stu-id="dbe55-152">Adding the data hierarchies of **Crates Sold at Farm** and **Crates Sold Wholesale** gives sums of those figures for each row and column.</span></span> 

<span data-ttu-id="dbe55-153">在示例中, "**服务器场**" 和 "**类型**" 都是行, 而 "发货箱销售额" 作为数据。</span><span class="sxs-lookup"><span data-stu-id="dbe55-153">In the example, both **Farm** and **Type** are rows, with the crate sales as the data.</span></span> 

![显示基于其来源的服务器场的不同水果的总销售额的数据透视表。](../images/excel-pivots-data-hierarchy.png)

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    // "Farm" and "Type" are the hierarchies on which the aggregation is based
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));

    // "Crates Sold at Farm" and "Crates Sold Wholesale" are the hierarchies
    // that will have their data aggregated (summed in this case)
    pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Crates Sold at Farm"));
    pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Crates Sold Wholesale"));

    await context.sync();
});
```

## <a name="change-aggregation-function"></a><span data-ttu-id="dbe55-155">更改聚合函数</span><span class="sxs-lookup"><span data-stu-id="dbe55-155">Change aggregation function</span></span>

<span data-ttu-id="dbe55-156">数据层次结构的值已聚合。</span><span class="sxs-lookup"><span data-stu-id="dbe55-156">Data hierarchies have their values aggregated.</span></span> <span data-ttu-id="dbe55-157">对于数字的数据集, 默认情况下, 这是一个总和。</span><span class="sxs-lookup"><span data-stu-id="dbe55-157">For datasets of numbers, this is a sum by default.</span></span> <span data-ttu-id="dbe55-158">该`summarizeBy`属性基于[AggregationFunction](/javascript/api/excel/excel.aggregationfunction)类型定义此行为。</span><span class="sxs-lookup"><span data-stu-id="dbe55-158">The `summarizeBy` property defines this behavior based on an [AggregationFunction](/javascript/api/excel/excel.aggregationfunction) type.</span></span>

<span data-ttu-id="dbe55-159">当前支持的聚合函数类型为`Sum`、 `Count` `Average` `Max` `Min` `Product` `CountNumbers` `StandardDeviation` `StandardDeviationP` `Variance`、、、、、、、、、和`Automatic` (默认值`VarianceP`)。</span><span class="sxs-lookup"><span data-stu-id="dbe55-159">The currently supported aggregation function types are `Sum`, `Count`, `Average`, `Max`, `Min`, `Product`, `CountNumbers`, `StandardDeviation`, `StandardDeviationP`, `Variance`, `VarianceP`, and `Automatic` (the default).</span></span>

<span data-ttu-id="dbe55-160">下面的代码示例将聚合更改为数据的平均值。</span><span class="sxs-lookup"><span data-stu-id="dbe55-160">The following code samples changes the aggregation to be averages of the data.</span></span>

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.dataHierarchies.load("no-properties-needed");
    await context.sync();

    // changing the aggregation from the default sum to an average of all the values in the hierarchy
    pivotTable.dataHierarchies.items[0].summarizeBy = Excel.AggregationFunction.average;
    pivotTable.dataHierarchies.items[1].summarizeBy = Excel.AggregationFunction.average;
    await context.sync();
});
```

## <a name="change-calculations-with-a-showasrule"></a><span data-ttu-id="dbe55-161">使用 ShowAsRule 更改计算</span><span class="sxs-lookup"><span data-stu-id="dbe55-161">Change calculations with a ShowAsRule</span></span>

<span data-ttu-id="dbe55-162">默认情况下, 数据透视表将单独聚合其行和列层次结构的数据。</span><span class="sxs-lookup"><span data-stu-id="dbe55-162">PivotTables, by default, aggregate the data of their row and column hierarchies independently.</span></span> <span data-ttu-id="dbe55-163">[ShowAsRule](/javascript/api/excel/excel.showasrule)将数据层次结构更改为基于数据透视表中的其他项的输出值。</span><span class="sxs-lookup"><span data-stu-id="dbe55-163">A [ShowAsRule](/javascript/api/excel/excel.showasrule) changes the data hierarchy to output values based on other items in the PivotTable.</span></span>

<span data-ttu-id="dbe55-164">`ShowAsRule`对象具有三个属性:</span><span class="sxs-lookup"><span data-stu-id="dbe55-164">The `ShowAsRule` object has three properties:</span></span>

-   <span data-ttu-id="dbe55-165">`calculation`: 要应用于数据层次结构的相对计算的类型 (默认值为`none`)。</span><span class="sxs-lookup"><span data-stu-id="dbe55-165">`calculation`: The type of relative calculation to apply to the data hierarchy (the default is `none`).</span></span>
-   <span data-ttu-id="dbe55-166">`baseField`: 层次结构中的字段, 其中包含在应用计算之前的基础数据。</span><span class="sxs-lookup"><span data-stu-id="dbe55-166">`baseField`: The field within the hierarchy containing the base data before the calculation is applied.</span></span> <span data-ttu-id="dbe55-167">[透视字段](/javascript/api/excel/excel.pivotfield)的名称通常与其父层次结构的名称相同。</span><span class="sxs-lookup"><span data-stu-id="dbe55-167">The [PivotField](/javascript/api/excel/excel.pivotfield) usually has the same name as its parent hierarchy.</span></span>
-   <span data-ttu-id="dbe55-168">`baseItem`: 个人[PivotItem](/javascript/api/excel/excel.pivotitem)根据计算类型与基本字段的值进行比较。</span><span class="sxs-lookup"><span data-stu-id="dbe55-168">`baseItem`: The individual [PivotItem](/javascript/api/excel/excel.pivotitem) compared against the values of the base fields based on the calculation type.</span></span> <span data-ttu-id="dbe55-169">并非所有计算都需要此字段。</span><span class="sxs-lookup"><span data-stu-id="dbe55-169">Not all calculations require this field.</span></span>

<span data-ttu-id="dbe55-170">以下示例将场数据层次结构中的 " **Crates**总数" 的计算设置为列总计的百分比。</span><span class="sxs-lookup"><span data-stu-id="dbe55-170">The following example sets the calculation on the **Sum of Crates Sold at Farm** data hierarchy to be a percentage of the column total.</span></span> <span data-ttu-id="dbe55-171">我们仍希望将粒度扩展到水果类型级别, 因此我们将使用**类型**行层次结构及其基础字段。</span><span class="sxs-lookup"><span data-stu-id="dbe55-171">We still want the granularity to extend to the fruit type level, so we’ll use the **Type** row hierarchy and its underlying field.</span></span> <span data-ttu-id="dbe55-172">该示例还将**服务器场**作为第一个行的层次结构, 因此服务器场总数将显示每个服务器场也负责生成的百分比。</span><span class="sxs-lookup"><span data-stu-id="dbe55-172">The example also has **Farm** as the first row hierarchy, so the farm total entries display the percentage each farm is responsible for producing as well.</span></span>

![显示与每个场中的单个服务器场和各个水果类型的总和相关的水果销售百分比的数据透视表。](../images/excel-pivots-showas-percentage.png)

``` TypeScript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    const farmDataHierarchy = pivotTable.dataHierarchies.getItem("Sum of Crates Sold at Farm");

    farmDataHierarchy.load("showAs");
    await context.sync();

    // show the crates of each fruit type sold at the farm as a percentage of the column's total
    let farmShowAs = farmDataHierarchy.showAs;
    farmShowAs.calculation = Excel.ShowAsCalculation.percentOfColumnTotal;
    farmShowAs.baseField = pivotTable.rowHierarchies.getItem("Type").fields.getItem("Type");
    farmDataHierarchy.showAs = farmShowAs; 
    farmDataHierarchy.name = "Percentage of Total Farm Sales";

    await context.sync();
});
```

<span data-ttu-id="dbe55-174">上面的示例将计算设置为相对于单个行层次结构的列。</span><span class="sxs-lookup"><span data-stu-id="dbe55-174">The previous example set the calculation to the column, relative to an individual row hierarchy.</span></span> <span data-ttu-id="dbe55-175">当计算与单个项目相关时, 请使用`baseItem`属性。</span><span class="sxs-lookup"><span data-stu-id="dbe55-175">When the calculation relates to an individual item, use the `baseItem` property.</span></span>

<span data-ttu-id="dbe55-176">下面的示例演示了`differenceFrom`计算。</span><span class="sxs-lookup"><span data-stu-id="dbe55-176">The following example shows the `differenceFrom` calculation.</span></span> <span data-ttu-id="dbe55-177">它显示服务器场与 "服务器场" 相关的 "销售数据" 层次结构条目的差异。</span><span class="sxs-lookup"><span data-stu-id="dbe55-177">It displays the difference of the farm crate sales data hierarchy entries relative to those of “A Farms”.</span></span>
<span data-ttu-id="dbe55-178">`baseField`是**服务器场**, 因此我们看到其他服务器场之间的差异, 以及每种类型的类似水果 (在此示例中**类型**也是行层次结构) 的细目。</span><span class="sxs-lookup"><span data-stu-id="dbe55-178">The `baseField` is **Farm**, so we see the differences between the other farms, as well as breakdowns for each type of like fruit (**Type** is also a row hierarchy in this example).</span></span>

![显示 "一群" 和其他 "服务器场" 之间的水果销售差异的数据透视表。](../images/excel-pivots-showas-differencefrom.png)

``` TypeScript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    const farmDataHierarchy = pivotTable.dataHierarchies.getItem("Sum of Crates Sold at Farm");

    farmDataHierarchy.load("showAs");
    await context.sync();

    // show the difference between crate sales of the "A Farms" and the other farms
    // this difference is both aggregated and shown for individual fruit types (where applicable)
    let farmShowAs = farmDataHierarchy.showAs;
    farmShowAs.calculation = Excel.ShowAsCalculation.differenceFrom;
    farmShowAs.baseField = pivotTable.rowHierarchies.getItem("Farm").fields.getItem("Farm");
    farmShowAs.baseItem = pivotTable.rowHierarchies.getItem("Farm").fields.getItem("Farm").items.getItem("A Farms");
    farmDataHierarchy.showAs = farmShowAs;
    farmDataHierarchy.name = "Difference from A Farms";
    await context.sync();
});
```

## <a name="pivottable-layouts"></a><span data-ttu-id="dbe55-182">数据透视表布局</span><span class="sxs-lookup"><span data-stu-id="dbe55-182">PivotTable layouts</span></span>

<span data-ttu-id="dbe55-183">[PivotLayout](/javascript/api/excel/excel.pivotlayout)定义层次结构及其数据的位置。</span><span class="sxs-lookup"><span data-stu-id="dbe55-183">A [PivotLayout](/javascript/api/excel/excel.pivotlayout) defines the placement of hierarchies and their data.</span></span> <span data-ttu-id="dbe55-184">您可以访问布局以确定存储数据的区域。</span><span class="sxs-lookup"><span data-stu-id="dbe55-184">You access the layout to determine the ranges where data is stored.</span></span>

<span data-ttu-id="dbe55-185">下图显示了哪些布局函数调用对应于数据透视表的区域。</span><span class="sxs-lookup"><span data-stu-id="dbe55-185">The following diagram shows which layout function calls correspond to which ranges of the PivotTable.</span></span>

![显示由布局的 get range 函数返回的数据透视表的节的图表。](../images/excel-pivots-layout-breakdown.png)

<span data-ttu-id="dbe55-187">下面的代码演示如何通过布局获取数据透视表数据的最后一行。</span><span class="sxs-lookup"><span data-stu-id="dbe55-187">The following code demonstrates how to get the last row of the PivotTable data by going through the layout.</span></span> <span data-ttu-id="dbe55-188">然后将这些值汇总到一起以进行总计。</span><span class="sxs-lookup"><span data-stu-id="dbe55-188">Those values are then summed together for a grand total.</span></span>

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    // get the totals for each data hierarchy from the layout
    const range = pivotTable.layout.getDataBodyRange();
    const grandTotalRange = range.getLastRow();
    grandTotalRange.load("address");
    await context.sync();

    // sum the totals from the PivotTable data hierarchies and place them in a new range
    const masterTotalRange = context.workbook.worksheets.getActiveWorksheet().getRange("B27:C27");
    masterTotalRange.formulas = [["All Crates", "=SUM(" + grandTotalRange.address + ")"]];
    await context.sync();
});
```

<span data-ttu-id="dbe55-189">数据透视表具有三种布局样式: 紧凑、大纲和表格。</span><span class="sxs-lookup"><span data-stu-id="dbe55-189">PivotTables have three layout styles: Compact, Outline, and Tabular.</span></span> <span data-ttu-id="dbe55-190">我们在前面的示例中看到了压缩样式。</span><span class="sxs-lookup"><span data-stu-id="dbe55-190">We’ve seen the compact style in the previous examples.</span></span> 

<span data-ttu-id="dbe55-191">下面的示例分别使用大纲样式和表格样式。</span><span class="sxs-lookup"><span data-stu-id="dbe55-191">The following examples use the outline and tabular styles, respectively.</span></span> <span data-ttu-id="dbe55-192">此代码示例演示如何在不同的布局之间循环。</span><span class="sxs-lookup"><span data-stu-id="dbe55-192">The code sample shows how to cycle between the different layouts.</span></span>

### <a name="outline-layout"></a><span data-ttu-id="dbe55-193">大纲布局</span><span class="sxs-lookup"><span data-stu-id="dbe55-193">Outline layout</span></span>

![使用大纲布局的数据透视表。](../images/excel-pivots-outline-layout.png)

### <a name="tabular-layout"></a><span data-ttu-id="dbe55-195">表格布局</span><span class="sxs-lookup"><span data-stu-id="dbe55-195">Tabular layout</span></span>

![使用表格布局的数据透视表。](../images/excel-pivots-tabular-layout.png)

## <a name="change-hierarchy-names"></a><span data-ttu-id="dbe55-197">更改层次结构名称</span><span class="sxs-lookup"><span data-stu-id="dbe55-197">Change hierarchy names</span></span>

<span data-ttu-id="dbe55-198">层次结构字段是可编辑的。</span><span class="sxs-lookup"><span data-stu-id="dbe55-198">Hierarchy fields are editable.</span></span> <span data-ttu-id="dbe55-199">下面的代码演示如何更改两个数据层次结构的显示名称。</span><span class="sxs-lookup"><span data-stu-id="dbe55-199">The following code demonstrates how to change the displayed names of two data hierarchies.</span></span>

```typescript
await Excel.run(async (context) => {
    const dataHierarchies = context.workbook.worksheets.getActiveWorksheet()
        .pivotTables.getItem("Farm Sales").dataHierarchies;
    dataHierarchies.load("no-properties-needed");
    await context.sync();

    // changing the displayed names of these entries
    dataHierarchies.items[0].name = "Farm Sales";
    dataHierarchies.items[1].name = "Wholesale";
    await context.sync();
});
```

## <a name="delete-a-pivottable"></a><span data-ttu-id="dbe55-200">删除数据透视表</span><span class="sxs-lookup"><span data-stu-id="dbe55-200">Delete a PivotTable</span></span>

<span data-ttu-id="dbe55-201">使用它们的名称删除数据透视表。</span><span class="sxs-lookup"><span data-stu-id="dbe55-201">PivotTables are deleted by using their name.</span></span>

```typescript
await Excel.run(async (context) => {
    context.workbook.worksheets.getItem("Pivot").pivotTables.getItem("Farm Sales").delete();

    await context.sync();
});
```

## <a name="see-also"></a><span data-ttu-id="dbe55-202">另请参阅</span><span class="sxs-lookup"><span data-stu-id="dbe55-202">See also</span></span>

- [<span data-ttu-id="dbe55-203">Excel JavaScript API 基本编程概念</span><span class="sxs-lookup"><span data-stu-id="dbe55-203">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="dbe55-204">Excel JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="dbe55-204">Excel JavaScript API Reference</span></span>](/javascript/api/excel)
