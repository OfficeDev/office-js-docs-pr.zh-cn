---
title: 使用 Excel JavaScript API 处理数据透视表
description: 使用 Excel JavaScript API 创建数据透视表并与其组件交互。
ms.date: 09/21/2018
ms.openlocfilehash: 5245665bad2933df205bcda29e226a965de1c356
ms.sourcegitcommit: 64da9ed76d22b14df745b1f0ef97a8f5194400e4
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/03/2018
ms.locfileid: "25361022"
---
# <a name="work-with-pivottables-using-the-excel-javascript-api"></a><span data-ttu-id="423b8-103">使用 Excel JavaScript API 处理数据透视表</span><span class="sxs-lookup"><span data-stu-id="423b8-103">Work with ranges using the Excel JavaScript API</span></span>

<span data-ttu-id="423b8-104">数据透视表可简化更大的数据集。</span><span class="sxs-lookup"><span data-stu-id="423b8-104">PivotTables streamline larger data sets.</span></span> <span data-ttu-id="423b8-105">它们允许分组数据的快速操作。</span><span class="sxs-lookup"><span data-stu-id="423b8-105">They allow the quick manipulation of grouped data.</span></span> <span data-ttu-id="423b8-106">Excel JavaScript API 允许加载项创建数据透视表并与其组件交互。</span><span class="sxs-lookup"><span data-stu-id="423b8-106">The Excel JavaScript API lets your add-in create PivotTables and interact with their components.</span></span> 

<span data-ttu-id="423b8-107">如果不熟悉数据透视表的功能，尝试以最终用户的身份了解它们的功能。</span><span class="sxs-lookup"><span data-stu-id="423b8-107">If you are unfamiliar with the functionality of PivotTables, consider exploring them as an end-user.</span></span> <span data-ttu-id="423b8-108">请参阅[创建数据透视表分析工作表数据](https://support.office.com/en-us/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables)，了解这些工具的入门指导。</span><span class="sxs-lookup"><span data-stu-id="423b8-108">See [Create a PivotTable to analyze worksheet data](https://support.office.com/en-us/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) for a good primer on these tools.</span></span> 

<span data-ttu-id="423b8-109">本文提供了常见方案的代码示例。</span><span class="sxs-lookup"><span data-stu-id="423b8-109">This article provides code samples for common scenarios.</span></span> <span data-ttu-id="423b8-110">参阅[**PivotTable**](https://docs.microsoft.com/javascript/api/excel/excel.pivottable) 和 [**PivotTableCollection**](https://docs.microsoft.com/javascript/api/excel/excel.pivottable)，加深对数据透视表 API 的理解。</span><span class="sxs-lookup"><span data-stu-id="423b8-110">To further your understanding of the PivotTable API, see [**PivotTable**](https://docs.microsoft.com/javascript/api/excel/excel.pivottable) and [**PivotTableCollection**](https://docs.microsoft.com/javascript/api/excel/excel.pivottable).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="423b8-111">目前不支持使用 OLAP 创建的数据透视表。</span><span class="sxs-lookup"><span data-stu-id="423b8-111">PivotTables created with OLAP are not currently supported.</span></span>

## <a name="hierarchies"></a><span data-ttu-id="423b8-112">层次结构</span><span class="sxs-lookup"><span data-stu-id="423b8-112">Hierarchies</span></span>

<span data-ttu-id="423b8-113">数据透视表基于四种层次结构类别构成：行、列、数据和筛选器。</span><span class="sxs-lookup"><span data-stu-id="423b8-113">PivotTables are organized based on four hierarchy categories: row, column, data, and filter.</span></span> <span data-ttu-id="423b8-114">本文将通篇使用以下描述各农场水果销售情况的数据。</span><span class="sxs-lookup"><span data-stu-id="423b8-114">The following data describing fruit sales from various farms will be used throughout this article.</span></span>

![来自不同农场的不同类型水果销售的集合。](../images/excel-pivots-raw-data.png)

<span data-ttu-id="423b8-116">此数据具有五个层次结构：**Farm**、**Type**、**Classification**、**Crates Sold at Farm** 和 **Crates Sold Wholesale**。</span><span class="sxs-lookup"><span data-stu-id="423b8-116">This data has five hierarchies: **Farms**, **Type**, **Classification**, **Crates Sold at Farm**, and **Crates Sold Wholesale**.</span></span> <span data-ttu-id="423b8-117">每个层次结构只能存在于四个类别中的一个类别。</span><span class="sxs-lookup"><span data-stu-id="423b8-117">Each hierarchy can only exist in one of the four categories.</span></span> <span data-ttu-id="423b8-118">如果 **Type** 添加到列层次结构，然后又添加到行层次结构，则其仅保留于后者。</span><span class="sxs-lookup"><span data-stu-id="423b8-118">If **Type** is added to column hierarchies and then added to row hierarchies, it only remains in the latter.</span></span>

<span data-ttu-id="423b8-119">行和列的层次结构定义如何分组数据。</span><span class="sxs-lookup"><span data-stu-id="423b8-119">Row and column hierarchies define how data will be grouped.</span></span> <span data-ttu-id="423b8-120">例如，**Farms** 的行层次结构会将来自同一农场的所有数据集归集在一起。</span><span class="sxs-lookup"><span data-stu-id="423b8-120">For example, a row hierarchy of **Farms** will group together all the data sets from the same farm.</span></span> <span data-ttu-id="423b8-121">选择行和列层次结构来定义数据透视表的方向。</span><span class="sxs-lookup"><span data-stu-id="423b8-121">The choice between row and column hierarchy defines the orientation of the PivotTable.</span></span>

<span data-ttu-id="423b8-122">数据层次结构是基于行和列层次结构进行聚合的值。</span><span class="sxs-lookup"><span data-stu-id="423b8-122">Data hierarchies are the values to be aggregated based on the row and column hierarchies.</span></span> <span data-ttu-id="423b8-123">具有 **Farms** 的行层次结构和 **Crates Sold Wholesale** 的数据层次结构的数据透视表显示每个农场所有不同水果的总和（默认）。</span><span class="sxs-lookup"><span data-stu-id="423b8-123">A PivotTable with a row hierarchy of **Farms** and a data hierarchy of **Crates Sold Wholesale** shows the sum total (by default) of all the different fruits for each farm.</span></span>

<span data-ttu-id="423b8-124">筛选器层次结构基于已筛选类型中的值包含或排除来自透视的数据。</span><span class="sxs-lookup"><span data-stu-id="423b8-124">Filter hierarchies include or exclude data from the pivot based on values within that filtered type.</span></span> <span data-ttu-id="423b8-125">选择了 **Organic** 类型的 **Classification** 筛选器层次结构仅显示有机水果的数据。</span><span class="sxs-lookup"><span data-stu-id="423b8-125">A filter hierarchy of **Classification** with the type **Organic** selected only shows data for organic fruit.</span></span>

<span data-ttu-id="423b8-126">这同样是农场数据，一旁是数据透视表。</span><span class="sxs-lookup"><span data-stu-id="423b8-126">Here is the farm data again, alongside a PivotTable.</span></span> <span data-ttu-id="423b8-127">数据透视表使用 **Farm** 和 **Type** 作为行层次结构，**Crates Sold at Farm** 和 **Crates Sold Wholesale** 作为数据层次结构 （带默认的 sum 汇总函数），**Classification** 作为筛选器层次结构（选中 **Organic**）。</span><span class="sxs-lookup"><span data-stu-id="423b8-127">The PivotTable is using **Farm** and **Type** as the row hierarchies, **Crates Sold at Farm** and **Crates Sold Wholesale** as the data hierarchies (with the default aggregation function of sum), and **Classification** as a filter hierarchy (with **Organic** selected).</span></span> 

![具有行、数据和筛选器层次结构的数据透视表旁边的水果销售数据的选定内容。](../images/excel-pivot-table-and-data.png)

<span data-ttu-id="423b8-129">可通过 JavaScript API 或通过 Excel UI 生成这个数据透视表。</span><span class="sxs-lookup"><span data-stu-id="423b8-129">This PivotTable could be generated through the JavaScript API or through the Excel UI.</span></span> <span data-ttu-id="423b8-130">两个选项均可通过加载项实现进一步的操作。</span><span class="sxs-lookup"><span data-stu-id="423b8-130">Both options allow for further manipulation through add-ins.</span></span>

## <a name="create-a-pivottable"></a><span data-ttu-id="423b8-131">创建数据透视表</span><span class="sxs-lookup"><span data-stu-id="423b8-131">Create a PivotTable and PivotChart</span></span>

<span data-ttu-id="423b8-132">数据透视表需要有名称、源和目标。</span><span class="sxs-lookup"><span data-stu-id="423b8-132">PivotTables need a name, source, and destination.</span></span> <span data-ttu-id="423b8-133">源可以是范围地址或表名（作为 `Range`、`string` 或 `Table`类型传递）。</span><span class="sxs-lookup"><span data-stu-id="423b8-133">The source can be a range address or table name (passed as a `Range`, `string`, or `Table` type).</span></span> <span data-ttu-id="423b8-134">目标是某一范围地址（作为`Range` 或 `string`给定）。</span><span class="sxs-lookup"><span data-stu-id="423b8-134">The destination is a range address (given as either a `Range` or `string`).</span></span> <span data-ttu-id="423b8-135">以下示例显示各种数据透视表的创建技术。</span><span class="sxs-lookup"><span data-stu-id="423b8-135">The following samples show various PivotTable creation techniques.</span></span>

### <a name="create-a-pivottable-with-range-addresses"></a><span data-ttu-id="423b8-136">创建带范围地址的数据透视表</span><span class="sxs-lookup"><span data-stu-id="423b8-136">Create a PivotTable with range addresses</span></span>

```typescript
await Excel.run(async (context) => {
    // creating a PivotTable named "Farm Sales" on the current worksheet at cell A22 with data from the range A1:E21
    context.workbook.worksheets.getActiveWorksheet().pivotTables.add("Farm Sales", "A1:E21", "A22");

    await context.sync();
});
```

### <a name="create-a-pivottable-with-range-objects"></a><span data-ttu-id="423b8-137">创建带范围对象的数据透视表</span><span class="sxs-lookup"><span data-stu-id="423b8-137">Create a PivotTable with Range objects</span></span>

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

### <a name="create-a-pivottable-at-the-workbook-level"></a><span data-ttu-id="423b8-138">在工作簿级别创建数据透视表</span><span class="sxs-lookup"><span data-stu-id="423b8-138">Create a PivotTable at the workbook level</span></span>

```typescript
await Excel.run(async (context) => {
    // creating a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data is from the worksheet "DataWorksheet" across the range A1:E21
    context.workbook.pivotTables.add("Farm Sales", "DataWorksheet!A1:E21", "PivotWorksheet!A2");

    await context.sync();
});
```

## <a name="use-an-existing-pivottable"></a><span data-ttu-id="423b8-139">使用现有数据透视表</span><span class="sxs-lookup"><span data-stu-id="423b8-139">Use an existing PivotTable</span></span>

<span data-ttu-id="423b8-140">手动创建的数据透视表亦可通过工作簿或者单个工作表的数据透视表集合进行访问。</span><span class="sxs-lookup"><span data-stu-id="423b8-140">Manually created PivotTables are also accessible through the PivotTable collection of the workbook or of individual worksheets.</span></span> 

<span data-ttu-id="423b8-141">以下代码获取工作簿中的第一个数据透视表。</span><span class="sxs-lookup"><span data-stu-id="423b8-141">The following code gets the first PivotTable in the workbook.</span></span> <span data-ttu-id="423b8-142">然后给出了表的名称，便于以后参考。</span><span class="sxs-lookup"><span data-stu-id="423b8-142">It then gives the table a name for easy future reference.</span></span>

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.pivotTables.getItem("My Pivot");
    await context.sync();
});
```

## <a name="add-rows-and-columns-to-a-pivottable"></a><span data-ttu-id="423b8-143">向数据透视表添加行和列</span><span class="sxs-lookup"><span data-stu-id="423b8-143">Add rows and columns to a PivotTable</span></span>

<span data-ttu-id="423b8-144">行和列按字段值相关的方式透视数据。</span><span class="sxs-lookup"><span data-stu-id="423b8-144">Rows and columns pivot the data around those fields’ values.</span></span>

<span data-ttu-id="423b8-145">添加 **Farm** 列可按每个农场的所有销售情况透视数据。</span><span class="sxs-lookup"><span data-stu-id="423b8-145">Adding the **Farm** column pivots all the sales around each farm.</span></span> <span data-ttu-id="423b8-146">添加 **Type** 和 **Classification** 行，可基于销售的水果以及该水果是否为有机等条件而将数据作进一步的分解。</span><span class="sxs-lookup"><span data-stu-id="423b8-146">Adding the **Type** and **Classification** rows further breaks down the data based on what fruit was sold and whether it was organic or not.</span></span>

![带 Farm 列以及 Type 和 Classification 行的数据透视表。](../images/excel-pivots-table-rows-and-columns.png)

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));
    
    pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem("Farm"));

    await context.sync();
});
```

<span data-ttu-id="423b8-148">还可拥有仅带行或列的数据透视表。</span><span class="sxs-lookup"><span data-stu-id="423b8-148">You can also have a PivotTable with only rows or columns.</span></span>

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));
    
    await context.sync();
});
```

## <a name="add-data-hierarchies-to-the-pivottable"></a><span data-ttu-id="423b8-149">向数据透视表添加数据层次结构</span><span class="sxs-lookup"><span data-stu-id="423b8-149">Add data hierarchies to the PivotTable</span></span>

<span data-ttu-id="423b8-150">基于行和列，数据层次结构结合信息填充数据透视表。</span><span class="sxs-lookup"><span data-stu-id="423b8-150">Data hierarchies fill the PivotTable with information to combine based on the rows and columns.</span></span> <span data-ttu-id="423b8-151">添加**Crates Sold at Farm** 和 **Crates Sold Wholesale** 的数据层次结构给出每行和每列数字的总和。</span><span class="sxs-lookup"><span data-stu-id="423b8-151">Adding the data hierarchies of **Crates Sold at Farm** and **Crates Sold Wholesale** gives sums of those figures for each row and column.</span></span> 

<span data-ttu-id="423b8-152">在示例中，**Farm** 和 **Type** 都是行，销售箱数作为数据。</span><span class="sxs-lookup"><span data-stu-id="423b8-152">In the example, both **Farm** and **Type** are rows, with the crate sales as the data.</span></span> 

![显示基于水果来源农场的不同水果总销售情况的数据透视表。](../images/excel-pivots-data-hierarchy.png)

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

## <a name="change-aggregation-function"></a><span data-ttu-id="423b8-154">更改汇总函数</span><span class="sxs-lookup"><span data-stu-id="423b8-154">Change aggregation function</span></span>

<span data-ttu-id="423b8-155">数据层次结构将其数值汇总。</span><span class="sxs-lookup"><span data-stu-id="423b8-155">Data hierarchies have their values aggregated.</span></span> <span data-ttu-id="423b8-156">对于数字的数据集，默认情况下，这是总和。</span><span class="sxs-lookup"><span data-stu-id="423b8-156">For datasets of numbers, this is a sum by default.</span></span> <span data-ttu-id="423b8-157">`summarizeBy` 属性基于 `AggregrationFunction` 类型定义此行为。</span><span class="sxs-lookup"><span data-stu-id="423b8-157">The `summarizeBy` property defines this behavior based on an `AggregrationFunction` type.</span></span> 

<span data-ttu-id="423b8-158">当前支持的汇总函数类型为 `Sum`、`Count`、`Average`、`Max` `Min`、`Product`、`CountNumbers`、`StandardDeviation`、`StandardDeviationP`、`Variance`、`VarianceP` 和 `Automatic`（默认）。</span><span class="sxs-lookup"><span data-stu-id="423b8-158">The currently supported aggregation function types are `Sum`, `Count`, `Average`, `Max`, `Min`, `Product`, `CountNumbers`, `StandardDeviation`, `StandardDeviationP`, `Variance`, `VarianceP`, and `Automatic` (the default).</span></span>

<span data-ttu-id="423b8-159">以下代码示例将数据汇总更改为平均值。</span><span class="sxs-lookup"><span data-stu-id="423b8-159">The following code samples changes the aggregation to be averages of the data.</span></span>

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

## <a name="change-calculations-with-a-showasrule"></a><span data-ttu-id="423b8-160">更改 ShowAsRule 计算</span><span class="sxs-lookup"><span data-stu-id="423b8-160">Change calculations with a ShowAsRule</span></span>

<span data-ttu-id="423b8-161">数据透视表默认情况下独立汇总其行和列层次结构的数据。</span><span class="sxs-lookup"><span data-stu-id="423b8-161">PivotTables, by default, aggregate the data of their row and column hierarchies independently.</span></span> <span data-ttu-id="423b8-162">`ShowAsRule` 数据层次结构更改为基于数据透视表中的其他项输出值。</span><span class="sxs-lookup"><span data-stu-id="423b8-162">A `ShowAsRule` changes the data hierarchy to output values based on other items in the PivotTable.</span></span>

<span data-ttu-id="423b8-163"> `ShowAsRule` 对象具有三个属性：</span><span class="sxs-lookup"><span data-stu-id="423b8-163">The `ShowAsRule` object has three properties:</span></span>
-   <span data-ttu-id="423b8-164">`calculation`：相对于数据层次结构的计算类型（默认是 `none`）。</span><span class="sxs-lookup"><span data-stu-id="423b8-164">`calculation`: The type of relative calculation to apply to the data hierarchy (the default is `none`).</span></span>
-   <span data-ttu-id="423b8-165">`baseField`：应用计算前层次结构中包含基准数据的字段。</span><span class="sxs-lookup"><span data-stu-id="423b8-165">`baseField`: The field within the hierarchy containing the base data before the calculation is applied.</span></span> <span data-ttu-id="423b8-166"> `PivotField` 通常与其父层次结构具有相同的名称。</span><span class="sxs-lookup"><span data-stu-id="423b8-166">The `PivotField` usually has the same name as its parent hierarchy.</span></span>
-   <span data-ttu-id="423b8-167">`baseItem`：根据计算类型的基本字段的值进行比较的单个项。</span><span class="sxs-lookup"><span data-stu-id="423b8-167">`baseItem`: The individual item compared against the values of the base fields based on the calculation type.</span></span> <span data-ttu-id="423b8-168">并非所有计算都需要此字段。</span><span class="sxs-lookup"><span data-stu-id="423b8-168">Not all calculations require this field.</span></span>

<span data-ttu-id="423b8-169">下面的示例将对 **Sum of Crates Sold at Farm** 数据层次结构列执行的计算设置为列总计的百分比。</span><span class="sxs-lookup"><span data-stu-id="423b8-169">The following example sets the calculation on the **Sum of Crates Sold at Farm** data hierarchy to be a percentage of the column total.</span></span> <span data-ttu-id="423b8-170">我们仍希望将粒度级别扩展至水果类型，因此我们将使用 **Type** 行层次结构及其基础字段。</span><span class="sxs-lookup"><span data-stu-id="423b8-170">We still want the granularity to extend to the fruit type level, so we’ll use the **Type** row hierarchy and its underlying field.</span></span> <span data-ttu-id="423b8-171">此示例还以 **Farm** 作为第一个行层次结构，因此农场的总项也显示每个农场负责产出的百分比。</span><span class="sxs-lookup"><span data-stu-id="423b8-171">The example also has **Farm** as the first row hierarchy, so the farm total entries display the percentage each farm is responsible for producing as well.</span></span>

![数据透视表显示每个农场的水果销售额相对于总销售额的百分比，以及单个农场各果品所占销售额的百分比。](../images/excel-pivots-showas-percentage.png)

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

<span data-ttu-id="423b8-173">上面的示例将计算设置为列中相对于单个行层次结构。</span><span class="sxs-lookup"><span data-stu-id="423b8-173">The previous example set the calculation to the column, relative to an individual row hierarchy.</span></span> <span data-ttu-id="423b8-174">当计算与单项相关时，使用 `baseItem` 属性。</span><span class="sxs-lookup"><span data-stu-id="423b8-174">When the calculation relates to an individual item, use the `baseItem` property.</span></span> 

<span data-ttu-id="423b8-175">下面的示例演示 `differenceFrom` 计算。</span><span class="sxs-lookup"><span data-stu-id="423b8-175">The following example shows the request.</span></span> <span data-ttu-id="423b8-176">它会显示农场销售箱数层次结构条目相对于 "A Farms" 的差异。</span><span class="sxs-lookup"><span data-stu-id="423b8-176">It displays the difference of the farm crate sales data hierarchy entries relative to those of “A Farms”.</span></span> <span data-ttu-id="423b8-177"> `baseField` 是 *\*Farm*\*，以便我们可以看到其他农场之间的差异，以及每种类似果品（*\*Type** 也是在此示例中的行层次结构）之间的差异。</span><span class="sxs-lookup"><span data-stu-id="423b8-177">The `baseField` is **Farm**, so we see the differences between the other farms, as well as breakdowns for each type of like fruit (**Type** is also a row hierarchy in this example).</span></span>

![显示 "A Farms" 和其他农场水果销售之间差异的数据透视表。](../images/excel-pivots-showas-differencefrom.png)

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

## <a name="pivottable-layouts"></a><span data-ttu-id="423b8-181">数据透视表布局</span><span class="sxs-lookup"><span data-stu-id="423b8-181">PivotTable layouts</span></span>

<span data-ttu-id="423b8-182">数据透视表布局定义层次结构及其数据的位置。</span><span class="sxs-lookup"><span data-stu-id="423b8-182">A PivotTable layout defines the placement of hierarchies and their data.</span></span> <span data-ttu-id="423b8-183">访问布局来确定存储数据区域的范围。</span><span class="sxs-lookup"><span data-stu-id="423b8-183">You access the layout to determine the ranges where data is stored.</span></span> 

<span data-ttu-id="423b8-184">下图显示了哪个布局函数调用对应哪个数据透视表范围。</span><span class="sxs-lookup"><span data-stu-id="423b8-184">The following diagram shows which layout function calls correspond to which ranges of the PivotTable.</span></span>

![此图显示数据透视表的哪些部分是由布局的获取范围函数返回的。](../images/excel-pivots-layout-breakdown.png)

<span data-ttu-id="423b8-186">下面的代码演示了如何通过布局获取数据透视表数据的最后一行。</span><span class="sxs-lookup"><span data-stu-id="423b8-186">The following code demonstrates how to get the last row of the PivotTable data by going through the layout.</span></span> <span data-ttu-id="423b8-187">然后对这些值求和获得总计。</span><span class="sxs-lookup"><span data-stu-id="423b8-187">Those values are then summed together for a grand total.</span></span>

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

<span data-ttu-id="423b8-188">数据透视表有三种布局样式：压缩、大纲和表格。</span><span class="sxs-lookup"><span data-stu-id="423b8-188">PivotTables have three layout styles: Compact, Outline, and Tabular.</span></span> <span data-ttu-id="423b8-189">在上面的示例中我们看到过压缩样式。</span><span class="sxs-lookup"><span data-stu-id="423b8-189">We’ve seen the compact style in the previous examples.</span></span> 

<span data-ttu-id="423b8-190">下面的示例分别使用大纲和表格样式。</span><span class="sxs-lookup"><span data-stu-id="423b8-190">The following examples use the outline and tabular styles, respectively.</span></span> <span data-ttu-id="423b8-191">代码示例显示如何在不同的布局之间转换。</span><span class="sxs-lookup"><span data-stu-id="423b8-191">The code sample shows how to cycle between the different layouts.</span></span>

### <a name="outline-layout"></a><span data-ttu-id="423b8-192">大纲版式</span><span class="sxs-lookup"><span data-stu-id="423b8-192">Outline layout</span></span>

![使用大纲版式的数据透视表。](../images/excel-pivots-outline-layout.png)

### <a name="tabular-layout"></a><span data-ttu-id="423b8-194">表格版式</span><span class="sxs-lookup"><span data-stu-id="423b8-194">Tabular layout</span></span>

![使用表格版式的数据透视表。](../images/excel-pivots-tabular-layout.png)

## <a name="change-hierarchy-names"></a><span data-ttu-id="423b8-196">更改层次结构名称</span><span class="sxs-lookup"><span data-stu-id="423b8-196">Change hierarchy names</span></span>

<span data-ttu-id="423b8-197">层次结构字段为可编辑。</span><span class="sxs-lookup"><span data-stu-id="423b8-197">Hierarchy fields are editable.</span></span> <span data-ttu-id="423b8-198">下面的代码演示如何交换两个数据层次结构的显示名称。</span><span class="sxs-lookup"><span data-stu-id="423b8-198">The following code demonstrates how to change the displayed names of two data hierarchies.</span></span>

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

## <a name="delete-a-pivottable"></a><span data-ttu-id="423b8-199">删除数据透视表</span><span class="sxs-lookup"><span data-stu-id="423b8-199">Delete a PivotTable</span></span>

<span data-ttu-id="423b8-200">使用数据透视表的名称删除数据透视表。</span><span class="sxs-lookup"><span data-stu-id="423b8-200">PivotTables are deleted by using their name.</span></span>

```typescript
await Excel.run(async (context) => {
    context.workbook.worksheets.getItem("Pivot").pivotTables.getItem("Farm Sales").delete();

    await context.sync();
});
```

## <a name="see-also"></a><span data-ttu-id="423b8-201">另请参阅</span><span class="sxs-lookup"><span data-stu-id="423b8-201">See also</span></span>

- [<span data-ttu-id="423b8-202">Excel JavaScript API 核心概念</span><span class="sxs-lookup"><span data-stu-id="423b8-202">Excel JavaScript API core concepts</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="423b8-203">Excel JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="423b8-203">Excel JavaScript API reference</span></span>](https://docs.microsoft.com/javascript/api/excel)
