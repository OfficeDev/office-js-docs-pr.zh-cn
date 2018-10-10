---
title: 使用 Excel JavaScript API 处理数据透视表
description: 使用 Excel JavaScript API 创建数据透视表并与其组件交互。
ms.date: 09/21/2018
ms.openlocfilehash: 00dd982d4ba4de0db34277cd546b572d4394e258
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459278"
---
# <a name="work-with-pivottables-using-the-excel-javascript-api"></a><span data-ttu-id="8e955-103">使用 Excel JavaScript API 处理数据透视表</span><span class="sxs-lookup"><span data-stu-id="8e955-103">Work with ranges using the Excel JavaScript API</span></span>

<span data-ttu-id="8e955-p101">数据透视表可简化大型数据集。它们允许快速操作分组的数据。Excel JavaScript API 允许加载项创建数据透视表并与其组件进行交互。</span><span class="sxs-lookup"><span data-stu-id="8e955-p101">PivotTables streamline larger data sets. They allow the quick manipulation of grouped data. The Excel JavaScript API lets your add-in create PivotTables and interact with their components.</span></span> 

<span data-ttu-id="8e955-p102">如果不熟悉数据透视表的功能，尝试以最终用户的身份了解它们的功能。请参阅 [创建数据透视表以分析工作表数据](https://support.office.com/en-us/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) ，了解这些工具的入门指导。</span><span class="sxs-lookup"><span data-stu-id="8e955-p102">If you are unfamiliar with the functionality of PivotTables, consider exploring them as an end user. See [Create a PivotTable to analyze worksheet data](https://support.office.com/en-us/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) for a good primer on these tools.</span></span> 

<span data-ttu-id="8e955-p103">本文提供了常见方案的代码示例。若要进一步了解数据透视表 API，请参阅 [**数据透视表**](https://docs.microsoft.com/javascript/api/excel/excel.pivottable) 和 [**PivotTableCollection**](https://docs.microsoft.com/javascript/api/excel/excel.pivottable) 。</span><span class="sxs-lookup"><span data-stu-id="8e955-p103">This article provides code samples for common scenarios. To further your understanding of the PivotTable API, see [**PivotTable**](https://docs.microsoft.com/javascript/api/excel/excel.pivottable) and [**PivotTableCollection**](https://docs.microsoft.com/javascript/api/excel/excel.pivottable).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="8e955-111">目前不支持使用 OLAP 创建的数据透视表。</span><span class="sxs-lookup"><span data-stu-id="8e955-111">PivotTables created with OLAP are not currently supported.</span></span>

## <a name="hierarchies"></a><span data-ttu-id="8e955-112">层次结构</span><span class="sxs-lookup"><span data-stu-id="8e955-112">Hierarchies</span></span>

<span data-ttu-id="8e955-p104">数据透视表根据四个层次结构类别进行组织：行、列、数据和筛选器。以下描述不同农场的水果销售的数据将被用于本文全文中。</span><span class="sxs-lookup"><span data-stu-id="8e955-p104">PivotTables are organized based on four hierarchy categories: row, column, data, and filter. The following data describing fruit sales from various farms will be used throughout this article.</span></span>

![来自不同农场的不同类型水果销售的集合。](../images/excel-pivots-raw-data.png)

<span data-ttu-id="8e955-p105">此数据具有五个层次结构： **Farms** 、 **Type** 、 **Classification** 、 **Crates Sold at Farm** 和 **Crates Sold Wholesale** 。每个层次结构可以仅存在于四种类型之一。如果 **Type** 被添加到列层次结构，然后又被添加到行层次结构，则它仅在后者保留。</span><span class="sxs-lookup"><span data-stu-id="8e955-p105">This data has five hierarchies: **Farms**, **Type**, **Classification**, **Crates Sold at Farm**, and **Crates Sold Wholesale**. Each hierarchy can only exist in one of the four categories. If **Type** is added to column hierarchies and then added to row hierarchies, it only remains in the latter.</span></span>

<span data-ttu-id="8e955-p106">行和列的层次结构定义了数据如何被分组。例如， **Farms** 的行层次结构将同一个农场的所有数据集组合在一起。行和列的层次结构之间的选择定义了数据透视表的方向。</span><span class="sxs-lookup"><span data-stu-id="8e955-p106">Row and column hierarchies define how data will be grouped. For example, a row hierarchy of **Farms** will group together all the data sets from the same farm. The choice between row and column hierarchy defines the orientation of the PivotTable.</span></span>

<span data-ttu-id="8e955-p107">数据层次结构是基于行和列层次结构进行聚合的值。具有 **Farms** 的行层次结构和 **Crates Sold Wholesale** 的数据层次结构的数据透视表显示了每个农场的所有不同水果的总和（默认）。</span><span class="sxs-lookup"><span data-stu-id="8e955-p107">Data hierarchies are the values to be aggregated based on the row and column hierarchies. A PivotTable with a row hierarchy of **Farms** and a data hierarchy of **Crates Sold Wholesale** shows the sum total (by default) of all the different fruits for each farm.</span></span>

<span data-ttu-id="8e955-p108">筛选器层次结构根据该筛选类型中的值包含或排除枢纽的数据。选择 **Organic** 类型的 **Classification** 筛选器层次结构仅显示有机水果的数据。</span><span class="sxs-lookup"><span data-stu-id="8e955-p108">Filter hierarchies include or exclude data from the pivot based on values within that filtered type. A filter hierarchy of **Classification** with the type **Organic** selected only shows data for organic fruit.</span></span>

<span data-ttu-id="8e955-p109">这里又是农场数据，以及数据透视表。数据透视表使用 **Farm** 和 **Type** 作为行层次结构， **Crates Sold at Farm** 和 **Crates Sold Wholesale** 作为数据层次结构 （带总和的默认汇总函数），以及 **Classification**  作为筛选器层次结构（选中 **Organic**）。</span><span class="sxs-lookup"><span data-stu-id="8e955-p109">Here is the farm data again, alongside a PivotTable. The PivotTable is using **Farm** and **Type** as the row hierarchies, **Crates Sold at Farm** and **Crates Sold Wholesale** as the data hierarchies (with the default aggregation function of sum), and **Classification** as a filter hierarchy (with **Organic** selected).</span></span> 

![具有行、数据和筛选器层次结构的数据透视表旁边的水果销售数据的选定内容。](../images/excel-pivot-table-and-data.png)

<span data-ttu-id="8e955-p110">无法通过 JavaScript API 或 Excel 用户界面生成此数据透视表。这两个选项允许通过加载项进行进一步的操作。</span><span class="sxs-lookup"><span data-stu-id="8e955-p110">This PivotTable could be generated through the JavaScript API or through the Excel UI. Both options allow for further manipulation through add-ins.</span></span>

## <a name="create-a-pivottable"></a><span data-ttu-id="8e955-131">创建数据透视表</span><span class="sxs-lookup"><span data-stu-id="8e955-131">Create a PivotTable and PivotChart</span></span>

<span data-ttu-id="8e955-p111">数据透视表的需要名称、 源和目标。源可以是区域地址或表名称（作为 `Range` 、 `string`、或 `Table` 类型进行传递)。目标是某一区域地址 (作为 `Range` 或 `string` 给定)。下面的示例显示各种数据透视表创建技术。</span><span class="sxs-lookup"><span data-stu-id="8e955-p111">PivotTables need a name, source, and destination. The source can be a range address or table name (passed as a `Range`, `string`, or `Table` type). The destination is a range address (given as either a `Range` or `string`). The following samples show various PivotTable creation techniques.</span></span>

### <a name="create-a-pivottable-with-range-addresses"></a><span data-ttu-id="8e955-136">创建带范围地址的数据透视表</span><span class="sxs-lookup"><span data-stu-id="8e955-136">Create a PivotTable with range addresses</span></span>

```typescript
await Excel.run(async (context) => {
    // creating a PivotTable named "Farm Sales" on the current worksheet at cell A22 with data from the range A1:E21
    context.workbook.worksheets.getActiveWorksheet().pivotTables.add("Farm Sales", "A1:E21", "A22");

    await context.sync();
});
```

### <a name="create-a-pivottable-with-range-objects"></a><span data-ttu-id="8e955-137">创建带范围对象的数据透视表</span><span class="sxs-lookup"><span data-stu-id="8e955-137">Create a PivotTable with Range objects</span></span>

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

### <a name="create-a-pivottable-at-the-workbook-level"></a><span data-ttu-id="8e955-138">在工作簿级别创建数据透视表</span><span class="sxs-lookup"><span data-stu-id="8e955-138">Create a PivotTable at the workbook level</span></span>

```typescript
await Excel.run(async (context) => {
    // creating a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data is from the worksheet "DataWorksheet" across the range A1:E21
    context.workbook.pivotTables.add("Farm Sales", "DataWorksheet!A1:E21", "PivotWorksheet!A2");

    await context.sync();
});
```

## <a name="use-an-existing-pivottable"></a><span data-ttu-id="8e955-139">使用现有数据透视表</span><span class="sxs-lookup"><span data-stu-id="8e955-139">Use an existing PivotTable</span></span>

<span data-ttu-id="8e955-140">手动创建的数据透视表亦可通过工作簿或者单个工作表的数据透视表集合进行访问。</span><span class="sxs-lookup"><span data-stu-id="8e955-140">Manually created PivotTables are also accessible through the PivotTable collection of the workbook or of individual worksheets.</span></span> 

<span data-ttu-id="8e955-p112">以下代码获取工作簿中第一个数据透视表。然后给出了表的名称以方便用于将来的参考。</span><span class="sxs-lookup"><span data-stu-id="8e955-p112">The following code gets the first PivotTable in the workbook. It then gives the table a name for easy future reference.</span></span>

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.pivotTables.getItem("My Pivot");
    await context.sync();
});
```

## <a name="add-rows-and-columns-to-a-pivottable"></a><span data-ttu-id="8e955-143">向数据透视表添加行和列</span><span class="sxs-lookup"><span data-stu-id="8e955-143">Add rows and columns to a PivotTable</span></span>

<span data-ttu-id="8e955-144">行和列按字段值相关的方式透视数据。</span><span class="sxs-lookup"><span data-stu-id="8e955-144">Rows and columns pivot the data around those fields’ values.</span></span>

<span data-ttu-id="8e955-p113">添加 **Farm** 列使所有销售数据以每个农场为中心运行。添加 **Type** 和 **Classification** 行进一步根据哪些水果已售出以及它是否是有机的，以此对数据进行分解。</span><span class="sxs-lookup"><span data-stu-id="8e955-p113">Adding the **Farm** column pivots all the sales around each farm. Adding the **Type** and **Classification** rows further breaks down the data based on what fruit was sold and whether it was organic or not.</span></span>

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

<span data-ttu-id="8e955-148">还可拥有仅带行或列的数据透视表。</span><span class="sxs-lookup"><span data-stu-id="8e955-148">You can also have a PivotTable with only rows or columns.</span></span>

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));
    
    await context.sync();
});
```

## <a name="add-data-hierarchies-to-the-pivottable"></a><span data-ttu-id="8e955-149">向数据透视表添加数据层次结构</span><span class="sxs-lookup"><span data-stu-id="8e955-149">Add data hierarchies to the PivotTable</span></span>

<span data-ttu-id="8e955-p114">数据层次结构根据行和列在数据透视表中填充了信息。添加 **Crates Sold at Farm** 和 **Crates Sold Wholesale** 的数据层次结构为每行和列提供了那些图表的总和。</span><span class="sxs-lookup"><span data-stu-id="8e955-p114">Data hierarchies fill the PivotTable with information to combine based on the rows and columns. Adding the data hierarchies of **Crates Sold at Farm** and **Crates Sold Wholesale** gives sums of those figures for each row and column.</span></span> 

<span data-ttu-id="8e955-152">在示例中， **Farm** 和 **Type** 是行，而售出箱数属于数据。</span><span class="sxs-lookup"><span data-stu-id="8e955-152">In the example, both **Farm** and **Type** are rows, with the crate sales as the data.</span></span> 

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

## <a name="change-aggregation-function"></a><span data-ttu-id="8e955-154">更改汇总函数</span><span class="sxs-lookup"><span data-stu-id="8e955-154">Change aggregation function</span></span>

<span data-ttu-id="8e955-p115">数据层次结构可以聚合其值。对于数字的数据集，这是默认情况下的总和。 `summarizeBy` 属性根据 `AggregrationFunction` 类型定义了此行为。</span><span class="sxs-lookup"><span data-stu-id="8e955-p115">Data hierarchies have their values aggregated. For datasets of numbers, this is a sum by default. The `summarizeBy` property defines this behavior based on an `AggregrationFunction` type.</span></span> 

<span data-ttu-id="8e955-158">当前支持的汇总函数类型为 `Sum` 、 `Count` 、 `Average` 、 `Max` 、 `Min` 、 `Product` 、 `CountNumbers` 、 `StandardDeviation` 、 `StandardDeviationP` 、 `Variance` 、 `VarianceP` 、 和 `Automatic` （默认）。</span><span class="sxs-lookup"><span data-stu-id="8e955-158">The currently supported aggregation function types are `Sum`, `Count`, `Average`, `Max`, `Min`, `Product`, `CountNumbers`, `StandardDeviation`, `StandardDeviationP`, `Variance`, `VarianceP`, and `Automatic` (the default).</span></span>

<span data-ttu-id="8e955-159">以下代码示例将数据汇总更改为平均值。</span><span class="sxs-lookup"><span data-stu-id="8e955-159">The following code samples changes the aggregation to be averages of the data.</span></span>

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

## <a name="change-calculations-with-a-showasrule"></a><span data-ttu-id="8e955-160">更改 ShowAsRule 计算</span><span class="sxs-lookup"><span data-stu-id="8e955-160">Change calculations with a ShowAsRule</span></span>

<span data-ttu-id="8e955-p116">数据透视表在默认情况下独立聚合了它们行和列的层次结构的数据。一个 `ShowAsRule` 根据数据透视表中的其他项改变了数据层次结构的输出值。</span><span class="sxs-lookup"><span data-stu-id="8e955-p116">PivotTables, by default, aggregate the data of their row and column hierarchies independently. A `ShowAsRule` changes the data hierarchy to output values based on other items in the PivotTable.</span></span>

<span data-ttu-id="8e955-163">此 `ShowAsRule` 对象具有三种属性：</span><span class="sxs-lookup"><span data-stu-id="8e955-163">The `ShowAsRule` object has three properties:</span></span>
-   <span data-ttu-id="8e955-164">`calculation`：相对于数据层次结构的计算类型（默认是 `none`）。</span><span class="sxs-lookup"><span data-stu-id="8e955-164">`calculation`: The type of relative calculation to apply to the data hierarchy (the default is `none`).</span></span>
-   <span data-ttu-id="8e955-p117">`baseField`: 计算之前包含了基准数据的层次结构内的字段会被应用。 `PivotField` 通常具有其父层次结构相同的名称。</span><span class="sxs-lookup"><span data-stu-id="8e955-p117">`baseField`: The field within the hierarchy containing the base data before the calculation is applied. The `PivotField` usually has the same name as its parent hierarchy.</span></span>
-   <span data-ttu-id="8e955-p118">`baseItem`: 根据计算类型与基本字段的值进行比较的单项。并非所有计算都需要此字段。</span><span class="sxs-lookup"><span data-stu-id="8e955-p118">`baseItem`: The individual item compared against the values of the base fields based on the calculation type. Not all calculations require this field.</span></span>

<span data-ttu-id="8e955-p119">下面的示例将 **Sum of Crates Sold at Farm** 数据层次结构的计算设置为列总计的百分比。我们仍希望此粒度扩展到水果类型的级别，因此我们将使用 **Type** 行层次结构及其基础字段。此示例还将 **Farm** 作为第一个行层次结构，因此农场的总项还显示了每个农场负责生产的百分比。</span><span class="sxs-lookup"><span data-stu-id="8e955-p119">The following example sets the calculation on the **Sum of Crates Sold at Farm** data hierarchy to be a percentage of the column total. We still want the granularity to extend to the fruit type level, so we’ll use the **Type** row hierarchy and its underlying field. The example also has **Farm** as the first row hierarchy, so the farm total entries display the percentage each farm is responsible for producing as well.</span></span>

![数据透视表显示每个农场的水果销售额相对于总销售额的百分比，以及单个农场各水果类型所占销售额的百分比。](../images/excel-pivots-showas-percentage.png)

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

<span data-ttu-id="8e955-p120">上面的示例将计算设置为列中，与单个行层次结构相对。当计算与单项相关时，使用 `baseItem` 属性。</span><span class="sxs-lookup"><span data-stu-id="8e955-p120">The previous example set the calculation to the column, relative to an individual row hierarchy. When the calculation relates to an individual item, use the `baseItem` property.</span></span> 

<span data-ttu-id="8e955-p121">下面的示例演示了 `differenceFrom` 计算。它会显示相对于“A Farms”的农场售出箱数的数据层次结构条目的差异。 `baseField` 是 **Farm** ，因此我们可以看到其他农场之间的不同，以及每种类型的类似水果（**Type**  也是在此示例中的行层次结构） 的明细。</span><span class="sxs-lookup"><span data-stu-id="8e955-p121">The following example shows the `differenceFrom` calculation. It displays the difference of the farm crate sales data hierarchy entries relative to those of “A Farms”. The `baseField` is **Farm**, so we see the differences between the other farms, as well as breakdowns for each type of like fruit (**Type** is also a row hierarchy in this example).</span></span>

![数据透视表将显示"A Farms"和其他农场之间的水果销售差异。此时将显示农场的总水果销售和各种水果类型的销售这两个差异。如果"A Farms"并不销售某种类型的水果，将显示"# n/A"。](../images/excel-pivots-showas-differencefrom.png)

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

## <a name="pivottable-layouts"></a><span data-ttu-id="8e955-181">数据透视表布局</span><span class="sxs-lookup"><span data-stu-id="8e955-181">PivotTable layouts</span></span>

<span data-ttu-id="8e955-p123">数据透视表布局定义了层次结构和其数据的位置。您可访问此布局以确定存储数据的区域。</span><span class="sxs-lookup"><span data-stu-id="8e955-p123">A PivotTable layout defines the placement of hierarchies and their data. You access the layout to determine the ranges where data is stored.</span></span> 

<span data-ttu-id="8e955-184">下图显示了哪个布局函数调用对应哪个数据透视表范围。</span><span class="sxs-lookup"><span data-stu-id="8e955-184">The following diagram shows which layout function calls correspond to which ranges of the PivotTable.</span></span>

![此图显示数据透视表的哪些部分是由布局的获取范围函数返回的。](../images/excel-pivots-layout-breakdown.png)

<span data-ttu-id="8e955-p124">下面的代码演示如何通过仔细检查布局来获取数据透视表数据的最后一行。这些值随后相加以得出总和。</span><span class="sxs-lookup"><span data-stu-id="8e955-p124">The following code demonstrates how to get the last row of the PivotTable data by going through the layout. Those values are then summed together for a grand total.</span></span>

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

<span data-ttu-id="8e955-p125">数据透视表有三种布局样式：压缩、大纲和表格。我们已在上面的示例看到了压缩样式。</span><span class="sxs-lookup"><span data-stu-id="8e955-p125">PivotTables have three layout styles: Compact, Outline, and Tabular. We’ve seen the compact style in the previous examples.</span></span> 

<span data-ttu-id="8e955-p126">下面的示例分别使用大纲和表格样式。代码示例演示如何在不同的布局之间循环。</span><span class="sxs-lookup"><span data-stu-id="8e955-p126">The following examples use the outline and tabular styles, respectively. The code sample shows how to cycle between the different layouts.</span></span>

### <a name="outline-layout"></a><span data-ttu-id="8e955-192">大纲版式</span><span class="sxs-lookup"><span data-stu-id="8e955-192">Outline layout</span></span>

![使用大纲版式的数据透视表。](../images/excel-pivots-outline-layout.png)

### <a name="tabular-layout"></a><span data-ttu-id="8e955-194">表格版式</span><span class="sxs-lookup"><span data-stu-id="8e955-194">Tabular layout</span></span>

![使用表格版式的数据透视表。](../images/excel-pivots-tabular-layout.png)

## <a name="change-hierarchy-names"></a><span data-ttu-id="8e955-196">更改层次结构名称</span><span class="sxs-lookup"><span data-stu-id="8e955-196">Change hierarchy names</span></span>

<span data-ttu-id="8e955-p127">层次结构字段可以编辑。下面的代码演示了如何更改两个数据层次结构的显示名称。</span><span class="sxs-lookup"><span data-stu-id="8e955-p127">Hierarchy fields are editable. The following code demonstrates how to change the displayed names of two data hierarchies.</span></span>

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

## <a name="delete-a-pivottable"></a><span data-ttu-id="8e955-199">删除数据透视表</span><span class="sxs-lookup"><span data-stu-id="8e955-199">Delete a PivotTable</span></span>

<span data-ttu-id="8e955-200">使用数据透视表的名称删除数据透视表。</span><span class="sxs-lookup"><span data-stu-id="8e955-200">PivotTables are deleted by using their name.</span></span>

```typescript
await Excel.run(async (context) => {
    context.workbook.worksheets.getItem("Pivot").pivotTables.getItem("Farm Sales").delete();

    await context.sync();
});
```

## <a name="see-also"></a><span data-ttu-id="8e955-201">另请参阅</span><span class="sxs-lookup"><span data-stu-id="8e955-201">See also</span></span>

- [<span data-ttu-id="8e955-202">使用 Excel JavaScript API 的基本编程概念</span><span class="sxs-lookup"><span data-stu-id="8e955-202">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="8e955-203">Excel JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="8e955-203">Excel JavaScript API reference</span></span>](https://docs.microsoft.com/javascript/api/excel)
