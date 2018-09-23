---
title: 使用 Excel JavaScript API 处理数据透视表
description: 使用 Excel JavaScript API 创建数据透视表并与其组件交互。
ms.date: 09/21/2018
ms.openlocfilehash: b8704389ced3686858f488b2a50f80c22b1b8bd6
ms.sourcegitcommit: e7e4d08569a01c69168bb005188e9a1e628304b9
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/22/2018
ms.locfileid: "24967667"
---
# <a name="work-with-pivottables-using-the-excel-javascript-api"></a><span data-ttu-id="18c27-103">使用 Excel JavaScript API 处理数据透视表</span><span class="sxs-lookup"><span data-stu-id="18c27-103">Work with Events using the Excel JavaScript API</span></span>

<span data-ttu-id="18c27-104">数据透视表简化更大的数据集。</span><span class="sxs-lookup"><span data-stu-id="18c27-104">PivotTables streamline larger data sets.</span></span> <span data-ttu-id="18c27-105">它们允许分组数据的快速操作。</span><span class="sxs-lookup"><span data-stu-id="18c27-105">They allow the quick manipulation of grouped data.</span></span> <span data-ttu-id="18c27-106">Excel JavaScript API 允许加载项创建数据透视表并与其组件交互。</span><span class="sxs-lookup"><span data-stu-id="18c27-106">The Excel JavaScript API lets your add-in create PivotTables and interact with their components.</span></span> 

<span data-ttu-id="18c27-107">如果不熟悉数据透视表的功能，尝试以最终用户的身份了解它们的功能。</span><span class="sxs-lookup"><span data-stu-id="18c27-107">If you are unfamiliar with the functionality of PivotTables, consider exploring them as an end-user.</span></span> <span data-ttu-id="18c27-108">请参阅[创建数据透视表分析工作表数据](https://support.office.com/en-us/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables)，了解这些工具的入门指导。</span><span class="sxs-lookup"><span data-stu-id="18c27-108">See [Create a PivotTable to analyze worksheet data](https://support.office.com/en-us/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) for a good primer on these tools.</span></span> 

<span data-ttu-id="18c27-109">本文提供了常见方案的代码示例。</span><span class="sxs-lookup"><span data-stu-id="18c27-109">This article provides code samples for common scenarios.</span></span> <span data-ttu-id="18c27-110">参阅[**数据透视表**](https://docs.microsoft.com/javascript/api/excel/excel.pivottable)和[**PivotTableCollection**](https://docs.microsoft.com/javascript/api/excel/excel.pivottable)，加深对数据透视表 API 的理解。</span><span class="sxs-lookup"><span data-stu-id="18c27-110">To further your understanding of the PivotTable API, see [**PivotTable**](https://docs.microsoft.com/javascript/api/excel/excel.pivottable) and [**PivotTableCollection**](https://docs.microsoft.com/javascript/api/excel/excel.pivottable).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="18c27-111">目前不支持使用 OLAP 创建的数据透视表。</span><span class="sxs-lookup"><span data-stu-id="18c27-111">PivotTables created with OLAP are not currently supported.</span></span>

## <a name="hierarchies"></a><span data-ttu-id="18c27-112">层次结构</span><span class="sxs-lookup"><span data-stu-id="18c27-112">Hierarchies</span></span>

<span data-ttu-id="18c27-113">数据透视表基于四种层次结构类别构成：行、列、数据和筛选器。</span><span class="sxs-lookup"><span data-stu-id="18c27-113">PivotTables are organized based on four hierarchy categories: row, column, data, and filter.</span></span> <span data-ttu-id="18c27-114">本文将通篇使用以下描述各农场水果销售情况的数据。</span><span class="sxs-lookup"><span data-stu-id="18c27-114">The following data describing fruit sales from various farms will be used throughout this article.</span></span>

![来自不同农场的不同类型水果销售的集合。](../images/excel-pivots-raw-data.png)

<span data-ttu-id="18c27-116">此数据具有五个层次结构：**农场**、**类型**、**分类**、 **农场销售箱数**，和**批发箱数**。</span><span class="sxs-lookup"><span data-stu-id="18c27-116">This data has five hierarchies: **Farms**, **Type**, **Classification**, **Crates Sold at Farm**, and **Crates Sold Wholesale**.</span></span> <span data-ttu-id="18c27-117">每个层次结构只能存在于四个类别中的一个类别。</span><span class="sxs-lookup"><span data-stu-id="18c27-117">Each hierarchy can only exist in one of the four categories.</span></span> <span data-ttu-id="18c27-118">如果**类型**添加到列层次结构，然后又添加到行层次结构，则其仅保留至后者。</span><span class="sxs-lookup"><span data-stu-id="18c27-118">If **Type** is added to column hierarchies and then added to row hierarchies, it only remains in the latter.</span></span>

<span data-ttu-id="18c27-119">行和列的层次结构定义如何分组数据。</span><span class="sxs-lookup"><span data-stu-id="18c27-119">Row and column hierarchies define how data will be grouped.</span></span> <span data-ttu-id="18c27-120">例如，**农场**的行层次结构会将来自同一农场的所有数据集归集在一起。</span><span class="sxs-lookup"><span data-stu-id="18c27-120">For example, a row hierarchy of **Farms** will group together all the data sets from the same farm.</span></span> <span data-ttu-id="18c27-121">选择行和列层次结构来定义数据透视表的方向。</span><span class="sxs-lookup"><span data-stu-id="18c27-121">The choice between row and column hierarchy defines the orientation of the PivotTable.</span></span>

<span data-ttu-id="18c27-122">数据层次结构是基于行和列层次结构进行聚合的值。</span><span class="sxs-lookup"><span data-stu-id="18c27-122">Data hierarchies are the values to be aggregated based on the row and column hierarchies.</span></span> <span data-ttu-id="18c27-123">具有**农场**的行层次结构和**批发箱数**的数据层次结构的数据透视表显示每个农场所有不同水果的总和（默认）。</span><span class="sxs-lookup"><span data-stu-id="18c27-123">A PivotTable with a row hierarchy of **Farms** and a data hierarchy of **Crates Sold Wholesale** shows the sum total (by default) of all the different fruits for each farm.</span></span>

<span data-ttu-id="18c27-124">筛选器层次结构基于已筛选类型中的值包含或排除来自透视的数据。</span><span class="sxs-lookup"><span data-stu-id="18c27-124">Filter hierarchies include or exclude data from the pivot based on values within that filtered type.</span></span> <span data-ttu-id="18c27-125">选择了**有机**类型的**分类**筛选器层次结构仅显示有机水果的数据。</span><span class="sxs-lookup"><span data-stu-id="18c27-125">A filter hierarchy of **Classification** with the type **Organic** selected only shows data for organic fruit.</span></span>

<span data-ttu-id="18c27-126">这还是农场数据，一旁是数据透视表。</span><span class="sxs-lookup"><span data-stu-id="18c27-126">Here is the farm data again, alongside a PivotTable.</span></span> <span data-ttu-id="18c27-127">数据透视表使用 **农场**和**类型**作为行层次结构，**农场销售箱数**和**批发箱数**作为数据层次结构 （默认总和的聚合函数），**分类**作为筛选器层次结构（选中**有机**）。</span><span class="sxs-lookup"><span data-stu-id="18c27-127">The PivotTable is using **Farm** and **Type** as the row hierarchies, **Crates Sold at Farm** and **Crates Sold Wholesale** as the data hierarchies (with the default aggregation function of sum), and **Classification** as a filter hierarchy (with **Organic** selected).</span></span> 

![具有行、数据和筛选器层次结构的数据透视表旁边果销售数据的选定内容。](../images/excel-pivot-table-and-data.png)

<span data-ttu-id="18c27-129">可通过 JavaScript API 或通过 Excel UI 生成这个数据透视表。</span><span class="sxs-lookup"><span data-stu-id="18c27-129">This PivotTable could be generated through the JavaScript API or through the Excel UI.</span></span> <span data-ttu-id="18c27-130">两个选项均可通过加载项实现进一步的操作。</span><span class="sxs-lookup"><span data-stu-id="18c27-130">Both options allow for further manipulation through add-ins.</span></span>

## <a name="create-a-pivottable"></a><span data-ttu-id="18c27-131">创建数据透视表</span><span class="sxs-lookup"><span data-stu-id="18c27-131">Create a PivotTable and PivotChart</span></span>

<span data-ttu-id="18c27-132">数据透视表需要有名称、源和目标。</span><span class="sxs-lookup"><span data-stu-id="18c27-132">PivotTables need a name, source, and destination.</span></span> <span data-ttu-id="18c27-133">源可以是范围地址或表名（作为`Range`、`string`或`Table`类型传递）。</span><span class="sxs-lookup"><span data-stu-id="18c27-133">The source can be a range address or table name (passed as a `Range`, `string`, or `Table` type).</span></span> <span data-ttu-id="18c27-134">目标是某一范围地址（作为`Range` 或 `string`给定）。</span><span class="sxs-lookup"><span data-stu-id="18c27-134">The destination is a range address (given as either a `Range` or `string`).</span></span> <span data-ttu-id="18c27-135">以下示例显示各种数据透视表的创建技术。</span><span class="sxs-lookup"><span data-stu-id="18c27-135">The following samples show various PivotTable creation techniques.</span></span>

### <a name="create-a-pivottable-with-range-addresses"></a><span data-ttu-id="18c27-136">创建带范围地址的数据透视表</span><span class="sxs-lookup"><span data-stu-id="18c27-136">Create a PivotTable with range addresses</span></span>

```typescript
await Excel.run(async (context) => {
    // creating a PivotTable named "Farm Sales" created on the current worksheet at cell A22 with data from the range A1:E21
    context.workbook.worksheets.getActiveWorksheet()
        .pivotTables.add("Farm Sales", "A1:E21", "A22");

    await context.sync();
});
```

### <a name="create-a-pivottable-with-range-objects"></a><span data-ttu-id="18c27-137">创建带范围对象的数据透视表</span><span class="sxs-lookup"><span data-stu-id="18c27-137">Create a PivotTable with Range objects</span></span>

```typescript
await Excel.run(async (context) => {    
    // creating a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data comes from the worksheet "DataWorksheet" across the range A1:E21
    const rangeToAnalyze = context.workbook.worksheets.getItem("DataWorksheet").getRange("A1:E21");
    const rangeToPlacePivot = context.workbook.worksheets.getItem("PivotWorksheet").getRange("A2");
    context.workbook.worksheets.getItem("PivotWorksheet").pivotTables.add("Farm Sales", rangeToAnalyze, rangeToPlacePivot);
    
    await context.sync();
});
```

### <a name="create-a-pivottable-at-the-workbook-level"></a><span data-ttu-id="18c27-138">在工作簿级别创建数据透视表</span><span class="sxs-lookup"><span data-stu-id="18c27-138">Create a PivotTable at the workbook level</span></span>

```typescript
await Excel.run(async (context) => {
    // creating a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data is from the worksheet "DataWorksheet" across the range A1:E21
    context.workbook.pivotTables.add("Farm Sales", "DataWorksheet!A1:E21", "PivotWorksheet!A2");

    await context.sync();
});
```

## <a name="use-an-existing-pivottable"></a><span data-ttu-id="18c27-139">使用现有数据透视表</span><span class="sxs-lookup"><span data-stu-id="18c27-139">Use an existing PivotTable</span></span>

<span data-ttu-id="18c27-140">手动创建的数据透视表亦可通过工作簿或者单个工作表的数据透视表集合进行访问。</span><span class="sxs-lookup"><span data-stu-id="18c27-140">Manually created PivotTables are also accessible through the PivotTable collection of the workbook or of individual worksheets.</span></span> 

<span data-ttu-id="18c27-141">以下代码获取工作簿中的第一个数据透视表。</span><span class="sxs-lookup"><span data-stu-id="18c27-141">The following code gets the first PivotTable in the workbook.</span></span> <span data-ttu-id="18c27-142">然后给出了表的名称，便于以后参考。</span><span class="sxs-lookup"><span data-stu-id="18c27-142">It then gives the table a name for easy future reference.</span></span>

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.pivotTables.getItem("My Pivot");
    await context.sync();
});
```

## <a name="add-rows-and-columns-to-a-pivottable"></a><span data-ttu-id="18c27-143">向数据透视表添加行和列</span><span class="sxs-lookup"><span data-stu-id="18c27-143">Add rows and columns to a PivotTable</span></span>

<span data-ttu-id="18c27-144">行和列透视与那些字段值相关的数据。</span><span class="sxs-lookup"><span data-stu-id="18c27-144">Rows and columns pivot the data around those fields’ values.</span></span>

<span data-ttu-id="18c27-145">添加 **农场**列透视每个农场的所有销售情况。</span><span class="sxs-lookup"><span data-stu-id="18c27-145">Adding the **Farm** column pivots all the sales around each farm.</span></span> <span data-ttu-id="18c27-146">添加**类型**和**分类**行，可基于销售的水果以及该水果是否为有机等条件而将数据作进一步的分解。</span><span class="sxs-lookup"><span data-stu-id="18c27-146">Adding the **Type** and **Classification** rows further breaks down the data based on what fruit was sold and whether it was organic or not.</span></span>

![带农场列和类型及分类行的数据透视表。](../images/excel-pivots-table-rows-and-columns.png)

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));
    
    pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem("Farm"));

    await context.sync();
});
```

<span data-ttu-id="18c27-148">还可拥有仅带行或列的数据透视表。</span><span class="sxs-lookup"><span data-stu-id="18c27-148">You can also have a PivotTable with only rows or columns.</span></span>

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));
    
    await context.sync();
});
```

## <a name="add-data-hierarchies-to-the-pivottable"></a><span data-ttu-id="18c27-149">向数据透视表添加数据层次结构</span><span class="sxs-lookup"><span data-stu-id="18c27-149">Add data hierarchies to the PivotTable</span></span>

<span data-ttu-id="18c27-150">基于行和列，数据层次结构结合信息填充数据透视表。</span><span class="sxs-lookup"><span data-stu-id="18c27-150">Data hierarchies fill the PivotTable with information to combine based on the rows and columns.</span></span> <span data-ttu-id="18c27-151">添加**农场销售箱数**和**批发箱数**的数据层次结构给出每行和每列数字的总和。</span><span class="sxs-lookup"><span data-stu-id="18c27-151">Adding the data hierarchies of **Crates Sold at Farm** and **Crates Sold Wholesale** gives sums of those figures for each row and column.</span></span> 

<span data-ttu-id="18c27-152">在示例中，**农场**和**类型**都是行，销售箱数作为数据。</span><span class="sxs-lookup"><span data-stu-id="18c27-152">In the example, both **Farm** and **Type** are rows, with the crate sales as the data.</span></span> 

![显示基于水果来源农场的不同水果总销售情况的数据透视表。](../images/excel-pivots-data-hierarchy.png)

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    // "Farm" and "Type" are the hierarchies on which the aggregation is based
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));

    // "Crates Sold at Farm" and "Crates Sold Wholesale" are the heirarchies that will have their data aggregated (summed in this case)
    pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Crates Sold at Farm"));
    pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Crates Sold Wholesale"));

    await context.sync();
});
```

## <a name="change-aggregation-function"></a><span data-ttu-id="18c27-154">更改聚合函数</span><span class="sxs-lookup"><span data-stu-id="18c27-154">Change aggregation function</span></span>

<span data-ttu-id="18c27-155">数据层次结构将其数值聚合。</span><span class="sxs-lookup"><span data-stu-id="18c27-155">Data hierarchies have their values aggregated.</span></span> <span data-ttu-id="18c27-156">对于数字的数据集，默认情况下，这是总和。</span><span class="sxs-lookup"><span data-stu-id="18c27-156">For datasets of numbers, this is a sum by default.</span></span> <span data-ttu-id="18c27-157">属性基于类型定义此行为 `AggregrationFunction`。`summarizeBy`</span><span class="sxs-lookup"><span data-stu-id="18c27-157">The `summarizeBy` property defines this behavior based on an `AggregrationFunction` type.</span></span> 

<span data-ttu-id="18c27-158">当前支持的聚合函数类型为`Sum`、`Count`、`Average`、`Max` `Min`、`Product`、`CountNumbers`、`StandardDeviation`、`StandardDeviationP`、`Variance`、`VarianceP`和`Automatic` （默认）。</span><span class="sxs-lookup"><span data-stu-id="18c27-158">The currently supported aggregation function types are `Sum`, `Count`, `Average`, `Max`, `Min`, `Product`, `CountNumbers`, `StandardDeviation`, `StandardDeviationP`, `Variance`, `VarianceP`, and `Automatic` (the default).</span></span>

<span data-ttu-id="18c27-159">以下代码示例更改了数据平均值的聚合。</span><span class="sxs-lookup"><span data-stu-id="18c27-159">The following code samples changes the aggregation to be averages of the data.</span></span>

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

## <a name="pivottable-layouts"></a><span data-ttu-id="18c27-160">数据透视表布局</span><span class="sxs-lookup"><span data-stu-id="18c27-160">PivotTable layouts</span></span>

<span data-ttu-id="18c27-161">数据透视表布局定义层次结构及其数据的位置。</span><span class="sxs-lookup"><span data-stu-id="18c27-161">A PivotTable layout defines the placement of hierarchies and their data.</span></span> <span data-ttu-id="18c27-162">访问布局来确定存储数据区域的范围。</span><span class="sxs-lookup"><span data-stu-id="18c27-162">You access the layout to determine the ranges where data is stored.</span></span> 

<span data-ttu-id="18c27-163">下图显示了哪个布局函数调用对应哪个数据透视表范围。</span><span class="sxs-lookup"><span data-stu-id="18c27-163">The following diagram shows which layout function calls correspond to which ranges of the PivotTable.</span></span>

![此图显示数据透视表的哪些部分是由布局的获取范围函数返回的。](../images/excel-pivots-layout-breakdown.png)

<span data-ttu-id="18c27-165">下面的代码演示了如何通过布局获取数据透视表数据的最后一行。</span><span class="sxs-lookup"><span data-stu-id="18c27-165">The following code demonstrates how to get the last row of the PivotTable data by going through the layout.</span></span> <span data-ttu-id="18c27-166">然后对这些值求和获得总计。</span><span class="sxs-lookup"><span data-stu-id="18c27-166">Those values are then summed together for a grand total.</span></span>


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

<span data-ttu-id="18c27-167">数据透视表有三种布局样式：压缩、大纲和表格。</span><span class="sxs-lookup"><span data-stu-id="18c27-167">PivotTables have three layout styles: Compact, Outline, and Tabular.</span></span> <span data-ttu-id="18c27-168">在上面的示例中我们看到过压缩样式。</span><span class="sxs-lookup"><span data-stu-id="18c27-168">We’ve seen the compact style in the previous examples.</span></span> 

<span data-ttu-id="18c27-169">下面的示例分别使用大纲和表格样式。</span><span class="sxs-lookup"><span data-stu-id="18c27-169">The following examples use the outline and tabular styles, respectively.</span></span> <span data-ttu-id="18c27-170">代码示例显示如何在不同的布局之间转换。</span><span class="sxs-lookup"><span data-stu-id="18c27-170">The code sample shows how to cycle between the different layouts.</span></span>

### <a name="outline-layout"></a><span data-ttu-id="18c27-171">大纲版式</span><span class="sxs-lookup"><span data-stu-id="18c27-171">Outline layout</span></span>

![使用大纲版式的数据透视表。](../images/excel-pivots-outline-layout.png)

### <a name="tabular-layout"></a><span data-ttu-id="18c27-173">表格版式</span><span class="sxs-lookup"><span data-stu-id="18c27-173">Tabular layout</span></span>

![使用表格版式的数据透视表。](../images/excel-pivots-tabular-layout.png)

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.layout.load("layoutType");
    await context.sync();
    
    // cycling through layout styles
    if (pivotTable.layout.layoutType === "Compact") {
        pivotTable.layout.layoutType = "Outline";
    } else if (pivotTable.layout.layoutType === "Outline") {
        pivotTable.layout.layoutType = "Tabular";
    } else {
        pivotTable.layout.layoutType = "Compact";
    }
    
    await context.sync();
});
```

## <a name="change-hierarchy-names"></a><span data-ttu-id="18c27-175">更改层次结构名称</span><span class="sxs-lookup"><span data-stu-id="18c27-175">Change hierarchy names</span></span>

<span data-ttu-id="18c27-176">层次结构字段为可编辑。</span><span class="sxs-lookup"><span data-stu-id="18c27-176">Hierarchy fields are editable.</span></span> <span data-ttu-id="18c27-177">下面的代码演示如何交换两个数据层次结构的显示名称。</span><span class="sxs-lookup"><span data-stu-id="18c27-177">The following code demonstrates how to change the displayed names of two data hierarchies.</span></span>

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

## <a name="delete-a-pivottable"></a><span data-ttu-id="18c27-178">删除数据透视表</span><span class="sxs-lookup"><span data-stu-id="18c27-178">Delete a PivotTable</span></span>

<span data-ttu-id="18c27-179">使用数据透视表的名称删除数据透视表。</span><span class="sxs-lookup"><span data-stu-id="18c27-179">PivotTables are deleted by using their name.</span></span>

```typescript
await Excel.run(async (context) => {
    context.workbook.worksheets.getItem("Pivot").pivotTables.getItem("Farm Sales").delete();

    await context.sync();
});
```

## <a name="see-also"></a><span data-ttu-id="18c27-180">另请参阅</span><span class="sxs-lookup"><span data-stu-id="18c27-180">See also</span></span>

- [<span data-ttu-id="18c27-181">Excel JavaScript API 核心概念</span><span class="sxs-lookup"><span data-stu-id="18c27-181">Excel JavaScript API core concepts</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="18c27-182">Excel JavaScript API 引用</span><span class="sxs-lookup"><span data-stu-id="18c27-182">Excel JavaScript API reference</span></span>](https://docs.microsoft.com/javascript/api/excel)
