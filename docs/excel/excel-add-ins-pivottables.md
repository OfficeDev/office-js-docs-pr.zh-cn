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
# <a name="work-with-pivottables-using-the-excel-javascript-api"></a>使用 Excel JavaScript API 处理数据透视表

数据透视表可简化大型数据集。它们允许快速操作分组的数据。Excel JavaScript API 允许加载项创建数据透视表并与其组件进行交互。 

如果不熟悉数据透视表的功能，尝试以最终用户的身份了解它们的功能。请参阅 [创建数据透视表以分析工作表数据](https://support.office.com/en-us/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) ，了解这些工具的入门指导。 

本文提供了常见方案的代码示例。若要进一步了解数据透视表 API，请参阅 [**数据透视表**](https://docs.microsoft.com/javascript/api/excel/excel.pivottable) 和 [**PivotTableCollection**](https://docs.microsoft.com/javascript/api/excel/excel.pivottable) 。

> [!IMPORTANT]
> 目前不支持使用 OLAP 创建的数据透视表。

## <a name="hierarchies"></a>层次结构

数据透视表根据四个层次结构类别进行组织：行、列、数据和筛选器。以下描述不同农场的水果销售的数据将被用于本文全文中。

![来自不同农场的不同类型水果销售的集合。](../images/excel-pivots-raw-data.png)

此数据具有五个层次结构： **Farms** 、 **Type** 、 **Classification** 、 **Crates Sold at Farm** 和 **Crates Sold Wholesale** 。每个层次结构可以仅存在于四种类型之一。如果 **Type** 被添加到列层次结构，然后又被添加到行层次结构，则它仅在后者保留。

行和列的层次结构定义了数据如何被分组。例如， **Farms** 的行层次结构将同一个农场的所有数据集组合在一起。行和列的层次结构之间的选择定义了数据透视表的方向。

数据层次结构是基于行和列层次结构进行聚合的值。具有 **Farms** 的行层次结构和 **Crates Sold Wholesale** 的数据层次结构的数据透视表显示了每个农场的所有不同水果的总和（默认）。

筛选器层次结构根据该筛选类型中的值包含或排除枢纽的数据。选择 **Organic** 类型的 **Classification** 筛选器层次结构仅显示有机水果的数据。

这里又是农场数据，以及数据透视表。数据透视表使用 **Farm** 和 **Type** 作为行层次结构， **Crates Sold at Farm** 和 **Crates Sold Wholesale** 作为数据层次结构 （带总和的默认汇总函数），以及 **Classification**  作为筛选器层次结构（选中 **Organic**）。 

![具有行、数据和筛选器层次结构的数据透视表旁边的水果销售数据的选定内容。](../images/excel-pivot-table-and-data.png)

无法通过 JavaScript API 或 Excel 用户界面生成此数据透视表。这两个选项允许通过加载项进行进一步的操作。

## <a name="create-a-pivottable"></a>创建数据透视表

数据透视表的需要名称、 源和目标。源可以是区域地址或表名称（作为 `Range` 、 `string`、或 `Table` 类型进行传递)。目标是某一区域地址 (作为 `Range` 或 `string` 给定)。下面的示例显示各种数据透视表创建技术。

### <a name="create-a-pivottable-with-range-addresses"></a>创建带范围地址的数据透视表

```typescript
await Excel.run(async (context) => {
    // creating a PivotTable named "Farm Sales" on the current worksheet at cell A22 with data from the range A1:E21
    context.workbook.worksheets.getActiveWorksheet().pivotTables.add("Farm Sales", "A1:E21", "A22");

    await context.sync();
});
```

### <a name="create-a-pivottable-with-range-objects"></a>创建带范围对象的数据透视表

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

### <a name="create-a-pivottable-at-the-workbook-level"></a>在工作簿级别创建数据透视表

```typescript
await Excel.run(async (context) => {
    // creating a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data is from the worksheet "DataWorksheet" across the range A1:E21
    context.workbook.pivotTables.add("Farm Sales", "DataWorksheet!A1:E21", "PivotWorksheet!A2");

    await context.sync();
});
```

## <a name="use-an-existing-pivottable"></a>使用现有数据透视表

手动创建的数据透视表亦可通过工作簿或者单个工作表的数据透视表集合进行访问。 

以下代码获取工作簿中第一个数据透视表。然后给出了表的名称以方便用于将来的参考。

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.pivotTables.getItem("My Pivot");
    await context.sync();
});
```

## <a name="add-rows-and-columns-to-a-pivottable"></a>向数据透视表添加行和列

行和列按字段值相关的方式透视数据。

添加 **Farm** 列使所有销售数据以每个农场为中心运行。添加 **Type** 和 **Classification** 行进一步根据哪些水果已售出以及它是否是有机的，以此对数据进行分解。

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

还可拥有仅带行或列的数据透视表。

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));
    
    await context.sync();
});
```

## <a name="add-data-hierarchies-to-the-pivottable"></a>向数据透视表添加数据层次结构

数据层次结构根据行和列在数据透视表中填充了信息。添加 **Crates Sold at Farm** 和 **Crates Sold Wholesale** 的数据层次结构为每行和列提供了那些图表的总和。 

在示例中， **Farm** 和 **Type** 是行，而售出箱数属于数据。 

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

## <a name="change-aggregation-function"></a>更改汇总函数

数据层次结构可以聚合其值。对于数字的数据集，这是默认情况下的总和。 `summarizeBy` 属性根据 `AggregrationFunction` 类型定义了此行为。 

当前支持的汇总函数类型为 `Sum` 、 `Count` 、 `Average` 、 `Max` 、 `Min` 、 `Product` 、 `CountNumbers` 、 `StandardDeviation` 、 `StandardDeviationP` 、 `Variance` 、 `VarianceP` 、 和 `Automatic` （默认）。

以下代码示例将数据汇总更改为平均值。

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

## <a name="change-calculations-with-a-showasrule"></a>更改 ShowAsRule 计算

数据透视表在默认情况下独立聚合了它们行和列的层次结构的数据。一个 `ShowAsRule` 根据数据透视表中的其他项改变了数据层次结构的输出值。

此 `ShowAsRule` 对象具有三种属性：
-   `calculation`：相对于数据层次结构的计算类型（默认是 `none`）。
-   `baseField`: 计算之前包含了基准数据的层次结构内的字段会被应用。 `PivotField` 通常具有其父层次结构相同的名称。
-   `baseItem`: 根据计算类型与基本字段的值进行比较的单项。并非所有计算都需要此字段。

下面的示例将 **Sum of Crates Sold at Farm** 数据层次结构的计算设置为列总计的百分比。我们仍希望此粒度扩展到水果类型的级别，因此我们将使用 **Type** 行层次结构及其基础字段。此示例还将 **Farm** 作为第一个行层次结构，因此农场的总项还显示了每个农场负责生产的百分比。

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

上面的示例将计算设置为列中，与单个行层次结构相对。当计算与单项相关时，使用 `baseItem` 属性。 

下面的示例演示了 `differenceFrom` 计算。它会显示相对于“A Farms”的农场售出箱数的数据层次结构条目的差异。 `baseField` 是 **Farm** ，因此我们可以看到其他农场之间的不同，以及每种类型的类似水果（**Type**  也是在此示例中的行层次结构） 的明细。

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

## <a name="pivottable-layouts"></a>数据透视表布局

数据透视表布局定义了层次结构和其数据的位置。您可访问此布局以确定存储数据的区域。 

下图显示了哪个布局函数调用对应哪个数据透视表范围。

![此图显示数据透视表的哪些部分是由布局的获取范围函数返回的。](../images/excel-pivots-layout-breakdown.png)

下面的代码演示如何通过仔细检查布局来获取数据透视表数据的最后一行。这些值随后相加以得出总和。

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

数据透视表有三种布局样式：压缩、大纲和表格。我们已在上面的示例看到了压缩样式。 

下面的示例分别使用大纲和表格样式。代码示例演示如何在不同的布局之间循环。

### <a name="outline-layout"></a>大纲版式

![使用大纲版式的数据透视表。](../images/excel-pivots-outline-layout.png)

### <a name="tabular-layout"></a>表格版式

![使用表格版式的数据透视表。](../images/excel-pivots-tabular-layout.png)

## <a name="change-hierarchy-names"></a>更改层次结构名称

层次结构字段可以编辑。下面的代码演示了如何更改两个数据层次结构的显示名称。

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

## <a name="delete-a-pivottable"></a>删除数据透视表

使用数据透视表的名称删除数据透视表。

```typescript
await Excel.run(async (context) => {
    context.workbook.worksheets.getItem("Pivot").pivotTables.getItem("Farm Sales").delete();

    await context.sync();
});
```

## <a name="see-also"></a>另请参阅

- [使用 Excel JavaScript API 的基本编程概念](excel-add-ins-core-concepts.md)
- [Excel JavaScript API 参考](https://docs.microsoft.com/javascript/api/excel)
