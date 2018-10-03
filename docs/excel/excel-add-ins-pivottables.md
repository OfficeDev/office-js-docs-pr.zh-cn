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
# <a name="work-with-pivottables-using-the-excel-javascript-api"></a>使用 Excel JavaScript API 处理数据透视表

数据透视表可简化更大的数据集。 它们允许分组数据的快速操作。 Excel JavaScript API 允许加载项创建数据透视表并与其组件交互。 

如果不熟悉数据透视表的功能，尝试以最终用户的身份了解它们的功能。 请参阅[创建数据透视表分析工作表数据](https://support.office.com/en-us/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables)，了解这些工具的入门指导。 

本文提供了常见方案的代码示例。 参阅[**PivotTable**](https://docs.microsoft.com/javascript/api/excel/excel.pivottable) 和 [**PivotTableCollection**](https://docs.microsoft.com/javascript/api/excel/excel.pivottable)，加深对数据透视表 API 的理解。

> [!IMPORTANT]
> 目前不支持使用 OLAP 创建的数据透视表。

## <a name="hierarchies"></a>层次结构

数据透视表基于四种层次结构类别构成：行、列、数据和筛选器。 本文将通篇使用以下描述各农场水果销售情况的数据。

![来自不同农场的不同类型水果销售的集合。](../images/excel-pivots-raw-data.png)

此数据具有五个层次结构：**Farm**、**Type**、**Classification**、**Crates Sold at Farm** 和 **Crates Sold Wholesale**。 每个层次结构只能存在于四个类别中的一个类别。 如果 **Type** 添加到列层次结构，然后又添加到行层次结构，则其仅保留于后者。

行和列的层次结构定义如何分组数据。 例如，**Farms** 的行层次结构会将来自同一农场的所有数据集归集在一起。 选择行和列层次结构来定义数据透视表的方向。

数据层次结构是基于行和列层次结构进行聚合的值。 具有 **Farms** 的行层次结构和 **Crates Sold Wholesale** 的数据层次结构的数据透视表显示每个农场所有不同水果的总和（默认）。

筛选器层次结构基于已筛选类型中的值包含或排除来自透视的数据。 选择了 **Organic** 类型的 **Classification** 筛选器层次结构仅显示有机水果的数据。

这同样是农场数据，一旁是数据透视表。 数据透视表使用 **Farm** 和 **Type** 作为行层次结构，**Crates Sold at Farm** 和 **Crates Sold Wholesale** 作为数据层次结构 （带默认的 sum 汇总函数），**Classification** 作为筛选器层次结构（选中 **Organic**）。 

![具有行、数据和筛选器层次结构的数据透视表旁边的水果销售数据的选定内容。](../images/excel-pivot-table-and-data.png)

可通过 JavaScript API 或通过 Excel UI 生成这个数据透视表。 两个选项均可通过加载项实现进一步的操作。

## <a name="create-a-pivottable"></a>创建数据透视表

数据透视表需要有名称、源和目标。 源可以是范围地址或表名（作为 `Range`、`string` 或 `Table`类型传递）。 目标是某一范围地址（作为`Range` 或 `string`给定）。 以下示例显示各种数据透视表的创建技术。

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

以下代码获取工作簿中的第一个数据透视表。 然后给出了表的名称，便于以后参考。

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.pivotTables.getItem("My Pivot");
    await context.sync();
});
```

## <a name="add-rows-and-columns-to-a-pivottable"></a>向数据透视表添加行和列

行和列按字段值相关的方式透视数据。

添加 **Farm** 列可按每个农场的所有销售情况透视数据。 添加 **Type** 和 **Classification** 行，可基于销售的水果以及该水果是否为有机等条件而将数据作进一步的分解。

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

基于行和列，数据层次结构结合信息填充数据透视表。 添加**Crates Sold at Farm** 和 **Crates Sold Wholesale** 的数据层次结构给出每行和每列数字的总和。 

在示例中，**Farm** 和 **Type** 都是行，销售箱数作为数据。 

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

数据层次结构将其数值汇总。 对于数字的数据集，默认情况下，这是总和。 `summarizeBy` 属性基于 `AggregrationFunction` 类型定义此行为。 

当前支持的汇总函数类型为 `Sum`、`Count`、`Average`、`Max` `Min`、`Product`、`CountNumbers`、`StandardDeviation`、`StandardDeviationP`、`Variance`、`VarianceP` 和 `Automatic`（默认）。

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

数据透视表默认情况下独立汇总其行和列层次结构的数据。 `ShowAsRule` 数据层次结构更改为基于数据透视表中的其他项输出值。

 `ShowAsRule` 对象具有三个属性：
-   `calculation`：相对于数据层次结构的计算类型（默认是 `none`）。
-   `baseField`：应用计算前层次结构中包含基准数据的字段。  `PivotField` 通常与其父层次结构具有相同的名称。
-   `baseItem`：根据计算类型的基本字段的值进行比较的单个项。 并非所有计算都需要此字段。

下面的示例将对 **Sum of Crates Sold at Farm** 数据层次结构列执行的计算设置为列总计的百分比。 我们仍希望将粒度级别扩展至水果类型，因此我们将使用 **Type** 行层次结构及其基础字段。 此示例还以 **Farm** 作为第一个行层次结构，因此农场的总项也显示每个农场负责产出的百分比。

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

上面的示例将计算设置为列中相对于单个行层次结构。 当计算与单项相关时，使用 `baseItem` 属性。 

下面的示例演示 `differenceFrom` 计算。 它会显示农场销售箱数层次结构条目相对于 "A Farms" 的差异。  `baseField` 是 **Farm**，以便我们可以看到其他农场之间的差异，以及每种类似果品（**Type** 也是在此示例中的行层次结构）之间的差异。

![显示 "A Farms" 和其他农场水果销售之间差异的数据透视表。 它同时显示农场的水果销售总额和果品销售额的差异。 如果 "A Farms" 未不销售某一类型的水果，将显示 "# n/A"。](../images/excel-pivots-showas-differencefrom.png)

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

数据透视表布局定义层次结构及其数据的位置。 访问布局来确定存储数据区域的范围。 

下图显示了哪个布局函数调用对应哪个数据透视表范围。

![此图显示数据透视表的哪些部分是由布局的获取范围函数返回的。](../images/excel-pivots-layout-breakdown.png)

下面的代码演示了如何通过布局获取数据透视表数据的最后一行。 然后对这些值求和获得总计。

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

数据透视表有三种布局样式：压缩、大纲和表格。 在上面的示例中我们看到过压缩样式。 

下面的示例分别使用大纲和表格样式。 代码示例显示如何在不同的布局之间转换。

### <a name="outline-layout"></a>大纲版式

![使用大纲版式的数据透视表。](../images/excel-pivots-outline-layout.png)

### <a name="tabular-layout"></a>表格版式

![使用表格版式的数据透视表。](../images/excel-pivots-tabular-layout.png)

## <a name="change-hierarchy-names"></a>更改层次结构名称

层次结构字段为可编辑。 下面的代码演示如何交换两个数据层次结构的显示名称。

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

- [Excel JavaScript API 核心概念](excel-add-ins-core-concepts.md)
- [Excel JavaScript API 参考](https://docs.microsoft.com/javascript/api/excel)
