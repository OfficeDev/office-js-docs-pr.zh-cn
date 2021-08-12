---
title: 使用 JavaScript API Excel数据透视表
description: 使用 Excel JavaScript API 创建数据透视表并与其组件交互。
ms.date: 07/02/2021
localization_priority: Normal
ms.openlocfilehash: 7be7fe8a4f7dcb2509943f7fd03fbb8739312e87874583bbe97b8139ab83c6b5
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/07/2021
ms.locfileid: "57084243"
---
# <a name="work-with-pivottables-using-the-excel-javascript-api"></a>使用 JavaScript API Excel数据透视表

数据透视表可简化更大的数据集。 它们允许对分组数据进行快速操作。 借助 Excel JavaScript API，加载项可以创建数据透视表并与其组件交互。 本文介绍数据透视表如何由 Office JavaScript API 表示，并提供关键方案的代码示例。

如果您不熟悉数据透视表的功能，请考虑以最终用户模式探索它们。
有关 [这些工具的良好基础，](https://support.office.com/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) 请参阅创建数据透视表以分析工作表数据。

> [!IMPORTANT]
> 当前不支持使用 OLAP 创建的数据透视表。 也不支持 Power Pivot。

## <a name="object-model"></a>对象模型

数据[透视表](/javascript/api/excel/excel.pivottable)是 JavaScript API 中数据透视表Office对象。

- `Workbook.pivotTables`和 是分别包含工作簿和工作表中的 `Worksheet.pivotTables` 数据透视[](/javascript/api/excel/excel.pivottable)表的[PivotTableCollection。](/javascript/api/excel/excel.pivottablecollection)
- 数据[透视表](/javascript/api/excel/excel.pivottable)包含具有[多个 PivotHierarchies 的 PivotHierarchyCollection。](/javascript/api/excel/excel.pivothierarchycollection) [](/javascript/api/excel/excel.pivothierarchy)
- 可以将[这些 PivotHierarchies](/javascript/api/excel/excel.pivothierarchy)添加到特定层次结构集合中，以定义数据透视表 (数据透视表数据透视表) 。 [](#hierarchies)
- [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy)包含一个只有一个 PivotField 的[PivotFieldCollection。](/javascript/api/excel/excel.pivotfieldcollection) [](/javascript/api/excel/excel.pivotfield) 如果设计扩展为包含 OLAP 数据透视表，这可能会更改。
- 只要[将字段](/javascript/api/excel/excel.pivotfield)的[PivotHierarchy](/javascript/api/excel/excel.pivothierarchy)分配给层次结构类别，透视字段就可以应用一个或多个[PivotFilter。](/javascript/api/excel/excel.pivotfilters)
- 透视[字段](/javascript/api/excel/excel.pivotfield)包含具有[多个 PivotItems 的 PivotItemCollection。](/javascript/api/excel/excel.pivotitemcollection) [](/javascript/api/excel/excel.pivotitem)
- [数据透视表](/javascript/api/excel/excel.pivottable)包含[一个 PivotLayout，](/javascript/api/excel/excel.pivotlayout)它定义[透视字段](/javascript/api/excel/excel.pivotfield)和[PivotItems](/javascript/api/excel/excel.pivotitem)在工作表中的显示位置。 布局还控制数据透视表的一些显示设置。

让我们看一下这些关系如何应用于一些示例数据。 以下数据描述了来自各种服务器场的菜品销售情况。 它将是整篇文章中的示例。

![不同服务器场中不同类型的新鲜品销售的集合。](../images/excel-pivots-raw-data.png)

此菜场销售数据将用于创建数据透视表。 每列（如 **Types）** 都是 `PivotHierarchy` 。 " **类型** "层次结构包含 **"类型"** 字段。 The **Types** field contains the items **Apple**， **Kiwi**， **Orange**， **Orange**， and **Orange**.

### <a name="hierarchies"></a>Hierarchies

数据透视表基于四个层次结构类别进行组织：[行](/javascript/api/excel/excel.rowcolumnpivothierarchy)、[列](/javascript/api/excel/excel.rowcolumnpivothierarchy)[、数据和](/javascript/api/excel/excel.datapivothierarchy)[筛选器](/javascript/api/excel/excel.filterpivothierarchy)。

前面显示的服务器场数据有五个层次结构：Farms、Type、Classification、Crates **Sold at Farm** 和 **Crates SoldRate。**    每个层次结构只能存在于四个类别之一中。 如果将 **Type** 添加到列层次结构中，则不能也位于行、数据或筛选器层次结构中。 如果 **Type** 随后添加到行层次结构中，则从列层次结构中删除它。 无论层次结构分配是通过 Excel UI 还是 JavaScript API Excel，此行为都是相同的。

行和列层次结构定义数据的分组方法。 例如，服务器场的行 **层次结构将同** 一服务器场的所有数据集组合在一起。 行和列层次结构之间的选择定义数据透视表的方向。

数据层次结构是基于行和列层次结构聚合的值。 具有服务器场的行层次结构和数据层次结构"销售商品"的数据层次结构的数据透视表显示每个服务器场的所有不同客户 () 的默认总和。

筛选器层次结构基于该筛选类型中的值包含或排除数据透视表中的数据。 "分类"的筛选器 **层次结构（** 已选择 **"有机** "类型）只显示有机菜的数据。

下面再次是数据透视表旁的服务器场数据。 数据透视表使用 **"服务器场**"和"类型"作为行层次结构，将"场中销售"和"出售的百货"作为数据层次结构 (其默认聚合函数为 sum) ，将 **Classification** 用作筛选器层次结构 (（选择 **"Organic**) "）。

![数据透视表旁边包含一组包含行、数据和筛选器层次结构的新鲜菜销售数据。](../images/excel-pivot-table-and-data.png)

此数据透视表可以通过 JavaScript API 或 Excel UI 生成。 这两个选项都允许通过外接程序进一步操作。

## <a name="create-a-pivottable"></a>创建数据透视表

数据透视表需要名称、源和目标。 源可以是区域地址或表名称， (、 或 键入 `Range` `string` `Table`) 。 目标地址是一个范围 (给定为 `Range` 或 `string`) 。
以下示例显示了各种数据透视表创建技术。

### <a name="create-a-pivottable-with-range-addresses"></a>创建具有区域地址的数据透视表

```js
Excel.run(function (context) {
    // Create a PivotTable named "Farm Sales" on the current worksheet at cell
    // A22 with data from the range A1:E21.
    context.workbook.worksheets.getActiveWorksheet().pivotTables.add(
      "Farm Sales", "A1:E21", "A22");

    return context.sync();
});
```

### <a name="create-a-pivottable-with-range-objects"></a>创建包含 Range 对象的数据透视表

```js
Excel.run(function (context) {
    // Create a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data comes from the worksheet "DataWorksheet" across the range A1:E21.
    var rangeToAnalyze = context.workbook.worksheets.getItem("DataWorksheet").getRange("A1:E21");
    var rangeToPlacePivot = context.workbook.worksheets.getItem("PivotWorksheet").getRange("A2");
    context.workbook.worksheets.getItem("PivotWorksheet").pivotTables.add(
      "Farm Sales", rangeToAnalyze, rangeToPlacePivot);

    return context.sync();
});
```

### <a name="create-a-pivottable-at-the-workbook-level"></a>在工作簿级别创建数据透视表

```js
Excel.run(function (context) {
    // Create a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data is from the worksheet "DataWorksheet" across the range A1:E21.
    context.workbook.pivotTables.add(
        "Farm Sales", "DataWorksheet!A1:E21", "PivotWorksheet!A2");

    return context.sync();
});
```

## <a name="use-an-existing-pivottable"></a>使用现有数据透视表

通过工作簿或单个工作表的数据透视表集合，也可以访问手动创建的数据透视表。 下面的代码从工作簿获取名为 **My Pivot** 的数据透视表。

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.pivotTables.getItem("My Pivot");
    return context.sync();
});
```

## <a name="add-rows-and-columns-to-a-pivottable"></a>向数据透视表添加行和列

行和列围绕这些字段的值透视数据。

添加 **"服务器场** "列可透视每个服务器场的所有销售额。 添加 **"类型** "和" **分类** "行会进一步分解数据，这些数据基于所出售的菜以及它是否是有机的。

![具有"服务器场"列和"类型"和"分类"行的数据透视表。](../images/excel-pivots-table-rows-and-columns.png)

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));

    pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem("Farm"));

    return context.sync();
});
```

也可以只包含行或列的数据透视表。

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));

    return context.sync();
});
```

## <a name="add-data-hierarchies-to-the-pivottable"></a>向数据透视表添加数据层次结构

数据层次结构使用基于行和列组合的信息填充数据透视表。 添加"场出售"和"出售 **的** 箱形"的数据层次结构会提供每一行和每一列的数据总和。

在示例中 **，Farm** 和 **Type** 都是行，以销售为数据。

![一个数据透视表，它基于不同动物的服务器场显示其总销售额。](../images/excel-pivots-data-hierarchy.png)

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    // "Farm" and "Type" are the hierarchies on which the aggregation is based.
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));

    // "Crates Sold at Farm" and "Crates Sold Wholesale" are the hierarchies
    // that will have their data aggregated (summed in this case).
    pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Crates Sold at Farm"));
    pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Crates Sold Wholesale"));

    return context.sync();
});
```

## <a name="pivottable-layouts-and-getting-pivoted-data"></a>数据透视表布局和获取透视数据

[PivotLayout](/javascript/api/excel/excel.pivotlayout)定义层次结构及其数据的位置。 访问布局以确定存储数据的范围。

下图显示了哪些布局函数调用对应于数据透视表的哪些区域。

![显示布局的 get range 函数返回数据透视表的哪些节的图表。](../images/excel-pivots-layout-breakdown.png)

### <a name="get-data-from-the-pivottable"></a>从数据透视表获取数据

布局定义数据透视表在工作表中的显示方式。 这意味着对象 `PivotLayout` 控制用于数据透视表元素的范围。 使用布局提供的范围获取数据透视表收集和聚合的数据。 特别是，使用 `PivotLayout.getDataBodyRange` 访问数据透视表生成的数据。

下面的代码演示了如何获取数据透视表数据的最后一行，方法为在上一示例) 中浏览布局 (服务器场中"销售的箱形总和"和"销售的箱形总和"列。  然后，这些值汇总在一起，最终总计显示在数据透视表数据透视表外部 (**E30** 单元格) 。

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    // Get the totals for each data hierarchy from the layout.
    var range = pivotTable.layout.getDataBodyRange();
    var grandTotalRange = range.getLastRow();
    grandTotalRange.load("address");
    return context.sync().then(function () {
        // Sum the totals from the PivotTable data hierarchies and place them in a new range, outside of the PivotTable.
        var masterTotalRange = context.workbook.worksheets.getActiveWorksheet().getRange("E30");
        masterTotalRange.formulas = [["=SUM(" + grandTotalRange.address + ")"]];
    });
});
```

### <a name="layout-types"></a>布局类型

数据透视表具有三种布局样式：精简、大纲和表格。 我们已在之前的示例中看到简洁样式。

下面的示例分别使用大纲样式和表格样式。 该代码示例演示如何在不同布局之间循环。

#### <a name="outline-layout"></a>大纲布局

![使用大纲布局的数据透视表。](../images/excel-pivots-outline-layout.png)

#### <a name="tabular-layout"></a>表格布局

![使用表格布局的数据透视表。](../images/excel-pivots-tabular-layout.png)

#### <a name="pivotlayout-type-switch-code-sample"></a>PivotLayout 类型开关代码示例

```js
Excel.run(function (context) {
    // Change the PivotLayout.type to a new type.
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.layout.load("layoutType");
    return context.sync().then(function () {
        // Cycle between the three layout types.
        if (pivotTable.layout.layoutType === "Compact") {
            pivotTable.layout.layoutType = "Outline";
        } else if (pivotTable.layout.layoutType === "Outline") {
            pivotTable.layout.layoutType = "Tabular";
        } else {
            pivotTable.layout.layoutType = "Compact";
        }
    
        return context.sync();
    });
});
```

### <a name="other-pivotlayout-functions"></a>其他 PivotLayout 函数

默认情况下，数据透视表会根据需要调整行和列大小。 此操作在数据透视表刷新时完成。 `PivotLayout.autoFormat` 指定该行为。 当 为 时，加载项进行的任何行或列大小更改将 `autoFormat` 持续存在 `false` 。 此外，数据透视表的默认设置在数据透视表中保留任何自定义 (如填充和字体) 。 设置为 `PivotLayout.preserveFormatting` `false` 以在刷新时应用默认格式。

还 `PivotLayout` 控制标题和总行设置、空数据单元格的显示方式以及 [替换文字](https://support.microsoft.com/topic/add-alternative-text-to-a-shape-picture-chart-smartart-graphic-or-other-object-44989b2a-903c-4d9a-b742-6a75b451c669) 选项。 [PivotLayout](/javascript/api/excel/excel.pivotlayout)引用提供了这些功能的完整列表。

下面的代码示例使空数据单元格显示字符串，将正文区域的格式设置为一致的水平对齐方式，并确保即使在数据透视表刷新后，格式更改 `"--"` 也保持不变。

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.pivotTables.getItem("Farm Sales");
    var pivotLayout = pivotTable.layout;

    // Set a default value for an empty cell in the PivotTable. This doesn't include cells left blank by the layout.
    pivotLayout.emptyCellText = "--";

    // Set the text alignment to match the rest of the PivotTable.
    pivotLayout.getDataBodyRange().format.horizontalAlignment = Excel.HorizontalAlignment.right;

    // Ensure empty cells are filled with a default value.
    pivotLayout.fillEmptyCells = true;

    // Ensure that the format settings persist, even after the PivotTable is refreshed and recalculated.
    pivotLayout.preserveFormatting = true;
    return context.sync();
});
```

## <a name="delete-a-pivottable"></a>删除数据透视表

使用数据透视表的名称删除数据透视表。

```js
Excel.run(function (context) {
    context.workbook.worksheets.getItem("Pivot").pivotTables.getItem("Farm Sales").delete();
    return context.sync();
});
```

## <a name="filter-a-pivottable"></a>筛选数据透视表

筛选数据透视表数据的主要方法是使用 PivotFilter。 切片器提供了一种不太灵活的备用筛选方法。

[PivotFilters](/javascript/api/excel/excel.pivotfilters)基于数据透视表的四个层次结构类别[](#hierarchies)筛选数据 (筛选器、列、行和) 。 有四种类型的 PivotFilter，允许基于日历日期的筛选、字符串分析、数字比较和基于自定义输入的筛选。

[切片器](/javascript/api/excel/excel.slicer)可应用于数据透视表和常规Excel表。 应用于数据透视表时，切片器的功能与 [PivotManualFilter](#pivotmanualfilter) 类似，并允许基于自定义输入进行筛选。 与 PivotFilter 不同，切片器具有Excel [UI 组件](https://support.office.com/article/Use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d)。 使用 `Slicer` 类，你可以创建此 UI 组件、管理筛选并控制其视觉外观。

### <a name="filter-with-pivotfilters"></a>使用 PivotFilter 进行筛选

[PivotFilters](/javascript/api/excel/excel.pivotfilters)允许您基于四个层次结构类别筛选数据透视表[](#hierarchies)数据 (筛选器、列、行和) 。 在数据透视表对象模型中， `PivotFilters` 应用到透视 [字段](/javascript/api/excel/excel.pivotfield)，并且 `PivotField` 每个都可以分配一个或多个 `PivotFilters` 。 若要将 PivotFilter 应用于透视字段，必须将字段对应的 [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) 分配给层次结构类别。

#### <a name="types-of-pivotfilters"></a>PivotFilter 的类型

| 筛选器类型 | 筛选目的 | Excel JavaScript API 参考 |
|:--- |:--- |:--- |
| DateFilter | 基于日历日期的筛选。 | [PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter) |
| LabelFilter | 文本比较筛选。 | [PivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter) |
| ManualFilter | 自定义输入筛选。 | [PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter) |
| ValueFilter | 数字比较筛选。 | [PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter) |

#### <a name="create-a-pivotfilter"></a>创建 PivotFilter

若要使用数据透视表等 (`Pivot*Filter` 数据透视表 `PivotDateFilter`) ，请对透视字段应用 [筛选器](/javascript/api/excel/excel.pivotfield)。 以下四个代码示例显示如何使用四种类型的 PivotFilter。

##### <a name="pivotdatefilter"></a>PivotDateFilter

第一个代码示例将 [PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter) 应用于 **Date Updated** PivotField，隐藏 **2020-08-01 之前的任何数据**。

> [!IMPORTANT]
> `Pivot*Filter`无法将 应用于透视字段，除非将该字段的 PivotHierarchy 分配给层次结构类别。 在下面的代码示例中，必须先将 添加到数据透视表的类别中，然后才能 `dateHierarchy` `rowHierarchies` 用于筛选。

```js
Excel.run(function (context) {
    // Get the PivotTable and the date hierarchy.
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    var dateHierarchy = pivotTable.rowHierarchies.getItemOrNullObject("Date Updated");
    
    return context.sync().then(function () {
        // PivotFilters can only be applied to PivotHierarchies that are being used for pivoting.
        // If it's not already there, add "Date Updated" to the hierarchies.
        if (dateHierarchy.isNullObject) {
          dateHierarchy = pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Date Updated"));
        }

        // Apply a date filter to filter out anything logged before August.
        var filterField = dateHierarchy.fields.getItem("Date Updated");
        var dateFilter = {
          condition: Excel.DateFilterCondition.afterOrEqualTo,
          comparator: {
            date: "2020-08-01",
            specificity: Excel.FilterDatetimeSpecificity.month
          }
        };
        filterField.applyFilter({ dateFilter: dateFilter });
        
        return context.sync();
    });
});
```

> [!NOTE]
> 以下三个代码段仅显示特定于筛选器的摘要，而不是完整 `Excel.run` 调用。

##### <a name="pivotlabelfilter"></a>PivotLabelFilter

第二个代码段演示如何将 [PivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter) 应用到 **Type** PivotField，使用 属性排除以 `LabelFilterCondition.beginsWith` 字母 L 开头 **的标签**。

```js
    // Get the "Type" field.
    var filterField = pivotTable.hierarchies.getItem("Type").fields.getItem("Type");

    // Filter out any types that start with "L" ("Lemons" and "Limes" in this case).
    var filter: Excel.PivotLabelFilter = {
      condition: Excel.LabelFilterCondition.beginsWith,
      substring: "L",
      exclusive: true
    };

    // Apply the label filter to the field.
    filterField.applyFilter({ labelFilter: filter });
```

##### <a name="pivotmanualfilter"></a>PivotManualFilter

第三个代码段使用 [PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter) 将手动筛选器应用于 **Classification** 字段，以筛选出不包含 **分类 Organic 的数据**。

```js
    // Apply a manual filter to include only a specific PivotItem (the string "Organic").
    var filterField = classHierarchy.fields.getItem("Classification");
    var manualFilter = { selectedItems: ["Organic"] };
    filterField.applyFilter({ manualFilter: manualFilter });
```

##### <a name="pivotvaluefilter"></a>PivotValueFilter

若要比较数字，请将值筛选器与 [PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter)一同使用，如最终代码片段中所示。 比较服务器场透视字段的数据与"已出售的 Crate PivotField"数据透视表中的数据，仅包括销售的箱总和超过 `PivotValueFilter` **值 500 的服务器场**。  

```js
    // Get the "Farm" field.
    var filterField = pivotTable.hierarchies.getItem("Farm").fields.getItem("Farm");
    
    // Filter to only include rows with more than 500 wholesale crates sold.
    var filter: Excel.PivotValueFilter = {
      condition: Excel.ValueFilterCondition.greaterThan,
      comparator: 500,
      value: "Sum of Crates Sold Wholesale"
    };
    
    // Apply the value filter to the field.
    filterField.applyFilter({ valueFilter: filter });
```

#### <a name="remove-pivotfilters"></a>删除 PivotFilter

若要删除所有 PivotFilter，请对各个 `clearAllFilters` 透视字段应用该方法，如下面的代码示例所示。

```js
Excel.run(function (context) {
    // Get the PivotTable.
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.hierarchies.load("name");
    
    return context.sync().then(function () {
        // Clear the filters on each PivotField.
        pivotTable.hierarchies.items.forEach(function (hierarchy) {
          hierarchy.fields.getItem(hierarchy.name).clearAllFilters();
        });
        return context.sync();
    });
});
```

### <a name="filter-with-slicers"></a>使用切片器筛选

[切片器](/javascript/api/excel/excel.slicer)允许从数据透视表或Excel筛选数据。 切片器使用指定列或透视字段的值筛选相应的行。 这些值存储为 中的 [SlicerItem](/javascript/api/excel/excel.sliceritem) 对象 `Slicer` 。 加载项可以调整这些筛选器，就像用户 (UI Excel[一](https://support.office.com/article/Use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d)) 。 切片器位于绘图层中工作表的顶部，如以下屏幕截图所示。

![筛选数据透视表上的数据的切片器。](../images/excel-slicer.png)

> [!NOTE]
> 本节中介绍的技术侧重于如何使用连接到数据透视表的切片器。 相同的技术也适用于使用连接到表的切片器。

#### <a name="create-a-slicer"></a>创建切片器

可以使用 方法在工作簿或工作表中 `Workbook.slicers.add` 创建切片 `Worksheet.slicers.add` 器。 这样做会将切片器添加到指定 或 对象的[SlicerCollection。](/javascript/api/excel/excel.slicercollection) `Workbook` `Worksheet` `SlicerCollection.add`方法具有三个参数：

- `slicerSource`：新切片器所基于的数据源。 它可以是 `PivotTable` 、 `Table` 或 字符串，表示 或 的名称或 `PivotTable` `Table` ID。
- `sourceField`：数据源中要筛选的字段。 它可以是 `PivotField` 、 `TableColumn` 或 字符串，表示 或 的名称或 `PivotField` `TableColumn` ID。
- `slicerDestination`：将在其中新建切片器的工作表。 它可以是 `Worksheet` 对象，或者是 的名称或 `Worksheet` ID。 通过 访问 时， `SlicerCollection` 不需要此参数 `Worksheet.slicers` 。 在这种情况下，集合的工作表用作目标。

下面的代码示例向透视工作表 **添加新切片器** 。 切片器的来源是场 **销售** 数据透视表，并且使用 **Type 数据进行** 筛选。 该切片器也命名为 **"切片器** "，供将来参考。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Pivot");
    var slicer = sheet.slicers.add(
        "Farm Sales" /* The slicer data source. For PivotTables, this can be the PivotTable object reference or name. */,
        "Type" /* The field in the data to filter by. For PivotTables, this can be a PivotField object reference or ID. */
    );
    slicer.name = "Fruit Slicer";
    return context.sync();
});
```

#### <a name="filter-items-with-a-slicer"></a>使用切片器筛选项

切片器使用 中的项筛选数据透视表 `sourceField` 。 `Slicer.selectItems`方法设置保留在切片器中的项。 这些项目作为 传递给 方法， `string[]` 表示项的键。 包含这些项目的任何行都保留在数据透视表的聚合中。 后续调用 `selectItems` ，用于将列表设置为这些调用中指定的键。

> [!NOTE]
> `Slicer.selectItems`如果传递的项不在数据源中，则 `InvalidArgument` 会引发错误。 可通过 属性（即 `Slicer.slicerItems` [SlicerItemCollection ）验证内容](/javascript/api/excel/excel.sliceritemcollection)。

下面的代码示例显示了为切片器选择的三个项目：**橙色**、**橙色****和橙色**。

```js
Excel.run(function (context) {
    var slicer = context.workbook.slicers.getItem("Fruit Slicer");
    // Anything other than the following three values will be filtered out of the PivotTable for display and aggregation.
    slicer.selectItems(["Lemon", "Lime", "Orange"]);
    return context.sync();
});
```

若要从切片器中删除所有筛选器，请使用 `Slicer.clearFilters` 方法，如以下示例所示。

```js
Excel.run(function (context) {
    var slicer = context.workbook.slicers.getItem("Fruit Slicer");
    slicer.clearFilters();
    return context.sync();
});
```

#### <a name="style-and-format-a-slicer"></a>设置切片器样式和格式

加载项可以通过属性调整切片器显示 `Slicer` 设置。 下面的代码示例将样式设置为 **SlicerStyleLight6**，将切片器顶部的文本设置为 **"菜** 类型"，将切片器放在绘图层上 **位置 (395，15) ，** 将切片器的大小设置为 **135x150** 像素。

```js
Excel.run(function (context) {
    var slicer = context.workbook.slicers.getItem("Fruit Slicer");
    slicer.caption = "Fruit Types";
    slicer.left = 395;
    slicer.top = 15;
    slicer.height = 135;
    slicer.width = 150;
    slicer.style = "SlicerStyleLight6";
    return context.sync();
});
```

#### <a name="delete-a-slicer"></a>删除切片器

若要删除切片器，请调用 `Slicer.delete` 方法。 下面的代码示例从当前工作表中删除第一个切片器。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.slicers.getItemAt(0).delete();
    return context.sync();
});
```

## <a name="change-aggregation-function"></a>更改聚合函数

数据层次结构聚合了它们的值。 对于数字数据集，默认情况下这是一个和。 属性 `summarizeBy` 根据 [AggregationFunction](/javascript/api/excel/excel.aggregationfunction) 类型定义此行为。

当前支持的聚合函数类型是 `Sum` 、和 `Count` (`Average` `Max` `Min` `Product` `CountNumbers` `StandardDeviation` `StandardDeviationP` `Variance` `VarianceP` `Automatic` 默认) 。

下面的代码示例将聚合更改为数据的平均值。

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.dataHierarchies.load("no-properties-needed");
    return context.sync().then(function() {

        // Change the aggregation from the default sum to an average of all the values in the hierarchy.
        pivotTable.dataHierarchies.items[0].summarizeBy = Excel.AggregationFunction.average;
        pivotTable.dataHierarchies.items[1].summarizeBy = Excel.AggregationFunction.average;
        return context.sync();
    });
});
```

## <a name="change-calculations-with-a-showasrule"></a>使用 ShowAsRule 更改计算

默认情况下，数据透视表独立聚合其行和列层次结构的数据。 [ShowAsRule](/javascript/api/excel/excel.showasrule)根据数据透视表中的其他项将数据层次结构更改为输出值。

对象 `ShowAsRule` 有三个属性：

- `calculation`：应用于数据层次结构的相对计算类型 (默认值为 `none`) 。
- `baseField`： [在应用](/javascript/api/excel/excel.pivotfield) 计算之前，层次结构中包含基本数据的透视字段。 由于Excel数据透视表具有到字段的一对一的层次结构映射，因此您将使用相同的名称访问层次结构和字段。
- `baseItem`：单个 [PivotItem](/javascript/api/excel/excel.pivotitem) 与基于计算类型的基字段的值进行比较。 并非所有计算都需要此字段。

下面的示例将场中销售的箱值总 **和数据层次结构** 的计算结果设定为列总计的百分比。
我们仍希望粒度扩展到树状类型级别，因此我们将使用 **Type** 行层次结构及其基础字段。
此示例还将 **Farm** 作为第一行层次结构，因此服务器场的总条目也显示每个服务器场负责生成百分比。

![一个数据透视表，其中显示每个场中各场和各个新鲜菜类型的新鲜菜销售额占总和的百分比。](../images/excel-pivots-showas-percentage.png)

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    var farmDataHierarchy = pivotTable.dataHierarchies.getItem("Sum of Crates Sold at Farm");

    farmDataHierarchy.load("showAs");
    return context.sync().then(function () {

        // Show the crates of each fruit type sold at the farm as a percentage of the column's total.
        var farmShowAs = farmDataHierarchy.showAs;
        farmShowAs.calculation = Excel.ShowAsCalculation.percentOfColumnTotal;
        farmShowAs.baseField = pivotTable.rowHierarchies.getItem("Type").fields.getItem("Type");
        farmDataHierarchy.showAs = farmShowAs;
        farmDataHierarchy.name = "Percentage of Total Farm Sales";
    });
});
```

前面的示例将相对于单个行层次结构的字段的计算设置为列。 当计算与单个项目相关时，请使用 `baseItem` 属性。

以下示例显示了 `differenceFrom` 计算。 它显示服务器场包含销售数据层次结构条目相对于服务器场 **的条目的区别**。
为 Farm，因此我们将看到其他服务器场之间的差异，以及每种类型的类似树 (Type 的细目也是此示例中的行层次结构 `baseField`) 。  

![一个数据透视表，显示"A Farms"和其他服务器场之间的菜品销售差异。 这显示服务器场的菜品总销售额和新鲜菜类型销售额的差值。 如果"A Farms"未出售特定类型的新鲜菜，则#N/A"。](../images/excel-pivots-showas-differencefrom.png)

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    var farmDataHierarchy = pivotTable.dataHierarchies.getItem("Sum of Crates Sold at Farm");

    farmDataHierarchy.load("showAs");
    return context.sync().then(function () {
        // Show the difference between crate sales of the "A Farms" and the other farms.
        // This difference is both aggregated and shown for individual fruit types (where applicable).
        var farmShowAs = farmDataHierarchy.showAs;
        farmShowAs.calculation = Excel.ShowAsCalculation.differenceFrom;
        farmShowAs.baseField = pivotTable.rowHierarchies.getItem("Farm").fields.getItem("Farm");
        farmShowAs.baseItem = pivotTable.rowHierarchies.getItem("Farm").fields.getItem("Farm").items.getItem("A Farms");
        farmDataHierarchy.showAs = farmShowAs;
        farmDataHierarchy.name = "Difference from A Farms";
    });
});
```

## <a name="change-hierarchy-names"></a>更改层次结构名称

层次结构字段是可编辑的。 以下代码演示如何更改两个数据层次结构的显示名称。

```js
Excel.run(function (context) {
    var dataHierarchies = context.workbook.worksheets.getActiveWorksheet()
        .pivotTables.getItem("Farm Sales").dataHierarchies;
    dataHierarchies.load("no-properties-needed");
    return context.sync().then(function () {
        // changing the displayed names of these entries
        dataHierarchies.items[0].name = "Farm Sales";
        dataHierarchies.items[1].name = "Wholesale";
    });
});
```

## <a name="see-also"></a>另请参阅

- [Excel 加载项中的 Word JavaScript 对象模型](excel-add-ins-core-concepts.md)
- [ExcelJavaScript API 参考](/javascript/api/excel)
