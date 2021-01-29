---
title: 使用 Excel JavaScript API 处理数据透视表
description: 使用 Excel JavaScript API 创建数据透视表并与其组件交互。
ms.date: 01/26/2021
localization_priority: Normal
ms.openlocfilehash: 9832322d40bbeb247685ff2498bdce42975c0377
ms.sourcegitcommit: 3123b9819c5225ee45a5312f64be79e46cbd0e3c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/29/2021
ms.locfileid: "50043909"
---
# <a name="work-with-pivottables-using-the-excel-javascript-api"></a>使用 Excel JavaScript API 处理数据透视表

数据透视表可简化较大的数据集。 它们允许快速操作分组数据。 Excel JavaScript API 允许加载项创建数据透视表并与其组件交互。 本文介绍数据透视表如何由 Office JavaScript API 表示，并提供关键方案的代码示例。

如果您不熟悉数据透视表的功能，请考虑以最终用户模式探索它们。
请参阅 ["创建数据透视表"以分析工作表数据](https://support.office.com/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) ，了解这些工具的良好基础。

> [!IMPORTANT]
> 当前不支持使用 OLAP 创建的数据透视表。 也不支持 Power Pivot。

## <a name="object-model"></a>对象模型

数据 [透视表](/javascript/api/excel/excel.pivottable) 是 Office JavaScript API 中数据透视表的中心对象。

- `Workbook.pivotTables`和分别包含工作簿和工作表中的数据透视表的数据透视 `Worksheet.pivotTables` 表[Collection。](/javascript/api/excel/excel.pivottablecollection) [](/javascript/api/excel/excel.pivottable)
- 数据[透视表](/javascript/api/excel/excel.pivottable)包含[具有多个 PivotHierarchies 的 PivotHierarchyCollection。](/javascript/api/excel/excel.pivothierarchycollection) [](/javascript/api/excel/excel.pivothierarchy)
- 可以将[这些 PivotHierarchies](/javascript/api/excel/excel.pivothierarchy)添加到特定层次结构集合中，以定义数据透视表 (数据透视表) 。 [](#hierarchies)
- [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy)包含只有一个透视字段的[PivotFieldCollection。](/javascript/api/excel/excel.pivotfieldcollection) [](/javascript/api/excel/excel.pivotfield) 如果设计扩展为包含 OLAP 数据透视表，这可能会更改。
- 只要[将字段](/javascript/api/excel/excel.pivotfield)的[PivotHierarchy](/javascript/api/excel/excel.pivothierarchy)分配给层次结构类别，透视字段就可以应用一个或多个[PivotFilter。](/javascript/api/excel/excel.pivotfilters) 
- 透视[字段](/javascript/api/excel/excel.pivotfield)包含具有多个[PivotItems 的 PivotItemCollection。](/javascript/api/excel/excel.pivotitemcollection) [](/javascript/api/excel/excel.pivotitem)
- 数据 [透视表](/javascript/api/excel/excel.pivottable) 包含 [一个 PivotLayout，](/javascript/api/excel/excel.pivotlayout) 用于定义 [数据透视字段](/javascript/api/excel/excel.pivotfield) 和 [PivotItems](/javascript/api/excel/excel.pivotitem) 在工作表中的显示位置。

让我们看一下这些关系如何应用于一些示例数据。 以下数据描述了各种服务器场中的水产品销售。 它将是本文中的示例。

![不同服务器场中不同类型的新鲜品销售的集合。](../images/excel-pivots-raw-data.png)

此服务器场销售数据将用于制作数据透视表。 每一列（如 **类型**）都是一个 `PivotHierarchy` 。 " **类型** "层次结构包含 **"类型"** 字段。 "**类型**"字段包含项 **Apple、Kiwi、Orange、Orange** 和 **Orange。**   

### <a name="hierarchies"></a>Hierarchies

数据透视表基于四个层次结构类别进行组织：[行](/javascript/api/excel/excel.rowcolumnpivothierarchy)、[列](/javascript/api/excel/excel.rowcolumnpivothierarchy)[、数据和](/javascript/api/excel/excel.datapivothierarchy)[筛选器](/javascript/api/excel/excel.filterpivothierarchy)。

前面显示的服务器场数据有五个层次结构：**场**、类型、**分类**、场中出售的箱和 **出售的箱**。 每个层次结构只能存在于四个类别之一中。 如果将 **Type** 添加到列层次结构中，则不能也位于行、数据或筛选器层次结构中。 如果 **Type** 随后添加到行层次结构中，则从列层次结构中删除类型。 无论层次结构分配是通过 Excel UI 还是 Excel JavaScript API 完成，此行为都是相同的。

行和列层次结构定义数据的分组方法。 例如，服务器场的 **行层次结构将** 同一服务器场的所有数据集组合在一起。 行层次结构和列层次结构之间的选择定义数据透视表的方向。

数据层次结构是基于行和列层次结构聚合的值。 具有服务器场的行层次结构和"出售的箱"的数据层次结构的数据透视表显示每个服务器场的所有不同 (中) 的总计值。

筛选器层次结构基于该筛选类型中的值包含或排除透视表中的数据。 已选择"有机"类型的 **分类** 的筛选器层次结构只显示有机水的数据。

下面是与数据透视表一起再次包含的服务器场数据。 数据透视表使用"服务器场"和"类型"作为行层次结构，将"在服务器场中出售"和"已出售的库存"作为数据层次结构 (，其默认聚合函数为 sum) ，将 **Classification** 用作筛选器层次结构 (，并选择了") "。 

![数据透视表旁边具有行、数据和筛选器层次结构的一系列新鲜品销售数据。](../images/excel-pivot-table-and-data.png)

此数据透视表可以通过 JavaScript API 或 Excel UI 生成。 这两个选项均允许通过加载项进一步操作。

## <a name="create-a-pivottable"></a>创建数据透视表

数据透视表需要名称、源和目标。 源可以是区域地址或表名称， (`Range` 作为 `string` ，传递，或 `Table` 键入) 。 目标地址是给定为 a 或 (`Range` 的范围 `string`) 。
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

### <a name="create-a-pivottable-with-range-objects"></a>使用 Range 对象创建数据透视表

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

## <a name="use-an-existing-pivottable"></a>使用现有的数据透视表

手动创建的数据透视表也可通过工作簿或单个工作表的数据透视表集合访问。 以下代码从工作簿获取名为 **My Pivot** 的数据透视表。

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.pivotTables.getItem("My Pivot");
    return context.sync();
});
```

## <a name="add-rows-and-columns-to-a-pivottable"></a>向数据透视表添加行和列

行和列围绕这些字段的值透视数据。

添加 **"服务器场** "列可透视每个服务器场的所有销售额。 添加 **"类型** " **和** "分类"行可进一步根据所出售的树和是否有机来分解数据。

![具有"服务器场"列和"类型和分类"行的数据透视表。](../images/excel-pivots-table-rows-and-columns.png)

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));

    pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem("Farm"));

    return context.sync();
});
```

还可以使数据透视表仅包含行或列。

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

数据层次结构使用基于行和列组合的信息填充数据透视表。 添加"在服务器场中出售的箱"和"出售的箱"**的数据** 层次结构会提供每行和每列的数据总和。

在示例中 **，Farm** 和 **Type** 都是行，以箱销售作为数据。

![一个数据透视表，显示基于它们所来自的服务器场的不同树的总销售额。](../images/excel-pivots-data-hierarchy.png)

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

下图显示了哪些布局函数调用对应于数据透视表的哪些范围。

![显示数据透视表的哪些部分由布局的获取范围函数返回的图表。](../images/excel-pivots-layout-breakdown.png)

### <a name="get-data-from-the-pivottable"></a>从数据透视表获取数据

布局定义数据透视表在工作表中的显示方式。 这意味着对象 `PivotLayout` 控制用于数据透视表元素的范围。 使用布局提供的范围获取由数据透视表收集和聚合的数据。 特别是，用于 `PivotLayout.getDataBodyRange` 访问数据透视表生成的结果。

下面的代码演示了如何获取数据透视表数据的最后一行，方法为浏览布局 (前面示例) 中"在服务器场中销售的箱总和"和"销售的"列的总计。  然后，这些值汇总在一起，最终总计显示在数据透视表外部的 **单元格 E30** (中) 。

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

数据透视表具有三种布局样式：精简、大纲和表格。 我们已经在之前的示例中看到过紧凑样式。

以下示例分别使用大纲样式和表格样式。 该代码示例演示如何在不同布局之间循环。

#### <a name="outline-layout"></a>大纲布局

![使用大纲布局的数据透视表。](../images/excel-pivots-outline-layout.png)

#### <a name="tabular-layout"></a>表格布局

![使用表格布局的数据透视表。](../images/excel-pivots-tabular-layout.png)

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

[PivotFilters](/javascript/api/excel/excel.pivotfilters) 根据数据透视表 [的四个](#hierarchies) 层次结构类别筛选数据， (筛选、列、行和) 。 有四种类型的 PivotFilter，允许基于日历日期的筛选、字符串分析、数字比较和基于自定义输入的筛选。 

[切片器](/javascript/api/excel/excel.slicer) 可以应用于数据透视表和常规 Excel 表。 应用于数据透视表时，切片器的功能与 [PivotManualFilter](#pivotmanualfilter) 类似，并允许基于自定义输入进行筛选。 与 PivotFilter 不同，切片器具有 [Excel UI 组件](https://support.office.com/article/Use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d)。 通过 `Slicer` 该类，你可以创建此 UI 组件、管理筛选并控制其视觉外观。 

### <a name="filter-with-pivotfilters"></a>使用 PivotFilter 筛选

[PivotFilters](/javascript/api/excel/excel.pivotfilters)允许您基于四个层次结构类别筛选数据透视表[](#hierarchies) (筛选器、列、行和) 。 在数据透视表对象模型中， `PivotFilters` 应用于 [数据透视字段](/javascript/api/excel/excel.pivotfield)，并且 `PivotField` 每个字段都可以分配一个或多个 `PivotFilters` 。 若要将 PivotFilters 应用于透视字段，字段对应的 [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) 必须分配给层次结构类别。 

#### <a name="types-of-pivotfilters"></a>PivotFilters 的类型

| 筛选器类型 | 筛选目的 | Excel JavaScript API 参考 |
|:--- |:--- |:--- |
| DateFilter | 基于日历日期的筛选。 | [PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter) |
| LabelFilter | 文本比较筛选。 | [PivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter) |
| ManualFilter | 自定义输入筛选。 | [PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter) |
| ValueFilter | 数字比较筛选。 | [PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter) |

#### <a name="create-a-pivotfilter"></a>创建 PivotFilter

若要筛选具有数据透视表 (`Pivot*Filter` 数据透视表) ，请 `PivotDateFilter` 对透视 [字段应用筛选器](/javascript/api/excel/excel.pivotfield)。 以下四个代码示例显示了如何使用这四种类型的 PivotFilter。 

##### <a name="pivotdatefilter"></a>PivotDateFilter

第一个代码示例将 [PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter) 应用于 **Date Updated** PivotField，隐藏 **2020-08-01 之前的任何数据**。 

> [!IMPORTANT] 
> A `Pivot*Filter` 不能应用于透视字段，除非该字段的 PivotHierarchy 分配给层次结构类别。 在下面的代码示例中，必须先将数据透视表添加到数据透视表的类别中，然后才能 `dateHierarchy` `rowHierarchies` 用于筛选。

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
> 以下三个代码段仅显示特定于筛选器的摘录，而不是完整 `Excel.run` 调用。

##### <a name="pivotlabelfilter"></a>PivotLabelFilter

第二个代码段演示如何将 [PivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter) 应用到 **Type** PivotField，使用该属性排除以字母 L 开头 `LabelFilterCondition.beginsWith` **的标签**。 

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

第三个代码段将 [具有 PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter) 的手动筛选器应用于 **Classification** 字段，筛选出不包含 **分类"组织"的数据**。 

```js
    // Apply a manual filter to include only a specific PivotItem (the string "Organic").
    var filterField = classHierarchy.fields.getItem("Classification");
    var manualFilter = { selectedItems: ["Organic"] };
    filterField.applyFilter({ manualFilter: manualFilter });
```

##### <a name="pivotvaluefilter"></a>PivotValueFilter

若要比较数字，请将值筛选器与 [PivotValueFilter 一](/javascript/api/excel/excel.pivotvaluefilter)同使用，如最终代码片段中所示。 该 `PivotValueFilter` 比较的服务器场透视字段的数据与 **Crates Sold PivotField，** 包括仅出售的箱总和超过 **值 500** 的服务器场的数据。 

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

#### <a name="remove-pivotfilters"></a>删除 PivotFilters

若要删除所有 PivotFilter，请对各个透视字段应用该方法， `clearAllFilters` 如下面的代码示例所示。 

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

[切片器](/javascript/api/excel/excel.slicer) 允许从 Excel 数据透视表或表筛选数据。 切片器使用指定列或透视字段的值筛选相应的行。 这些值存储为 [SlicerItem](/javascript/api/excel/excel.sliceritem) 对象 `Slicer` 。 加载项可以调整这些筛选器，用户也可以 ([Excel UI](https://support.office.com/article/Use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d)) 。 切片器位于绘图层中工作表的顶部，如以下屏幕截图所示。

![数据透视表上的切片器筛选数据。](../images/excel-slicer.png)

> [!NOTE]
> 本节中介绍的技术侧重于如何使用连接到数据透视表的切片器。 相同的技术也适用于使用连接到表的切片器。

#### <a name="create-a-slicer"></a>创建切片器

可以使用方法在工作簿或工作表中 `Workbook.slicers.add` 创建切片 `Worksheet.slicers.add` 器。 这样做会将切片器添加到指定或对象的[SlicerCollection。](/javascript/api/excel/excel.slicercollection) `Workbook` `Worksheet` 该方法 `SlicerCollection.add` 具有三个参数：

- `slicerSource`：新切片器所基于的数据源。 它可以是 `PivotTable` 一个 `Table` ，或字符串，表示的名称或 ID `PivotTable` 的或 `Table` 。
- `sourceField`：要筛选的数据源中的字段。 它可以是 `PivotField` 一个 `TableColumn` ，或字符串，表示的名称或 ID `PivotField` 的或 `TableColumn` 。
- `slicerDestination`：将在其中新建切片器的工作表。 它可以是 `Worksheet` 对象或名称或 `Worksheet` ID。 通过 访问时，不需要 `SlicerCollection` 此参数 `Worksheet.slicers` 。 在这种情况下，集合的工作表用作目标。

下面的代码示例向数据透视表添加新 **切片** 器。 切片器的来源是服务器场 **销售** 数据透视表，并且使用 **Type** 数据进行筛选。 该切片器也名为 **"切片器"，** 供将来参考。

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

切片器筛选数据透视表，并筛选数据透视表中的项 `sourceField` 。 `Slicer.selectItems`该方法设置保留在切片器中的项。 这些项目作为表示项的键传递给方法 `string[]` 。 包含这些项目的任何行都保留在数据透视表的聚合中。 后续调用 `selectItems` 以将列表设置为这些调用中指定的键。

> [!NOTE]
> 如果 `Slicer.selectItems` 传递的项不在数据源中，则 `InvalidArgument` 会引发错误。 可通过属性（即 `Slicer.slicerItems` [SlicerItemCollection）验证内容](/javascript/api/excel/excel.sliceritemcollection)。

下面的代码示例显示了为切片器选择的三个项目：**青** 绿色、酸 **橙色****和橙色**。

```js
Excel.run(function (context) {
    var slicer = context.workbook.slicers.getItem("Fruit Slicer");
    // Anything other than the following three values will be filtered out of the PivotTable for display and aggregation.
    slicer.selectItems(["Lemon", "Lime", "Orange"]);
    return context.sync();
});
```

若要从切片器中删除所有筛选器，请使用该方法 `Slicer.clearFilters` ，如以下示例所示。

```js
Excel.run(function (context) {
    var slicer = context.workbook.slicers.getItem("Fruit Slicer");
    slicer.clearFilters();
    return context.sync();
});
```

#### <a name="style-and-format-a-slicer"></a>设置切片器样式和格式

加载项可以通过属性调整切片器显示 `Slicer` 设置。 下面的代码示例将样式设置为 **SlicerStyleLight6，** 将切片器顶部的文本设置为 **"树** 类型"，将切片器放在绘图层上位置 **(395，15) ，** 将切片器的大小设置为 **135x150** 像素。

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

若要删除切片器，请调用 `Slicer.delete` 该方法。 下面的代码示例从当前工作表中删除第一个切片器。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.slicers.getItemAt(0).delete();
    return context.sync();
});
```

## <a name="change-aggregation-function"></a>更改聚合函数

数据层次结构聚合了它们的值。 对于数字数据集，默认情况下这是一个和。 该属性 `summarizeBy` 根据 [AggregationFunction](/javascript/api/excel/excel.aggregationfunction) 类型定义此行为。

当前支持的聚合函数类型是 `Sum` ， 和 (`Count` `Average` `Max` `Min` `Product` `CountNumbers` `StandardDeviation` `StandardDeviationP` `Variance` `VarianceP` `Automatic` 默认) 。

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

该对象 `ShowAsRule` 具有三个属性：

- `calculation`：应用于数据层次结构的相对计算类型 (默认值为 `none`) 。
- `baseField`： [在应用](/javascript/api/excel/excel.pivotfield) 计算之前，层次结构中包含基数据的数据透视字段。 由于 Excel 数据透视表具有层次结构到字段的一对一映射，因此您将使用相同的名称访问层次结构和字段。
- `baseItem`：单个 [PivotItem](/javascript/api/excel/excel.pivotitem) 与基于计算类型的基字段的值进行比较。 并非所有计算都需要此字段。

以下示例将服务器场数据层次结构中销售的箱总和的计算结果设定为列总计的百分比。
我们仍希望粒度扩展到结果类型级别，因此我们将使用 **"** 类型"行层次结构及其基础字段。
此示例还将 **Farm** 作为第一行层次结构，因此服务器场总条目也显示每个服务器场负责生产的百分比。

![一个数据透视表，显示每个服务器场中每个服务器场和各个新鲜菜类型相对于总销售额的新鲜情况百分比。](../images/excel-pivots-showas-percentage.png)

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

前面的示例相对于单个行层次结构的字段将计算设置为列。 当计算与单个项目相关时，请使用 `baseItem` 该属性。

以下示例显示 `differenceFrom` 计算。 它显示服务器场包含销售数据层次结构条目与服务器场的条目 **的差**。
它是服务器场，因此我们将看到其他服务器场之间的差异，以及每种类型的类似树的细目 (Type 也是此示例中的行层次结构 `baseField`) 。  

![一个数据透视表，显示"A Farms"和其他服务器场之间的新鲜品销售差异。 这显示了服务器场的总新鲜销售额和新鲜菜类型的销售差异。 如果"A Farms"未出售特定类型的#N，则会显示"#N/A"。](../images/excel-pivots-showas-differencefrom.png)

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
- [Excel JavaScript API 参考](/javascript/api/excel)
