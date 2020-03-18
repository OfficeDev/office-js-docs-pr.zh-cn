---
title: 使用 Excel JavaScript API 处理数据透视表
description: 使用 Excel JavaScript API 创建数据透视表并与其组件进行交互。
ms.date: 01/22/2020
localization_priority: Normal
ms.openlocfilehash: ec7d7ccd7f040185e31b59693827c31d5dab8372
ms.sourcegitcommit: a0262ea40cd23f221e69bcb0223110f011265d13
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42688572"
---
# <a name="work-with-pivottables-using-the-excel-javascript-api"></a><span data-ttu-id="8884f-103">使用 Excel JavaScript API 处理数据透视表</span><span class="sxs-lookup"><span data-stu-id="8884f-103">Work with PivotTables using the Excel JavaScript API</span></span>

<span data-ttu-id="8884f-104">数据透视表精简了较大的数据集。</span><span class="sxs-lookup"><span data-stu-id="8884f-104">PivotTables streamline larger data sets.</span></span> <span data-ttu-id="8884f-105">它们允许快速操作分组数据。</span><span class="sxs-lookup"><span data-stu-id="8884f-105">They allow the quick manipulation of grouped data.</span></span> <span data-ttu-id="8884f-106">Excel JavaScript API 允许你的外接程序创建数据透视表并与其组件进行交互。</span><span class="sxs-lookup"><span data-stu-id="8884f-106">The Excel JavaScript API lets your add-in create PivotTables and interact with their components.</span></span> <span data-ttu-id="8884f-107">本文介绍了 Office JavaScript API 如何表示数据透视表，并提供了主要方案的代码示例。</span><span class="sxs-lookup"><span data-stu-id="8884f-107">This article describes how PivotTables are represented by the Office JavaScript API and provides code samples for key scenarios.</span></span>

<span data-ttu-id="8884f-108">如果您对数据透视表的功能不熟悉，请考虑将其作为最终用户来浏览。</span><span class="sxs-lookup"><span data-stu-id="8884f-108">If you are unfamiliar with the functionality of PivotTables, consider exploring them as an end user.</span></span>
<span data-ttu-id="8884f-109">有关这些工具的最佳入门知识，请参阅[创建数据透视表以分析工作表数据](https://support.office.com/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables)。</span><span class="sxs-lookup"><span data-stu-id="8884f-109">See [Create a PivotTable to analyze worksheet data](https://support.office.com/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) for a good primer on these tools.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="8884f-110">目前不支持使用 OLAP 创建的数据透视表。</span><span class="sxs-lookup"><span data-stu-id="8884f-110">PivotTables created with OLAP are not currently supported.</span></span> <span data-ttu-id="8884f-111">此外，也不支持 Power Pivot。</span><span class="sxs-lookup"><span data-stu-id="8884f-111">There is also no support for Power Pivot.</span></span>

## <a name="object-model"></a><span data-ttu-id="8884f-112">对象模型</span><span class="sxs-lookup"><span data-stu-id="8884f-112">Object model</span></span>

<span data-ttu-id="8884f-113">[数据透视表](/javascript/api/excel/excel.pivottable)是 OFFICE JavaScript API 中数据透视表的中心对象。</span><span class="sxs-lookup"><span data-stu-id="8884f-113">The [PivotTable](/javascript/api/excel/excel.pivottable) is the central object for PivotTables in the Office JavaScript API.</span></span>

- <span data-ttu-id="8884f-114">`Workbook.pivotTables`并且`Worksheet.pivotTables`是分别在工作簿和工作表中包含[数据透视](/javascript/api/excel/excel.pivottable)表的[PivotTableCollections](/javascript/api/excel/excel.pivottablecollection) 。</span><span class="sxs-lookup"><span data-stu-id="8884f-114">`Workbook.pivotTables` and `Worksheet.pivotTables` are [PivotTableCollections](/javascript/api/excel/excel.pivottablecollection) that contain the [PivotTables](/javascript/api/excel/excel.pivottable) in the workbook and worksheet, respectively.</span></span>
- <span data-ttu-id="8884f-115">[数据透视表](/javascript/api/excel/excel.pivottable)包含具有多个[PivotHierarchies](/javascript/api/excel/excel.pivothierarchy)的[PivotTableCollections](/javascript/api/excel/excel.pivottablecollection) 。</span><span class="sxs-lookup"><span data-stu-id="8884f-115">A [PivotTable](/javascript/api/excel/excel.pivottable) contains a [PivotTableCollections](/javascript/api/excel/excel.pivottablecollection) that has multiple [PivotHierarchies](/javascript/api/excel/excel.pivothierarchy).</span></span>
- <span data-ttu-id="8884f-116">[PivotHierarchy](/javascript/api/excel/excel.pivothierarchy)包含一个仅具有一个[透视字段](/javascript/api/excel/excel.pivotfield)的[PivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection) 。</span><span class="sxs-lookup"><span data-stu-id="8884f-116">A [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) contains a [PivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection) that has exactly one [PivotField](/javascript/api/excel/excel.pivotfield).</span></span> <span data-ttu-id="8884f-117">如果设计扩展以包含 OLAP 数据透视表，则可能会发生更改。</span><span class="sxs-lookup"><span data-stu-id="8884f-117">If the design expands to include OLAP PivotTables, this may change.</span></span>
- <span data-ttu-id="8884f-118">[透视字段](/javascript/api/excel/excel.pivotfield)包含具有多个[PivotItems](/javascript/api/excel/excel.pivotitem)的[PivotItemCollection](/javascript/api/excel/excel.pivotitemcollection) 。</span><span class="sxs-lookup"><span data-stu-id="8884f-118">A [PivotField](/javascript/api/excel/excel.pivotfield) contains a [PivotItemCollection](/javascript/api/excel/excel.pivotitemcollection) that has multiple [PivotItems](/javascript/api/excel/excel.pivotitem).</span></span>
- <span data-ttu-id="8884f-119">[数据透视表](/javascript/api/excel/excel.pivottable)包含一个[PivotLayout](/javascript/api/excel/excel.pivotlayout) ，用于定义在工作表中显示[透视字段](/javascript/api/excel/excel.pivotfield)和[PivotItems](/javascript/api/excel/excel.pivotitem)的位置。</span><span class="sxs-lookup"><span data-stu-id="8884f-119">A [PivotTable](/javascript/api/excel/excel.pivottable) contains a [PivotLayout](/javascript/api/excel/excel.pivotlayout) that defines where the [PivotFields](/javascript/api/excel/excel.pivotfield) and [PivotItems](/javascript/api/excel/excel.pivotitem) are displayed in the worksheet.</span></span>

<span data-ttu-id="8884f-120">让我们来看看这些关系如何应用于一些示例数据。</span><span class="sxs-lookup"><span data-stu-id="8884f-120">Let's look at how these relationships apply to some example data.</span></span> <span data-ttu-id="8884f-121">以下数据介绍了来自不同服务器场的水果销售。</span><span class="sxs-lookup"><span data-stu-id="8884f-121">The following data describes fruit sales from various farms.</span></span> <span data-ttu-id="8884f-122">本主题将是本文中的示例。</span><span class="sxs-lookup"><span data-stu-id="8884f-122">It will be the example throughout this article.</span></span>

![来自不同服务器场的不同类型的水果销售的集合。](../images/excel-pivots-raw-data.png)

<span data-ttu-id="8884f-124">此水果可使用的销售数据将用于制作数据透视表。</span><span class="sxs-lookup"><span data-stu-id="8884f-124">This fruit farm sales data will be used to make a PivotTable.</span></span> <span data-ttu-id="8884f-125">每个列（例如 "**类型**"） `PivotHierarchy`是。</span><span class="sxs-lookup"><span data-stu-id="8884f-125">Each column, such as **Types**, is a `PivotHierarchy`.</span></span> <span data-ttu-id="8884f-126">"**类型**" 层次结构包含 "**类型**" 字段。</span><span class="sxs-lookup"><span data-stu-id="8884f-126">The **Types** hierarchy contains the **Types** field.</span></span> <span data-ttu-id="8884f-127">"**类型**" 字段包含**苹果**、 **Kiwi**、**柠檬**、**酸**橙色和**橙色**的项。</span><span class="sxs-lookup"><span data-stu-id="8884f-127">The **Types** field contains the items **Apple**, **Kiwi**, **Lemon**, **Lime**, and **Orange**.</span></span>

### <a name="hierarchies"></a><span data-ttu-id="8884f-128">Hierarchies</span><span class="sxs-lookup"><span data-stu-id="8884f-128">Hierarchies</span></span>

<span data-ttu-id="8884f-129">数据透视表基于四种层次结构类别进行组织：[行](/javascript/api/excel/excel.rowcolumnpivothierarchy)、[列](/javascript/api/excel/excel.rowcolumnpivothierarchy)、[数据](/javascript/api/excel/excel.datapivothierarchy)和[筛选器](/javascript/api/excel/excel.filterpivothierarchy)。</span><span class="sxs-lookup"><span data-stu-id="8884f-129">PivotTables are organized based on four hierarchy categories: [row](/javascript/api/excel/excel.rowcolumnpivothierarchy), [column](/javascript/api/excel/excel.rowcolumnpivothierarchy), [data](/javascript/api/excel/excel.datapivothierarchy), and [filter](/javascript/api/excel/excel.filterpivothierarchy).</span></span>

<span data-ttu-id="8884f-130">前面显示的服务器场数据具有五个层次**结构：服务器场、\*\*\*\*类型**、**分类**、**服务器场中销售的 Crates**和**Crates 销售批发**。</span><span class="sxs-lookup"><span data-stu-id="8884f-130">The farm data shown earlier has five hierarchies: **Farms**, **Type**, **Classification**, **Crates Sold at Farm**, and **Crates Sold Wholesale**.</span></span> <span data-ttu-id="8884f-131">每个层次结构只能存在于四个类别之一中。</span><span class="sxs-lookup"><span data-stu-id="8884f-131">Each hierarchy can only exist in one of the four categories.</span></span> <span data-ttu-id="8884f-132">如果**Type**添加到列层次结构，则它不能也在行、数据或筛选器层次结构中。</span><span class="sxs-lookup"><span data-stu-id="8884f-132">If **Type** is added to column hierarchies, it cannot also be in the row, data, or filter hierarchies.</span></span> <span data-ttu-id="8884f-133">如果**类型**随后添加到行层次结构中，则会将其从列层次结构中删除。</span><span class="sxs-lookup"><span data-stu-id="8884f-133">If **Type** is subsequently added to row hierarchies, it is removed from the column hierarchies.</span></span> <span data-ttu-id="8884f-134">无论是通过 Excel UI 还是 Excel JavaScript Api 执行层次结构分配，此行为都相同。</span><span class="sxs-lookup"><span data-stu-id="8884f-134">This behavior is the same whether hierarchy assignment is done through the Excel UI or the Excel JavaScript APIs.</span></span>

<span data-ttu-id="8884f-135">行和列层次结构定义数据的分组方式。</span><span class="sxs-lookup"><span data-stu-id="8884f-135">Row and column hierarchies define how data will be grouped.</span></span> <span data-ttu-id="8884f-136">例如，**服务器场**的行层次结构将把来自同一个服务器场的所有数据集组合在一起。</span><span class="sxs-lookup"><span data-stu-id="8884f-136">For example, a row hierarchy of **Farms** will group together all the data sets from the same farm.</span></span> <span data-ttu-id="8884f-137">在行和列层次结构之间进行选择，以定义数据透视表的方向。</span><span class="sxs-lookup"><span data-stu-id="8884f-137">The choice between row and column hierarchy defines the orientation of the PivotTable.</span></span>

<span data-ttu-id="8884f-138">数据层次结构是要根据行和列层次结构聚合的值。</span><span class="sxs-lookup"><span data-stu-id="8884f-138">Data hierarchies are the values to be aggregated based on the row and column hierarchies.</span></span> <span data-ttu-id="8884f-139">具有**服务器场**的行层次结构和**Crates 销售**的数据层次结构的数据透视表显示每个服务器场的所有不同 fruits 的总和总计（默认值）。</span><span class="sxs-lookup"><span data-stu-id="8884f-139">A PivotTable with a row hierarchy of **Farms** and a data hierarchy of **Crates Sold Wholesale** shows the sum total (by default) of all the different fruits for each farm.</span></span>

<span data-ttu-id="8884f-140">筛选器层次结构基于该筛选类型中的值包括或排除数据透视表中的数据。</span><span class="sxs-lookup"><span data-stu-id="8884f-140">Filter hierarchies include or exclude data from the pivot based on values within that filtered type.</span></span> <span data-ttu-id="8884f-141">选定类型为 "**有机**" 的**分类**筛选器层次结构仅显示用于随机水果的数据。</span><span class="sxs-lookup"><span data-stu-id="8884f-141">A filter hierarchy of **Classification** with the type **Organic** selected only shows data for organic fruit.</span></span>

<span data-ttu-id="8884f-142">下面是数据透视表旁边的服务器场数据。</span><span class="sxs-lookup"><span data-stu-id="8884f-142">Here is the farm data again, alongside a PivotTable.</span></span> <span data-ttu-id="8884f-143">数据透视表使用**服务器场**和**类型**作为行层次结构， **Crates 在服务器场**和**Crates 销售批发**作为数据层次结构（具有 sum 的默认聚合函数）和**分类**作为筛选器层次结构（带有**随机**选择的）。</span><span class="sxs-lookup"><span data-stu-id="8884f-143">The PivotTable is using **Farm** and **Type** as the row hierarchies, **Crates Sold at Farm** and **Crates Sold Wholesale** as the data hierarchies (with the default aggregation function of sum), and **Classification** as a filter hierarchy (with **Organic** selected).</span></span>

![选择了具有行、数据和筛选器层次结构的数据透视表旁边的水果销售数据。](../images/excel-pivot-table-and-data.png)

<span data-ttu-id="8884f-145">此数据透视表可通过 JavaScript API 或 Excel UI 生成。</span><span class="sxs-lookup"><span data-stu-id="8884f-145">This PivotTable could be generated through the JavaScript API or through the Excel UI.</span></span> <span data-ttu-id="8884f-146">这两个选项都允许通过外接程序进行进一步操作。</span><span class="sxs-lookup"><span data-stu-id="8884f-146">Both options allow for further manipulation through add-ins.</span></span>

## <a name="create-a-pivottable"></a><span data-ttu-id="8884f-147">创建数据透视表</span><span class="sxs-lookup"><span data-stu-id="8884f-147">Create a PivotTable</span></span>

<span data-ttu-id="8884f-148">数据透视表需要名称、源和目标。</span><span class="sxs-lookup"><span data-stu-id="8884f-148">PivotTables need a name, source, and destination.</span></span> <span data-ttu-id="8884f-149">源可以是区域地址或表名称（作为`Range`、 `string`或`Table`类型传递）。</span><span class="sxs-lookup"><span data-stu-id="8884f-149">The source can be a range address or table name (passed as a `Range`, `string`, or `Table` type).</span></span> <span data-ttu-id="8884f-150">目标是区域地址（指定为`Range`或`string`）。</span><span class="sxs-lookup"><span data-stu-id="8884f-150">The destination is a range address (given as either a `Range` or `string`).</span></span>
<span data-ttu-id="8884f-151">下面的示例展示了各种数据透视表创建技术。</span><span class="sxs-lookup"><span data-stu-id="8884f-151">The following samples show various PivotTable creation techniques.</span></span>

### <a name="create-a-pivottable-with-range-addresses"></a><span data-ttu-id="8884f-152">创建包含区域地址的数据透视表</span><span class="sxs-lookup"><span data-stu-id="8884f-152">Create a PivotTable with range addresses</span></span>

```js
Excel.run(function (context) {
    // Create a PivotTable named "Farm Sales" on the current worksheet at cell
    // A22 with data from the range A1:E21.
    context.workbook.worksheets.getActiveWorksheet().pivotTables.add(
      "Farm Sales", "A1:E21", "A22");

    return context.sync();
});
```

### <a name="create-a-pivottable-with-range-objects"></a><span data-ttu-id="8884f-153">创建包含 Range 对象的数据透视表</span><span class="sxs-lookup"><span data-stu-id="8884f-153">Create a PivotTable with Range objects</span></span>

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

### <a name="create-a-pivottable-at-the-workbook-level"></a><span data-ttu-id="8884f-154">在工作簿级别创建数据透视表</span><span class="sxs-lookup"><span data-stu-id="8884f-154">Create a PivotTable at the workbook level</span></span>

```js
Excel.run(function (context) {
    // Create a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data is from the worksheet "DataWorksheet" across the range A1:E21.
    context.workbook.pivotTables.add(
        "Farm Sales", "DataWorksheet!A1:E21", "PivotWorksheet!A2");

    return context.sync();
});
```

## <a name="use-an-existing-pivottable"></a><span data-ttu-id="8884f-155">使用现有数据透视表</span><span class="sxs-lookup"><span data-stu-id="8884f-155">Use an existing PivotTable</span></span>

<span data-ttu-id="8884f-156">手动创建的数据透视表也可通过工作簿或单个工作表的数据透视表集合进行访问。</span><span class="sxs-lookup"><span data-stu-id="8884f-156">Manually created PivotTables are also accessible through the PivotTable collection of the workbook or of individual worksheets.</span></span> <span data-ttu-id="8884f-157">下面的代码从工作簿中获取名为 **"我的透视**表" 的数据透视表。</span><span class="sxs-lookup"><span data-stu-id="8884f-157">The following code gets a PivotTable named **My Pivot** from the workbook.</span></span>

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.pivotTables.getItem("My Pivot");
    return context.sync();
});
```

## <a name="add-rows-and-columns-to-a-pivottable"></a><span data-ttu-id="8884f-158">向数据透视表添加行和列</span><span class="sxs-lookup"><span data-stu-id="8884f-158">Add rows and columns to a PivotTable</span></span>

<span data-ttu-id="8884f-159">数据透视这些字段的值周围的行和列。</span><span class="sxs-lookup"><span data-stu-id="8884f-159">Rows and columns pivot the data around those fields’ values.</span></span>

<span data-ttu-id="8884f-160">添加 "**服务器场**" 列将每个服务器场的所有销售额枢轴分布。</span><span class="sxs-lookup"><span data-stu-id="8884f-160">Adding the **Farm** column pivots all the sales around each farm.</span></span> <span data-ttu-id="8884f-161">添加 "**类型**" 和 "**分类**" 行会根据所售的水果和是否为 "有随机" 来进一步细分数据。</span><span class="sxs-lookup"><span data-stu-id="8884f-161">Adding the **Type** and **Classification** rows further breaks down the data based on what fruit was sold and whether it was organic or not.</span></span>

![具有服务器场列和类型和分类行的数据透视表。](../images/excel-pivots-table-rows-and-columns.png)

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));

    pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem("Farm"));

    return context.sync();
});
```

<span data-ttu-id="8884f-163">您还可以拥有仅包含行或列的数据透视表。</span><span class="sxs-lookup"><span data-stu-id="8884f-163">You can also have a PivotTable with only rows or columns.</span></span>

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));

    return context.sync();
});
```

## <a name="add-data-hierarchies-to-the-pivottable"></a><span data-ttu-id="8884f-164">向数据透视表中添加数据层次结构</span><span class="sxs-lookup"><span data-stu-id="8884f-164">Add data hierarchies to the PivotTable</span></span>

<span data-ttu-id="8884f-165">数据层次结构使用要基于行和列进行组合的信息填充数据透视表。</span><span class="sxs-lookup"><span data-stu-id="8884f-165">Data hierarchies fill the PivotTable with information to combine based on the rows and columns.</span></span> <span data-ttu-id="8884f-166">添加 Crates 的数据层次结构**在服务器场**和**Crates**售出销售批发为每个行和列提供这些数字的总和。</span><span class="sxs-lookup"><span data-stu-id="8884f-166">Adding the data hierarchies of **Crates Sold at Farm** and **Crates Sold Wholesale** gives sums of those figures for each row and column.</span></span>

<span data-ttu-id="8884f-167">在示例中，"**服务器场**" 和 "**类型**" 都是行，而 "发货箱销售额" 作为数据。</span><span class="sxs-lookup"><span data-stu-id="8884f-167">In the example, both **Farm** and **Type** are rows, with the crate sales as the data.</span></span>

![显示基于其来源的服务器场的不同水果的总销售额的数据透视表。](../images/excel-pivots-data-hierarchy.png)

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

## <a name="slicers"></a><span data-ttu-id="8884f-169">切片器</span><span class="sxs-lookup"><span data-stu-id="8884f-169">Slicers</span></span>

<span data-ttu-id="8884f-170">[切片](/javascript/api/excel/excel.slicer)器允许从 Excel 数据透视表或表中筛选数据。</span><span class="sxs-lookup"><span data-stu-id="8884f-170">[Slicers](/javascript/api/excel/excel.slicer) allow data to be filtered from an Excel PivotTable or table.</span></span> <span data-ttu-id="8884f-171">切片器使用指定的列或透视字段中的值筛选相应的行。</span><span class="sxs-lookup"><span data-stu-id="8884f-171">A slicer uses values from a specified column or PivotField to filter corresponding rows.</span></span> <span data-ttu-id="8884f-172">这些值存储为中[SlicerItem](/javascript/api/excel/excel.sliceritem)的`Slicer`SlicerItem 对象。</span><span class="sxs-lookup"><span data-stu-id="8884f-172">These values are stored as [SlicerItem](/javascript/api/excel/excel.sliceritem) objects in the `Slicer`.</span></span> <span data-ttu-id="8884f-173">你的外接程序可以按照用户（[通过 EXCEL UI](https://support.office.com/article/Use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d)）的方式调整这些筛选器。</span><span class="sxs-lookup"><span data-stu-id="8884f-173">Your add-in can adjust these filters, as can users ([through the Excel UI](https://support.office.com/article/Use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d)).</span></span> <span data-ttu-id="8884f-174">切片器位于绘图层中的工作表的顶部，如以下屏幕截图所示。</span><span class="sxs-lookup"><span data-stu-id="8884f-174">The slicer sits on top of the worksheet in the drawing layer, as shown in the following screenshot.</span></span>

![数据透视表上的切片器筛选数据。](../images/excel-slicer.png)

> [!NOTE]
> <span data-ttu-id="8884f-176">本节中介绍的技术重点介绍如何使用连接到数据透视表的切片器。</span><span class="sxs-lookup"><span data-stu-id="8884f-176">The techniques described in this section focus on how to use slicers connected to PivotTables.</span></span> <span data-ttu-id="8884f-177">同样的技术也适用于使用连接到表的切片器。</span><span class="sxs-lookup"><span data-stu-id="8884f-177">The same techniques also apply to using slicers connected to tables.</span></span>

### <a name="create-a-slicer"></a><span data-ttu-id="8884f-178">创建切片器</span><span class="sxs-lookup"><span data-stu-id="8884f-178">Create a slicer</span></span>

<span data-ttu-id="8884f-179">您可以使用`Workbook.slicers.add`方法或`Worksheet.slicers.add`方法在工作簿或工作表中创建切片器。</span><span class="sxs-lookup"><span data-stu-id="8884f-179">You can create a slicer in a workbook or worksheet by using the `Workbook.slicers.add` method or `Worksheet.slicers.add` method.</span></span> <span data-ttu-id="8884f-180">这样做会将切片器添加[SlicerCollection](/javascript/api/excel/excel.slicercollection)到指定`Workbook`对象或`Worksheet`对象的 SlicerCollection。</span><span class="sxs-lookup"><span data-stu-id="8884f-180">Doing so adds a slicer to the [SlicerCollection](/javascript/api/excel/excel.slicercollection) of the specified `Workbook` or `Worksheet` object.</span></span> <span data-ttu-id="8884f-181">该`SlicerCollection.add`方法具有三个参数：</span><span class="sxs-lookup"><span data-stu-id="8884f-181">The `SlicerCollection.add` method has three parameters:</span></span>

- <span data-ttu-id="8884f-182">`slicerSource`：新切片器所基于的数据源。</span><span class="sxs-lookup"><span data-stu-id="8884f-182">`slicerSource`: The data source on which the new slicer is based.</span></span> <span data-ttu-id="8884f-183">它可以`PivotTable`是、或`Table`字符串，代表`PivotTable`或`Table`的名称或 ID。</span><span class="sxs-lookup"><span data-stu-id="8884f-183">It can be a `PivotTable`, `Table`, or string representing the name or ID of a `PivotTable` or `Table`.</span></span>
- <span data-ttu-id="8884f-184">`sourceField`：要作为筛选依据的数据源中的字段。</span><span class="sxs-lookup"><span data-stu-id="8884f-184">`sourceField`: The field in the data source by which to filter.</span></span> <span data-ttu-id="8884f-185">它可以`PivotField`是、或`TableColumn`字符串，代表`PivotField`或`TableColumn`的名称或 ID。</span><span class="sxs-lookup"><span data-stu-id="8884f-185">It can be a `PivotField`, `TableColumn`, or string representing the name or ID of a `PivotField` or `TableColumn`.</span></span>
- <span data-ttu-id="8884f-186">`slicerDestination`：将在其中创建新切片器的工作表。</span><span class="sxs-lookup"><span data-stu-id="8884f-186">`slicerDestination`: The worksheet where the new slicer will be created.</span></span> <span data-ttu-id="8884f-187">它可以是一个`Worksheet`对象，也可以是的名称或`Worksheet`ID。</span><span class="sxs-lookup"><span data-stu-id="8884f-187">It can be a `Worksheet` object or the name or ID of a `Worksheet`.</span></span> <span data-ttu-id="8884f-188">通过`Worksheet.slicers`访问时，此参数`SlicerCollection`是不必要的。</span><span class="sxs-lookup"><span data-stu-id="8884f-188">This parameter is unnecessary when the `SlicerCollection` is accessed through `Worksheet.slicers`.</span></span> <span data-ttu-id="8884f-189">在这种情况下，集合的工作表将用作目标。</span><span class="sxs-lookup"><span data-stu-id="8884f-189">In this case, the collection's worksheet is used as the destination.</span></span>

<span data-ttu-id="8884f-190">下面的代码示例向**数据透视**表中添加一个新的切片器。</span><span class="sxs-lookup"><span data-stu-id="8884f-190">The following code sample adds a new slicer to the **Pivot** worksheet.</span></span> <span data-ttu-id="8884f-191">切片器的源是**服务器场销售**数据透视表和使用**类型**数据的筛选器。</span><span class="sxs-lookup"><span data-stu-id="8884f-191">The slicer's source is the **Farm Sales** PivotTable and filters using the **Type** data.</span></span> <span data-ttu-id="8884f-192">切片器也称为**水果切片器**，以供将来参考。</span><span class="sxs-lookup"><span data-stu-id="8884f-192">The slicer is also named **Fruit Slicer** for future reference.</span></span>

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

### <a name="filter-items-with-a-slicer"></a><span data-ttu-id="8884f-193">使用切片器筛选项目</span><span class="sxs-lookup"><span data-stu-id="8884f-193">Filter items with a slicer</span></span>

<span data-ttu-id="8884f-194">切片器使用中的项筛选数据透视`sourceField`表。</span><span class="sxs-lookup"><span data-stu-id="8884f-194">The slicer filters the PivotTable with items from the `sourceField`.</span></span> <span data-ttu-id="8884f-195">该`Slicer.selectItems`方法将设置切片器中保留的项。</span><span class="sxs-lookup"><span data-stu-id="8884f-195">The `Slicer.selectItems` method sets the items that remain in the slicer.</span></span> <span data-ttu-id="8884f-196">这些项作为 a `string[]`传递给方法，表示项的键。</span><span class="sxs-lookup"><span data-stu-id="8884f-196">These items are passed to the method as a `string[]`, representing the keys of the items.</span></span> <span data-ttu-id="8884f-197">包含这些项目的任何行仍保留在数据透视表的聚合中。</span><span class="sxs-lookup"><span data-stu-id="8884f-197">Any rows containing those items remain in the PivotTable's aggregation.</span></span> <span data-ttu-id="8884f-198">随后调用`selectItems`将列表设置为在这些调用中指定的键。</span><span class="sxs-lookup"><span data-stu-id="8884f-198">Subsequent calls to `selectItems` set the list to the keys specified in those calls.</span></span>

> [!NOTE]
> <span data-ttu-id="8884f-199">如果`Slicer.selectItems`向传递的项不在数据源中，则会引发`InvalidArgument`错误。</span><span class="sxs-lookup"><span data-stu-id="8884f-199">If `Slicer.selectItems` is passed an item that's not in the data source, an `InvalidArgument` error is thrown.</span></span> <span data-ttu-id="8884f-200">可以通过`Slicer.slicerItems`属性来验证内容，这是[SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection)。</span><span class="sxs-lookup"><span data-stu-id="8884f-200">The contents can be verified through the `Slicer.slicerItems` property, which is a [SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection).</span></span>

<span data-ttu-id="8884f-201">下面的代码示例显示为切片器选择了三个项目：**柠檬**、**酸橙色**和**橙色**。</span><span class="sxs-lookup"><span data-stu-id="8884f-201">The following code sample shows three items being selected for the slicer: **Lemon**, **Lime**, and **Orange**.</span></span>

```js
Excel.run(function (context) {
    var slicer = context.workbook.slicers.getItem("Fruit Slicer");
    // Anything other than the following three values will be filtered out of the PivotTable for display and aggregation.
    slicer.selectItems(["Lemon", "Lime", "Orange"]);
    return context.sync();
});
```

<span data-ttu-id="8884f-202">若要从切片器中删除所有筛选器`Slicer.clearFilters` ，请使用方法，如下面的示例所示。</span><span class="sxs-lookup"><span data-stu-id="8884f-202">To remove all filters from the slicer, use the `Slicer.clearFilters` method, as shown in the following sample.</span></span>

```js
Excel.run(function (context) {
    var slicer = context.workbook.slicers.getItem("Fruit Slicer");
    slicer.clearFilters();
    return context.sync();
});
```

### <a name="style-and-format-a-slicer"></a><span data-ttu-id="8884f-203">为切片器设置样式和格式</span><span class="sxs-lookup"><span data-stu-id="8884f-203">Style and format a slicer</span></span>

<span data-ttu-id="8884f-204">您的外接可以通过`Slicer`属性调整切片器的显示设置。</span><span class="sxs-lookup"><span data-stu-id="8884f-204">You add-in can adjust a slicer's display settings through `Slicer` properties.</span></span> <span data-ttu-id="8884f-205">下面的代码示例将样式设置为**SlicerStyleLight6**，将切片器顶部的文本设置为**水果类型**，将切片器放置在绘图层上的位置 **（395，15）** ，并将切片器的大小设置为**135x150**像素。</span><span class="sxs-lookup"><span data-stu-id="8884f-205">The following code sample sets the style to **SlicerStyleLight6**, sets the text at the top of the slicer to **Fruit Types**, places the slicer at the position **(395, 15)** on the drawing layer, and sets the slicer's size to **135x150** pixels.</span></span>

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

### <a name="delete-a-slicer"></a><span data-ttu-id="8884f-206">删除切片器</span><span class="sxs-lookup"><span data-stu-id="8884f-206">Delete a slicer</span></span>

<span data-ttu-id="8884f-207">若要删除切片器，请`Slicer.delete`调用方法。</span><span class="sxs-lookup"><span data-stu-id="8884f-207">To delete a slicer, call the `Slicer.delete` method.</span></span> <span data-ttu-id="8884f-208">下面的代码示例从当前工作表中删除第一个切片器。</span><span class="sxs-lookup"><span data-stu-id="8884f-208">The following code sample deletes the first slicer from the current worksheet.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.slicers.getItemAt(0).delete();
    return context.sync();
});
```

## <a name="change-aggregation-function"></a><span data-ttu-id="8884f-209">更改聚合函数</span><span class="sxs-lookup"><span data-stu-id="8884f-209">Change aggregation function</span></span>

<span data-ttu-id="8884f-210">数据层次结构的值已聚合。</span><span class="sxs-lookup"><span data-stu-id="8884f-210">Data hierarchies have their values aggregated.</span></span> <span data-ttu-id="8884f-211">对于数字的数据集，默认情况下，这是一个总和。</span><span class="sxs-lookup"><span data-stu-id="8884f-211">For datasets of numbers, this is a sum by default.</span></span> <span data-ttu-id="8884f-212">该`summarizeBy`属性基于[AggregationFunction](/javascript/api/excel/excel.aggregationfunction)类型定义此行为。</span><span class="sxs-lookup"><span data-stu-id="8884f-212">The `summarizeBy` property defines this behavior based on an [AggregationFunction](/javascript/api/excel/excel.aggregationfunction) type.</span></span>

<span data-ttu-id="8884f-213">当前支持的聚合函数类型为`Sum`、 `Count` `Average` `Max` `Min` `Product` `CountNumbers` `StandardDeviation` `StandardDeviationP` `Variance`、、、、、、、、、和`Automatic` （默认值`VarianceP`）。</span><span class="sxs-lookup"><span data-stu-id="8884f-213">The currently supported aggregation function types are `Sum`, `Count`, `Average`, `Max`, `Min`, `Product`, `CountNumbers`, `StandardDeviation`, `StandardDeviationP`, `Variance`, `VarianceP`, and `Automatic` (the default).</span></span>

<span data-ttu-id="8884f-214">下面的代码示例将聚合更改为数据的平均值。</span><span class="sxs-lookup"><span data-stu-id="8884f-214">The following code samples changes the aggregation to be averages of the data.</span></span>

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

## <a name="change-calculations-with-a-showasrule"></a><span data-ttu-id="8884f-215">使用 ShowAsRule 更改计算</span><span class="sxs-lookup"><span data-stu-id="8884f-215">Change calculations with a ShowAsRule</span></span>

<span data-ttu-id="8884f-216">默认情况下，数据透视表将单独聚合其行和列层次结构的数据。</span><span class="sxs-lookup"><span data-stu-id="8884f-216">PivotTables, by default, aggregate the data of their row and column hierarchies independently.</span></span> <span data-ttu-id="8884f-217">[ShowAsRule](/javascript/api/excel/excel.showasrule)将数据层次结构更改为基于数据透视表中的其他项的输出值。</span><span class="sxs-lookup"><span data-stu-id="8884f-217">A [ShowAsRule](/javascript/api/excel/excel.showasrule) changes the data hierarchy to output values based on other items in the PivotTable.</span></span>

<span data-ttu-id="8884f-218">`ShowAsRule`对象具有三个属性：</span><span class="sxs-lookup"><span data-stu-id="8884f-218">The `ShowAsRule` object has three properties:</span></span>

- <span data-ttu-id="8884f-219">`calculation`：要应用于数据层次结构的相对计算的类型（默认值为`none`）。</span><span class="sxs-lookup"><span data-stu-id="8884f-219">`calculation`: The type of relative calculation to apply to the data hierarchy (the default is `none`).</span></span>
- <span data-ttu-id="8884f-220">`baseField`：在应用计算之前包含基础数据的层次结构中的[透视字段](/javascript/api/excel/excel.pivotfield)。</span><span class="sxs-lookup"><span data-stu-id="8884f-220">`baseField`: The [PivotField](/javascript/api/excel/excel.pivotfield) in the hierarchy containing the base data before the calculation is applied.</span></span> <span data-ttu-id="8884f-221">由于 Excel 数据透视表的层次结构与字段的一对一映射，因此将使用相同的名称来访问层次结构和字段。</span><span class="sxs-lookup"><span data-stu-id="8884f-221">Since Excel PivotTables have a one-to-one mapping of hierarchy to field, you'll use the same name to access both the hierarchy and the field.</span></span>
- <span data-ttu-id="8884f-222">`baseItem`：个人[PivotItem](/javascript/api/excel/excel.pivotitem)根据计算类型与基本字段的值进行比较。</span><span class="sxs-lookup"><span data-stu-id="8884f-222">`baseItem`: The individual [PivotItem](/javascript/api/excel/excel.pivotitem) compared against the values of the base fields based on the calculation type.</span></span> <span data-ttu-id="8884f-223">并非所有计算都需要此字段。</span><span class="sxs-lookup"><span data-stu-id="8884f-223">Not all calculations require this field.</span></span>

<span data-ttu-id="8884f-224">以下示例将场数据层次结构中的 " **Crates**总数" 的计算设置为列总计的百分比。</span><span class="sxs-lookup"><span data-stu-id="8884f-224">The following example sets the calculation on the **Sum of Crates Sold at Farm** data hierarchy to be a percentage of the column total.</span></span>
<span data-ttu-id="8884f-225">我们仍希望将粒度扩展到水果类型级别，因此我们将使用**类型**行层次结构及其基础字段。</span><span class="sxs-lookup"><span data-stu-id="8884f-225">We still want the granularity to extend to the fruit type level, so we’ll use the **Type** row hierarchy and its underlying field.</span></span>
<span data-ttu-id="8884f-226">该示例还将**服务器场**作为第一个行的层次结构，因此服务器场总数将显示每个服务器场也负责生成的百分比。</span><span class="sxs-lookup"><span data-stu-id="8884f-226">The example also has **Farm** as the first row hierarchy, so the farm total entries display the percentage each farm is responsible for producing as well.</span></span>

![显示与每个场中的单个服务器场和各个水果类型的总和相关的水果销售百分比的数据透视表。](../images/excel-pivots-showas-percentage.png)

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

<span data-ttu-id="8884f-228">上面的示例将计算设置为相对于单个行层次结构的字段的列。</span><span class="sxs-lookup"><span data-stu-id="8884f-228">The previous example set the calculation to the column, relative to the field of an individual row hierarchy.</span></span> <span data-ttu-id="8884f-229">当计算与单个项目相关时，请使用`baseItem`属性。</span><span class="sxs-lookup"><span data-stu-id="8884f-229">When the calculation relates to an individual item, use the `baseItem` property.</span></span>

<span data-ttu-id="8884f-230">下面的示例演示了`differenceFrom`计算。</span><span class="sxs-lookup"><span data-stu-id="8884f-230">The following example shows the `differenceFrom` calculation.</span></span> <span data-ttu-id="8884f-231">它显示场中相对于**服务器场**中的销售数据层次结构条目的不同之处。</span><span class="sxs-lookup"><span data-stu-id="8884f-231">It displays the difference of the farm crate sales data hierarchy entries relative to those of **A Farms**.</span></span>
<span data-ttu-id="8884f-232">`baseField`是**服务器场**，因此我们看到其他服务器场之间的差异，以及每种类型的类似水果（在此示例中**类型**也是行层次结构）的细目。</span><span class="sxs-lookup"><span data-stu-id="8884f-232">The `baseField` is **Farm**, so we see the differences between the other farms, as well as breakdowns for each type of like fruit (**Type** is also a row hierarchy in this example).</span></span>

![显示 "一群" 和其他 "服务器场" 之间的水果销售差异的数据透视表。](../images/excel-pivots-showas-differencefrom.png)

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

## <a name="pivottable-layouts"></a><span data-ttu-id="8884f-236">数据透视表布局</span><span class="sxs-lookup"><span data-stu-id="8884f-236">PivotTable layouts</span></span>

<span data-ttu-id="8884f-237">[PivotLayout](/javascript/api/excel/excel.pivotlayout)定义层次结构及其数据的位置。</span><span class="sxs-lookup"><span data-stu-id="8884f-237">A [PivotLayout](/javascript/api/excel/excel.pivotlayout) defines the placement of hierarchies and their data.</span></span> <span data-ttu-id="8884f-238">您可以访问布局以确定存储数据的区域。</span><span class="sxs-lookup"><span data-stu-id="8884f-238">You access the layout to determine the ranges where data is stored.</span></span>

<span data-ttu-id="8884f-239">下图显示了哪些布局函数调用对应于数据透视表的区域。</span><span class="sxs-lookup"><span data-stu-id="8884f-239">The following diagram shows which layout function calls correspond to which ranges of the PivotTable.</span></span>

![显示由布局的 get range 函数返回的数据透视表的节的图表。](../images/excel-pivots-layout-breakdown.png)

<span data-ttu-id="8884f-241">下面的代码演示如何通过布局获取数据透视表数据的最后一行。</span><span class="sxs-lookup"><span data-stu-id="8884f-241">The following code demonstrates how to get the last row of the PivotTable data by going through the layout.</span></span> <span data-ttu-id="8884f-242">然后将这些值汇总到一起以进行总计。</span><span class="sxs-lookup"><span data-stu-id="8884f-242">Those values are then summed together for a grand total.</span></span>

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    // Get the totals for each data hierarchy from the layout.
    var range = pivotTable.layout.getDataBodyRange();
    var grandTotalRange = range.getLastRow();
    grandTotalRange.load("address");
    return context.sync().then(function () {
        // Sum the totals from the PivotTable data hierarchies and place them in a new range.
        var masterTotalRange = context.workbook.worksheets.getActiveWorksheet().getRange("B27:C27");
        masterTotalRange.formulas = [["All Crates", "=SUM(" + grandTotalRange.address + ")"]];
    });
});
```

<span data-ttu-id="8884f-243">数据透视表具有三种布局样式：紧凑、大纲和表格。</span><span class="sxs-lookup"><span data-stu-id="8884f-243">PivotTables have three layout styles: Compact, Outline, and Tabular.</span></span> <span data-ttu-id="8884f-244">我们在前面的示例中看到了压缩样式。</span><span class="sxs-lookup"><span data-stu-id="8884f-244">We’ve seen the compact style in the previous examples.</span></span>

<span data-ttu-id="8884f-245">下面的示例分别使用大纲样式和表格样式。</span><span class="sxs-lookup"><span data-stu-id="8884f-245">The following examples use the outline and tabular styles, respectively.</span></span> <span data-ttu-id="8884f-246">此代码示例演示如何在不同的布局之间循环。</span><span class="sxs-lookup"><span data-stu-id="8884f-246">The code sample shows how to cycle between the different layouts.</span></span>

### <a name="outline-layout"></a><span data-ttu-id="8884f-247">大纲布局</span><span class="sxs-lookup"><span data-stu-id="8884f-247">Outline layout</span></span>

![使用大纲布局的数据透视表。](../images/excel-pivots-outline-layout.png)

### <a name="tabular-layout"></a><span data-ttu-id="8884f-249">表格布局</span><span class="sxs-lookup"><span data-stu-id="8884f-249">Tabular layout</span></span>

![使用表格布局的数据透视表。](../images/excel-pivots-tabular-layout.png)

## <a name="change-hierarchy-names"></a><span data-ttu-id="8884f-251">更改层次结构名称</span><span class="sxs-lookup"><span data-stu-id="8884f-251">Change hierarchy names</span></span>

<span data-ttu-id="8884f-252">层次结构字段是可编辑的。</span><span class="sxs-lookup"><span data-stu-id="8884f-252">Hierarchy fields are editable.</span></span> <span data-ttu-id="8884f-253">下面的代码演示如何更改两个数据层次结构的显示名称。</span><span class="sxs-lookup"><span data-stu-id="8884f-253">The following code demonstrates how to change the displayed names of two data hierarchies.</span></span>

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

## <a name="delete-a-pivottable"></a><span data-ttu-id="8884f-254">删除数据透视表</span><span class="sxs-lookup"><span data-stu-id="8884f-254">Delete a PivotTable</span></span>

<span data-ttu-id="8884f-255">使用它们的名称删除数据透视表。</span><span class="sxs-lookup"><span data-stu-id="8884f-255">PivotTables are deleted by using their name.</span></span>

```js
Excel.run(function (context) {
    context.workbook.worksheets.getItem("Pivot").pivotTables.getItem("Farm Sales").delete();
    return context.sync();
});
```

## <a name="see-also"></a><span data-ttu-id="8884f-256">另请参阅</span><span class="sxs-lookup"><span data-stu-id="8884f-256">See also</span></span>

- [<span data-ttu-id="8884f-257">Excel JavaScript API 基本编程概念</span><span class="sxs-lookup"><span data-stu-id="8884f-257">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="8884f-258">Excel JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="8884f-258">Excel JavaScript API Reference</span></span>](/javascript/api/excel)
